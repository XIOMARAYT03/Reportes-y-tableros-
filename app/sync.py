"""Servicio de sincronizacion del Kardex con Odoo 18.

Objetivos:
  * Desacoplar la lectura del kardex de la latencia del ERP: los movimientos
    se descargan de forma incremental a una base SQLite local y el reporte se
    arma desde ahi.
  * Detectar altas y modificaciones de movimientos por marca de agua
    (``write_date`` + ``id``) sin volver a traer el historico completo.
  * No duplicar movimientos: la clave primaria es el id de Odoo y la
    escritura es un ``INSERT ... ON CONFLICT DO UPDATE`` dentro de una
    transaccion.
  * Registrar el estado, la fecha de ultima sincronizacion y los errores para
    poder revisarlos desde la interfaz.
  * Recuperarse de desconexiones: la marca de agua solo avanza cuando la
    pagina se guardo correctamente, de modo que un fallo reintenta esa pagina.

El bus de Odoo se verifica en tiempo de ejecucion (canal de notificaciones de
Odoo <= 16 por longpolling o WebSocket de Odoo >= 17).  Si la instalacion no
lo expone, el servicio funciona con sondeo periodico y el intervalo real queda
informado en ``estado()``.
"""
import json
import os
import sqlite3
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
RUTA_DB_DEFECTO = os.path.join(HERE, "inventario.db")
INTERVALO_DEFECTO = 60
PAGINA = 500
PAGINA_MIN = 25
CAMPOS_MOVIMIENTO = [
    "id", "name", "date", "quantity", "product_qty", "product_id", "product_uom",
    "company_id", "state", "location_id", "location_dest_id", "picking_id",
    "picking_type_id", "purchase_line_id", "sale_line_id", "origin_returned_move_id",
    "origin", "write_date", "create_uid",
]

ESQUEMA = """
CREATE TABLE IF NOT EXISTS sync_estado (
    modelo TEXT PRIMARY KEY,
    marca_agua_fecha TEXT,
    marca_agua_id INTEGER,
    ultimo_intento TEXT,
    ultimo_exito TEXT,
    estado TEXT,
    error TEXT,
    filas INTEGER DEFAULT 0,
    completo INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS sync_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT, nivel TEXT, modelo TEXT, mensaje TEXT
);
CREATE TABLE IF NOT EXISTS movimientos (
    id INTEGER PRIMARY KEY,
    producto_id INTEGER, company_id INTEGER, nombre TEXT, fecha TEXT, write_date TEXT,
    estado TEXT, cantidad REAL, product_qty REAL, uom_id INTEGER,
    origen_id INTEGER, destino_id INTEGER, uso_origen TEXT, uso_destino TEXT,
    picking_id INTEGER, picking_name TEXT, picking_seq TEXT, picking_code TEXT,
    picking_type_id INTEGER, origen_retorno_id INTEGER,
    compra_line_id INTEGER, venta_line_id INTEGER, referencia TEXT,
    usuario_id INTEGER, sincronizado_en TEXT
);
CREATE INDEX IF NOT EXISTS ix_mov_producto ON movimientos (producto_id, fecha);
CREATE INDEX IF NOT EXISTS ix_mov_write ON movimientos (write_date, id);
CREATE TABLE IF NOT EXISTS capas (
    id INTEGER PRIMARY KEY, move_id INTEGER, producto_id INTEGER, company_id INTEGER,
    cantidad REAL, valor REAL, costo REAL
);
CREATE INDEX IF NOT EXISTS ix_capas_move ON capas (move_id);
CREATE TABLE IF NOT EXISTS productos (
    id INTEGER PRIMARY KEY, codigo TEXT, nombre TEXT, uom_id INTEGER, company_id INTEGER,
    categoria TEXT, metodo_costo TEXT, valuacion TEXT, precio_estandar REAL,
    qty_fisica REAL, qty_reservado REAL, qty_disponible REAL, actualizado TEXT
);
CREATE INDEX IF NOT EXISTS ix_prod_codigo ON productos (codigo);
CREATE TABLE IF NOT EXISTS ubicaciones (
    id INTEGER PRIMARY KEY, usage TEXT, company_id INTEGER, nombre TEXT, complete_name TEXT
);
CREATE TABLE IF NOT EXISTS unidades (
    id INTEGER PRIMARY KEY, nombre TEXT, category_id INTEGER, factor REAL,
    factor_inv REAL, precision REAL
);
CREATE TABLE IF NOT EXISTS quants (
    producto_id INTEGER, location_id INTEGER, cantidad REAL, reservado REAL,
    PRIMARY KEY (producto_id, location_id)
);
CREATE INDEX IF NOT EXISTS ix_quants_prod ON quants (producto_id);
CREATE TABLE IF NOT EXISTS conciliaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, empresa INTEGER, producto_id INTEGER,
    codigo TEXT, unidad TEXT, stock_odoo REAL, stock_kardex REAL, diferencia REAL,
    diferencia_rel REAL, estado TEXT, detalle TEXT
);
CREATE INDEX IF NOT EXISTS ix_conc_empresa ON conciliaciones (empresa, ts);
"""


def _ahora():
    return datetime.now()


def _iso(dt=None):
    return (dt or _ahora()).strftime("%Y-%m-%d %H:%M:%S")


class ServicioSync:
    """Servicio de sincronizacion incremental con Odoo (hilo unico)."""

    _instancia = None
    _instancia_lock = threading.Lock()

    def __init__(self, sr, id_of, ref, url, db_path=None, intervalo=None,
                 empresas=None, log=None):
        self._sr = sr
        self._id_of = id_of
        self._ref = ref
        self.url = (url or "").rstrip("/")
        self.ruta = db_path or os.environ.get("KARDEX_DB") or RUTA_DB_DEFECTO
        self.intervalo = int(intervalo or os.environ.get("KARDEX_SYNC_INTERVALO_SEG")
                             or INTERVALO_DEFECTO)
        self.empresas = list(empresas or [])
        self._log = log or (lambda *a: None)
        self._stop = threading.Event()
        self._hilo = None
        self._local = threading.local()
        self._lock = threading.Lock()
        self._estado = {"activo": False, "estado": "detenido", "error": "",
                        "ultima_sincronizacion": None, "ultimo_intento": None,
                        "duracion": 0.0, "filas": 0, "ciclo": 0}
        self._bus = {"verificado": False, "disponible": False, "detalle": []}
        self._inicial_completa = False
        self._tipos_cache = None
        self._tipos_expira = 0
        self._crear_esquema()

    # ----------------------------------------------------------------- #
    # Base de datos
    # ----------------------------------------------------------------- #
    def _crear_esquema(self):
        conn = self._conexion()
        with conn:
            conn.executescript(ESQUEMA)
            self._migrar_esquema(conn)

    def _migrar_esquema(self, conn):
        """Agrega a las tablas existentes las columnas incorporadas despues."""
        cambios = [("movimientos", "usuario_id", "INTEGER")]
        for tabla, columna, tipo in cambios:
            cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({tabla})")}
            if cols and columna not in cols:
                conn.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {tipo}")

    def _conexion(self):
        conn = getattr(self._local, "conn", None)
        if conn is None:
            conn = sqlite3.connect(self.ruta, timeout=30)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            self._local.conn = conn
        return conn

    # ----------------------------------------------------------------- #
    # Arranque / parada (instancia unica)
    # ----------------------------------------------------------------- #
    @classmethod
    def instancia(cls, *a, **k):
        with cls._instancia_lock:
            if cls._instancia is None:
                cls._instancia = cls(*a, **k)
            return cls._instancia

    def arrancar(self):
        with self._lock:
            if self._hilo and self._hilo.is_alive():
                return False
            self._stop.clear()
            self._hilo = threading.Thread(target=self._bucle, name="kardex-sync", daemon=True)
            self._hilo.start()
            self._estado["activo"] = True
            return True

    def detener(self):
        self._stop.set()
        with self._lock:
            self._hilo = None
        self._estado["activo"] = False

    def _bucle(self):
        # Espera inicial corta: el servidor web debe responder de inmediato.
        if self._stop.wait(5):
            return
        while not self._stop.is_set():
            try:
                self.sincronizar(motivo="programado")
            except Exception as exc:                      # pragma: no cover - defensivo
                self._estado["estado"] = "error"
                self._estado["error"] = str(exc)
                self._log(f"[sync] error en el ciclo programado: {exc}")
            self._stop.wait(self.intervalo)

    # ----------------------------------------------------------------- #
    # Registro
    # ----------------------------------------------------------------- #
    def _log_bd(self, nivel, modelo, mensaje):
        try:
            conn = self._conexion()
            with conn:
                conn.execute("INSERT INTO sync_log (ts, nivel, modelo, mensaje) "
                             "VALUES (?,?,?,?)", (_iso(), nivel, modelo, str(mensaje)[:2000]))
        except sqlite3.Error:
            pass

    def _estado_modelo(self, modelo):
        conn = self._conexion()
        fila = conn.execute("SELECT * FROM sync_estado WHERE modelo=?", (modelo,)).fetchone()
        if not fila:
            conn.execute("INSERT INTO sync_estado (modelo, estado) VALUES (?,?)",
                         (modelo, "nuevo"))
            with conn:
                conn.commit()
            fila = conn.execute("SELECT * FROM sync_estado WHERE modelo=?", (modelo,)).fetchone()
        return dict(fila)

    def _guardar_estado(self, modelo, **campos):
        conn = self._conexion()
        sets = ", ".join(f"{k}=?" for k in campos)
        vals = list(campos.values()) + [modelo]
        with conn:
            conn.execute(f"UPDATE sync_estado SET {sets} WHERE modelo=?", vals)

    # ----------------------------------------------------------------- #
    # Sincronizacion
    # ----------------------------------------------------------------- #
    def sincronizar(self, completo=False, motivo="manual", producto_id=None):
        """Ejecuta un ciclo de sincronizacion. Devuelve un resumen."""
        inicio = time.time()
        with self._lock:
            if self._estado.get("sincronizando"):
                return {"ok": False, "mensaje": "Ya hay una sincronizacion en curso",
                        "estado": dict(self._estado)}
            self._estado["sincronizando"] = True
            self._estado["ultimo_intento"] = _iso()
            self._estado["estado"] = "sincronizando"
            self._estado["error"] = ""
        resumen = {"ok": True, "inicio": _iso()}
        try:
            conn = self._conexion()
            with conn:
                conn.commit()
            self._sincronizar_catalogos()
            filas_mov, productos = self._sincronizar_movimientos(completo=completo,
                                                                 producto_id=producto_id)
            filas_prod = self._sincronizar_quants(productos)
            if completo:
                self._guardar_estado("stock.move", completo=1)
            resumen.update({
                "movimientos": filas_mov,
                "quants": filas_prod,
                "productos": len(productos),
                "duracion": round(time.time() - inicio, 2),
                "completo": completo,
                "motivo": motivo,
            })
            self._estado.update({"estado": "ok", "error": "",
                                 "ultima_sincronizacion": _iso(),
                                 "duracion": round(time.time() - inicio, 2),
                                 "filas": filas_mov, "ciclo": self._estado.get("ciclo", 0) + 1})
            self._log_bd("info", "sync", f"Ciclo {motivo}: {filas_mov} movimientos, "
                                       f"{filas_prod} quants, "
                                       f"{resumen['duracion']}s")
        except Exception as exc:
            self._estado["estado"] = "error"
            self._estado["error"] = str(exc)
            self._log_bd("error", "sync", f"Ciclo {motivo} fallo: {exc}")
            self._log(f"[sync] error: {exc}")
            resumen.update({"ok": False, "error": str(exc)})
        finally:
            with self._lock:
                self._estado["sincronizando"] = False
        return resumen

    # ---------------- catalogos ------------------------------------- #
    def _sincronizar_catalogos(self):
        locs = self._sr("stock.location", [], ["id", "usage", "company_id", "name", "complete_name"])
        unis = self._sr("uom.uom", [], ["id", "name", "category_id", "factor", "factor_inv", "rounding"])
        conn = self._conexion()
        with conn:
            conn.executemany(
                "INSERT INTO ubicaciones (id, usage, company_id, nombre, complete_name) "
                "VALUES (?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET usage=excluded.usage, "
                "company_id=excluded.company_id, nombre=excluded.nombre, "
                "complete_name=excluded.complete_name",
                [(l["id"], l.get("usage"), self._id_of(l.get("company_id")),
                  l.get("name") or "", l.get("complete_name") or "") for l in locs])
            conn.executemany(
                "INSERT INTO unidades (id, nombre, category_id, factor, factor_inv, precision) "
                "VALUES (?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET nombre=excluded.nombre, "
                "category_id=excluded.category_id, factor=excluded.factor, "
                "factor_inv=excluded.factor_inv, precision=excluded.precision",
                [(u["id"], u.get("name") or "", self._id_of(u.get("category_id")),
                  u.get("factor") or 1.0, u.get("factor_inv") or 1.0,
                  u.get("rounding") or 0.01) for u in unis])

    # ---------------- movimientos ------------------------------------ #
    def _dominio_movimientos(self, completo, producto_id=None):
        """Dominio base de movimientos (la marca de agua se agrega al paginar)."""
        dom = [("state", "=", "done")]
        if producto_id:
            dom.append(("product_id", "=", producto_id))
        return dom

    def sincronizar_producto(self, producto_id):
        """Sincroniza unicamente los movimientos de un producto."""
        return self.sincronizar(motivo="producto", producto_id=producto_id)

    def lanzar(self, completo=False, motivo="manual", producto_id=None):
        """Lanza una sincronizacion en segundo plano y devuelve de inmediato.

        La sincronizacion inicial completa puede tardar varios minutos, por
        lo que nunca debe ejecutarse dentro de una peticion web.
        """
        with self._lock:
            if self._estado.get("sincronizando"):
                return {"ok": False, "mensaje": "Ya hay una sincronizacion en curso"}
            hilo = threading.Thread(target=self.sincronizar, args=(completo,),
                                    kwargs={"motivo": motivo, "producto_id": producto_id},
                                    daemon=True, name="kardex-sync-manual")
            hilo.start()
        return {"ok": True, "mensaje": "Sincronizacion iniciada en segundo plano"}

    def _tipos_picking(self):
        """Mapa id -> {seq, code, name} de los tipos de picking (cacheado)."""
        if getattr(self, "_tipos_cache", None) and not self._tipos_expira:
            return self._tipos_cache
        tipos = self._sr("stock.picking.type", [], ["id", "sequence_code", "code", "name"])
        self._tipos_cache = {t["id"]: {"seq": t.get("sequence_code"), "code": t.get("code"),
                                      "name": t.get("name")} for t in tipos}
        self._tipos_expira = time.time() + 3600
        return self._tipos_cache

    def _existe_tabla(self, nombre):
        conn = self._conexion()
        r = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                         (nombre,)).fetchone()
        return r is not None

    def _sincronizar_movimientos(self, completo=False, producto_id=None):
        """Descarga los movimientos completados pagina a pagina.

        La marca de agua es lexicografica (write_date, id) y solo avanza cuando
        la pagina quedo guardada: si el ERP falla a mitad de camino, el ciclo
        siguiente retoma exactamente donde se quedo, sin perder ni duplicar
        movimientos.
        """
        conn = self._conexion()
        est = self._estado_modelo("stock.move")
        if completo:
            with conn:
                conn.execute("DELETE FROM movimientos")
                conn.execute("DELETE FROM capas")
                conn.execute("UPDATE sync_estado SET marca_agua_fecha=NULL, marca_agua_id=0, "
                             "completo=0 WHERE modelo='stock.move'")
            est = self._estado_modelo("stock.move")
        if producto_id:
            base = [("state", "=", "done"), ("product_id", "=", producto_id)]
            orden = "id asc"
        else:
            base = [("state", "=", "done")]
            orden = "write_date asc, id asc"
        wf, wid = est.get("marca_agua_fecha"), est.get("marca_agua_id") or 0
        if producto_id:
            wf, wid = None, 0
        total = 0
        productos = set()
        pagina = PAGINA
        while True:
            dom = list(base)
            if wf:
                dom += ["|", ("write_date", ">", wf),
                        "&", ("write_date", "=", wf), ("id", ">", wid)]
            try:
                filas = self._sr("stock.move", dom, CAMPOS_MOVIMIENTO,
                                 order=orden, limit=pagina)
            except Exception:
                if pagina <= PAGINA_MIN:
                    raise
                pagina = max(PAGINA_MIN, pagina // 2)
                self._log_bd("warn", "stock.move",
                             f"Pagina reducida a {pagina} por error de transporte")
                continue
            if not filas:
                break
            self._guardar_movimientos(filas)
            for f in filas:
                pid = self._id_of(f.get("product_id"))
                if pid:
                    productos.add(pid)
            total += len(filas)
            ultima = filas[-1]
            wf = ultima.get("write_date") or wf
            wid = ultima["id"]
            self._guardar_estado("stock.move", marca_agua_fecha=wf, marca_agua_id=wid,
                                  ultimo_exito=_iso(), estado="ok", filas=total)
            self._estado["filas"] = total
            if len(filas) < pagina:
                break
        self._sincronizar_capas_de_productos(producto_id)
        return total, productos

    def _guardar_movimientos(self, filas):
        """Guarda los movimientos (upsert) y devuelve el estado para la marca de agua."""
        locs = {l["id"]: l for l in self._leer_ubicaciones()}
        pts = self._tipos_picking()
        conn = self._conexion()
        datos = []
        for m in filas:
            pt = pts.get(self._id_of(m.get("picking_type_id"))) or {}
            src = locs.get(self._id_of(m.get("location_id")) or 0, {})
            dst = locs.get(self._id_of(m.get("location_dest_id")) or 0, {})
            datos.append((
                m["id"], self._id_of(m.get("product_id")), self._id_of(m.get("company_id")),
                m.get("name") or "", m.get("date") or "", m.get("write_date") or "",
                m.get("state") or "", m.get("quantity") or 0.0, m.get("product_qty") or 0.0,
                self._id_of(m.get("product_uom")), self._id_of(m.get("location_id")),
                self._id_of(m.get("location_dest_id")), src.get("usage"), dst.get("usage"),
                self._id_of(m.get("picking_id")), None, pt.get("seq"), pt.get("code"),
                self._id_of(m.get("picking_type_id")),
                self._id_of(m.get("origin_returned_move_id")),
                self._id_of(m.get("purchase_line_id")), self._id_of(m.get("sale_line_id")),
                m.get("origin") or "", self._id_of(m.get("create_uid")), _iso(),
            ))
        with conn:
            conn.executemany(
                "INSERT INTO movimientos (id, producto_id, company_id, nombre, fecha, write_date, "
                "estado, cantidad, product_qty, uom_id, origen_id, destino_id, uso_origen, "
                "uso_destino, picking_id, picking_name, picking_seq, picking_code, "
                "picking_type_id, origen_retorno_id, compra_line_id, venta_line_id, referencia, "
                "usuario_id, sincronizado_en) VALUES (" + ",".join(["?"] * 25) + ") "
                "ON CONFLICT(id) DO UPDATE SET producto_id=excluded.producto_id, "
                "company_id=excluded.company_id, nombre=excluded.nombre, fecha=excluded.fecha, "
                "write_date=excluded.write_date, estado=excluded.estado, "
                "cantidad=excluded.cantidad, product_qty=excluded.product_qty, "
                "uom_id=excluded.uom_id, origen_id=excluded.origen_id, "
                "destino_id=excluded.destino_id, uso_origen=excluded.uso_origen, "
                "uso_destino=excluded.uso_destino, picking_id=excluded.picking_id, "
                "picking_seq=excluded.picking_seq, picking_code=excluded.picking_code, "
                "picking_type_id=excluded.picking_type_id, "
                "origen_retorno_id=excluded.origen_retorno_id, "
                "compra_line_id=excluded.compra_line_id, venta_line_id=excluded.venta_line_id, "
                "referencia=excluded.referencia, usuario_id=excluded.usuario_id, "
                "sincronizado_en=excluded.sincronizado_en",
                datos)

    def _avanzar_marca_agua(self, filas, completo):
        if not filas:
            return
        wf = max((f.get("write_date") or "") for f in filas)
        wid = max(f["id"] for f in filas if (f.get("write_date") or "") == wf) if wf else 0
        self._guardar_estado("stock.move", marca_agua_fecha=wf, marca_agua_id=wid,
                              ultimo_exito=_iso(), estado="ok", filas=len(filas))

    def _leer_ubicaciones(self):
        conn = self._conexion()
        return [dict(r) for r in conn.execute(
            "SELECT id, usage, company_id, nombre, complete_name FROM ubicaciones")]

    def _sincronizar_capas_de_productos(self, producto_id=None):
        """Descarga las capas de valorizacion de los movimientos sincronizados."""
        conn = self._conexion()
        if producto_id:
            ids = [r[0] for r in conn.execute(
                "SELECT DISTINCT id FROM movimientos WHERE producto_id=?", (producto_id,))]
        else:
            ids = [r[0] for r in conn.execute(
                "SELECT DISTINCT id FROM movimientos WHERE id NOT IN "
                "(SELECT DISTINCT move_id FROM capas)")]
        total = 0
        for i in range(0, len(ids), 400):
            lote = ids[i:i + 400]
            capas = self._sr("stock.valuation.layer", [("stock_move_id", "in", lote)],
                             ["id", "stock_move_id", "product_id", "company_id", "quantity",
                              "value", "unit_cost"])
            with conn:
                conn.executemany(
                    "INSERT INTO capas (id, move_id, producto_id, company_id, cantidad, valor, "
                    "costo) VALUES (?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET "
                    "move_id=excluded.move_id, producto_id=excluded.producto_id, "
                    "company_id=excluded.company_id, cantidad=excluded.cantidad, "
                    "valor=excluded.valor, costo=excluded.costo",
                    [(c["id"], self._id_of(c.get("stock_move_id")),
                      self._id_of(c.get("product_id")), self._id_of(c.get("company_id")),
                      c.get("quantity") or 0.0, c.get("value") or 0.0,
                      c.get("unit_cost") or 0.0) for c in capas])
            total += len(capas)
        return total

    def _sincronizar_quants(self, producto_ids=None):
        """Actualiza el stock fisico (quants) de los productos indicados."""
        if producto_ids is None:
            producto_ids = []
        ids = sorted({p for p in producto_ids if p})
        if not ids:
            return 0
        conn = self._conexion()
        total = 0
        for i in range(0, len(ids), 200):
            lote = ids[i:i + 200]
            quants = self._sr("stock.quant", [("product_id", "in", lote)],
                              ["product_id", "location_id", "quantity", "reserved_quantity"])
            prods = self._sr("product.product", [("id", "in", lote)],
                             ["id", "default_code", "name", "uom_id", "standard_price"])
            suma = {}
            for q in quants:
                pid = self._id_of(q.get("product_id"))
                d = suma.setdefault(pid, [0.0, 0.0])
                d[0] += q.get("quantity") or 0.0
                d[1] += q.get("reserved_quantity") or 0.0
            with conn:
                conn.executemany(
                    "INSERT INTO quants (producto_id, location_id, cantidad, reservado) "
                    "VALUES (?,?,?,?) ON CONFLICT(producto_id, location_id) DO UPDATE SET "
                    "cantidad=excluded.cantidad, reservado=excluded.reservado",
                    [(self._id_of(q.get("product_id")), self._id_of(q.get("location_id")),
                      q.get("quantity") or 0.0, q.get("reserved_quantity") or 0.0)
                     for q in quants])
            filas = []
            for p in prods:
                fisico, reservado = suma.get(p["id"], [0.0, 0.0])
                filas.append((p["id"], p.get("default_code") or "", p.get("name") or "",
                              self._id_of(p.get("uom_id")), None, None, None, None,
                              p.get("standard_price") or 0.0, fisico, reservado,
                              fisico - reservado, _iso()))
            with conn:
                conn.executemany(
                    "INSERT INTO productos (id, codigo, nombre, uom_id, company_id, "
                    "categoria, metodo_costo, valuacion, precio_estandar, qty_fisica, "
                    "qty_reservado, qty_disponible, actualizado) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET "
                    "codigo=excluded.codigo, nombre=excluded.nombre, uom_id=excluded.uom_id, "
                    "precio_estandar=excluded.precio_estandar, qty_fisica=excluded.qty_fisica, "
                    "qty_reservado=excluded.qty_reservado, qty_disponible=excluded.qty_disponible, "
                    "actualizado=excluded.actualizado", filas)
            total += len(filas)
        return total

    # ----------------------------------------------------------------- #
    # Verificacion del bus de Odoo
    # ----------------------------------------------------------------- #
    def verificar_bus(self, usuario=None, clave=None, db=None):
        """Comprueba si la instalacion expone el canal de notificaciones de Odoo.

        Odoo <= 16: ``/longpolling/poll`` (HTTP).  Odoo >= 17: ``/websocket``.
        Ambos exigen una sesion web (``session_id``), que se obtiene con
        ``/web/session/authenticate``.  Se documenta el resultado real.
        """
        detalle = []
        if not self.url:
            return {"verificado": True, "disponible": False, "detalle": ["URL de Odoo no configurada"]}
        base = self.url.split("://", 1)[-1]
        host, _, port = base.partition(":")
        puerto = int(port or 80)
        http_port = 8072
        # 1) sesion web (necesaria para cualquier canal)
        sesion = None
        try:
            cuerpo = json.dumps({"jsonrpc": "2.0", "method": "call",
                                 "params": {"db": db, "login": usuario, "password": clave}}).encode()
            req = urllib.request.Request(f"http://{host}:{puerto}/web/session/authenticate",
                                         data=cuerpo,
                                         headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read().decode())
            sesion = (data.get("result") or {}).get("session_id")
            detalle.append("web/session/authenticate: %s" % ("ok" if sesion else
                                                               "denegado por el ERP"))
        except Exception as exc:
            detalle.append("web/session/authenticate fallo: %s" % exc)
        # 2) websocket (Odoo >= 17)
        ws = None
        try:
            with urllib.request.urlopen(f"http://{host}:{http_port}/websocket", timeout=8) as r:
                ws = r.status
            detalle.append("/websocket (Odoo >= 17): HTTP %s" % ws)
        except urllib.error.HTTPError as exc:
            ws = exc.code
            detalle.append("/websocket (Odoo >= 17): HTTP %s (no disponible)" % exc.code)
        except Exception as exc:
            ws = None
            detalle.append("/websocket (Odoo >= 17): sin respuesta (%s)" % type(exc).__name__)
        # 3) longpolling (Odoo <= 16)
        try:
            cuerpo = json.dumps({"id": 1, "params": {"channels": ["stock.move"],
                                                     "last_notif_id": 0, "timeout": 5}}).encode()
            req = urllib.request.Request(f"http://{host}:{http_port}/longpolling/poll",
                                         data=cuerpo,
                                         headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=12) as r:
                detalle.append("/longpolling/poll (Odoo <= 16): HTTP %s" % r.status)
                largo = True
        except urllib.error.HTTPError as exc:
            detalle.append("/longpolling/poll (Odoo <= 16): HTTP %s (no disponible)" % exc.code)
            largo = False
        except Exception as exc:
            detalle.append("/longpolling/poll (Odoo <= 16): sin respuesta (%s)" % type(exc).__name__)
            largo = False
        disponible = bool(sesion) and ((ws and ws != 404) or largo)
        self._bus = {"verificado": True, "disponible": disponible, "detalle": detalle,
                     "verificado_en": _iso()}
        if not disponible:
            self._log_bd("info", "bus", "Bus no disponible: " + " | ".join(detalle))
        return self._bus

    # ----------------------------------------------------------------- #
    # Estado y lectura del espejo
    # ----------------------------------------------------------------- #
    def listo(self):
        """True si el espejo local puede usarse para armar el kardex."""
        est = self._estado_modelo("stock.move")
        if not est.get("completo"):
            return False
        ultimo = est.get("ultimo_exito")
        if not ultimo:
            return False
        try:
            delta = (_ahora() - datetime.strptime(ultimo, "%Y-%m-%d %H:%M:%S")).total_seconds()
        except ValueError:
            return False
        return delta <= max(self.intervalo * 3, 300)

    def estado(self):
        est = self._estado_modelo("stock.move")
        conn = self._conexion()
        try:
            total = conn.execute("SELECT COUNT(*) FROM movimientos").fetchone()[0]
        except sqlite3.Error:
            total = 0
        return {
            "activo": bool(self._hilo and self._hilo.is_alive()),
            "sincronizando": bool(self._estado.get("sincronizando")),
            "estado": self._estado.get("estado"),
            "error": self._estado.get("error") or "",
            "intervalo_seg": self.intervalo,
            "ultima_sincronizacion": est.get("ultimo_exito") or self._estado.get("ultima_sincronizacion"),
            "ultimo_intento": self._estado.get("ultimo_intento"),
            "duracion": self._estado.get("duracion"),
            "marca_agua": {"write_date": est.get("marca_agua_fecha"), "id": est.get("marca_agua_id")},
            "movimientos_mirrored": total,
            "sincronizacion_inicial_completa": bool(est.get("completo")),
            "espejo_disponible": self.listo(),
            "bus": self._bus,
            "modo": "espejo local (SQLite)" if self.listo() else "consulta directa a Odoo",
            "base_datos": self.ruta,
        }

    def errores(self, limite=50):
        conn = self._conexion()
        filas = conn.execute("SELECT ts, nivel, modelo, mensaje FROM sync_log "
                             "WHERE nivel='error' ORDER BY id DESC LIMIT ?", (limite,)).fetchall()
        return [dict(f) for f in filas]

    def _filas(self, sql, params=()):
        conn = self._conexion()
        return [dict(f) for f in conn.execute(sql, params).fetchall()]

    def movimientos_producto(self, producto_id):
        return self._filas("SELECT * FROM movimientos WHERE producto_id=? ORDER BY fecha, id",
                           (producto_id,))

    def capas_producto(self, producto_id):
        return self._filas("SELECT * FROM capas WHERE producto_id=?", (producto_id,))

    def ubicaciones(self):
        return self._filas("SELECT * FROM ubicaciones")

    def unidades(self):
        return self._filas("SELECT * FROM unidades")

    def producto(self, producto_id):
        filas = self._filas("SELECT * FROM productos WHERE id=?", (producto_id,))
        return filas[0] if filas else None

    def quants_producto(self, producto_id, solo_alcance=None):
        """Quants del producto (opcionalmente filtrados por ubicaciones internas)."""
        if self.listo():
            filas = self._filas("SELECT * FROM quants WHERE producto_id=?", (producto_id,))
            out = [{"location_id": f["location_id"], "quantity": f["cantidad"],
                    "reserved_quantity": f["reservado"]} for f in filas]
        else:
            out = self._sr("stock.quant", [("product_id", "=", producto_id)],
                           ["location_id", "quantity", "reserved_quantity"])
            out = [{"location_id": self._id_of(q.get("location_id")),
                    "quantity": q.get("quantity") or 0.0,
                    "reserved_quantity": q.get("reserved_quantity") or 0.0} for q in out]
        if solo_alcance is None:
            return out
        return [q for q in out if self._id_of(q.get("location_id")) in solo_alcance]

    def productos_con_movimientos(self, limite=200):
        """Ids de los productos con mas movimientos sincronizados."""
        conn = self._conexion()
        filas = conn.execute("SELECT producto_id, COUNT(*) c FROM movimientos "
                              "WHERE producto_id IS NOT NULL GROUP BY producto_id "
                              "ORDER BY c DESC LIMIT ?", (limite,)).fetchall()
        return [f[0] for f in filas]

    def conciliaciones(self, empresa=None, limite=100):
        if empresa:
            return self._filas("SELECT * FROM conciliaciones WHERE empresa=? "
                               "ORDER BY id DESC LIMIT ?", (empresa, limite))
        return self._filas("SELECT * FROM conciliaciones ORDER BY id DESC LIMIT ?", (limite,))

    def guardar_conciliacion(self, empresa, datos):
        conn = self._conexion()
        with conn:
            conn.execute(
                "INSERT INTO conciliaciones (ts, empresa, producto_id, codigo, unidad, "
                "stock_odoo, stock_kardex, diferencia, diferencia_rel, estado, detalle) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (_iso(), empresa, datos.get("producto_id"), datos.get("codigo"),
                 datos.get("unidad"), datos.get("stock_odoo"), datos.get("stock_kardex"),
                 datos.get("diferencia"), datos.get("diferencia_relativa"),
                 datos.get("estado"), json.dumps(datos.get("detalle") or {}, ensure_ascii=False)))
