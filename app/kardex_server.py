"""Kardex Fimas - servidor local de reportes de inventario (Odoo) con control de acceso."""
import http.client
import io
import json
import os
import threading
import time
import traceback
import urllib.parse
import xml.parsers.expat
import xmlrpc.client
from collections import defaultdict
from datetime import datetime, timedelta
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import devoluciones
import inventario
import odoo_auth
import permisos
import sync as sync_mov

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("KARDEX_PORT", "8765"))
HOST = os.environ.get("KARDEX_HOST", "0.0.0.0")

# Ruta del .env con las credenciales del ERP (Odoo). En el servidor se define
# con KARDEX_ENV_PATH; si no existe se busca un .env junto al programa y, como
# ultimo recurso, la ruta legacy usada en desarrollo local.
_LEGACY_ENV = r"C:\Users\Usuario\my-app\gestion-comercial-odoo\backend\.env"


def _resolver_env():
    ruta = os.environ.get("KARDEX_ENV_PATH") or os.path.join(HERE, ".env")
    if os.path.isfile(ruta):
        return ruta
    if os.path.isfile(_LEGACY_ENV):
        return _LEGACY_ENV
    return ruta

TIPO_ENTRADA = {"IMP": "Ingreso (Importaci\u00f3n)", "IN": "Ingreso (Recepci\u00f3n)",
                "DEVCL": "Devoluci\u00f3n de cliente"}
MSG_LOGIN = ("El correo electronico o la contrasena ingresada no son correctos. "
             "Verifique sus credenciales e intente nuevamente.")
MSG_NO_ADMIN = "Lo siento, no tienes permisos de administrador para realizar esta acci\u00f3n."
_lock = threading.RLock()
_rpc = threading.local()
_cache = {"kardex": {}, "products": {}}
# Vigencia del cache del kardex (segundos). El ERP se vuelve a consultar al
# expirar o cuando el usuario pide un refresco explicito.
CACHE_TTL_SEG = 60


def load_env(path):
    cfg = {}
    if not os.path.isfile(path):
        raise SystemExit(
            f"[Kardex] No se encontro el archivo de configuracion .env en:\n  {path}\n"
            "Defina la variable de entorno KARDEX_ENV_PATH o cree un .env junto al programa "
            "(ver .env.example).")
    for line in open(path, encoding="utf-8"):
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()
    return cfg


ENV_PATH = _resolver_env()
CFG = load_env(ENV_PATH)
URL, DB, USER, PWD = CFG["odoo_url"], CFG["odoo_db"], CFG["odoo_username"], CFG["odoo_password"]


def _proxy():
    """ServerProxy por hilo: xmlrpc.client no es seguro entre hilos."""
    p = getattr(_rpc, "models", None)
    if p is None:
        p = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object", allow_none=True)
        _rpc.models = p
    return p


def _connect():
    """Abre un proxy limpio y autentica la sesion de Odoo."""
    common = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(DB, USER, PWD, {})
    if not uid:
        raise RuntimeError("No se pudo autenticar contra el ERP (Odoo)")
    return uid


def _execute(model, method, args, kwargs=None):
    """Ejecuta una llamada RPC con reintento ante fallo de transporte.

    xmlrpc.client no es seguro entre hilos, por eso cada hilo usa su propio
    ServerProxy (`_rpc` es un thread-local).  Las llamadas NO se serializan
    con un lock global: dos usuarios pueden consultar el ERP en paralelo y
    solo se reintenta (una vez) cuando falla el transporte, descartando el
    proxy del hilo para que la siguiente llamada use una conexion nueva.
    """
    ultimo = None
    for intento in range(2):
        try:
            return _proxy().execute_kw(DB, UID, PWD, model, method, args, kwargs or {})
        except xmlrpc.client.Fault:
            raise
        except (http.client.HTTPException, xmlrpc.client.ProtocolError,
                xmlrpc.client.ResponseError, OSError,
                xml.parsers.expat.ExpatError) as exc:
            ultimo = exc
            _rpc.models = None
            time.sleep(0.3 * (intento + 1))
    raise ultimo


UID = _connect()
odoo_auth.configurar(URL, DB, USER, PWD)

# Tamano de pagina para las lecturas al ERP. El servidor Odoo trunca las
# respuestas XML-RPC muy grandes, por eso las lecturas se hacen por paginas.
PAGE = 500
PAGE_MIN = 25


def sr(model, domain, fields, order=None, limit=None):
    """search_read paginado sobre el ERP.

    Divide la lectura en paginas (limit/offset) para que ninguna respuesta
    XML-RPC sea tan grande que el ERP la trunque. Si aun asi una pagina llega
    incompleta, se reduce el tamano de pagina y se reintenta.
    """
    base = {"fields": fields}
    if order:
        base["order"] = order
    total = limit if limit else None
    out, offset, size = [], 0, PAGE
    while True:
        restante = None if total is None else total - len(out)
        if restante is not None and restante <= 0:
            break
        n = size if restante is None else min(size, restante)
        try:
            filas = _execute(model, "search_read", [domain], dict(base, limit=n, offset=offset))
        except xml.parsers.expat.ExpatError:
            if size <= PAGE_MIN:
                raise
            size = max(PAGE_MIN, size // 2)
            continue
        out.extend(filas)
        offset += len(filas)
        if len(filas) < n:
            break
    return out


def id_of(v):
    if isinstance(v, dict):
        return v.get("id")
    if isinstance(v, (list, tuple)):
        return v[0] if v else None
    return v


def ref(v):
    return v[1] if isinstance(v, (list, tuple)) and len(v) > 1 else None


def load_global():
    return {
        "picking_types": {p["id"]: p for p in sr("stock.picking.type", [], ["id", "name", "code", "sequence_code"])},
        "companies": sr("res.company", [], ["id", "name"]),
        "warehouses": sr("stock.warehouse", [], ["id", "name", "code"]),
    }


G = load_global()
devoluciones.configurar(sr, id_of, ref, URL)
inventario.configurar(sr, id_of, ref, URL)
_loc_cache, _partner_cache = {}, {}


def servicio_sync():
    """Servicio de sincronizacion con Odoo (instancia unica, creada de forma perezosa)."""
    return sync_mov.ServicioSync.instancia(sr, id_of, ref, URL, log=print)


def locations(ids):
    todo = sorted(set(i for i in ids if i and i not in _loc_cache))
    for l in sr("stock.location", [("id", "in", todo)], ["id", "name", "complete_name", "usage"]):
        _loc_cache[l["id"]] = l
    return [_loc_cache[i] for i in ids if i in _loc_cache]


def partners(ids):
    todo = sorted(set(i for i in ids if i and i not in _partner_cache))
    if todo:
        for p in sr("res.partner", [("id", "in", todo)], ["id", "name", "vat", "ref"]):
            _partner_cache[p["id"]] = p
    return [_partner_cache[i] for i in ids if i in _partner_cache]


def _by_ids(model, ids, fields):
    ids = sorted({i for i in ids if i})
    if not ids:
        return {}
    return {r["id"]: r for r in sr(model, [("id", "in", ids)], fields)}


def _quants(product_id):
    """Quants (stock fisico) del producto en todas sus ubicaciones."""
    svc = servicio_sync()
    if svc.listo():
        return svc.quants_producto(product_id)
    return sr("stock.quant", [("product_id", "=", product_id)],
              ["location_id", "quantity", "reserved_quantity", "inventory_quantity"])


def _datos_movimientos(product_id):
    """Movimientos y capas de valorizacion del producto.

    Se usa el espejo local cuando la sincronizacion esta al dia y, si no,
    se consulta el ERP.  El formato de salida es identico en ambos casos.
    """
    svc = servicio_sync()
    if svc and svc.listo():
        movs = []
        for m in svc.movimientos_producto(product_id):
            movs.append({
                "id": m["id"], "date": m["fecha"], "name": m["nombre"],
                "quantity": m["cantidad"], "product_qty": m["product_qty"],
                "product_uom": m["uom_id"], "location_id": m["origen_id"],
                "location_dest_id": m["destino_id"], "state": m["estado"],
                "picking_id": m["picking_id"], "picking_type_id": m["picking_type_id"],
                "seq_picking": m["picking_seq"], "picking_code": m["picking_code"],
                "compra_line_id": m["compra_line_id"], "venta_line_id": m["venta_line_id"],
                "origen_retorno_id": m["origen_retorno_id"],
                "company_id": m["company_id"], "referencia": m["referencia"],
                "create_uid": m.get("usuario_id"),
            })
        capas = defaultdict(list)
        for c in svc.capas_producto(product_id):
            capas[c["move_id"]].append({"quantity": c["cantidad"], "value": c["valor"],
                                        "unit_cost": c["costo"]})
        return movs, dict(capas), "espejo local (SQLite)"

    moves = sr("stock.move", [("product_id", "=", product_id), ("state", "=", "done")],
               ["id", "name", "date", "quantity", "product_qty", "product_uom", "company_id",
                "location_id", "location_dest_id", "state", "picking_id", "picking_type_id",
                "purchase_line_id", "sale_line_id", "origin_returned_move_id", "origin",
                "create_uid"],
               order="date asc, id asc")
    movs = [{"id": m["id"], "date": m["date"], "name": m.get("name") or "",
             "quantity": m.get("quantity") or 0.0, "product_qty": m.get("product_qty"),
             "product_uom": id_of(m.get("product_uom")),
             "location_id": id_of(m.get("location_id")),
             "location_dest_id": id_of(m.get("location_dest_id")),
             "state": m.get("state"), "picking_id": id_of(m.get("picking_id")),
             "picking_type_id": id_of(m.get("picking_type_id")),
             "seq_picking": None, "picking_code": None,
             "compra_line_id": id_of(m.get("purchase_line_id")),
             "venta_line_id": id_of(m.get("sale_line_id")),
             "origen_retorno_id": id_of(m.get("origin_returned_move_id")),
             "company_id": id_of(m.get("company_id")),
             "referencia": m.get("origin") or "",
             "create_uid": id_of(m.get("create_uid"))} for m in moves]
    capas = defaultdict(list)
    for l in sr("stock.valuation.layer", [("product_id", "=", product_id)],
                ["stock_move_id", "quantity", "unit_cost", "value"]):
        capas[id_of(l["stock_move_id"])].append({"quantity": l.get("quantity") or 0.0,
                                                "value": l.get("value") or 0.0,
                                                "unit_cost": l.get("unit_cost") or 0.0})
    return movs, dict(capas), "consulta directa a Odoo"


def _enriquecer_movimientos(movs):
    """Agrega documento, picking, contraparte y responsable a cada movimiento."""
    if not movs:
        return movs
    picks = _by_ids("stock.picking", [m.get("picking_id") for m in movs],
                    ["id", "name", "partner_id", "origin", "state", "user_id"])
    pol = _by_ids("purchase.order.line", [m.get("compra_line_id") for m in movs],
                  ["id", "order_id"])
    sol = _by_ids("sale.order.line", [m.get("venta_line_id") for m in movs], ["id", "order_id"])
    pos = _by_ids("purchase.order", [id_of(v.get("order_id")) for v in pol.values()],
                  ["id", "name", "partner_id"])
    sos = _by_ids("sale.order", [id_of(v.get("order_id")) for v in sol.values()],
                  ["id", "name", "partner_id"])
    noms = {p["id"]: p for p in partners(
        [id_of(v.get("partner_id")) for v in list(picks.values()) + list(pos.values())
         + list(sos.values())])}
    # Quien realizo el movimiento: el responsable del traslado en Odoo
    # (stock.picking.user_id) y, si no hubiera, quien creo el movimiento.
    usuarios = _by_ids("res.users",
                        [m.get("create_uid") for m in movs] +
                        [id_of((picks.get(m.get("picking_id")) or {}).get("user_id"))
                         for m in movs],
                        ["id", "name", "login"])
    for m in movs:
        pk = picks.get(m.get("picking_id")) or {}
        po = pos.get(id_of((pol.get(m.get("compra_line_id")) or {}).get("order_id"))) or {}
        so = sos.get(id_of((sol.get(m.get("venta_line_id")) or {}).get("order_id"))) or {}
        m["picking_name"] = pk.get("name")
        doc = pk.get("name") or m.get("name") or m.get("referencia") or "N/A"
        if po.get("name"):
            doc += f" (OC {po['name']})"
        elif so.get("name"):
            doc += f" (Venta {so['name']})"
        m["doc"] = doc
        contraparte = ""
        if po.get("partner_id"):
            contraparte = noms.get(id_of(po.get("partner_id")), {}).get("name")
        elif so.get("partner_id"):
            contraparte = noms.get(id_of(so.get("partner_id")), {}).get("name")
        elif pk.get("partner_id"):
            contraparte = noms.get(id_of(pk.get("partner_id")), {}).get("name")
        m["contraparte"] = contraparte or ""
        usuario = (usuarios.get(id_of(pk.get("user_id")))
                   or usuarios.get(m.get("create_uid")) or {})
        m["usuario"] = usuario.get("name") or ""
    return movs


def build_kardex(company_id, product_id, refresh=False):
    """Arma el kardex de un producto.

    La clasificacion de entradas, salidas y transferencias la hace el motor de
    inventario (`app/inventario.py`) a partir de las ubicaciones de origen y
    destino y del alcance de la empresa (ubicaciones internas con
    company_id = empresa).  No se usa el codigo del tipo de picking para decidir
    el signo ni la existencia de capas de valorizacion.

    Reglas aplicadas (ver `inventario.clasificar`):
      * origen interno  -> destino interno : transferencia interna, efecto 0
        (PICK, PACK, OUT e INT no cuentan dos veces la misma mercaderia).
      * origen interno  -> destino externo : salida del almacen.
      * origen externo  -> destino interno : entrada al almacen.
      * origen externo  -> destino externo : sin efecto sobre el almacen.  Este
        es el caso de los ajustes y los transitos entre companias (p. ej.
        "Transito Inter-Companias" -> "Ajuste de Inventario"), que antes se
        contaban como entrada y descuadraban las existencias.
      * Todo traslado entre empresas (salida por una ubicacion interna y
        entrada por otra) se refleja una sola vez en cada empresa.

    Las cantidades se normalizan a la unidad base del producto y el saldo se
    concilia contra `stock.quant` de las mismas ubicaciones internas.

    La fuente de datos es el espejo local sincronizado cuando esta al dia y,
    en caso contrario, una consulta directa al ERP.
    """
    key = f"{company_id}:{product_id}"
    ahora = time.time()
    with _lock:
        entrada = _cache["kardex"].get(key)
        if entrada and not refresh and ahora - entrada[0] < CACHE_TTL_SEG:
            return entrada[1]

    prod = sr("product.product", [("id", "=", product_id)],
              ["id", "display_name", "name", "default_code", "uom_id", "categ_id",
               "product_tmpl_id", "standard_price", "type"])[0]
    tmpl = sr("product.template", [("id", "=", id_of(prod["product_tmpl_id"]))],
              ["cost_method", "valuation", "standard_price", "uom_id", "uom_po_id"])[0]
    # Preferir la UoM del producto, si no existe usar la del template
    if not prod.get("uom_id") and tmpl.get("uom_id"):
        prod["uom_id"] = tmpl["uom_id"]

    alm = inventario.almacenes(company_id)
    locs = inventario.ubicaciones()
    movimientos, capas, origen_datos = _datos_movimientos(product_id)
    movimientos = _enriquecer_movimientos(movimientos)
    quants = _quants(product_id)
    # Stock fisico de referencia: quants que estan dentro del alcance.  Son los
    # mismos que usa la conciliacion, de modo que la comparacion se hace sobre
    # el mismo universo de ubicaciones que el que genera el kardex.
    quants_alcance = [q for q in quants if id_of(q.get("location_id")) in alm]

    calc = inventario.calcular_kardex(company_id, prod, movimientos, capas, alm, locs,
                                      quants=quants)
    estado_sync = servicio_sync().estado()
    conciliacion = inventario.conciliar(company_id, prod, calc["kpi"]["saldo"], quants,
                                        movimientos=movimientos, capas=capas, alm=alm,
                                        locs=locs)
    conciliacion["producto_id"] = product_id
    conciliacion["ultima_sincronizacion"] = estado_sync.get("ultima_sincronizacion")

    payload = {
        "ok": True,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "empresa": {"id": company_id,
                    "nombre": next((c["name"] for c in G["companies"] if c["id"] == company_id), "N/A")},
        "producto": {"id": product_id, "codigo": prod["default_code"] or "N/A",
                     "nombre": prod["name"], "descripcion": prod["display_name"],
                     "unidad": calc["unidad"], "uom_id": calc["uom_id"],
                     "precision": calc["precision"],
                     "categoria": ref(prod["categ_id"]) or "N/A",
                     "costo_promedio_erp": prod["standard_price"] or 0.0,
                     "metodo": {"average": "Promedio ponderado (average) - real_time",
                                "fifo": "FIFO - real_time",
                                "standard": "Costo estándar"}.get(tmpl["cost_method"], tmpl["cost_method"]),
                     "valuacion": tmpl["valuation"]},
        "kpi": calc["kpi"],
        "rows": calc["rows"],
        "internals": calc["internals"],
        "externos": calc["externos"],
        "conciliacion": conciliacion,
        "alcance": {"empresa": company_id,
                    "ubicaciones_internas": len(alm),
                    "quants": len(quants_alcance)},
        "origen_datos": origen_datos,
        "sincronizacion": estado_sync,
    }
    if not calc["rows"]:
        payload.update({"ok": False,
                        "sin_movimientos": f"No se encontraron movimientos registrados para el ítem "
                                           f"{prod['default_code']} - {prod['name']} en el período "
                                           f"indicado para {payload['empresa']['nombre']}.",
                        "cancelados": 0, "total_moves": len(movimientos)})
    else:
        payload["desde"], payload["hasta"] = calc["rows"][0]["date"][:10], calc["rows"][-1]["date"][:10]
    payload["notas"] = [
        f"Fuente de datos: {origen_datos}. Movimientos completados (stock.move, estado done) "
        f"y capas de valorización (stock.valuation.layer) del producto {product_id}.",
        f"Unidad base del producto: {calc['unidad']} (precisión {calc['precision']}). Las "
        f"cantidades de los movimientos se convertieron a esta unidad con los factores "
        f"configurados en Odoo; las conversiones entre categorías incompatibles se rechazan.",
        f"Método de valuación del ERP: {payload['producto']['metodo']}; el valor contable de cada "
        f"movimiento corresponde a su capa de valorización. Las cantidades físicas NO se toman "
        f"de las capas.",
        f"Alcance: {len(alm)} ubicaciones internas de la empresa {payload['empresa']['nombre']}. "
        f"El signo de cada movimiento se determina por su ubicación de origen y destino "
        f"(entrada / salida / transferencia interna); los movimientos entre ubicaciones virtuales "
        f"(ajuste, tránsito, producción) no alteran el stock del almacén.",
        f"Transferencias internas excluidas del detalle: {calc['kpi']['internos']} movimientos "
        f"({calc['kpi']['internos_qty']:,.2f} {calc['unidad']}); se listan en la hoja secundaria del Excel.",
        f"Conciliación contra el stock físico de Odoo: {calc['kpi']['fisico']:,.2f} "
        f"{calc['unidad']} (quants del alcance) frente a {calc['kpi']['saldo']:,.2f} "
        f"{calc['unidad']} calculados por movimientos. Diferencia: "
        f"{conciliacion['diferencia']:,.2f} ({conciliacion['estado']}).",
    ] + (calc.get("notas_extra") or [])
    with _lock:
        _cache["kardex"][key] = (time.time(), payload)
    return payload


def _productos_conciliables(empresa, limite, svc):
    """Ids de los productos a conciliar y el origen de la seleccion.

    Con el espejo local al dia se toman los productos con mas movimientos.  Si
    el espejo aun no esta disponible (no se ha corrido la sincronizacion
    completa) se consultan directamente en Odoo los productos que tienen
    existencias en las ubicaciones internas de la empresa, de modo que el
    reporte nunca queda vacio por falta de espejo.
    """
    ids = svc.productos_con_movimientos(limite)
    if ids:
        return ids, "espejo local (SQLite)"
    alm = inventario.almacenes(empresa)
    quants = sr("stock.quant",
                [("location_id", "in", sorted(alm)), ("quantity", "!=", 0)],
                ["product_id", "quantity"])
    magnitudes = {}
    for q in quants or []:
        pid = id_of(q.get("product_id"))
        if pid:
            magnitudes[pid] = magnitudes.get(pid, 0.0) + (q.get("quantity") or 0.0)
    top = sorted(magnitudes.items(), key=lambda kv: -abs(kv[1]))[:limite]
    return [pid for pid, _ in top], "consulta directa a Odoo (stock.quant)"


def conciliacion_empresa(empresa, limite=50):
    """Concilia el saldo calculado por movimientos con el stock fisico de Odoo.

    Para cada producto se comparan, en la misma unidad base y sobre las
    mismas ubicaciones internas de la empresa, el stock fisico (quants) y el
    saldo que resulta de los movimientos completados.  No se modifica nunca
    el inventario del ERP.
    """
    svc = servicio_sync()
    alm = inventario.almacenes(empresa)
    locs = inventario.ubicaciones()
    ids, origen_productos = _productos_conciliables(empresa, limite, svc)
    prods = _by_ids("product.product", ids,
                    ["id", "default_code", "name", "uom_id", "standard_price"])
    salida = []
    for pid in ids:
        p = prods.get(pid)
        if not p:
            continue
        movs, capas, _ = _datos_movimientos(pid)
        movs = _enriquecer_movimientos(movs)
        quants = svc.quants_producto(pid, alm)
        calc = inventario.calcular_kardex(empresa, p, movs, capas, alm, locs, quants=quants)
        conc = inventario.conciliar(empresa, p, calc["kpi"]["saldo"], quants,
                                    movimientos=movs, capas=capas, alm=alm, locs=locs)
        conc["producto_id"] = pid
        conc["origen_productos"] = origen_productos
        conc["ultima_sincronizacion"] = svc.estado().get("ultima_sincronizacion")
        svc.guardar_conciliacion(empresa, conc)
        salida.append(conc)
    salida.sort(key=lambda c: abs(c["diferencia"]), reverse=True)
    return salida


def search_products(company_id, q, limit=40):
    q = (q or "").strip()
    if q:
        dom = ["|", ("default_code", "ilike", q), ("name", "ilike", q)]
    else:
        dom = [("type", "in", ["consu", "product"])]
    rows = sr("product.product", dom, ["id", "default_code", "name", "display_name"],
              order="default_code asc, name asc", limit=limit)
    out = []
    for r in rows:
        out.append({"id": r["id"], "codigo": r["default_code"] or f"ID{r['id']}",
                    "nombre": r["name"], "display": r["display_name"]})
    return out


def filter_rows(payload, f_from, f_to, tipo, q):
    rows = payload["rows"]
    q = (q or "").strip().lower()
    out = []
    for r in rows:
        d = r["date"][:10]
        if f_from and d < f_from:
            continue
        if f_to and d > f_to:
            continue
        if tipo and r["tipo"] != tipo:
            continue
        if q and q not in (r["doc"] + " " + r["origen"] + " " + r["tipo"] + " " +
                           (r.get("usuario") or "")).lower():
            continue
        out.append(r)
    return out


def filter_internals(payload, f_from, f_to, tipo, q):
    """Aplica los mismos filtros que filter_rows a las transferencias internas."""
    q = (q or "").strip().lower()
    out = []
    for i in payload["internals"]:
        d = i["date"][:10]
        if f_from and d < f_from:
            continue
        if f_to and d > f_to:
            continue
        if tipo and i["tipo"] != tipo:
            continue
        texto = " ".join(str(i.get(k) or "")
                         for k in ("doc", "ubicacion", "origen", "tipo", "usuario")).lower()
        if q and q not in texto:
            continue
        out.append(i)
    return out


HEAD_FILL = PatternFill("solid", fgColor="0F172A")
INT_FILL = PatternFill("solid", fgColor="F1F5F9")
INT_FONT = Font(size=10, color="64748B", italic=True)
TITLE_FONT = Font(bold=True, size=14, color="0F172A")
LBL = Font(bold=True, size=10, color="475569")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _celda_origen(texto, usuario):
    """Origen/destino del movimiento con el responsable en una segunda linea."""
    if usuario:
        return f"{texto}\nRealizado por: {usuario}"
    return texto


def build_xlsx(payload, rows, frm, incluir_internas=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Kardex"
    prod, k = payload["producto"], payload["kpi"]
    ws["A1"] = f"REPORTE KARDEX \u2014 {prod['codigo']} \u2014 {prod['nombre']}"
    ws["A1"].font = TITLE_FONT
    meta = [
        ("Empresa", payload["empresa"]["nombre"]),
        ("Unidad de medida", prod["unidad"]),
        ("Categor\u00eda", prod["categoria"]),
        ("M\u00e9todo de valuaci\u00f3n", prod["metodo"]),
        ("Per\u00edodo del reporte", (frm.get("desde") or payload.get("desde") or "todas las fechas") +
                                 " \u2013 " + (frm.get("hasta") or payload.get("hasta") or "hoy")),
        ("Filtros aplicados", frm.get("texto", "(sin filtros)")),
        ("Fuente de datos", payload.get("origen_datos", "")),
        ("Ultima sincronizacion", (payload.get("sincronizacion") or {}).get("ultima_sincronizacion") or "n/d"),
        ("Generado", payload["generado"]),
    ]
    r = 3
    for label, value in meta:
        ws.cell(row=r, column=1, value=label).font = LBL
        ws.cell(row=r, column=2, value=value)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="RESUMEN").font = Font(bold=True, size=12)
    r += 1
    kpis = [("Entradas totales", k["entradas"]), ("Salidas totales", k["salidas"]),
            ("Saldo final", k["saldo"]), ("Valor del saldo", k["valor"]),
            ("Costo promedio final", k["costo_prom"]), ("Movimientos", k["movimientos"]),
            ("Transferencias internas", k["internos"]),
            ("Stock fisico en Odoo", k["fisico"]), ("Disponible", k["disponible"]),
            ("Reservado", k["reservado"])]
    for i, (label, value) in enumerate(kpis):
        c = 1 + (i % 4) * 2
        rr = r + (i // 4) * 2
        ws.cell(row=rr, column=c, value=label).font = LBL
        cell = ws.cell(row=rr, column=c + 1, value=value)
        cell.font = Font(bold=True)
        if "Valor" in label or "Costo" in label:
            cell.number_format = '"$"#,##0.00'
        else:
            cell.number_format = "#,##0.00"
    r += len(kpis) // 4 * 2 + 2

    headers = ["Fecha", "Tipo Mvto.", "Num. Doc.", "Origen / Destino", "Entradas", "Salidas",
               "Costo Unit.", "Saldo", "Valor Total Saldo"]
    head_row = r
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=c, value=h)
        cell.fill = HEAD_FILL
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    int_rows = (filter_internals(payload, frm.get("desde", ""), frm.get("hasta", ""),
                                 frm.get("tipo", ""), frm.get("q", "")) if incluir_internas else [])
    items = [(mv["date"], str(mv.get("doc") or ""), "k", mv) for mv in rows]
    items += [(i["date"], str(i.get("doc") or ""), "i", i) for i in int_rows]
    items.sort(key=lambda t: (t[0], t[1]))

    tot_e = tot_s = 0.0
    ultimo_kardex = None
    for i, (fecha, _doc, clase, mv) in enumerate(items, start=1):
        rr = head_row + i
        f = fecha.split("-")[2] + "/" + fecha.split("-")[1] + "/" + fecha.split("-")[0]
        if clase == "k":
            vals = [f, mv["tipo"], mv["doc"], _celda_origen(mv["origen"], mv.get("usuario")),
                    mv["entrada"], mv["salida"], mv["cost"], mv["saldo"], mv["valor_total"]]
            tot_e += mv["entrada"]
            tot_s += mv["salida"]
            ultimo_kardex = mv
        else:
            vals = [f, "Transferencia interna", mv["doc"],
                    _celda_origen(" → ".join(x for x in (mv.get("ubicacion"), mv.get("origen"))
                                             if x), mv.get("usuario")),
                    None, None, None, None, None]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 4))
            if c in (5, 6, 8):
                cell.number_format = "#,##0.00"
            if c in (7, 9):
                cell.number_format = '"$"#,##0.00'
            if clase == "i":
                cell.fill = INT_FILL
                cell.font = INT_FONT
            elif i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")

    last = head_row + len(items)
    tr = last + 1
    resumen = f"{len(rows)} movimientos"
    if int_rows:
        resumen += f" + {len(int_rows)} transferencias internas (solo informativo)"
    ws.cell(row=tr, column=1, value="TOTALES").font = Font(bold=True)
    ws.cell(row=tr, column=4, value=resumen)
    ref = ultimo_kardex or {}
    for c, v in ((5, tot_e), (6, tot_s), (8, ref.get("saldo", 0)), (9, ref.get("valor_total", 0))):
        cell = ws.cell(row=tr, column=c, value=v)
        cell.font = Font(bold=True)
        cell.number_format = "#,##0.00" if c in (5, 6, 8) else '"$"#,##0.00'
    for c in range(1, 10):
        ws.cell(row=tr, column=c).fill = PatternFill("solid", fgColor="E2E8F0")
        ws.cell(row=tr, column=c).border = BORDER

    widths = [12, 24, 34, 46, 12, 12, 13, 12, 18]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    ws.auto_filter.ref = f"A{head_row}:I{last}"
    ws.row_dimensions[head_row].height = 30

    nr = tr + 2
    ws.cell(row=nr, column=1, value="NOTAS").font = Font(bold=True, size=12)
    notas = list(payload["notas"])
    if int_rows:
        notas.insert(0, f"Las {len(int_rows)} transferencias internas se muestran en gris solo para "
                        "referencia: no afectan el saldo, el costo promedio ni la valoracion del inventario.")
    for i, n in enumerate(notas, start=1):
        ws.cell(row=nr + i, column=1, value=f"{i}. {n}").alignment = Alignment(wrap_text=False)

    if payload["internals"]:
        w2 = wb.create_sheet("Transferencias internas")
        h2 = ["Fecha", "Tipo", "Num. Doc.", "Cantidad", "Ubicación destino", "Origen", "Usuario"]
        for c, h in enumerate(h2, start=1):
            cell = w2.cell(row=1, column=c, value=h)
            cell.fill = HEAD_FILL
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.border = BORDER
        for i, mv in enumerate(payload["internals"], start=2):
            f = mv["date"].split("-")[2] + "/" + mv["date"].split("-")[1] + "/" + mv["date"].split("-")[0]
            for c, v in enumerate([f, mv["tipo"], mv["doc"], mv["cantidad"], mv["ubicacion"],
                                   mv["origen"], mv.get("usuario") or ""], start=1):
                cell = w2.cell(row=i, column=c, value=v)
                cell.border = BORDER
                if c == 4:
                    cell.number_format = "#,##0.00"
        for c, w in enumerate([12, 10, 32, 12, 40, 22, 22], start=1):
            w2.column_dimensions[get_column_letter(c)].width = w
        w2.freeze_panes = "A2"
        w2.auto_filter.ref = f"A1:G{1 + len(payload['internals'])}"

    conc = payload.get("conciliacion")
    if conc:
        w3 = wb.create_sheet("Conciliacion")
        w3["A1"] = "CONCILIACION CONTRA EL STOCK FISICO DE ODOO"
        w3["A1"].font = TITLE_FONT
        filas = [("Producto", conc.get("codigo") or prod["codigo"]),
                 ("Descripcion", conc.get("descripcion") or prod["nombre"]),
                 ("Unidad", conc.get("unidad") or prod["unidad"]),
                 ("Stock fisico en Odoo", conc.get("stock_odoo")),
                 ("Saldo del kardex", conc.get("stock_kardex")),
                 ("Diferencia", conc.get("diferencia")),
                 ("Estado", conc.get("estado")),
                 ("Ultima sincronizacion", conc.get("ultima_sincronizacion") or "n/d")]
        for i, (label, value) in enumerate(filas, start=3):
            w3.cell(row=i, column=1, value=label).font = LBL
            cell = w3.cell(row=i, column=2, value=value)
            if isinstance(value, (int, float)):
                cell.number_format = "#,##0.00"
        r0 = len(filas) + 4
        w3.cell(row=r0, column=1, value="EXISTENCIAS POR UBICACION").font = Font(bold=True, size=12)
        cab = ["Ubicacion", "Existencia", "Reservado"]
        for c, h in enumerate(cab, start=1):
            cell = w3.cell(row=r0 + 1, column=c, value=h)
            cell.fill = HEAD_FILL
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.border = BORDER
        for i, q in enumerate(conc.get("detalle_quants") or [], start=r0 + 2):
            for c, v in enumerate([q.get("ubicacion"), q.get("cantidad"), q.get("reservado")], start=1):
                cell = w3.cell(row=i, column=c, value=v)
                cell.border = BORDER
                if c in (2, 3):
                    cell.number_format = "#,##0.00"
        causas = conc.get("posibles_causas") or []
        if causas:
            rc = r0 + 2 + len(conc.get("detalle_quants") or []) + 1
            w3.cell(row=rc, column=1, value="CAUSAS POSIBLES DE LA DIFERENCIA").font = Font(bold=True, size=12)
            for i, x in enumerate(causas, start=rc + 1):
                w3.cell(row=i, column=1, value=f"- {x}")
        for c, w in enumerate([34, 18, 16], start=1):
            w3.column_dimensions[get_column_letter(c)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_csv(payload, rows):
    head = ["Fecha", "Tipo Mvto.", "Num. Doc.", "Origen / Destino", "Entradas", "Salidas",
            "Costo Unit.", "Saldo", "Valor Total Saldo"]
    out = [",".join(head)]
    for mv in rows:
        d = mv["date"].split("-")
        f = f"{d[2]}/{d[1]}/{d[0]}"
        cells = [f, mv["tipo"], mv["doc"], _celda_origen(mv["origen"], mv.get("usuario")),
                 f"{mv['entrada']:.2f}", f"{mv['salida']:.2f}",
                 f"{mv['cost']:.4f}", f"{mv['saldo']:.2f}", f"{mv['valor_total']:.2f}"]
        out.append(",".join('"' + c.replace('"', '""') + '"' for c in cells))
    return ("\ufeff" + "\n".join(out)).encode("utf-8")


def build_despacho(empresa, filtro_id, d0, d1):
    """Tiempos de despacho por fase (picking/packing/output) en un rango de fechas.

    Puerto del reporte original /TiempoDespachoOutput/ de la app gestion-comercial.
    d0 y d1 son datetime (desde y hasta, inclusive).
    """
    inicio = f"{d0:%Y-%m-%d} 05:00:00"
    if d0.date() == d1.date():
        dn = d1 + timedelta(days=1)
        fin = f"{dn:%Y-%m-%d} 04:59:59"
    else:
        fin = f"{d1:%Y-%m-%d} 04:59:59"

    vacio = {"ok": True, "user_id_filtro_aplicado": filtro_id,
             "kpi_promedios_generales": {"promedio_picking": "0h 0m",
                                         "promedio_packing": "0h 0m",
                                         "promedio_output": "0h 0m"},
             "total_pedidos_analizados": 0, "detalles_por_pedido": [],
             "mensaje": "No se encontraron movimientos de inventario en este rango de fechas."}

    pickings = sr("stock.picking",
                  [("state", "=", "done"), ("company_id", "=", empresa),
                   ("date_done", ">=", inicio), ("date_done", "<=", fin)],
                  ["sale_id"])
    sale_ids = sorted({id_of(p.get("sale_id")) for p in pickings
                       if id_of(p.get("sale_id"))})
    if not sale_ids:
        return vacio

    bulk = sr("stock.picking",
              [("sale_id", "in", sale_ids), ("state", "=", "done")],
              ["id", "name", "picking_type_id", "date_done", "user_id", "create_date",
               "sale_id", "partner_id"])

    # Conteo de items por picking. Se consulta stock.move en lugar de traer el
    # one2many move_ids_without_package: con miles de pedidos esa respuesta
    # XML-RPC llega truncada y el reporte falla.
    conteo_items = {}
    picking_ids = [p["id"] for p in bulk]
    for i in range(0, len(picking_ids), 500):
        for m in sr("stock.move", [("picking_id", "in", picking_ids[i:i + 500])],
                    ["picking_id"]):
            pid = id_of(m.get("picking_id"))
            conteo_items[pid] = conteo_items.get(pid, 0) + 1

    por_pedido = {}
    for p in bulk:
        sid = id_of(p.get("sale_id"))
        if sid:
            por_pedido.setdefault(sid, []).append(p)

    fmt = "%Y-%m-%d %H:%M:%S"
    resultados = []
    mapa_fechas_fin = {}

    def mins(a, b):
        return max(0, int((datetime.strptime(b, fmt) - datetime.strptime(a, fmt)).total_seconds() // 60))

    def legible(m):
        h, mm = divmod(m, 60)
        return f"{h}h {mm}m"

    for s_id in sale_ids:
        ops = por_pedido.get(s_id, [])
        if not ops:
            continue
        fases = {"picking": None, "packing": None, "output": None}
        nombre_pedido = f"ID: {s_id}"
        nombre_cliente = "Cliente no especificado"
        incluir = False
        n_items = 0
        n_conting = 0

        for p in ops:
            if isinstance(p.get("sale_id"), (list, tuple)) and len(p["sale_id"]) > 1:
                nombre_pedido = p["sale_id"][1]
            if isinstance(p.get("partner_id"), (list, tuple)) and len(p["partner_id"]) > 1:
                nombre_cliente = p["partner_id"][1]
            operacion = (ref(p.get("picking_type_id")) or "").lower()
            resp_id = id_of(p.get("user_id")) or 0
            resp_nom = ref(p.get("user_id")) or "No asignado"
            if filtro_id == 0 or resp_id == filtro_id:
                incluir = True
            datos_fase = {"documento": p.get("name", ""), "responsable": resp_nom,
                          "inicio_utc": p.get("create_date"), "fin_utc": p.get("date_done")}
            n_moves = conteo_items.get(p.get("id"), 0)
            if n_moves:
                if "pick" in operacion:
                    n_items = n_moves
                elif n_conting == 0:
                    n_conting = n_moves
            if "pick" in operacion:
                fases["picking"] = datos_fase
            elif any(x in operacion for x in ("pack", "emp", "prep")):
                fases["packing"] = datos_fase
            elif any(x in operacion for x in ("out", "deliv", "salida", "entrega")):
                fases["output"] = datos_fase

        if not incluir:
            continue
        if n_items == 0:
            n_items = n_conting

        rep = {"pedido_id": s_id, "pedido_nombre": nombre_pedido, "cliente": nombre_cliente,
               "total_items": int(n_items), "tiempos_fases_minutos": {},
               "tiempos_fases_legible": {}, "detalles_operacion": {}}

        if fases["picking"]:
            m = mins(fases["picking"]["inicio_utc"], fases["picking"]["fin_utc"])
            rep["tiempos_fases_minutos"]["picking"] = m
            rep["tiempos_fases_legible"]["picking"] = legible(m)
            rep["detalles_operacion"]["picking"] = {"documento": fases["picking"]["documento"],
                                                    "responsable": fases["picking"]["responsable"],
                                                    "fecha": fases["picking"]["fin_utc"][:10]}

        if fases["packing"] and fases["picking"]:
            m = mins(fases["picking"]["fin_utc"], fases["packing"]["fin_utc"])
            rep["tiempos_fases_minutos"]["packing"] = m
            rep["tiempos_fases_legible"]["packing"] = legible(m)
            rep["detalles_operacion"]["packing"] = {"documento": fases["packing"]["documento"],
                                                    "responsable": fases["packing"]["responsable"],
                                                    "fecha": fases["packing"]["fin_utc"][:10]}

        if fases["output"]:
            origen = fases["packing"] or fases["picking"]
            if origen:
                m = mins(origen["fin_utc"], fases["output"]["fin_utc"])
                rep["tiempos_fases_minutos"]["output"] = m
                rep["tiempos_fases_legible"]["output"] = legible(m)
                rep["detalles_operacion"]["output"] = {"documento": fases["output"]["documento"],
                                                       "responsable": fases["output"]["responsable"],
                                                       "fecha": fases["output"]["fin_utc"][:10]}

        if rep["tiempos_fases_minutos"]:
            resultados.append(rep)
            fase_final = "output" if "output" in rep["detalles_operacion"] else (
                "packing" if "packing" in rep["detalles_operacion"] else "picking")
            mapa_fechas_fin[s_id] = datetime.strptime(fases[fase_final]["fin_utc"], fmt)

    if not resultados:
        return vacio

    # tiempos muertos / transicion entre pedidos de cada despachador
    jornadas = {}
    for r in resultados:
        det = r["detalles_operacion"]
        fase_final = "output" if "output" in det else ("packing" if "packing" in det else "picking")
        nombre = det[fase_final]["responsable"]
        jornadas.setdefault(nombre, []).append(
            {"pedido": r, "fecha_fin": mapa_fechas_fin.get(r["pedido_id"], datetime.min)})

    for lista in jornadas.values():
        lista.sort(key=lambda x: x["fecha_fin"])
        for i, item in enumerate(lista):
            pedido = item["pedido"]
            if i == 0:
                pedido["tiempo_espera_previo_minutos"] = 0
                pedido["tiempo_espera_previo_legible"] = "Inicio de jornada"
            else:
                brecha = max(0, int((item["fecha_fin"] - lista[i - 1]["fecha_fin"]).total_seconds() // 60))
                pedido["tiempo_espera_previo_minutos"] = brecha
                pedido["tiempo_espera_previo_legible"] = legible(brecha)

    def promedio(fase):
        vals = [r["tiempos_fases_minutos"][fase] for r in resultados
                if fase in r["tiempos_fases_minutos"]]
        return legible(int(sum(vals) / len(vals))) if vals else "0h 0m"

    return {"ok": True, "user_id_filtro_aplicado": filtro_id,
            "kpi_promedios_generales": {"promedio_picking": promedio("picking"),
                                        "promedio_packing": promedio("packing"),
                                        "promedio_output": promedio("output")},
            "total_pedidos_analizados": len(resultados),
            "detalles_por_pedido": resultados}


def build_despacho_xlsx(empresa, filtro_id, d0, d1, responsable=""):
    """Exporta el dashboard de tiempos de despacho (misma salida que la app externa)."""
    rep = build_despacho(empresa, filtro_id, d0, d1)
    pedidos = list(rep.get("detalles_por_pedido", []))
    responsable = (responsable or "").strip()

    def resp_de(p):
        det = p.get("detalles_operacion", {})
        for fase in ("output", "packing", "picking"):
            v = (det.get(fase) or {}).get("responsable")
            if v:
                return v
        return "Sin asignar"

    def first(p, key):
        det = p.get("detalles_operacion", {})
        for fase in ("output", "packing", "picking"):
            v = (det.get(fase) or {}).get(key)
            if v:
                return v
        return "\u2014"

    if responsable:
        pedidos = [p for p in pedidos if resp_de(p) == responsable]

    wb = Workbook()
    ws = wb.active
    ws.title = "Despacho"
    ws["A1"] = f"TIEMPOS DE DESPACHO \u2014 {d0:%d/%m/%Y} \u2013 {d1:%d/%m/%Y}"
    ws["A1"].font = TITLE_FONT
    headers = ["Pedido", "Cliente", "Responsable", "Documento", "Items",
               "Picking (min)", "Picking", "Packing (min)", "Packing",
               "Output (min)", "Output", "Espera Previo (min)", "Espera Previo"]
    head_row = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=c, value=h)
        cell.fill = HEAD_FILL
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for i, p in enumerate(pedidos, start=1):
        mn = p.get("tiempos_fases_minutos", {})
        lg = p.get("tiempos_fases_legible", {})
        vals = [p.get("pedido_nombre"), p.get("cliente"), first(p, "responsable"),
                first(p, "documento"), p.get("total_items", 0),
                mn.get("picking", 0), lg.get("picking", "\u2014"),
                mn.get("packing", 0), lg.get("packing", "\u2014"),
                mn.get("output", 0), lg.get("output", "\u2014"),
                p.get("tiempo_espera_previo_minutos", 0),
                p.get("tiempo_espera_previo_legible", "\u2014")]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=head_row + i, column=c, value=v)
            cell.border = BORDER
            if c in (6, 8, 10, 12):
                cell.number_format = "#,##0"
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
    widths = [16, 30, 20, 18, 8, 13, 10, 13, 10, 13, 10, 17, 14]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    ws.auto_filter.ref = f"A{head_row}:M{head_row + len(pedidos)}"
    ws.row_dimensions[head_row].height = 28

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "KardexFimas/1.0"

    def log_message(self, *a):
        pass

    # ---------- utilidades ----------
    def _send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, code, data, headers=None):
        self._send(code, json.dumps(data, ensure_ascii=False), headers=headers)

    def _html(self, name):
        path = os.path.join(HERE, name)
        if not os.path.exists(path):
            return self._send(404, "Pagina no encontrada", "text/plain; charset=utf-8")
        with open(path, "rb") as fh:
            self._send(200, fh.read(), "text/html; charset=utf-8")

    def _redirect(self, destino):
        self.send_response(302)
        self.send_header("Location", destino)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _cookie_token(self):
        raw = self.headers.get("Cookie")
        if not raw:
            return None
        try:
            jar = SimpleCookie()
            jar.load(raw)
        except Exception:
            return None
        m = jar.get(odoo_auth.COOKIE)
        return m.value if m else None

    def _usuario(self):
        return odoo_auth.usuario_de_sesion(self._cookie_token())

    def _guardar_cookie(self, token):
        return (f"{odoo_auth.COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax; "
                f"Max-Age={odoo_auth.SESION_HORAS * 3600}")

    def _borrar_cookie(self):
        return f"{odoo_auth.COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"

    def _requiere(self):
        """Devuelve (usuario, None) o (None, respuesta 401/302 ya enviada)."""
        u = self._usuario()
        if u:
            return u, None
        if self.path.startswith("/api/"):
            self._json(401, {"error": "no_autorizado", "mensaje": MSG_LOGIN})
        else:
            self._redirect("/login")
        return None, True

    def _requiere_admin(self):
        """Como _requiere, pero exige rol admin en Odoo."""
        u, cortado = self._requiere()
        if cortado:
            return None, True
        if u.get("rol") != "admin":
            self._json(403, {"error": "sin_permiso", "mensaje": MSG_NO_ADMIN})
            return None, True
        return u, False

    def _leer_json(self):
        largo = int(self.headers.get("Content-Length") or 0)
        if not largo:
            return {}
        try:
            return json.loads(self.rfile.read(largo).decode("utf-8")) or {}
        except (ValueError, UnicodeDecodeError):
            return {}

    # ---------- GET ----------
    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        one = lambda k, d=None: (q.get(k) or [d])[0]
        try:
            if u.path in ("/login", "/login.html"):
                return self._html("login.html")
            if u.path == "/api/auth/estado":
                return self._json(200, {"autenticado": bool(self._usuario()),
                                        "requiere_instalacion": False})
            if u.path == "/api/auth/me":
                usr = self._usuario()
                if not usr:
                    return self._json(401, {"error": "no_autorizado", "mensaje": MSG_LOGIN})
                return self._json(200, {"usuario": {**usr,
                                                    "modulos": permisos.obtener(usr["id"])}})

            usr, cortado = self._requiere()
            if cortado:
                return
            if u.path in ("/", "/index.html", "/kardex", "/kardex_app.html"):
                return self._html("kardex_app.html")
            if u.path in ("/permisos", "/permisos.html"):
                if usr.get("rol") != "admin":
                    return self._redirect("/")
                return self._html("permisos.html")
            if u.path in ("/configuracion", "/configuracion.html"):
                if usr.get("rol") != "admin":
                    return self._redirect("/")
                return self._html("configuracion.html")
            if u.path in ("/despacho", "/despacho.html"):
                return self._html("despacho.html")
            if u.path in ("/devoluciones", "/devoluciones.html"):
                if not permisos.puede(usr["id"], "devoluciones"):
                    nombre = permisos.modulos_activos().get(
                        "devoluciones", "Devoluciones pendientes de compras puntuales")
                    return self._redirect(f"/?sin_permiso={nombre}")
                return self._html("devoluciones.html")
            if u.path in ("/usuarios", "/usuarios.html"):
                return self._redirect("/")
            if u.path == "/api/companies" and not (
                    permisos.puede(usr["id"], "kardex") or permisos.puede(usr["id"], "despacho")):
                nombre = permisos.modulos_activos().get("kardex", "Kardex")
                return self._json(403, {"error": "sin_permiso",
                                        "mensaje": f"Lo siento, no tienes permiso para acceder "
                                                   f"al m\u00f3dulo {nombre}."})
            if u.path in ("/api/products", "/api/kardex",
                          "/api/export", "/api/export/csv", "/api/sync/estado",
                          "/api/sync/errores", "/api/conciliacion") and not permisos.puede(usr["id"], "kardex"):
                nombre = permisos.modulos_activos().get("kardex", "Kardex")
                return self._json(403, {"error": "sin_permiso",
                                        "mensaje": f"Lo siento, no tienes permiso para acceder "
                                                   f"al m\u00f3dulo {nombre}."})
            if u.path in ("/api/despacho", "/api/despachadores", "/api/despacho/export") and not (
                    permisos.puede(usr["id"], "despacho")):
                nombre = permisos.modulos_activos().get("despacho", "Despacho")
                return self._json(403, {"error": "sin_permiso",
                                        "mensaje": f"Lo siento, no tienes permiso para acceder "
                                                   f"al m\u00f3dulo {nombre}."})
            if u.path in ("/api/devoluciones", "/api/devoluciones/export",
                          "/api/devoluciones/detalle") and not (
                    permisos.puede(usr["id"], "devoluciones")):
                nombre = permisos.modulos_activos().get(
                    "devoluciones", "Devoluciones pendientes de compras puntuales")
                return self._json(403, {"error": "sin_permiso",
                                        "mensaje": f"Lo siento, no tienes permiso para acceder "
                                                   f"al m\u00f3dulo {nombre}."})
            if u.path == "/api/permisos":
                if usr.get("rol") != "admin":
                    return self._json(403, {"error": "sin_permiso", "mensaje": MSG_NO_ADMIN})
                gid = odoo_auth.group_system_id()
                filas = sr("res.users", [("active", "=", True)],
                           ["id", "name", "login", "email", "groups_id"], order="name asc")
                usuarios = [{"id": r["id"], "nombre": r["name"], "login": r["login"],
                             "email": r["email"] or "",
                             "rol": "admin" if gid and gid in (r.get("groups_id") or [])
                                    else "usuario_estándar",
                             "modulos": permisos.obtener(r["id"])} for r in filas]
                return self._json(200, {"modulos": permisos.listar_modulos(),
                                        "usuarios": usuarios})
            if u.path == "/api/companies":
                return self._json(200, G["companies"])
            if u.path == "/api/config":
                if usr.get("rol") != "admin":
                    return self._json(403, {"error": "sin_permiso", "mensaje": MSG_NO_ADMIN})
                return self._json(200, {"ok": True, "sistema": "Fimas",
                                        "origen": "Odoo ERP \u00b7 tiempo real",
                                        "url_odoo": URL, "base_datos": DB,
                                        "companias": [{"id": c["id"], "nombre": c["name"]}
                                                      for c in G["companies"]],
                                        "modulos": permisos.modulos_activos()})
            if u.path == "/api/products":
                return self._json(200, search_products(one("company_id"), one("q")))
            if u.path == "/api/kardex":
                cid, pid = int(one("company_id")), int(one("product_id"))
                refrescar = one("refresh", "0") in ("1", "true", "si")
                return self._json(200, build_kardex(cid, pid, refresh=refrescar))
            if u.path == "/api/sync/estado":
                return self._json(200, servicio_sync().estado())
            if u.path == "/api/sync/errores":
                return self._json(200, {"errores": servicio_sync().errores(
                    int(one("limite", "50") or 50))})
            if u.path == "/api/conciliacion":
                cid = int(one("company_id", "0") or 0)
                if not cid:
                    return self._json(400, {"error": "parametros",
                                            "mensaje": "Indique la compa\u00f1eda (company_id)."})
                filas = conciliacion_empresa(cid, int(one("limite", "50") or 50))
                return self._json(200, {"empresa": cid, "conciliaciones": filas})
            if u.path == "/api/despachadores":
                filas = sr("res.users", [("active", "=", True)], ["id", "name"], order="name asc")
                return self._json(200, [{"id": r["id"], "nombre": r["name"]} for r in filas])
            if u.path == "/api/despacho":
                try:
                    d0 = datetime.strptime(one("desde", ""), "%Y-%m-%d")
                    d1 = datetime.strptime(one("hasta", ""), "%Y-%m-%d")
                except ValueError:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "Indique el rango de fechas (desde / hasta) "
                                                       "en formato YYYY-MM-DD."})
                if d1 < d0:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "La fecha 'hasta' no puede ser menor "
                                                       "que la fecha 'desde'."})
                empresa = int(one("empresa", "0") or 0)
                if not empresa:
                    return self._json(400, {"error": "empresa",
                                            "mensaje": "Seleccione una compa\u00f1\u00eda."})
                user_id = int(one("user_id", "0") or 0)
                return self._json(200, build_despacho(empresa, user_id, d0, d1))
            if u.path == "/api/despacho/export":
                try:
                    d0 = datetime.strptime(one("desde", ""), "%Y-%m-%d")
                    d1 = datetime.strptime(one("hasta", ""), "%Y-%m-%d")
                except ValueError:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "Indique el rango de fechas (desde / hasta) "
                                                       "en formato YYYY-MM-DD."})
                if d1 < d0:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "La fecha 'hasta' no puede ser menor "
                                                       "que la fecha 'desde'."})
                empresa = int(one("empresa", "0") or 0)
                if not empresa:
                    return self._json(400, {"error": "empresa",
                                            "mensaje": "Seleccione una compa\u00f1\u00eda."})
                user_id = int(one("user_id", "0") or 0)
                data = build_despacho_xlsx(empresa, user_id, d0, d1, one("responsable", ""))
                fname = f"despacho_{datetime.now():%Y%m%d_%H%M}.xlsx"
                return self._send(200, data,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{fname}"'})
            if u.path in ("/api/devoluciones", "/api/devoluciones/export",
                          "/api/devoluciones/detalle"):
                desde, hasta = one("desde", ""), one("hasta", "")
                try:
                    if desde:
                        datetime.strptime(desde, "%Y-%m-%d")
                    if hasta:
                        datetime.strptime(hasta, "%Y-%m-%d")
                except ValueError:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "Indique el rango de fechas (desde / hasta) "
                                                       "en formato YYYY-MM-DD."})
                if desde and hasta and hasta < desde:
                    return self._json(400, {"error": "fechas",
                                            "mensaje": "La fecha 'hasta' no puede ser menor "
                                                       "que la fecha 'desde'."})
                if u.path == "/api/devoluciones/detalle":
                    mid = int(one("move_id", "0") or 0)
                    if not mid:
                        return self._json(400, {"error": "parametros",
                                                "mensaje": "Indique el movimiento de devolucion."})
                    return self._json(200, devoluciones.detalle(devoluciones.COMPANY_FINPAC, mid))
                data = devoluciones.build(devoluciones.COMPANY_FINPAC, desde, hasta,
                                          refresh=one("refresh", "0") in ("1", "true", "si"))
                if u.path == "/api/devoluciones/export":
                    fname = f"Devoluciones_puntuales_FINPAC_{datetime.now():%Y%m%d_%H%M}.xlsx"
                    return self._send(200, devoluciones.build_xlsx(data),
                                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": f'attachment; filename="{fname}"'})
                return self._json(200, data)
            if u.path in ("/api/export", "/api/export/csv"):
                cid, pid = int(one("company_id")), int(one("product_id"))
                payload = build_kardex(cid, pid, refresh=True)
                if not payload.get("ok"):
                    return self._json(404, {"error": "sin_movimientos",
                                            "mensaje": payload.get("sin_movimientos", "Sin movimientos")})
                frm = {"desde": one("desde", ""), "hasta": one("hasta", ""),
                       "tipo": one("tipo", ""), "q": one("q", ""),
                       "texto": (one("q", "") or "(sin filtros)").upper(),
                       "internas": one("internas", "0") in ("1", "true", "si")}
                rows = filter_rows(payload, frm["desde"], frm["hasta"], frm["tipo"], one("q", ""))
                fname = f"Kardex_{payload['producto']['codigo']}_{datetime.now():%Y%m%d_%H%M}.xlsx"
                if u.path == "/api/export/csv":
                    return self._send(200, build_csv(payload, rows), "text/csv; charset=utf-8",
                                      {"Content-Disposition": f'attachment; filename="{fname[:-5]}.csv"'})
                return self._send(200, build_xlsx(payload, rows, frm, frm["internas"]),
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{fname}"'})
            return self._json(404, {"error": "ruta_no_encontrada"})
        except Exception:
            traceback.print_exc()
            self._json(500, {"error": "error_interno"})

    # ---------- POST ----------
    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        try:
            if u.path == "/api/auth/login":
                data = self._leer_json()
                res = odoo_auth.autenticar(URL, DB, data.get("email"), data.get("password"),
                                           USER, PWD)
                if not res.get("ok"):
                    if res.get("codigo") in ("odoo_error", "odoo_offline"):
                        print(f"[login] odoo no disponible: {res.get('codigo')}")
                        return self._json(502, {"error": "odoo_offline",
                                                "mensaje": "No fue posible consultar el ERP (Odoo). "
                                                           "Intente nuevamente en unos minutos."})
                    print(f"[login] fallo: {res.get('codigo')}")
                    return self._json(401, {"error": "credenciales", "mensaje": MSG_LOGIN})
                token = odoo_auth.crear_sesion(res["usuario"])
                u = res["usuario"]
                print(f"[login] ok: {u['login']} (odoo uid {u['id']})")
                return self._json(200, {"ok": True, "usuario": u},
                                  {"Set-Cookie": self._guardar_cookie(token)})
            if u.path == "/api/auth/logout":
                odoo_auth.cerrar_sesion(self._cookie_token())
                return self._json(200, {"ok": True}, {"Set-Cookie": self._borrar_cookie()})
            if u.path == "/api/sync/sincronizar":
                if not permisos.puede(usr["id"], "kardex"):
                    return self._json(403, {"error": "sin_permiso",
                                            "mensaje": "No tiene permiso sobre el m\u00f3dulo Kardex."})
                data = self._leer_json()
                completo = bool(data.get("completo"))
                pid = data.get("producto_id")
                if completo and usr.get("rol") != "admin":
                    return self._json(403, {"error": "sin_permiso",
                                            "mensaje": "Solo un administrador puede lanzar la "
                                                       "sincronizaci\u00f3n completa."})
                res = servicio_sync().lanzar(completo=completo, motivo="manual",
                                             producto_id=int(pid) if pid else None)
                return self._json(200 if res.get("ok") else 409, res)
            if u.path.rstrip("/") == "/api/permisos/modulos":
                _, cortado = self._requiere_admin()
                if cortado:
                    return
                data = self._leer_json()
                res = permisos.crear_modulo(data.get("clave"), data.get("nombre"),
                                            data.get("descripcion"))
                return self._json(200 if res.get("ok") else 400, res)
            return self._json(404, {"error": "ruta_no_encontrada"})
        except Exception:
            traceback.print_exc()
            self._json(500, {"error": "error_interno"})

    # ---------- PUT / DELETE ----------
    def do_PUT(self):
        u = urllib.parse.urlparse(self.path)
        try:
            usr, cortado = self._requiere_admin()
            if cortado:
                return
            data = self._leer_json()
            if u.path.rstrip("/") == "/api/permisos":
                res = permisos.actualizar(data.get("odoo_uid"), data.get("modulos"))
                return self._json(200 if res.get("ok") else 400, res)
            partes = u.path.rstrip("/").split("/")
            if len(partes) == 5 and partes[1:4] == ["api", "permisos", "modulos"]:
                res = permisos.actualizar_modulo(partes[4], data.get("nombre"),
                                                 data.get("descripcion"), data.get("activo"))
                return self._json(200 if res.get("ok") else 400, res)
            return self._json(404, {"error": "ruta_no_encontrada"})
        except Exception:
            traceback.print_exc()
            self._json(500, {"error": "error_interno"})

    def do_DELETE(self):
        u = urllib.parse.urlparse(self.path)
        try:
            _, cortado = self._requiere_admin()
            if cortado:
                return
            partes = u.path.rstrip("/").split("/")
            if len(partes) == 5 and partes[1:4] == ["api", "permisos", "modulos"]:
                res = permisos.eliminar_modulo(partes[4])
                return self._json(200 if res.get("ok") else 400, res)
            return self._json(404, {"error": "ruta_no_encontrada"})
        except Exception:
            traceback.print_exc()
            self._json(500, {"error": "error_interno"})


if __name__ == "__main__":
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    import socket
    try:
        _s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        _s.connect(("8.8.8.8", 80))
        lan_ip = _s.getsockname()[0]
        _s.close()
    except Exception:
        lan_ip = "127.0.0.1"
    print(f"Kardex Fimas en http://localhost:{PORT}/")
    print(f"  red local: http://{lan_ip}:{PORT}/  (host de arranque: {HOST})")
    print(f"  empresas ERP: {[c['name'] for c in G['companies']]}")
    print("  ingreso: cuentas de Odoo (XML-RPC), sesiones locales de 8 h")

    # Servicio de sincronizacion con Odoo: hilo unico, en segundo plano.
    svc = servicio_sync()
    svc.arrancar()
    print(f"  sincronizacion: activa cada {svc.intervalo}s | espejo local: {svc.ruta}")

    # Verificacion del canal de notificaciones de Odoo (no bloquea el arranque).
    def _verificar_bus():
        try:
            info = servicio_sync().verificar_bus(USER, PWD, DB)
            if info.get("disponible"):
                print("  bus de Odoo: disponible")
            else:
                print("  bus de Odoo: no disponible -> sincronizacion periodica cada "
                      f"{svc.intervalo}s")
                for d in info.get("detalle", []):
                    print(f"      - {d}")
        except Exception as exc:                      # pragma: no cover - defensivo
            print(f"  bus de Odoo: verificacion fallida ({exc})")

    threading.Thread(target=_verificar_bus, daemon=True, name="kardex-bus").start()

    srv.serve_forever()