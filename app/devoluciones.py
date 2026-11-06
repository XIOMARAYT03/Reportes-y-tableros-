"""Devoluciones pendientes de compras puntuales - FINPAC.

Identifica los productos de compras PUNTUALES/LOCALES de FINPAC que un cliente
devolvio y que a la fecha siguen en el inventario sin venderse.

Modulo de solo lectura sobre Odoo: NO escribe en el ERP, NO genera cobros ni
descuentos, y NO modifica el Kardex existente. Reutiliza la conexion XML-RPC y
la logica de inventario del servidor por inyeccion (evita import circular).

Como se determina una "compra puntual" (verificable en Odoo)
------------------------------------------------------------
Una devolucion solo entra al reporte si la compra que probo el producto es una
compra puntual/local de FINPAC. La compra se localiza por tres vias verificables
y se acepta la primera que coincida con el producto devuelto:

  1. Via documental directa: la linea de venta original tiene purchase_line_ids
     (campo propio del ERP) que apuntan a las lineas de la orden de compra.
     Se revisan TODAS las lineas de la orden de venta que venden el mismo
     producto: la compra puntual puede estar registrada en la linea hermana.
  2. Via documental de la OC citada: el movimiento de salida original (el que
     el cliente recibio y luego devolvio) cita el codigo de la OC en su
     documento origen, p.ej. "P14563,S25196 - 260015499/TUB137: Salida>Clientes".
  3. Via trazabilidad FIFO: la OC cuya capa fue consumida por la salida original
     (la que entrego las unidades al cliente que devuelve).

El vinculo usado se informa en el campo "vinculo" de cada fila:
  "orden_venta" o "documento" = vinculo documental del ERP;
  "trazabilidad" = la OC se dedujo de la traza FIFO de las capas.

La compra se descarta cuando:
  - esta marcada como importacion (purchase.order.is_import = True), o
  - pertenece a otra empresa (el mismo codigo de OC existe en FINPAC y en
    MASERVEN: solo se usa la OC de FINPAC), o
  - su proveedor es una empresa del grupo (MASERVEN): son compras internas /
    traspasos entre empresas, no compras locales de FINPAC, o
  - no esta vinculada a la orden de venta original (la OC no la cita en origin
    ni en group_id) y la linea de venta no usa una ruta comercial puntual
    ("Compra Puntual (En Oficina)", "Cross-Dock").


Como se calcula la cantidad pendiente
--------------------------------------
Odoo no permite identificar fisicamente las unidades devueltas, asi que se usa
una trazabilidad cronologica FIFO por CAPAS sobre los movimientos reales:

  - Cada entrada (recepcion de compra, ajuste positivo, devolucion de cliente)
    abre una capa de stock con la fecha del movimiento.
  - Cada salida (venta a cliente, ajuste de inventario, devolucion a proveedor,
    traspaso a otra empresa) consume las capas mas ANTIGUAS primero (FIFO).
  - Los traslados internos (PICK/PACK/PRCH/INT: interno -> interno) NO son
    ventas: no abren ni consumen capas.
  - La capa de la devolucion conserva su saldo mientras no sea consumida: esa
    es la CANTIDAD PENDIENTE. Si otro vendedor vendio esas unidades despues,
    su salida consume la capa de la devolucion y la pendiente baja. Si se
    vendieron todas, la devolucion no aparece en el reporte.
  - Las compras posteriores del mismo codigo abren capas nuevas: el stock
    guardado despues NUNCA se atribuye a la devolucion original.

El pool se siembra con el stock inicial no trazado (stock vivo de las
ubicaciones internas de FINPAC menos el delta de los movimientos) de modo que
la traza sea consistente con el inventario real del ERP.
"""
import io
import re
import threading
import time
from collections import defaultdict, deque
from datetime import datetime

# ---- Configuracion / acceso a Odoo inyectados por kardex_server ----
_sr = None
_id_of = None
_ref = None
_odoo_url = None
_lock = threading.RLock()
_cache = {}
_cache_reporte = {}

# El reporte es pesado (trazabilidad FIFO de todo el historial de los productos):
# se cachea unos minutos para que el usuario no espere en cada consulta.
CACHE_REPORTE_SEG = 300

COMPANY_FINPAC = 1
COMPANY_FINPAC_NOMBRE = "FINPAC"
NO_ID = "No identificado"
EPS = 1e-6

# Rutas comerciales que identifican una compra puntual/local en el ERP.
RUTAS_PUNTUALES = ("puntual", "cross-dock", "cross dock")

# Codigo de OC local de FINPAC (P12345) citado en los documentos de origen.
_RE_OC = re.compile(r"\bP\d{4,6}\b")


def configurar(sr, id_of, ref, url):
    global _sr, _id_of, _ref, _odoo_url
    _sr, _id_of, _ref, _odoo_url = sr, id_of, ref, url


def _oid(v):
    if _id_of:
        return _id_of(v)
    return v[0] if isinstance(v, (list, tuple)) and v else v


def _label(v):
    return v[1] if isinstance(v, (list, tuple)) and len(v) > 1 else None


def _query(model, domain, fields, order=None, limit=None):
    if _sr is None:
        raise RuntimeError("devoluciones no configurado")
    return _sr(model, domain, fields, order=order, limit=limit)


def _bulk(model, ids, fields):
    """Busca por ids en lotes (evita dominios gigantes)."""
    ids = sorted({i for i in ids if i})
    out = {}
    for i in range(0, len(ids), 400):
        chunk = ids[i:i + 400]
        for r in _query(model, [("id", "in", chunk)], fields):
            out[r["id"]] = r
    return out


def _link(model, rec_id):
    return (f"{_odoo_url}/web#id={rec_id}&model={model}&view_type=form"
            if _odoo_url else "")


# --------------------------------------------------------------------------- #
# Ubicaciones y rutas comerciales
# --------------------------------------------------------------------------- #
def _ubicaciones():
    """Mapa id -> (usage, company_id) de todas las ubicaciones del ERP."""
    with _lock:
        if "locs" not in _cache:
            m = {}
            for l in _query("stock.location", [], ["id", "usage", "company_id", "complete_name"]):
                m[l["id"]] = (l.get("usage"), _oid(l.get("company_id")))
            _cache["locs"] = m
        return _cache["locs"]


def _almacenes(company_id):
    """Ids de las ubicaciones internas (de almacen) de la empresa."""
    key = ("alm", company_id)
    with _lock:
        if key not in _cache:
            locs = _ubicaciones()
            _cache[key] = {i for i, (u, c) in locs.items()
                           if u == "internal" and c == company_id}
        return _cache[key]


def _rutas_puntuales():
    """Ids de las rutas comerciales de compra puntual/local."""
    with _lock:
        if "rutas" not in _cache:
            ids = set()
            for r in _query("stock.route", [], ["id", "name"]):
                nombre = (r.get("name") or "").lower()
                if any(h in nombre for h in RUTAS_PUNTUALES):
                    ids.add(r["id"])
            _cache["rutas"] = ids
        return _cache["rutas"]


def _usage(loc_id):
    return (_ubicaciones().get(loc_id) or ("", None))[0]


# --------------------------------------------------------------------------- #
# Clasificacion de movimientos sobre las ubicaciones de la empresa
# --------------------------------------------------------------------------- #
CLASE_TRASLADO = "traslado"    # interno -> interno (no es venta)
CLASE_ENTRADA = "entrada"      # hacia el almacen de la empresa
CLASE_SALIDA = "salida"        # desde el almacen de la empresa
CLASE_EXTERNO = "externo"      # no afecta al almacen de la empresa


def _clasificar(m, alm):
    """Clase del movimiento respecto de las ubicaciones de la empresa.

    Se usa el uso de la ubicacion (customer/supplier/inventory/transit/
    internal) combinado con la pertenencia de la ubicacion a la empresa: solo
    las ubicaciones internas de FINPAC forman el almacen analizado.
    """
    src, dst = _oid(m.get("location_id")), _oid(m.get("location_dest_id"))
    src_int = src in alm
    dst_int = dst in alm
    if src_int and dst_int:
        return CLASE_TRASLADO, "traslado"
    if dst_int and not src_int:
        u = _usage(src)
        if u == "customer":
            return CLASE_ENTRADA, "devolucion"
        if u == "supplier":
            return CLASE_ENTRADA, "compra"
        if u == "inventory":
            return CLASE_ENTRADA, "ajuste_positivo"
        if u == "transit":
            return CLASE_ENTRADA, "traspaso_entrada"
        return CLASE_ENTRADA, "entrada"
    if src_int and not dst_int:
        u = _usage(dst)
        if u == "customer":
            return CLASE_SALIDA, "venta"
        if u == "inventory":
            return CLASE_SALIDA, "ajuste_negativo"
        if u == "supplier":
            return CLASE_SALIDA, "devolucion_proveedor"
        if u == "transit":
            return CLASE_SALIDA, "traspaso"
        return CLASE_SALIDA, "salida"
    return CLASE_EXTERNO, "externo"


# --------------------------------------------------------------------------- #
# Traza cronologica FIFO por capas
# --------------------------------------------------------------------------- #
MOV_CAMPOS_TRACE = ["id", "date", "quantity", "product_id", "location_id", "location_dest_id",
                    "origin_returned_move_id", "origin"]
MOV_CAMPOS_FULL = ["id", "name", "date", "quantity", "product_id", "location_id",
                   "location_dest_id", "company_id", "purchase_line_id", "sale_line_id",
                   "origin_returned_move_id", "origin"]


def _movimientos(company_id, product_ids, completo=False):
    """Movimientos done de los productos, agrupados y en orden cronologico.

    completo=True agrega los campos descriptive (name, purchase_line_id, ...);
    la traza FIFO solo necesita el juego corto de campos, que hace la lectura
    mucho mas liviana.
    """
    out = defaultdict(list)
    ids = sorted({p for p in product_ids if p})
    campos = MOV_CAMPOS_FULL if completo else MOV_CAMPOS_TRACE
    for i in range(0, len(ids), 50):
        chunk = ids[i:i + 50]
        for m in _query("stock.move",
                        [("product_id", "in", chunk), ("state", "=", "done")],
                        campos,
                        order="date asc, id asc"):
            out[_oid(m.get("product_id"))].append(m)
    return out


def _stock_vivo(company_id, product_ids, alm):
    """Stock vivo del producto en las ubicaciones internas de la empresa."""
    out = defaultdict(float)
    ids = sorted({p for p in product_ids if p})
    lista_alm = sorted(alm)
    for i in range(0, len(ids), 50):
        chunk = ids[i:i + 50]
        for q in _query("stock.quant", [("product_id", "in", chunk)],
                        ["product_id", "location_id", "quantity"]):
            if _oid(q.get("location_id")) in alm:
                out[_oid(q.get("product_id"))] += q.get("quantity") or 0.0
    return out


def _rastreo(movs, alm, inicial):
    """Traza FIFO por capas de un producto. Devuelve {move_id: pendiente}.

    inicial: stock inicial no trazado (capas mas antiguas). Las salidas
    consumen las capas mas antiguas primero; el excedente que no cubre ninguna
    capa se considera stock previo no trazado y no se imputa a la devolucion.
    """
    capas = deque()
    if inicial > EPS:
        capas.append([inicial, "", "inicial", None, None])
    pendientes = defaultdict(float)
    consumo = defaultdict(lambda: defaultdict(float))
    for m in movs:
        clase, sub = _clasificar(m, alm)
        q = abs(m.get("quantity") or 0.0)
        if q <= EPS or clase in (CLASE_TRASLADO, CLASE_EXTERNO):
            continue
        if clase == CLASE_ENTRADA:
            po = _po_codigo(m.get("origin"))
            if sub == "devolucion" and not po:
                # las unidades devueltas son las que el cliente recibio: la
                # compra de origen es la que consumio la salida original.
                po = _dominante(consumo.get(_oid(m.get("origin_returned_move_id"))))
            capas.append([q, m.get("date") or "", sub, m["id"], po])
        elif clase == CLASE_SALIDA:
            resto = q
            while resto > EPS and capas:
                c = capas[0]
                toma = c[0] if c[0] <= resto else resto
                c[0] -= toma
                resto -= toma
                if c[4]:
                    consumo[m["id"]][c[4]] += toma
                if c[0] <= EPS:
                    capas.popleft()
    for c in capas:
        if c[2] == "devolucion" and c[3]:
            pendientes[c[3]] += c[0]
    return pendientes, consumo


# --------------------------------------------------------------------------- #
# Vinculo devolucion -> compra puntual
# --------------------------------------------------------------------------- #
def _ocs_citadas(*textos):
    """Codigos de OC (P12345) presentes en los textos indicados."""
    out = []
    for t in textos:
        if not t:
            continue
        for mm in _RE_OC.finditer(str(t)):
            code = mm.group(0)
            if code not in out:
                out.append(code)
    return out


def _po_codigo(texto):
    """Primer codigo de OC (P12345) citado en un texto."""
    mm = _RE_OC.search(str(texto or ""))
    return mm.group(0) if mm else None


def _dominante(con):
    """Codigo de OC mas consumido dentro de un diccionario {oc: cantidad}."""
    if not con:
        return None
    return max(con.items(), key=lambda kv: kv[1])[0]


def _es_importada(po):
    return bool(po.get("is_import"))


def _proveedores_grupo():
    """Partner ids de las empresas del grupo (MASERVEN y FINPAC).

    Una OC cuyo proveedor es una de esas empresas NO es una compra local:
    es una compra interna/traspaso entre empresas del grupo, y este modulo es
    exclusivo de FINPAC, por lo que no se reporta.
    """
    with _lock:
        if "grupo" not in _cache:
            ids = set()
            for c in _query("res.company", [], ["id", "partner_id"]):
                pid = _oid(c.get("partner_id"))
                if pid:
                    ids.add(pid)
            _cache["grupo"] = ids
        return _cache["grupo"]


def _es_compra_interna(po):
    """True si el proveedor de la OC es una empresa del grupo (no MASERVEN)."""
    pid = _oid(po.get("partner_id"))
    return bool(pid) and pid in _proveedores_grupo()


def _po_cita_venta(po, so_id, so_nombre):
    """True si la OC esta vinculada a la orden de venta original."""
    if so_id and _oid(po.get("group_id")) == so_id:
        return True
    if so_nombre and so_nombre in (po.get("origin") or ""):
        return True
    return (po.get("sale_order_count") or 0) > 0


def _buscar_compra(product_id, sl_orden, pols, pos, so_id, so_nombre, ocs_citadas,
                    lineas_por_codigo):
    """Localiza la OC puntual que probo el producto devuelto.

    Se revisan TODAS las lineas de la orden de venta original que venden el
    mismo producto: en este ERP la compra puntual puede quedar registrada en
    otra linea de la misma orden (la que uso la ruta "Compra Puntual").

    Devuelve un dict con:
      po / pl      : la OC y su linea de compra (None si no hay)
      vinculo      : "orden_venta" | "documento" (ambos documentales)
      internas     : OC descartadas por ser compras internas entre empresas
                     del grupo (MASERVEN/FINPAC)
      importadas   : OC descartadas por ser ordenes de importacion (is_import)
    """
    rutas = _rutas_puntuales()
    ruta_puntual = any(_oid(l.get("route_id")) in rutas for l in sl_orden)
    res = {"po": None, "pl": None, "vinculo": None, "internas": [], "importadas": []}

    def _candidato(pl):
        """OC de una linea de compra, o None si esa OC no es admisible."""
        po = pos.get(_oid(pl.get("order_id")))
        if not po or _oid(po.get("company_id")) != COMPANY_FINPAC:
            return None
        if _oid(pl.get("product_id")) and _oid(pl.get("product_id")) != product_id:
            return None
        if _es_importada(po):
            res["importadas"].append(po)
            return None
        if _es_compra_interna(po):
            res["internas"].append(po)
            return None
        return po

    # via 1: vinculo documental directo de la linea de venta
    for l in sl_orden:
        for plid in (l.get("purchase_line_ids") or []):
            pl = pols.get(plid)
            if not pl:
                continue
            po = _candidato(pl)
            if po is None:
                continue
            if _po_cita_venta(po, so_id, so_nombre) or ruta_puntual:
                res.update(po=po, pl=pl, vinculo="orden_venta")
                return res
    # via 2: OC citada en el documento de la salida original / devolucion
    for code in ocs_citadas:
        for pl in lineas_por_codigo.get(code, []):
            po = _candidato(pl)
            if po is None:
                continue
            if _po_cita_venta(po, so_id, so_nombre) or ruta_puntual:
                res.update(po=po, pl=pl, vinculo="documento")
                return res
    return res


# Texto que explica como se determino el vinculo devolucion -> compra.
VINCULO_DOC_TXT = ("Vinculo documental: la orden de venta original apunta a esta OC "
                   "de compra puntual de FINPAC (linea de compra registrada en la orden).")
VINCULO_TRAZ_TXT = ("Vinculo por trazabilidad FIFO: la OC se dedujo de la capa de stock "
                    "que surtio la entrega al cliente que devuelve. No hay vinculo "
                    "documental de compra en la orden de venta; es una deduccin por "
                    "trazabilidad de capas.")



# --------------------------------------------------------------------------- #
# Reporte
# --------------------------------------------------------------------------- #
def build(company_id=COMPANY_FINPAC, desde="", hasta="", refresh=False):
    """Devoluciones pendientes de compras puntuales de FINPAC.

    company_id se fuerza a FINPAC (1): este modulo es exclusivo de esa empresa.
    desde/hasta acotan las devoluciones (fecha del movimiento de devolucion).
    refresh=True fuerza a reconsultar el ERP (no usa la cache del reporte).
    """
    company_id = COMPANY_FINPAC
    clave_cache = (desde, hasta)
    if not refresh:
        ent = _cache_reporte.get(clave_cache)
        if ent and time.time() - ent[0] < CACHE_REPORTE_SEG:
            return ent[1]
    locs = _ubicaciones()
    alm = _almacenes(company_id)

    vacio = {"ok": True, "empresa": {"id": company_id, "nombre": COMPANY_FINPAC_NOMBRE},
             "items": [], "total": 0, "candidatas": 0, "incluidas": 0,
             "excluidas": 0, "exclusiones": [], "generado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
             "mensaje": "No se encontraron devoluciones pendientes de compras puntuales."}

    # 1) devoluciones terminadas en el rango. Candidata = movimiento done que
    #    entra al almacen desde una ubicacion de cliente (todas las devoluciones,
    #    no solo las del tipo de operacion DEVCL).
    dom = [("company_id", "=", company_id), ("state", "=", "done"),
           ("location_id.usage", "=", "customer")]
    if desde:
        dom.append(("date", ">=", desde + " 00:00:00"))
    if hasta:
        dom.append(("date", "<=", hasta + " 23:59:59"))
    mv = _query("stock.move", dom,
                ["id", "name", "date", "product_id", "quantity", "picking_id",
                 "sale_line_id", "origin_returned_move_id", "origin", "location_usage",
                 "location_dest_usage", "company_id", "location_id", "location_dest_id"],
                order="date asc, id asc", limit=20000)
    # solo las que entran al almacen: una venta no es devolucion
    mv = [m for m in mv if _usage(_oid(m.get("location_dest_id"))) == "internal"]
    if not mv:
        return vacio
    candidatas_total = len(mv)

    # 2) datos relacionados
    picks = _bulk("stock.picking", [_oid(m.get("picking_id")) for m in mv],
                  ["id", "name", "origin", "partner_id", "sale_id"])
    ret_ids = [_oid(m.get("origin_returned_move_id")) for m in mv]
    ret_moves = _bulk("stock.move", ret_ids,
                      ["id", "name", "origin", "sale_line_id", "picking_id", "quantity", "date"])
    # la linea de venta original: la de la devolucion o la de la salida original
    sl_ids = []
    for m in mv:
        sl = _oid(m.get("sale_line_id"))
        if not sl:
            sl = _oid((ret_moves.get(_oid(m.get("origin_returned_move_id"))) or {}).get("sale_line_id"))
        sl_ids.append(sl)
    sols = _bulk("sale.order.line", sl_ids,
                 ["id", "order_id", "product_id", "route_id", "purchase_line_ids",
                  "product_uom_qty", "qty_returned", "name"])
    so_ids = [_oid((sols.get(i) or {}).get("order_id")) for i in sl_ids]
    sos = _bulk("sale.order", so_ids,
                ["id", "name", "partner_id", "user_id", "date_order", "state", "origin"])
    # todas las lineas de las ordenes de venta originales (el vinculo de compra
    # puntual puede estar en otra linea de la misma orden con el mismo producto)
    sl_orden = {}
    if so_ids:
        for i in range(0, len(so_ids), 300):
            for l in _query("sale.order.line", [("order_id", "in", so_ids[i:i + 300])],
                            ["id", "order_id", "product_id", "route_id", "purchase_line_ids"]):
                sl_orden[l["id"]] = l
    prod_ids = sorted({_oid(m.get("product_id")) for m in mv if m.get("product_id")})
    prods = _bulk("product.product", prod_ids,
                  ["id", "default_code", "name", "display_name", "uom_id",
                   "qty_available", "free_qty"])

    # lineas de compra vinculadas a las lineas de venta (vinculo documental)
    pl_ids = sorted({p for l in sl_orden.values()
                     for p in (l.get("purchase_line_ids") or [])})
    pols = _bulk("purchase.order.line", pl_ids,
                 ["id", "order_id", "product_id", "price_unit", "product_qty", "qty_received"])

    # 3) agrupar devoluciones por (producto, linea de venta)
    grupos = {}
    for m, slid in zip(mv, sl_ids):
        pid = _oid(m.get("product_id"))
        clave = (pid, slid)
        g = grupos.get(clave)
        if g is None:
            g = grupos[clave] = {"cantidad": 0.0, "fecha": m["date"], "moves": [],
                                 "fecha_min": m["date"], "primer_move": m}
        g["cantidad"] += m.get("quantity") or 0.0
        g["moves"].append(m["id"])
        if m["date"] < g["fecha_min"]:
            g["fecha_min"] = m["date"]
            g["primer_move"] = m

    # 4) traza cronologica FIFO por producto (una consulta por lote)
    movs_por_producto = _movimientos(company_id, prod_ids)
    stock_vivo = _stock_vivo(company_id, prod_ids, alm)
    pendientes_por_move, consumo_por_move = {}, {}
    for pid in prod_ids:
        movs = movs_por_producto.get(pid, [])
        inicial = stock_vivo.get(pid, 0.0) - _delta_movimientos(movs, alm)
        pend, cons = _rastreo(movs, alm, inicial)
        pendientes_por_move[pid] = pend
        consumo_por_move[pid] = cons

    # 5) OC que supplying las unidades (trazabilidad FIFO de la salida original)
    #    y OC citadas en los documentos de los movimientos.
    traza_oc = {}
    codigos = set()
    for m, slid in zip(mv, sl_ids):
        ret_m = ret_moves.get(_oid(m.get("origin_returned_move_id"))) or {}
        pk = picks.get(_oid(m.get("picking_id"))) or {}
        codigos.update(_ocs_citadas(m.get("origin"), ret_m.get("origin"),
                                    ret_m.get("name"), pk.get("origin")))
        pid = _oid(m.get("product_id"))
        salida = _oid(m.get("origin_returned_move_id"))
        po_fis = _dominante(consumo_por_move.get(pid, {}).get(salida))
        if po_fis:
            codigos.add(po_fis)
            traza_oc.setdefault((pid, slid), po_fis)
    campos_po = ["id", "name", "partner_id", "date_order", "date_approve", "is_import",
                 "origin", "group_id", "sale_order_count", "state", "company_id"]
    # el mismo codigo de OC existe en FINPAC y en MASERVEN: solo la de FINPAC
    pos_por_codigo = {}
    if codigos:
        nombres = set(codigos)
        nombres.update("PO" + c[1:] for c in codigos if c.startswith("P"))
        for po in _query("purchase.order",
                         [("name", "in", sorted(nombres)), ("company_id", "=", company_id)],
                         campos_po):
            pos_por_codigo[po["name"]] = po
    lineas_por_codigo = defaultdict(list)
    pos_por_id = {po["id"]: po for po in pos_por_codigo.values()}
    if pos_por_id:
        po_ids = sorted(pos_por_id)
        for i in range(0, len(po_ids), 300):
            for pl in _query("purchase.order.line", [("order_id", "in", po_ids[i:i + 300])],
                             ["id", "order_id", "product_id", "price_unit"]):
                po = pos_por_id.get(_oid(pl.get("order_id")))
                if po:
                    lineas_por_codigo[po["name"]].append(pl)

    pos = _bulk("purchase.order",
                [_oid(p.get("order_id")) for p in pols.values()] + list(pos_por_id),
                campos_po)

    items = []
    exclusiones = []
    for (pid, slid), g in grupos.items():
        prod = prods.get(pid) or {}
        sl = sols.get(slid) or {}
        so = sos.get(_oid(sl.get("order_id"))) or {}
        so_id = _oid(so.get("id"))
        so_nombre = so.get("name")
        primer = g["primer_move"]
        ret_m = ret_moves.get(_oid(primer.get("origin_returned_move_id"))) or {}
        pk = picks.get(_oid(primer.get("picking_id"))) or {}

        cand_ref = {"producto_id": pid, "codigo": prod.get("default_code") or NO_ID,
                    "orden_venta": so_nombre or NO_ID, "fecha": g["fecha_min"],
                    "cantidad": round(g["cantidad"], 4)}

        # ---- trazabilidad: unidades devueltas que siguen en el almacen ----
        pendientes = pendientes_por_move.get(pid, {})
        cant_devuelta = round(g["cantidad"], 4)
        pendiente = round(sum(pendientes.get(mid, 0.0) for mid in g["moves"]), 4)
        if pendiente <= EPS:
            exclusiones.append({**cand_ref, "motivo": "vendido_completamente",
                                "detalle": f"Las {cant_devuelta:g} u. devueltas salieron "
                                           f"por completo despues de la devolucion."})
            continue

        # ---- compra puntual vinculada ----
        ocs = _ocs_citadas(primer.get("origin"), ret_m.get("origin"),
                            ret_m.get("name"), pk.get("origin"))
        lineas_ov = [l for l in sl_orden.values()
                     if _oid(l.get("order_id")) == so_id
                     and _oid(l.get("product_id")) == pid]
        res = _buscar_compra(pid, lineas_ov, pols, pos, so_id, so_nombre, ocs,
                             lineas_por_codigo)
        po, pl, origen_vinculo = res["po"], res["pl"], res["vinculo"]
        internas, importadas = res["internas"], res["importadas"]
        if po is None:
            po_fis = traza_oc.get((pid, slid))
            po_fisico = pos_por_codigo.get(po_fis) if po_fis else None
            if po_fisico is not None and not _es_importada(po_fisico) \
                    and not _es_compra_interna(po_fisico):
                po = po_fisico
                pl = next((x for x in lineas_por_codigo.get(po_fis, [])
                           if _oid(x.get("product_id")) == pid), None)
                origen_vinculo = "trazabilidad"
            else:
                if po_fisico is not None:
                    if _es_importada(po_fisico):
                        importadas.append(po_fisico)
                    else:
                        internas.append(po_fisico)
                if importadas:
                    ocs_imp = ", ".join(sorted({p.get("name") or "" for p in importadas})[:3])
                    motivo = ("La OC (%s) es una orden de importacion (is_import): no es "
                              "una compra local, por lo que no se reporta." % ocs_imp)
                    exclusiones.append({**cand_ref, "motivo": "compra_importada",
                                        "detalle": motivo})
                elif internas:
                    ocs_int = ", ".join(sorted({p.get("name") or "" for p in internas})[:3])
                    motivo = ("La OC (%s) es una compra interna entre empresas del grupo "
                              "(MASERVEN): este modulo es solo de FINPAC." % ocs_int)
                    exclusiones.append({**cand_ref, "motivo": "compra_interna_grupo",
                                        "detalle": motivo})
                else:
                    if ocs:
                        motivo = ("La OC citada (%s) no es una compra puntual de FINPAC "
                                  "vinculada a la venta original (es importacion o no "
                                  "corresponde a esta venta)." % ", ".join(ocs[:3]))
                    else:
                        motivo = ("No se encontro una orden de compra puntual vinculada a la "
                                  "venta original (ni vinculo de compra en las lineas de la "
                                  "orden de venta ni OC citada en los movimientos).")
                    exclusiones.append({**cand_ref, "motivo": "sin_compra_puntual",
                                        "detalle": motivo})
                continue

        cliente = _label(so.get("partner_id")) or NO_ID
        vendedor = _label(so.get("user_id")) or NO_ID
        proveedor = _label(po.get("partner_id")) or NO_ID
        fecha_oc = (po.get("date_approve") or po.get("date_order") or "")[:10] or NO_ID
        costo = pl.get("price_unit") if pl else None
        if costo is None:
            costo = 0.0
        stock_actual = round(prod.get("qty_available") or 0.0, 4)
        aviso = ""
        if pendiente > stock_actual + 0.01:
            aviso = (f"La cantidad pendiente ({pendiente:g} u.) es mayor que el stock "
                     f"actual del producto ({stock_actual:g} u.): el inventario y los "
                     f"movimientos de devolucion no cuadran; revise el caso.")
        # tipo de vinculo: DOC = documental confirmado, TRAZ = deducido por FIFO
        vinculo_tipo = "TRAZ" if origen_vinculo == "trazabilidad" else "DOC"
        vinculo_detalle = (VINCULO_TRAZ_TXT if vinculo_tipo == "TRAZ"
                           else VINCULO_DOC_TXT)

        items.append({
            "cliente": cliente,
            "orden_venta": so_nombre or NO_ID,
            "codigo": prod.get("default_code") or NO_ID,
            "descripcion": prod.get("name") or NO_ID,
            "cantidad_devuelta": cant_devuelta,
            "cantidad_pendiente": pendiente,
            "stock_actual": stock_actual,
            "orden_compra": po.get("name") or NO_ID,
            "proveedor": proveedor,
            "fecha_oc": fecha_oc,
            "costo_unitario": round(costo or 0.0, 6),
            "vendedor": vendedor,
            "producto_id": pid,
            "company_id": company_id,
            "fecha_devolucion": g["fecha_min"][:10],
            "moves": g["moves"],
            "move_id": primer["id"],
            "vinculo": origen_vinculo,
            "vinculo_tipo": vinculo_tipo,
            "vinculo_detalle": vinculo_detalle,
            "aviso": aviso,
            "devolucion_odoo": _link("stock.move", primer["id"]),
            "oc_odoo": _link("purchase.order", _oid(po.get("id"))),
        })

    items.sort(key=lambda i: (i["fecha_oc"], i["codigo"], i["orden_venta"]))
    data = {"ok": True, "empresa": {"id": company_id, "nombre": COMPANY_FINPAC_NOMBRE},
            "items": items, "total": len(items),
            "candidatas": candidatas_total,
            "incluidas": len(items),
            "excluidas": len(exclusiones),
            "exclusiones": exclusiones,
            "generado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "mensaje": "" if items else "No se encontraron devoluciones pendientes "
                                         "con los filtros aplicados."}
    with _lock:
        _cache_reporte[clave_cache] = (time.time(), data)
    return data


def _delta_movimientos(movs, alm):
    """Delta de inventario de los movimientos sobre el almacen de la empresa."""
    d = 0.0
    for m in movs:
        clase, _sub = _clasificar(m, alm)
        q = abs(m.get("quantity") or 0.0)
        if clase == CLASE_ENTRADA:
            d += q
        elif clase == CLASE_SALIDA:
            d -= q
    return d


# --------------------------------------------------------------------------- #
# Traza de movimientos de un caso (para revision)
# --------------------------------------------------------------------------- #
def detalle(company_id, move_id):
    """Movimientos que sustentan un caso de devolucion."""
    company_id = COMPANY_FINPAC
    alm = _almacenes(company_id)
    mv = _query("stock.move", [("id", "=", move_id)], ["id", "name", "date", "product_id",
                 "quantity", "picking_id", "sale_line_id"])
    if not mv:
        return {"ok": False, "mensaje": "Caso no encontrado."}
    m = mv[0]
    pid = _oid(m.get("product_id"))
    movs = _movimientos(company_id, [pid], completo=True).get(pid, [])
    stock = _stock_vivo(company_id, [pid], alm).get(pid, 0.0)
    inicial = stock - _delta_movimientos(movs, alm)
    pendientes, _consumo = _rastreo(movs, alm, inicial)
    etiquetas = {"traslado": "Traslado interno (no es venta)",
                 "venta": "Venta / entrega a cliente",
                 "ajuste_negativo": "Ajuste de inventario (-)",
                 "ajuste_positivo": "Ajuste de inventario (+)",
                 "devolucion_proveedor": "Devolucion a proveedor",
                 "traspaso": "Traspaso a otra empresa",
                 "traspaso_entrada": "Ingreso por traspaso",
                 "compra": "Recepcion de compra",
                 "devolucion": "Devolucion de cliente",
                 "entrada": "Entrada", "salida": "Salida", "externo": "Fuera de la empresa"}
    filas = []
    saldo = inicial
    for x in movs:
        clase, sub = _clasificar(x, alm)
        q = abs(x.get("quantity") or 0.0)
        if clase == CLASE_ENTRADA:
            saldo += q
        elif clase == CLASE_SALIDA:
            saldo -= q
        filas.append({
            "fecha": x["date"], "nombre": x["name"], "cantidad": round(x.get("quantity") or 0.0, 4),
            "tipo": etiquetas.get(sub, sub),
            "origen": _label(x.get("origin")) or (x.get("origin") or ""),
            "pendiente": round(pendientes.get(x["id"], 0.0), 4) if sub == "devolucion" else None,
            "saldo": round(saldo, 4),
        })
    return {"ok": True, "producto_id": pid, "stock_actual": round(stock, 4),
            "stock_inicial_no_trazado": round(inicial, 4),
            "devolucion": {"id": m["id"], "nombre": m["name"], "fecha": m["date"][:10],
                           "cantidad": round(m.get("quantity") or 0.0, 4)},
            "movimientos": filas}


# --------------------------------------------------------------------------- #
# Exportacion a Excel
# --------------------------------------------------------------------------- #
def build_xlsx(data):
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Devoluciones pendientes"

    head_fill = PatternFill("solid", fgColor="0F172A")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = ("DEVOLUCIONES PENDIENTES DE COMPRAS PUNTUALES \u2014 "
                f"{data['empresa']['nombre']}")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws["A2"] = f"Generado: {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(size=10, color="64748B")

    headers = ["CLIENTE", "ORDEN DE VENTA ORIGINAL", "C\u00d3DIGO", "DESCRIPCI\u00d3N",
               "CANTIDAD DEVUELTA", "CANTIDAD PENDIENTE", "STOCK ACTUAL", "ORDEN DE COMPRA",
               "PROVEEDOR", "FECHA DE O/C", "COSTO UNITARIO", "VENDEDOR"]
    head_row = 4
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=c, value=h)
        cell.fill = head_fill
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    r = head_row
    for it in data["items"]:
        r += 1
        vals = [it["cliente"], it["orden_venta"], it["codigo"], it["descripcion"],
                it["cantidad_devuelta"], it["cantidad_pendiente"], it["stock_actual"],
                it["orden_compra"], it["proveedor"], it["fecha_oc"],
                it["costo_unitario"], it["vendedor"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (1, 4, 9)))
            if c in (5, 6, 7):
                cell.number_format = "#,##0.00"
            if c == 11:
                cell.number_format = '"$"#,##0.000000'
            if r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
        # vinculo de la OC y avisos (comentarios: el cuadro mantiene las 12 columnas)
        es_traz = it.get("vinculo_tipo") == "TRAZ" or it.get("vinculo") == "trazabilidad"
        vinculo = (it.get("vinculo_detalle") or
                   ("Trazabilidad FIFO: la OC se dedujo de la capa de stock que surtio la "
                    "entrega al cliente que devuelve."
                    if es_traz else
                    "Vinculo documental: la orden de venta original apunta a esta OC "
                    "(compra puntual de FINPAC)."))
        ws.cell(row=r, column=8).comment = Comment(vinculo, "Devoluciones pendientes")
        if es_traz:
            # se distingue a simple vista el vinculo deducido del documental
            celda = ws.cell(row=r, column=8)
            celda.fill = PatternFill("solid", fgColor="FEF3C7")
            celda.font = Font(bold=True, color="92400E", size=11)
        if it.get("aviso"):
            ws.cell(row=r, column=6).comment = Comment(it["aviso"], "Devoluciones pendientes")
            for c in range(1, 13):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFF7ED")

    if not data["items"]:
        ws.cell(row=head_row + 1, column=1,
                value="Sin devoluciones pendientes de compras puntuales en el rango indicado.")

    widths = [30, 20, 12, 40, 16, 16, 13, 16, 32, 12, 14, 28]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    if data["items"]:
        ws.auto_filter.ref = f"A{head_row}:L{head_row + len(data['items'])}"
    ws.row_dimensions[head_row].height = 30

    nr = head_row + len(data["items"]) + 3
    ws.cell(row=nr, column=1, value="NOTAS").font = Font(bold=True, size=11)
    notas = [
        f"Modulo exclusivo de {data['empresa']['nombre']}: solo compras puntuales/locales "
        "vinculadas a la orden de venta original de la devolucion.",
        "No se incluyen las compras internas con MASERVEN (traspasos entre empresas del "
        "grupo): este reporte es solo de FINPAC.",
        "CLIENTE: de la orden de venta original. VENDEDOR: de la orden de venta original.",
        "CANTIDAD DEVUELTA: unidades devueltas por el cliente en el rango indicado.",
        "CANTIDAD PENDIENTE: unidades devueltas que siguen en inventario. Se calcula con "
        "trazabilidad cronologica FIFO por capas sobre los movimientos reales: las salidas "
        "(ventas, ajustes, devoluciones a proveedor) consumen primero las unidades mas "
        "antiguas, de modo que las unidades devueltas se descuentan cuando otro vendedor "
        "las vende. Los traslados internos no son ventas.",
        "STOCK ACTUAL: stock total del producto en el ERP (no es el stock de la devolucion).",
        "ORDEN DE COMPRA: la OC de FINPAC que surtio el producto. En el comentario de la "
        "celda se indica el vinculo: documental (la orden de venta apunta a la OC) o por "
        "trazabilidad FIFO (la OC se dedujo de la capa que entrego las unidades). Las OCs "
        "con vinculo por trazabilidad FIFO van resaltadas en ambar.",
        "Las ordenes de compra marcadas como importacion (is_import) no se reportan: este "
        "modulo es de compras puntuales/locales.",
        "Las compras posteriores del mismo codigo abren capas nuevas: ese stock no se "
        "atribuye a la devolucion original.",
        "Si Odoo no permite identificar fisicamente las mismas unidades, la cantidad "
        "pendiente es el resultado de la trazabilidad cronologica descrita, no una "
        "certificacion fisica de unidades.",
        "Este modulo no genera cobros ni modifica el ERP.",
    ]
    for i, n in enumerate(notas, start=1):
        ws.cell(row=nr + i, column=1, value=f"{i}. {n}")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
