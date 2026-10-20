"""Devoluciones pendientes de compras locales - FINPAC.

Identifica UNICAMENTE los productos de compras locales/puntuales que un cliente
devolvio (movimientos DEVCL terminados) y que todavia tienen unidades pendientes
de salida, es decir, que no se vendieron por completo despues de la devolucion.

Modulo de solo lectura sobre Odoo: NO escribe en el ERP, NO genera cobros ni
descuentos, y NO modifica el Kardex existente. Reutiliza la conexion XML-RPC y
la logica de inventario del servidor por inyeccion (evita import circular).

El vendedor se obtiene SIEMPRE de la orden de venta original de la devolucion
(sale.order.user_id), nunca de la compra.
"""
import io
import re
import threading
from collections import defaultdict
from datetime import datetime

# ---- Configuracion / acceso a Odoo inyectados por kardex_server ----
_sr = None
_id_of = None
_ref = None
_odoo_url = None
_lock = threading.RLock()
_cache = {}

COMPANY_FINPAC = 1
COMPANY_FINPAC_NOMBRE = "FINPAC"
NO_ID = "No identificado"
REVISAR = "Revisar"

# Route comercial que marca una compra puntual / en oficina (verificado en Odoo: id 17).
PUNTUAL_HINTS = ("puntual", "en oficina")
# Documento origen con el que una OC cita un pedido de venta (Sxxxxx).
_RE_SALE = re.compile(r"S\d{4,6}")


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
    ids = sorted({i for i in ids if i})
    if not ids:
        return {}
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
# Descubrimiento de tipos de operacion y rutas (cacheado en memoria)
# --------------------------------------------------------------------------- #
def _devcl_types(company_id):
    key = ("devcl", company_id)
    with _lock:
        if key not in _cache:
            _cache[key] = [t["id"] for t in _query(
                "stock.picking.type",
                [("sequence_code", "=", "DEVCL"), ("code", "=", "incoming"),
                 ("company_id", "=", company_id)], ["id"])]
        return _cache[key]


def _puntual_route_ids():
    key = "routes_puntual"
    with _lock:
        if key not in _cache:
            _cache[key] = [r["id"] for r in _query("stock.route", [], ["id", "name"])
                           if any(h in (r.get("name") or "").lower() for h in PUNTUAL_HINTS)]
        return _cache[key]


def _outgoing_type_ids(company_id):
    key = ("out", company_id)
    with _lock:
        if key not in _cache:
            _cache[key] = [t["id"] for t in _query(
                "stock.picking.type",
                [("code", "=", "outgoing"), ("company_id", "=", company_id)], ["id"])]
        return _cache[key]


# --------------------------------------------------------------------------- #
# Vinculacion devolucion -> compra local
# --------------------------------------------------------------------------- #
def _es_importada(po):
    """Criterio verificable de compra importada (excluida del reporte)."""
    return bool(po.get("is_import"))


# Rutas comerciales que acreditan una compra puntual/local en FINPAC.
PUNTUAL_ROUTE_IDS = (17, 4)  # "Compra Puntual (En Oficina)", "Finpac: Cross-Dock"

# Codigo de OC local de FINPAC citado en textos de documento (p.ej. "P14563",
# "S25196 - 260015499"). Se extrae para vincular la devolucion con su compra.
_RE_PO = re.compile(r"\bP\d{4,6}\b")


def _po_names_en_texto(*textos):
    """Codigos de OC (formato P####) presentes en uno o varios textos."""
    out = []
    for t in textos:
        if not t:
            continue
        for mm in _RE_PO.finditer(str(t)):
            code = mm.group(0)
            if code not in out:
                out.append(code)
    return out


def _compra_local_de_linea(sl, pols, pos, origin_po_names=None, pols_por_nombre=None):
    """Devuelve (po, pl) de la compra local que origino la devolucion, o (None,_).

    Vias verificables, en orden de fiabilidad:
      1. purchase_line_ids documental de la linea de venta.
      2. Codigo(s) de OC citados en el documento (devolucion y/o salida original),
         p.ej. origin = "P14563" o "S25196 - 260015499".
    Se descartan las POs marcadas como importacion. Si la linea de venta tiene
    producto, la OC candidata debe contener ese mismo producto (evita cruzar
    compras de otro item citadas en el mismo documento). No fusiona varias:
    devuelve la primera compra local comprobable.
    """
    # via 1: vinculo documental directo
    for plid in (sl.get("purchase_line_ids") or []):
        pl = pols.get(plid)
        if not pl:
            continue
        po = pos.get(_oid(pl.get("order_id")))
        if not po or _es_importada(po):
            continue
        return po, pl
    # via 2: OC citada como origen (de la devolucion o de la salida original)
    if origin_po_names and pols_por_nombre:
        pid_sl = _oid(sl.get("product_id"))
        for name in origin_po_names:
            for pl in pols_por_nombre.get((name or "").strip(), []):
                po = pos.get(_oid(pl.get("order_id")))
                if not po or _es_importada(po):
                    continue
                if pid_sl and _oid(pl.get("product_id")) and _oid(pl.get("product_id")) != pid_sl:
                    continue
                return po, pl
    return None, None


def _compras_locales_por_producto(product_id, pol_por_producto, pos):
    """Compras locales conocidas del producto (para detectar ambiguedad/adicionales)."""
    out = []
    for pl in pol_por_producto.get(product_id, []):
        po = pos.get(_oid(pl.get("order_id")))
        if po and not _es_importada(po):
            out.append((po, pl))
    return out


# --------------------------------------------------------------------------- #
# Analisis posterior: rastreo cronologico de la devolucion en el inventario
# --------------------------------------------------------------------------- #
# La clasificacion de cada movimiento se hace por la NATURALEZA de sus
# ubicaciones (usage de stock.location), no solo por el tipo de operacion:
#   - sale a      -> usage "customer"  (venta / entrega real)
#   - entra desde -> usage "supplier"  (recepcion real de compra)
#   - ajuste      -> usage "inventory" (ajuste de inventario / desecho)
#   - interno     -> ambos extremos internos (no altera la disponibilidad neta)
# Esto es verificable en Odoo y estable ante cambios de flujos/almacenes.
_VIRTUAL_USAGE_CUSTOMER = "customer"
_VIRTUAL_USAGE_SUPPLIER = "supplier"
_VIRTUAL_USAGE_INVENTORY = "inventory"


def _ubicaciones_virtuales(company_id):
    """Mapa id -> usage para las ubicaciones que NO son internas (cacheado)."""
    key = ("locs_usage", company_id)
    with _lock:
        if key not in _cache:
            m = {}
            for l in _query("stock.location", [("company_id", "=", company_id)],
                            ["id", "usage"]):
                m[l["id"]] = l.get("usage")
            # las ubicaciones compartidas (Socios/Clientes, Proveedores, virtuales)
            # pueden tener company_id vacio; se traen aparte.
            for l in _query("stock.location", [("company_id", "=", False)],
                            ["id", "usage"]):
                m.setdefault(l["id"], l.get("usage"))
            _cache[key] = m
        return _cache[key]


def _movimientos_productos(company_id, product_ids):
    """Todos los movimientos done de los productos, agrupados por producto.

    Una sola consulta por lote en lugar de una por caso: el historial completo de
    cada producto se reutiliza para rastrear el destino de las unidades devueltas
    y detectar stock nuevo (compras/ajustes posteriores) del mismo codigo.
    """
    out = defaultdict(list)
    ids = sorted({p for p in product_ids if p})
    for i in range(0, len(ids), 200):
        chunk = ids[i:i + 200]
        for m in _query("stock.move",
                        [("company_id", "=", company_id), ("product_id", "in", chunk),
                         ("state", "=", "done")],
                        ["id", "date", "quantity", "picking_type_id", "product_id",
                         "location_id", "location_dest_id", "is_inventory", "reference"],
                        order="date asc, id asc"):
            out[_oid(m.get("product_id"))].append(m)
    return out


# Usages que representan salida definitiva del stock disponible del almacen.
_USAGE_SALIDA = {"customer", "inventory", "production"}
_USAGE_INTERNO = {"internal", "transit"}


def _clasificar_movimiento(m, locs_usage):
    """Clase de un movimiento segun la NATURALEZA de sus ubicaciones.

    Devuelve uno de:
      "venta"              -> salida a cliente (uso customer).
      "salida_ajuste"      -> egreso por ajuste/desecho/produccion (uso inventory/
                              production como destino, o sea el almacen pierde stock).
      "entrada"            -> ingreso real de proveedor o ajuste positivo/producto.
      "devolucion_cliente" -> entra desde customer a una ubicacion interna.
      "interno"            -> reubicacion entre ubicaciones internas/transito
                              (NO altera la disponibilidad neta; se ignora).
      "otro"               -> patron no reconocido (se ignora para el saldo).

    Regla general por usages, no por tipo de operacion ni por is_inventory:
    un traslado interno registrado como ajuste (interno -> interno, is_inventory)
    es una reubicacion, NO una salida.
    """
    src = _oid(m.get("location_id"))
    dst = _oid(m.get("location_dest_id"))
    u_src = locs_usage.get(src)
    u_dst = locs_usage.get(dst)

    src_int = u_src in _USAGE_INTERNO
    dst_int = u_dst in _USAGE_INTERNO

    # 1) devolucion de cliente: sale de Clientes y entra al almacen
    if u_src == _VIRTUAL_USAGE_CUSTOMER and dst_int:
        return "devolucion_cliente"
    # 2) venta / entrega a cliente
    if u_dst == _VIRTUAL_USAGE_CUSTOMER and src_int:
        return "venta"
    # 3) ingreso real de proveedor (o desde inventario virtual hacia el almacen)
    if u_src == _VIRTUAL_USAGE_SUPPLIER and dst_int:
        return "entrada"
    # 4) ingreso por ajuste positivo: de inventario virtual/produccion hacia el almacen
    if u_src in _USAGE_SALIDA and dst_int and u_src != _VIRTUAL_USAGE_CUSTOMER:
        return "entrada"
    # 5) egreso por ajuste/desecho/produccion: del almacen hacia inventario virtual
    if src_int and u_dst in _USAGE_SALIDA and u_dst != _VIRTUAL_USAGE_CUSTOMER:
        return "salida_ajuste"
    # 6) reubicacion interna (incluye traslados registrados como ajuste)
    if src_int and dst_int:
        return "interno"
    return "otro"


def _rastrear(movs, fecha_dev, move_ids, cant_devuelta, locs_usage):
    """Rastrea cronologicamente el destino de las unidades devueltas.

    Criterio consistente (no certeza de unidades fisicas): se mantiene un saldo
    FIFO de la devolucion. Las salidas (ventas/ajustes negativos) consumen primero
    las unidades devueltas y, si ya no quedan, el stock nuevo. Las entradas
    (compras/ajustes positivos) NO se atribuyen a la devolucion: son stock nuevo.

    Devuelve dict con:
      - vendida: unidades de la devolucion que salieron (imputadas FIFO).
      - pendiente: unidades de la devolucion que siguen en inventario.
      - stock_nuevo_posterior: unidades ingresadas despues (compras/ajustes +),
        que hacen que el stock actual NO equivalga a la devolucion.
      - compras_posteriores/ajustes_posteriores/ventas_posteriores: desglose.
      - primer_evento, ultimo_evento: fechas extremas consideradas.
    """
    move_ids = set(move_ids)
    dev_restante = float(cant_devuelta)
    vendida = 0.0
    entradas = 0.0
    ajustes_in = 0.0
    ajustes_out = 0.0
    ventas = 0.0
    primer = None
    ultimo = None

    for m in movs:
        if m["id"] in move_ids:
            continue
        if m["date"] < fecha_dev:
            continue
        q = abs(m.get("quantity") or 0.0)
        if q == 0:
            continue
        clase = _clasificar_movimiento(m, locs_usage)
        if primer is None:
            primer = m["date"]
        ultimo = m["date"]

        if clase in ("interno", "otro", "devolucion_cliente"):
            # reubicacion interna u otro patron: no altera la disponibilidad
            continue
        if clase == "entrada":
            entradas += q
            continue
        # clases que sacan unidades del almacen disponible (venta / ajuste de egreso)
        if clase == "venta":
            ventas += q
        elif clase == "salida_ajuste":
            ajustes_out += q
        # imputacion FIFO: primero consume lo devuelto
        consume = min(dev_restante, q)
        if consume > 0:
            dev_restante -= consume
            vendida += consume

    return {
        "vendida": round(vendida, 4),
        "pendiente": round(max(0.0, dev_restante), 4),
        "stock_nuevo": round(entradas + ajustes_in, 4),
        "entradas": round(entradas, 4),
        "ajustes_in": round(ajustes_in, 4),
        "ajustes_out": round(ajustes_out, 4),
        "ventas": round(ventas, 4),
        "primer_evento": primer,
        "ultimo_evento": ultimo,
    }


# --------------------------------------------------------------------------- #
# Reporte
# --------------------------------------------------------------------------- #
def build(company_id=COMPANY_FINPAC, desde="", hasta="", vendedor=""):
    """Construye el reporte de devoluciones pendientes de compras locales.

    company_id se fuerza a FINPAC (1): este modulo es exclusivo de esa empresa.
    """
    company_id = COMPANY_FINPAC
    devcl = _devcl_types(company_id)
    out_types = _outgoing_type_ids(company_id)
    puntual = _puntual_route_ids()
    locs_usage = _ubicaciones_virtuales(company_id)
    customer_ids = [i for i, u in locs_usage.items() if u == _VIRTUAL_USAGE_CUSTOMER]

    vacio = {"ok": True, "empresa": {"id": company_id, "nombre": COMPANY_FINPAC_NOMBRE},
             "items": [], "total": 0, "vendedores": [], "candidatas": 0,
             "incluidas": 0, "excluidas": 0, "exclusiones": [],
             "mensaje": "No se encontraron devoluciones pendientes de compras locales."}
    if not devcl:
        vacio["mensaje"] = "No existe un tipo de operacion DEVCL para FINPAC."
        return vacio

    # 1) devoluciones terminadas en el rango. Candidata = movimiento done cuyo
    # ORIGEN es una ubicacion de cliente (usage customer) y DESTINO una ubicacion
    # interna. Esto cubre TODAS las devoluciones, no solo el picking_type DEVCL.
    dom = [("company_id", "=", company_id), ("state", "=", "done"),
           ("location_id", "in", customer_ids)]
    if desde:
        dom.append(("date", ">=", desde + " 00:00:00"))
    if hasta:
        dom.append(("date", "<=", hasta + " 23:59:59"))
    mv_raw = _query("stock.move", dom,
                    ["id", "name", "date", "product_id", "quantity", "picking_id",
                     "picking_type_id", "location_id", "location_dest_id",
                     "sale_line_id", "origin_returned_move_id"],
                    order="date asc, id asc", limit=20000)
    # solo las que entran al almacen (destino interno): una venta no es devolucion
    mv = [m for m in mv_raw if locs_usage.get(_oid(m.get("location_dest_id"))) in _USAGE_INTERNO]
    candidatas_total = len(mv)
    if not mv:
        return vacio

    # 2) datos relacionados
    picks = _bulk("stock.picking", [_oid(m.get("picking_id")) for m in mv],
                  ["id", "name", "partner_id", "origin", "sale_id", "purchase_id"])
    # salida original (para leer la OC citada en su documento origen)
    ret_move_ids = sorted({_oid(m.get("origin_returned_move_id"))
                           for m in mv if m.get("origin_returned_move_id")})
    ret_moves = _bulk("stock.move", ret_move_ids,
                      ["id", "picking_id", "origin", "reference"])
    ret_picks = _bulk("stock.picking",
                      [(_oid(r.get("picking_id")) if r.get("picking_id") else None)
                       for r in ret_moves.values()],
                      ["id", "name", "origin"])
    sl_ids = [_oid(m.get("sale_line_id")) for m in mv]
    sols = _bulk("sale.order.line", sl_ids,
                 ["id", "order_id", "route_id", "purchase_line_ids", "product_id"])
    sos = _bulk("sale.order", [_oid((sols.get(i) or {}).get("order_id")) for i in sl_ids],
                ["id", "name", "partner_id", "user_id"])
    prods = _bulk("product.product", [_oid(m.get("product_id")) for m in mv],
                  ["id", "default_code", "name", "display_name", "uom_id"])
    partners = _bulk("res.partner",
                     [_oid((picks.get(_oid(m.get("picking_id"))) or {}).get("partner_id"))
                      for m in mv] +
                     [_oid((sos.get(_oid((sols.get(i) or {}).get("order_id"))) or {}).get("user_id"))
                      for i in sl_ids],
                     ["id", "name"])

    # POs de las lineas devueltas (documental)
    pl_ids = sorted({pid for i in sl_ids
                     for pid in ((sols.get(i) or {}).get("purchase_line_ids") or [])})
    pols = _bulk("purchase.order.line", pl_ids,
                 ["id", "order_id", "product_id", "price_unit", "product_qty", "qty_received"])
    pos = _bulk("purchase.order", [_oid(p.get("order_id")) for p in pols.values()],
                ["id", "name", "partner_id", "date_order", "date_approve", "is_import", "origin"])

    # compras locales conocidas por producto (deteccion de adicionales / ambiguedad)
    prod_ids = sorted({_oid(m.get("product_id")) for m in mv if m.get("product_id")})
    pol_por_producto = defaultdict(list)
    for i in range(0, len(prod_ids), 300):
        chunk = prod_ids[i:i + 300]
        for pl in _query("purchase.order.line", [("product_id", "in", chunk)],
                         ["id", "order_id", "product_id", "price_unit"]):
            pol_por_producto[_oid(pl.get("product_id"))].append(pl)
    pos_extra = _bulk("purchase.order",
                      [_oid(pl.get("order_id")) for lst in pol_por_producto.values() for pl in lst],
                      ["id", "name", "partner_id", "date_order", "date_approve", "is_import", "origin"])
    pos_all = {**pos, **pos_extra}

    # lineas de compra indexadas por codigo de OC (para la via origin -> OC citada)
    pols_por_nombre = defaultdict(list)
    for lst in pol_por_producto.values():
        for pl in lst:
            po = pos_all.get(_oid(pl.get("order_id")))
            if po and po.get("name"):
                pl = dict(pl)
                pl["order_id"] = po["id"]
                pols_por_nombre[po["name"]].append(pl)

    # costo unitario de la devolucion (capa de valorizacion) como respaldo
    costo_dev = {}
    for lay in _query("stock.valuation.layer",
                      [("stock_move_id", "in", [m["id"] for m in mv])],
                      ["stock_move_id", "quantity", "unit_cost"]):
        mid = _oid(lay.get("stock_move_id"))
        q, uc = lay.get("quantity") or 0.0, lay.get("unit_cost") or 0.0
        acc = costo_dev.setdefault(mid, [0.0, 0.0])
        acc[0] += q
        acc[1] += q * uc
    costo_dev = {k: (v[1] / v[0] if v[0] else None) for k, v in costo_dev.items()}

    # historial completo de movimientos de los productos (una consulta por lote)
    movs_por_producto = _movimientos_productos(company_id, prod_ids)

    # 3) agrupar devoluciones por (producto, pedido de venta) para no duplicar
    grupos = defaultdict(lambda: {"cantidad": 0.0, "fecha_min": None, "move_ids": [],
                                  "primera_move": None})
    for m in mv:
        sl = sols.get(_oid(m.get("sale_line_id"))) or {}
        so = sos.get(_oid(sl.get("order_id"))) or {}
        clave = (_oid(m.get("product_id")), _oid(so.get("order_id")) or _oid(m.get("sale_line_id")),
                 _oid(m.get("sale_line_id")))
        g = grupos[clave]
        g["cantidad"] += m.get("quantity") or 0.0
        g["move_ids"].append(m["id"])
        if g["fecha_min"] is None or m["date"] < g["fecha_min"]:
            g["fecha_min"] = m["date"]
            g["primera_move"] = m
        g["_m"] = m
        g["_sl"] = sl
        g["_so"] = so

    items = []
    vendedores = {}
    exclusiones = []  # auditoria de cobertura
    for (pid, _soid, slid), g in grupos.items():
        prod = prods.get(pid) or {}
        m = g["_m"]
        sl = g["_sl"]
        so = g["_so"]
        pk = picks.get(_oid(m.get("picking_id"))) or {}

        cand_ref = {"move_id": m["id"], "fecha": g["fecha_min"],

                    "codigo": prod.get("default_code") or NO_ID,
                    "orden_venta": so.get("name") or NO_ID,
                    "cantidad": round(g["cantidad"], 4)}

        # vendedor desde la orden de venta original (nunca de la compra)
        vend_nombre = _label(so.get("user_id")) or ""
        if not vend_nombre:
            vend_nombre = (partners.get(_oid(so.get("user_id"))) or {}).get("name") or ""
        if not vend_nombre:
            vend_nombre = NO_ID

        # compra local de origen: via documental (purchase_line_ids) o via OC citada
        # en el documento de devolucion y/o en la salida original.
        ret_m = ret_moves.get(_oid(m.get("origin_returned_move_id"))) or {}
        ret_pk = ret_picks.get(_oid(ret_m.get("picking_id"))) or {}
        origin_po_names = _po_names_en_texto(
            pk.get("origin"), ret_m.get("origin"), ret_m.get("reference"),
            ret_pk.get("origin"))
        po, pl = _compra_local_de_linea(sl, pols, pos_all,
                                        origin_po_names=origin_po_names,
                                        pols_por_nombre=pols_por_nombre)
        if po is None:
            exclusiones.append({**cand_ref, "motivo": "sin_compra_local",
                                "detalle": "No se pudo vincular una compra local/puntual "
                                           "(sin purchase_line_ids ni OC P#### citada en "
                                           "la devolucion ni en la salida original)."})
            continue

        vendedores[vend_nombre] = True

        # cantidad devuelta acumulada del producto+pedido
        cant_dev = round(g["cantidad"], 4)

        # rastreo cronologico del destino de las unidades devueltas
        tr = _rastrear(movs_por_producto.get(pid, []), g["fecha_min"], g["move_ids"],
                       cant_dev, locs_usage)
        vendida = tr["vendida"]
        pendiente = tr["pendiente"]

        # compras locales ADICIONALES del mismo codigo posteriores a la devolucion:
        # refuerzan que el stock actual no es necesariamente el devuelto.
        fecha_dev_dia = g["fecha_min"][:10]
        compras_extra = []
        for cpo, cpl in _compras_locales_por_producto(pid, pol_por_producto, pos_all):
            if po and _oid(cpo["id"]) == _oid(po["id"]):
                continue
            f = (cpo.get("date_approve") or cpo.get("date_order") or "")[:10]
            if f and f >= fecha_dev_dia:
                compras_extra.append((cpo, cpl))

        # ---- estado del caso ----
        if pendiente <= 0:
            estado = "VENDIDO COMPLETAMENTE"
        elif vendida > 0:
            estado = "PARCIALMENTE VENDIDO"
        else:
            estado = "PENDIENTE"

        # exclusion: si se vendio por completo despues de la devolucion, no es un
        # caso pendiente y no se lista (pero queda registrado en la auditoria).
        if estado == "VENDIDO COMPLETAMENTE":
            exclusiones.append({**cand_ref, "motivo": "vendido_completamente",
                                "detalle": f"Las {cant_dev:g} u. devueltas salieron por completo "
                                           f"despu\u00e9s de la devoluci\u00f3n."})
            continue

        # avisos: stock nuevo posterior o compras adicionales => la cantidad
        # pendiente es una estimacion por asignacion FIFO, no certeza fisica.
        ambiguo = bool(tr["stock_nuevo"] > 0 or compras_extra)
        aviso = ""
        if ambiguo:
            partes = []
            if tr["entradas"] > 0:
                partes.append(f"{tr['entradas']:g} u. recibidas de proveedor")
            if compras_extra:
                partes.append(f"{len(compras_extra)} compra(s) posterior(es) del mismo c\u00f3digo")
            aviso = ("Hubo " + " y ".join(partes) +
                     " despu\u00e9s de la devoluci\u00f3n: la cantidad pendiente se estima por "
                     "asignaci\u00f3n FIFO, no es certeza f\u00edsica de las mismas unidades.")

        costo_unit = (pl.get("price_unit") if pl and pl.get("price_unit") else None) \
            or costo_dev.get(m["id"]) or 0.0
        contacto = (partners.get(_oid(pk.get("partner_id"))) or {}).get("name") \
            or (partners.get(_oid(so.get("partner_id"))) or {}).get("name") or NO_ID
        proveedor = (partners.get(_oid(po.get("partner_id"))) or {}).get("name") \
            or _label(po.get("partner_id")) or NO_ID

        items.append({
            "move_id": m["id"],
            "move_ids": g["move_ids"],
            "contacto": contacto,
            "orden_venta": so.get("name") or NO_ID,
            "codigo": prod.get("default_code") or NO_ID,
            "descripcion": prod.get("name") or NO_ID,
            "cantidad_devuelta": cant_dev,
            "cantidad_vendida_despues": vendida,
            "cantidad_pendiente": pendiente,
            "estado": estado,
            "stock_nuevo_posterior": tr["stock_nuevo"],
            "entradas_posteriores": tr["entradas"],
            "ajustes_entrada_posteriores": tr["ajustes_in"],
            "ajustes_salida_posteriores": tr["ajustes_out"],
            "ventas_posteriores": tr["ventas"],
            "cantidad_estimada": ambiguo,
            "aviso": aviso,
            "orden_compra": po.get("name") or NO_ID,
            "proveedor": proveedor,
            "fecha_oc": (po.get("date_approve") or po.get("date_order") or "")[:10] or NO_ID,
            "costo_unitario": round(costo_unit or 0.0, 6),
            "vendedor": vend_nombre or NO_ID,
            "producto_id": pid,
            "company_id": company_id,
            "devolucion_odoo": _link("stock.move", m["id"]),
            "oc_odoo": _link("purchase.order", _oid(po.get("id"))),
        })

    # 4) filtro por vendedor (sobre el resultado ya depurado)
    if vendedor:
        v = vendedor.strip().lower()
        items = [i for i in items if v in (i["vendedor"] or "").lower()]

    items.sort(key=lambda i: (i["fecha_oc"], i["codigo"], i["orden_venta"]))
    return {"ok": True, "empresa": {"id": company_id, "nombre": COMPANY_FINPAC_NOMBRE},
            "items": items, "total": len(items),
            "vendedores": sorted(vendedores.keys()),
            "candidatas": candidatas_total,
            "incluidas": len(items),
            "excluidas": len(exclusiones),
            "exclusiones": exclusiones,
            "mensaje": "" if items else "No se encontraron devoluciones pendientes con los filtros aplicados."}


# --------------------------------------------------------------------------- #
# Detalle de movimientos que sustentan un caso (para revision)
# --------------------------------------------------------------------------- #
def detalle(company_id, move_id):
    company_id = COMPANY_FINPAC
    mv = _query("stock.move", [("id", "=", move_id), ("company_id", "=", company_id)],
                ["id", "name", "date", "product_id", "quantity", "picking_id", "sale_line_id"])
    if not mv:
        return {"ok": False, "mensaje": "Caso no encontrado."}
    m = mv[0]
    pid = _oid(m.get("product_id"))
    locs_usage = _ubicaciones_virtuales(company_id)
    movs = _query("stock.move",
                  [("company_id", "=", company_id), ("product_id", "=", pid), ("state", "=", "done")],
                  ["id", "name", "date", "quantity", "picking_type_id", "picking_id",
                   "sale_line_id", "location_id", "location_dest_id", "is_inventory",
                   "reference"], order="date asc, id asc")
    picks = _bulk("stock.picking", [_oid(x.get("picking_id")) for x in movs],
                  ["id", "name", "partner_id"])
    types = {t["id"]: t for t in _query("stock.picking.type", [], ["id", "sequence_code", "code"])}
    etiquetas = {
        "venta": "Salida / Venta", "entrada": "Entrada / Compra o ajuste (+)",
        "salida_ajuste": "Ajuste de inventario (\u2212) / desecho / producci\u00f3n",
        "devolucion_cliente": "Devoluci\u00f3n de cliente", "interno": "Reubicaci\u00f3n interna",
        "otro": "Otro movimiento",
    }
    filas = []
    for x in movs:
        pt = types.get(_oid(x.get("picking_type_id"))) or {}
        clase = _clasificar_movimiento(x, locs_usage)
        filas.append({
            "fecha": x["date"], "nombre": x["name"],
            "documento": (picks.get(_oid(x.get("picking_id"))) or {}).get("name") or "",
            "tipo": pt.get("sequence_code") or ("Ajuste/Interno" if not pt else pt.get("code")),
            "clase": clase,
            "clase_etiqueta": etiquetas.get(clase, clase),
            "es_salida": clase == "venta",
            "es_devolucion": clase == "devolucion_cliente",
            "cantidad": round(x.get("quantity") or 0.0, 4),
            "sale_line": _oid(x.get("sale_line_id")),
        })
    return {"ok": True, "producto_id": pid,
            "devolucion": {"id": m["id"], "nombre": m["name"], "fecha": m["date"][:10],
                           "cantidad": round(m.get("quantity") or 0.0, 4)},
            "movimientos": filas}


# --------------------------------------------------------------------------- #
# Exportacion a Excel (mismas columnas del reporte)
# --------------------------------------------------------------------------- #
def build_xlsx(data):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Devoluciones pendientes"

    head_fill = PatternFill("solid", fgColor="0F172A")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = ("DEVOLUCIONES PENDIENTES DE COMPRAS LOCALES \u2014 "
                f"{data['empresa']['nombre']}")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws["A2"] = f"Generado: {datetime.now():%d/%m/%Y %H:%M}"
    ws["A2"].font = Font(size=10, color="64748B")

    headers = ["CONTACTO", "ORDEN DE VENTA", "C\u00d3DIGO", "DESCRIPCI\u00d3N",
               "CANTIDAD DEVUELTA", "CANTIDAD VENDIDA DESPU\u00c9S", "CANTIDAD PENDIENTE",
               "ESTADO", "ORDEN DE COMPRA", "PROVEEDOR", "FECHA DE O/C",
               "COSTO UNIT.", "VENDEDOR"]
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
        vals = [it["contacto"], it["orden_venta"], it["codigo"], it["descripcion"],
                it["cantidad_devuelta"], it["cantidad_vendida_despues"], it["cantidad_pendiente"],
                it["estado"], it["orden_compra"], it["proveedor"], it["fecha_oc"],
                it["costo_unitario"], it["vendedor"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (1, 4, 10)))
            if c in (5, 6, 7):
                cell.number_format = "#,##0.00"
            if c == 8:
                cell.alignment = Alignment(vertical="top", horizontal="center")
            if c == 12:
                cell.number_format = '"$"#,##0.000000'
            if it.get("cantidad_estimada"):
                cell.fill = PatternFill("solid", fgColor="FFFBEB")

    if not data["items"]:
        ws.cell(row=head_row + 1, column=1, value="Sin devoluciones pendientes en el rango indicado.")

    widths = [26, 14, 12, 40, 14, 18, 16, 20, 16, 28, 12, 13, 30]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    if data["items"]:
        ws.auto_filter.ref = f"A{head_row}:M{head_row + len(data['items'])}"
    ws.row_dimensions[head_row].height = 30

    nr = head_row + len(data["items"]) + 3
    ws.cell(row=nr, column=1, value="NOTAS").font = Font(bold=True, size=11)
    notas = [
        "Solo compras locales / puntuales de FINPAC. Se excluyen las compras de importaci\u00f3n.",
        "CANTIDAD DEVUELTA: unidades que devolvi\u00f3 originalmente el cliente (movimientos DEVCL).",
        "CANTIDAD VENDIDA DESPU\u00c9S: unidades de la devoluci\u00f3n que salieron a cliente despu\u00e9s "
        "de la devoluci\u00f3n (rastreo cronol\u00f3gico de movimientos).",
        "CANTIDAD PENDIENTE: unidades de la devoluci\u00f3n que siguen en inventario (devuelta \u2212 vendida). "
        "No es el stock total del producto.",
        "ESTADO: PENDIENTE (nada vendido), PARCIALMENTE VENDIDO o VENDIDO COMPLETAMENTE (excluido del reporte).",
        "Las filas resaltadas tuvieron stock nuevo posterior (compras o ajustes del mismo c\u00f3digo): "
        "la cantidad pendiente es una estimaci\u00f3n por asignaci\u00f3n FIFO, no certeza f\u00edsica de las mismas unidades.",
        "Las transferencias internas no alteran la disponibilidad y no se imputan a la devoluci\u00f3n.",
        "El vendedor proviene de la orden de venta original de la devoluci\u00f3n.",
        "Este m\u00f3dulo no genera cobros ni modifica el ERP.",
    ]
    for i, n in enumerate(notas, start=1):
        ws.cell(row=nr + i, column=1, value=f"{i}. {n}")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
