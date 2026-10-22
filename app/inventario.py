"""Motor de inventario del Kardex: clasificacion de movimientos por ubicacion,
conversion de unidades de medida, calculo de saldos y conciliacion contra Odoo.

Las reglas provienen de una auditoria hecha sobre los datos reales del ERP
(Odoo 18.0+e) y corrigen el criterio anterior, que clasifiaba los movimientos
por el codigo del tipo de picking y por la existencia de capas de valorizacion.

ALCANCE (que es el stock fisico de una empresa)
    El alcance son las ubicaciones con ``usage='internal'`` y
    ``company_id = <empresa>``.  El efecto de un movimiento sobre ese alcance
    depende solo de sus dos ubicaciones:

        origen interno -> destino interno : traslado      (efecto 0)
        origen interno -> destino fuera   : salida        (efecto -cantidad)
        origen fuera   -> destino interno : entrada       (efecto +cantidad)
        ninguno en el alcance             : externo       (efecto 0)

    Como ``stock.move.quantity`` es siempre positiva, la direccion del par de
    ubicaciones es la que define el signo.  Las operaciones multietapa
    (PICK -> PACK -> OUT) no se duplican: PICK y PACK son interno->interno y
    solo OUT abandona el alcance.  Las ubicaciones virtuales (inventory,
    transit, production, view) no pertenecen al alcance, de modo que un ajuste
    fisico se contabiliza por su efecto real (entra o sale de una ubicacion
    interna) y nunca como una operacion de almacen.

UNIDADES DE MEDIDA
    ``stock.move.quantity`` esta expresado en el ``product_uom`` del movimiento
    mientras que ``stock.valuation.layer.quantity`` lo esta en la unidad base
    del producto.  La conversion se resuelve en local con los factores
    configurados en Odoo (``uom.uom.factor_inv`` relativo a la unidad de
    referencia de la categoria) y se aplica el ``rounding`` de la unidad
    destino.  No se invoca ``uom.uom._compute_quantity`` porque es un metodo
    privado que no puede llamarse por XML-RPC.  Si las unidades pertenecen a
    categorias distintas la conversion se rechaza de forma explicita.

VALORIZACION
    El valor contable sale de ``stock.valuation.layer`` y las cantidades
    fisicas nunca se toman de las capas: son la cantidad convertida del
    movimiento.  Cuando el movimiento no tiene capa (productos sin valuacion
    en tiempo real, o movimientos entre ubicaciones virtuales) la fila se
    valora al costo promedio o estandar y se marca como estimada.
"""
import os
import threading
import time
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_CEILING, ROUND_DOWN, ROUND_HALF_UP

# --------------------------------------------------------------------------- #
# Dependencias inyectadas (mismo patron que devoluciones.py)
# --------------------------------------------------------------------------- #
_sr = None
_id_of = None
_ref = None
_url = None
_lock = threading.RLock()
_cache = {}
_ordenes = {}

# Clases de movimiento respecto del alcance de la empresa
CLASE_ENTRADA = "entrada"
CLASE_SALIDA = "salida"
CLASE_TRASLADO = "traslado"
CLASE_EXTERNO = "externo"

REDONDEOS = {"HALF-UP": ROUND_HALF_UP, "UP": ROUND_CEILING, "DOWN": ROUND_DOWN}


def configurar(sr, id_of, ref, url=None):
    """Inyecta las funciones de acceso al ERP (patron de devoluciones.py)."""
    global _sr, _id_of, _ref, _url
    _sr, _id_of, _ref = sr, id_of, ref
    if url:
        _url = url


def _metodo_redondeo():
    return os.environ.get("KARDEX_REDONDEO", "HALF-UP").upper()


# --------------------------------------------------------------------------- #
# Catalogos
# --------------------------------------------------------------------------- #
def uoms(refrescar=False):
    """Catalogo de unidades de medida: id -> {nombre, categoria, factor, factor_inv, precision}."""
    with _lock:
        if refrescar or "uoms" not in _cache:
            filas = _sr("uom.uom", [], ["id", "name", "category_id", "factor", "factor_inv",
                                        "rounding", "uom_type", "active"])
            cat = {}
            for f in filas:
                cat[f["id"]] = {
                    "id": f["id"],
                    "nombre": f.get("name") or "",
                    "category_id": _id_of(f.get("category_id")),
                    "factor": f.get("factor") or 1.0,
                    "factor_inv": f.get("factor_inv") or 1.0,
                    "precision": f.get("rounding") or 0.01,
                    "tipo": f.get("uom_type") or "",
                    "activo": bool(f.get("active", True)),
                }
            _cache["uoms"] = cat
        return _cache["uoms"]


def ubicaciones(refrescar=False):
    """Catalogo de ubicaciones: id -> {nombre, usage, company_id}."""
    with _lock:
        if refrescar or "locs" not in _cache:
            filas = _sr("stock.location", [], ["id", "name", "complete_name", "usage",
                                               "company_id", "active"])
            cat = {}
            for f in filas:
                cat[f["id"]] = {
                    "id": f["id"],
                    "nombre": f.get("name") or "",
                    "complete_name": f.get("complete_name") or f.get("name") or "",
                    "usage": f.get("usage") or "",
                    "company_id": _id_of(f.get("company_id")),
                }
            _cache["locs"] = cat
        return _cache["locs"]


def almacenes(empresa, refrescar=False):
    """Ids de las ubicaciones internas que forman el alcance de la empresa."""
    with _lock:
        clave = ("alm", empresa)
        if refrescar or clave not in _cache:
            locs = ubicaciones(refrescar)
            _cache[clave] = {i for i, l in locs.items()
                             if l.get("usage") == "internal" and l.get("company_id") == empresa}
        return _cache[clave]


def invalidar_catalogos():
    with _lock:
        _cache.pop("uoms", None)
        _cache.pop("locs", None)
        for k in [k for k in _cache if isinstance(k, tuple) and k and k[0] == "alm"]:
            _cache.pop(k, None)


def nombre_unidad(uom_id):
    cat = uoms()
    u = cat.get(uom_id)
    return u["nombre"] if u else "Unidades"


# --------------------------------------------------------------------------- #
# Conversion de unidades
# --------------------------------------------------------------------------- #
def _exponente(precision):
    """Exponente decimal de la precision (0.01 -> -2)."""
    try:
        d = Decimal(str(precision)).normalize()
    except Exception:
        return -2
    return d.as_tuple().exponent


def redondear(valor, precision, metodo=None):
    """Redondea a la precision de la unidad de medida (Decimal, sin error de coma flotante)."""
    if valor is None:
        return 0.0
    metodo = metodo or _metodo_redondeo()
    modo = REDONDEOS.get(metodo, ROUND_HALF_UP)
    exp = _exponente(precision or 0.01)
    try:
        q = Decimal(str(valor)).quantize(Decimal(1).scaleb(exp), rounding=modo)
    except Exception:
        return float(valor)
    return float(q)


def convertir(cantidad, uom_origen, uom_destino, metodo=None):
    """Convierte una cantidad entre unidades usando los factores de Odoo.

    Devuelve un diccionario con el resultado y la trazabilidad:
    ``ok``, ``cantidad`` (normalizada), ``cantidad_original``, ``factor``,
    ``precision``, ``motivo`` (vacio si la conversion fue valida) y
    ``categoria_compatible``.
    """
    res = {"ok": False, "cantidad": 0.0, "cantidad_original": float(cantidad or 0.0),
           "factor": None, "precision": None, "motivo": "", "categoria_compatible": False,
           "uom_origen": uom_origen, "uom_destino": uom_destino}
    if cantidad is None:
        res["motivo"] = "sin_cantidad"
        return res
    cat = uoms()
    o, d = cat.get(uom_origen), cat.get(uom_destino)
    if o is None or d is None:
        res["motivo"] = "unidad_desconocida"
        return res
    if o["category_id"] != d["category_id"]:
        res["motivo"] = "categorias_incompatibles"
        return res
    res["categoria_compatible"] = True
    if o["id"] == d["id"]:
        res.update(ok=True, cantidad=float(cantidad), factor=1.0, precision=d["precision"])
        return res
    fi = o["factor_inv"] or 1.0
    fd = d["factor_inv"] or 1.0
    if not fd:
        res["motivo"] = "factor_invalido"
        return res
    factor = fi / fd
    res.update(ok=True, cantidad=redondear(float(cantidad) * factor, d["precision"], metodo),
               factor=factor, precision=d["precision"])
    return res


# --------------------------------------------------------------------------- #
# Clasificacion de movimientos por ubicacion
# --------------------------------------------------------------------------- #
def _sub_entrada(uso_origen, locs):
    if uso_origen == "customer":
        return "devolucion_cliente"
    if uso_origen == "supplier":
        return "compra"
    if uso_origen == "inventory":
        return "ajuste_positivo"
    if uso_origen == "transit":
        return "traspaso_entrada"
    if uso_origen == "production":
        return "produccion"
    if uso_origen == "view":
        return "vista"
    return "entrada"


def _sub_salida(uso_destino, locs):
    if uso_destino == "customer":
        return "venta"
    if uso_destino == "supplier":
        return "devolucion_proveedor"
    if uso_destino == "inventory":
        return "ajuste_negativo"
    if uso_destino == "transit":
        return "traspaso_salida"
    if uso_destino == "production":
        return "consumo_produccion"
    if uso_destino == "view":
        return "vista"
    return "salida"


def clasificar(mov, alm, locs=None):
    """Clasifica un movimiento respecto del alcance `alm` (ids de ubicaciones internas).

    Devuelve (clase, subtipo, efecto) donde el efecto ya viene con signo y en
    la unidad base del producto.  La clase depende unicamente de las
    ubicaciones de origen y destino, nunca del tipo de picking.
    """
    locs = locs if locs is not None else ubicaciones()
    src = _id_of(mov.get("location_id"))
    dst = _id_of(mov.get("location_dest_id"))
    src_int = src in alm
    dst_int = dst in alm
    l_src = locs.get(src, {})
    l_dst = locs.get(dst, {})
    if src_int and dst_int:
        return CLASE_TRASLADO, "traslado", 0.0
    if dst_int:
        sub = _sub_entrada(l_src.get("usage"), locs)
        if l_src.get("usage") == "internal":
            sub = "traspaso_entre_empresas"
        return CLASE_ENTRADA, sub, 1.0
    if src_int:
        sub = _sub_salida(l_dst.get("usage"), locs)
        if l_dst.get("usage") == "internal":
            sub = "traspaso_entre_empresas"
        return CLASE_SALIDA, sub, -1.0
    return CLASE_EXTERNO, "externo", 0.0


ETIQUETAS = {
    "compra": "Ingreso (Recepción)",
    "devolucion_cliente": "Devolución de cliente",
    "venta": "Venta (Despacho)",
    "devolucion_proveedor": "Devolución a proveedor",
    "ajuste_positivo": "Ajuste de inventario",
    "ajuste_negativo": "Ajuste de inventario",
    "traspaso_entrada": "Traspaso hacia almacén",
    "traspaso_salida": "Traspaso desde almacén",
    "traspaso_entre_empresas": "Traspaso entre empresas",
    "produccion": "Producción (producto terminado)",
    "consumo_produccion": "Producción (consumo de material)",
    "vista": "Movimiento de vista",
    "entrada": "Entrada de inventario",
    "salida": "Salida de inventario",
    "traslado": "Transferencia interna",
    "externo": "Movimiento externo",
}


def etiqueta(clase, subtipo, seq_picking=None):
    """Etiqueta legible del movimiento (la clase NO se deduce de aqui)."""
    if clase == CLASE_TRASLADO:
        return ETIQUETAS["traslado"]
    if clase == CLASE_EXTERNO:
        return ETIQUETAS["externo"]
    if subtipo == "compra" and seq_picking == "IMP":
        return "Ingreso (Importación)"
    return ETIQUETAS.get(subtipo, ETIQUETAS.get(clase, "Movimiento"))


# --------------------------------------------------------------------------- #
# Calculo del kardex
# --------------------------------------------------------------------------- #
def _redondeo_unidad(valor, precision):
    return redondear(valor, precision)


def quants_del_alcance(quants, alm, uom_base=None, precision=0.01):
    """Suma el stock fisico de los quants que pertenecen al alcance.

    Solo se toman los quants cuya ubicacion esta en `alm` (ubicaciones internas
    de la empresa): los quants de clientes, proveedores, ajustes o transito
    netean a cero y no representan existencias del almacen.

    Las cantidades se normalizan a la unidad base del producto: `stock.quant`
    ya las guarda en esa unidad, pero si el quant trae otra se convierte con los
    factores de unidades del ERP.

    Devuelve (fisico, reservado, detalle) con `detalle` una lista de dicts
    {"ubicacion": id, "cantidad": float, "reservado": float}.
    """
    fisico = reservado = 0.0
    detalle = []
    for q in quants or []:
        lid = _id_of(q.get("location_id"))
        if lid not in alm:
            continue
        cant_q = q.get("quantity") or 0.0
        cant_r = q.get("reserved_quantity") or 0.0
        uom_q = q.get("product_uom") or q.get("uom_id") or uom_base
        if uom_q and uom_base and uom_q != uom_base:
            conv_q = convertir(cant_q, uom_q, uom_base)
            if conv_q.get("ok"):
                cant_q = conv_q["cantidad"]
            conv_r = convertir(cant_r, uom_q, uom_base)
            if conv_r.get("ok"):
                cant_r = conv_r["cantidad"]
        fisico += cant_q
        reservado += cant_r
        detalle.append({"ubicacion": lid, "cantidad": cant_q, "reservado": cant_r})
    return (_redondeo_unidad(fisico, precision),
            _redondeo_unidad(reservado, precision),
            detalle)


def calcular_kardex(empresa, producto, movimientos, capas, alm, locs=None,
                    quants=None, notas_extra=None):
    """Arma el kardex de un producto a partir de sus movimientos completados.

    `movimientos` : lista de dicts con al menos id, date, quantity, product_uom,
                    location_id, location_dest_id, state.
    `capas`       : dict move_id -> [{"quantity":..,"value":..,"unit_cost":..}]
    `alm`         : ids de ubicaciones internas del alcance.
    `quants`      : opcional, [{"location_id":..,"quantity":..,"reserved_quantity":..}]
    """
    locs = locs if locs is not None else ubicaciones()
    uom_base = _id_of(producto.get("uom_id"))
    unidad = nombre_unidad(uom_base)
    precision = uoms().get(uom_base, {}).get("precision", 0.01)
    prod_id = producto.get("id")

    filas, traslados, externos, avisos = [], [], [], []
    saldo = 0.0
    valor = 0.0
    valor_entradas = 0.0
    qty_entradas = 0.0
    costo_estandar = producto.get("standard_price") or 0.0

    for mov in sorted(movimientos, key=lambda m: (str(m.get("date") or ""), m.get("id") or 0)):
        if mov.get("state") != "done":
            continue
        clase, subtipo, signo = clasificar(mov, alm, locs)

        # --- cantidad fisica normalizada a la unidad base del producto ---
        uom_mov = _id_of(mov.get("product_uom")) or uom_base
        cant = mov.get("quantity") or 0.0
        conv = convertir(cant, uom_mov, uom_base)
        if conv["ok"]:
            cantidad = conv["cantidad"]
        else:
            cantidad = mov.get("product_qty")
            if cantidad is None:
                cantidad = cant
            avisos.append({
                "tipo": "unidad_no_convertible",
                "movimiento": mov.get("name") or mov.get("id"),
                "detalle": conv["motivo"],
            })

        # --- efecto sobre el alcance ---
        efecto = signo * cantidad
        doc_mov = mov.get("name") or ""
        l_src = locs.get(_id_of(mov.get("location_id")), {})
        l_dst = locs.get(_id_of(mov.get("location_dest_id")), {})
        origen_txt = l_src.get("complete_name") or "N/A"
        destino_txt = l_dst.get("complete_name") or "N/A"

        if clase == CLASE_TRASLADO:
            traslados.append({
                "date": mov.get("date"), "tipo": mov.get("seq_picking") or "TRASLADO",
                "doc": mov.get("picking_name") or doc_mov,
                "cantidad": cantidad, "ubicacion": destino_txt,
                "origen": f"{origen_txt} -> {destino_txt}",
                "subtipo": subtipo, "usuario": mov.get("usuario") or "",
            })
            continue
        if clase == CLASE_EXTERNO:
            externos.append({"date": mov.get("date"), "nombre": doc_mov,
                             "cantidad": cantidad, "origen": origen_txt,
                             "destino": destino_txt,
                             "usuario": mov.get("usuario") or ""})
            continue

        # --- valoracion (capas reales del ERP) ---
        lay = capas.get(mov.get("id"), [])
        lq = sum((l.get("quantity") or 0.0) for l in lay)
        lval = sum((l.get("value") or 0.0) for l in lay)
        estimado = False
        if lay and abs(lq) > 1e-9:
            costo = lval / lq
            valor_mov = lval
        else:
            estimado = True
            costo = costo_estandar or 0.0
            valor_mov = efecto * costo

        saldo += efecto
        valor += valor_mov
        if efecto > 0:
            valor_entradas += valor_mov
            qty_entradas += efecto
        elif efecto < 0:
            valor_entradas += valor_mov
            qty_entradas += efecto

        filas.append({
            "date": mov.get("date"),
            "tipo": etiqueta(clase, subtipo, mov.get("seq_picking")),
            "clase": clase,
            "subtipo": subtipo,
            "doc": mov.get("doc") or mov.get("picking_name") or doc_mov,
            "origen": (f"Origen: {origen_txt}  →  Destino: {destino_txt}"
                       + (f"  ·  {mov.get('contraparte')}" if mov.get("contraparte") else "")),
            "usuario": mov.get("usuario") or "",
            "entrada": efecto if efecto > 0 else 0.0,
            "salida": -efecto if efecto < 0 else 0.0,
            "qty": round(efecto, 6),
            "cantidad": round(abs(efecto), 6),
            "cantidad_original": cant,
            "uom_original": nombre_unidad(uom_mov),
            "cost": round(costo, 6),
            "valor": round(valor_mov, 4),
            "valor_total": round(valor, 2),
            "saldo": round(saldo, 6),
            "costo_estimado": estimado,
            "uom": unidad,
            "referencias": mov.get("referencias") or [],
        })

    saldo = _redondeo_unidad(saldo, precision)
    valor = round(valor, 2)
    costo_prom_actual = (valor / saldo) if saldo > 0 else None
    costo_pond = (valor_entradas / qty_entradas) if qty_entradas else None

    # Stock fisico: solo los quants de las ubicaciones internas del alcance
    # (los quants de clientes, proveedores o ajustes netean a cero y no
    # representan existencias del almacen).
    fisico, reservado, _detalle_q = quants_del_alcance(quants, alm, uom_base, precision)

    kpi = {
        "entradas": round(sum(f["entrada"] for f in filas), 4),
        "salidas": round(sum(f["salida"] for f in filas), 4),
        "saldo": round(saldo, 4),
        "valor": valor,
        "costo_prom": round(costo_prom_actual, 6) if costo_prom_actual else 0.0,
        "costo_prom_ponderado": round(costo_pond, 6) if costo_pond else 0.0,
        "movimientos": len(filas),
        "internos": len(traslados),
        "internos_qty": round(sum(t["cantidad"] for t in traslados), 4),
        "externos": len(externos),
        "fisico": round(fisico, 4),
        "reservado": round(reservado, 4),
        "disponible": round(fisico - reservado, 4),
        "cancelados": 0,
        "avisos": avisos,
    }
    return {"rows": filas, "internals": traslados, "externos": externos,
            "kpi": kpi, "unidad": unidad, "uom_id": uom_base,
            "precision": precision, "notas_extra": notas_extra or []}


# --------------------------------------------------------------------------- #
# Conciliacion contra el stock fisico de Odoo
# --------------------------------------------------------------------------- #
def conciliar(empresa, producto, saldo_kardex, quants, movimientos=None,
              capas=None, alm=None, uom_base=None, locs=None):
    """Compara el saldo calculado por movimientos con el stock fisico de Odoo.

    El stock de referencia es la suma de `stock.quant` **exclusivamente** de
    las ubicaciones internas del alcance de la empresa y en la unidad base del
    producto.  No se usa `qty_available` porque mezcla empresas y alcances.
    """
    locs = locs if locs is not None else ubicaciones()
    alm = alm if alm is not None else almacenes(empresa)
    uom_base = uom_base or _id_of(producto.get("uom_id"))
    unidad = nombre_unidad(uom_base)
    precision = uoms().get(uom_base, {}).get("precision", 0.01)

    fisico, reservado, detalle_q = quants_del_alcance(quants, alm, uom_base, precision)
    detalle_quants = [{"ubicacion": locs.get(d["ubicacion"], {}).get("complete_name", ""),
                        "cantidad": d["cantidad"],
                        "reservado": d["reservado"]}
                       for d in detalle_q]
    diferencia = _redondeo_unidad(float(saldo_kardex) - fisico, precision)
    rel = (diferencia / fisico) if fisico else None

    if abs(diferencia) < max(precision / 2.0, 1e-9):
        estado = "conciliado"
    else:
        estado = "diferencia"

    # Movimientos sospechosos: los que tocan ubicaciones virtuales o que no
    #quares con el alcance y podrian explicar la diferencia.
    sospechosos = []
    if movimientos:
        for mov in movimientos:
            if mov.get("state") != "done":
                continue
            clase, subtipo, signo = clasificar(mov, alm, locs)
            uso_s = locs.get(_id_of(mov.get("location_id")), {}).get("usage")
            uso_d = locs.get(_id_of(mov.get("location_dest_id")), {}).get("usage")
            toca_virtual = (uso_s in ("inventory", "transit", "production", "view")
                            or uso_d in ("inventory", "transit", "production", "view"))
            if toca_virtual:
                sospechosos.append({
                    "id": mov.get("id"), "fecha": mov.get("date"),
                    "nombre": mov.get("name") or "",
                    "clase": clase, "subtipo": subtipo,
                    "cantidad": mov.get("quantity") or 0.0,
                    "origen": locs.get(_id_of(mov.get("location_id")), {}).get("complete_name", ""),
                    "destino": locs.get(_id_of(mov.get("location_dest_id")), {}).get("complete_name", ""),
                    "motivo": "toca ubicacion virtual",
                })
    # Quantidades fuera del alcance (inventario, transito) del mismo producto
    fuera = []
    for q in quants or []:
        lid = _id_of(q.get("location_id"))
        if lid in alm:
            continue
        loc = locs.get(lid, {})
        if loc.get("usage") in ("inventory", "transit", "production", "view"):
            fuera.append({"ubicacion": loc.get("complete_name", ""),
                          "uso": loc.get("usage"), "cantidad": q.get("quantity") or 0.0})

    causas = []
    sin_movimientos = not tiene_movimientos_en_alcance(movimientos, alm, locs)
    if sin_movimientos:
        causas.append("No hay movimientos completados que muevan el producto dentro del "
                      "almacen; el saldo historico podria estar incompleto.")
    if estado == "diferencia":
        if sospechosos:
            causas.append("Hay %d movimiento(s) completados que tocan ubicaciones virtuales "
                          "(ajuste, transito, produccion): su efecto real puede explicar la "
                          "diferencia." % len(sospechosos))
        if fuera:
            causas.append("El producto tiene %d cantidad(es) en ubicaciones fuera del alcance "
                          "(inventario/transito): revise si deben estar en el almacen." % len(fuera))
        if not sin_movimientos:
            causas.append("Compare el historico de movimientos con el conteo fisico del almacen "
                          "para identificar el origen exacto de la diferencia.")

    return {
        "codigo": producto.get("default_code") or "",
        "descripcion": producto.get("name") or "",
        "unidad": unidad,
        "stock_odoo": round(fisico, 4),
        "stock_kardex": round(float(saldo_kardex), 4),
        "diferencia": round(diferencia, 4),
        "diferencia_relativa": round(rel, 6) if rel is not None else None,
        "estado": estado,
        "posibles_causas": causas,
        "movimientos_sospechosos": sospechosos[:50],
        "quants_fuera_de_alcance": fuera[:50],
        "detalle_quants": detalle_quants[:50],
        "reservado": round(reservado, 4),
    }


def tiene_movimientos_en_alcance(movimientos, alm, locs=None):
    """True si el producto tiene movimientos que entran o salen del alcance."""
    locs = locs if locs is not None else ubicaciones()
    for mov in movimientos or []:
        if mov.get("state") != "done":
            continue
        clase, _, _ = clasificar(mov, alm, locs)
        if clase in (CLASE_ENTRADA, CLASE_SALIDA):
            return True
    return False
