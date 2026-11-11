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
from datetime import date, datetime, time as dt_time, timedelta
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
            })
        capas = defaultdict(list)
        for c in svc.capas_producto(product_id):
            capas[c["move_id"]].append({"quantity": c["cantidad"], "value": c["valor"],
                                        "unit_cost": c["costo"]})
        return movs, dict(capas), "espejo local (SQLite)"

    moves = sr("stock.move", [("product_id", "=", product_id), ("state", "=", "done")],
               ["id", "name", "date", "quantity", "product_qty", "product_uom", "company_id",
                "location_id", "location_dest_id", "state", "picking_id", "picking_type_id",
                "purchase_line_id", "sale_line_id", "origin_returned_move_id", "origin"],
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
             "referencia": m.get("origin") or ""} for m in moves]
    capas = defaultdict(list)
    for l in sr("stock.valuation.layer", [("product_id", "=", product_id)],
                ["stock_move_id", "quantity", "unit_cost", "value"]):
        capas[id_of(l["stock_move_id"])].append({"quantity": l.get("quantity") or 0.0,
                                                "value": l.get("value") or 0.0,
                                                "unit_cost": l.get("unit_cost") or 0.0})
    return movs, dict(capas), "consulta directa a Odoo"


def _enriquecer_movimientos(movs):
    """Agrega documento, picking y contraparte a cada movimiento."""
    if not movs:
        return movs
    picks = _by_ids("stock.picking", [m.get("picking_id") for m in movs],
                    ["id", "name", "partner_id", "origin", "state"])
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
    return movs


def build_kardex(company_id, product_id, refresh=False):
    """Arma el kardex de un producto.

    La clasificacion de entradas, salidas y transferencias la hace el motor de
    inventario (`app/inventario.py`) a partir de las ubicaciones de origen y
    destino y del alcance de la empresa (ubicaciones internas con
    company_id = empresa).  No se usa el codigo del tipo de picking para decidir
    el signo ni la existencia de capas de valorizacion.

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
              ["cost_method", "valuation", "standard_price"])[0]

    alm = inventario.almacenes(company_id)
    locs = inventario.ubicaciones()
    movimientos, capas, origen_datos = _datos_movimientos(product_id)
    movimientos = _enriquecer_movimientos(movimientos)
    quants = _quants(product_id)

    calc = inventario.calcular_kardex(company_id, prod, movimientos, capas, alm, locs,
                                      quants=quants)
    estado_sync = servicio_sync().estado()
    conciliacion = inventario.conciliar(company_id, prod, calc["kpi"]["saldo"], quants,
                                        movimientos=movimientos, capas=capas, alm=alm)
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
                                "standard": "Costo estÃ¡ndar"}.get(tmpl["cost_method"], tmpl["cost_method"]),
                     "valuacion": tmpl["valuation"]},
        "kpi": calc["kpi"],
        "rows": calc["rows"],
        "internals": calc["internals"],
        "externos": calc["externos"],
        "conciliacion": conciliacion,
        "origen_datos": origen_datos,
        "sincronizacion": estado_sync,
    }
    if not calc["rows"]:
        payload.update({"ok": False,
                        "sin_movimientos": f"No se encontraron movimientos registrados para el Ã­tem "
                                           f"{prod['default_code']} - {prod['name']} en el perÃ­odo "
                                           f"indicado para {payload['empresa']['nombre']}.",
                        "cancelados": 0, "total_moves": len(movimientos)})
    else:
        payload["desde"], payload["hasta"] = calc["rows"][0]["date"][:10], calc["rows"][-1]["date"][:10]
    payload["notas"] = [
        f"Fuente de datos: {origen_datos}. Movimientos completados (stock.move, estado done) "
        f"y capas de valorizaciÃ³n (stock.valuation.layer) del producto {product_id}.",
        f"Unidad base del producto: {calc['unidad']} (precisiÃ³n {calc['precision']}). Las "
        f"cantidades de los movimientos se convertieron a esta unidad con los factores "
        f"configurados en Odoo; las conversiones entre categorÃ­as incompatibles se rechazan.",
        f"MÃ©todo de valuaciÃ³n del ERP: {payload['producto']['metodo']}; el valor contable de cada "
        f"movimiento corresponde a su capa de valorizaciÃ³n. Las cantidades fÃ­sicas NO se toman "
        f"de las capas.",
        f"Alcance: {len(alm)} ubicaciones internas de la empresa {payload['empresa']['nombre']}. "
        f"El signo de cada movimiento se determina por su ubicaciÃ³n de origen y destino "
        f"(entrada / salida / transferencia interna); los movimientos entre ubicaciones virtuales "
        f"(ajuste, trÃ¡nsito, producciÃ³n) no alteran el stock del almacÃ©n.",
        f"Transferencias internas excluidas del detalle: {calc['kpi']['internos']} movimientos "
        f"({calc['kpi']['internos_qty']:,.2f} {calc['unidad']}); se listan en la hoja secundaria del Excel.",
        f"ConciliaciÃ³n contra el stock fÃ­sico de Odoo: {calc['kpi']['fisico']:,.2f} "
        f"{calc['unidad']} (quants del alcance) frente a {calc['kpi']['saldo']:,.2f} "
        f"{calc['unidad']} calculados por movimientos. Diferencia: "
        f"{conciliacion['diferencia']:,.2f} ({conciliacion['estado']}).",
    ] + (calc.get("notas_extra") or [])
    with _lock:
        _cache["kardex"][key] = (time.time(), payload)
    return payload


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
    ids = svc.productos_con_movimientos(limite)
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
                                    movimientos=movs, capas=capas, alm=alm)
        conc["producto_id"] = pid
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
        if q and q not in (r["doc"] + " " + r["origen"] + " " + r["tipo"]).lower():
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
        texto = " ".join(str(i.get(k) or "") for k in ("doc", "ubicacion", "origen", "tipo")).lower()
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
            vals = [f, mv["tipo"], mv["doc"], mv["origen"], mv["entrada"], mv["salida"],
                    mv["cost"], mv["saldo"], mv["valor_total"]]
            tot_e += mv["entrada"]
            tot_s += mv["salida"]
            ultimo_kardex = mv
        else:
            vals = [f, "Transferencia interna", mv["doc"],
                    " \u2192 ".join(x for x in (mv.get("ubicacion"), mv.get("origen")) if x),
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
        h2 = ["Fecha", "Tipo", "Num. Doc.", "Cantidad", "Ubicaci\u00f3n destino", "Origen"]
        for c, h in enumerate(h2, start=1):
            cell = w2.cell(row=1, column=c, value=h)
            cell.fill = HEAD_FILL
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.border = BORDER
        for i, mv in enumerate(payload["internals"], start=2):
            f = mv["date"].split("-")[2] + "/" + mv["date"].split("-")[1] + "/" + mv["date"].split("-")[0]
            for c, v in enumerate([f, mv["tipo"], mv["doc"], mv["cantidad"], mv["ubicacion"], mv["origen"]], start=1):
                cell = w2.cell(row=i, column=c, value=v)
                cell.border = BORDER
                if c == 4:
                    cell.number_format = "#,##0.00"
        for c, w in enumerate([12, 10, 32, 12, 40, 22], start=1):
            w2.column_dimensions[get_column_letter(c)].width = w
        w2.freeze_panes = "A2"
        w2.auto_filter.ref = f"A1:F{1 + len(payload['internals'])}"

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
        cells = [f, mv["tipo"], mv["doc"], mv["origen"], f"{mv['entrada']:.2f}", f"{mv['salida']:.2f}",
                 f"{mv['cost']:.4f}", f"{mv['saldo']:.2f}", f"{mv['valor_total']:.2f}"]
        out.append(",".join('"' + c.replace('"', '""') + '"' for c in cells))
    return ("\ufeff" + "\n".join(out)).encode("utf-8")


_TZ_GUAYAQUIL = timedelta(hours=-5)
_HORARIO_INICIO = dt_time(7, 0)
_HORARIO_FIN = dt_time(16, 0)
_FERIADOS_CACHE = {}


def _pascua(anio):
    a = anio % 19
    b = anio // 100
    c = anio % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(anio, mes, dia)


def _trasladar_feriado(d, trasladable):
    if not trasladable:
        return d
    dia_semana = d.weekday()
    if dia_semana == 1:
        return d - timedelta(days=1)
    if dia_semana in (2, 3):
        return d + timedelta(days=4 - dia_semana)
    if dia_semana == 5:
        return d - timedelta(days=1)
    if dia_semana == 6:
        return d + timedelta(days=1)
    return d


def _feriados_ecuador(anio):
    cache = _FERIADOS_CACHE.get(anio)
    if cache is not None:
        return cache
    pascua = _pascua(anio)
    base = [
        (date(anio, 1, 1), False),
        (pascua - timedelta(days=48), False),
        (pascua - timedelta(days=47), False),
        (pascua - timedelta(days=2), False),
        (date(anio, 5, 1), True),
        (date(anio, 5, 24), True),
        (date(anio, 7, 25), True),
        (date(anio, 8, 10), True),
        (date(anio, 10, 9), True),
        (date(anio, 11, 2), True),
        (date(anio, 11, 3), True),
        (date(anio, 12, 25), False),
    ]
    feriados = {_trasladar_feriado(d, trasladable) for d, trasladable in base}
    _FERIADOS_CACHE[anio] = feriados
    return feriados


def _es_dia_habil(d):
    return d.weekday() < 5 and d not in _feriados_ecuador(d.year)


def _siguiente_dia_habil(d):
    d = d + timedelta(days=1)
    while not _es_dia_habil(d):
        d += timedelta(days=1)
    return d


def _minutos_habiles(desde_utc, hasta_utc):
    ini = desde_utc + _TZ_GUAYAQUIL
    fin = hasta_utc + _TZ_GUAYAQUIL
    if fin <= ini:
        return 0

    d = ini.date()
    hora = ini.time()
    if _es_dia_habil(d):
        if hora < _HORARIO_INICIO:
            cursor = datetime.combine(d, _HORARIO_INICIO)
        elif hora < _HORARIO_FIN:
            cursor = ini
        else:
            cursor = datetime.combine(_siguiente_dia_habil(d), _HORARIO_INICIO)
    else:
        cursor = datetime.combine(_siguiente_dia_habil(d), _HORARIO_INICIO)

    total = 0.0
    while cursor < fin:
        d = cursor.date()
        if _es_dia_habil(d):
            apertura = datetime.combine(d, _HORARIO_INICIO)
            cierre = datetime.combine(d, _HORARIO_FIN)
            tramo_ini = max(cursor, apertura)
            tramo_fin = min(fin, cierre)
            if tramo_fin > tramo_ini:
                total += (tramo_fin - tramo_ini).total_seconds()
        cursor = datetime.combine(d + timedelta(days=1), _HORARIO_INICIO)
    return int(total // 60)


def _fecha_hora_local(valor):
    if not valor:
        return ""
    try:
        return (datetime.strptime(valor, "%Y-%m-%d %H:%M:%S") + _TZ_GUAYAQUIL).strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return ""


def _fecha_local_corta(valor):
    """Fecha 'YYYY-MM-DD' en hora local de Ecuador a partir de un UTC 'YYYY-MM-DD HH:MM:SS'.

    Odoo guarda las fechas en UTC; recortar el texto sin convertir la zona
    horaria desplaza al dia siguiente todo el trabajo hecho despues de las
    19:00 local (UTC-5).
    """
    if not valor:
        return ""
    try:
        return (datetime.strptime(valor, "%Y-%m-%d %H:%M:%S") + _TZ_GUAYAQUIL).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return ""


_ESTADOS_ASIGNADO = ("assigned", "disponible", "listo", "ready")
_ESTADOS_HECHO = ("done", "hecho", "realizado", "completado")


def _parse_lista_usuarios(raw):
    """Ids y nombres exactos (en mayusculas) de una lista separada por comas."""
    uids, nombres = set(), set()
    for tok in (x.strip() for x in (raw or "").split(",")):
        if not tok:
            continue
        if tok.isdigit():
            uids.add(int(tok))
        else:
            nombres.add(tok.upper())
    return uids, nombres


# Usuarios que nunca se reportan como responsables de una fase de despacho:
# el personal de facturacion valida traslados de bodega y escribe sobre ellos
# el numero de autorizacion SRI, pero ese tramite administrativo no puede
# figurar como ejecucion de PICK / PACK / OUT en el tablero ni en el grafico.
# Se configura en el .env (despacho_excluir_usuarios) con ids o nombres
# exactos de Odoo, separados por comas.
_EXCL_RESP_IDS, _EXCL_RESP_NOMBRES = _parse_lista_usuarios(
    CFG.get("despacho_excluir_usuarios", ""))


def _resp_excluido(uid, nombre):
    return ((uid or 0) in _EXCL_RESP_IDS or
            (nombre or "").strip().upper() in _EXCL_RESP_NOMBRES)


# Ordenes que no se asignaron a ningun despachador y fueron ejecutadas
# directamente por estos usuarios (despacho_confirmacion_directa en el
# .env, ids o nombres de Odoo separados por comas): al no existir una
# asignacion que consultar, la referencia de inicio del pedido pasa a ser
# la fecha de confirmacion de la orden de venta (sale.order.date_order).
_CONF_DIRECTO_IDS, _CONF_DIRECTO_NOMBRES = _parse_lista_usuarios(
    CFG.get("despacho_confirmacion_directa", ""))


def _ejecutor_directo_conf(uid, nombre):
    return ((uid or 0) in _CONF_DIRECTO_IDS or
            (nombre or "").strip().upper() in _CONF_DIRECTO_NOMBRES)


# Una misma persona puede haber trabajado con mas de un usuario de Odoo (una
# cuenta generica tipo 'usuario002' y luego su cuenta propia): para el
# tablero y sus reportes se unifican en una sola identidad. Se configura en
# el .env (despacho_alias_usuarios) con pares "id_origen:id_canonico"
# separados por comas; el nombre mostrado es el del usuario canonico.
def _parse_alias_usuarios(raw):
    alias = {}
    for tok in (x.strip() for x in (raw or "").split(",")):
        if not tok or ":" not in tok:
            continue
        origen, canonico = (x.strip() for x in tok.split(":", 1))
        if origen.isdigit() and canonico.isdigit() and origen != canonico:
            alias[int(origen)] = int(canonico)
    return alias


_ALIAS_RESP_IDS = _parse_alias_usuarios(CFG.get("despacho_alias_usuarios", ""))
_ALIAS_NOMBRES = {}
_ALIAS_LOCK = threading.Lock()


def _nombre_canonico(uid):
    """Nombre actual del usuario canonico (cache en memoria)."""
    if uid not in _ALIAS_NOMBRES:
        nombre = None
        try:
            fila = sr("res.users", [("id", "=", uid)], ["name"])
            nombre = fila[0]["name"] if fila else None
        except Exception as exc:
            print(f"[despacho] No se pudo leer el usuario canonico {uid}: {exc}")
        with _ALIAS_LOCK:
            _ALIAS_NOMBRES[uid] = nombre
    return _ALIAS_NOMBRES[uid]


def _canon_usuario(uid, nombre=None):
    """(id, nombre) unificados de un usuario de Odoo.

    Si el id pertenece a una identidad duplicada se devuelve el usuario
    canonico y su nombre; en caso contrario los datos entran sin cambios.
    """
    canonico = _ALIAS_RESP_IDS.get(uid or 0)
    if canonico:
        return canonico, _nombre_canonico(canonico)
    return uid or 0, nombre


# Despachadores que figuran como responsables del tablero de despachos (el
# filtro "Despachador" de la pantalla solo ofrece esas personas; el resto de
# usuarios de Odoo no aparece). El calculo replica como el tablero decide el
# responsable de cada fase: asignado (stock.picking.user_id) o validador
# (write_uid de una escritura contemporanea a la validacion).
_DESPACHADORES_TTL = 900            # vigencia del cache, en segundos
_despachadores_cache = {"ts": 0.0, "datos": None}
_SISTEMA_UIDS = {1}                 # OdooBot: escribe en masa, nunca despacha


def _tipos_fases_despacho():
    """Tipos de operacion cuyos traslados forman las fases del tablero.

    El tablero identifica cada fase por el nombre del traslado (token
    PICK / PACK / OUT, que hereda del codigo de secuencia del tipo) o, en su
    defecto, por el nombre del tipo de operacion. Aqui se seleccionan los
    mismos tipos con esas mismas claves.
    """
    tipos = []
    for t in (G.get("picking_types") or {}).values():
        nombre = (t.get("name") or "").strip().lower()
        seq = (t.get("sequence_code") or "").strip().lower()
        if not nombre and not seq:
            continue
        if ("pick" in nombre
                or any(k in nombre for k in ("pack", "emp", "prep"))
                or any(k in nombre for k in ("out", "deliv", "salida", "entrega"))
                or seq in ("pick", "pack", "out")):
            tipos.append(t["id"])
    return sorted(tipos)


def _despachadores_tablero():
    """Personas que figuran como responsables del tablero de despachos.

    - Asignados: ``user_id`` de traslados de las fases con orden de venta y
      sin cancelar; el tablero los muestra como responsables cuando no hay
      rastro de quien valido.
    - Validadores: ``write_uid`` de traslados validados cuya escritura cae
      dentro de la ventana de validacion (|write_date - date_done| <= 120
      s, la misma que usa el respaldo validador_write del tablero). Al
      validar, Odoo deja write_uid en quien valido; una escritura posterior
      (p. ej. el numero SRI que anade facturacion) no puede figurar como
      ejecucion de bodega.
    """
    ahora = time.time()
    if (_despachadores_cache["datos"] is not None
            and ahora - _despachadores_cache["ts"] < _DESPACHADORES_TTL):
        return _despachadores_cache["datos"]
    datos = []
    tipos = _tipos_fases_despacho()
    empresas = [c["id"] for c in G.get("companies") or []]
    if tipos and empresas:
        base = [("picking_type_id", "in", tipos), ("sale_id", "!=", False),
                ("state", "!=", "cancel"), ("company_id", "in", empresas)]
        candidatos = {}
        try:
            # Asignados (read_group agrega en el ERP: una sola consulta).
            for g in _execute("stock.picking", "read_group",
                              [base, ["user_id"], ["user_id"]]):
                uid = id_of(g.get("user_id"))
                if uid and uid not in _SISTEMA_UIDS:
                    candidatos[uid] = ref(g.get("user_id"))
            # Candidatos a validador: ultimo escritor de traslados validados.
            solo_valid = {}
            for g in _execute("stock.picking", "read_group",
                              [base + [("state", "=", "done")],
                               ["write_uid"], ["write_uid"]]):
                uid = id_of(g.get("write_uid"))
                if (not uid or uid in _SISTEMA_UIDS or uid in candidatos
                        or _resp_excluido(uid, ref(g.get("write_uid")))):
                    continue
                solo_valid[uid] = ref(g.get("write_uid"))
            # Confirmacion del validador: al menos una escritura dentro de la
            # ventana de validacion. La lectura se acota (y ordena por fecha
            # descendente) para que quien escribe en masa sin validar, p. ej.
            # los planificadores de Odoo, quede descartado sin leer todo su
            # historial.
            for uid, nombre in solo_valid.items():
                filas = sr("stock.picking",
                           base + [("state", "=", "done"), ("write_uid", "=", uid)],
                           ["write_date", "date_done"],
                           order="write_date desc", limit=300)
                for p in filas:
                    try:
                        delta = abs((datetime.fromisoformat(p["write_date"]) -
                                     datetime.fromisoformat(p["date_done"])).total_seconds())
                    except (ValueError, TypeError):
                        continue
                    if delta <= 120:
                        candidatos[uid] = nombre
                        break
        except Exception as exc:
            print(f"[despacho] No se pudo calcular los despachadores del tablero: {exc}")
            return []
        unicos = {}
        for uid, nombre in candidatos.items():
            if _resp_excluido(uid, nombre):
                continue
            cid, cnom = _canon_usuario(uid, nombre)
            unicos[cid] = {"id": cid, "nombre": cnom or nombre or f"uid {cid}"}
        datos = sorted(unicos.values(), key=lambda x: (x["nombre"] or ""))
        _despachadores_cache.update(ts=ahora, datos=datos)
    return datos


def _valor_vacio(valor):
    return valor in (None, False, "", 0) or str(valor).strip().lower() in (
        "false", "none", "ninguno", "sin asignar")


def _tracking_pickings(picking_ids):
    """Fechas y responsables reales tomados del Chatter (mail.message).

    Devuelve {picking_id: {"asignado": utc, "hecho": utc,
                           "hecho_por": nombre, "hecho_por_id": uid}}.

    - asignado: primer cambio de ``user_id`` de vacio a un despachador, o
      primer cambio de ``state`` a 'assigned' (Disponible).
    - hecho: cambio de ``state`` a 'done' (Hecho) o el mensaje de validacion
      del traslado (subtipo stock.mt_picking_done). Solo se usa como respaldo
      cuando ``stock.picking.date_done`` no esta disponible.
    - hecho_por / hecho_por_id: usuario que realmente valido la operacion
      (autor del mensaje de validacion). Si el traslado se asigno a una
      persona pero lo ejecuto otra, manda quien lo ejecuto; sin rastro en el
      Chatter no se informan y queda el ``user_id`` asignado.

    La lectura evita buscar ``mail.message`` con dominios masivos: en Odoo 18
    ese modelo resuelve el acceso DESPUES de aplicar la ventana limit/offset
    y con LEFT JOINs a las notificaciones del usuario conectado, de modo que
    en lotes grandes recorta y duplica filas silenciosamente (un dominio con
    search_count 1599 devolvia paginas de ~400 y mensajes de validacion
    perdidos). En su lugar se buscan los ``mail.tracking.value`` atravesando
    ``mail_message_id`` (subconsulta SQL plana que no pasa por ese filtro) y
    los mensajes se leen por id con ``read()``, que tampoco pasa por
    ``_search``.
    """
    resultado = {}
    ids = sorted({i for i in picking_ids if i})
    if not ids:
        return resultado

    # En esta version de Odoo el campo auditado es ``field_id`` (many2one a
    # ir.model.fields); se resuelven los ids de 'user_id' y 'state' para no
    # traer los demas cambios del documento.
    try:
        campos = {f["id"]: f["name"]
                  for f in sr("ir.model.fields",
                              [("model", "=", "stock.picking"),
                               ("name", "in", ["user_id", "state"])],
                              ["id", "name"])}
    except Exception as exc:
        print(f"[despacho] No se pudo leer ir.model.fields: {exc}")
        return resultado
    if not campos:
        return resultado

    # Subtipo oficial de validacion del traslado: el mensaje que lo lleva lo
    # genera la accion de validar y su autor es el usuario que la ejecuto.
    subtipo_done = set()
    try:
        for d in sr("ir.model.data",
                    [("module", "=", "stock"), ("name", "=", "mt_picking_done"),
                     ("model", "=", "mail.message.subtype")],
                    ["res_id"]):
            subtipo_done.add(d["res_id"])
    except Exception as exc:
        print(f"[despacho] No se pudo leer el subtipo de validacion: {exc}")

    tracking = []
    try:
        for i in range(0, len(ids), 300):
            tracking.extend(sr("mail.tracking.value",
                               [("mail_message_id.model", "=", "stock.picking"),
                                ("mail_message_id.res_id", "in", ids[i:i + 300]),
                                ("field_id", "in", list(campos))],
                               ["mail_message_id", "field_id", "old_value_char",
                                "new_value_char", "old_value_integer",
                                "new_value_integer"],
                               order="id asc"))
    except Exception as exc:
        print(f"[despacho] No se pudo leer mail.tracking.value: {exc}")
        return resultado

    # Mensajes que llevan esos cambios, leidos por id: ver nota en el
    # docstring sobre por que no se buscan con un dominio masivo.
    msg_ids = sorted({id_of(tv.get("mail_message_id")) for tv in tracking} - {None})
    mensajes = []
    for i in range(0, len(msg_ids), 100):
        lote = msg_ids[i:i + 100]
        try:
            mensajes.extend(_execute("mail.message", "read", [lote],
                                     {"fields": ["id", "res_id", "date",
                                                 "author_id", "subtype_id"]}))
        except xmlrpc.client.Fault:
            for mid in lote:
                try:
                    mensajes.extend(_execute("mail.message", "read", [[mid]],
                                             {"fields": ["id", "res_id", "date",
                                                         "author_id", "subtype_id"]}))
                except xmlrpc.client.Fault as exc:
                    print(f"[despacho] Mensaje {mid} ilegible: {exc.faultString[:80]}")

    fecha_msg = {m["id"]: m.get("date") for m in mensajes}
    pick_de_msg = {m["id"]: id_of(m.get("res_id")) for m in mensajes}
    autor_de_msg = {m["id"]: m.get("author_id") for m in mensajes}

    def anotar_validacion(pid, fecha, autor):
        datos = resultado.setdefault(pid, {})
        actual = datos.get("hecho")
        if actual is None or fecha < actual:
            datos["hecho"] = fecha
            datos["_autor_validacion"] = autor

    for m in mensajes:
        pid = pick_de_msg.get(m["id"])
        fecha = fecha_msg.get(m["id"])
        if not pid or not fecha:
            continue
        if id_of(m.get("subtype_id")) in subtipo_done:
            anotar_validacion(pid, fecha, autor_de_msg.get(m["id"]))

    for tv in tracking:
        msg_id = id_of(tv.get("mail_message_id"))
        pid = pick_de_msg.get(msg_id)
        fecha = fecha_msg.get(msg_id)
        if not pid or not fecha:
            continue
        campo = campos.get(id_of(tv.get("field_id")))
        datos = resultado.setdefault(pid, {})
        if campo == "user_id":
            nuevo = tv.get("new_value_integer") or tv.get("new_value_char")
            viejo = tv.get("old_value_integer") or tv.get("old_value_char")
            if nuevo and _valor_vacio(viejo):
                if not datos.get("asignado") or fecha < datos["asignado"]:
                    datos["asignado"] = fecha
        elif campo == "state":
            nuevo = (tv.get("new_value_char") or "").strip().lower()
            if any(x in nuevo for x in _ESTADOS_ASIGNADO):
                if not datos.get("asignado") or fecha < datos["asignado"]:
                    datos["asignado"] = fecha
            elif any(x in nuevo for x in _ESTADOS_HECHO):
                anotar_validacion(pid, fecha, autor_de_msg.get(msg_id))

    # El autor del mensaje es un res.partner; se resuelve al usuario interno
    # (res.users) para mostrar su nombre y poder filtrar el tablero por el
    # despachador que realmente ejecuto la operacion.
    autores = sorted({id_of(d.get("_autor_validacion"))
                      for d in resultado.values() if d.get("_autor_validacion")})
    autor_a_usuario = {}
    if autores:
        try:
            # Se incluyen usuarios archivados: la identidad de un validador
            # historico no debe perderse porque su cuenta de Odoo ya no este
            # activa (se unifica despues con _canon_usuario si procede).
            for u in sr("res.users", [("partner_id", "in", autores),
                                      ("active", "in", [True, False])],
                        ["id", "name", "partner_id"]):
                autor_a_usuario[id_of(u.get("partner_id"))] = u
        except Exception as exc:
            print(f"[despacho] No se pudo leer res.users de validadores: {exc}")
    for datos in resultado.values():
        autor = id_of(datos.pop("_autor_validacion", None))
        usuario = autor_a_usuario.get(autor)
        # La fecha de validacion ("hecho") se conserva siempre; la autoria
        # solo se informa si el validador no es un usuario excluido del
        # tablero (p.ej. facturacion), porque entonces no es ejecucion de
        # bodega y el responsable queda resuelto por asignacion.
        if usuario and not _resp_excluido(usuario.get("id"), usuario.get("name")):
            datos["hecho_por_id"], datos["hecho_por"] = _canon_usuario(
                usuario.get("id"), usuario.get("name"))
    return resultado


def _fase_por_documento(nombre):
    """Tipo de operacion segun el nombre del traslado (PICK / PACK / OUT).

    Se exige el token completo entre barras para no confundir, por ejemplo,
    un picking con un nombre que solo contenga la palabra como subcadena.
    """
    partes = [x.strip().upper() for x in (nombre or "").split("/")]
    if "PICK" in partes:
        return "picking"
    if "PACK" in partes:
        return "packing"
    if "OUT" in partes:
        return "output"
    return None


def build_despacho(empresa, filtro_id, d0, d1, rol="todo"):
    """Tiempos de despacho por fase (picking/packing/output) en un rango de fechas.

    Puerto del reporte original /TiempoDespachoOutput/ de la app gestion-comercial.
    d0 y d1 son datetime (desde y hasta, inclusive). ``rol`` limita las
    participaciones visibles a pick, pack u out (todo = sin limite).
    """
    rol = (rol or "todo").lower()
    if rol not in ("todo", "pick", "pack", "out"):
        rol = "todo"
    # Un filtro pedido con una identidad duplicada (p. ej. un favorito del
    # navegador con el id de la cuenta generica) se resuelve al canonico,
    # igual que los responsables del reporte.
    filtro_id = _ALIAS_RESP_IDS.get(filtro_id or 0, filtro_id or 0)
    # Ventana en hora de Ecuador (UTC-5): desde las 00:00 del dia "desde"
    # hasta las 23:59:59 del dia "hasta", ambos inclusive en fecha local
    # (05:00 UTC = 00:00 local; el dia "hasta" cierra 04:59:59 UTC del dia
    # siguiente). Antes el dia "hasta" quedaba excluido en rangos de varios
    # dias, porque el cierre se ponia a las 04:59:59 UTC del mismo dia.
    inicio = f"{d0:%Y-%m-%d} 05:00:00"
    fin = f"{d1 + timedelta(days=1):%Y-%m-%d} 04:59:59"

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

    # Se traen TODAS las fases del pedido (completadas o abiertas) para poder
    # aplicar la cadena PICK -> PACK -> OUT sin inventar estados. La fecha de
    # origen del pedido solo se usa para el Lead Time, nunca como inicio de fase.
    bulk = sr("stock.picking",
              [("sale_id", "in", sale_ids), ("state", "!=", "cancel")],
              ["id", "name", "picking_type_id", "state", "date_done", "user_id",
               "create_date", "date_assigned_user", "sale_id", "partner_id",
               "write_uid", "write_date", "x_studio_fecha_y_hora_de_orden"])

    # Fechas reales de asignacion y cierre extraidas del Chatter (mail.message),
    # en lugar de las fechas genericas de creacion del traslado.
    tracking = _tracking_pickings([p["id"] for p in bulk])

    # Fecha de confirmacion de cada orden de venta (date_order: en Odoo es la
    # fecha de confirmacion de las ordenes ya validadas). Es la columna
    # "Fecha y hora de confirmacion" del tablero y la referencia de inicio de
    # los pedidos que no se asignaron a ningun despachador y fueron
    # ejecutados directamente por los usuarios de despacho_confirmacion_directa.
    confirmacion_pedido = {}
    try:
        for s in sr("sale.order", [("id", "in", sale_ids)], ["id", "date_order"]):
            confirmacion_pedido[s["id"]] = s.get("date_order")
    except Exception as exc:
        print(f"[despacho] No se pudo leer la fecha de confirmacion de ordenes: {exc}")

    # Conteo de items por picking. Se consulta stock.move en lugar de traer el
    # one2many move_ids_without_package: con miles de pedidos esa respuesta
    # XML-RPC llega truncada y el reporte falla.
    conteo_items = {}
    fin_por_picking = {}
    picking_ids = [p["id"] for p in bulk]
    for i in range(0, len(picking_ids), 500):
        for m in sr("stock.move", [("picking_id", "in", picking_ids[i:i + 500])],
                    ["picking_id", "state", "date"]):
            pid = id_of(m.get("picking_id"))
            conteo_items[pid] = conteo_items.get(pid, 0) + 1
            # Hora efectiva en que cada movimiento pasa a "Hecho".
            if m.get("state") == "done" and m.get("date"):
                if not fin_por_picking.get(pid) or m["date"] > fin_por_picking[pid]:
                    fin_por_picking[pid] = m["date"]

    por_pedido = {}
    for p in bulk:
        sid = id_of(p.get("sale_id"))
        if sid:
            por_pedido.setdefault(sid, []).append(p)

    fmt = "%Y-%m-%d %H:%M:%S"
    ahora_dt = datetime.utcnow()
    ahora = ahora_dt.strftime(fmt)
    resultados = []
    mapa_fechas_fin = {}
    fases_orden = (("pick", "picking"), ("pack", "packing"), ("out", "output"))

    def legible(m):
        h, mm = divmod(m, 60)
        return f"{h}h {mm}m"

    def rank(datos):
        return (1 if datos.get("completado") else 0, datos.get("fin_utc") or "",
                datos.get("inicio_utc") or "")

    def aplicar_tiempos(datos, prev_fin_dt):
        """Espera + ejecucion de una fase, medida solo en horario laboral.

        La espera corre desde el fin de la fase anterior hasta que esta fase
        arranca (asignacion/creacion). Si el documento se creo antes de que la
        fase anterior terminara, el inicio operativo se ajusta al fin anterior.
        """
        ini_dt = (datetime.strptime(datos["inicio_utc"], fmt)
                  if datos.get("inicio_utc") else None)
        ini_oper = ini_dt
        if prev_fin_dt and ini_oper and ini_oper < prev_fin_dt:
            ini_oper = prev_fin_dt
        espera = _minutos_habiles(prev_fin_dt, ini_oper) if prev_fin_dt and ini_oper else 0
        if datos.get("completado") and datos.get("fin_utc"):
            fin_dt = datetime.strptime(datos["fin_utc"], fmt)
        else:
            fin_dt = ahora_dt
        ejecucion = _minutos_habiles(ini_oper, fin_dt) if ini_oper and fin_dt else 0
        # Una fase ya validada ('done') nunca se muestra como 0h 0m: si el
        # traslado duro menos de un minuto se informa 1 minuto.
        if datos.get("completado") and ejecucion == 0:
            ejecucion = 1
        datos["inicio_operativo_utc"] = ini_oper.strftime(fmt) if ini_oper else None
        datos["espera_minutos"] = espera
        datos["espera_legible"] = legible(espera)
        datos["ejecucion_minutos"] = ejecucion
        datos["ejecucion_legible"] = legible(ejecucion)
        datos["minutos_netos"] = espera + ejecucion
        datos["legible"] = legible(datos["minutos_netos"])
        return fin_dt

    def fase_info(datos):
        return {"responsable": datos.get("responsable"),
                "documento": datos.get("documento", ""),
                "asignado": _fecha_hora_local(datos.get("asignado_utc")),
                "inicio": _fecha_hora_local(datos.get("inicio_utc")),
                "fin": _fecha_hora_local(datos.get("fin_utc")),
                "estado": "completado" if datos.get("completado") else "en_proceso",
                "completado": bool(datos.get("completado")),
                "espera_minutos": datos.get("espera_minutos", 0),
                "espera_legible": datos.get("espera_legible", "0h 0m"),
                "ejecucion_minutos": datos.get("ejecucion_minutos", 0),
                "ejecucion_legible": datos.get("ejecucion_legible", "0h 0m"),
                "minutos_netos": datos.get("minutos_netos", 0),
                "legible": datos.get("legible", "0h 0m")}

    for s_id in sale_ids:
        ops = por_pedido.get(s_id, [])
        if not ops:
            continue
        todas = {"picking": None, "packing": None, "output": None}
        visibles = {"picking": None, "packing": None, "output": None}
        nombre_pedido = f"ID: {s_id}"
        nombre_cliente = "Cliente no especificado"
        n_items = 0
        n_conting = 0

        for p in ops:
            if isinstance(p.get("sale_id"), (list, tuple)) and len(p["sale_id"]) > 1:
                nombre_pedido = p["sale_id"][1]
            if isinstance(p.get("partner_id"), (list, tuple)) and len(p["partner_id"]) > 1:
                nombre_cliente = p["partner_id"][1]
            # Cada fase se identifica por su propio documento, de forma
            # independiente: /PICK/, /PACK/ y /OUT/ en el nombre del traslado.
            fase = _fase_por_documento(p.get("name"))
            if not fase:
                operacion = (ref(p.get("picking_type_id")) or "").lower()
                if "pick" in operacion:
                    fase = "picking"
                elif any(x in operacion for x in ("pack", "emp", "prep")):
                    fase = "packing"
                elif any(x in operacion for x in ("out", "deliv", "salida", "entrega")):
                    fase = "output"
            if not fase:
                continue
            resp_id = id_of(p.get("user_id")) or 0
            completado = p.get("state") == "done"
            tr = tracking.get(p.get("id")) or {}
            # Responsable real de la fase: quien valido la operacion en Odoo
            # (autor del mensaje de validacion en el Chatter). Si el traslado
            # se asigno a una persona y lo ejecuto otra, manda quien lo
            # ejecuto. Sin rastro de validacion queda el usuario asignado y,
            # como ultimo recurso en traslados validados sin asignar, el
            # write_uid de una escritura contemporanea a la validacion
            # (ventana de 2 minutos): al validar, Odoo deja write_uid en
            # quien valido, pero despues el personal de facturacion escribe
            # el numero de autorizacion SRI en el mismo traslado y ese cambio
            # posterior NO puede figurar como ejecucion de bodega.
            validador_write = None
            validador_write_id = 0
            if completado and not tr.get("hecho_por"):
                wd, dd = p.get("write_date"), p.get("date_done")
                if wd and dd:
                    try:
                        delta = abs((datetime.fromisoformat(wd) -
                                     datetime.fromisoformat(dd)).total_seconds())
                        wuid = id_of(p.get("write_uid")) or 0
                        if delta <= 120 and not _resp_excluido(wuid, ref(p.get("write_uid"))):
                            validador_write_id, validador_write = _canon_usuario(
                                wuid, ref(p.get("write_uid")))
                    except ValueError:
                        validador_write = None
            # El usuario asignado tambien pasa por el mismo filtro: un
            # excluido (facturacion) nunca figura como responsable y la fase
            # queda sin responsable si no hay otro candidato.
            resp_asignado = ref(p.get("user_id"))
            if _resp_excluido(resp_id, resp_asignado):
                resp_id, resp_asignado = 0, None
            # Identidades duplicadas de una misma persona se reportan como el
            # usuario canonico (nombre incluido).
            resp_id, resp_asignado = _canon_usuario(resp_id, resp_asignado)
            resp_nombre = (tr.get("hecho_por") or validador_write or
                           resp_asignado or None)
            # El filtro por despachador sigue al ejecutor real.
            resp_filtro = (tr.get("hecho_por_id") or validador_write_id or
                           resp_id or 0)
            # Inicio operativo: asignacion registrada en el Chatter (cambio de
            # user_id de vacio al despachador o paso a 'assigned'). Si no hay
            # rastro de auditoria se usan las fechas del traslado como respaldo.
            inicio_utc = (tr.get("asignado") or p.get("date_assigned_user") or
                          p.get("create_date"))
            # Cierre: date_done del stock.picking cuando su estado es 'done'
            # (respaldo: cambio a 'Hecho' en el Chatter y ultimo stock.move).
            fin_utc = None
            if completado:
                fin_utc = (p.get("date_done") or tr.get("hecho") or
                           fin_por_picking.get(p.get("id")))
            datos_fase = {
                "documento": p.get("name", ""),
                "responsable": resp_nombre,
                "inicio_utc": inicio_utc,
                "fin_utc": fin_utc,
                "asignado_utc": (tr.get("asignado") or p.get("date_assigned_user") or
                                 (p.get("create_date") if resp_id else None)),
                "completado": completado,
            }
            n_moves = conteo_items.get(p.get("id"), 0)
            if n_moves:
                if fase == "picking":
                    n_items = n_moves
                elif n_conting == 0:
                    n_conting = n_moves
            if not todas[fase] or rank(datos_fase) > rank(todas[fase]):
                todas[fase] = datos_fase
            if filtro_id == 0 or resp_filtro == filtro_id:
                if not visibles[fase] or rank(datos_fase) > rank(visibles[fase]):
                    visibles[fase] = datos_fase

        if not any(visibles.values()):
            continue
        if filtro_id and rol in ("pick", "pack", "out"):
            fase_rol = {"pick": "picking", "pack": "packing", "out": "output"}[rol]
            if not visibles.get(fase_rol):
                continue
        if n_items == 0:
            n_items = n_conting

        # Pedidos que no se asignaron a ningun despachador y fueron ejecutados
        # directamente (p.ej. Jesus Sanchez): sin asignacion que consultar, la
        # referencia de inicio pasa a ser la fecha de confirmacion de la orden
        # de venta. Solo se rellena la referencia (asignado/inicio del ciclo);
        # el inicio operativo de cada fase y sus tiempos no se tocan.
        if confirmacion_pedido:
            sin_asignar = not any((tracking.get(p.get("id")) or {}).get("asignado") or
                                  id_of(p.get("user_id")) or p.get("date_assigned_user")
                                  for p in ops)
            if sin_asignar:
                directo = any(
                    _ejecutor_directo_conf((tracking.get(p.get("id")) or {}).get("hecho_por_id"),
                                           (tracking.get(p.get("id")) or {}).get("hecho_por"))
                    for p in ops)
                if directo:
                    fecha_conf = confirmacion_pedido.get(s_id)
                    if fecha_conf:
                        for fase in ("picking", "packing", "output"):
                            datos = todas.get(fase)
                            if datos and not datos.get("asignado_utc"):
                                datos["asignado_utc"] = fecha_conf

        # Inicio del ciclo: cuando el pedido se asigna a un despachador
        # (primer cambio de user_id registrado en el Chatter), no la fecha de
        # confirmacion de la orden. Solo se usa para el Lead Time y la espera
        # inicial del PICK.
        inicio_ciclo = None
        for fase in ("picking", "packing", "output"):
            datos = todas.get(fase)
            valor = datos.get("asignado_utc") if datos else None
            if valor and (inicio_ciclo is None or valor < inicio_ciclo):
                inicio_ciclo = valor
        if not inicio_ciclo:
            inicio_ciclo = next((p.get("x_studio_fecha_y_hora_de_orden") for p in ops
                                 if p.get("x_studio_fecha_y_hora_de_orden")), None)
        if not inicio_ciclo:
            inicio_ciclo = min((p.get("create_date") for p in ops if p.get("create_date")),
                               default=None)
        prev_fin_dt = datetime.strptime(inicio_ciclo, fmt) if inicio_ciclo else None
        ultimo_fin_dt = prev_fin_dt

        # Tiempo de cada fase = espera desde el fin anterior + ejecucion.
        for clave, fase in fases_orden:
            datos = todas.get(fase)
            if not datos:
                continue
            prev_para_visible = prev_fin_dt
            fin_dt = aplicar_tiempos(datos, prev_fin_dt)
            visible = visibles.get(fase)
            if visible is not None and visible is not datos:
                aplicar_tiempos(visible, prev_para_visible)
            if fin_dt:
                ultimo_fin_dt = fin_dt
                prev_fin_dt = fin_dt

        rep = {"pedido_id": s_id, "pedido_nombre": nombre_pedido,
               "cliente": nombre_cliente, "total_items": int(n_items),
               "fases": {}, "tiempos_fases_minutos": {},
               "tiempos_fases_legible": {}, "detalles_operacion": {},
               "estados_fases": {},
               "fecha_confirmacion": _fecha_hora_local(confirmacion_pedido.get(s_id))}

        # Documentos de las tres operaciones, aunque el filtro de responsable
        # oculte alguna fase: PICK / PACK / OUT.
        rep["documentos"] = {clave: (todas.get(fase) or {}).get("documento", "")
                             for clave, fase in fases_orden}

        # Estado completo de la cadena, aunque el filtro oculte al responsable.
        for clave, fase in fases_orden:
            datos = todas.get(fase)
            rep["estados_fases"][clave] = {
                "existe": bool(datos),
                "completado": bool(datos and datos.get("completado")),
            }

        for clave, fase in fases_orden:
            datos = visibles.get(fase)
            if not datos:
                continue
            info = fase_info(datos)
            rep["fases"][clave] = info
            rep["tiempos_fases_minutos"][fase] = info["minutos_netos"]
            rep["tiempos_fases_legible"][fase] = info["legible"]
            rep["detalles_operacion"][fase] = {
                "documento": datos.get("documento", ""),
                "responsable": datos.get("responsable"),
                "asignado": _fecha_hora_local(datos.get("asignado_utc")),
                # Fecha de referencia en hora local (antes se recortaba el
                # texto en UTC): es el dia en que la fase se completo (o
                # arranco si sigue abierta), sin desfase de zona horaria.
                "fecha": _fecha_local_corta(datos.get("fin_utc") or datos.get("inicio_utc")),
                "completado": bool(datos.get("completado")),
            }
            rep[f"responsable_{clave}"] = info["responsable"]
            rep[f"inicio_{clave}"] = info["inicio"]
            rep[f"fin_{clave}"] = info["fin"]
            rep[f"minutos_{clave}"] = info["minutos_netos"]
            rep[f"tiempo_{clave}"] = info["legible"]

        # Lead Time: desde la creacion del pedido hasta la ultima fase
        # ejecutada, o hasta ahora si la fase actual sigue abierta.
        lead = None
        if ultimo_fin_dt:
            desde_dt = datetime.strptime(inicio_ciclo, fmt) if inicio_ciclo else None
            if desde_dt is None:
                primera = todas["picking"] or todas["packing"] or todas["output"]
                desde_dt = (datetime.strptime(primera["inicio_operativo_utc"], fmt)
                            if primera and primera.get("inicio_operativo_utc") else None)
            if desde_dt:
                lead = _minutos_habiles(desde_dt, ultimo_fin_dt)
        rep["lead_time_minutos"] = lead
        rep["lead_time_legible"] = legible(lead) if lead is not None else None

        # Ciclo total del pedido: asignacion al despachador -> date_done de la
        # ultima fase (OUT cuando existe). Si esa fase ya esta 'done' el reloj
        # se congela con su date_done y no usa la fecha actual.
        fase_ciclo = ("output" if todas.get("output") else
                      ("packing" if todas.get("packing") else "picking"))
        datos_ciclo = todas.get(fase_ciclo) or {}
        fin_ciclo = (datos_ciclo.get("fin_utc")
                     if datos_ciclo.get("completado") else None)
        inicio_ciclo_dt = datetime.strptime(inicio_ciclo, fmt) if inicio_ciclo else None
        ciclo = None
        if fin_ciclo and inicio_ciclo_dt:
            ciclo = _minutos_habiles(inicio_ciclo_dt,
                                     datetime.strptime(fin_ciclo, fmt))
        rep["terminado"] = bool(datos_ciclo.get("completado"))
        rep["ciclo_total_minutos"] = ciclo
        rep["ciclo_total_legible"] = legible(ciclo) if ciclo is not None else None

        if rep["tiempos_fases_minutos"]:
            fase_final = ("output" if "output" in rep["detalles_operacion"] else
                          ("packing" if "packing" in rep["detalles_operacion"] else "picking"))
            rep["fecha_asignacion"] = rep["detalles_operacion"][fase_final]["asignado"]
            resultados.append(rep)
            fin_jornada = visibles[fase_final].get("fin_utc") or ahora
            mapa_fechas_fin[s_id] = datetime.strptime(fin_jornada, fmt)

    if not resultados:
        return vacio

    # tiempos muertos / transicion entre pedidos de cada despachador
    jornadas = {}
    for r in resultados:
        det = r["detalles_operacion"]
        fase_final = "output" if "output" in det else ("packing" if "packing" in det else "picking")
        nombre = det[fase_final]["responsable"]
        if not nombre:
            r["tiempo_espera_previo_minutos"] = 0
            r["tiempo_espera_previo_legible"] = "Sin asignar"
            continue
        jornadas.setdefault(nombre, []).append(
            {"pedido": r, "fecha_fin": mapa_fechas_fin.get(r["pedido_id"], datetime.min)})

    for lista in jornadas.values():
        lista.sort(key=lambda x: x["fecha_fin"])
        for i, item in enumerate(lista):
            pedido = item["pedido"]
            if pedido.get("terminado"):
                # El ciclo ya cerro (OUT 'done'): se congela el cronometro y
                # se informa el estado final en lugar de una espera abierta.
                pedido["tiempo_espera_previo_minutos"] = pedido.get("ciclo_total_minutos") or 0
                pedido["tiempo_espera_previo_legible"] = "Terminado"
            elif i == 0:
                pedido["tiempo_espera_previo_minutos"] = 0
                pedido["tiempo_espera_previo_legible"] = "Inicio de jornada"
            else:
                brecha = _minutos_habiles(lista[i - 1]["fecha_fin"], item["fecha_fin"])
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


def build_despacho_xlsx(empresa, filtro_id, d0, d1, responsable="", rol="todo"):
    """Exporta el dashboard de tiempos de despacho (misma salida que la app externa)."""
    rep = build_despacho(empresa, filtro_id, d0, d1, rol)
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

    def docs_de(p):
        """Documentos de las tres operaciones (PICK / PACK / OUT)."""
        docs = p.get("documentos") or {}
        partes = [docs.get(k) for k in ("pick", "pack", "out") if docs.get(k)]
        return " / ".join(partes) or first(p, "documento")

    if responsable:
        pedidos = [p for p in pedidos if resp_de(p) == responsable]

    wb = Workbook()
    ws = wb.active
    ws.title = "Despacho"
    ws["A1"] = f"TIEMPOS DE DESPACHO \u2014 {d0:%d/%m/%Y} \u2013 {d1:%d/%m/%Y}"
    ws["A1"].font = TITLE_FONT
    headers = ["Pedido", "Cliente", "Items",
               "Resp. PICK", "Inicio PICK", "Fin PICK", "PICK total (min)",
               "Resp. PACK", "Inicio PACK", "Fin PACK", "PACK total (min)",
               "Resp. OUT", "Inicio OUT", "Fin OUT", "OUT total (min)",
               "Documento", "Fecha y hora de confirmación", "Espera entre pedidos (min)", "Espera entre pedidos",
               "Trans. PICK\u2192PACK", "Trans. PACK\u2192OUT", "Lead Time"]
    head_row = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=c, value=h)
        cell.fill = HEAD_FILL
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for i, p in enumerate(pedidos, start=1):
        fase_pick = p.get("fases", {}).get("pick") or {}
        fase_pack = p.get("fases", {}).get("pack") or {}
        fase_out = p.get("fases", {}).get("out") or {}
        trans = p.get("transiciones") or {}
        vals = [p.get("pedido_nombre"), p.get("cliente"), p.get("total_items", 0),
                fase_pick.get("responsable", "\u2014"),
                fase_pick.get("inicio", "\u2014"),
                fase_pick.get("fin", "\u2014"),
                fase_pick.get("minutos_netos", 0),
                fase_pack.get("responsable", "\u2014"),
                fase_pack.get("inicio", "\u2014"),
                fase_pack.get("fin", "\u2014"),
                fase_pack.get("minutos_netos", 0),
                fase_out.get("responsable", "\u2014"),
                fase_out.get("inicio", "\u2014"),
                fase_out.get("fin", "\u2014"),
                fase_out.get("minutos_netos", 0),
                docs_de(p), p.get("fecha_confirmacion") or p.get("fecha_asignacion") or "\u2014",
                p.get("tiempo_espera_previo_minutos", 0),
                p.get("tiempo_espera_previo_legible", "\u2014"),
                (trans.get("pick_pack") or {}).get("legible", "\u2014"),
                (trans.get("pack_out") or {}).get("legible", "\u2014"),
                p.get("lead_time_legible") or "\u2014"]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=head_row + i, column=c, value=v)
            cell.border = BORDER
            if c in (7, 11, 15, 18):
                cell.number_format = "#,##0"
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
    widths = [16, 30, 8,
              20, 17, 17, 13,
              20, 17, 17, 13,
              20, 17, 17, 13,
              18, 17, 17, 14, 16, 16, 14]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    ws.auto_filter.ref = f"A{head_row}:V{head_row + len(pedidos)}"
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
                                    else "usuario_estÃ¡ndar",
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
                # Solo quienes figuran como responsables del tablero
                # (asignados o validadores de las fases PICK/PACK/OUT, de
                # cualquiera de las companias), no todos los usuarios de
                # Odoo: ventas, administracion y cuentas sin actividad de
                # bodega no aparecen en el filtro.
                return self._json(200, _despachadores_tablero())
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
                rol = (one("rol", "todo") or "todo").lower()
                return self._json(200, build_despacho(empresa, user_id, d0, d1, rol))
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
                rol = (one("rol", "todo") or "todo").lower()
                data = build_despacho_xlsx(empresa, user_id, d0, d1,
                                           one("responsable", ""), rol)
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
            usr, cortado = self._requiere()
            if cortado:
                return
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