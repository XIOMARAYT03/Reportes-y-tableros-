"""Kardex Fimas - servidor local de reportes de inventario (Odoo) con control de acceso."""
import http.client
import io
import json
import os
import threading
import time
import traceback
import urllib.parse
import xmlrpc.client
from collections import defaultdict
from datetime import datetime, timedelta
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import devoluciones
import odoo_auth
import permisos

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
    """Ejecuta una llamada RPC con exclusion mutua y un reintento ante fallo de transporte.

    xmlrpc.client.ServerProxy no es seguro entre hilos y las conexiones HTTP/1.1
    pueden quedar en estado Idle o Request-sent; por eso se crea un proxy nuevo
    por llamada y se serializan las llamadas con un RLock.
    """
    ultimo = None
    for intento in range(2):
        try:
            with _lock:
                _rpc.models = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object", allow_none=True)
                return _proxy().execute_kw(DB, UID, PWD, model, method, args, kwargs or {})
        except xmlrpc.client.Fault:
            raise
        except (http.client.HTTPException, xmlrpc.client.ProtocolError,
                xmlrpc.client.ResponseError, OSError) as exc:
            ultimo = exc
            _rpc.models = None
            time.sleep(0.3 * (intento + 1))
    raise ultimo


UID = _connect()
odoo_auth.configurar(URL, DB, USER, PWD)


def sr(model, domain, fields, order=None, limit=None):
    kw = {"fields": fields}
    if order:
        kw["order"] = order
    if limit:
        kw["limit"] = limit
    return _execute(model, "search_read", [domain], kw)


def id_of(v):
    if isinstance(v, (list, tuple)):
        return v[0] if v else None
    return v


def ref(v):
    return v[1] if isinstance(v, (list, tuple)) and len(v) > 1 else None


# El modulo de devoluciones reutiliza la conexion Odoo de este servidor por
# inyeccion (no importa kardex_server, evitando el import circular).
devoluciones.configurar(sr=sr, id_of=id_of, ref=ref, url=URL)


def load_global():
    return {
        "picking_types": {p["id"]: p for p in sr("stock.picking.type", [], ["id", "name", "code", "sequence_code"])},
        "companies": sr("res.company", [], ["id", "name"]),
        "warehouses": sr("stock.warehouse", [], ["id", "name", "code"]),
    }


G = load_global()
_loc_cache, _partner_cache = {}, {}


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


def build_kardex(company_id, product_id):
    key = f"{company_id}:{product_id}"
    with _lock:
        if key in _cache["kardex"]:
            return _cache["kardex"][key]

    prod = sr("product.product", [("id", "=", product_id)],
              ["id", "display_name", "name", "default_code", "uom_id", "categ_id",
               "product_tmpl_id", "standard_price", "type"])[0]
    tmpl = sr("product.template", [("id", "=", id_of(prod["product_tmpl_id"]))],
              ["cost_method", "valuation", "standard_price"])[0]

    payload = {
        "ok": True,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "empresa": {"id": company_id,
                    "nombre": next((c["name"] for c in G["companies"] if c["id"] == company_id), "N/A")},
        "producto": {"id": product_id, "codigo": prod["default_code"] or "N/A",
                     "nombre": prod["name"], "descripcion": prod["display_name"],
                     "unidad": ref(prod["uom_id"]) or "Unidades",
                     "categoria": ref(prod["categ_id"]) or "N/A",
                     "costo_promedio_erp": prod["standard_price"] or 0.0,
                     "metodo": {"average": "Promedio ponderado (average) - real_time",
                                "fifo": "FIFO - real_time",
                                "standard": "Costo est\u00e1ndar"}.get(tmpl["cost_method"], tmpl["cost_method"]),
                     "valuacion": tmpl["valuation"]},
        "kpi": {}, "rows": [], "internals": [],
    }

    moves = sr("stock.move", [("product_id", "=", product_id), ("company_id", "=", company_id)],
               ["id", "name", "date", "quantity", "product_qty", "picking_id", "picking_type_id",
                "location_id", "location_dest_id", "state", "purchase_line_id", "sale_line_id",
                "origin_returned_move_id", "origin"], order="date asc, id asc")
    done = [m for m in moves if m["state"] == "done"]
    if not done:
        payload.update({"ok": False,
                        "sin_movimientos": f"No se encontraron movimientos registrados para el \u00edtem "
                                           f"{prod['default_code']} - {prod['name']} en el per\u00edodo "
                                           f"indicado para {payload['empresa']['nombre']}.",
                        "cancelados": len(moves) - len(done), "total_moves": len(moves)})
        with _lock:
            _cache["kardex"][key] = payload
        return payload

    layers = sr("stock.valuation.layer", [("product_id", "=", product_id), ("company_id", "=", company_id)],
                ["stock_move_id", "quantity", "unit_cost", "value"], order="id asc")
    layer_by_move = defaultdict(list)
    for l in layers:
        layer_by_move[id_of(l["stock_move_id"])].append(l)

    def by_ids(model, ids, fields):
        ids = sorted({i for i in ids if i})
        if not ids:
            return {}
        return {r["id"]: r for r in sr(model, [("id", "in", ids)], fields)}

    orig = by_ids("stock.move", [id_of(m.get("origin_returned_move_id")) for m in done],
                  ["id", "purchase_line_id", "sale_line_id", "origin"])
    picks = by_ids("stock.picking", [id_of(m.get("picking_id")) for m in done],
                   ["id", "name", "partner_id", "origin", "state"])
    pol = by_ids("purchase.order.line",
                 [id_of(m.get("purchase_line_id")) for m in done] +
                 [id_of(v.get("purchase_line_id")) for v in orig.values()],
                 ["id", "order_id", "price_subtotal"])
    sol = by_ids("sale.order.line",
                 [id_of(m.get("sale_line_id")) for m in done] +
                 [id_of(v.get("sale_line_id")) for v in orig.values()], ["id", "order_id"])
    pos = by_ids("purchase.order", [id_of(v["order_id"]) for v in pol.values()],
                 ["id", "name", "partner_id", "date_order", "state"])
    sos = by_ids("sale.order", [id_of(v["order_id"]) for v in sol.values()],
                 ["id", "name", "partner_id", "date_order", "state"])
    locs = {l["id"]: l for l in locations([id_of(v) for m in done for v in (m.get("location_id"), m.get("location_dest_id"))])}
    names = partners([id_of(v.get("partner_id")) for v in list(picks.values()) + list(pos.values()) + list(sos.values())])
    pname = {p["id"]: p for p in names}

    rows, internals = [], []
    for m in done:
        pt = G["picking_types"].get(id_of(m.get("picking_type_id")))
        code = pt["sequence_code"] if pt else None
        tcode = pt["code"] if pt else None
        pk = picks.get(id_of(m.get("picking_id")), {})
        om = orig.get(id_of(m.get("origin_returned_move_id")), {})
        po_line = pol.get(id_of(m.get("purchase_line_id"))) or pol.get(id_of(om.get("purchase_line_id")))
        so_line = sol.get(id_of(m.get("sale_line_id"))) or sol.get(id_of(om.get("sale_line_id")))
        po = pos.get(id_of((po_line or {}).get("order_id")), {})
        so = sos.get(id_of((so_line or {}).get("order_id")), {})
        partner = ref(pk.get("partner_id"))
        po_partner = ref(po.get("partner_id"))
        lay = layer_by_move.get(m["id"], [])
        lq = sum(l["quantity"] for l in lay)
        lval = sum(l["value"] for l in lay)

        if code in ("PICK", "PACK", "PRCH", "INT"):
            internals.append({"date": m["date"], "tipo": pt["sequence_code"], "doc": pk.get("name"),
                              "cantidad": m["quantity"], "ubicacion": ref(locs.get(id_of(m.get("location_id")), {}).get("complete_name")),
                              "origen": m.get("origin") or "N/A"})
            continue
        src_u = (locs.get(id_of(m.get("location_id"))) or {}).get("usage")
        dst_u = (locs.get(id_of(m.get("location_dest_id"))) or {}).get("usage")
        if code is None and src_u in ("internal", "transit") and dst_u in ("internal", "transit"):
            internals.append({"date": m["date"], "tipo": "AJ-INT", "doc": m["name"], "cantidad": m["quantity"],
                              "ubicacion": ref((locs.get(id_of(m.get("location_dest_id"))) or {}).get("complete_name")),
                              "origen": m.get("origin") or "N/A"})
            continue

        if lq:
            qty, value = lq, lval
            cost = lval / lq if lq else 0.0
        else:
            qty = -m["quantity"] if tcode == "outgoing" else m["quantity"]
            cost = prod["standard_price"] or 0.0
            value = qty * cost

        if tcode == "incoming":
            tipo = TIPO_ENTRADA.get(code, "Entrada")
            contraparte = po_partner or partner
            etiqueta = "Cliente" if code == "DEVCL" else "Proveedor"
            origen = f"{etiqueta}: {contraparte or 'No especificado'}"
            doc = pk.get("name") or m.get("origin") or m["name"] or "N/A"
            if po.get("name"):
                doc += f" (OC {po['name']})"
            elif so.get("name"):
                doc += f" (Venta {so['name']})"
        elif tcode == "outgoing":
            tipo = "Devoluci\u00f3n a proveedor" if code == "DEVP" else "Venta (Despacho)"
            origen = f"Destino: {partner or 'No especificado'}"
            doc = pk.get("name") or m.get("origin") or m["name"] or "N/A"
            if so.get("name"):
                doc += f" (Venta {so['name']})"
            elif po.get("name"):
                doc += f" (OC {po['name']})"
        else:
            tipo = "Ajuste de inventario"
            if qty > 0:
                origen = "Ubic. destino: " + (ref(m.get("location_dest_id")) or "No especificado")
            else:
                origen = "Ubic. origen: " + (ref(m.get("location_id")) or "No especificado") + " (Ajuste a inventario f\u00edsico)"
            doc = m["name"] or "Ajuste de inventario"

        rows.append({"date": m["date"], "tipo": tipo, "doc": doc, "origen": origen,
                     "qty": round(qty, 4), "cost": round(cost, 6), "valor": round(value, 4)})

    rows.sort(key=lambda r: (r["date"], 0 if r["qty"] > 0 else 1))
    saldo = valor = 0.0
    for r in rows:
        saldo += r["qty"]
        valor += r["valor"]
        r["saldo"] = round(saldo, 4)
        r["valor_total"] = round(valor, 2)
        r["costo_prom"] = round(valor / saldo, 6) if saldo > 0 else 0.0
        r["entrada"] = r["qty"] if r["qty"] > 0 else 0
        r["salida"] = -r["qty"] if r["qty"] < 0 else 0

    ent = [r for r in rows if r["qty"] > 0]
    sal = [r for r in rows if r["qty"] < 0]
    payload["kpi"] = {
        "entradas": round(sum(r["entrada"] for r in ent), 2),
        "salidas": round(sum(r["salida"] for r in sal), 2),
        "saldo": rows[-1]["saldo"], "valor": rows[-1]["valor_total"],
        "costo_prom": rows[-1]["costo_prom"], "movimientos": len(rows),
        "internos": len(internals), "internos_qty": round(sum(i["cantidad"] for i in internals), 2),
        "cancelados": len(moves) - len(done),
    }
    payload["rows"] = rows
    payload["internals"] = internals
    payload["desde"], payload["hasta"] = rows[0]["date"][:10], rows[-1]["date"][:10]
    payload["notas"] = [
        f"Fuente: Odoo \u2014 stock.move / stock.move.line (estado done) y stock.valuation.layer. Producto ID {product_id}, compa\u00f1\u00eda {company_id}.",
        f"M\u00e9todo de valuaci\u00f3n del ERP: {payload['producto']['metodo']}; el costo unitario de cada movimiento corresponde a su capa de valorizaci\u00f3n.",
        f"Transferencias internas excluidas del detalle: {payload['kpi']['internos']} movimientos ({payload['kpi']['internos_qty']:,.2f} unidades), se listan en la hoja secundaria del Excel.",
        "Importes sin IVA. Valores queried en tiempo real desde el ERP.",
    ]
    with _lock:
        _cache["kardex"][key] = payload
    return payload


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
        ("Per\u00edodo del reporte", frm.get("desde", payload.get("desde")) + " \u2013 " + frm.get("hasta", payload.get("hasta"))),
        ("Filtros aplicados", frm.get("texto", "(sin filtros)")),
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
            ("Transferencias internas", k["internos"])]
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
              ["name", "picking_type_id", "date_done", "user_id", "create_date",
               "sale_id", "move_ids_without_package", "partner_id"])

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
            moves = p.get("move_ids_without_package")
            if moves and isinstance(moves, list):
                if "pick" in operacion:
                    n_items = len(moves)
                elif n_conting == 0:
                    n_conting = len(moves)
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
                          "/api/export", "/api/export/csv") and not permisos.puede(usr["id"], "kardex"):
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
            if u.path.startswith("/api/devoluciones") and not (
                    permisos.puede(usr["id"], "devoluciones")):
                nombre = permisos.modulos_activos().get("devoluciones", "Devoluciones")
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
                return self._json(200, build_kardex(cid, pid))
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
            if u.path == "/api/devoluciones":
                data = devoluciones.build(devoluciones.COMPANY_FINPAC,
                                          desde=one("desde", ""), hasta=one("hasta", ""),
                                          vendedor=one("vendedor", ""))
                return self._json(200, data)
            if u.path == "/api/devoluciones/detalle":
                try:
                    mid = int(one("move_id", "0") or 0)
                except ValueError:
                    return self._json(400, {"error": "move_id",
                                            "mensaje": "Indique el caso a consultar."})
                res = devoluciones.detalle(devoluciones.COMPANY_FINPAC, mid)
                return self._json(200 if res.get("ok") else 404, res)
            if u.path == "/api/devoluciones/export":
                data = devoluciones.build(devoluciones.COMPANY_FINPAC,
                                          desde=one("desde", ""), hasta=one("hasta", ""),
                                          vendedor=one("vendedor", ""))
                xlsx = devoluciones.build_xlsx(data)
                fname = f"devoluciones_pendientes_{datetime.now():%Y%m%d_%H%M}.xlsx"
                return self._send(200, xlsx,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{fname}"'})
            if u.path in ("/api/export", "/api/export/csv"):
                cid, pid = int(one("company_id")), int(one("product_id"))
                payload = build_kardex(cid, pid)
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
    srv.serve_forever()