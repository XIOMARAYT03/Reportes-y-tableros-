"""Pruebas del servicio de sincronizacion (ERP ficticio, sin conexion).

    python -m unittest discover -s app/tests -v
"""
import os
import sqlite3
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sync  # noqa: E402


def _id_of(valor):
    if valor is None or valor is False:
        return None
    if isinstance(valor, (list, tuple)):
        return valor[0] if valor else None
    if isinstance(valor, dict):
        return valor.get("id")
    return valor


def _ref(valor):
    if isinstance(valor, (list, tuple)) and len(valor) > 1:
        return valor[1]
    return valor


class FakeERP:
    """ERP minimo con movimientos, quants, unidades, ubicaciones y capas."""

    def __init__(self, movimientos):
        self.movimientos = list(movimientos)
        self.llamadas = []
        self.fallar_en_pagina = None      # 1-based: pagina que debe fallar
        self.paginas_servidas = 0

    def sr(self, modelo, dominio=None, campos=None, order=None, limit=None, offset=None):
        self.llamadas.append((modelo, list(dominio or []), limit))
        if modelo == "uom.uom":
            return [{"id": 1, "name": "Units", "category_id": [1, "Units"], "factor": 1.0,
                     "factor_inv": 1.0, "rounding": 0.01, "uom_type": "bigger", "active": True}]
        if modelo == "stock.location":
            return [{"id": 10, "name": "Stock", "complete_name": "C1/Stock", "usage": "internal",
                     "company_id": 1, "active": True},
                    {"id": 30, "name": "Suppliers", "complete_name": "Suppliers",
                     "usage": "supplier", "company_id": 1, "active": True}]
        if modelo == "stock.picking.type":
            return [{"id": 5, "sequence_code": "IN", "code": "incoming", "name": "Receipts"}]
        if modelo == "stock.move":
            return self._movimientos(dominio, limit)
        if modelo == "stock.quant":
            return [{"id": 1, "location_id": [10, "Stock"], "quantity": 100.0,
                     "reserved_quantity": 0.0, "inventory_quantity": 0.0, "product_id": 99}]
        if modelo == "stock.valuation.layer":
            return [{"id": 1, "stock_move_id": [m["id"], m["name"]], "quantity": m["quantity"],
                     "value": m["quantity"] * 5.0, "unit_cost": 5.0, "product_id": 99}
                    for m in self.movimientos]
        return []

    def _hoja(self, cond, m):
        campo, op, valor = cond
        if campo == "write_date":
            a, b = (m.get("write_date") or ""), (valor or "")
        elif campo == "id":
            a, b = m["id"], valor
        elif campo == "state":
            a, b = m.get("state"), valor
        elif campo == "product_id":
            a, b = m.get("product_id"), valor
        else:
            return True
        if op == ">":
            return a > b
        if op == "<":
            return a < b
        if op == "=":
            return a == b
        if op in ("!=",):
            return a != b
        if op == "in":
            return a in (valor or [])
        return True

    def _eval(self, tokens, i, m):
        """Evalua un dominio en notacion polaca (&, |, !) como lo hace Odoo."""
        if i >= len(tokens):
            return True, i
        tok = tokens[i]
        if tok == "|":
            izq, j = self._eval(tokens, i + 1, m)
            der, j = self._eval(tokens, j, m)
            return (izq or der), j
        if tok == "&":
            izq, j = self._eval(tokens, i + 1, m)
            der, j = self._eval(tokens, j, m)
            return (izq and der), j
        if tok == "!":
            r, j = self._eval(tokens, i + 1, m)
            return (not r), j
        return self._hoja(tok, m), i + 1

    def _match(self, tokens, m):
        i, res = 0, True
        while i < len(tokens):
            r, i = self._eval(tokens, i, m)
            res = res and r
        return res

    def _movimientos(self, dominio, limit):
        filas = [m for m in self.movimientos if self._match(list(dominio or []), m)]
        filas.sort(key=lambda m: (m["write_date"] or "", m["id"]))
        self.paginas_servidas += 1
        if self.fallar_en_pagina and self.paginas_servidas == self.fallar_en_pagina:
            raise OSError("fallo simulado de transporte")
        return filas[:limit] if limit else filas


def movimiento(id, write_date, cantidad=10.0):
    return {"id": id, "name": f"MV/{id}", "date": "2026-01-01 10:00:00",
            "quantity": cantidad, "product_qty": cantidad, "product_id": 99,
            "product_uom": [1, "Units"], "company_id": [1, "FIMAS"], "state": "done",
            "location_id": [30, "Suppliers"], "location_dest_id": [10, "Stock"],
            "picking_id": [100, "P/001"], "picking_type_id": [5, "Receipts"],
            "purchase_line_id": False, "sale_line_id": False,
            "origin_returned_move_id": False, "origin": "PO001", "write_date": write_date}


class TestSync(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="kardex_sync_")
        self.ruta = os.path.join(self.dir, "inventario.db")
        self.erp = FakeERP([movimiento(i, f"2026-01-{i:02d} 08:00:00") for i in range(1, 8)])
        self.svc = sync.ServicioSync(self.erp.sr, _id_of, _ref, "http://localhost:8069",
                                     db_path=self.ruta, intervalo=60)

    def tearDown(self):
        try:
            os.remove(self.ruta)
        except OSError:
            pass

    def test_guarda_los_movimientos(self):
        res = self.svc.sincronizar(motivo="test")
        self.assertTrue(res["ok"], res)
        self.assertEqual(len(self.svc.movimientos_producto(99)), 7)
        self.assertEqual(self.svc.estado()["estado"], "ok")

    def test_la_marca_de_agua_avanza(self):
        self.svc.sincronizar(motivo="test")
        est = self.svc._estado_modelo("stock.move")
        self.assertEqual(est["marca_agua_id"], 7)
        self.assertTrue(est["marca_agua_fecha"])

    def test_es_idempotente(self):
        self.svc.sincronizar(motivo="test")
        self.svc.sincronizar(motivo="test2")
        self.assertEqual(len(self.svc.movimientos_producto(99)), 7)

    def test_solo_trae_lo_nuevo_por_marca_de_agua(self):
        self.svc.sincronizar(motivo="test")
        self.erp.movimientos.append(movimiento(9, "2026-02-01 08:00:00"))
        self.svc.sincronizar(motivo="test2")
        ids = [m["id"] for m in self.svc.movimientos_producto(99)]
        self.assertIn(9, ids)
        self.assertEqual(len(ids), 8)

    def test_reanuda_tras_un_fallo_de_transporte(self):
        # La primera pagina se guarda; la segunda falla.  La marca de agua
        # debe quedar en la primera pagina para reintentar la que falta.
        self.erp.fallar_en_pagina = 2
        sync.PAGINA = 4
        try:
            res = self.svc.sincronizar(motivo="test")
        finally:
            sync.PAGINA = 500
        self.assertFalse(res["ok"])
        est = self.svc._estado_modelo("stock.move")
        self.assertEqual(est["marca_agua_id"], 4)
        self.assertEqual(len(self.svc.movimientos_producto(99)), 4)
        # El siguiente ciclo completa lo que falto
        res2 = self.svc.sincronizar(motivo="test2")
        self.assertTrue(res2["ok"], res2)
        self.assertEqual(len(self.svc.movimientos_producto(99)), 7)

    def test_reduce_la_pagina_ante_un_error_de_transporte(self):
        self.erp.fallar_en_pagina = 2
        sync.PAGINA = 400
        try:
            self.svc.sincronizar(motivo="test")
        finally:
            sync.PAGINA = 500
        # tras reintentar con pagina mas chica el ciclo debe terminar
        self.assertGreaterEqual(self.svc._estado_modelo("stock.move")["marca_agua_id"], 1)

    def test_guarda_los_quants_por_ubicacion(self):
        self.svc.sincronizar(motivo="test")
        quants = self.svc.quants_producto(99)
        self.assertTrue(quants)
        self.assertEqual(quants[0]["location_id"], 10)
        self.assertAlmostEqual(quants[0]["quantity"], 100.0)

    def test_lanzar_devuelve_de_inmediato(self):
        res = self.svc.lanzar(motivo="manual")
        self.assertTrue(res["ok"])
        # el ciclo corre en segundo plano
        for _ in range(100):
            if self.svc.estado()["estado"] == "ok" and self.svc.movimientos_producto(99):
                break
            import time
            time.sleep(0.05)
        self.assertEqual(len(self.svc.movimientos_producto(99)), 7)

    def test_estado_y_errores(self):
        self.svc.sincronizar(motivo="test")
        est = self.svc.estado()
        self.assertEqual(est["intervalo_seg"], 60)
        self.assertTrue(est["ultima_sincronizacion"])
        self.assertIsInstance(self.svc.errores(10), list)

    def test_el_esquema_crea_todas_las_tablas(self):
        self.svc.sincronizar(motivo="test")
        conn = sqlite3.connect(self.ruta)
        tablas = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        conn.close()
        for t in ("sync_estado", "sync_log", "movimientos", "capas", "productos",
                  "ubicaciones", "unidades", "quants", "conciliaciones"):
            self.assertIn(t, tablas)


if __name__ == "__main__":
    unittest.main(verbosity=2)
