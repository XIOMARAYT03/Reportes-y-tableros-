"""Prueba en vivo contra Odoo: el kardex de CIN058 debe cuadrar con el stock.

    python app/tests/test_live_cin058.py

El saldo de referencia (297) corresponde al estado del ERP antes de la venta
del 22/10/2026.  La comprobacion principal no es ese numero sino la igualdad
entre el saldo calculado por movimientos y el stock fisico real (`stock.quant`)
de las mismas ubicaciones internas de la empresa; el 297 solo se exige cuando
la consulta al ERP confirma que el almacen sigue en ese saldo.

Si el ERP no esta accesible la prueba se omite (no falla) para que el resto
de la bateria se pueda ejecutar sin conexion.
"""
import os
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

CODIGO = os.environ.get("KARDEX_TEST_CODIGO", "CIN058")
SALDO_ESPERADO = float(os.environ.get("KARDEX_TEST_SALDO", "297"))
EMPRESA = int(os.environ.get("KARDEX_TEST_EMPRESA", "1"))
FECHA_AJUSTE = "2026-03-09"


class TestLiveCIN058(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            import kardex_server as ks
        except Exception as exc:                      # pragma: no cover
            raise unittest.SkipTest(f"El servidor no pudo importarse: {exc}")
        try:
            prod = ks.sr("product.product", [("default_code", "=", CODIGO)],
                         ["id", "default_code", "name"])
        except Exception as exc:                      # pragma: no cover
            raise unittest.SkipTest(f"El ERP no esta accesible: {exc}")
        if not prod:
            raise unittest.SkipTest(f"El producto {CODIGO} no existe en el ERP")
        cls.ks = ks
        cls.pid = prod[0]["id"]
        cls.codigo = prod[0]["default_code"]

        # El resto de la bateria deja un ERP ficticio instalado en `inventario`:
        # se restaura el acceso real para que el kardex se arme con Odoo.
        import inventario
        inventario.configurar(ks.sr, ks.id_of, ks.ref)
        inventario.invalidar_catalogos()
        cls.alm = inventario.almacenes(EMPRESA)
        cls.locs = inventario.ubicaciones()

    # ------------------------------------------------------------------ #
    def stock_odoo(self):
        """Stock fisico real del producto en las ubicaciones internas del alcance."""
        quants = self.ks.sr("stock.quant", [("product_id", "=", self.pid)],
                            ["location_id", "quantity", "reserved_quantity"])
        total = 0.0
        for q in quants:
            if self.ks.id_of(q.get("location_id")) in self.alm:
                total += q.get("quantity") or 0.0
        return total

    def kardex(self):
        p = self.ks.build_kardex(EMPRESA, self.pid, refresh=True)
        self.assertTrue(p["ok"], p.get("sin_movimientos"))
        return p

    # ------------------------------------------------------------------ #
    def test_el_kardex_coincide_con_el_stock_de_odoo(self):
        """Regla principal: el saldo por movimientos = existencias en Odoo."""
        p = self.kardex()
        odoo = self.stock_odoo()
        self.assertAlmostEqual(p["kpi"]["saldo"], odoo, places=2,
                               msg=f"saldo kardex {p['kpi']['saldo']} vs "
                                   f"stock Odoo {odoo}")
        self.assertAlmostEqual(p["kpi"]["fisico"], odoo, places=2)
        self.assertAlmostEqual(p["conciliacion"]["stock_odoo"], odoo, places=2)
        self.assertAlmostEqual(p["conciliacion"]["diferencia"], 0.0, places=2)
        self.assertEqual(p["conciliacion"]["estado"], "conciliado",
                         f"diferencia de conciliacion: {p['conciliacion']['diferencia']}")

    def test_saldo_y_conciliacion(self):
        p = self.kardex()
        odoo = self.stock_odoo()
        k = p["kpi"]
        # El saldo de referencia solo se exige si el ERP sigue en ese saldo.
        if abs(odoo - SALDO_ESPERADO) < 0.005:
            self.assertAlmostEqual(k["saldo"], SALDO_ESPERADO, places=2,
                                   msg=f"saldo esperado {SALDO_ESPERADO}, "
                                       f"obtenido {k['saldo']}")
        else:
            self.assertAlmostEqual(k["saldo"], odoo, places=2,
                                   msg=f"saldo esperado {odoo} (stock de Odoo), "
                                       f"obtenido {k['saldo']}")
        # Las entradas y salidas del kardex cuadran con el saldo.
        self.assertAlmostEqual(k["entradas"] - k["salidas"], k["saldo"], places=2)

    def test_el_movimiento_virtual_no_altera_el_saldo(self):
        """El ajuste ENCERAR (internal->inventory -76, transit->inventory +2,
        inventory->internal +47) debe dejar el saldo en el stock real, no +2."""
        p = self.kardex()
        filas = [r for r in p["rows"] if r["date"][:10] == FECHA_AJUSTE]
        ajuste = sum(r["entrada"] - r["salida"] for r in filas)
        # El efecto neto sobre el almacen debe ser -29 (los movimientos que
        # pasan por ubicaciones virtuales se contabilizan por su efecto real
        # segun el alcance).
        self.assertAlmostEqual(ajuste, -29.0, places=2,
                               msg=f"efecto de los movimientos del {FECHA_AJUSTE}: "
                                   f"{ajuste} (se esperaba -29)")
        self.assertAlmostEqual(p["kpi"]["saldo"], self.stock_odoo(), places=2)

    def test_el_ajuste_virtual_no_aparece_como_entrada(self):
        """Ningun movimiento entre dos ubicaciones virtuales puede generar fila.

        Es el caso que descuadraba CIN058: "Transito Inter-Companias" ->
        "Ajuste de Inventario" se contaba como entrada y sumaba 2 unidades
        que nunca estuvieron en el almacen.
        """
        p = self.kardex()
        movimientos = self.ks._datos_movimientos(self.pid)[0]
        fuera = []
        for m in movimientos:
            if m.get("state") != "done":
                continue
            s = self.ks.id_of(m.get("location_id"))
            d = self.ks.id_of(m.get("location_dest_id"))
            if s in self.alm or d in self.alm:
                continue
            fuera.append((s, d))
        self.assertTrue(fuera, "se esperaba al menos un movimiento fuera del alcance")
        for s, d in fuera:
            nom_s = self.locs.get(s, {}).get("complete_name", "")
            nom_d = self.locs.get(d, {}).get("complete_name", "")
            if not (nom_s and nom_d):
                continue
            for fila in p["rows"]:
                self.assertFalse(nom_s in fila["origen"] and nom_d in fila["origen"],
                                 f"movimiento virtual contado como fila del kardex: "
                                 f"{nom_s} -> {nom_d}")

    def test_el_espejo_local_da_el_mismo_resultado(self):
        """Con el espejo sincronizado el resultado debe ser identico."""
        svc = self.ks.servicio_sync()
        if not svc.listo():
            self.skipTest("El espejo local aun no esta completo")
        p = self.ks.build_kardex(EMPRESA, self.pid, refresh=True)
        self.assertEqual(p["origen_datos"], "espejo local (SQLite)")
        self.assertAlmostEqual(p["kpi"]["saldo"], self.stock_odoo(), places=2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
