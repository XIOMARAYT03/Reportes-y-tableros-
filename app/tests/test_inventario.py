"""Pruebas unitarias del motor de inventario (sin conexion al ERP).

    python -m unittest discover -s app/tests -v

Las ubicaciones y unidades son ficticias pero replican la estructura real de
Odoo 18 (usage, company_id, category_id, factor, factor_inv, rounding).
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import inventario  # noqa: E402

# --------------------------------------------------------------------------- #
# ERP ficticio
# --------------------------------------------------------------------------- #
UOMS = [
    {"id": 1, "name": "Units", "category_id": [1, "Units"], "factor": 1.0,
     "factor_inv": 1.0, "rounding": 0.01, "uom_type": "bigger", "active": True},
    # En Odoo factor = ratio de la UoM respecto de la unidad de referencia de la
    # categoria (1 docena = 1/12 unidades) y factor_inv = su reciproco.
    {"id": 2, "name": "Dozen", "category_id": [1, "Units"], "factor": 1.0 / 12.0,
     "factor_inv": 12.0, "rounding": 0.01, "uom_type": "bigger", "active": True},
    {"id": 3, "name": "Box", "category_id": [1, "Units"], "factor": 1.0 / 6.0,
     "factor_inv": 6.0, "rounding": 1.0, "uom_type": "bigger", "active": True},
    {"id": 4, "name": "Kg", "category_id": [2, "Weight"], "factor": 1.0,
     "factor_inv": 1.0, "rounding": 0.01, "uom_type": "bigger", "active": True},
    {"id": 5, "name": "Pallet", "category_id": [1, "Units"], "factor": 1.0,
     "factor_inv": 1.0, "rounding": 1.0, "uom_type": "reference", "active": True},
]

LOCS = [
    {"id": 10, "name": "Stock", "complete_name": "C1/Stock", "usage": "internal",
     "company_id": 1, "active": True},
    {"id": 11, "name": "Input", "complete_name": "C1/Input", "usage": "internal",
     "company_id": 1, "active": True},
    {"id": 12, "name": "Output", "complete_name": "C1/Output", "usage": "internal",
     "company_id": 1, "active": True},
    {"id": 20, "name": "Stock", "complete_name": "C2/Stock", "usage": "internal",
     "company_id": 2, "active": True},
    {"id": 30, "name": "Suppliers", "complete_name": "Suppliers", "usage": "supplier",
     "company_id": 1, "active": True},
    {"id": 31, "name": "Customers", "complete_name": "Customers", "usage": "customer",
     "company_id": 1, "active": True},
    {"id": 40, "name": "Inventory", "complete_name": "Virtual/Inventory", "usage": "inventory",
     "company_id": 1, "active": True},
    {"id": 41, "name": "Transit", "complete_name": "Virtual/Transit", "usage": "transit",
     "company_id": 1, "active": True},
    {"id": 42, "name": "Production", "complete_name": "Virtual/Production", "usage": "production",
     "company_id": 1, "active": True},
    {"id": 43, "name": "View", "complete_name": "C1/View", "usage": "view",
     "company_id": 1, "active": True},
]


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
    if isinstance(valor, dict):
        return valor.get("display_name") or valor.get("name")
    return valor


def fake_sr(modelo, dominio=None, campos=None, order=None, limit=None, offset=None):
    datos = {"uom.uom": UOMS, "stock.location": LOCS}.get(modelo, [])
    return [dict(f) for f in datos]


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
ALM = {10, 11, 12}


def mov(id=1, fecha="2026-01-01 10:00:00", cantidad=10.0, origen=30, destino=10,
        uom=1, estado="done", ptype=None, producto=99, empresa=1):
    return {"id": id, "date": fecha, "quantity": cantidad, "product_qty": cantidad,
            "product_uom": uom, "location_id": origen, "location_dest_id": destino,
            "state": estado, "picking_type_id": ptype, "product_id": producto,
            "company_id": empresa, "name": f"MOV/{id}", "picking_id": None,
            "purchase_line_id": None, "sale_line_id": None, "origin_returned_move_id": None}


PROD = {"id": 99, "default_code": "TEST01", "name": "Producto de prueba",
        "uom_id": [1, "Units"], "standard_price": 2.0, "categ_id": [1, "Cat"]}


def kardex(movimientos, quants=None, empresa=1, capas=None):
    alm = inventario.almacenes(empresa)
    locs = inventario.ubicaciones()
    return inventario.calcular_kardex(empresa, PROD, movimientos, capas or {},
                                      alm, locs, quants=quants or [])


def tearDownModule():
    """Restaura el acceso real al ERP.

    El ERP ficticio de estas pruebas queda instalado en `inventario` y no debe
    filtrarse a otras pruebas (por ejemplo las que consultan Odoo en vivo).
    """
    try:
        import kardex_server as ks
        inventario.configurar(ks.sr, ks.id_of, ks.ref)
    except Exception:
        inventario.configurar(None, None, None)
    inventario.invalidar_catalogos()


# --------------------------------------------------------------------------- #
# Pruebas
# --------------------------------------------------------------------------- #
class TestCatalogos(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def test_almacenes_solo_internas_de_la_empresa(self):
        self.assertEqual(inventario.almacenes(1), {10, 11, 12})
        self.assertEqual(inventario.almacenes(2), {20})

    def test_catalogo_de_unidades(self):
        u = inventario.uoms()
        self.assertAlmostEqual(u[2]["factor_inv"], 12.0, places=6)
        self.assertAlmostEqual(u[3]["precision"], 1.0)
        self.assertEqual(u[4]["category_id"], 2)


class TestClasificacion(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def clas(self, m, empresa=1):
        alm = inventario.almacenes(empresa)
        return inventario.clasificar(m, alm, inventario.ubicaciones())

    def test_entrada_desde_proveedor(self):
        clase, sub, signo = self.clas(mov(origen=30, destino=10))
        self.assertEqual(clase, "entrada")
        self.assertEqual(sub, "compra")
        self.assertEqual(signo, 1.0)

    def test_salida_a_cliente(self):
        clase, sub, signo = self.clas(mov(origen=10, destino=31))
        self.assertEqual(clase, "salida")
        self.assertEqual(signo, -1.0)

    def test_traslado_interno_no_mueve_stock(self):
        clase, sub, signo = self.clas(mov(origen=10, destino=11))
        self.assertEqual(clase, "traslado")
        self.assertEqual(signo, 0.0)

    def test_virtual_a_virtual_no_mueve_stock(self):
        clase, _, signo = self.clas(mov(origen=40, destino=41))
        self.assertEqual(signo, 0.0)
        self.assertNotEqual(clase, "entrada")
        self.assertNotEqual(clase, "salida")

    def test_ajuste_desde_inventario_es_entrada(self):
        clase, _, signo = self.clas(mov(origen=40, destino=10))
        self.assertEqual(clase, "entrada")
        self.assertEqual(signo, 1.0)

    def test_ajuste_hacia_inventario_es_salida(self):
        clase, _, signo = self.clas(mov(origen=10, destino=40))
        self.assertEqual(clase, "salida")
        self.assertEqual(signo, -1.0)

    def test_traspaso_entre_empresas_entra_en_la_que_recibe(self):
        # origen en la empresa 2, destino en la empresa 1
        clase, sub, signo = self.clas(mov(origen=20, destino=10))
        self.assertEqual(clase, "entrada")
        self.assertEqual(sub, "traspaso_entre_empresas")
        # y al revés, en la empresa 2 es una salida
        clase, sub, signo = self.clas(mov(origen=20, destino=10), empresa=2)
        self.assertEqual(clase, "salida")
        self.assertEqual(sub, "traspaso_entre_empresas")

    def test_devolucion_de_cliente_es_entrada(self):
        clase, sub, signo = self.clas(mov(origen=31, destino=10))
        self.assertEqual(clase, "entrada")
        self.assertEqual(signo, 1.0)

    def test_devolucion_a_proveedor_es_salida(self):
        clase, sub, signo = self.clas(mov(origen=10, destino=30))
        self.assertEqual(clase, "salida")
        self.assertEqual(signo, -1.0)

    def test_produccion_es_entrada_y_consumo_es_salida(self):
        self.assertEqual(self.clas(mov(origen=42, destino=10))[0], "entrada")
        self.assertEqual(self.clas(mov(origen=10, destino=42))[0], "salida")

    def test_ubicacion_de_vista_no_altera_el_stock(self):
        clase, _, signo = self.clas(mov(origen=43, destino=10))
        self.assertEqual(signo, 1.0)   # entra al almacen desde una vista
        clase, _, signo = self.clas(mov(origen=10, destino=43))
        self.assertEqual(signo, -1.0)

    def test_el_tipo_de_picking_no_define_el_signo(self):
        # XD (cross dock) con ubicaciones internas no es una entrada
        clase, _, signo = self.clas(mov(origen=10, destino=12, ptype="XD"))
        self.assertEqual(clase, "traslado")
        self.assertEqual(signo, 0.0)


class TestConversion(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def test_misma_unidad(self):
        r = inventario.convertir(10, 1, 1)
        self.assertTrue(r["ok"])
        self.assertAlmostEqual(r["cantidad"], 10.0)

    def test_conversion_desde_caja_a_unidad(self):
        r = inventario.convertir(3, 3, 1)
        self.assertTrue(r["ok"])
        self.assertAlmostEqual(r["cantidad"], 18.0)

    def test_conversion_desde_decena_a_unidad(self):
        r = inventario.convertir(2, 2, 1)
        self.assertTrue(r["ok"])
        self.assertAlmostEqual(r["cantidad"], 24.0)

    def test_categorias_incompatibles_se_rechazan(self):
        r = inventario.convertir(5, 1, 4)
        self.assertFalse(r["ok"])
        self.assertIn("categor", r["motivo"].lower())

    def test_redondeo_a_la_precision_del_destino(self):
        # 1/3 de caja = 2 unidades (caja = 6 unidades, precision 0.01)
        r = inventario.convertir(1.0 / 3.0, 3, 1)
        self.assertTrue(r["ok"])
        self.assertAlmostEqual(r["cantidad"], 2.0)

    def test_redondeo_al_entero_cuando_la_unidad_lo_exige(self):
        # destino "Pallet" con rounding 1.0 -> 2.4 unidades se redondea a 2
        r = inventario.convertir(2.4, 1, 5)
        self.assertTrue(r["ok"])
        self.assertEqual(r["cantidad"], 2.0)
        self.assertEqual(r["precision"], 1.0)

    def test_la_conversion_a_la_misma_unidad_no_altera_el_valor(self):
        r = inventario.convertir(1.0 / 3.0, 3, 3)
        self.assertTrue(r["ok"])
        self.assertAlmostEqual(r["cantidad"], 1.0 / 3.0, places=6)


class TestKardex(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def test_saldo_de_entradas_y_salidas(self):
        ms = [mov(id=1, fecha="2026-01-01 08:00:00", cantidad=100, origen=30, destino=10),
              mov(id=2, fecha="2026-01-02 08:00:00", cantidad=30, origen=10, destino=31)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["entradas"], 100.0)
        self.assertAlmostEqual(r["kpi"]["salidas"], 30.0)
        self.assertAlmostEqual(r["kpi"]["saldo"], 70.0)
        self.assertEqual(len(r["rows"]), 2)

    def test_el_orden_de_las_fechas_no_afecta_el_saldo(self):
        ms = [mov(id=1, fecha="2026-01-05 08:00:00", cantidad=50, origen=30, destino=10),
              mov(id=2, fecha="2026-01-01 08:00:00", cantidad=20, origen=10, destino=31)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 30.0)
        self.assertEqual(r["rows"][0]["date"], "2026-01-01 08:00:00")

    def test_las_transferencias_internas_no_alteran_el_saldo(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=40, origen=10, destino=11, ptype="PICK"),
              mov(id=3, cantidad=40, origen=11, destino=12, ptype="PACK")]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 100.0)
        self.assertEqual(r["kpi"]["internos"], 2)
        self.assertEqual(len(r["rows"]), 1)

    def test_los_movimientos_virtuales_no_alteran_el_saldo(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=100, origen=40, destino=41),
              mov(id=3, cantidad=5, origen=10, destino=40)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 95.0)

    def test_las_cantidades_se_convierten_a_la_unidad_base(self):
        ms = [mov(id=1, cantidad=2, origen=30, destino=10, uom=2)]   # 2 docenas
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["entradas"], 24.0)
        self.assertEqual(r["unidad"], "Units")

    def test_producto_sin_movimientos(self):
        r = kardex([])
        self.assertEqual(r["rows"], [])
        self.assertAlmostEqual(r["kpi"]["saldo"], 0.0)

    def test_saldo_cero_tras_una_entrada_y_una_salida_iguales(self):
        ms = [mov(id=1, cantidad=75, origen=30, destino=10),
              mov(id=2, cantidad=75, origen=10, destino=31)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 0.0)

    def test_los_movimientos_no_completados_se_ignoran(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=40, origen=10, destino=31, estado="cancel")]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 100.0)
        self.assertEqual(r["kpi"]["movimientos"], 1)

    def test_el_valor_usa_las_capas_del_erp(self):
        ms = [mov(id=1, cantidad=10, origen=30, destino=10)]
        capas = {1: [{"quantity": 10.0, "value": 55.0, "unit_cost": 5.5}]}
        r = kardex(ms, capas=capas)
        self.assertAlmostEqual(r["kpi"]["valor"], 55.0)
        self.assertAlmostEqual(r["rows"][0]["cost"], 5.5)


class TestRegresionCIN058(unittest.TestCase):
    """Caso reportado en CIN058: el kardex no cuadraba con las existencias.

    El movimiento de ajuste entre dos ubicaciones virtuales (transito ->
    inventario) se contaba antes como una entrada y sumaba mercaderia que nunca
    estuvo en el almacen.
    """

    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def test_el_ajuste_entre_ubicaciones_virtuales_no_es_entrada(self):
        # 41 = Virtual/Transit, 40 = Virtual/Inventory (ambas fuera del almacen)
        ms = [mov(id=1, fecha="2026-03-09 08:00:00", cantidad=100, origen=30, destino=10),
              mov(id=2, fecha="2026-03-09 09:00:00", cantidad=2, origen=41, destino=40),
              mov(id=3, fecha="2026-03-09 10:00:00", cantidad=47, origen=40, destino=10)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["saldo"], 147.0)
        self.assertAlmostEqual(r["kpi"]["entradas"], 147.0)
        self.assertAlmostEqual(r["kpi"]["salidas"], 0.0)
        # el ajuste virtual no genera fila en el detalle
        self.assertEqual(len(r["rows"]), 2)
        for fila in r["rows"]:
            self.assertNotIn("Virtual/Transit", fila["origen"])

    def test_el_ajuste_virtual_no_cambia_el_saldo(self):
        base = [mov(id=1, cantidad=100, origen=30, destino=10)]
        con_ajuste = base + [mov(id=2, cantidad=2, origen=41, destino=40)]
        self.assertAlmostEqual(kardex(con_ajuste)["kpi"]["saldo"],
                               kardex(base)["kpi"]["saldo"])

    def test_la_entrada_desde_ajuste_de_inventario_si_cuenta(self):
        # al revés: del inventario virtual al almacen si es una entrada real
        ms = [mov(id=1, cantidad=30, origen=40, destino=10)]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["entradas"], 30.0)
        self.assertAlmostEqual(r["kpi"]["saldo"], 30.0)

    def test_el_traspaso_entre_empresas_se_cuenta_una_vez_en_cada_empresa(self):
        # salida de la empresa 1 hacia el transito
        ms = [mov(id=1, cantidad=100, origen=10, destino=41, empresa=1)]
        r1 = kardex(ms, empresa=1)
        self.assertAlmostEqual(r1["kpi"]["salidas"], 100.0)
        self.assertAlmostEqual(r1["kpi"]["saldo"], -100.0)
        # y entrada en la empresa 2 (otra empresa, otra ubicacion interna)
        ms2 = [mov(id=2, cantidad=100, origen=41, destino=20, empresa=2)]
        r2 = kardex(ms2, empresa=2)
        self.assertAlmostEqual(r2["kpi"]["entradas"], 100.0)
        self.assertAlmostEqual(r2["kpi"]["saldo"], 100.0)
        # en la empresa 2 el movimiento de salida no cuenta
        r3 = kardex(ms, empresa=2)
        self.assertAlmostEqual(r3["kpi"]["saldo"], 0.0)

    def test_la_cadena_pick_pack_out_no_duplica_la_salida(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=100, origen=10, destino=11, ptype="PICK"),
              mov(id=3, cantidad=100, origen=11, destino=12, ptype="PACK"),
              mov(id=4, cantidad=100, origen=12, destino=31, ptype="OUT")]
        r = kardex(ms)
        self.assertAlmostEqual(r["kpi"]["salidas"], 100.0)
        self.assertAlmostEqual(r["kpi"]["saldo"], 0.0)
        self.assertEqual(r["kpi"]["internos"], 2)
        self.assertEqual(len(r["rows"]), 2)

    def test_el_kardex_se_concilia_con_los_quants_del_alcance(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=2, origen=41, destino=40),
              mov(id=3, cantidad=30, origen=10, destino=31)]
        quants = [{"location_id": 10, "quantity": 70.0, "reserved_quantity": 0.0},
                  {"location_id": 40, "quantity": -2.0, "reserved_quantity": 0.0},
                  {"location_id": 31, "quantity": 30.0, "reserved_quantity": 0.0}]
        r = kardex(ms, quants=quants)
        self.assertAlmostEqual(r["kpi"]["saldo"], 70.0)
        self.assertAlmostEqual(r["kpi"]["fisico"], 70.0)
        c = inventario.conciliar(1, PROD, r["kpi"]["saldo"], quants, movimientos=ms,
                                  alm=inventario.almacenes(1), locs=inventario.ubicaciones())
        self.assertEqual(c["estado"], "conciliado")
        self.assertAlmostEqual(c["diferencia"], 0.0)

    def test_cada_fila_lleva_el_usuario_que_realizo_el_movimiento(self):
        ms = [mov(id=1, cantidad=100, origen=30, destino=10),
              mov(id=2, cantidad=40, origen=10, destino=11, ptype="PICK")]
        ms[0]["usuario"] = "JUAN PEREZ"
        ms[1]["usuario"] = "MARIA LOPEZ"
        r = kardex(ms)
        self.assertEqual(r["rows"][0]["usuario"], "JUAN PEREZ")
        self.assertEqual(r["internals"][0]["usuario"], "MARIA LOPEZ")

    def test_la_fila_sin_usuario_no_falla(self):
        r = kardex([mov(id=1, cantidad=100, origen=30, destino=10)])
        self.assertEqual(r["rows"][0]["usuario"], "")


class TestConciliacion(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inventario.configurar(fake_sr, _id_of, _ref)
        inventario.invalidar_catalogos()

    def test_conciliado_cuando_coincide(self):
        alm = inventario.almacenes(1)
        locs = inventario.ubicaciones()
        ms = [mov(id=1, cantidad=100, origen=30, destino=10)]
        quants = [{"location_id": 10, "quantity": 100.0, "reserved_quantity": 10.0}]
        c = inventario.conciliar(1, PROD, 100.0, quants, movimientos=ms, alm=alm, locs=locs)
        self.assertEqual(c["estado"], "conciliado")
        self.assertAlmostEqual(c["diferencia"], 0.0)
        self.assertAlmostEqual(c["stock_odoo"], 100.0)
        self.assertAlmostEqual(c["reservado"], 10.0)

    def test_diferencia_detectada(self):
        alm = inventario.almacenes(1)
        locs = inventario.ubicaciones()
        ms = [mov(id=1, cantidad=100, origen=30, destino=10)]
        quants = [{"location_id": 10, "quantity": 90.0, "reserved_quantity": 0.0}]
        c = inventario.conciliar(1, PROD, 100.0, quants, movimientos=ms, alm=alm, locs=locs)
        self.assertEqual(c["estado"], "diferencia")
        self.assertAlmostEqual(c["diferencia"], 10.0)
        self.assertTrue(c["posibles_causas"])

    def test_los_quants_fuera_del_alcance_no_cuentan(self):
        alm = inventario.almacenes(1)
        locs = inventario.ubicaciones()
        ms = [mov(id=1, cantidad=100, origen=30, destino=10)]
        quants = [{"location_id": 10, "quantity": 100.0, "reserved_quantity": 0.0},
                  {"location_id": 31, "quantity": 50.0, "reserved_quantity": 0.0},
                  {"location_id": 40, "quantity": -20.0, "reserved_quantity": 0.0}]
        c = inventario.conciliar(1, PROD, 100.0, quants, movimientos=ms, alm=alm, locs=locs)
        self.assertAlmostEqual(c["stock_odoo"], 100.0)
        self.assertEqual(c["estado"], "conciliado")
        # solo se reportan las ubicaciones virtuales fuera de alcance
        self.assertEqual(len(c["quants_fuera_de_alcance"]), 1)

    def test_sin_movimientos_en_el_alcance(self):
        alm = inventario.almacenes(1)
        locs = inventario.ubicaciones()
        c = inventario.conciliar(1, PROD, 0.0, [], movimientos=[], alm=alm, locs=locs)
        self.assertEqual(c["estado"], "conciliado")
        self.assertTrue(c["posibles_causas"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
