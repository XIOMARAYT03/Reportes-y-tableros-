"""Permisos de acceso por usuario de Odoo y catalogo dinamico de modulos.

Persistencia en SQLite (permisos.db):
  - modulos: catalogo administrado por el admin desde /permisos (alta/baja/edicion).
  - permisos: modulos habilitados por cada usuario de Odoo (UID como llave).

Nunca se guardan credenciales. El modulo "kardex" es de sistema: no se puede
borrar ni desactivar, para no dejar al completo al sistema sin acceso.
"""
import os
import re
import sqlite3
import threading
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "permisos.db")

SISTEMA = {"kardex"}
CLAVE_RE = re.compile(r"^[a-z0-9_]{2,20}$")

_lock = threading.RLock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS modulos (
    clave        TEXT PRIMARY KEY,
    nombre       TEXT NOT NULL,
    descripcion  TEXT NOT NULL DEFAULT '',
    activo       INTEGER NOT NULL DEFAULT 1,
    creado_en    TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS permisos (
    odoo_uid      INTEGER PRIMARY KEY,
    modulos       TEXT NOT NULL,
    actualizado_en TEXT NOT NULL
);
"""


def _conn():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with _lock:
        con = _conn()
        try:
            con.executescript(SCHEMA)
            semillas = [
                ("kardex", "Kardex - Consulta de inventario",
                 "Consulta de saldos, movimientos y reportes de inventario"),
                ("despacho", "Despacho - Tiempos por fase",
                 "Tiempos de despacho por pedido: picking, packing y output"),
                ("devoluciones", "Devoluciones pendientes de compras puntuales - FINPAC",
                 "Devoluciones de clientes de FINPAC de compras puntuales que siguen "
                 "en el inventario sin venderse"),
            ]
            for clave, nombre, desc in semillas:
                if not con.execute("SELECT 1 FROM modulos WHERE clave=?", (clave,)).fetchone():
                    con.execute("INSERT INTO modulos (clave, nombre, descripcion, activo, creado_en) "
                                "VALUES (?,?,?,?,?)",
                                (clave, nombre, desc, 1,
                                 datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            con.commit()
        finally:
            con.close()


def listar_modulos():
    """Catalogo completo: [{clave, nombre, descripcion, activo, es_sistema}]."""
    with _lock:
        con = _conn()
        try:
            filas = con.execute("SELECT * FROM modulos ORDER BY nombre COLLATE NOCASE").fetchall()
        finally:
            con.close()
    return [{"clave": f["clave"], "nombre": f["nombre"], "descripcion": f["descripcion"] or "",
             "activo": bool(f["activo"]), "es_sistema": f["clave"] in SISTEMA}
            for f in filas]


def modulos_activos():
    """{clave: nombre} solo de los modulos activos."""
    return {m["clave"]: m["nombre"] for m in listar_modulos() if m["activo"]}


def crear_modulo(clave, nombre, descripcion=""):
    clave = (clave or "").strip().lower()
    nombre = (nombre or "").strip()
    if not CLAVE_RE.match(clave):
        return {"ok": False, "mensaje": "La clave debe tener de 2 a 20 caracteres "
                                        "con letras minúsculas, números o guion_bajo."}
    if len(nombre) < 2:
        return {"ok": False, "mensaje": "El nombre del módulo es demasiado corto."}
    with _lock:
        con = _conn()
        try:
            if con.execute("SELECT 1 FROM modulos WHERE clave=?", (clave,)).fetchone():
                return {"ok": False, "mensaje": f"Ya existe un módulo con la clave '{clave}'."}
            con.execute("INSERT INTO modulos (clave, nombre, descripcion, activo, creado_en) "
                        "VALUES (?,?,?,?,?)",
                        (clave, nombre, (descripcion or "").strip(), 1,
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "clave": clave}


def actualizar_modulo(clave, nombre=None, descripcion=None, activo=None):
    clave = (clave or "").strip().lower()
    with _lock:
        con = _conn()
        try:
            fila = con.execute("SELECT * FROM modulos WHERE clave=?", (clave,)).fetchone()
            if not fila:
                return {"ok": False, "mensaje": "El módulo no existe."}
            if clave in SISTEMA and activo is False:
                return {"ok": False, "mensaje": "El módulo Kardex es de sistema: no se puede desactivar."}
            nuevo_nombre = (nombre or "").strip() or fila["nombre"]
            nueva_desc = (descripcion or "").strip() if descripcion is not None else fila["descripcion"]
            nuevo_activo = fila["activo"] if activo is None else (1 if activo else 0)
            con.execute("UPDATE modulos SET nombre=?, descripcion=?, activo=? WHERE clave=?",
                        (nuevo_nombre, nueva_desc, nuevo_activo, clave))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "clave": clave}


def eliminar_modulo(clave):
    clave = (clave or "").strip().lower()
    if clave in SISTEMA:
        return {"ok": False, "mensaje": "El módulo Kardex es de sistema: no se puede eliminar."}
    with _lock:
        con = _conn()
        try:
            if not con.execute("SELECT 1 FROM modulos WHERE clave=?", (clave,)).fetchone():
                return {"ok": False, "mensaje": "El módulo no existe."}
            con.execute("DELETE FROM modulos WHERE clave=?", (clave,))
            # limpiar la clave de las asignaciones de usuarios
            for fila in con.execute("SELECT odoo_uid, modulos FROM permisos").fetchall():
                mods = [m for m in (fila["modulos"] or "").split(",") if m and m != clave]
                con.execute("UPDATE permisos SET modulos=? WHERE odoo_uid=?",
                            (",".join(mods), fila["odoo_uid"]))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "clave": clave}


def _parse(valor):
    return [m for m in (valor or "").split(",") if m]


def obtener(odoo_uid):
    """Modulos habilitados para el usuario (todos los activos por defecto si no hay fila)."""
    with _lock:
        con = _conn()
        try:
            fila = con.execute("SELECT modulos FROM permisos WHERE odoo_uid=?",
                               (odoo_uid,)).fetchone()
        finally:
            con.close()
    if not fila:
        return list(modulos_activos())
    return _parse(fila["modulos"])


def puede(odoo_uid, clave):
    """True si el modulo esta activo y el usuario lo tiene habilitado."""
    return clave in modulos_activos() and clave in obtener(odoo_uid)


def actualizar(odoo_uid, modulos):
    if not odoo_uid or int(odoo_uid) <= 0:
        return {"ok": False, "mensaje": "Usuario Odoo invalido."}
    if not isinstance(modulos, list):
        return {"ok": False, "mensaje": "Debe indicar la lista de modulos."}
    conocidos = {m["clave"] for m in listar_modulos()}
    invalidos = [m for m in modulos if m not in conocidos]
    if invalidos:
        return {"ok": False, "mensaje": f"Modulos no validos: {', '.join(invalidos)}"}
    orden = [m["clave"] for m in listar_modulos() if m["clave"] in set(modulos)]
    guardados = ",".join(orden)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _lock:
        con = _conn()
        try:
            con.execute("DELETE FROM permisos WHERE odoo_uid=?", (int(odoo_uid),))
            con.execute("INSERT INTO permisos (odoo_uid, modulos, actualizado_en) VALUES (?,?,?)",
                        (int(odoo_uid), guardados, ahora))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "modulos": orden}


init_db()
