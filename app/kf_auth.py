"""Kardex Fimas - autenticacion, sesiones y administracion de usuarios.

Persistencia en SQLite (usuarios.db). Contrasenas con PBKDF2-HMAC-SHA256 y
salt por usuario; nunca se almacenan en texto plano.
"""
import hashlib
import hmac
import os
import secrets
import sqlite3
import threading
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "usuarios.db")

MAX_USUARIOS = 10
MAX_INTENTOS = 5
BLOQUEO_MINUTOS = 1
SESION_HORAS = 8
ITERACIONES = 200_000
COOKIE = "kardex_session"

_lock = threading.RLock()
_local = threading.local()

SCHEMA = """
CREATE TABLE IF NOT EXISTS usuarios (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre            TEXT NOT NULL,
    email             TEXT NOT NULL UNIQUE,
    password_hash     TEXT NOT NULL,
    salt              TEXT NOT NULL,
    estado            TEXT NOT NULL DEFAULT 'activo',
    rol               TEXT NOT NULL DEFAULT 'usuario',
    intentos_fallidos INTEGER NOT NULL DEFAULT 0,
    bloqueado_hasta   TEXT,
    creado_en         TEXT NOT NULL,
    ultimo_acceso     TEXT
);
CREATE TABLE IF NOT EXISTS sesiones (
    token       TEXT PRIMARY KEY,
    usuario_id  INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
    creada_en   TEXT NOT NULL,
    expira_en   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_sesiones_usuario ON sesiones(usuario_id);
"""


def _conn():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with _lock:
        con = _conn()
        try:
            con.executescript(SCHEMA)
            con.commit()
            con.execute("DELETE FROM sesiones WHERE expira_en < ?", (ahora_iso(),))
            con.commit()
        finally:
            con.close()


def ahora():
    return datetime.now()


def ahora_iso():
    return ahora().strftime("%Y-%m-%d %H:%M:%S")


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), ITERACIONES)
    return dk.hex(), salt


def verificar_password(password, hash_hex, salt):
    calc, _ = hash_password(password, salt)
    return hmac.compare_digest(calc, hash_hex)


def email_valido(email):
    return bool(email) and "@" in email and "." in email.split("@")[-1] and len(email) >= 5


def contar_usuarios():
    with _lock:
        con = _conn()
        try:
            return con.execute("SELECT COUNT(*) c FROM usuarios").fetchone()["c"]
        finally:
            con.close()


def listar_usuarios():
    with _lock:
        con = _conn()
        try:
            return [dict(r) for r in con.execute(
                "SELECT id,nombre,email,estado,rol,intentos_fallidos,bloqueado_hasta,creado_en,ultimo_acceso "
                "FROM usuarios ORDER BY nombre COLLATE NOCASE")]
        finally:
            con.close()


def obtener_usuario(usuario_id):
    with _lock:
        con = _conn()
        try:
            row = con.execute("SELECT * FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
            return dict(row) if row else None
        finally:
            con.close()


def _normaliza_email(email):
    return (email or "").strip().lower()


def bootstrap_admin(nombre, email, password):
    with _lock:
        if contar_usuarios() > 0:
            return {"ok": False, "codigo": "ya_inicializado",
                    "mensaje": "El sistema ya tiene usuarios registrados."}
        return crear_usuario(nombre, email, password, rol="admin")


def crear_usuario(nombre, email, password, rol="usuario", estado="activo"):
    nombre = (nombre or "").strip()
    email = _normaliza_email(email)
    if len(nombre) < 3:
        return {"ok": False, "codigo": "nombre_invalido",
                "mensaje": "El nombre debe tener al menos 3 caracteres."}
    if not email_valido(email):
        return {"ok": False, "codigo": "email_invalido",
                "mensaje": "Ingrese un correo electronico valido."}
    if not password or len(password) < 8:
        return {"ok": False, "codigo": "password_corta",
                "mensaje": "La contrasena debe tener al menos 8 caracteres."}
    if rol not in ("admin", "usuario"):
        rol = "usuario"
    if estado not in ("activo", "inactivo"):
        estado = "activo"
    with _lock:
        con = _conn()
        try:
            if con.execute("SELECT COUNT(*) c FROM usuarios").fetchone()["c"] >= MAX_USUARIOS:
                return {"ok": False, "codigo": "limite_usuarios",
                        "mensaje": f"No se pueden registrar mas de {MAX_USUARIOS} usuarios."}
            if con.execute("SELECT 1 FROM usuarios WHERE email=?", (email,)).fetchone():
                return {"ok": False, "codigo": "email_duplicado",
                        "mensaje": "Ese correo electronico ya esta registrado."}
            pwd, salt = hash_password(password)
            cur = con.execute(
                "INSERT INTO usuarios (nombre,email,password_hash,salt,estado,rol,creado_en) "
                "VALUES (?,?,?,?,?,?,?)", (nombre, email, pwd, salt, estado, rol, ahora_iso()))
            con.commit()
            nuevo_id = cur.lastrowid
        finally:
            con.close()
    return {"ok": True, "id": nuevo_id, "usuario": {"id": nuevo_id, "nombre": nombre, "email": email,
                                                      "rol": rol, "estado": estado}}


def actualizar_usuario(usuario_id, nombre=None, email=None, rol=None, estado=None):
    with _lock:
        con = _conn()
        try:
            actual = con.execute("SELECT * FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
            if not actual:
                return {"ok": False, "codigo": "no_existe", "mensaje": "El usuario no existe."}
            if rol not in ("admin", "usuario"):
                rol = actual["rol"]
            if estado not in ("activo", "inactivo"):
                estado = actual["estado"]
            email = _normaliza_email(email) if email else actual["email"]
            if not email_valido(email):
                return {"ok": False, "codigo": "email_invalido",
                        "mensaje": "Ingrese un correo electronico valido."}
            if con.execute("SELECT 1 FROM usuarios WHERE email=? AND id<>?", (email, usuario_id)).fetchone():
                return {"ok": False, "codigo": "email_duplicado",
                        "mensaje": "Ese correo electronico ya esta registrado."}
            if rol == "admin":
                con.execute("UPDATE usuarios SET rol='usuario' WHERE rol='admin' AND id<>?", (usuario_id,))
            elif actual["rol"] == "admin":
                admins = con.execute("SELECT COUNT(*) c FROM usuarios "
                                     "WHERE rol='admin' AND id<>? AND estado='activo'",
                                     (usuario_id,)).fetchone()["c"]
                if admins == 0:
                    return {"ok": False, "codigo": "ultimo_admin",
                            "mensaje": "Debe existir al menos un administrador activo."}
            if estado == "inactivo" and actual["rol"] == "admin":
                admins = con.execute("SELECT COUNT(*) c FROM usuarios WHERE rol='admin' AND id<>? AND estado='activo'",
                                     (usuario_id,)).fetchone()["c"]
                if admins == 0:
                    return {"ok": False, "codigo": "ultimo_admin",
                            "mensaje": "Debe existir al menos un administrador activo."}
            con.execute("UPDATE usuarios SET nombre=?, email=?, rol=?, estado=? WHERE id=?",
                        ((nombre or actual["nombre"]).strip(), email, rol, estado, usuario_id))
            if estado == "inactivo":
                con.execute("DELETE FROM sesiones WHERE usuario_id=?", (usuario_id,))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "mensaje": "Usuario actualizado."}


def eliminar_usuario(usuario_id, actor_id):
    if usuario_id == actor_id:
        return {"ok": False, "codigo": "auto_bloqueo", "mensaje": "No puede eliminar su propio usuario."}
    with _lock:
        con = _conn()
        try:
            u = con.execute("SELECT * FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
            if not u:
                return {"ok": False, "codigo": "no_existe", "mensaje": "El usuario no existe."}
            if u["rol"] == "admin":
                admins = con.execute("SELECT COUNT(*) c FROM usuarios WHERE rol='admin' AND id<>? AND estado='activo'",
                                     (usuario_id,)).fetchone()["c"]
                if admins == 0:
                    return {"ok": False, "codigo": "ultimo_admin",
                            "mensaje": "Debe existir al menos un administrador activo."}
            con.execute("DELETE FROM sesiones WHERE usuario_id=?", (usuario_id,))
            con.execute("DELETE FROM usuarios WHERE id=?", (usuario_id,))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "mensaje": "Usuario eliminado."}


def restablecer_password(usuario_id, password, requiere_actual=False, actual=None, conservar_token=None):
    if not password or len(password) < 8:
        return {"ok": False, "codigo": "password_corta",
                "mensaje": "La contrasena debe tener al menos 8 caracteres."}
    with _lock:
        con = _conn()
        try:
            u = con.execute("SELECT * FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
            if not u:
                return {"ok": False, "codigo": "no_existe", "mensaje": "El usuario no existe."}
            if requiere_actual and not verificar_password(actual or "", u["password_hash"], u["salt"]):
                return {"ok": False, "codigo": "password_actual",
                        "mensaje": "La contrasena actual no es correcta."}
            pwd, salt = hash_password(password)
            con.execute("UPDATE usuarios SET password_hash=?, salt=?, intentos_fallidos=0, "
                        "bloqueado_hasta=NULL WHERE id=?", (pwd, salt, usuario_id))
            if conservar_token:
                con.execute("DELETE FROM sesiones WHERE usuario_id=? AND token<>?",
                            (usuario_id, conservar_token))
            else:
                con.execute("DELETE FROM sesiones WHERE usuario_id=?", (usuario_id,))
            con.commit()
        finally:
            con.close()
    return {"ok": True, "mensaje": "Contrasena actualizada."}


def autenticar(email, password):
    email = _normaliza_email(email)
    with _lock:
        con = _conn()
        try:
            u = con.execute("SELECT * FROM usuarios WHERE email=?", (email,)).fetchone()
            if not u:
                return {"ok": False, "codigo": "credenciales"}
            if u["estado"] != "activo":
                return {"ok": False, "codigo": "credenciales"}
            hasta = u["bloqueado_hasta"]
            if hasta:
                try:
                    limite = datetime.strptime(hasta, "%Y-%m-%d %H:%M:%S")
                except (TypeError, ValueError):
                    limite = ahora()
                if limite > ahora():
                    return {"ok": False, "codigo": "bloqueado",
                            "espera": max(1, int((limite - ahora()).total_seconds()))}
                con.execute("UPDATE usuarios SET intentos_fallidos=0, bloqueado_hasta=NULL WHERE id=?", (u["id"],))
            if not verificar_password(password or "", u["password_hash"], u["salt"]):
                intentos = u["intentos_fallidos"] + 1
                bloqueo = None
                if intentos >= MAX_INTENTOS:
                    bloqueo = (ahora() + timedelta(minutes=BLOQUEO_MINUTOS)).strftime("%Y-%m-%d %H:%M:%S")
                    intentos = 0
                con.execute("UPDATE usuarios SET intentos_fallidos=?, bloqueado_hasta=? WHERE id=?",
                            (intentos, bloqueo, u["id"]))
                con.commit()
                return {"ok": False, "codigo": "credenciales"}
            con.execute("UPDATE usuarios SET intentos_fallidos=0, bloqueado_hasta=NULL, ultimo_acceso=? WHERE id=?",
                        (ahora_iso(), u["id"]))
            con.commit()
            return {"ok": True, "usuario": {"id": u["id"], "nombre": u["nombre"], "email": u["email"],
                                            "rol": u["rol"], "estado": u["estado"]}}
        finally:
            con.close()


def crear_sesion(usuario_id):
    token = secrets.token_urlsafe(32)
    expira = (ahora() + timedelta(hours=SESION_HORAS)).strftime("%Y-%m-%d %H:%M:%S")
    with _lock:
        con = _conn()
        try:
            con.execute("INSERT INTO sesiones (token,usuario_id,creada_en,expira_en) VALUES (?,?,?,?)",
                        (token, usuario_id, ahora_iso(), expira))
            con.commit()
        finally:
            con.close()
    return token, expira


def usuario_de_sesion(token):
    if not token:
        return None
    with _lock:
        con = _conn()
        try:
            s = con.execute("SELECT * FROM sesiones WHERE token=?", (token,)).fetchone()
            if not s:
                return None
            try:
                expira = datetime.strptime(s["expira_en"], "%Y-%m-%d %H:%M:%S")
            except (TypeError, ValueError):
                con.execute("DELETE FROM sesiones WHERE token=?", (token,))
                con.commit()
                return None
            if expira <= ahora():
                con.execute("DELETE FROM sesiones WHERE token=?", (token,))
                con.commit()
                return None
            u = con.execute("SELECT * FROM usuarios WHERE id=?", (s["usuario_id"],)).fetchone()
            if not u or u["estado"] != "activo":
                con.execute("DELETE FROM sesiones WHERE token=?", (token,))
                con.commit()
                return None
            return {"id": u["id"], "nombre": u["nombre"], "email": u["email"], "rol": u["rol"],
                    "estado": u["estado"], "expira_en": s["expira_en"]}
        finally:
            con.close()


def cerrar_sesion(token):
    if not token:
        return
    with _lock:
        con = _conn()
        try:
            con.execute("DELETE FROM sesiones WHERE token=?", (token,))
            con.commit()
        finally:
            con.close()


def estado_sistema():
    total = contar_usuarios()
    return {"usuarios": total, "max_usuarios": MAX_USUARIOS, "max_intentos": MAX_INTENTOS,
            "bloqueo_minutos": BLOQUEO_MINUTOS, "sesion_horas": SESION_HORAS,
            "requiere_instalacion": total == 0}


init_db()