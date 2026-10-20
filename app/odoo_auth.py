"""Kardex Fimas - autenticacion contra Odoo y sesiones en memoria.

La contrasena nunca se almacena: se valida contra el ERP via XML-RPC
(common.authenticate) y la sesion local identifica al usuario por su UID de Odoo.
"""
import http.client
import secrets
import threading
import xmlrpc.client
from datetime import datetime, timedelta

SESION_HORAS = 8
COOKIE = "kardex_session"
ROL_ADMIN = "admin"
ROL_USUARIO = "usuario_estándar"

_lock = threading.RLock()
_sesiones = {}
_cfg = {"url": None, "db": None, "svc_user": None, "svc_pwd": None}
_grupo_system = {"id": None, "buscado": False}


def configurar(url, db, svc_user=None, svc_pwd=None):
    """Datos de conexion (cuenta de servicio) para resolver el rol de administrador."""
    _cfg.update(url=url, db=db, svc_user=svc_user, svc_pwd=svc_pwd)


def group_system_id():
    """res_id del grupo Odoo base.group_system (administrador de configuracion)."""
    if _grupo_system["buscado"]:
        return _grupo_system["id"]
    if not _cfg["url"] or not _cfg["svc_user"]:
        return None
    try:
        common = xmlrpc.client.ServerProxy(f"{_cfg['url']}/xmlrpc/2/common", allow_none=True)
        svc_uid = common.authenticate(_cfg["db"], _cfg["svc_user"], _cfg["svc_pwd"], {})
        if not svc_uid:
            return None
        obj = xmlrpc.client.ServerProxy(f"{_cfg['url']}/xmlrpc/2/object", allow_none=True)
        filas = obj.execute_kw(_cfg["db"], svc_uid, _cfg["svc_pwd"],
                               "ir.model.data", "search_read",
                               [[["model", "=", "res.groups"], ["module", "=", "base"],
                                 ["name", "=", "group_system"]]],
                               {"fields": ["res_id"], "limit": 1})
        if filas:
            _grupo_system["id"] = filas[0]["res_id"]
            _grupo_system["buscado"] = True
    except (OSError, http.client.HTTPException, xmlrpc.client.ProtocolError,
            xmlrpc.client.ResponseError, xmlrpc.client.Fault):
        return None
    return _grupo_system["id"]


def es_admin(groups_ids):
    gid = group_system_id()
    return bool(gid) and gid in (groups_ids or [])


def ahora():
    return datetime.now()


def _limpiar_vencidas():
    vencidas = [t for t, s in _sesiones.items() if s["expira"] <= ahora()]
    for t in vencidas:
        _sesiones.pop(t, None)


def autenticar(url, db, login, password, svc_user=None, svc_pwd=None):
    """Valida credenciales contra Odoo. Devuelve {ok, usuario} o {ok:False, codigo}."""
    login = (login or "").strip()
    password = password or ""
    if not login or not password:
        return {"ok": False, "codigo": "credenciales"}
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid = common.authenticate(db, login, password, {})
        if not uid and "@" in login and svc_user:
            svc = common.authenticate(db, svc_user, svc_pwd, {})
            if svc:
                obj = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
                filas = obj.execute_kw(db, svc, svc_pwd, "res.users", "search_read",
                                       [[["email", "=ilike", login]]],
                                       {"fields": ["login"], "limit": 2})
                if len(filas) == 1:
                    uid = common.authenticate(db, filas[0]["login"], password, {})
                    if uid:
                        login = filas[0]["login"]
        if not uid:
            return {"ok": False, "codigo": "credenciales"}
        obj = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
        datos = obj.execute_kw(db, uid, password, "res.users", "read", [uid],
                               {"fields": ["name", "login", "email", "groups_id"]})
        u = datos[0] if datos else {}
        rol = ROL_ADMIN if es_admin(u.get("groups_id")) else ROL_USUARIO
        return {"ok": True, "usuario": {"id": uid,
                                        "nombre": u.get("name") or login,
                                        "login": u.get("login") or login,
                                        "email": u.get("email") or "",
                                        "rol": rol}}
    except xmlrpc.client.Fault:
        return {"ok": False, "codigo": "odoo_error"}
    except (OSError, http.client.HTTPException, xmlrpc.client.ProtocolError,
            xmlrpc.client.ResponseError):
        return {"ok": False, "codigo": "odoo_offline"}


def crear_sesion(usuario):
    token = secrets.token_urlsafe(32)
    expira = ahora() + timedelta(hours=SESION_HORAS)
    with _lock:
        _limpiar_vencidas()
        _sesiones[token] = {"expira": expira, "usuario": dict(usuario)}
    return token


def usuario_de_sesion(token):
    if not token:
        return None
    with _lock:
        s = _sesiones.get(token)
        if not s:
            return None
        if s["expira"] <= ahora():
            _sesiones.pop(token, None)
            return None
        return dict(s["usuario"])


def cerrar_sesion(token):
    if not token:
        return
    with _lock:
        _sesiones.pop(token, None)
