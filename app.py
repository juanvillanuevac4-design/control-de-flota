"""Control de Flota: un reporte diario de la distribución de buses del depósito."""
import csv
import io
import json
import logging
import os
import re
import secrets
import time
from datetime import date, datetime, timedelta
from functools import wraps
from zoneinfo import ZoneInfo

import psycopg
from flask import Flask, Response, g, jsonify, redirect, request, send_from_directory, session
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from psycopg.rows import dict_row
from werkzeug.security import check_password_hash, generate_password_hash

from buses_iniciales import BUSES

TZ = ZoneInfo(os.environ.get("ZONA_HORARIA", "America/Santiago"))
def _limpiar_url(v):
    """Acepta también el snippet de Neon (psql 'postgresql://...') o la URL con comillas o espacios."""
    m = re.search(r"postgres(?:ql)?://[^\s'\"]+", v or "")
    return m.group(0) if m else (v or "").strip()


DATABASE_URL = _limpiar_url(os.environ.get("DATABASE_URL", "postgresql://localhost/flota"))
ADMIN_USUARIO = os.environ.get("ADMIN_USUARIO", "admin")
ADMIN_CLAVE = os.environ.get("ADMIN_CLAVE", "")

TERMINALES = ["Pie Andino", "Aguirre Luco", "Juanita", "Otro"]  # columna J "Terminal Real"
ESTADOS = ["Operativo", "Mantenimiento"]  # columna K "Operativo"
# Columnas de la hoja "Base"
COLUMNAS_BASE = ["Fecha", "Hora", "Encargado", "PLACA", "INTERNO", "TIPO BUS", "TIPO FLOTA", "MARCA CHASIS",
                 "TERMINAL OP", "Terminal Real", "Operativo", "Motivo mantenimiento", "Observacion",
                 "Último cambio", "Cambiado por"]
COLUMNAS_CAMBIOS = ["Fecha y hora", "Fecha reporte", "Tipo", "Usuario", "Nombre", "PLACA", "INTERNO", "TIPO BUS",
                    "TIPO FLOTA", "MARCA CHASIS", "TERMINAL OP", "Terminal Real", "Operativo", "Motivo mantenimiento",
                    "Observacion"]
MOTIVOS_INICIALES = ["Mecánica", "Eléctrica", "Carrocería", "Neumáticos", "Frenos", "Climatización", "Aire acondicionado",
                     "Revisión técnica", "Siniestro", "Validador / Torniquete", "Otro"]
CAMPOS = ("terminal_real", "operativo", "motivo", "observacion")

SCHEMA = """
CREATE TABLE IF NOT EXISTS buses (
    id SERIAL PRIMARY KEY,
    placa TEXT NOT NULL UNIQUE,
    interno TEXT NOT NULL UNIQUE,
    tipo_bus TEXT NOT NULL DEFAULT '',
    tipo_flota TEXT NOT NULL DEFAULT '',
    marca_chasis TEXT NOT NULL DEFAULT '',
    terminal_op TEXT NOT NULL DEFAULT '',
    ini_terminal TEXT NOT NULL DEFAULT 'Pie Andino',
    ini_estado TEXT NOT NULL DEFAULT 'Operativo',
    activo BOOLEAN NOT NULL DEFAULT TRUE
);
-- Un reporte por día
CREATE TABLE IF NOT EXISTS reportes (
    id SERIAL PRIMARY KEY,
    fecha DATE NOT NULL UNIQUE,
    hora TIME NOT NULL,
    encargado TEXT NOT NULL,
    registrado_por TEXT NOT NULL,
    modificado_por TEXT,
    modificado TIMESTAMPTZ,
    creado TIMESTAMPTZ NOT NULL DEFAULT now()
);
-- "Base": una fila por bus y por día (columnas A a L de la hoja Base)
CREATE TABLE IF NOT EXISTS base (
    id SERIAL PRIMARY KEY,
    reporte_id INT NOT NULL REFERENCES reportes(id) ON DELETE CASCADE,
    fecha DATE NOT NULL,
    hora TIME NOT NULL,
    encargado TEXT NOT NULL,
    placa TEXT NOT NULL,
    interno TEXT NOT NULL,
    tipo_bus TEXT NOT NULL,
    tipo_flota TEXT NOT NULL,
    marca_chasis TEXT NOT NULL,
    terminal_op TEXT NOT NULL,
    terminal_real TEXT NOT NULL CHECK (terminal_real IN ('Pie Andino','Aguirre Luco','Juanita','Otro')),
    operativo TEXT NOT NULL CHECK (operativo IN ('Operativo','Mantenimiento')),
    observacion TEXT NOT NULL DEFAULT '',
    UNIQUE (fecha, interno)
);
CREATE INDEX IF NOT EXISTS base_fecha ON base (fecha);
CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY,
    usuario TEXT NOT NULL UNIQUE,
    nombre TEXT NOT NULL,
    clave_hash TEXT NOT NULL,
    rol TEXT NOT NULL DEFAULT 'operador',
    activo BOOLEAN NOT NULL DEFAULT TRUE,
    creado TIMESTAMPTZ NOT NULL DEFAULT now()
);
ALTER TABLE usuarios DROP CONSTRAINT IF EXISTS usuarios_rol_check;
ALTER TABLE usuarios ADD CONSTRAINT usuarios_rol_check CHECK (rol IN ('admin','operador','invitado'));
ALTER TABLE base ADD COLUMN IF NOT EXISTS motivo TEXT NOT NULL DEFAULT '';
ALTER TABLE base ADD COLUMN IF NOT EXISTS cambio_ts TIMESTAMPTZ;
ALTER TABLE base ADD COLUMN IF NOT EXISTS cambio_por TEXT NOT NULL DEFAULT '';
-- Motivos de mantenimiento (lista desplegable con mantenedor)
CREATE TABLE IF NOT EXISTS motivos (
    id SERIAL PRIMARY KEY,
    nombre TEXT NOT NULL UNIQUE,
    activo BOOLEAN NOT NULL DEFAULT TRUE,
    orden INT NOT NULL DEFAULT 100
);
-- Historial: cada día parte con la foto completa ('inicial') y luego solo se agregan los cambios.
-- Cada fila guarda el estado completo del bus después del cambio, con quién y cuándo.
CREATE TABLE IF NOT EXISTS cambios (
    id BIGSERIAL PRIMARY KEY,
    ts TIMESTAMPTZ NOT NULL DEFAULT now(),
    fecha DATE NOT NULL,
    tipo TEXT NOT NULL CHECK (tipo IN ('inicial','cambio')),
    usuario TEXT NOT NULL,
    nombre TEXT NOT NULL DEFAULT '',
    interno TEXT NOT NULL,
    terminal_real TEXT NOT NULL,
    operativo TEXT NOT NULL,
    motivo TEXT NOT NULL DEFAULT '',
    observacion TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS cambios_interno_ts ON cambios (interno, ts DESC);
CREATE INDEX IF NOT EXISTS cambios_ts ON cambios (ts);
CREATE TABLE IF NOT EXISTS migraciones (nombre TEXT PRIMARY KEY, fecha TIMESTAMPTZ NOT NULL DEFAULT now());
-- Bitácora: una fila por cada acción (quién, cuándo, desde dónde, qué cambió y cómo estaba antes).
-- Solo se agregan filas: un disparador impide editarlas o borrarlas.
CREATE TABLE IF NOT EXISTS bitacora (
    id BIGSERIAL PRIMARY KEY,
    ts TIMESTAMPTZ NOT NULL DEFAULT now(),
    usuario TEXT NOT NULL,
    nombre TEXT NOT NULL DEFAULT '',
    rol TEXT NOT NULL DEFAULT '',
    origen TEXT NOT NULL,
    accion TEXT NOT NULL,
    fecha_reporte DATE,
    interno TEXT NOT NULL DEFAULT '',
    objeto TEXT NOT NULL DEFAULT '',
    campo TEXT NOT NULL DEFAULT '',
    antes TEXT NOT NULL DEFAULT '',
    despues TEXT NOT NULL DEFAULT '',
    detalle TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS bitacora_ts ON bitacora (ts);
CREATE INDEX IF NOT EXISTS bitacora_usuario_ts ON bitacora (usuario, ts);
CREATE OR REPLACE FUNCTION solo_agregar() RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    RAISE EXCEPTION 'La tabla % solo admite agregar registros: las correcciones se registran como un cambio nuevo', TG_TABLE_NAME;
END $$;
DROP TRIGGER IF EXISTS bitacora_solo_agregar ON bitacora;
CREATE TRIGGER bitacora_solo_agregar BEFORE UPDATE OR DELETE ON bitacora FOR EACH ROW EXECUTE FUNCTION solo_agregar();
DROP TRIGGER IF EXISTS bitacora_sin_vaciar ON bitacora;
CREATE TRIGGER bitacora_sin_vaciar BEFORE TRUNCATE ON bitacora FOR EACH STATEMENT EXECUTE FUNCTION solo_agregar();
"""

# El historial de estados (cambios) queda igual de protegido, después de las migraciones antiguas.
SCHEMA_PROTEGER_CAMBIOS = """
DROP TRIGGER IF EXISTS cambios_solo_agregar ON cambios;
CREATE TRIGGER cambios_solo_agregar BEFORE UPDATE OR DELETE ON cambios FOR EACH ROW EXECUTE FUNCTION solo_agregar();
DROP TRIGGER IF EXISTS cambios_sin_vaciar ON cambios;
CREATE TRIGGER cambios_sin_vaciar BEFORE TRUNCATE ON cambios FOR EACH STATEMENT EXECUTE FUNCTION solo_agregar();
"""

NOMBRE_CAMPO = {"terminal_real": "Terminal Real", "operativo": "Estado", "motivo": "Motivo mantenimiento",
                "observacion": "Observación"}
ORIGENES = {"arrastre": "Arrastre entre terminales", "doble_click": "Doble click", "detalle": "Ventana del bus",
            "pantalla": "Reporte diario"}
NOMBRE_ROL = {"admin": "Administrador", "operador": "Operador", "invitado": "Invitado"}
SISTEMA = {"usuario": "sistema", "nombre": "Sistema", "rol": "sistema"}


def registrar(cx, accion, origen, u=None, ts=None, **k):
    """Agrega una fila a la bitácora. k: fecha_reporte, interno, objeto, campo, antes, despues, detalle."""
    registrar_varias(cx, [dict(k, accion=accion, origen=origen)], u, ts)


def registrar_varias(cx, filas, u=None, ts=None):
    u = u or SISTEMA
    ts = ts or ahora()
    cx.cursor().executemany(
        """INSERT INTO bitacora (ts, usuario, nombre, rol, origen, accion, fecha_reporte, interno, objeto, campo, antes, despues, detalle)
           VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
        [(ts, u["usuario"], u.get("nombre", ""), u.get("rol", ""), f["origen"], f["accion"], f.get("fecha_reporte"),
          f.get("interno", ""), f.get("objeto", ""), f.get("campo", ""), str(f.get("antes", "") or ""),
          str(f.get("despues", "") or ""), f.get("detalle", "")) for f in filas],
    )

app = Flask(__name__, static_folder="static")
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("RENDER") is not None,
    PERMANENT_SESSION_LIFETIME=timedelta(days=7),
)
logging.basicConfig(level=logging.INFO)


def conectar():
    return psycopg.connect(
        DATABASE_URL, row_factory=dict_row, autocommit=False,
        connect_timeout=15, keepalives=1, keepalives_idle=30, keepalives_interval=10, keepalives_count=3,
    )


def ahora():
    return datetime.now(TZ)


def iniciar_bd():
    with conectar() as cx:
        cx.execute(SCHEMA)
        cx.cursor().executemany(
            """INSERT INTO buses (placa, interno, tipo_bus, tipo_flota, marca_chasis, terminal_op, ini_terminal, ini_estado)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING""",
            BUSES,
        )
        cx.cursor().executemany(
            "INSERT INTO motivos (nombre, orden) VALUES (%s,%s) ON CONFLICT DO NOTHING",
            [(m, i) for i, m in enumerate(MOTIVOS_INICIALES)],
        )
        # Reportes anteriores al historial: se cargan como foto 'inicial' a la hora en que se enviaron.
        cx.execute(
            """INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
               SELECT (b.fecha + b.hora) AT TIME ZONE %s, b.fecha, 'inicial', r.registrado_por, b.encargado, b.interno,
                      b.terminal_real, b.operativo, b.motivo, b.observacion
               FROM base b JOIN reportes r ON r.id = b.reporte_id
               WHERE NOT EXISTS (SELECT 1 FROM cambios c WHERE c.fecha = b.fecha)""",
            (str(TZ),),
        )
        # Una sola vez: observaciones escritas antes de existir el campo Motivo que son un motivo de la lista
        # (solo buses en Mantenimiento) pasan al campo Motivo.
        if cx.execute("INSERT INTO migraciones (nombre) VALUES ('obs_a_motivo') ON CONFLICT DO NOTHING RETURNING 1").fetchone():
            for tabla in ("base", "cambios"):
                n = cx.execute(
                    f"""UPDATE {tabla} t SET motivo = m.nombre, observacion = ''
                        FROM motivos m
                        WHERE t.motivo = '' AND t.operativo = 'Mantenimiento'
                          AND lower(trim(t.observacion)) = lower(m.nombre)"""
                ).rowcount
                app.logger.info("Migración obs_a_motivo: %s filas movidas en %s", n, tabla)
        # Una sola vez: lo mismo, pero sin importar el estado (el usuario pidió que el dato quede siempre en Motivo).
        if cx.execute("INSERT INTO migraciones (nombre) VALUES ('obs_a_motivo_todos') ON CONFLICT DO NOTHING RETURNING 1").fetchone():
            for tabla in ("base", "cambios"):
                filas = cx.execute(
                    f"""UPDATE {tabla} t SET motivo = m.nombre, observacion = ''
                        FROM motivos m
                        WHERE t.motivo = '' AND lower(trim(t.observacion)) = lower(m.nombre)
                        RETURNING t.fecha, t.interno, t.operativo, m.nombre"""
                ).fetchall()
                app.logger.info("Migración obs_a_motivo_todos: %s filas movidas en %s: %s", len(filas), tabla,
                                "; ".join(f"{f['fecha']} {f['interno']} {f['operativo']} -> {f['nombre']}" for f in filas[:40]))
            resto = cx.execute(
                """SELECT fecha, interno, operativo, observacion FROM base WHERE observacion <> ''
                   ORDER BY fecha DESC, interno LIMIT 40"""
            ).fetchall()
            app.logger.info("Observaciones que quedan: %s", "; ".join(
                f"{f['fecha']} {f['interno']} {f['operativo']}: {f['observacion']}" for f in resto) or "ninguna")
        # Una sola vez: dejar en la bitácora las correcciones que hizo obs_a_motivo_todos (15-11-2026),
        # que se aplicaron editando el historial antes de que existiera la bitácora.
        if cx.execute("INSERT INTO migraciones (nombre) VALUES ('bitacora_correcciones_obs') ON CONFLICT DO NOTHING RETURNING 1").fetchone():
            m = cx.execute("SELECT fecha FROM migraciones WHERE nombre='obs_a_motivo_todos'").fetchone()
            corregidas = cx.execute(
                f"""SELECT b.fecha, b.interno, b.motivo FROM base b JOIN motivos m ON m.nombre = b.motivo
                    WHERE b.operativo = 'Operativo' AND b.observacion = '' ORDER BY b.fecha, {ORDEN_INTERNO}"""
            ).fetchall() if m else []
            filas = []
            for f in corregidas:
                comun = dict(origen="Corrección de datos", accion="Corrección: observación pasada a motivo",
                             fecha_reporte=f["fecha"], interno=f["interno"],
                             detalle="Registrada después: se aplicó al historial antes de existir la bitácora")
                filas += [dict(comun, campo="Observación", antes=f["motivo"], despues=""),
                          dict(comun, campo="Motivo mantenimiento", antes="", despues=f["motivo"])]
            if filas:
                registrar_varias(cx, filas, ts=m["fecha"])
            app.logger.info("Bitácora: %s correcciones anteriores registradas", len(corregidas))
        cx.execute(SCHEMA_PROTEGER_CAMBIOS)
        hay_admin = cx.execute("SELECT 1 FROM usuarios WHERE rol='admin' AND activo").fetchone()
        if not hay_admin and ADMIN_CLAVE:
            cx.execute(
                """INSERT INTO usuarios (usuario, nombre, clave_hash, rol) VALUES (%s, 'Administrador', %s, 'admin')
                   ON CONFLICT (usuario) DO UPDATE SET clave_hash = EXCLUDED.clave_hash, rol='admin', activo=TRUE""",
                (ADMIN_USUARIO.lower(), generate_password_hash(ADMIN_CLAVE)),
            )
            registrar(cx, "Usuario administrador inicial creado", "Sistema", objeto=ADMIN_USUARIO.lower())
            app.logger.info("Administrador inicial '%s' creado", ADMIN_USUARIO)


def parse_fecha(s, defecto=None):
    try:
        return date.fromisoformat(s) if s else defecto
    except ValueError:
        return defecto


# ---------- Sesión ----------

def usuario_actual():
    return session.get("u")


def requiere_login(f):
    @wraps(f)
    def w(*a, **k):
        if not usuario_actual():
            return jsonify(error="Tu sesión terminó. Vuelve a iniciar sesión."), 401
        return f(*a, **k)
    return w


def requiere_admin(f):
    @wraps(f)
    def w(*a, **k):
        u = usuario_actual()
        if not u:
            return jsonify(error="Tu sesión terminó. Vuelve a iniciar sesión."), 401
        if u["rol"] != "admin":
            return jsonify(error="Solo un administrador puede hacer esto."), 403
        return f(*a, **k)
    return w


def requiere_escritura(f):
    """El perfil invitado puede usar todo, pero nada de lo que hace se guarda en la base."""
    @wraps(f)
    def w(*a, **k):
        u = usuario_actual()
        if not u:
            return jsonify(error="Tu sesión terminó. Vuelve a iniciar sesión."), 401
        if u["rol"] == "invitado":
            return jsonify(error="Perfil invitado: los cambios no se guardan en la base de datos."), 403
        return f(*a, **k)
    return w


def es_invitado():
    u = usuario_actual()
    return bool(u) and u["rol"] == "invitado"


def mapa_ficticio():
    """Interno y placa reales -> ficticios, para que el perfil invitado (demo pública) no vea la flota real.
    Se asignan por orden de interno, así que el orden y los grupos se mantienen; el mapeo es el mismo en toda la app."""
    if "mapa_ficticio" not in g:
        with conectar() as cx:
            buses = cx.execute(f"SELECT interno, placa FROM buses ORDER BY {ORDEN_INTERNO}").fetchall()
        m = {}
        for i, b in enumerate(buses):
            m[b["interno"]] = str(1001 + i)
            m[b["placa"]] = f"DM{i + 1:04d}"
        patron = re.compile(r"(?<![0-9A-Za-z])(" + "|".join(sorted(map(re.escape, m), key=len, reverse=True)) + r")(?![0-9A-Za-z])") if m else None
        g.mapa_ficticio = (m, patron)
    return g.mapa_ficticio


# Nombres reales de terminales -> ficticios para el perfil invitado ("Otro" se mantiene).
TERMINALES_FICTICIAS = {"pie andino": "Terminal Norte", "aguirre luco": "Terminal Centro", "juanita": "Terminal Sur"}
PATRON_TERMINALES = re.compile(r"\b(" + "|".join(TERMINALES_FICTICIAS) + r")\b", re.IGNORECASE)


def ocultar(v):
    """Reemplaza internos, placas y nombres de terminales reales por ficticios en cualquier valor (texto, lista o dict)."""
    if isinstance(v, str):
        m, patron = mapa_ficticio()
        v = patron.sub(lambda x: m[x.group(1)], v) if patron else v
        f = lambda x: TERMINALES_FICTICIAS[x.group(1).lower()]
        return PATRON_TERMINALES.sub(lambda x: f(x).upper() if x.group(1).isupper() else f(x), v)
    if isinstance(v, list):
        return [ocultar(x) for x in v]
    if isinstance(v, dict):
        return {k: ocultar(x) for k, x in v.items()}
    return v


def interno_real(valor):
    """Filtro escrito por el invitado (interno ficticio) -> interno real."""
    if not es_invitado():
        return valor
    m, _ = mapa_ficticio()
    inv = {f: r for r, f in m.items()}
    return inv.get(valor, valor)


@app.after_request
def internos_ficticios_para_invitado(resp):
    if resp.is_json and request.path.startswith("/api/") and es_invitado():
        resp.set_data(json.dumps(ocultar(resp.get_json()), ensure_ascii=False))
    return resp


@app.before_request
def verificar_sesion_vigente():
    """Cierra la sesión si el usuario fue desactivado o cambió de rol (se revisa cada minuto)."""
    u = usuario_actual()
    if not u or request.path.startswith("/static") or request.path == "/healthz":
        return
    if session.get("verificado", 0) > time.time() - 60:
        return
    with conectar() as cx:
        r = cx.execute("SELECT id, usuario, nombre, rol, activo FROM usuarios WHERE id=%s", (u["id"],)).fetchone()
    if not r or not r["activo"]:
        session.clear()
        return
    session["u"] = {"id": r["id"], "usuario": r["usuario"], "nombre": r["nombre"], "rol": r["rol"]}
    session["verificado"] = time.time()


@app.before_request
def solo_json_en_post():
    # Protección CSRF: los formularios de otros sitios no pueden enviar JSON con cookies.
    if request.method == "POST" and request.path.startswith("/api/") and not request.is_json:
        return jsonify(error="Solicitud no válida."), 400


def pagina(nombre, admin=False):
    u = usuario_actual()
    if not u:
        return redirect("/login")
    if admin and u["rol"] != "admin":
        return redirect("/")
    return send_from_directory(app.static_folder, nombre)


@app.get("/")
def index():
    return pagina("index.html")


@app.get("/dashboard")
def dashboard_page():
    return pagina("dashboard.html")


@app.get("/admin")
def admin_page():
    return pagina("admin.html", admin=True)


@app.get("/login")
def login_page():
    return send_from_directory(app.static_folder, "login.html")


@app.post("/api/login")
def api_login():
    d = request.get_json(silent=True) or {}
    usuario = (d.get("usuario") or "").strip().lower()
    clave = d.get("clave") or ""
    with conectar() as cx:
        r = cx.execute("SELECT * FROM usuarios WHERE lower(usuario)=%s AND activo", (usuario,)).fetchone()
    if not r or not check_password_hash(r["clave_hash"], clave):
        if r:
            with conectar() as cx:
                registrar(cx, "Ingreso fallido (contraseña incorrecta)", "Sesión", r)
        time.sleep(0.7)
        return jsonify(error="Usuario o contraseña incorrectos."), 401
    session.clear()
    session.permanent = True
    session["u"] = {"id": r["id"], "usuario": r["usuario"], "nombre": r["nombre"], "rol": r["rol"]}
    session["verificado"] = time.time()
    with conectar() as cx:
        registrar(cx, "Ingreso", "Sesión", r)
    return jsonify(ok=True)


@app.post("/api/logout")
def api_logout():
    u = usuario_actual()
    if u:
        with conectar() as cx:
            registrar(cx, "Salida", "Sesión", u)
    session.clear()
    return jsonify(ok=True)


@app.get("/api/yo")
@requiere_login
def api_yo():
    return jsonify(usuario_actual())


@app.post("/api/mi-clave")
@requiere_escritura
def api_mi_clave():
    d = request.get_json(silent=True) or {}
    actual, nueva = d.get("actual") or "", d.get("nueva") or ""
    if len(nueva) < 8:
        return jsonify(error="La nueva contraseña debe tener al menos 8 caracteres."), 400
    with conectar() as cx:
        r = cx.execute("SELECT clave_hash FROM usuarios WHERE id=%s", (usuario_actual()["id"],)).fetchone()
        if not check_password_hash(r["clave_hash"], actual):
            return jsonify(error="La contraseña actual no es correcta."), 400
        cx.execute("UPDATE usuarios SET clave_hash=%s WHERE id=%s",
                   (generate_password_hash(nueva), usuario_actual()["id"]))
        registrar(cx, "Cambió su propia contraseña", "Sesión", usuario_actual(), objeto=usuario_actual()["usuario"])
    return jsonify(ok=True)


# ---------- Reporte diario ----------

ORDEN_INTERNO = "CASE WHEN interno ~ '^[0-9]+$' THEN interno::int END, interno"


def fila_json(r):
    return {
        "placa": r["placa"], "interno": r["interno"], "tipo_bus": r["tipo_bus"], "tipo_flota": r["tipo_flota"],
        "marca_chasis": r["marca_chasis"], "terminal_op": r["terminal_op"],
        "terminal_real": r["terminal_real"], "operativo": r["operativo"], "motivo": r.get("motivo") or "",
        "observacion": r["observacion"], "cambio_por": r.get("cambio_por") or "",
        "cambio_ts": r["cambio_ts"].astimezone(TZ).strftime("%d-%m %H:%M") if r.get("cambio_ts") else "",
    }


def lista_motivos(cx):
    return [m["nombre"] for m in cx.execute("SELECT nombre FROM motivos WHERE activo ORDER BY orden, nombre").fetchall()]


def prellenado(cx, fecha):
    """Filas de un día sin reporte: estado del último reporte anterior (o el inicial del Excel)."""
    ant = cx.execute("SELECT id, fecha FROM reportes WHERE fecha < %s ORDER BY fecha DESC LIMIT 1", (fecha,)).fetchone()
    previas = {}
    if ant:
        for f in cx.execute("SELECT * FROM base WHERE reporte_id=%s", (ant["id"],)).fetchall():
            previas[f["interno"]] = f
    filas = []
    for b in cx.execute(f"SELECT * FROM buses WHERE activo ORDER BY {ORDEN_INTERNO}").fetchall():
        p = previas.get(b["interno"])
        filas.append({
            "placa": b["placa"], "interno": b["interno"], "tipo_bus": b["tipo_bus"], "tipo_flota": b["tipo_flota"],
            "marca_chasis": b["marca_chasis"], "terminal_op": b["terminal_op"],
            "terminal_real": p["terminal_real"] if p else b["ini_terminal"],
            "operativo": p["operativo"] if p else b["ini_estado"],
            "motivo": (p.get("motivo") or "") if p else "",
            "observacion": p["observacion"] if p else "",
            "cambio_por": "", "cambio_ts": "",
        })
    return ant, filas


@app.get("/api/reporte")
@requiere_login
def api_reporte():
    """Estado de una fecha (hoy por defecto). Si aún no hay reporte, viene prellenado con el día anterior."""
    hoy = ahora().date()
    fecha = parse_fecha(request.args.get("fecha"), hoy)
    if fecha > hoy:
        return jsonify(error="No se puede reportar una fecha futura."), 400
    with conectar() as cx:
        motivos = lista_motivos(cx)
        rep = cx.execute("SELECT * FROM reportes WHERE fecha=%s", (fecha,)).fetchone()
        if rep:
            filas = cx.execute(f"SELECT * FROM base WHERE reporte_id=%s ORDER BY {ORDEN_INTERNO}", (rep["id"],)).fetchall()
            ult = cx.execute("SELECT max(ts) AS t FROM cambios WHERE fecha=%s", (fecha,)).fetchone()["t"]
            return jsonify(
                fecha=fecha.isoformat(), existe=True,
                reporte={"hora": rep["hora"].strftime("%H:%M"), "encargado": rep["encargado"],
                         "registrado_por": rep["registrado_por"], "modificado_por": rep["modificado_por"],
                         "ultimo_cambio": ult.astimezone(TZ).strftime("%H:%M") if ult else None},
                filas=[fila_json(f) for f in filas], terminales=TERMINALES, estados=ESTADOS, motivos=motivos,
            )
        ant, filas = prellenado(cx, fecha)
    return jsonify(
        fecha=fecha.isoformat(), existe=False, prellenado_desde=ant["fecha"].isoformat() if ant else None,
        filas=filas, terminales=TERMINALES, estados=ESTADOS, motivos=motivos,
    )


@app.post("/api/reporte")
@requiere_escritura
def api_guardar_reporte():
    """Guarda solo los buses que cambiaron. El primer guardado del día crea el reporte con la foto completa."""
    d = request.get_json(silent=True) or {}
    u = usuario_actual()
    n = ahora()
    fecha = parse_fecha(d.get("fecha"), n.date())
    if fecha > n.date():
        return jsonify(error="No se puede reportar una fecha futura."), 400
    encargado = (d.get("encargado") or u["nombre"]).strip()[:120]
    entradas = {}
    for f in d.get("filas") or []:
        i = str(f.get("interno"))
        if f.get("terminal_real") not in TERMINALES:
            return jsonify(error=f"Terminal Real no válida en el bus {i}."), 400
        if f.get("operativo") not in ESTADOS:
            return jsonify(error=f"Estado no válido en el bus {i}."), 400
        entradas[i] = {
            "terminal_real": f["terminal_real"], "operativo": f["operativo"],
            "motivo": (f.get("motivo") or "").strip()[:120] if f["operativo"] == "Mantenimiento" else "",
            "observacion": (f.get("observacion") or "").strip()[:500],
            "origen": ORIGENES.get(f.get("origen"), ORIGENES["pantalla"]),
        }
    with conectar() as cx:
        cx.execute("SELECT pg_advisory_xact_lock(hashtext('reporte' || %s))", (fecha.isoformat(),))
        rep = cx.execute("SELECT * FROM reportes WHERE fecha=%s", (fecha,)).fetchone()
        nuevo = rep is None
        if nuevo:
            hora = n.time().replace(second=0, microsecond=0)
            rid = cx.execute(
                "INSERT INTO reportes (fecha, hora, encargado, registrado_por) VALUES (%s,%s,%s,%s) RETURNING id",
                (fecha, hora, encargado, u["usuario"]),
            ).fetchone()["id"]
            _, filas = prellenado(cx, fecha)
            cx.cursor().executemany(
                """INSERT INTO base (reporte_id, fecha, hora, encargado, placa, interno, tipo_bus, tipo_flota,
                                     marca_chasis, terminal_op, terminal_real, operativo, motivo, observacion)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                [(rid, fecha, hora, encargado, f["placa"], f["interno"], f["tipo_bus"], f["tipo_flota"],
                  f["marca_chasis"], f["terminal_op"], f["terminal_real"], f["operativo"], f["motivo"],
                  f["observacion"]) for f in filas],
            )
            cx.cursor().executemany(
                """INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
                   VALUES (%s,%s,'inicial',%s,%s,%s,%s,%s,%s,%s)""",
                [(n, fecha, u["usuario"], u["nombre"], f["interno"], f["terminal_real"], f["operativo"], f["motivo"],
                  f["observacion"]) for f in filas],
            )
            ant = cx.execute("SELECT fecha FROM reportes WHERE fecha < %s ORDER BY fecha DESC LIMIT 1", (fecha,)).fetchone()
            registrar(cx, "Reporte del día creado", ORIGENES["pantalla"], u, n, fecha_reporte=fecha,
                      detalle=f"{len(filas)} buses, " + (f"tomados del reporte del {ant['fecha']:%d-%m-%Y}" if ant
                                                          else "estado inicial del Excel") + f". Encargado: {encargado}")
        else:
            rid = rep["id"]
            if encargado != rep["encargado"]:
                registrar(cx, "Cambio de encargado del reporte", ORIGENES["pantalla"], u, n, fecha_reporte=fecha,
                          campo="Encargado", antes=rep["encargado"], despues=encargado)
            cx.execute("UPDATE reportes SET encargado=%s, modificado_por=%s, modificado=now() WHERE id=%s",
                       (encargado, u["usuario"], rid))
        actuales = {f["interno"]: f for f in cx.execute("SELECT * FROM base WHERE reporte_id=%s", (rid,)).fetchall()}
        faltan = [i for i in entradas if i not in actuales]
        if faltan:
            return jsonify(error="Buses que no están en el reporte: " + ", ".join(faltan)), 400
        cambiados = [i for i, e in entradas.items() if any(e[c] != (actuales[i][c] or "") for c in CAMPOS)]
        registrar_varias(cx, [
            dict(accion="Cambio de bus", origen=entradas[i]["origen"], fecha_reporte=fecha, interno=i,
                 campo=NOMBRE_CAMPO[c], antes=actuales[i][c] or "", despues=entradas[i][c])
            for i in cambiados for c in CAMPOS if entradas[i][c] != (actuales[i][c] or "")], u, n)
        for i in cambiados:
            e = entradas[i]
            cx.execute(
                """UPDATE base SET terminal_real=%s, operativo=%s, motivo=%s, observacion=%s, cambio_ts=%s, cambio_por=%s
                   WHERE reporte_id=%s AND interno=%s""",
                (e["terminal_real"], e["operativo"], e["motivo"], e["observacion"], n, u["nombre"], rid, i),
            )
            cx.execute(
                """INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
                   VALUES (%s,%s,'cambio',%s,%s,%s,%s,%s,%s,%s)""",
                (n, fecha, u["usuario"], u["nombre"], i, e["terminal_real"], e["operativo"], e["motivo"], e["observacion"]),
            )
    return jsonify(ok=True, creado=nuevo, cambios=len(cambiados), hora=n.strftime("%H:%M"))


@app.get("/api/reportes")
@requiere_login
def api_lista_reportes():
    with conectar() as cx:
        rs = cx.execute(
            """SELECT r.fecha, r.hora, r.encargado, count(*) FILTER (WHERE b.operativo='Operativo') AS op, count(b.id) AS total
               FROM reportes r LEFT JOIN base b ON b.reporte_id=r.id GROUP BY r.id ORDER BY r.fecha DESC LIMIT 60"""
        ).fetchall()
    return jsonify([{"fecha": r["fecha"].isoformat(), "hora": r["hora"].strftime("%H:%M"),
                     "encargado": r["encargado"], "operativos": r["op"], "total": r["total"]} for r in rs])


# ---------- Dashboard y descarga ----------

@app.get("/api/dashboard")
@requiere_login
def api_dashboard():
    """Filas compactas [fecha, interno, tipo_bus, marca, terminal_real, estado, observación] del rango pedido."""
    hasta = parse_fecha(request.args.get("hasta"), ahora().date())
    desde = parse_fecha(request.args.get("desde"), hasta - timedelta(days=29))
    with conectar() as cx:
        filas = estado_diario(cx, desde, hasta)
    return jsonify(
        desde=desde.isoformat(), hasta=hasta.isoformat(), terminales=TERMINALES,
        filas=[[f["fecha"].isoformat(), f["interno"], f["placa"], f["tipo_bus"], f["marca_chasis"],
                f["terminal_real"], f["operativo"], f["observacion"], f["motivo"]] for f in filas],
    )


def estado_diario(cx, desde, hasta):
    """Estado de cada bus al cierre de cada día, armado desde el historial (último registro de cada bus en ese reporte)."""
    return cx.execute(
        """SELECT x.*, r.hora, r.encargado, b.placa, b.tipo_bus, b.tipo_flota, b.marca_chasis, b.terminal_op,
                  ult.ts AS cambio_ts, ult.nombre AS cambio_por
           FROM (SELECT DISTINCT ON (fecha, interno) fecha, interno, terminal_real, operativo, motivo, observacion
                 FROM cambios WHERE fecha BETWEEN %s AND %s ORDER BY fecha, interno, ts DESC, id DESC) x
           JOIN buses b ON b.interno = x.interno
           JOIN reportes r ON r.fecha = x.fecha
           LEFT JOIN LATERAL (SELECT c.ts, c.nombre FROM cambios c
                              WHERE c.fecha = x.fecha AND c.interno = x.interno AND c.tipo = 'cambio'
                              ORDER BY c.ts DESC, c.id DESC LIMIT 1) ult ON TRUE
           ORDER BY x.fecha, CASE WHEN x.interno ~ '^[0-9]+$' THEN x.interno::int END, x.interno""", (desde, hasta)
    ).fetchall()


def filas_descarga():
    desde = parse_fecha(request.args.get("desde"), date(1900, 1, 1))
    hasta = parse_fecha(request.args.get("hasta"), date(2999, 12, 31))
    with conectar() as cx:
        filas = estado_diario(cx, desde, hasta)
    return [[f["fecha"], f["hora"].strftime("%H:%M"), f["encargado"], f["placa"], f["interno"], f["tipo_bus"],
             f["tipo_flota"], f["marca_chasis"], f["terminal_op"], f["terminal_real"], f["operativo"], f["motivo"],
             f["observacion"], f["cambio_ts"].astimezone(TZ).strftime("%d-%m-%Y %H:%M") if f["cambio_ts"] else "",
             f["cambio_por"]] for f in filas]


def momento_param():
    """?momento=AAAA-MM-DDTHH:MM en hora de Chile (por defecto, ahora)."""
    try:
        return datetime.fromisoformat(request.args.get("momento", "")).replace(tzinfo=TZ)
    except ValueError:
        return ahora()


def filas_estado(momento):
    """Estado de cada bus en un instante: el último registro hasta ese momento del reporte más reciente
    (una corrección hecha hoy a un día anterior no cambia el estado de hoy)."""
    with conectar() as cx:
        filas = cx.execute(
            f"""SELECT * FROM (
                  SELECT DISTINCT ON (c.interno) c.*, b.placa, b.tipo_bus, b.tipo_flota, b.marca_chasis, b.terminal_op
                  FROM cambios c JOIN buses b ON b.interno = c.interno
                  WHERE c.ts <= %s AND c.fecha <= %s ORDER BY c.interno, c.fecha DESC, c.ts DESC, c.id DESC) x
                ORDER BY {ORDEN_INTERNO}""", (momento + timedelta(seconds=59, microseconds=999999), momento.date())
        ).fetchall()
    return [[momento.strftime("%d-%m-%Y %H:%M"), f["placa"], f["interno"], f["tipo_bus"], f["tipo_flota"],
             f["marca_chasis"], f["terminal_op"], f["terminal_real"], f["operativo"], f["motivo"], f["observacion"],
             f["ts"].astimezone(TZ).strftime("%d-%m-%Y %H:%M"), f["nombre"]] for f in filas]


COLUMNAS_ESTADO = ["Estado al", "PLACA", "INTERNO", "TIPO BUS", "TIPO FLOTA", "MARCA CHASIS", "TERMINAL OP",
                   "Terminal Real", "Operativo", "Motivo mantenimiento", "Observacion", "Último cambio", "Cambiado por"]


def filas_cambios():
    desde = parse_fecha(request.args.get("desde"), date(1900, 1, 1))
    hasta = parse_fecha(request.args.get("hasta"), date(2999, 12, 31))
    with conectar() as cx:
        filas = cx.execute(
            """SELECT c.*, b.placa, b.tipo_bus, b.tipo_flota, b.marca_chasis, b.terminal_op
               FROM cambios c JOIN buses b ON b.interno = c.interno
               WHERE c.fecha BETWEEN %s AND %s ORDER BY c.ts, c.id""", (desde, hasta)
        ).fetchall()
    return [[f["ts"].astimezone(TZ).strftime("%d-%m-%Y %H:%M:%S"), f["fecha"].strftime("%d-%m-%Y"),
             "Foto inicial" if f["tipo"] == "inicial" else "Cambio", f["usuario"], f["nombre"], f["placa"], f["interno"],
             f["tipo_bus"], f["tipo_flota"], f["marca_chasis"], f["terminal_op"], f["terminal_real"], f["operativo"],
             f["motivo"], f["observacion"]] for f in filas]


COLUMNAS_BITACORA = ["Fecha y hora", "Usuario", "Nombre", "Rol", "Origen", "Acción", "Fecha reporte", "INTERNO",
                     "Objeto", "Campo", "Antes", "Después", "Detalle"]


def consulta_bitacora(cx, limite=None):
    """Bitácora filtrada por ?desde, ?hasta (días, hora de Chile), ?usuario, ?interno y ?tipo (buses|admin|sesion|descargas)."""
    a = request.args
    desde = parse_fecha(a.get("desde"), date(1900, 1, 1))
    hasta = parse_fecha(a.get("hasta"), date(2999, 12, 31))
    cond, par = ["(ts AT TIME ZONE %s)::date BETWEEN %s AND %s"], [str(TZ), desde, hasta]
    if a.get("usuario"):
        cond.append("usuario = %s"); par.append(a["usuario"])
    if a.get("interno"):
        cond.append("interno = %s"); par.append(interno_real(a["interno"].strip().upper()))
    tipo = {"buses": "(accion IN ('Cambio de bus','Reporte del día creado','Cambio de encargado del reporte') OR accion LIKE 'Corrección%%')",
            "admin": "origen = 'Administración'", "sesion": "origen = 'Sesión'", "descargas": "origen = 'Descarga'"}.get(a.get("tipo"))
    if tipo:
        cond.append(tipo)
    sql = f"SELECT * FROM bitacora WHERE {' AND '.join(cond)} ORDER BY ts DESC, id DESC"
    if limite:
        sql += f" LIMIT {int(limite)}"
    return cx.execute(sql, par).fetchall()


def fila_bitacora(f):
    return [f["ts"].astimezone(TZ).strftime("%d-%m-%Y %H:%M:%S"), f["usuario"], f["nombre"],
            NOMBRE_ROL.get(f["rol"], f["rol"]), f["origen"], f["accion"],
            f["fecha_reporte"].strftime("%d-%m-%Y") if f["fecha_reporte"] else "", f["interno"], f["objeto"],
            f["campo"], f["antes"], f["despues"], f["detalle"]]


@app.get("/api/bitacora")
@requiere_login
def api_bitacora():
    """Últimas acciones (hasta 300) y resumen por usuario del período, para el dashboard."""
    with conectar() as cx:
        filas = consulta_bitacora(cx)
        usuarios = cx.execute("SELECT DISTINCT usuario, nombre FROM bitacora ORDER BY nombre").fetchall()
    resumen = {}
    for f in filas:
        r = resumen.setdefault(f["usuario"], {"usuario": f["usuario"], "nombre": f["nombre"], "acciones": 0,
                                              "cambios_bus": 0, "buses": set(), "ultima": f["ts"]})
        r["acciones"] += 1
        if f["accion"] == "Cambio de bus":
            r["cambios_bus"] += 1
            r["buses"].add(f["interno"])
    res = sorted(({**r, "buses": len(r["buses"]), "ultima": r["ultima"].astimezone(TZ).strftime("%d-%m-%Y %H:%M")}
                  for r in resumen.values()), key=lambda r: -r["acciones"])
    return jsonify(total=len(filas), filas=[fila_bitacora(f) for f in filas[:300]], resumen=res,
                   usuarios=[{"usuario": u["usuario"], "nombre": u["nombre"]} for u in usuarios])


@app.get("/api/bitacora.<fmt>")
@requiere_login
def api_bitacora_descarga(fmt):
    with conectar() as cx:
        filas = [fila_bitacora(f) for f in consulta_bitacora(cx)]
        registrar_descarga(cx, "Bitácora de acciones", fmt)
    return (como_xlsx(COLUMNAS_BITACORA, filas, "bitacora_flota", "Bitácora") if fmt == "xlsx"
            else como_csv(COLUMNAS_BITACORA, filas, "bitacora_flota"))


def registrar_descarga(cx, que, fmt):
    filtros = ", ".join(f"{k}={v}" for k, v in request.args.items() if v)
    registrar(cx, "Descarga: " + que, "Descarga", usuario_actual(), detalle=f"{fmt}" + (f" ({filtros})" if filtros else ""))


def como_csv(columnas, filas, nombre):
    if es_invitado():
        filas = [[ocultar(v) for v in f] for f in filas]
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(columnas)
    for f in filas:
        w.writerow([v.strftime("%d-%m-%Y") if isinstance(v, date) else v for v in f])
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={nombre}.csv"})


def como_xlsx(columnas, filas, nombre, hoja):
    if es_invitado():
        filas = [[ocultar(v) for v in f] for f in filas]
    wb = Workbook()
    ws = wb.active
    ws.title = hoja
    ws.append(columnas)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F5F8B")
    for f in filas:
        ws.append(f)
        for k, v in enumerate(f, 1):
            if isinstance(v, date):
                ws.cell(ws.max_row, k).number_format = "DD-MM-YYYY"
    for k, col in enumerate(columnas, 1):
        ws.column_dimensions[ws.cell(1, k).column_letter].width = max(10, min(40, len(col) + 4))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    return Response(buf.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename={nombre}.xlsx"})


@app.get("/api/descarga.<fmt>")
@requiere_login
def api_descarga(fmt):
    filas = filas_descarga()
    with conectar() as cx:
        registrar_descarga(cx, "Base por día", fmt)
    return como_xlsx(COLUMNAS_BASE, filas, "base_flota", "Base") if fmt == "xlsx" else como_csv(COLUMNAS_BASE, filas, "base_flota")


@app.get("/api/estado.<fmt>")
@requiere_login
def api_estado(fmt):
    m = momento_param()
    filas = filas_estado(m)
    with conectar() as cx:
        registrar_descarga(cx, "Estado a una fecha y hora", fmt)
    nombre = "estado_flota_" + m.strftime("%Y%m%d_%H%M")
    return como_xlsx(COLUMNAS_ESTADO, filas, nombre, "Estado") if fmt == "xlsx" else como_csv(COLUMNAS_ESTADO, filas, nombre)


@app.get("/api/cambios.<fmt>")
@requiere_login
def api_cambios(fmt):
    filas = filas_cambios()
    with conectar() as cx:
        registrar_descarga(cx, "Historial de estados", fmt)
    return como_xlsx(COLUMNAS_CAMBIOS, filas, "historial_cambios", "Cambios") if fmt == "xlsx" else como_csv(COLUMNAS_CAMBIOS, filas, "historial_cambios")


# ---------- Administración ----------

@app.get("/api/admin/datos")
@requiere_admin
def api_admin_datos():
    with conectar() as cx:
        us = cx.execute("SELECT id, usuario, nombre, rol, activo FROM usuarios ORDER BY activo DESC, nombre").fetchall()
        bs = cx.execute(
            """SELECT id, placa, interno, tipo_bus, tipo_flota, marca_chasis, terminal_op, activo FROM buses
               ORDER BY activo DESC, CASE WHEN interno ~ '^[0-9]+$' THEN interno::int END, interno"""
        ).fetchall()
        ms = cx.execute("SELECT id, nombre, activo FROM motivos ORDER BY activo DESC, orden, nombre").fetchall()
    return jsonify(usuarios=us, buses=bs, motivos=ms, yo=usuario_actual())


@app.post("/api/admin/usuarios")
@requiere_admin
def api_admin_crear_usuario():
    d = request.get_json(silent=True) or {}
    usuario = (d.get("usuario") or "").strip().lower()
    nombre = (d.get("nombre") or "").strip()
    rol = d.get("rol") if d.get("rol") in ("admin", "operador", "invitado") else "operador"
    clave = d.get("clave") or ""
    if not usuario or not nombre:
        return jsonify(error="Completa el usuario y el nombre."), 400
    if len(clave) < 8:
        return jsonify(error="La contraseña debe tener al menos 8 caracteres."), 400
    with conectar() as cx:
        if cx.execute("SELECT 1 FROM usuarios WHERE lower(usuario)=%s", (usuario,)).fetchone():
            return jsonify(error="Ese nombre de usuario ya existe."), 409
        cx.execute("INSERT INTO usuarios (usuario, nombre, clave_hash, rol) VALUES (%s,%s,%s,%s)",
                   (usuario, nombre, generate_password_hash(clave), rol))
        registrar(cx, "Usuario creado", "Administración", usuario_actual(), objeto=usuario,
                  detalle=f"{nombre}, rol {NOMBRE_ROL.get(rol, rol)}")
    return jsonify(ok=True)


@app.post("/api/admin/usuarios/<int:uid>")
@requiere_admin
def api_admin_editar_usuario(uid):
    d = request.get_json(silent=True) or {}
    with conectar() as cx:
        u = cx.execute("SELECT * FROM usuarios WHERE id=%s FOR UPDATE", (uid,)).fetchone()
        if not u:
            return jsonify(error="El usuario no existe."), 404
        rol = d.get("rol", u["rol"])
        activo = bool(d.get("activo", u["activo"]))
        if rol not in ("admin", "operador", "invitado"):
            return jsonify(error="Rol no válido."), 400
        if u["rol"] == "admin" and (rol != "admin" or not activo):
            otros = cx.execute("SELECT count(*) AS n FROM usuarios WHERE rol='admin' AND activo AND id<>%s",
                               (uid,)).fetchone()
            if otros["n"] == 0:
                return jsonify(error="Debe quedar al menos un administrador activo."), 409
        if d.get("clave") and len(d["clave"]) < 8:
            return jsonify(error="La contraseña debe tener al menos 8 caracteres."), 400
        cx.execute("UPDATE usuarios SET rol=%s, activo=%s WHERE id=%s", (rol, activo, uid))
        filas = []
        if rol != u["rol"]:
            filas.append(dict(accion="Cambio de rol", campo="Rol", antes=NOMBRE_ROL.get(u["rol"], u["rol"]),
                              despues=NOMBRE_ROL.get(rol, rol)))
        if activo != u["activo"]:
            filas.append(dict(accion="Usuario activado" if activo else "Usuario desactivado", campo="Estado",
                              antes="Activo" if u["activo"] else "Inactivo", despues="Activo" if activo else "Inactivo"))
        if d.get("clave"):
            cx.execute("UPDATE usuarios SET clave_hash=%s WHERE id=%s", (generate_password_hash(d["clave"]), uid))
            filas.append(dict(accion="Contraseña cambiada por un administrador", campo="Contraseña"))
        registrar_varias(cx, [dict(f, origen="Administración", objeto=u["usuario"]) for f in filas], usuario_actual())
    return jsonify(ok=True)


@app.post("/api/admin/buses")
@requiere_admin
def api_admin_buses():
    """Agrega un bus a la nómina, o activa/desactiva uno existente ({id, activo})."""
    d = request.get_json(silent=True) or {}
    with conectar() as cx:
        if d.get("id"):
            b = cx.execute("SELECT * FROM buses WHERE id=%s", (d["id"],)).fetchone()
            if not b:
                return jsonify(error="El bus no existe."), 404
            activo = bool(d.get("activo"))
            cx.execute("UPDATE buses SET activo=%s WHERE id=%s", (activo, d["id"]))
            if activo != b["activo"]:
                registrar(cx, "Bus activado en la nómina" if activo else "Bus desactivado de la nómina", "Administración",
                          usuario_actual(), interno=b["interno"], objeto=b["placa"], campo="Estado en nómina",
                          antes="Activo" if b["activo"] else "Inactivo", despues="Activo" if activo else "Inactivo")
            return jsonify(ok=True)
        campos = {k: (d.get(k) or "").strip().upper() for k in
                  ("placa", "interno", "tipo_bus", "tipo_flota", "marca_chasis", "terminal_op")}
        if not campos["placa"] or not campos["interno"]:
            return jsonify(error="La placa y el número interno son obligatorios."), 400
        if cx.execute("SELECT 1 FROM buses WHERE placa=%s OR interno=%s",
                      (campos["placa"], campos["interno"])).fetchone():
            return jsonify(error="Ya existe un bus con esa placa o ese interno."), 409
        cx.execute(
            """INSERT INTO buses (placa, interno, tipo_bus, tipo_flota, marca_chasis, terminal_op)
               VALUES (%(placa)s,%(interno)s,%(tipo_bus)s,%(tipo_flota)s,%(marca_chasis)s,%(terminal_op)s)""", campos)
        registrar(cx, "Bus agregado a la nómina", "Administración", usuario_actual(), interno=campos["interno"],
                  objeto=campos["placa"], detalle=", ".join(v for v in (campos["tipo_bus"], campos["tipo_flota"],
                                                                         campos["marca_chasis"], campos["terminal_op"]) if v))
    return jsonify(ok=True)


@app.post("/api/admin/motivos")
@requiere_admin
def api_admin_motivos():
    """Agrega un motivo de mantenimiento, o activa/desactiva/renombra uno existente ({id, activo, nombre})."""
    d = request.get_json(silent=True) or {}
    nombre = (d.get("nombre") or "").strip()[:120]
    with conectar() as cx:
        if d.get("id"):
            m = cx.execute("SELECT * FROM motivos WHERE id=%s", (d["id"],)).fetchone()
            if not m:
                return jsonify(error="El motivo no existe."), 404
            if nombre and nombre != m["nombre"]:
                if cx.execute("SELECT 1 FROM motivos WHERE nombre=%s AND id<>%s", (nombre, d["id"])).fetchone():
                    return jsonify(error="Ya existe un motivo con ese nombre."), 409
                cx.execute("UPDATE motivos SET nombre=%s WHERE id=%s", (nombre, d["id"]))
                registrar(cx, "Motivo renombrado", "Administración", usuario_actual(), objeto=nombre, campo="Nombre",
                          antes=m["nombre"], despues=nombre)
            if "activo" in d and bool(d["activo"]) != m["activo"]:
                cx.execute("UPDATE motivos SET activo=%s WHERE id=%s", (bool(d["activo"]), d["id"]))
                registrar(cx, "Motivo activado" if d["activo"] else "Motivo desactivado", "Administración",
                          usuario_actual(), objeto=nombre or m["nombre"], campo="Estado",
                          antes="Activo" if m["activo"] else "Inactivo", despues="Activo" if d["activo"] else "Inactivo")
            return jsonify(ok=True)
        if not nombre:
            return jsonify(error="Escribe el nombre del motivo."), 400
        previo = cx.execute("SELECT activo FROM motivos WHERE nombre=%s", (nombre,)).fetchone()
        cx.execute("INSERT INTO motivos (nombre) VALUES (%s) ON CONFLICT (nombre) DO UPDATE SET activo=TRUE", (nombre,))
        if not previo or not previo["activo"]:
            registrar(cx, "Motivo agregado" if not previo else "Motivo activado", "Administración", usuario_actual(),
                      objeto=nombre)
    return jsonify(ok=True)


@app.errorhandler(Exception)
def error(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return jsonify(error=e.description), e.code
    app.logger.exception("Error en %s", request.path)
    return jsonify(error="Error del servidor: " + type(e).__name__), 500


@app.get("/healthz")
def healthz():
    return "ok"


iniciar_bd()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
