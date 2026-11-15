"""Control de Flota: un reporte diario de la distribución de buses del depósito."""
import csv
import io
import logging
import os
import re
import secrets
import time
from datetime import date, datetime, timedelta
from functools import wraps
from zoneinfo import ZoneInfo

import psycopg
from flask import Flask, Response, jsonify, redirect, request, send_from_directory, session
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
"""

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
        hay_admin = cx.execute("SELECT 1 FROM usuarios WHERE rol='admin' AND activo").fetchone()
        if not hay_admin and ADMIN_CLAVE:
            cx.execute(
                """INSERT INTO usuarios (usuario, nombre, clave_hash, rol) VALUES (%s, 'Administrador', %s, 'admin')
                   ON CONFLICT (usuario) DO UPDATE SET clave_hash = EXCLUDED.clave_hash, rol='admin', activo=TRUE""",
                (ADMIN_USUARIO.lower(), generate_password_hash(ADMIN_CLAVE)),
            )
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
        time.sleep(0.7)
        return jsonify(error="Usuario o contraseña incorrectos."), 401
    session.clear()
    session.permanent = True
    session["u"] = {"id": r["id"], "usuario": r["usuario"], "nombre": r["nombre"], "rol": r["rol"]}
    session["verificado"] = time.time()
    return jsonify(ok=True)


@app.post("/api/logout")
def api_logout():
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
        else:
            rid = rep["id"]
            cx.execute("UPDATE reportes SET encargado=%s, modificado_por=%s, modificado=now() WHERE id=%s",
                       (encargado, u["usuario"], rid))
        actuales = {f["interno"]: f for f in cx.execute("SELECT * FROM base WHERE reporte_id=%s", (rid,)).fetchall()}
        faltan = [i for i in entradas if i not in actuales]
        if faltan:
            return jsonify(error="Buses que no están en el reporte: " + ", ".join(faltan)), 400
        cambiados = [i for i, e in entradas.items() if any(e[c] != (actuales[i][c] or "") for c in CAMPOS)]
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
        filas = cx.execute(
            """SELECT fecha, interno, placa, tipo_bus, marca_chasis, terminal_real, operativo, observacion, motivo FROM base
               WHERE fecha BETWEEN %s AND %s ORDER BY fecha""", (desde, hasta)
        ).fetchall()
    return jsonify(
        desde=desde.isoformat(), hasta=hasta.isoformat(), terminales=TERMINALES,
        filas=[[f["fecha"].isoformat(), f["interno"], f["placa"], f["tipo_bus"], f["marca_chasis"],
                f["terminal_real"], f["operativo"], f["observacion"], f["motivo"]] for f in filas],
    )


def filas_descarga():
    desde = parse_fecha(request.args.get("desde"), date(1900, 1, 1))
    hasta = parse_fecha(request.args.get("hasta"), date(2999, 12, 31))
    with conectar() as cx:
        filas = cx.execute(
            f"SELECT * FROM base WHERE fecha BETWEEN %s AND %s ORDER BY fecha, {ORDEN_INTERNO}", (desde, hasta)
        ).fetchall()
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
    """Estado de cada bus en un instante: el último registro del historial hasta ese momento."""
    with conectar() as cx:
        filas = cx.execute(
            f"""SELECT * FROM (
                  SELECT DISTINCT ON (c.interno) c.*, b.placa, b.tipo_bus, b.tipo_flota, b.marca_chasis, b.terminal_op
                  FROM cambios c JOIN buses b ON b.interno = c.interno
                  WHERE c.ts <= %s ORDER BY c.interno, c.ts DESC, c.id DESC) x
                ORDER BY {ORDEN_INTERNO}""", (momento + timedelta(seconds=59, microseconds=999999),)
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


def como_csv(columnas, filas, nombre):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(columnas)
    for f in filas:
        w.writerow([v.strftime("%d-%m-%Y") if isinstance(v, date) else v for v in f])
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={nombre}.csv"})


def como_xlsx(columnas, filas, nombre, hoja):
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
    return como_xlsx(COLUMNAS_BASE, filas, "base_flota", "Base") if fmt == "xlsx" else como_csv(COLUMNAS_BASE, filas, "base_flota")


@app.get("/api/estado.<fmt>")
@requiere_login
def api_estado(fmt):
    m = momento_param()
    filas = filas_estado(m)
    nombre = "estado_flota_" + m.strftime("%Y%m%d_%H%M")
    return como_xlsx(COLUMNAS_ESTADO, filas, nombre, "Estado") if fmt == "xlsx" else como_csv(COLUMNAS_ESTADO, filas, nombre)


@app.get("/api/cambios.<fmt>")
@requiere_login
def api_cambios(fmt):
    filas = filas_cambios()
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
        if d.get("clave"):
            cx.execute("UPDATE usuarios SET clave_hash=%s WHERE id=%s", (generate_password_hash(d["clave"]), uid))
    return jsonify(ok=True)


@app.post("/api/admin/buses")
@requiere_admin
def api_admin_buses():
    """Agrega un bus a la nómina, o activa/desactiva uno existente ({id, activo})."""
    d = request.get_json(silent=True) or {}
    with conectar() as cx:
        if d.get("id"):
            cx.execute("UPDATE buses SET activo=%s WHERE id=%s", (bool(d.get("activo")), d["id"]))
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
    return jsonify(ok=True)


@app.post("/api/admin/motivos")
@requiere_admin
def api_admin_motivos():
    """Agrega un motivo de mantenimiento, o activa/desactiva/renombra uno existente ({id, activo, nombre})."""
    d = request.get_json(silent=True) or {}
    nombre = (d.get("nombre") or "").strip()[:120]
    with conectar() as cx:
        if d.get("id"):
            if nombre:
                if cx.execute("SELECT 1 FROM motivos WHERE nombre=%s AND id<>%s", (nombre, d["id"])).fetchone():
                    return jsonify(error="Ya existe un motivo con ese nombre."), 409
                cx.execute("UPDATE motivos SET nombre=%s WHERE id=%s", (nombre, d["id"]))
            if "activo" in d:
                cx.execute("UPDATE motivos SET activo=%s WHERE id=%s", (bool(d["activo"]), d["id"]))
            return jsonify(ok=True)
        if not nombre:
            return jsonify(error="Escribe el nombre del motivo."), 400
        cx.execute("INSERT INTO motivos (nombre) VALUES (%s) ON CONFLICT (nombre) DO UPDATE SET activo=TRUE", (nombre,))
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
