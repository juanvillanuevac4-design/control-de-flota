"""Control de Flota: un reporte diario de la distribución de buses del depósito."""
import base64
import csv
import hashlib
import io
import json
import logging
import os
import re
import secrets
import threading
import time
import unicodedata
from datetime import date, datetime, timedelta
from functools import wraps
from pathlib import Path
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

import psycopg
from flask import Flask, Response, g, jsonify, redirect, request, send_from_directory, session
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from psycopg.rows import dict_row
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash, generate_password_hash

from buses_iniciales import BUSES

TZ = ZoneInfo(os.environ.get("ZONA_HORARIA", "America/Santiago"))
def _limpiar_url(v):
    """Acepta también el snippet de Neon (psql 'postgresql://...') o la URL con comillas o espacios."""
    m = re.search(r"postgres(?:ql)?://[^\s'\"]+", v or "")
    return m.group(0) if m else (v or "").strip()


DATABASE_URL = _limpiar_url(os.environ.get("DATABASE_URL", "postgresql://localhost/flota"))
# Mínimo privilegio: la app trabaja con el rol flota_app (DATABASE_URL) y solo el arranque usa el dueño de la base
# (DATABASE_URL_ADMIN) para crear o actualizar tablas. Sin DATABASE_URL_ADMIN, todo usa DATABASE_URL como antes.
DATABASE_URL_ADMIN = _limpiar_url(os.environ.get("DATABASE_URL_ADMIN", "")) or DATABASE_URL
ROL_APP = "flota_app"
ADMIN_USUARIO = os.environ.get("ADMIN_USUARIO", "admin")
ADMIN_CLAVE = os.environ.get("ADMIN_CLAVE", "")

TERMINALES = ["Pie Andino", "Aguirre Luco", "Juanita", "Otro"]  # columna J "Terminal Real"
ESTADOS = ["Operativo", "En carga", "Mantenimiento"]  # columna K "Operativo"
# Estados que cuentan como flota disponible en dashboard y tableros. Decisión de Juan (2026-11-18):
# "En carga" (bus cargando batería) NO cuenta como disponible, aunque tampoco es mantenimiento.
DISPONIBLES = {"Operativo"}
# Columnas de la hoja "Base"
COLUMNAS_BASE = ["Fecha", "Hora", "Encargado", "PLACA", "INTERNO", "TIPO BUS", "TIPO FLOTA", "MARCA CHASIS",
                 "TERMINAL OP", "Terminal Real", "Operativo", "Motivo mantenimiento", "Observacion",
                 "Último cambio", "Cambiado por", "Sistema"]
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
    operativo TEXT NOT NULL CHECK (operativo IN ('Operativo','En carga','Mantenimiento')),
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
ALTER TABLE base DROP CONSTRAINT IF EXISTS base_operativo_check;
ALTER TABLE base ADD CONSTRAINT base_operativo_check CHECK (operativo IN ('Operativo','En carga','Mantenimiento'));
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
-- Parámetros editables por el administrador (p. ej. meta de disponibilidad): cada cambio es una fila nueva y vale la última.
CREATE TABLE IF NOT EXISTS parametros (
    id BIGSERIAL PRIMARY KEY,
    ts TIMESTAMPTZ NOT NULL DEFAULT now(),
    clave TEXT NOT NULL,
    valor TEXT NOT NULL
);  -- quién lo cambió queda en la bitácora
DROP TRIGGER IF EXISTS parametros_solo_agregar ON parametros;
CREATE TRIGGER parametros_solo_agregar BEFORE UPDATE OR DELETE ON parametros FOR EACH ROW EXECUTE FUNCTION solo_agregar();
DROP TRIGGER IF EXISTS parametros_sin_vaciar ON parametros;
CREATE TRIGGER parametros_sin_vaciar BEFORE TRUNCATE ON parametros FOR EACH STATEMENT EXECUTE FUNCTION solo_agregar();
-- En desuso desde 18-11-2026 (la meta pasó a ser % de disponibilidad en `parametros`); se conserva por su historial.
CREATE TABLE IF NOT EXISTS metas (
    id BIGSERIAL PRIMARY KEY,
    ts TIMESTAMPTZ NOT NULL DEFAULT now(),
    terminal TEXT NOT NULL,
    requeridos INT NOT NULL CHECK (requeridos >= 0)
);  -- quién la cambió queda en la bitácora
DROP TRIGGER IF EXISTS metas_solo_agregar ON metas;
CREATE TRIGGER metas_solo_agregar BEFORE UPDATE OR DELETE ON metas FOR EACH ROW EXECUTE FUNCTION solo_agregar();
DROP TRIGGER IF EXISTS metas_sin_vaciar ON metas;
CREATE TRIGGER metas_sin_vaciar BEFORE TRUNCATE ON metas FOR EACH STATEMENT EXECUTE FUNCTION solo_agregar();
-- Sistema del bus al que pertenece cada motivo (18-11-2026). Solo agregar: vale la última fila de cada motivo.
-- Los tableros agrupan por sistema; los reportes guardan el motivo tal como se escribió.
CREATE TABLE IF NOT EXISTS motivo_sistema (
    id BIGSERIAL PRIMARY KEY,
    ts TIMESTAMPTZ NOT NULL DEFAULT now(),
    motivo TEXT NOT NULL,
    sistema TEXT NOT NULL
);  -- quién lo asignó queda en la bitácora
DROP TRIGGER IF EXISTS motivo_sistema_solo_agregar ON motivo_sistema;
CREATE TRIGGER motivo_sistema_solo_agregar BEFORE UPDATE OR DELETE ON motivo_sistema FOR EACH ROW EXECUTE FUNCTION solo_agregar();
DROP TRIGGER IF EXISTS motivo_sistema_sin_vaciar ON motivo_sistema;
CREATE TRIGGER motivo_sistema_sin_vaciar BEFORE TRUNCATE ON motivo_sistema FOR EACH STATEMENT EXECUTE FUNCTION solo_agregar();
CREATE OR REPLACE VIEW motivo_sistema_actual AS
    SELECT DISTINCT ON (motivo) motivo, sistema FROM motivo_sistema ORDER BY motivo, ts DESC, id DESC;
"""

# Permisos del rol de la app: sin DELETE en ninguna tabla, y en cambios y bitácora solo leer y agregar.
# No es dueño de nada, así que tampoco puede borrar tablas ni los disparadores que protegen la bitácora.
PERMISOS_APP = f"""
GRANT USAGE ON SCHEMA public TO {ROL_APP};
REVOKE ALL ON ALL TABLES IN SCHEMA public FROM {ROL_APP};
GRANT SELECT, INSERT, UPDATE ON buses, reportes, base, usuarios, motivos TO {ROL_APP};
GRANT SELECT, INSERT ON cambios, bitacora, metas, parametros, motivo_sistema TO {ROL_APP};
GRANT SELECT ON motivo_sistema_actual TO {ROL_APP};
REVOKE ALL ON ALL SEQUENCES IN SCHEMA public FROM {ROL_APP};
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO {ROL_APP};
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
ORIGENES = {"arrastre": "Arrastre entre terminales", "doble_click": "Doble click", "selector": "Selector de estado", "detalle": "Ventana del bus",
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
# Render pone un proxy delante: con esto request.remote_addr es la IP real del cliente y request.is_secure ve el HTTPS.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
EN_RENDER = os.environ.get("RENDER") is not None
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=EN_RENDER,
    SESSION_COOKIE_NAME="__Host-sesion" if EN_RENDER else "sesion",
    # La sesión dura un turno (12 horas por defecto); después hay que volver a ingresar.
    PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.environ.get("SESION_HORAS", "12"))),
    MAX_CONTENT_LENGTH=2 * 1024 * 1024,  # ningún envío legítimo pesa más de unos 100 KB
)
logging.basicConfig(level=logging.INFO)


def conectar(url=None):
    return psycopg.connect(
        url or DATABASE_URL, row_factory=dict_row, autocommit=False,
        connect_timeout=15, keepalives=1, keepalives_idle=30, keepalives_interval=10, keepalives_count=3,
    )


def ahora():
    return datetime.now(TZ)


def iniciar_bd():
    try:
        _iniciar_bd()
    except psycopg.errors.InsufficientPrivilege:
        # La app ya corre con el rol limitado y no hay DATABASE_URL_ADMIN: las tablas existen, se sigue sin tocarlas.
        app.logger.warning("Esquema no actualizado: el rol de la app no puede crear tablas (falta DATABASE_URL_ADMIN)")
    with conectar() as cx:
        r = cx.execute("SELECT current_user AS yo, pg_get_userbyid(relowner) AS dueno FROM pg_class "
                       "WHERE relname = 'bitacora' AND relkind = 'r'").fetchone()
    if r and r["yo"] == r["dueno"]:
        app.logger.warning("La app usa el rol dueño de la base (%s): puede borrar la protección de la bitácora. "
                           "Ver README, sección Mínimo privilegio.", r["yo"])


def _iniciar_bd():
    with conectar(DATABASE_URL_ADMIN) as cx:
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
        if cx.execute("INSERT INTO migraciones (nombre) VALUES ('carga_excel_14_11') ON CONFLICT DO NOTHING RETURNING 1").fetchone():
            _carga_excel_14_11(cx)
        if cx.execute("INSERT INTO migraciones (nombre) VALUES ('motivos_por_sistema') ON CONFLICT DO NOTHING RETURNING 1").fetchone():
            _motivos_por_sistema(cx)
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
        if cx.execute("SELECT 1 FROM pg_roles WHERE rolname=%s", (ROL_APP,)).fetchone():
            cx.execute(PERMISOS_APP)
            app.logger.info("Permisos del rol %s actualizados", ROL_APP)


def _normalizar(t):
    return "".join(c for c in unicodedata.normalize("NFD", t.strip().lower()) if unicodedata.category(c) != "Mn")


def _carga_excel_14_11(cx):
    """Una sola vez (pedido de Juan Patricio, 18-11-2026): buses en mantenimiento del Excel del 14-11 con su motivo.
    1) Agrega a la lista los motivos que faltan (si ya existe uno igual salvo tildes o mayúsculas, se usa ese).
    2) Reporte del 14-11: cada bus del Excel queda en Mantenimiento con su motivo (si no hay reporte, se crea).
    3) Reportes posteriores: se arrastra el mismo estado, salvo en buses que alguien ya cambió a mano después
       del 14-11 (su dato es más nuevo y no se pisa). Todo se agrega como cambio nuevo en `cambios` y en la bitácora."""
    from carga_excel_14_11 import ARCHIVO, BUSES as CARGA, FECHA, MOTIVOS
    n = ahora()
    fecha = date.fromisoformat(FECHA)
    u = {"usuario": "sistema", "nombre": "Carga Excel 14-11 (Juan Patricio)", "rol": "sistema"}
    origen, detalle = "Carga desde Excel", f"Archivo {ARCHIVO}, pedido por Juan Patricio el 18-11-2026"
    # 1) Motivos
    existentes = {_normalizar(m["nombre"]): m for m in cx.execute("SELECT * FROM motivos").fetchall()}
    orden = (cx.execute("SELECT coalesce(max(orden), 0) AS o FROM motivos WHERE orden < 100").fetchone()["o"] or 0) + 1
    nombre_final = {}
    for m in MOTIVOS:
        e = existentes.get(_normalizar(m))
        if e:
            nombre_final[m] = e["nombre"]
            if not e["activo"]:
                cx.execute("UPDATE motivos SET activo=TRUE WHERE id=%s", (e["id"],))
                registrar(cx, "Motivo activado", origen, u, n, objeto=e["nombre"], campo="Estado",
                          antes="Inactivo", despues="Activo", detalle=detalle)
        else:
            cx.execute("INSERT INTO motivos (nombre, orden) VALUES (%s,%s)", (m, orden))
            orden += 1
            nombre_final[m] = m
            registrar(cx, "Motivo agregado", origen, u, n, objeto=m, detalle=detalle)
    carga = {i: nombre_final[m] for i, _, m in CARGA}
    # 2) Reporte del 14-11 (se crea si no existe, igual que el primer guardado del día)
    rep = cx.execute("SELECT * FROM reportes WHERE fecha=%s", (fecha,)).fetchone()
    if not rep:
        hora = n.time().replace(second=0, microsecond=0)
        rid = cx.execute("INSERT INTO reportes (fecha, hora, encargado, registrado_por) VALUES (%s,%s,%s,%s) RETURNING id",
                         (fecha, hora, u["nombre"], u["usuario"])).fetchone()["id"]
        _, filas = prellenado(cx, fecha)
        cx.cursor().executemany(
            """INSERT INTO base (reporte_id, fecha, hora, encargado, placa, interno, tipo_bus, tipo_flota,
                                 marca_chasis, terminal_op, terminal_real, operativo, motivo, observacion)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
            [(rid, fecha, hora, u["nombre"], f["placa"], f["interno"], f["tipo_bus"], f["tipo_flota"], f["marca_chasis"],
              f["terminal_op"], f["terminal_real"], f["operativo"], f["motivo"], f["observacion"]) for f in filas])
        cx.cursor().executemany(
            """INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
               VALUES (%s,%s,'inicial',%s,%s,%s,%s,%s,%s,%s)""",
            [(n, fecha, u["usuario"], u["nombre"], f["interno"], f["terminal_real"], f["operativo"], f["motivo"],
              f["observacion"]) for f in filas])
        registrar(cx, "Reporte del día creado", origen, u, n, fecha_reporte=fecha,
                  detalle=f"{len(filas)} buses, para cargar el Excel del 14-11. {detalle}")
    # 3) Del 14-11 en adelante
    resumen, respetados = [], {}
    for r in cx.execute("SELECT id, fecha FROM reportes WHERE fecha >= %s ORDER BY fecha", (fecha,)).fetchall():
        tocados = set() if r["fecha"] == fecha else {x["interno"] for x in cx.execute(
            """SELECT DISTINCT interno FROM cambios WHERE tipo='cambio' AND usuario <> 'sistema'
               AND fecha > %s AND fecha <= %s""", (fecha, r["fecha"])).fetchall()}
        actuales = {f["interno"]: f for f in cx.execute("SELECT * FROM base WHERE reporte_id=%s", (r["id"],)).fetchall()}
        aplicados = 0
        for i, motivo in carga.items():
            a = actuales.get(i)
            if not a or (a["operativo"] == "Mantenimiento" and a["motivo"] == motivo):
                continue
            if i in tocados:
                respetados.setdefault(r["fecha"].isoformat(), []).append(i)
                continue
            registrar_varias(cx, [dict(accion="Cambio de bus", origen=origen, fecha_reporte=r["fecha"], interno=i,
                                       campo=NOMBRE_CAMPO[c], antes=a[c] or "", despues=v, detalle=detalle)
                                  for c, v in (("operativo", "Mantenimiento"), ("motivo", motivo)) if (a[c] or "") != v], u, n)
            cx.execute("""UPDATE base SET operativo='Mantenimiento', motivo=%s, cambio_ts=%s, cambio_por=%s
                          WHERE reporte_id=%s AND interno=%s""", (motivo, n, u["nombre"], r["id"], i))
            cx.execute("""INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
                          VALUES (%s,%s,'cambio',%s,%s,%s,%s,'Mantenimiento',%s,%s)""",
                       (n, r["fecha"], u["usuario"], u["nombre"], i, a["terminal_real"], motivo, a["observacion"]))
            aplicados += 1
        if aplicados:
            cx.execute("UPDATE reportes SET modificado_por=%s, modificado=now() WHERE id=%s", (u["usuario"], r["id"]))
        resumen.append(f"{r['fecha']:%d-%m}: {aplicados}")
    app.logger.info("Carga Excel 14-11: buses actualizados por reporte: %s. No se pisaron (cambiados a mano después "
                    "del 14-11): %s", ", ".join(resumen), json.dumps(respetados) if respetados else "ninguno")


def _motivos_por_sistema(cx):
    """Una sola vez (pedido de Juan Patricio, 18-11-2026): motivos agrupados por sistema del bus.
    1) Registra a qué sistema pertenece cada motivo anterior (tabla motivo_sistema, solo agregar).
    2) La lista desplegable pasa a tener solo los sistemas; los motivos anteriores quedan inactivos.
    3) En el reporte más reciente, los buses en mantenimiento con un motivo anterior pasan al sistema como
       un cambio nuevo, y el motivo anterior queda en la observación. Los reportes anteriores no se tocan."""
    from motivos_sistema import MAPA, SISTEMAS
    n = ahora()
    u = {"usuario": "sistema", "nombre": "Motivos por sistema (Juan Patricio)", "rol": "sistema"}
    origen, detalle = "Corrección de datos", "Motivos agrupados por sistema del bus, pedido por Juan Patricio el 18-11-2026"
    motivos = {m["nombre"]: m for m in cx.execute("SELECT * FROM motivos").fetchall()}
    por_norma = {_normalizar(k): v for k, v in MAPA.items()}
    mapa = {nombre: por_norma[_normalizar(nombre)] for nombre in motivos if _normalizar(nombre) in por_norma}
    sin_sistema = sorted(m for m in motivos if m not in mapa and m not in SISTEMAS)
    # 1) Equivalencias
    cx.cursor().executemany("INSERT INTO motivo_sistema (ts, motivo, sistema) VALUES (%s,%s,%s)",
                            [(n, m, s) for m, s in sorted(mapa.items())])
    registrar_varias(cx, [dict(accion="Motivo asignado a sistema", origen=origen, objeto=m, campo="Sistema",
                               antes="", despues=s, detalle=detalle) for m, s in sorted(mapa.items())], u, n)
    # 2) Lista desplegable
    for i, s in enumerate(SISTEMAS):
        e = motivos.get(s)
        if not e:
            cx.execute("INSERT INTO motivos (nombre, orden) VALUES (%s,%s)", (s, i))
            registrar(cx, "Motivo agregado", origen, u, n, objeto=s, detalle=detalle)
        else:
            cx.execute("UPDATE motivos SET orden=%s, activo=TRUE WHERE id=%s", (i, e["id"]))
            if not e["activo"]:
                registrar(cx, "Motivo activado", origen, u, n, objeto=s, campo="Estado", antes="Inactivo",
                          despues="Activo", detalle=detalle)
    for m, e in motivos.items():
        if m in mapa and m not in SISTEMAS and e["activo"]:
            cx.execute("UPDATE motivos SET activo=FALSE WHERE id=%s", (e["id"],))
            registrar(cx, "Motivo desactivado", origen, u, n, objeto=m, campo="Estado", antes="Activo",
                      despues="Inactivo", detalle=f"Pasa a {mapa[m]}. {detalle}")
    # 3) Reporte más reciente
    r = cx.execute("SELECT id, fecha FROM reportes ORDER BY fecha DESC LIMIT 1").fetchone()
    aplicados = 0
    if r:
        for a in cx.execute("SELECT * FROM base WHERE reporte_id=%s AND operativo='Mantenimiento'", (r["id"],)).fetchall():
            s = mapa.get(a["motivo"])
            if not s or s == a["motivo"]:
                continue
            obs = (a["motivo"] + (" · " + a["observacion"] if a["observacion"] else ""))[:500]
            registrar_varias(cx, [dict(accion="Cambio de bus", origen=origen, fecha_reporte=r["fecha"], interno=a["interno"],
                                       campo=NOMBRE_CAMPO[c], antes=a[c] or "", despues=v, detalle=detalle)
                                  for c, v in (("motivo", s), ("observacion", obs))], u, n)
            cx.execute("""UPDATE base SET motivo=%s, observacion=%s, cambio_ts=%s, cambio_por=%s
                          WHERE reporte_id=%s AND interno=%s""", (s, obs, n, u["nombre"], r["id"], a["interno"]))
            cx.execute("""INSERT INTO cambios (ts, fecha, tipo, usuario, nombre, interno, terminal_real, operativo, motivo, observacion)
                          VALUES (%s,%s,'cambio',%s,%s,%s,%s,'Mantenimiento',%s,%s)""",
                       (n, r["fecha"], u["usuario"], u["nombre"], a["interno"], a["terminal_real"], s, obs))
            aplicados += 1
        if aplicados:
            cx.execute("UPDATE reportes SET modificado_por=%s, modificado=now() WHERE id=%s", (u["usuario"], r["id"]))
    app.logger.info("Motivos por sistema: %s equivalencias, %s buses pasados al sistema en el reporte del %s. "
                    "Motivos sin sistema (quedan igual): %s", len(mapa), aplicados, r and r["fecha"],
                    ", ".join(sin_sistema) or "ninguno")


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
    Se asignan por orden de interno, así que el orden y los grupos se mantienen; el mapeo es el mismo en toda la app.
    También los nombres y usuarios del personal (datos personales, Ley 19.628) pasan a "Usuario N"."""
    if "mapa_ficticio" not in g:
        with conectar() as cx:
            buses = cx.execute(f"SELECT interno, placa FROM buses ORDER BY {ORDEN_INTERNO}").fetchall()
            personas = cx.execute("SELECT id, usuario, nombre FROM usuarios ORDER BY id").fetchall()
        m = {}
        for i, b in enumerate(buses):
            m[b["interno"]] = str(1001 + i)
            m[b["placa"]] = f"DM{i + 1:04d}"
        yo = (usuario_actual() or {}).get("id")
        reservadas = {x.lower() for x in TERMINALES + ESTADOS + list(NOMBRE_ROL.values()) + ["Sistema", "Otro"]}
        for p in personas:
            if p["id"] == yo:
                continue
            for real, falso in ((p["nombre"], f"Usuario {p['id']}"), (p["usuario"], f"usuario{p['id']}")):
                if real and len(real) >= 3 and real.lower() not in reservadas and real not in m:
                    m[real] = falso
        patron = re.compile(r"(?<!\w)(" + "|".join(sorted(map(re.escape, m), key=len, reverse=True)) + r")(?!\w)") if m else None
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


def sin_listado_para_invitado(f):
    """El invitado ve totales y gráficos, pero ninguna descarga con el detalle por bus."""
    @wraps(f)
    def w(*a, **k):
        if es_invitado():
            return jsonify(error="Perfil invitado: puedes ver totales y gráficos, pero no el listado de buses."), 403
        return f(*a, **k)
    return w


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
        r = cx.execute("SELECT id, usuario, nombre, rol, activo, clave_hash FROM usuarios WHERE id=%s", (u["id"],)).fetchone()
    # Si la contraseña cambió, las sesiones abiertas con la contraseña anterior se cierran.
    if not r or not r["activo"] or session.get("hc") != huella_clave(r["clave_hash"]):
        session.clear()
        return
    session["u"] = {"id": r["id"], "usuario": r["usuario"], "nombre": r["nombre"], "rol": r["rol"]}
    session["verificado"] = time.time()


def huella_clave(clave_hash):
    """Huella corta de la contraseña guardada: cambia cuando cambia la contraseña, sin exponer el hash."""
    return hashlib.sha256((app.secret_key + clave_hash).encode()).hexdigest()[:16]


@app.before_request
def solo_json_en_post():
    # Protección CSRF: los formularios de otros sitios no pueden enviar JSON con cookies,
    # y además se rechaza cualquier POST que el navegador marque como venido de otro sitio.
    if request.method == "POST" and request.path.startswith("/api/"):
        if not request.is_json:
            return jsonify(error="Solicitud no válida."), 400
        origen = request.headers.get("Origin")
        if origen and urlparse(origen).netloc != request.host:
            return jsonify(error="Solicitud no válida."), 403


@app.before_request
def sin_html_directo():
    """Las pantallas solo se sirven por sus rutas (que exigen sesión), no como /static/archivo.html."""
    if request.path.startswith("/static/") and request.path.endswith(".html"):
        return jsonify(error="No encontrado."), 404


def _hashes_csp():
    """Hash de cada <script> en línea de las pantallas: solo esos scripts pueden ejecutarse (CSP)."""
    hs = set()
    for html in Path(app.static_folder).glob("*.html"):
        for js in re.findall(r"<script>(.*?)</script>", html.read_text(encoding="utf-8"), re.S):
            hs.add("'sha256-" + base64.b64encode(hashlib.sha256(js.encode()).digest()).decode() + "'")
    return " ".join(sorted(hs))


CSP = ("default-src 'self'; script-src 'self' " + _hashes_csp() + "; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
       "img-src 'self' data:; connect-src 'self'; font-src 'self' https://fonts.gstatic.com; object-src 'none'; base-uri 'none'; "
       "form-action 'self'; frame-ancestors 'none'")


@app.after_request
def cabeceras_seguridad(resp):
    h = resp.headers
    h["Content-Security-Policy"] = CSP
    h["X-Content-Type-Options"] = "nosniff"
    h["X-Frame-Options"] = "DENY"
    h["Referrer-Policy"] = "same-origin"
    h["Permissions-Policy"] = "camera=(), microphone=(), geolocation=(), payment=()"
    h["Cross-Origin-Opener-Policy"] = "same-origin"
    if request.is_secure or EN_RENDER:
        h["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    if request.path.startswith("/api/") or request.path in ("/", "/dashboard", "/tableros", "/reiterativos", "/admin", "/login", "/privacidad"):
        h["Cache-Control"] = "no-store"  # datos de la flota y del personal: que no queden en cachés
    return resp


# ---------- Límite de intentos de ingreso ----------
# En memoria: sirve porque la app corre con un solo proceso (gunicorn -w 1). Se reinicia con cada despliegue.
VENTANA_INTENTOS = 15 * 60
MAX_FALLOS_USUARIO = 5   # por usuario en 15 minutos
MAX_FALLOS_IP = 20       # por IP en 15 minutos (muchos usuarios distintos desde el mismo lugar)
_fallos = {}
_fallos_lock = threading.Lock()
# Hash de relleno: un usuario que no existe tarda lo mismo que una contraseña incorrecta.
_HASH_RELLENO = generate_password_hash(secrets.token_hex(16))


def _fallos_recientes(clave, t):
    lst = [x for x in _fallos.get(clave, []) if x > t - VENTANA_INTENTOS]
    if lst:
        _fallos[clave] = lst
    else:
        _fallos.pop(clave, None)
    return lst


def ingreso_bloqueado(usuario, ip):
    t = time.time()
    with _fallos_lock:
        return (len(_fallos_recientes("u:" + usuario, t)) >= MAX_FALLOS_USUARIO
                or len(_fallos_recientes("ip:" + ip, t)) >= MAX_FALLOS_IP)


def anotar_fallo(usuario, ip):
    """Anota un ingreso fallido; devuelve True cuando ese fallo bloquea al usuario."""
    t = time.time()
    with _fallos_lock:
        if len(_fallos) > 10000:  # tope de memoria ante un ataque con muchos usuarios inventados
            for k in list(_fallos):
                _fallos_recientes(k, t)
        _fallos.setdefault("ip:" + ip, []).append(t)
        lst = _fallos.setdefault("u:" + usuario, [])
        lst.append(t)
        return len(_fallos_recientes("u:" + usuario, t)) == MAX_FALLOS_USUARIO


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


@app.get("/tableros")
def tableros_page():
    return pagina("tableros.html")


@app.get("/reiterativos")
def reiterativos_page():
    return pagina("reiterativos.html")


@app.get("/admin")
def admin_page():
    return pagina("admin.html", admin=True)


@app.get("/login")
def login_page():
    return send_from_directory(app.static_folder, "login.html")


def _env_html(nombre, defecto):
    v = os.environ.get(nombre, "").strip() or defecto
    return v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


@app.get("/privacidad")
def privacidad_page():
    """Aviso de privacidad (Ley 19.628 / Ley 21.719). Responsable, contacto y plazo vienen de variables de entorno."""
    html = (Path(app.static_folder) / "privacidad.html").read_text(encoding="utf-8")
    for k, d in (("PRIVACIDAD_RESPONSABLE", "la empresa que opera el depósito"),
                 ("PRIVACIDAD_CONTACTO", "el administrador del sistema"),
                 ("PRIVACIDAD_RETENCION", "3 años desde cada registro. Pasado ese plazo, el registro de acciones se borra y "
                                          "en el historial de la flota tu nombre y usuario se reemplazan por \"anonimizado\"")):
        html = html.replace("{{" + k + "}}", _env_html(k, d))
    return Response(html, mimetype="text/html")


@app.post("/api/login")
def api_login():
    d = request.get_json(silent=True) or {}
    usuario = (d.get("usuario") or "").strip().lower()
    clave = d.get("clave") or ""
    if not isinstance(clave, str) or len(usuario) > 80 or len(clave) > 200:
        return jsonify(error="Usuario o contraseña incorrectos."), 401
    ip = request.remote_addr or "?"
    if ingreso_bloqueado(usuario, ip):
        return jsonify(error="Demasiados intentos fallidos. Espera 15 minutos antes de volver a intentar."), 429
    with conectar() as cx:
        r = cx.execute("SELECT * FROM usuarios WHERE lower(usuario)=%s AND activo", (usuario,)).fetchone()
    if not check_password_hash(r["clave_hash"] if r else _HASH_RELLENO, clave) or not r:
        bloquea = anotar_fallo(usuario, ip)
        if r:
            with conectar() as cx:
                registrar(cx, "Ingreso fallido (contraseña incorrecta)", "Sesión", r)
                if bloquea:
                    registrar(cx, "Usuario bloqueado 15 minutos por intentos fallidos", "Sesión", r,
                              detalle=f"{MAX_FALLOS_USUARIO} intentos fallidos seguidos")
        else:
            app.logger.warning("Ingreso fallido con un usuario que no existe (%s caracteres)", len(usuario))
        time.sleep(0.5)
        return jsonify(error="Usuario o contraseña incorrectos."), 401
    with _fallos_lock:
        _fallos.pop("u:" + usuario, None)
    session.clear()
    session.permanent = True
    session["u"] = {"id": r["id"], "usuario": r["usuario"], "nombre": r["nombre"], "rol": r["rol"]}
    session["hc"] = huella_clave(r["clave_hash"])
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
        if nueva == actual:
            return jsonify(error="La nueva contraseña debe ser distinta de la actual."), 400
        nuevo_hash = generate_password_hash(nueva)
        cx.execute("UPDATE usuarios SET clave_hash=%s WHERE id=%s", (nuevo_hash, usuario_actual()["id"]))
        registrar(cx, "Cambió su propia contraseña", "Sesión", usuario_actual(), objeto=usuario_actual()["usuario"])
    session["hc"] = huella_clave(nuevo_hash)  # esta sesión sigue; las otras sesiones abiertas se cierran
    return jsonify(ok=True)


@app.get("/api/mis-datos.csv")
@requiere_login
def api_mis_datos():
    """Derecho de acceso: cada persona descarga sus datos de usuario y todas sus acciones registradas."""
    u = usuario_actual()
    with conectar() as cx:
        p = cx.execute("SELECT usuario, nombre, rol, activo, creado FROM usuarios WHERE id=%s", (u["id"],)).fetchone()
        acciones = cx.execute("SELECT * FROM bitacora WHERE usuario=%s ORDER BY ts, id", (u["usuario"],)).fetchall()
        registrar(cx, "Descarga: mis datos personales", "Descarga", u, detalle=f"{len(acciones)} acciones")
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["Usuario", "Nombre", "Rol", "Activo", "Creado"])
    w.writerow([p["usuario"], p["nombre"], NOMBRE_ROL.get(p["rol"], p["rol"]), "Sí" if p["activo"] else "No",
                p["creado"].astimezone(TZ).strftime("%d-%m-%Y %H:%M")])
    w.writerow([])
    w.writerow(COLUMNAS_BITACORA)
    for f in acciones:
        w.writerow(fila_bitacora(f))
    # Sin ocultar(): son los propios datos de la persona, aunque sea invitado.
    return Response("\ufeff" + buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=mis_datos.csv"})


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
            """SELECT r.fecha, r.hora, r.encargado, count(*) FILTER (WHERE b.operativo = ANY(%s)) AS op, count(b.id) AS total
               FROM reportes r LEFT JOIN base b ON b.reporte_id=r.id GROUP BY r.id ORDER BY r.fecha DESC LIMIT 60""",
            (sorted(DISPONIBLES),)).fetchall()
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
    # El invitado recibe las mismas filas para contar, pero sin interno, placa ni observación.
    listado = not es_invitado()
    return jsonify(
        desde=desde.isoformat(), hasta=hasta.isoformat(), terminales=TERMINALES, listado=listado,
        estados=ESTADOS, disponibles=sorted(DISPONIBLES),
        filas=[[f["fecha"].isoformat(), f["interno"] if listado else "", f["placa"] if listado else "", f["tipo_bus"],
                f["marca_chasis"], f["terminal_real"], f["operativo"], f["observacion"] if listado else "", f["sistema"],
                f["motivo"] if listado else ""]
               for f in filas],
    )


def estado_diario(cx, desde, hasta):
    """Estado de cada bus al cierre de cada día, armado desde el historial (último registro de cada bus en ese reporte)."""
    return cx.execute(
        """SELECT x.*, r.hora, r.encargado, b.placa, b.tipo_bus, b.tipo_flota, b.marca_chasis, b.terminal_op,
                  ult.ts AS cambio_ts, ult.nombre AS cambio_por,
                  coalesce(ms.sistema, x.motivo) AS sistema
           FROM (SELECT DISTINCT ON (fecha, interno) fecha, interno, terminal_real, operativo, motivo, observacion
                 FROM cambios WHERE fecha BETWEEN %s AND %s ORDER BY fecha, interno, ts DESC, id DESC) x
           LEFT JOIN motivo_sistema_actual ms ON ms.motivo = x.motivo
           JOIN buses b ON b.interno = x.interno
           JOIN reportes r ON r.fecha = x.fecha
           LEFT JOIN LATERAL (SELECT c.ts, c.nombre FROM cambios c
                              WHERE c.fecha = x.fecha AND c.interno = x.interno AND c.tipo = 'cambio'
                              ORDER BY c.ts DESC, c.id DESC LIMIT 1) ult ON TRUE
           ORDER BY x.fecha, CASE WHEN x.interno ~ '^[0-9]+$' THEN x.interno::int END, x.interno""", (desde, hasta)
    ).fetchall()


# ---------- Tableros: evolución, mantenimiento, disponibilidad por hora y calidad del reporte ----------
# Todo sale del historial `cambios` (solo-agregar), de `reportes` y de `parametros` (meta). Los buses se identifican
# solo por su número interno (sin placa). El invitado recibe solo agregados; la calidad del reporte es solo para admin.
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MAX_DIAS_TABLEROS = 731     # dos años
MAX_DIAS_HORAS = 120        # la disponibilidad por hora usa los últimos 120 días del rango
DIAS_PREVIOS = 180          # historial previo para reconstruir episodios que empezaron antes del rango
DIAS_REINCIDENCIA = 7


def _horas_punta():
    """HORAS_PUNTA="6-8,17-19": horas (0 a 23) consideradas punta. Cada hora cubre de HH:00 a HH:59."""
    hs = set()
    for parte in os.environ.get("HORAS_PUNTA", "6-8,17-19").split(","):
        try:
            a, _, b = parte.strip().partition("-")
            hs.update(range(int(a), int(b or a) + 1))
        except ValueError:
            pass
    return sorted(h for h in hs if 0 <= h <= 23)


def _periodo(d, grano):
    if grano == "semana":
        ini = d - timedelta(days=d.weekday())
        return ini.isoformat(), f"Sem {ini.isocalendar()[1]} · {ini:%d-%m}"
    if grano == "mes":
        return f"{d:%Y-%m}", f"{MESES[d.month - 1]} {d.year}"
    return d.isoformat(), f"{d:%d-%m}"


def _contar(dic, k, n=1):
    dic[k] = dic.get(k, 0) + n


def _evolucion(diario, grano):
    """Promedio diario de buses operativos y en mantenimiento por período, y de buses en mantenimiento por motivo."""
    dias = {}
    for f in diario:
        x = dias.setdefault(f["fecha"], {"op": 0, "ca": 0, "mt": 0, "mot": {}})
        if f["operativo"] == "En carga":
            x["ca"] += 1
        elif f["operativo"] != "Mantenimiento":
            x["op"] += 1
        else:
            x["mt"] += 1
            _contar(x["mot"], f["sistema"] or "Sin motivo")
    per = {}
    for d in sorted(dias):
        clave, etiqueta = _periodo(d, grano)
        p = per.setdefault(clave, {"etiqueta": etiqueta, "desde": d, "dias": 0, "op": 0, "ca": 0, "mt": 0, "mot": {}})
        p["dias"] += 1; p["op"] += dias[d]["op"]; p["ca"] += dias[d]["ca"]; p["mt"] += dias[d]["mt"]; p["hasta"] = d
        for k, n in dias[d]["mot"].items():
            _contar(p["mot"], k, n)
    out = []
    for p in per.values():
        n, t = p["dias"], p["op"] + p["ca"] + p["mt"]
        disp = p["op"] + (p["ca"] if "En carga" in DISPONIBLES else 0)
        out.append({"etiqueta": p["etiqueta"], "desde": p["desde"].isoformat(), "hasta": p["hasta"].isoformat(), "dias": n,
                    "op": round(p["op"] / n, 1), "ca": round(p["ca"] / n, 1), "mt": round(p["mt"] / n, 1),
                    "disp": round(disp / t * 100, 1) if t else None,
                    "motivos": {k: round(v / n, 1) for k, v in p["mot"].items()}})
    return out


def _episodios(eventos, semilla, inicio_hist):
    """Episodios de mantenimiento: desde que un bus pasa a Mantenimiento hasta que vuelve a Operativo.
    Los cambios se recorren por día del reporte y, dentro del día, por hora. Una corrección hecha después
    sobre un día pasado cuenta al cierre de ese día (no a la hora en que se corrigió)."""
    cierre, abiertos, eps = {}, {}, []
    for i, s in semilla.items():
        cierre[i] = s["operativo"]
        if s["operativo"] == "Mantenimiento":
            abiertos[i] = {"interno": i, "inicio": inicio_hist, "fin": None, "motivo": s["motivo"] or "Sin motivo",
                           "terminal": s["terminal_real"], "desde_antes": True}
            eps.append(abiertos[i])
    for e in eventos:
        i = e["interno"]
        t = min(e["ts"], datetime(e["fecha"].year, e["fecha"].month, e["fecha"].day, 23, 59, 59, tzinfo=TZ))
        if e["operativo"] == "Mantenimiento":
            if i not in abiertos:
                abiertos[i] = {"interno": i, "inicio": t, "fin": None, "motivo": e["motivo"] or "Sin motivo",
                               "terminal": e["terminal_real"], "desde_antes": cierre.get(i) is None}
                eps.append(abiertos[i])
            elif e["motivo"]:
                abiertos[i]["motivo"] = e["motivo"]   # el motivo se suele precisar después: vale el último
        elif i in abiertos:
            ep = abiertos.pop(i)
            ep["fin"] = max(t, ep["inicio"])
        cierre[i] = e["operativo"]
    return eps


def _horas(ep, fin_def):
    return ((ep["fin"] or fin_def) - ep["inicio"]).total_seconds() / 3600


def _mediana(vs):
    vs = sorted(vs)
    if not vs:
        return None
    m = len(vs) // 2
    return round(vs[m] if len(vs) % 2 else (vs[m - 1] + vs[m]) / 2, 1)


def _resumen_tiempos(eps, clave, fin_def):
    """Tiempo fuera de servicio por motivo o terminal: episodios terminados (mediana y promedio) y abiertos."""
    g = {}
    for ep in eps:
        x = g.setdefault(ep[clave], {"cerrados": [], "abiertos": 0})
        if ep["fin"]:
            x["cerrados"].append(_horas(ep, fin_def))
        else:
            x["abiertos"] += 1
    out = [{"nombre": k, "reparados": len(v["cerrados"]), "abiertos": v["abiertos"], "mediana_h": _mediana(v["cerrados"]),
            "promedio_h": round(sum(v["cerrados"]) / len(v["cerrados"]), 1) if v["cerrados"] else None,
            "total_h": round(sum(v["cerrados"]), 1)} for k, v in g.items()]
    return sorted(out, key=lambda r: -(r["total_h"] or 0))


def _por_hora(eventos, semilla, desde, hasta, punta):
    """% de buses operativos al final de cada hora de cada día, y [operativos, total] por terminal en hora punta.
    En cada momento vale el último registro hasta ese momento del reporte más reciente
    (la misma regla que la descarga 'Estado a una fecha y hora')."""
    limite = ahora()
    actual = {i: (date.min, s["operativo"], s["terminal_real"]) for i, s in semilla.items()}
    evs = sorted(eventos, key=lambda e: (e["ts"], e["fecha"]))
    pendientes, j, dias, d = [], 0, [], desde
    while d <= hasta:
        resto = []
        for e in pendientes:   # registros de un reporte de fecha futura: valen desde ese día
            (resto.append(e) if e["fecha"] > d else _aplicar(actual, e))
        pendientes = resto
        fila, pt = [], {}
        for h in range(24):
            momento = datetime(d.year, d.month, d.day, h, 59, 59, 999999, tzinfo=TZ)
            if momento - timedelta(minutes=59, seconds=59) > limite:
                fila.append(None); continue
            while j < len(evs) and evs[j]["ts"] <= momento:
                (pendientes.append(evs[j]) if evs[j]["fecha"] > d else _aplicar(actual, evs[j])); j += 1
            if not actual:   # antes del primer reporte: sin datos
                fila.append(None); continue
            op = sum(1 for x in actual.values() if x[1] in DISPONIBLES)
            fila.append(round(op / len(actual) * 100, 1))
            if h in punta:
                por_t = {}
                for x in actual.values():
                    c = por_t.setdefault(x[2], [0, 0])
                    c[0] += x[1] in DISPONIBLES; c[1] += 1
                pt[h] = por_t
        if any(v is not None for v in fila):
            dias.append({"fecha": d.isoformat(), "h": fila, "punta": pt})
        d += timedelta(days=1)

    def perfil(filtro):
        out = []
        for h in range(24):
            vs = [x["h"][h] for x in dias if filtro(date.fromisoformat(x["fecha"]).weekday()) and x["h"][h] is not None]
            out.append(round(sum(vs) / len(vs), 1) if vs else None)
        return out
    return dias, {"habil": perfil(lambda w: w < 5), "finde": perfil(lambda w: w >= 5)}


def _aplicar(actual, e):
    previo = actual.get(e["interno"])
    if previo is None or e["fecha"] >= previo[0]:
        actual[e["interno"]] = (e["fecha"], e["operativo"], e["terminal_real"])


def _semilla(cx, antes_de):
    return {r["interno"]: r for r in cx.execute(
        """SELECT DISTINCT ON (c.interno) c.interno, c.operativo, coalesce(ms.sistema, c.motivo) AS motivo,
                  c.terminal_real FROM cambios c
           JOIN buses b ON b.interno = c.interno AND b.activo
           LEFT JOIN motivo_sistema_actual ms ON ms.motivo = c.motivo
           WHERE c.fecha < %s ORDER BY c.interno, c.fecha DESC, c.ts DESC, c.id DESC""", (antes_de,)).fetchall()}


def _eventos(cx, desde, hasta):
    return cx.execute(
        """SELECT c.interno, c.fecha, c.ts, c.tipo, c.usuario, c.nombre, c.operativo,
                  coalesce(ms.sistema, c.motivo) AS motivo, c.terminal_real FROM cambios c
           JOIN buses b ON b.interno = c.interno AND b.activo
           LEFT JOIN motivo_sistema_actual ms ON ms.motivo = c.motivo
           WHERE c.fecha BETWEEN %s AND %s ORDER BY c.fecha, c.ts, c.id""", (desde, hasta)).fetchall()


META_DEFECTO = float(os.environ.get("META_DISPONIBILIDAD", "95"))


def meta_disponibilidad(cx):
    """Meta de disponibilidad de la flota (% de buses operativos). La fija el administrador; por defecto 95%."""
    r = cx.execute("SELECT valor FROM parametros WHERE clave='meta_disponibilidad' ORDER BY ts DESC, id DESC LIMIT 1").fetchone()
    try:
        return float(r["valor"]) if r else META_DEFECTO
    except ValueError:
        return META_DEFECTO


def _calidad(cx, desde, hasta, eventos):
    """Quién reporta, a qué hora y cuántas correcciones hace (solo administradores)."""
    reps = cx.execute("SELECT fecha, hora, encargado, registrado_por, creado FROM reportes WHERE fecha BETWEEN %s AND %s ORDER BY fecha",
                      (desde, hasta)).fetchall()
    por_u, por_dia = {}, {}
    for e in eventos:
        if e["tipo"] != "cambio":
            continue
        local = e["ts"].astimezone(TZ)
        corr = local.date() > e["fecha"]
        u = por_u.setdefault(e["usuario"], {"usuario": e["usuario"], "nombre": e["nombre"], "cambios": 0, "correcciones": 0,
                                            "dias": set(), "horas": []})
        u["cambios"] += 1; u["correcciones"] += corr; u["dias"].add(local.date()); u["horas"].append(local.hour + local.minute / 60)
        x = por_dia.setdefault(e["fecha"], {"cambios": 0, "correcciones": 0, "usuarios": set()})
        x["cambios"] += 1; x["correcciones"] += corr; x["usuarios"].add(e["nombre"] or e["usuario"])
    hhmm = lambda h: f"{int(h):02d}:{int(round(h % 1 * 60)) % 60:02d}" if h is not None else ""
    usuarios = [{"usuario": u["usuario"], "nombre": u["nombre"], "cambios": u["cambios"], "correcciones": u["correcciones"],
                 "pct_correcciones": round(u["correcciones"] / u["cambios"] * 100, 1), "dias": len(u["dias"]),
                 "hora_mediana": hhmm(_mediana(u["horas"]))} for u in sorted(por_u.values(), key=lambda u: -u["cambios"])]
    dias = []
    for r in reps:
        x = por_dia.get(r["fecha"], {"cambios": 0, "correcciones": 0, "usuarios": set()})
        creado = r["creado"].astimezone(TZ)
        dias.append({"fecha": r["fecha"].isoformat(), "hora_reporte": r["hora"].strftime("%H:%M"),
                     "creado": creado.strftime("%d-%m %H:%M"), "tarde": creado.date() > r["fecha"],
                     "encargado": r["encargado"], "registrado_por": r["registrado_por"], "cambios": x["cambios"],
                     "correcciones": x["correcciones"], "usuarios": sorted(x["usuarios"])})
    return {"usuarios": usuarios, "dias": dias}


def _rango_pedido():
    a = request.args
    hoy = ahora().date()
    hasta = min(parse_fecha(a.get("hasta"), hoy), hoy)
    desde = parse_fecha(a.get("desde"), hasta - timedelta(days=29))
    if desde > hasta:
        desde, hasta = hasta, desde
    desde = max(desde, hasta - timedelta(days=MAX_DIAS_TABLEROS - 1))
    try:
        umbral = min(max(int(a.get("umbral", 2)), 2), 20)
    except ValueError:
        umbral = 2
    return desde, hasta, umbral


def _episodios_del_rango(eventos, semilla, previo, desde, hasta):
    """Episodios que tocan el rango (todos) y los que empiezan en él (rango). Un episodio que terminó después
    del rango se trata como abierto al final del rango."""
    fin_def = min(ahora(), datetime(hasta.year, hasta.month, hasta.day, 23, 59, 59, tzinfo=TZ))
    ini_rango = datetime(desde.year, desde.month, desde.day, tzinfo=TZ)
    todos = [ep for ep in _episodios(eventos, semilla, datetime(previo.year, previo.month, previo.day, tzinfo=TZ))
             if ep["inicio"] <= fin_def]
    for ep in todos:
        if ep["fin"] and ep["fin"] > fin_def:
            ep["fin"] = None
    return todos, [ep for ep in todos if ep["inicio"] >= ini_rango], fin_def


@app.get("/api/tableros")
@requiere_login
def api_tableros():
    desde, hasta, umbral = _rango_pedido()
    grano = request.args.get("grano") if request.args.get("grano") in ("dia", "semana", "mes") else "dia"
    h_desde = max(desde, hasta - timedelta(days=MAX_DIAS_HORAS - 1))
    previo = desde - timedelta(days=DIAS_PREVIOS)
    punta = _horas_punta()
    u = usuario_actual()
    with conectar() as cx:
        diario = estado_diario(cx, desde, hasta)
        eventos = _eventos(cx, previo, hasta)
        semilla_ep = _semilla(cx, previo)
        semilla_h = _semilla(cx, h_desde)
        meta = meta_disponibilidad(cx)
        calidad = _calidad(cx, desde, hasta, [e for e in eventos if e["fecha"] >= desde]) if u["rol"] == "admin" else None
    listado = not es_invitado()

    # Episodios: los del rango son los que empiezan en él; los abiertos al final sirven para el ranking.
    todos, rango, fin_def = _episodios_del_rango(eventos, semilla_ep, previo, desde, hasta)
    abiertos = sorted((ep for ep in todos if not ep["fin"]), key=lambda ep: ep["inicio"])

    # Reincidencia: vuelve a mantenimiento hasta 7 días después de haber vuelto a operar.
    ultimo_fin, reincid, en_rango = {}, [], {id(ep) for ep in rango}
    for ep in sorted(todos, key=lambda ep: ep["inicio"]):
        prev = ultimo_fin.get(ep["interno"])
        if id(ep) in en_rango and prev and (ep["inicio"] - prev[0]).total_seconds() <= DIAS_REINCIDENCIA * 86400:
            reincid.append({"interno": ep["interno"], "reparado": prev[0].strftime("%d-%m %H:%M"),
                            "volvio": ep["inicio"].strftime("%d-%m %H:%M"), "dias": round((ep["inicio"] - prev[0]).total_seconds() / 86400, 1),
                            "motivo_antes": prev[1], "motivo": ep["motivo"], "terminal": ep["terminal"]})
        if ep["fin"]:
            ultimo_fin[ep["interno"]] = (ep["fin"], ep["motivo"])

    # Reiterativos y Pareto de motivos (por ingresos a mantenimiento en el rango).
    por_bus, pareto, dias_mt, dias_mt_mot = {}, {}, {}, {}
    for ep in rango:
        b = por_bus.setdefault(ep["interno"], {"ingresos": 0, "motivos": {}, "ultimo": ep["inicio"]})
        b["ingresos"] += 1; _contar(b["motivos"], ep["motivo"]); b["ultimo"] = max(b["ultimo"], ep["inicio"])
        _contar(pareto, ep["motivo"])
    for f in diario:
        if f["operativo"] == "Mantenimiento":
            _contar(dias_mt, f["interno"]); _contar(dias_mt_mot, f["sistema"] or "Sin motivo")
    reit = sorted(((i, b) for i, b in por_bus.items() if b["ingresos"] >= umbral),
                  key=lambda x: (-x[1]["ingresos"], -dias_mt.get(x[0], 0)))
    distrib = {}
    for b in por_bus.values():
        _contar(distrib, b["ingresos"])
    total_ing, acum, filas_pareto = len(rango), 0, []
    for k, n in sorted(pareto.items(), key=lambda kv: -kv[1]):
        acum += n
        filas_pareto.append({"motivo": k, "ingresos": n, "pct": round(n / total_ing * 100, 1),
                             "acumulado": round(acum / total_ing * 100, 1), "dias_bus": dias_mt_mot.get(k, 0)})

    # Comparación entre terminales.
    term = {t: {"terminal": t, "dias_bus": 0, "op": 0} for t in TERMINALES}
    n_dias = len({f["fecha"] for f in diario}) or 1
    for f in diario:
        x = term.setdefault(f["terminal_real"], {"terminal": f["terminal_real"], "dias_bus": 0, "op": 0})
        x["dias_bus"] += 1; x["op"] += f["operativo"] in DISPONIBLES
    tiempos_t = {r["nombre"]: r for r in _resumen_tiempos(rango, "terminal", fin_def)}
    reinc_t = {}
    for r in reincid:
        _contar(reinc_t, r["terminal"])
    ing_t = {}
    for ep in rango:
        _contar(ing_t, ep["terminal"])
    terminales = [{"terminal": t, "buses_prom": round(x["dias_bus"] / n_dias, 1), "op_prom": round(x["op"] / n_dias, 1),
                   "disp": round(x["op"] / x["dias_bus"] * 100, 1) if x["dias_bus"] else None, "ingresos": ing_t.get(t, 0),
                   "mediana_h": tiempos_t.get(t, {}).get("mediana_h"), "reincidencias": reinc_t.get(t, 0)}
                  for t, x in term.items()]

    # Disponibilidad por hora y hora punta frente a la meta.
    dias_h, perfiles = _por_hora([e for e in eventos if e["fecha"] >= h_desde], semilla_h, h_desde, hasta, punta)
    # % de cada terminal en cada hora punta frente a la meta de disponibilidad.
    pct = lambda c: round(c[0] / c[1] * 100, 1) if c and c[1] else None
    cumplimiento = []
    for t in TERMINALES:
        vals = [v for d in dias_h for h in d["punta"] if (v := pct(d["punta"][h].get(t))) is not None]
        cumplimiento.append({"terminal": t, "horas": len(vals), "prom": round(sum(vals) / len(vals), 1) if vals else None,
                             "minimo": min(vals) if vals else None,
                             "pct_cumple": round(sum(v >= meta for v in vals) / len(vals) * 100, 1) if vals else None})
    tot_punta = [pct([sum(c[0] for c in d["punta"][h].values()), sum(c[1] for c in d["punta"][h].values())])
                 for d in dias_h for h in d["punta"]]
    tot_punta = [v for v in tot_punta if v is not None]
    punta_dias = [{"fecha": d["fecha"], "min": {t: min((v for h in d["punta"] if (v := pct(d["punta"][h].get(t))) is not None), default=None)
                                                 for t in TERMINALES}} for d in dias_h if d["punta"]]

    ahora_ = ahora()
    return jsonify(
        desde=desde.isoformat(), hasta=hasta.isoformat(), grano=grano, umbral=umbral, listado=listado,
        terminales=TERMINALES, horas_punta=punta, dias_reincidencia=DIAS_REINCIDENCIA, en_carga_disponible="En carga" in DISPONIBLES,
        evolucion=_evolucion(diario, grano),
        tiempos={"motivo": _resumen_tiempos(rango, "motivo", fin_def), "terminal": list(tiempos_t.values()),
                 "mediana_h": _mediana([_horas(ep, fin_def) for ep in rango if ep["fin"]]), "reparados": sum(1 for ep in rango if ep["fin"])},
        reiterativos={"ingresos": total_ing, "buses_con_ingreso": len(por_bus), "reiterativos": len(reit),
                      "ingresos_reiterativos": sum(b["ingresos"] for _, b in reit),
                      "distribucion": [[k, distrib[k]] for k in sorted(distrib)]},
        reincidencia={"total": len(reincid), "tasa": round(len(reincid) / total_ing * 100, 1) if total_ing else None,
                      "lista": reincid if listado else []},
        pareto=filas_pareto,
        detenidos={"total": len(abiertos), "mas_de_7": sum(1 for ep in abiertos if (fin_def - ep["inicio"]).days >= 7),
                   "lista": [{"interno": ep["interno"], "desde": ep["inicio"].strftime("%d-%m-%Y %H:%M"),
                              "dias": round((fin_def - ep["inicio"]).total_seconds() / 86400, 1), "desde_antes": ep["desde_antes"],
                              "motivo": ep["motivo"], "terminal": ep["terminal"]} for ep in abiertos] if listado else [],
                   "por_motivo": sorted(_resumen_contar(abiertos, "motivo").items(), key=lambda kv: -kv[1])},
        terminales_cmp=terminales,
        horas={"desde": h_desde.isoformat(), "dias": [{"fecha": d["fecha"], "h": d["h"]} for d in dias_h], **perfiles},
        meta=meta,
        punta={"cumplimiento": cumplimiento, "dias": punta_dias,
               "flota": {"prom": round(sum(tot_punta) / len(tot_punta), 1) if tot_punta else None,
                         "pct_cumple": round(sum(v >= meta for v in tot_punta) / len(tot_punta) * 100, 1) if tot_punta else None}},
        calidad=calidad, generado=ahora_.strftime("%d-%m-%Y %H:%M"),
    )


@app.get("/api/reiterativos")
@requiere_login
def api_reiterativos():
    """Buses que entran a mantenimiento `umbral` veces o más en el rango, con sus tipos de falla (motivos) y cada ingreso.
    Se identifican solo por interno. El invitado recibe solo agregados."""
    desde, hasta, umbral = _rango_pedido()
    previo = desde - timedelta(days=DIAS_PREVIOS)
    with conectar() as cx:
        diario = estado_diario(cx, desde, hasta)
        eventos = _eventos(cx, previo, hasta)
        semilla = _semilla(cx, previo)
        tipos = {r["interno"]: r["tipo_bus"] for r in cx.execute("SELECT interno, tipo_bus FROM buses").fetchall()}
    todos, rango, fin_def = _episodios_del_rango(eventos, semilla, previo, desde, hasta)
    dias_mt = {}
    for f in diario:
        if f["operativo"] == "Mantenimiento":
            _contar(dias_mt, f["interno"])
    # Fin del episodio anterior de cada bus, para marcar las reincidencias (vuelve a fallar <= 7 días después).
    ultimo_fin, reinc = {}, set()
    for ep in sorted(todos, key=lambda ep: ep["inicio"]):
        prev = ultimo_fin.get(ep["interno"])
        if prev and (ep["inicio"] - prev).total_seconds() <= DIAS_REINCIDENCIA * 86400:
            reinc.add(id(ep))
        if ep["fin"]:
            ultimo_fin[ep["interno"]] = ep["fin"]
    por_bus = {}
    for ep in sorted(rango, key=lambda ep: ep["inicio"]):
        por_bus.setdefault(ep["interno"], []).append(ep)
    distrib = {}
    for eps in por_bus.values():
        _contar(distrib, len(eps))
    reit = {i: eps for i, eps in por_bus.items() if len(eps) >= umbral}
    mot_ing, mot_buses, mot_horas = {}, {}, {}
    buses = []
    for i, eps in reit.items():
        mot = {}
        for ep in eps:
            _contar(mot, ep["motivo"]); _contar(mot_ing, ep["motivo"]); _contar(mot_horas, ep["motivo"], _horas(ep, fin_def))
        for k in mot:
            _contar(mot_buses, k)
        principal = max(mot.items(), key=lambda kv: kv[1])
        buses.append({
            "interno": i, "tipo_bus": tipos.get(i, ""), "ingresos": len(eps), "dias_mt": dias_mt.get(i, 0),
            "horas_fuera": round(sum(_horas(ep, fin_def) for ep in eps), 1), "motivos": mot,
            "principal": principal[0], "repite": principal[1] > 1, "reincidencias": sum(id(ep) in reinc for ep in eps),
            "detenido": eps[-1]["fin"] is None, "primero": eps[0]["inicio"].strftime("%d-%m-%Y"),
            "ultimo": eps[-1]["inicio"].strftime("%d-%m-%Y"),
            "episodios": [{"inicio": ep["inicio"].strftime("%d-%m-%Y %H:%M"),
                           "fin": ep["fin"].strftime("%d-%m-%Y %H:%M") if ep["fin"] else "",
                           "horas": round(_horas(ep, fin_def), 1), "motivo": ep["motivo"], "terminal": ep["terminal"],
                           "reincidencia": id(ep) in reinc} for ep in eps]})
    buses.sort(key=lambda b: (-b["ingresos"], -b["horas_fuera"]))
    motivos = [{"motivo": k, "ingresos": n, "buses": mot_buses[k], "horas": round(mot_horas[k], 1)}
               for k, n in sorted(mot_ing.items(), key=lambda kv: -kv[1])]
    listado = not es_invitado()
    return jsonify(
        desde=desde.isoformat(), hasta=hasta.isoformat(), umbral=umbral, listado=listado, dias_reincidencia=DIAS_REINCIDENCIA,
        generado=ahora().strftime("%d-%m-%Y %H:%M"),
        resumen={"ingresos": len(rango), "buses_con_ingreso": len(por_bus), "reiterativos": len(reit),
                 "ingresos_reiterativos": sum(len(e) for e in reit.values()),
                 "horas_reiterativos": round(sum(b["horas_fuera"] for b in buses), 1),
                 "horas_total": round(sum(_horas(ep, fin_def) for ep in rango), 1),
                 "detenidos": sum(b["detenido"] for b in buses), "mismo_motivo": sum(b["repite"] for b in buses),
                 "distribucion": [[k, distrib[k]] for k in sorted(distrib)]},
        motivos=motivos, buses=buses if listado else [])


def _resumen_contar(eps, clave):
    out = {}
    for ep in eps:
        _contar(out, ep[clave])
    return out


@app.post("/api/admin/meta")
@requiere_admin
def api_admin_meta():
    """Meta de disponibilidad de la flota en % ({meta}). Cada cambio es una fila nueva y queda en la bitácora."""
    d = request.get_json(silent=True) or {}
    try:
        n = round(float(str(d.get("meta")).replace(",", ".")), 1)
        assert 0 < n <= 100
    except (TypeError, ValueError, AssertionError):
        return jsonify(error="Escribe un porcentaje entre 0 y 100."), 400
    with conectar() as cx:
        antes = meta_disponibilidad(cx)
        if antes == n:
            return jsonify(ok=True)
        cx.execute("INSERT INTO parametros (clave, valor) VALUES ('meta_disponibilidad', %s)", (str(n),))
        registrar(cx, "Meta de disponibilidad", "Administración", usuario_actual(), objeto="Flota",
                  campo="Meta de disponibilidad (%)", antes=f"{antes:g}", despues=f"{n:g}")
    return jsonify(ok=True)


def filas_descarga():
    desde = parse_fecha(request.args.get("desde"), date(1900, 1, 1))
    hasta = parse_fecha(request.args.get("hasta"), date(2999, 12, 31))
    with conectar() as cx:
        filas = estado_diario(cx, desde, hasta)
    return [[f["fecha"], f["hora"].strftime("%H:%M"), f["encargado"], f["placa"], f["interno"], f["tipo_bus"],
             f["tipo_flota"], f["marca_chasis"], f["terminal_op"], f["terminal_real"], f["operativo"], f["motivo"],
             f["observacion"], f["cambio_ts"].astimezone(TZ).strftime("%d-%m-%Y %H:%M") if f["cambio_ts"] else "",
             f["cambio_por"], f["sistema"] if f["operativo"] == "Mantenimiento" else ""] for f in filas]


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
        cond.append("usuario = %s"); par.append(interno_real(a["usuario"]))
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
    listado = not es_invitado()
    return jsonify(total=len(filas), filas=[fila_bitacora(f) for f in filas[:300]] if listado else [],
                   listado=listado, resumen=res,
                   usuarios=[{"usuario": u["usuario"], "nombre": u["nombre"]} for u in usuarios])


@app.get("/api/bitacora.<fmt>")
@requiere_login
@sin_listado_para_invitado
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
@sin_listado_para_invitado
def api_descarga(fmt):
    filas = filas_descarga()
    with conectar() as cx:
        registrar_descarga(cx, "Base por día", fmt)
    return como_xlsx(COLUMNAS_BASE, filas, "base_flota", "Base") if fmt == "xlsx" else como_csv(COLUMNAS_BASE, filas, "base_flota")


@app.get("/api/estado.<fmt>")
@requiere_login
@sin_listado_para_invitado
def api_estado(fmt):
    m = momento_param()
    filas = filas_estado(m)
    with conectar() as cx:
        registrar_descarga(cx, "Estado a una fecha y hora", fmt)
    nombre = "estado_flota_" + m.strftime("%Y%m%d_%H%M")
    return como_xlsx(COLUMNAS_ESTADO, filas, nombre, "Estado") if fmt == "xlsx" else como_csv(COLUMNAS_ESTADO, filas, nombre)


@app.get("/api/cambios.<fmt>")
@requiere_login
@sin_listado_para_invitado
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
        meta = meta_disponibilidad(cx)
    return jsonify(usuarios=us, buses=bs, motivos=ms, yo=usuario_actual(), meta=meta)


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
            nuevo_hash = generate_password_hash(d["clave"])
            cx.execute("UPDATE usuarios SET clave_hash=%s WHERE id=%s", (nuevo_hash, uid))
            if uid == usuario_actual()["id"]:
                session["hc"] = huella_clave(nuevo_hash)
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
    return jsonify(error="Error del servidor. Intenta de nuevo; si sigue, avisa al administrador."), 500


@app.get("/healthz")
def healthz():
    return "ok"


iniciar_bd()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
