"""Retención de datos personales: 3 años (Ley 21.719). Única excepción a la regla de solo agregar.

Se corre a mano, nunca desde la app, con el rol dueño de la base:
    DATABASE_URL_ADMIN=postgresql://... python purgar.py             # solo muestra qué haría
    DATABASE_URL_ADMIN=postgresql://... python purgar.py --ejecutar  # lo hace

Qué hace con lo anterior al corte (hoy menos 3 años):
- bitacora: borra las filas (es el registro de actividad de cada persona).
- cambios, reportes y base: conserva el estado de los buses, pero reemplaza los nombres y usuarios por "anonimizado".
Todo en una transacción, y deja en la bitácora un evento con el corte y las cantidades.
"""
import os
import sys
from datetime import date

import psycopg

ANIOS = 3
ANON = "anonimizado"


def main():
    url = os.environ.get("DATABASE_URL_ADMIN")
    if not url:
        sys.exit("Falta DATABASE_URL_ADMIN (la URL del dueño de la base).")
    ejecutar = "--ejecutar" in sys.argv
    hoy = date.today()
    corte = hoy.replace(year=hoy.year - ANIOS) if not (hoy.month == 2 and hoy.day == 29) else date(hoy.year - ANIOS, 2, 28)
    with psycopg.connect(url) as cx:
        n = {
            "bitácora (filas borradas)": cx.execute(
                "SELECT count(*) FROM bitacora WHERE (ts AT TIME ZONE 'America/Santiago')::date < %s", (corte,)).fetchone()[0],
            "historial (filas anonimizadas)": cx.execute(
                "SELECT count(*) FROM cambios WHERE fecha < %s AND usuario <> %s", (corte, ANON)).fetchone()[0],
            "reportes (anonimizados)": cx.execute(
                "SELECT count(*) FROM reportes WHERE fecha < %s AND encargado <> %s", (corte, ANON)).fetchone()[0],
        }
        print(f"Corte: todo lo anterior al {corte:%d-%m-%Y}")
        for k, v in n.items():
            print(f"  {k}: {v}")
        if not any(n.values()):
            print("Nada que purgar.")
            return
        if not ejecutar:
            print("Simulación: no se cambió nada. Agrega --ejecutar para aplicarlo.")
            return
        cx.execute("ALTER TABLE bitacora DISABLE TRIGGER bitacora_solo_agregar")
        cx.execute("ALTER TABLE cambios DISABLE TRIGGER cambios_solo_agregar")
        cx.execute("DELETE FROM bitacora WHERE (ts AT TIME ZONE 'America/Santiago')::date < %s", (corte,))
        cx.execute("UPDATE cambios SET usuario=%s, nombre=%s WHERE fecha < %s AND usuario <> %s", (ANON, ANON, corte, ANON))
        cx.execute("""UPDATE reportes SET encargado=%s, registrado_por=%s,
                      modificado_por=CASE WHEN modificado_por IS NULL THEN NULL ELSE %s END
                      WHERE fecha < %s AND encargado <> %s""", (ANON, ANON, ANON, corte, ANON))
        cx.execute("UPDATE base SET encargado=%s, cambio_por=CASE WHEN cambio_por='' THEN '' ELSE %s END WHERE fecha < %s",
                   (ANON, ANON, corte))
        cx.execute("ALTER TABLE bitacora ENABLE TRIGGER bitacora_solo_agregar")
        cx.execute("ALTER TABLE cambios ENABLE TRIGGER cambios_solo_agregar")
        cx.execute(
            """INSERT INTO bitacora (usuario, nombre, rol, origen, accion, detalle)
               VALUES ('sistema', 'Sistema', 'sistema', 'Retención de datos', 'Purga por plazo de retención (3 años)', %s)""",
            (f"Anterior al {corte:%d-%m-%Y}: " + "; ".join(f"{k}: {v}" for k, v in n.items()),))
    print("Listo. La purga quedó registrada en la bitácora.")


if __name__ == "__main__":
    main()
