# Control de Flota

Registro diario de la distribución de la flota del depósito, basado en `Distrubucion_Buses.xlsx`.

- **Nómina de buses** (hoja *Distrubucion*): 130 buses, sembrados en la tabla `buses` al primer arranque. El estado inicial de cada bus viene de la hoja *Distribucion* (Patio Real / Observacion: PANNE = Mantenimiento).
- **Un reporte por día**: tabla `reportes` (fecha única) y tabla `base` con las columnas de la hoja *Base*: Fecha, Hora, Encargado, PLACA, INTERNO, TIPO BUS, TIPO FLOTA, MARCA CHASIS, TERMINAL OP, Terminal Real, Operativo, Observacion.
- **Prellenado**: al abrir el día, cada bus trae la Terminal Real, el estado y la observación del último reporte anterior. Solo se cambia lo que se movió.
- **Listas**: Terminal Real = Pie Andino, Aguirre Luco, Juanita, Otro. Operativo = Operativo, Mantenimiento.
- Hay un solo reporte por fecha. Cualquier usuario puede crearlo o corregirlo (también días anteriores); queda `registrado_por` y `modificado_por`.

## Pantallas

- `/` Reporte diario en vista **Por terminal** (propuesta C elegida; las propuestas A Tabla y B Tarjetas quedan en el historial de git).
- `/dashboard` Descarga (Excel o CSV por rango de fechas) y dashboard **Foto del día** con etiquetas de datos (propuestas 2 y 3 en el historial de git).
- `/admin` Usuarios y contraseñas (crear, cambiar rol, activar/desactivar, cambiar clave) y nómina de buses.

## Configuración (Render)

| Variable | Uso |
|---|---|
| `DATABASE_URL` | Cadena de conexión de Neon (`postgresql://...?sslmode=require`) |
| `ADMIN_USUARIO` / `ADMIN_CLAVE` | Administrador inicial (se crea solo si no existe ningún administrador) |
| `SECRET_KEY` | Firma de sesión (Render la genera) |
| `ZONA_HORARIA` | `America/Santiago` |

## Ejecutar localmente

```bash
pip install -r requirements.txt
DATABASE_URL=postgresql://... ADMIN_CLAVE=algo-seguro gunicorn -w 1 --threads 4 app:app
```
