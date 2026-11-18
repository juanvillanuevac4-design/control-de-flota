# Control de Flota

Registro diario de la distribución de la flota del depósito, basado en `Distrubucion_Buses.xlsx`.

- **Nómina de buses** (hoja *Distrubucion*): 130 buses, sembrados en la tabla `buses` al primer arranque. El estado inicial de cada bus viene de la hoja *Distribucion* (Patio Real / Observacion: PANNE = Mantenimiento).
- **Un reporte por día**: tabla `reportes` (fecha única) y tabla `base` con las columnas de la hoja *Base*: Fecha, Hora, Encargado, PLACA, INTERNO, TIPO BUS, TIPO FLOTA, MARCA CHASIS, TERMINAL OP, Terminal Real, Operativo, Observacion.
- **Prellenado**: al abrir el día, cada bus trae la Terminal Real, el estado y la observación del último reporte anterior. Solo se cambia lo que se movió.
- **Listas**: Terminal Real = Pie Andino, Aguirre Luco, Juanita, Otro. Operativo = Operativo, Mantenimiento.
- Hay un solo reporte por fecha. Cualquier usuario puede crearlo o corregirlo (también días anteriores); queda `registrado_por` y `modificado_por`.

- **Guardado automático**: cada cambio (arrastrar, doble click, "Aplicar y guardar") se guarda al instante; "Guardar corrección" fuerza el guardado. El primer guardado del día crea el reporte con la foto completa.
- **Historial** (`cambios`): cada día parte con la foto completa (`inicial`) y luego se agrega una fila por cada bus que cambia, con fecha y hora, usuario y el estado completo del bus. Con eso se arma el estado de la flota en cualquier fecha y hora.
- **Motivos de mantenimiento** (`motivos`): lista desplegable al marcar Mantenimiento, con mantenedor en `/admin`.
- **Perfil invitado**: usa todo, pero sus cambios quedan solo en su pantalla; el servidor rechaza cualquier escritura.

## Pantallas

- `/` Reporte diario en vista **Por terminal** (propuesta C elegida; las propuestas A Tabla y B Tarjetas quedan en el historial de git).
- `/dashboard` Descarga (Excel o CSV por rango de fechas) y dashboard **Foto del día** con etiquetas de datos (propuestas 2 y 3 en el historial de git).
- `/tableros` Tableros (todo sale del historial `cambios`, solo-agregar): evolución por estatus por día, semana o mes; mantenimiento por motivo; % de disponibilidad por hora (curva hábil/fin de semana y mapa de calor día × hora); hora punta frente a la flota requerida por terminal; tiempo fuera de servicio por motivo; Pareto de motivos; reincidencia ≤ 7 días tras volver a operar; buses detenidos hace más tiempo; comparación entre terminales; calidad del reporte (solo admin). Los buses se identifican solo por interno; el invitado ve solo agregados.
- `/reiterativos` Buses reiterativos (entran a mantenimiento N veces o más en el rango): matriz bus × tipo de falla, cada ingreso con fecha, duración, terminal y marca de reincidencia; filtro por interno y falla. Invitado: solo totales.
- `/admin` Usuarios y contraseñas (crear, cambiar rol, activar/desactivar, cambiar clave) y nómina de buses.

## Configuración (Render)

| Variable | Uso |
|---|---|
| `DATABASE_URL` | Cadena de conexión de Neon (`postgresql://...?sslmode=require`) |
| `ADMIN_USUARIO` / `ADMIN_CLAVE` | Administrador inicial (se crea solo si no existe ningún administrador) |
| `SECRET_KEY` | Firma de sesión (Render la genera) |
| `ZONA_HORARIA` | `America/Santiago` |
| `HORAS_PUNTA` | Horas punta para Tableros, por defecto `6-8,17-19` (cada hora cubre HH:00 a HH:59) |

## Ejecutar localmente

```bash
pip install -r requirements.txt
DATABASE_URL=postgresql://... ADMIN_CLAVE=algo-seguro gunicorn -w 1 --threads 4 app:app
```

## Seguridad

- Contraseñas con hash (scrypt de Werkzeug). Ingreso: máximo 5 intentos fallidos por usuario y 20 por IP cada 15 minutos (luego 15 minutos de bloqueo, queda en la bitácora). El límite vive en memoria: vale porque la app corre con un proceso (`-w 1`) y se reinicia con cada despliegue.
- Sesión de 12 horas (`SESION_HORAS`), cookie `__Host-` segura, HttpOnly y SameSite=Lax. Cambiar una contraseña cierra las demás sesiones de ese usuario en menos de un minuto.
- Permisos por rol revisados en el servidor en cada ruta (`requiere_login`, `requiere_admin`, `requiere_escritura`, `sin_listado_para_invitado`).
- CSRF: los POST de la API solo aceptan JSON y se rechazan si el navegador indica otro origen.
- Cabeceras: CSP (solo los scripts propios, por hash), HSTS, X-Frame-Options DENY, nosniff, Referrer-Policy, `Cache-Control: no-store` en datos.
- Todas las consultas SQL van parametrizadas. Los errores no muestran detalles internos. Envíos de más de 2 MB se rechazan.
- Perfil invitado: internos, placas y nombres del personal ficticios en todas las respuestas (datos personales, Ley 19.628).

## Mínimo privilegio en la base (Neon)

La app puede trabajar con un rol `flota_app` que no es dueño de nada: no puede borrar filas, ni editar `cambios` o `bitacora`, ni quitar los disparadores que las protegen. Solo el arranque usa el dueño (`DATABASE_URL_ADMIN`) para crear tablas y darle a `flota_app` sus permisos (se reaplican en cada arranque).

1. En el **SQL Editor de Neon** (crear el rol por SQL, no desde la consola, para que no quede en `neon_superuser`):
   ```sql
   CREATE TEMP TABLE _c AS SELECT replace(gen_random_uuid()::text || gen_random_uuid()::text, '-', '') AS clave;
   DO $$ BEGIN EXECUTE format('CREATE ROLE flota_app LOGIN PASSWORD %L', (SELECT clave FROM _c)); END $$;
   SELECT clave, pg_has_role('flota_app', 'neon_superuser', 'member') AS es_superusuario FROM _c;
   ```
   Copia `clave` (y confirma que `es_superusuario` sea `false`).
2. En **Render → Environment**: crea `DATABASE_URL_ADMIN` con la URL actual de `DATABASE_URL`, y cambia `DATABASE_URL` por la misma URL con `flota_app:<clave>` en lugar del usuario y contraseña. Guarda y despliega.
3. En los logs del arranque debe aparecer `Permisos del rol flota_app actualizados` y no debe aparecer el aviso `La app usa el rol dueño de la base`.

## Privacidad (Ley 19.628 / Ley 21.719)

- `/privacidad`: aviso de privacidad, enlazado desde el ingreso y la barra superior. Completa `PRIVACIDAD_RESPONSABLE` (razón social), `PRIVACIDAD_CONTACTO` (correo para derechos) en Render. `PRIVACIDAD_RETENCION` ya dice 3 años por defecto.
- "Mis datos" (barra superior): cada persona descarga sus datos y todas sus acciones registradas (derecho de acceso). La descarga queda en la bitácora.
- **Retención: 3 años.** `purgar.py` borra la bitácora anterior al corte y anonimiza nombres y usuarios en el historial (el estado de los buses se conserva). Es la única excepción a la regla de solo agregar: se corre a mano con `DATABASE_URL_ADMIN`, primero sin argumentos (simulación) y luego con `--ejecutar`, y deja en la bitácora un evento con el corte y las cantidades. Correrlo una vez al año (la primera vez con datos que purgar será en noviembre de 2029).
