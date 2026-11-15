# Política de privacidad de Control de Flota

> Borrador para revisión de un abogado. Última actualización: [FECHA].

## 1. Quién es el responsable

El responsable del tratamiento de los datos personales de esta aplicación es **[RAZÓN SOCIAL]**, RUT **[RUT]**, con domicilio en **[DOMICILIO]** (en adelante, "la Empresa"). Para cualquier consulta sobre privacidad, escribir a **[CORREO DE PRIVACIDAD]**, a cargo de **[CARGO RESPONSABLE]**.

Esta política se rige por la Ley N° 19.628 sobre Protección de la Vida Privada, con las modificaciones de la Ley N° 21.719.

## 2. Qué es la aplicación

Control de Flota es una herramienta interna para registrar a diario la distribución de los buses del depósito (terminal, estado operativo, motivo de mantenimiento y observaciones). Solo la usan personas con una cuenta creada por un administrador de la Empresa. No está dirigida al público ni a menores de edad.

## 3. Qué datos personales tratamos

| Dato | De quién | Para qué |
|---|---|---|
| Nombre de usuario, nombre y rol (administrador, operador o invitado) | Usuarios de la app | Identificar a quien ingresa y definir qué puede hacer |
| Contraseña, guardada solo como resumen cifrado (*hash*); nadie puede leerla | Usuarios | Verificar el ingreso |
| Fecha de creación de la cuenta y si está activa | Usuarios | Administrar las cuentas |
| Bitácora de acciones: usuario, nombre, rol, fecha y hora, pantalla desde la que actuó, acción, y el valor anterior y el nuevo | Usuarios | Saber quién hizo cada cambio y cuándo, y reconstruir el estado de la flota |
| Ingresos, salidas e intentos de ingreso fallidos | Usuarios | Seguridad de las cuentas |
| Descargas de reportes (qué se descargó y cuándo) | Usuarios | Control de la información que sale de la app |
| Nombre del encargado de cada reporte diario | Usuarios | Responsabilidad del reporte |
| Cookie de sesión | Usuarios | Mantener la sesión abierta hasta 7 días |
| Dirección IP y datos técnicos de la conexión | Cualquier persona que abra la app | Los registra el proveedor de hosting en sus logs de acceso y los recibe Google al cargar las fuentes tipográficas. La app no los guarda en su base de datos |

La app también guarda datos de los buses (placa, número interno, tipo, marca, terminal). Son datos de bienes de la Empresa y, por sí solos, no identifican a personas.

**No tratamos** datos sensibles (salud, origen étnico, opiniones políticas, datos biométricos, etc.), ni datos de pasajeros. El campo *Observación* es de texto libre y no debe usarse para escribir datos de personas (ver los Términos de uso).

## 4. Base de licitud

Tratamos estos datos porque son necesarios para:

- **Cumplir el contrato** de trabajo o de prestación de servicios con cada usuario, del que forma parte el uso de esta herramienta.
- **El interés legítimo de la Empresa** en controlar su operación, en la seguridad de la información y en poder acreditar quién hizo cada registro ante reclamos o fiscalizaciones. Ese interés no prevalece sobre los derechos del usuario: solo se registran acciones dentro de la app, no se monitorea al usuario fuera de ella.

No usamos los datos para publicidad, no los vendemos y no tomamos decisiones automatizadas sobre las personas.

## 5. Cuánto tiempo los conservamos

- **Cuenta de usuario:** mientras la persona tenga relación con la Empresa. Al terminar, la cuenta se desactiva (no puede ingresar).
- **Bitácora e historial de cambios:** **[5] años** desde cada registro. Estos registros no se pueden modificar ni borrar desde la app: es una medida deliberada para que la trazabilidad sea confiable. Las correcciones se registran como un evento nuevo que deja a la vista el valor anterior. Cumplido el plazo, la Empresa los eliminará o anonimizará.
- **Logs del proveedor de hosting:** según los plazos del proveedor.

## 6. Con quién los compartimos

Solo con los proveedores que hacen funcionar la app, que actúan como encargados del tratamiento por cuenta de la Empresa:

| Proveedor | Servicio | País |
|---|---|---|
| Render Services, Inc. | Hosting de la aplicación | Estados Unidos (Oregón) |
| Neon (Databricks, Inc.) | Base de datos PostgreSQL | [REGIÓN DE NEON] |
| Google LLC (Google Fonts) | Fuentes tipográficas; recibe la IP del navegador | Estados Unidos |

Esto constituye una **transferencia internacional de datos**. La Empresa la realiza con las garantías que exige la ley, como las cláusulas contractuales de protección de datos de cada proveedor. Dentro de la Empresa, los datos los ven los usuarios de la app según su rol: los reportes y la bitácora son visibles para los usuarios con cuenta.

También podremos entregar datos a autoridades cuando una ley o una orden judicial lo exija.

## 7. Cómo protegemos los datos

- Conexión cifrada (HTTPS) y base de datos con conexión cifrada.
- Contraseñas guardadas solo como resumen cifrado, con mínimo de 8 caracteres.
- Cookie de sesión protegida (no accesible desde scripts, solo por HTTPS, con protección contra solicitudes de otros sitios).
- Permisos por rol: el rol invitado no puede guardar nada.
- Bitácora que no se puede alterar.

Si ocurre una vulneración de seguridad que pueda afectar los datos, la Empresa la informará a la Agencia de Protección de Datos Personales y, cuando corresponda, a las personas afectadas, según lo exige la ley.

## 8. Tus derechos

Como titular de los datos puedes ejercer los derechos de **acceso, rectificación, supresión, oposición, portabilidad y bloqueo**. Para hacerlo, escribe a **[CORREO DE PRIVACIDAD]** indicando tu nombre, tu usuario en la app y qué pides. Es gratuito. Responderemos dentro del plazo legal (30 días corridos, prorrogables una vez por igual plazo con aviso).

Sobre la **supresión** de la bitácora: como esos registros sirven para acreditar la operación y defender a la Empresa y a los propios usuarios ante reclamos, podremos conservarlos durante el plazo del punto 5 aunque se pida su eliminación, en los casos que la ley lo permite. Si así ocurre, te lo explicaremos por escrito.

Si no quedas conforme con la respuesta, puedes reclamar ante la **Agencia de Protección de Datos Personales**.

## 9. Cambios a esta política

Si cambia esta política, publicaremos la nueva versión en la app con su fecha. Si el cambio es importante, avisaremos a los usuarios antes de que rija.
