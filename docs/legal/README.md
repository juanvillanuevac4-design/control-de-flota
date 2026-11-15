# Textos legales (borrador base)

> **Borrador para revisión de un abogado.** No es asesoría legal ni está validado. Se redactó a partir de lo que la app hace hoy (código en `app.py`, noviembre 2026) y de la Ley 19.628 con las modificaciones de la Ley 21.719, que entra en vigencia el 1 de diciembre de 2026.

| Archivo | Para qué sirve | Quién lo ve |
|---|---|---|
| [politica-privacidad.md](politica-privacidad.md) | Qué datos se tratan, para qué, cuánto tiempo, con quién se comparten y cómo ejercer derechos | Pública, enlazada desde el ingreso |
| [terminos-uso.md](terminos-uso.md) | Reglas de uso de la app para los usuarios internos | Pública, enlazada desde el ingreso |
| [aviso-tratamiento.md](aviso-tratamiento.md) | Aviso corto que se entrega al crear la cuenta (deber de información) | Se entrega a cada usuario nuevo |

En la app se leen en `/legal` (enlace al pie de la pantalla de ingreso y de cada pantalla).

## Antes de publicar, completar

Los campos entre corchetes: `[RAZÓN SOCIAL]`, `[RUT]`, `[DOMICILIO]`, `[CORREO DE PRIVACIDAD]`, `[CARGO RESPONSABLE]`, `[REGIÓN DE NEON]`, `[FECHA]`.

## Puntos que el abogado debe decidir

1. **Bitácora inmutable vs. derecho de supresión.** La bitácora y el historial (`bitacora`, `cambios`) no se pueden editar ni borrar (lo impide un disparador en la base). Eso choca con el derecho de supresión. El texto lo justifica en la trazabilidad operacional y la defensa ante reclamos, y fija un plazo de conservación. El abogado debe confirmar esa base y el plazo (se propone 5 años; ajustable).
2. **Base de licitud.** Se propone el contrato de trabajo o de servicios y el interés legítimo de la empresa, no el consentimiento, porque en una relación laboral el consentimiento rara vez es libre.
3. **Transferencia internacional.** Render (hosting) opera en Oregón, EE.UU. Neon (base de datos) opera en la región elegida al crearla. Google Fonts recibe la IP de quien abre la app. Se necesitan garantías adecuadas (cláusulas contractuales de los proveedores). Revisar los acuerdos de tratamiento de datos (DPA) de Render y Neon.
4. **Campo Observación (texto libre).** Puede terminar con datos de conductores o terceros. Los términos lo prohíben, pero conviene instruirlo también por escrito.
5. **Registros de acceso del proveedor.** La app no guarda direcciones IP, pero Render sí las registra en sus logs de acceso. Se declara en la política.
6. **Política de reembolso:** no se incluye porque la app no cobra a sus usuarios.
