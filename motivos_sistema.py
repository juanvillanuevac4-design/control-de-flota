"""Motivos de mantenimiento agrupados por sistema del bus (pedido de Juan Patricio, 18-11-2026).

El historial no se edita: los reportes antiguos conservan su motivo original y los tableros lo agrupan
por sistema leyendo la tabla `motivo_sistema` (solo agregar; vale la última fila de cada motivo).
"""

# Lista nueva de motivos, en el orden de la lista desplegable.
SISTEMAS = ["Motor", "Transmisión", "Frenos", "Suspensión y dirección", "Eléctrico", "Climatización", "Neumáticos",
            "Carrocería", "Puertas", "Batería/carga", "Siniestro y vandalismo", "Revisión técnica", "Certificación",
            "Mantenimiento preventivo", "Validador / Torniquete", "Donante",
            "Andes Motors (garantía)", "Otros"]

# Motivo anterior -> sistema. Los que ya son un sistema (Frenos, Neumáticos, ...) quedan igual.
MAPA = {
    "Motor": "Motor", "Aceite motor": "Motor", "Refrigeración": "Motor", "Mecánica": "Motor",
    "Frenos": "Frenos",
    "Suspensión": "Suspensión y dirección", "Dirección": "Suspensión y dirección",
    "Eléctrica": "Eléctrico", "Motor de arranque": "Eléctrico", "Baterías": "Eléctrico",  # baterías de 24 V de los diésel
    "Validador / Torniquete": "Validador / Torniquete",
    "Climatización": "Climatización", "Aire acondicionado": "Climatización",
    "Neumáticos": "Neumáticos", "Carrocería": "Carrocería",
    "Siniestro": "Siniestro y vandalismo", "Vandalismo": "Siniestro y vandalismo",
    "Revisión técnica": "Revisión técnica", "Certificación": "Certificación",
    "Mantenimiento preventivo": "Mantenimiento preventivo", "Check list": "Mantenimiento preventivo",
    "Donante": "Donante", "Andes Motors (garantía)": "Andes Motors (garantía)",
    "Otro": "Otros",
}
