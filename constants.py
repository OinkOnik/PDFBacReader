# constants.py

"""
Archivo de constantes para la aplicación de extracción de datos PDF.
Contiene listas de títulos a buscar y patrones a excluir en los documentos PDF.
Módulos relacionados:
- data_extraction.py: Usa estas constantes para extraer datos
"""

# Títulos base que solo aparecen una vez
BASE_TITLES = [
    "Fecha de Reporte",
    "Correlativo",
    "Indique número de SS",
    "#Oportunidad",
    "Número Afiliado Gestión Afiliado principal",
    "Nombre del Afiliado",
    "Nombre del oficial técnico que brinda servicio",
    "Evaluaciones a realizar",
    "Entrega de Papelería y Cantidad",
    "Cierre de gestión",
    "Nombre persona que atiende",
    "Detalle de trabajo realizado para cierre de gestión",
    "¿Es posible capturar el correo electrónico del comercio?",
    "Fecha resolución",
    "Validación fecha",
    "Hora de llegada",
    "Hora de salida",
    "Tipo de terminal instalada, reprogramada o retirada",
    "¿POS GSM Prestada?",
    "Cantidad GSM",
    "Datos de terminal",
    "¿Instalar SIM adicional?",
    "Fotografía de SIM",
    "¿El datáfono instalado lleva código QR?",
    "Tipo de gestiones",
    "Atención por",
    "Técnico que atiende",
    "Revisión General en cualquier visita",
]

# Títulos que se repiten hasta 20 veces (en el orden correcto)
REPEATING_TITLES = [
    "Actualización en Sistema Adquirente",
    "Esta serie fue",
    "Esta serie lleva SIM",
    "Modelo de Terminal",
    "Número de SIM",
    "Número de Serie",
    "Número de Terminal"
]

# El número máximo de repeticiones
MAX_REPETITIONS = 20

# Generar la lista completa de títulos a extraer
TITLES_TO_EXTRACT = BASE_TITLES.copy()
TITLES_TO_EXTRACT.extend(REPEATING_TITLES)

# Generar la lista completa de todos los posibles títulos (para asegurar que estén en el Excel)
ALL_POSSIBLE_TITLES = BASE_TITLES.copy()
for title in REPEATING_TITLES:
    ALL_POSSIBLE_TITLES.append(title)  # Título base sin número
    for i in range(2, MAX_REPETITIONS + 1):
        ALL_POSSIBLE_TITLES.append(f"{title} ({i})")

# Para el formato de terminal agrupado
TERMINAL_FORMATTED_TITLES = []
for i in range(1, MAX_REPETITIONS + 1):
    prefix = "" if i == 1 else f" {i}"
    for title in REPEATING_TITLES:
        TERMINAL_FORMATTED_TITLES.append(f"Terminal{prefix} - {title}")

# Patrones que deben ser excluidos (como instrucciones o elementos de navegación)
PATTERNS_TO_EXCLUDE = [
    r"^Page \d+$",
    r"^Powered by",
    r"^F-COM -",
    r"^Para BAC Credomatic",
    r"^https?://",
]
