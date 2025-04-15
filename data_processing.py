# data_processing.py

"""
Funciones de procesamiento y transformación de datos extraídos de PDFs.
Contiene lógica para organizar, agrupar y formatear los datos estructurados
según las necesidades del proyecto.

Módulos relacionados:
- constants.py: Proporciona listas de títulos y patrones
- data_extraction.py: Utiliza estas funciones para procesar los datos extraídos
"""

import pandas as pd
from constants import (
    REPEATING_TITLES,
    MAX_REPETITIONS,
    BASE_TITLES,
    TERMINAL_FORMATTED_TITLES,
    ALL_POSSIBLE_TITLES
)


def process_terminal_data(data):
    """
    Agrupa los campos repetidos de terminales en bloques numerados
    y genera claves formateadas 'Terminal - Campo', 'Terminal 2 - Campo', ...
    Elimina las claves originales de los títulos repetidos.

    Args:
        data (dict): Diccionario con los datos extraídos
    """
    # 1) Agrupar valores por índice de terminal
    grouped = {}
    for field in REPEATING_TITLES:
        for i in range(1, MAX_REPETITIONS + 1):
            original_key = field if i == 1 else f"{field} ({i})"
            if original_key in data:
                grouped.setdefault(i, {})[field] = data[original_key]
                # Eliminar inmediatamente la clave original para evitar duplicados
                data.pop(original_key)

    # 2) Insertar claves formateadas en orden ascendente
    for i in sorted(grouped.keys()):
        prefix = "" if i == 1 else f" {i}"
        for field in REPEATING_TITLES:
            if field in grouped[i]:
                formatted_key = f"Terminal{prefix} - {field}"
                data[formatted_key] = grouped[i][field]


def merge_dataframes(df_list):
    """
    Combina múltiples DataFrames en uno solo, asegurando que todas las columnas
    estén presentes y ordenadas correctamente.

    Args:
        df_list (list): Lista de DataFrames a combinar

    Returns:
        pd.DataFrame: DataFrame combinado con todas las columnas ordenadas
    """
    if not df_list:
        return pd.DataFrame()

    all_columns = set()
    for df in df_list:
        all_columns.update(df.columns)

    # Añadir solo los títulos base y los títulos formateados para Terminal
    # No incluir los títulos repetidos sin formato (Actualización en Sistema Adquirente (2), etc.)
    all_columns.update(BASE_TITLES)  # Solo títulos base, no ALL_POSSIBLE_TITLES
    all_columns.update(TERMINAL_FORMATTED_TITLES)
    all_columns.add('Nombre del Archivo')

    # Definir orden deseado
    non_repeating_columns = [
        'Nombre del Archivo',
        'Fecha de Reporte',
        'Correlativo',
        'Número Afiliado Gestión Afiliado principal',
        'Nombre del Afiliado',
        '#Oportunidad',
        'Atención por',
        'Cantidad GSM',
        'Cierre de gestión',
        'Detalle de trabajo realizado para cierre de gestión',
        'Entrega de Papelería y Cantidad',
        'Evaluaciones a realizar',
        'Fecha resolución',
        'Hora de llegada',
        'Hora de salida',
        'Indique número de SS',
        'Nombre del oficial técnico que brinda servicio',
        'Nombre persona que atiende',
        'Revisión General en cualquier visita',
        'Tipo de gestiones',
        'Tipo de terminal instalada, reprogramada o retirada',
        'Técnico que atiende',
        'Validación fecha',
        '¿El datáfono instalado lleva código QR?',
        '¿Es posible capturar el correo electrónico del comercio?',
        '¿Instalar SIM adicional?',
        '¿POS GSM Prestada?'
    ]

    repeating_columns = []
    ordered_repeating_fields = [
        "Actualización en Sistema Adquirente",
        "Esta serie fue",
        "Esta serie lleva SIM",
        "Modelo de Terminal",
        "Número de SIM",
        "Número de Serie",
        "Número de Terminal",
        "Comentario"
    ]

    for i in range(1, MAX_REPETITIONS + 1):
        prefix = "" if i == 1 else f" {i}"
        for field in ordered_repeating_fields:
            repeating_columns.append(f"Terminal{prefix} - {field}")

    ordered_columns = [
        col for col in non_repeating_columns + repeating_columns
        if col in all_columns
    ]

    # Aquí es donde se añaden columnas adicionales que no deberían estar
    # Solo incluir columnas que no sean de los títulos repetidos sin formato
    remaining_columns = sorted([
        col for col in all_columns - set(ordered_columns)
        if not any(col.startswith(title) and ('(' in col) for title in REPEATING_TITLES)
    ])

    all_columns_ordered = ordered_columns + remaining_columns

    complete_dfs = []
    for df in df_list:
        # Primero eliminar las columnas no deseadas de cada DataFrame
        df_cols = list(df.columns)
        cols_to_drop = [col for col in df_cols if
                        any(col.startswith(title) and ('(' in col) for title in REPEATING_TITLES)]
        if cols_to_drop:
            df = df.drop(columns=cols_to_drop)

        # Crear un nuevo DataFrame con todas las columnas requeridas
        # Esto evita la fragmentación al añadir columnas una por una
        new_data = {}
        for col in all_columns_ordered:
            if col in df.columns:
                new_data[col] = df[col]
            else:
                new_data[col] = [None] * len(df)

        # Crear un nuevo DataFrame sin fragmentación
        new_df = pd.DataFrame(new_data, columns=all_columns_ordered)
        complete_dfs.append(new_df)

    result_df = pd.concat(complete_dfs, ignore_index=True)
    return result_df