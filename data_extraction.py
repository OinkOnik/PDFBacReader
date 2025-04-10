# data_extraction.py

"""
Funciones para extraer datos específicos de archivos PDF.
Contiene lógica para parsear y limpiar datos extraídos según sus tipos.
Maneja múltiples ocurrencias de los mismos títulos y organiza los datos
en una estructura coherente.

Módulos relacionados:
- constants.py: Proporciona listas de títulos y patrones
- pdf_processor.py: Utiliza estas funciones para procesar PDFs
"""

import re
import fitz  # PyMuPDF
import pandas as pd
import os
from constants import (
    TITLES_TO_EXTRACT,
    PATTERNS_TO_EXCLUDE,
    ALL_POSSIBLE_TITLES,
    TERMINAL_FORMATTED_TITLES,
    MAX_REPETITIONS,
    REPEATING_TITLES,
    BASE_TITLES
)


def extract_data_from_pdf(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)

        full_text = ""
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            full_text += page.get_text()

        pdf_document.close()

        lines = full_text.split('\n')
        data = {}

        data['Nombre del Archivo'] = os.path.basename(pdf_path)

        title_counters = {}
        i = 0

        # Lista de títulos con tratamiento especial
        special_extraction_titles = [
            "Correlativo",
            "Número Afiliado Gestión Afiliado principal",
            "Atención por",
            "Nombre del oficial técnico que brinda servicio",
            "Validación fecha"
        ]

        # Títulos que requieren extracción multilinea
        multiline_titles = [
            "Revisión General en cualquier visita"
        ]

        # Primera pasada: buscar títulos exactos incluyendo los especiales
        while i < len(lines):
            line = lines[i].strip()

            # Verificar si la línea actual comienza exactamente con uno de los títulos
            for title in TITLES_TO_EXTRACT:
                if line == title or line.startswith(title + ":") or line.startswith(title + " "):
                    # Contador para múltiples ocurrencias
                    title_counters[title] = title_counters.get(title, 0) + 1
                    if title_counters[title] > 1:
                        key = f"{title} ({title_counters[title]})"
                    else:
                        key = title

                    # Si el valor está en la misma línea tras dos puntos
                    if ":" in line:
                        possible_value = line.split(":", 1)[1].strip()
                        if possible_value:
                            data[key] = possible_value
                            i += 1
                            break

                    # Extraer valor en la siguiente línea(s)
                    value = ""
                    j = i + 1

                    special_cases = {
                        "Detalle de trabajo realizado para cierre de gestión": "Ubicación del comercio",
                        "Evaluaciones a realizar": "¿Comercio tiene Stickers actualizados?",
                        "Nombre persona que atiende": "Firma:",
                        "Tipo de gestiones": "Indique si entregó rollos de papel",
                        "Hora de salida": None,
                        "Hora de llegada": None
                    }
                    stop_pattern = special_cases.get(title)

                    # Manejo para títulos multilínea
                    if title in multiline_titles:
                        multiline_value = []
                        while j < len(lines):
                            next_line = lines[j].strip()
                            if any(next_line.startswith(t) for t in TITLES_TO_EXTRACT) or \
                               re.match(r'^[A-ZÁÉÍÓÚÑa-záéíóúñ0-9\s#¿?]+:', next_line):
                                break
                            if next_line and not any(re.search(pat, next_line) for pat in PATTERNS_TO_EXCLUDE):
                                multiline_value.append(next_line)
                            j += 1
                        if multiline_value:
                            data[key] = "\n".join(multiline_value)
                        i = j
                        break

                    # Extracción normal de un solo valor
                    while j < len(lines):
                        next_line = lines[j].strip()
                        if stop_pattern and next_line.startswith(stop_pattern):
                            break
                        if title in ["Hora de salida", "Hora de llegada"]:
                            time_match = re.search(
                                r'\d{1,2}:[\d]{2}\s*[APMapm]{2}(?:\s*(?:GMT|UTC)?[+-]?\d{1,2}:\d{2})?',
                                next_line
                            )
                            if time_match:
                                value = time_match.group(0).strip()
                            else:
                                value = next_line.strip()
                            j += 1
                            break
                        if title in special_extraction_titles:
                            if next_line and not any(next_line.startswith(t) for t in TITLES_TO_EXTRACT):
                                value = next_line
                                j += 1
                                break
                        if title != "Nombre del Afiliado" and (
                                any(next_line.startswith(t) for t in TITLES_TO_EXTRACT) or
                                re.match(r'^[A-ZÁÉÍÓÚÑa-záéíóúñ0-9\s#¿?]+:', next_line)
                        ):
                            break
                        if any(re.search(pat, next_line) for pat in PATTERNS_TO_EXCLUDE):
                            j += 1
                            continue
                        if next_line and not any(next_line.startswith(t) for t in TITLES_TO_EXTRACT):
                            value = next_line
                            j += 1
                            break
                        j += 1

                    if value:
                        data[key] = value
                    i = j
                    break
            else:
                i += 1

        # Segunda pasada: títulos especiales si faltan
        for title in special_extraction_titles:
            if title not in data:
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    if title.lower() in line.lower():
                        if ":" in line:
                            data[title] = line.split(":", 1)[1].strip()
                        elif i + 1 < len(lines) and lines[i + 1].strip():
                            data[title] = lines[i + 1].strip()
                    i += 1

        # Búsqueda específica para "Revisión General"
        if "Revisión General en cualquier visita" not in data:
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                if line.startswith("Revisión General"):
                    j = i + 1
                    multiline_value = []
                    while j < len(lines):
                        next_line = lines[j].strip()
                        if any(next_line.startswith(t) for t in TITLES_TO_EXTRACT) or \
                           re.match(r'^[A-ZÁÉÍÓÚÑa-záéíóúñ0-9\s#¿?]+:', next_line):
                            break
                        if next_line and not any(re.search(pat, next_line) for pat in PATTERNS_TO_EXCLUDE):
                            multiline_value.append(next_line)
                        j += 1
                    if multiline_value:
                        data["Revisión General en cualquier visita"] = "\n".join(multiline_value)
                    break
                i += 1

        # Extracción extra de horas en las últimas páginas
        try:
            pdf_document = fitz.open(pdf_path)
            last_pages_text = ""
            for page_num in range(max(0, len(pdf_document) - 2), len(pdf_document)):
                last_pages_text += pdf_document[page_num].get_text()
            pdf_document.close()

            for label in ["Hora de llegada", "Hora de salida"]:
                if not any(key.startswith(label) for key in data.keys()):
                    pattern = rf"{label}[:\s]*([\d]{{1,2}}:[\d]{{2}}\s*[APMapm]{{2}}(?:\s*(?:GMT|UTC)?[+-]?\d{{1,2}}:\d{{2}})?)"
                    match = re.search(pattern, last_pages_text)
                    if match:
                        data[label] = match.group(1).strip()
        except Exception:
            pass

        # Patrones especiales de búsqueda si faltan campos
        special_patterns = {
            "Correlativo": r"Correlativo[:\s]*([^\n]+)",
            "Número Afiliado Gestión Afiliado principal": r"(?:Número|N[úu]mero)\s*Afiliado\s*Gesti[óo]n\s*Afiliado\s*principal[:\s]*([^\n]+)",
            "Atención por": r"Atenci[óo]n\s*por[:\s]*([^\n]+)",
            "Nombre del oficial técnico que brinda servicio": r"Nombre\s*del\s*oficial\s*t[ée]cnico[:\s]*([^\n]+)",
            "Validación fecha": r"Validaci[óo]n\s*fecha[:\s]*([^\n]+)"
        }
        for field, pattern in special_patterns.items():
            if field not in data:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    data[field] = match.group(1).strip()

        # Patrón multiline para "Revisión General"
        special_multiline_pattern = {
            "Revisión General en cualquier visita": r"Revisi[óo]n\s*General\s*en\s*cualquier\s*visita:?([\s\S]*?)(?=(?:" +
                                                  "|".join(TITLES_TO_EXTRACT) + r"|$))"
        }
        for field, pattern in special_multiline_pattern.items():
            if field not in data:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    captured_text = match.group(1).strip()
                    cleaned_lines = [
                        ln.strip() for ln in captured_text.split('\n')
                        if ln.strip() and not any(re.search(pat, ln) for pat in PATTERNS_TO_EXCLUDE)
                    ]
                    if cleaned_lines:
                        data[field] = "\n".join(cleaned_lines)

        # Agrupamos y formateamos los datos de terminales
        process_terminal_data(data)

        return data

    except Exception as e:
        print(f"Error al procesar el PDF {pdf_path}: {str(e)}")
        return None


def process_terminal_data(data):
    """
    Agrupa los campos repetidos de terminales en bloques numerados
    y genera claves formateadas 'Terminal - Campo', 'Terminal 2 - Campo', ...
    Elimina las claves originales de los títulos repetidos.
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
        'Fotografía de SIM',
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
        "Número de Terminal"
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
