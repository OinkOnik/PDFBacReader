# data_extraction.py

"""
Funciones para extraer datos crudos de archivos PDF.
Contiene la lógica principal para parsear y extraer información de PDFs.

Módulos relacionados:
- constants.py: Proporciona listas de títulos y patrones
- data_processing.py: Procesa los datos extraídos para su estructuración
- pdf_processor.py: Utiliza estas funciones para procesar PDFs
"""

import re
import fitz  # PyMuPDF
import os
from constants import (
    TITLES_TO_EXTRACT,
    PATTERNS_TO_EXCLUDE,
    ALL_POSSIBLE_TITLES,
    TERMINAL_FORMATTED_TITLES,
    REPEATING_TITLES,
    BASE_TITLES
)
from data_processing import process_terminal_data, merge_dataframes


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
            "Validación fecha",
            "Entrega de Papelería y Cantidad"  # Añadido a la lista
        ]

        # Títulos que requieren extracción multilinea
        multiline_titles = [
            "Revisión General en cualquier visita"
        ]

        # Primera pasada: buscar títulos exactos incluyendo los especiales
        while i < len(lines):
            line = lines[i].strip()

            for title in TITLES_TO_EXTRACT:
                if line == title or line.startswith(title + ":") or line.startswith(title + " "):
                    title_counters[title] = title_counters.get(title, 0) + 1
                    key = f"{title} ({title_counters[title]})" if title_counters[title] > 1 else title

                    if ":" in line:
                        possible_value = line.split(":", 1)[1].strip()
                        if possible_value:
                            data[key] = possible_value
                            i += 1
                            break

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

        # ✅ PROCESAMIENTO ESPECIAL PARA LA TABLA "Entrega de Papelería y Cantidad"
        if "Entrega de Papelería y Cantidad" in data:
            material_table_pattern = r"Entrega de Papelería y Cantidad.*?Material\s+Cantidad\s+(.*?)(?:Gestión de Papelería|$)"
            table_match = re.search(material_table_pattern, full_text, re.DOTALL)

            if table_match:
                table_content = table_match.group(1).strip()
                rows = re.findall(r"([^\n]+?)\s+(\d+)(?:\s*$|\n)", table_content)

                if rows:
                    materials_data = []
                    for material, quantity in rows:
                        materials_data.append(f"{material.strip()}: {quantity.strip()}")
                    data["Entrega de Papelería y Cantidad"] = "\n".join(materials_data)

            elif data.get("Entrega de Papelería y Cantidad") == "":
                i = 0
                while i < len(lines):
                    if "Material" in lines[i] and "Cantidad" in lines[i]:
                        material_line = i + 1
                        if material_line < len(lines) and lines[material_line].strip():
                            parts = re.split(r'\s{2,}', lines[material_line].strip())
                            if len(parts) >= 2:
                                material = parts[0].strip()
                                quantity = parts[-1].strip()
                                data["Entrega de Papelería y Cantidad"] = f"{material}: {quantity}"
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

        process_terminal_data(data)

        return data

    except Exception as e:
        print(f"Error al procesar el PDF {pdf_path}: {str(e)}")
        return None
