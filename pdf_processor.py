# pdf_processor.py
"""
Clase para procesar PDFs en segundo plano usando QThread.
Maneja la extracción de datos sin bloquear la interfaz de usuario.

Módulos relacionados:
- data_extraction.py: Contiene las funciones de extracción de datos
- pdf_extractor_app.py: Utiliza esta clase para procesar PDFs
"""

import re
import pandas as pd
from PyQt6.QtCore import QThread, pyqtSignal
from data_extraction import extract_data_from_pdf, merge_dataframes
from constants import BASE_TITLES, REPEATING_TITLES, ALL_POSSIBLE_TITLES, TERMINAL_FORMATTED_TITLES

class PDFExtractorThread(QThread):
    """Hilo para procesar PDFs sin bloquear la interfaz"""
    progress_updated = pyqtSignal(int)
    extraction_finished = pyqtSignal(pd.DataFrame)
    error_occurred = pyqtSignal(str)

    def __init__(self, pdf_files):
        """
        Inicializa el hilo de extracción.

        Args:
            pdf_files (list): Lista de rutas a archivos PDF
        """
        super().__init__()
        self.pdf_files = pdf_files
        self.running = True

    def get_terminal_sort_key(self, column_name):
        """
        Crea una clave de ordenación segura para nombres de columna de terminal.

        Args:
            column_name (str): Nombre de la columna

        Returns:
            tuple: Clave de ordenación (número de terminal, parte del campo)
        """
        # Intentar extraer un número después de "Terminal"
        match = re.search(r"Terminal\s*(\d+)", column_name)
        terminal_num = int(match.group(1)) if match else 1

        # Obtener la parte después del guión si existe
        field_part = column_name.split(" - ")[1] if " - " in column_name else column_name

        return (terminal_num, field_part)

    def run(self):
        """Procesa los PDFs y emite señales de progreso y finalización"""
        try:
            # Lista para almacenar los resultados de cada PDF
            all_data = []

            # Procesar cada archivo PDF
            total_files = len(self.pdf_files)
            for i, pdf_file in enumerate(self.pdf_files):
                if not self.running:
                    break

                # Extraer datos del PDF
                data = extract_data_from_pdf(pdf_file)
                if data:
                    # Convertir el diccionario a DataFrame (una sola fila)
                    df = pd.DataFrame([data])
                    all_data.append(df)

                # Actualizar progreso
                progress = int((i + 1) / total_files * 100)
                self.progress_updated.emit(progress)

            # Crear DataFrame con todos los resultados
            if all_data and self.running:
                # Combinar todos los DataFrames asegurando que tengan todas las columnas posibles
                result_df = merge_dataframes(all_data)

                # Definir el orden de las columnas
                # 1. Primero las columnas importantes
                important_cols = ['Nombre del Archivo', 'Fecha de Reporte', 'Correlativo',
                                 'Número Afiliado Gestión', 'Nombre del Afiliado']

                # 2. Luego los títulos base (una sola vez)
                ordered_cols = [col for col in important_cols if col in result_df.columns]

                # 3. Añadir el resto de títulos base en orden alfabético
                other_base_cols = [col for col in BASE_TITLES
                                  if col in result_df.columns and col not in ordered_cols]
                other_base_cols.sort()
                ordered_cols.extend(other_base_cols)

                # 4. Añadir las columnas de terminal formateadas
                terminal_cols = [col for col in result_df.columns
                                if "Terminal" in col and " - " in col]

                # Ordenar columnas de terminal usando la función segura de extracción de número
                terminal_cols.sort(key=self.get_terminal_sort_key)
                ordered_cols.extend(terminal_cols)

                # 5. Finalmente, añadir cualquier columna restante
                remaining_cols = [col for col in result_df.columns
                                 if col not in ordered_cols]
                remaining_cols.sort()
                ordered_cols.extend(remaining_cols)

                # Reordenar el DataFrame
                result_df = result_df[ordered_cols]

                self.extraction_finished.emit(result_df)
            elif self.running:
                # Si no hay datos, crear un DataFrame vacío con todas las columnas posibles
                empty_df = pd.DataFrame(columns=['Nombre del Archivo'] + ALL_POSSIBLE_TITLES + TERMINAL_FORMATTED_TITLES)
                self.error_occurred.emit("No se pudieron extraer datos de los PDFs seleccionados")
                self.extraction_finished.emit(empty_df)

        except Exception as e:
            import traceback
            error_msg = f"Error durante la extracción: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)  # Imprimir detalles en la consola para diagnóstico
            if self.running:
                self.error_occurred.emit(f"Error durante la extracción: {str(e)}")

    def stop(self):
        """Detiene el procesamiento"""
        self.running = False