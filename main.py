import sys
import traceback
from PyQt6.QtWidgets import QApplication
from pdf_extractor_app import PDFExtractorApp


def exception_hook(exctype, value, tb):
    """Captura excepciones no manejadas para diagnóstico"""
    traceback_formated = traceback.format_exception(exctype, value, tb)
    traceback_string = "".join(traceback_formated)
    print(f"Error no manejado: {traceback_string}")
    sys._excepthook(exctype, value, tb)
    sys.exit(1)


if __name__ == "__main__":
    # Registrar manejador de excepciones
    sys._excepthook = sys.excepthook
    sys.excepthook = exception_hook

    try:
        app = QApplication(sys.argv)
        window = PDFExtractorApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error al iniciar la aplicación: {str(e)}")
        traceback.print_exc()