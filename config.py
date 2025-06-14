from pathlib import Path

# Ruta base del proyecto
BASE_DIR = Path(__file__).resolve().parent

CSV_PATH = BASE_DIR / "Data" / "Superstore.csv"

IMAGE_DIR=BASE_DIR / "Data_with_image"

EXCEL_EXPORT_PATH = BASE_DIR / "Exported_Data" / "SuperstoreSummary.xlsx"
