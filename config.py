import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent  # Путь к главной папке приложения

SERVER_URL = "http://localhost:8000/calculate/"

CSV_PATH = BASE_DIR / "сsv_files"
DRILLING_CSV_PATH = CSV_PATH / "drilling_fluid"

IMAGE_PATH = BASE_DIR / "images"
ICONS_PATH = IMAGE_PATH / "icons"

DB_PATH = BASE_DIR / "db_files"
path4 = "msh_files/"
path5 = "projects/"
path6 = "project/"
