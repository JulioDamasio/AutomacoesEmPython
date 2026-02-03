# database/connection.py
import duckdb
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "cgsodb.duckdb"

def get_connection():
    return duckdb.connect(str(DB_PATH))

print("Conexão executada com sucesso!")            