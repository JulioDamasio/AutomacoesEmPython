# database/setup_db.py
from pathlib import Path
from connection import get_connection

BASE_DIR = Path(__file__).parent
TABLES_DIR = BASE_DIR / "tables"

def executar_sql(caminho_sql: Path):
    print(f"📄 Executando: {caminho_sql.name}")
    with open(caminho_sql, "r", encoding="utf-8") as f:
        sql = f.read()

    con = get_connection()
    con.execute(sql)
    con.close()

def criar_tabelas_budget():
    pasta = TABLES_DIR / "tables_budget"
    for sql_file in pasta.glob("*.sql"):
        executar_sql(sql_file)

def criar_tabelas_financial():
    pasta = TABLES_DIR / "tables_financial"
    for sql_file in pasta.glob("*.sql"):
        executar_sql(sql_file)

if __name__ == "__main__":
    criar_tabelas_budget()
    criar_tabelas_financial()
    print("✅ Banco estruturado com sucesso")