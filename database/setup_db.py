# database/setup_db.py
from connection import get_connection
from tables import create_test_table

def setup_database():
    conn = get_connection()
    create_test_table(conn)
    conn.close()
    print("✅ Tabela de teste criada com sucesso.")

if __name__ == "__main__":
    setup_database()