import pandas as pd
from pathlib import Path
from connection import get_connection
import string
from datetime import date, timedelta


def get_date_columns(con, table_name: str):
    query = f"""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_name = '{table_name}'
          AND data_type = 'DATE'
    """
    return [row[0] for row in con.execute(query).fetchall()]


def get_decimal_columns(con, table_name: str):
    query = f"""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_name = '{table_name}'
          AND data_type LIKE 'DECIMAL%'
    """
    return [row[0] for row in con.execute(query).fetchall()]


def datas_validas_para_carga():
    hoje = date.today()
    weekday = hoje.weekday()  # segunda=0, domingo=6

    if weekday == 0:  # segunda-feira
        return [
            hoje - timedelta(days=1),  # domingo
            hoje - timedelta(days=2),  # sábado
            hoje - timedelta(days=3),  # sexta
        ]
    else:
        return [hoje - timedelta(days=1)]


def excel_to_table(
    excel_path: Path,
    table_name: str,
    column_map: dict,
    data_start_row: int
):
    print(f"📥 Lendo arquivo: {excel_path.name}")

    def col_letter_to_index(letter):
        return string.ascii_uppercase.index(letter.upper())

    excel_cols = [col_letter_to_index(c) for c in column_map.keys()]

    df = pd.read_excel(
        excel_path,
        header=None,
        skiprows=data_start_row - 1,
        usecols=excel_cols
    )

    df.columns = list(column_map.values())
    df = df.dropna(how="all")

    print(f"📊 Linhas lidas: {len(df)}")

    con = get_connection()

    # 🔎 Tipos da tabela
    date_columns = get_date_columns(con, table_name)
    decimal_columns = get_decimal_columns(con, table_name)

    # 📅 Converter DATE
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col],
                dayfirst=True,
                errors="coerce"
            ).dt.date

    # 💰 Converter DECIMAL (pt-BR)
    for col in decimal_columns:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 🔄 Ajuste de sinal por RO - Evento
    if "ro_evento" in df.columns and "valor_absoluto" in df.columns:
        df["valor_absoluto"] = df.apply(
            lambda row: abs(row["valor_absoluto"])
            if str(row["ro_evento"]) == "301206"
            else -abs(row["valor_absoluto"]),
            axis=1
        )

    # 📆 Filtro por datas válidas
    if "emissao_dia" in df.columns:
        datas_validas = datas_validas_para_carga()
        df = df[df["emissao_dia"].isin(datas_validas)]

        print(f"📅 Datas válidas para carga: {datas_validas}")
        print(f"📊 Linhas após filtro de data: {len(df)}")

    if df.empty:
        print("⚠️ Nenhuma linha válida para inserir. Encerrando.")
        con.close()
        return

    con.register("df_temp", df)

    columns_insert = ["id"] + list(df.columns)
    columns_select = list(df.columns)

    print(f"💾 Inserindo dados em {table_name}...")

    con.execute(f"""
        INSERT INTO {table_name} ({", ".join(columns_insert)})
        SELECT
            (SELECT COALESCE(MAX(id), 0) FROM {table_name})
            + ROW_NUMBER() OVER (),
            {", ".join(columns_select)}
        FROM df_temp
    """)

    con.close()
    print("✅ Inserção concluída\n")