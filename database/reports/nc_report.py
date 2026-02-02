import pandas as pd
from pathlib import Path
from database.connection import get_connection
from utils.format_valores import formatar_contabil


def generate_nc_report(selected_dates, output_path: Path):
    output_dir = Path(output_path)

    con = get_connection()
    
    dates_sql = ",".join(
        f"DATE '{d.isoformat()}'" for d in selected_dates
    )

    query = f"""
        SELECT
            emissao_dia,
            emitente_ug,
            favorecido_doc,
            ro_evento,
            ptres,
            numero_nc,
            plano_interno,
            natureza_despeza,
            siafi,
            ABS(valor_absoluto) AS valor_absoluto,
            regexp_extract(
                upper(observacao),
                'TED[: ]+([0-9]{{4,6}})',
                1
            ) AS ted
        FROM notas_credito
        WHERE emissao_dia IN ({dates_sql})
        AND ro_evento <> '301206'
    """
    df = con.execute(query).df()
    con.close()

    # 🧱 Ajustes finais
    df['emissao_dia'] = pd.to_datetime(df['emissao_dia'], errors='coerce')
    df['Emissão - Dia'] = df['emissao_dia'].dt.strftime('%d/%m/%Y')
    
    # ➕ Garantir valor positivo
    df['valor_absoluto'] = df['valor_absoluto'].abs()

    # 💰 Formatar valor
    df['valor_absoluto'] = df['valor_absoluto'].apply(formatar_contabil)

    # 🔤 Renomear colunas
    df.rename(columns={
        'emitente_ug': 'Emitente - UG',
        'favorecido_doc': 'Favorecido Doc.',
        'ro_evento': 'NC - Evento',
        'ptres': 'NC - PTRES',
        'numero_nc': 'NC',
        'plano_interno': 'NC - Plano Interno',
        'natureza_despeza': 'NC - Natureza Despesa',
        'siafi': 'NC - Transferência',
        'valor_absoluto': 'NC - Valor Linha',
        'ted': 'TED'
    }, inplace=True)

    colunas_finais = [
        'Emissão - Dia',
        'Emitente - UG',
        'Favorecido Doc.',
        'NC - Evento',
        'NC - PTRES',
        'NC',
        'NC - Plano Interno',
        'NC - Natureza Despesa',
        'NC - Transferência',
        'NC - Valor Linha',
        'TED'
    ]

    df = df[colunas_finais]

    # 📄 Arquivo de saída
    nome_data = selected_dates[0].strftime("%Y-%m-%d")

    if df.empty:
        output_file = output_dir / f"NC {nome_data} SEM DADOS.csv"
    else:
        output_file = output_dir / f"NC {nome_data}.csv"

    df.to_csv(output_file, index=False, sep=";", encoding="utf-8-sig")
    
    print(f"✅ Arquivo de NC {nome_data} gerado com sucesso")
    return df