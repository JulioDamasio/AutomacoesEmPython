import pandas as pd
import csv
from pathlib import Path
from database.connection import get_connection
from utils.format_valores import formatar_contabil


def generate_pf_legado_report(selected_dates, output_path: Path):
    output_dir = Path(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    con = get_connection()
    
    dates_sql = ",".join(
        f"DATE '{d.isoformat()}'" for d in selected_dates
    )

    query = f"""
        SELECT
            emissao_dia,
            pf_numero,
            emitente_ug,
            emitente_gestao,
            gestao_descricao,
            favorecido_doc,
            favorecido_doc_descricao,
            pf_evento,
            pf_categoria_gasto,
            fonte_recurso,
            vinculacao_pagamento,
            siafi,
            valor_absoluto
        FROM notas_de_financeiro
        WHERE emissao_dia IN ({dates_sql})
          AND CAST(emitente_ug AS VARCHAR) <> '152734'
        ORDER BY emissao_dia, pf_numero
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

    df['Coluna D'] = df['pf_evento']
    df['Coluna F'] = df['fonte_recurso']

    df.rename(columns={
        'pf_numero': 'PF',
        'emitente_ug': 'Emitente - UG',
        'emitente_gestao': 'Emitente - Gestão',
        'gestao_descricao': '',
        'favorecido_doc': 'Favorecido Doc.',
        'favorecido_doc_descricao' : '',
        'pf_evento': 'PF - Evento',
        'pf_categoria_gasto': 'PF - Categoria Gasto',
        'fonte_recurso': 'PF - Fonte Recursos',
        'vinculacao_pagamento': 'PF - Vinculação Pagamento',
        'siafi': 'PF - Inscrição',
        'valor_absoluto': 'PF - Valor Linha'
    }, inplace=True)

    colunas_finais = [
        'Emissão - Dia', 'PF', 'Emitente - UG', 'Emitente - Gestão',
        'Coluna D', 'Favorecido Doc.', 'Coluna F', 'PF - Evento',
        'PF - Categoria Gasto', 'PF - Fonte Recursos',
        'PF - Vinculação Pagamento', 'PF - Inscrição', 'PF - Valor Linha'
    ]

    df = df[colunas_finais]

    # 📄 Arquivo de saída
    nome_data = selected_dates[0].strftime("%Y-%m-%d")

    if df.empty:
        output_file = output_dir / f"PF {nome_data} SEM DADOS.csv"
    else:
        output_file = output_dir / f"PF {nome_data}.csv"

    df.to_csv(output_file, index=False, sep=";", encoding="utf-8-sig")
    
    print(f"✅ Arquivo de PF {nome_data} gerado com sucesso")
    return df