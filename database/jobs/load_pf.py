from pathlib import Path
from load_pf_to_duckdb import excel_to_table


excel_to_table(
    excel_path=Path(
        r"W:\B - TED\7 - AUTOMAÇÃO\Banco de Dados\Financeiro\PF Legado - Exercício 2026.xlsx"
    ),
    table_name="notas_de_financeiro",
    data_start_row=7,
    column_map={
        "A": "emissao_dia",
        "B": "pf_numero",
        "C": "emitente_ug",
        "D": "emitente_ug_descricao",
        "E": "emitente_gestao",
        "F": "gestao_descricao",
        "G": "favorecido_doc",
        "H": "favorecido_doc_descricao",
        "I": "pf_evento",
        "J": "pf_evento_descricao",
        "K": "pf_categoria_gasto",
        "M": "fonte_recurso",
        "N": "fonte_recurso_descricao",
        "O": "vinculacao_pagamento",
        "P": "vinculacao_descricao",
        "Q": "siafi",
        "R": "valor_absoluto",
    }
)