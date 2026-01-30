from pathlib import Path
from load_nc_to_duckdb import excel_to_table

excel_to_table(
    excel_path=Path(
        r"W:\B - TED\7 - AUTOMAÇÃO\Banco de Dados\Orçamentário\NC funcionando - EXERCÍCIO 2026.xlsx"
    ),
    table_name="notas_credito",
    data_start_row=7,
    column_map={
        "A": "emissao_dia",
        "B": "emitente_ug",
        "C": "emitente_ug_descricao",
        "D": "favorecido_doc",
        "E": "favorecido_doc_descricao",
        "F": "ro_evento",
        "G": "descricao_evento",
        "H": "ptres",
        "I": "numero_nc",
        "J": "plano_interno",
        "K": "plano_interno_descricao",
        "M": "natureza_despeza",
        "N": "ce",
        "O": "natureza_despeza_descricao",
        "P": "fonte_recurso",
        "Q": "fonte_recurso_descricao",
        "R": "siafi",
        "T": "valor_absoluto",
        "U": "observacao",
        "V": "lancado_por",
    }
)