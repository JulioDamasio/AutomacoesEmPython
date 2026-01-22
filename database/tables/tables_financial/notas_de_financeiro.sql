CREATE TABLE IF NOT EXISTS notas_de_financeiro (

    id BIGINT,
    emissao_dia DATE,
    pf_numero VARCHAR,
    emitente_ug VARCHAR,
    emitente_ug_descricao VARCHAR,
    emitente_gestao VARCHAR,
    gestao_descricao VARCHAR,
    favorecido_doc VARCHAR,
    favorecido_doc_descricao VARCHAR,
    pf_evento VARCHAR,
    pf_evento_descricao VARCHAR,
    pf_categoria_gasto VARCHAR,
    fonte_recurso VARCHAR,
    fonte_recurso_descricao VARCHAR,
    vinculacao_pagamento VARCHAR,
    vinculacao_descricao VARCHAR,
    siafi VARCHAR,
    valor_absoluto DECIMAL(18,2)

);