import os
import fitz  # PyMuPDF
from datetime import datetime
import re

def extract_and_highlight_keywords(folder_path, keywords):
    output_document = fitz.open()  # Documento de saída único

    # Filtra apenas os arquivos PDF na pasta
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    
    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        try:
            pdf_document = fitz.open(file_path)

            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text("text")  # Obter texto da página
                highlighted = False

                for keyword in keywords:
                    # Cria uma expressão regular para buscar a palavra-chave completa
                    keyword_regex = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
                    for match in keyword_regex.finditer(text):
                        # Encontra as posições do texto correspondente
                        instances = page.search_for(match.group(0), quads=True)
                        for inst in instances:
                            highlight = page.add_highlight_annot(inst)
                            highlighted = True

                if highlighted:
                    # Adicionar a página atual ao documento de saída
                    output_document.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

            pdf_document.close()

        except Exception as e:
            print(f"Erro ao processar o arquivo {filename}: {e}")

    if len(output_document) > 0:
        # Salvar o documento de saída com todas as páginas destacadas
        today = datetime.today().strftime('%d-%m-%y')
        output_filename = f"{today} Resumo.pdf"
        output_path = os.path.join(folder_path, output_filename)
        output_document.save(output_path)
        output_document.close()

        print(f"Todas as páginas destacadas foram salvas em: {output_path}")
    else:
        print("Nenhum termo encontrado nos arquivos PDF.")

# Caminho da pasta e palavras-chave
folder_path = r"W:\B - TED\7 - AUTOMAÇÃO\Imprensa Nacional"
keywords = [
    'Termo de execução descentralizada', 'TED', 'Ministério da educação', 'SECADI',
    'SECRETARIA DE EDUCAÇÃO CONTINUADA ALFABETIZAÇÃO DE JOVENS E ADULTOS, DIVERSIDADE E INCLUSÃO', 'Secretaria de Educação Continuada, Alfabetização de Jovens e Adultos',
    'SESU', 'SETEC', 'SEB', 'Secretaria de Educação Básica', 'Secretaria de Educação Superior',
    'Secretaria de Educação Profissional e Tecnológica', 'SERES', '"MEC"', 'servidores da administração publica federal direta', 'MINISTRO DE ESTADO DA EDUCAÇÃO','Administração direta','Plano Interno - PI', 'Plano Interno', '"SPO"', 'Subsecretaria de planejamento e orçamento', 'Subsecretaria de Tecnologia da Informação e Comunicação', 'STIC', 'Secretaria de Educação Continuada, Alfabetização de Jovens e Adultos, Diversidade e Inclusão', 'TRANSFEREGOV', 'Transferegov', 'Transferegov.br', 'Lei Orçamentária'
]

# Extração e destaque dos termos nos PDFs
extract_and_highlight_keywords(folder_path, keywords)