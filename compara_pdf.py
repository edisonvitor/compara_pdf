import os
import re
import fitz
import pandas as pd
from collections import defaultdict
from tkinter import Tk
from tkinter.filedialog import askdirectory

def extract_text_from_pdf(pdf_path):
    """Extrai o texto de um PDF usando PyMuPDF."""
    try:
        # Abre o PDF e extrai o texto
        with fitz.open(pdf_path) as pdf_document:
            text = ""
            for page in pdf_document:
                text += page.get_text()
        return text
    except Exception as e:
        # mostrar o erro e retorna uma string vazia
        print(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""

def extrair_nome_principal(texto):
    """Extrai o nome principal do docente."""
    
    padrao_nome = r'Nome\s+([^\n]+)'
    match = re.search(padrao_nome, texto)
    if match:
        return match.group(1).strip()
    return None

def extrair_nomes_citacoes(texto):
    """Extrai as variações dos nomes em citações bibliográficas."""
    padrao_citacoes = r'Nome em citações\s+bibliográficas\s+(.*?)\s+Lattes iD'
    match = re.search(padrao_citacoes, texto, re.DOTALL)
    if match:
        citacoes = match.group(1).replace('\n', ' ').strip()
        return [nome.strip() for nome in citacoes.split(';') if nome.strip()]
    return []

def compare_references_with_docentes(articles_folder, docentes_folder):
    """Compara referências bibliográficas dos artigos com os nomes dos docentes."""
    results = []
    docentes_info = {}

    # Lista todos os PDFs nas duas pastas
    articles_pdfs = [f for f in os.listdir(articles_folder) if f.endswith('.pdf')]
    docentes_pdfs = [f for f in os.listdir(docentes_folder) if f.endswith('.pdf')]

    # Extração dos nomes dos docentes dos PDFs da banca
    for docente_pdf in docentes_pdfs:
        docente_text = extract_text_from_pdf(os.path.join(docentes_folder, docente_pdf))
        if docente_text:
            nome_principal = extrair_nome_principal(docente_text)
            nomes_citacoes = extrair_nomes_citacoes(docente_text)
            if nome_principal and nomes_citacoes:
                docentes_info[nome_principal] = {
                    "nomes_citacoes": set(nomes_citacoes),
                    "citacoes": 0,
                    "arquivos": []
                }

    # Para cada PDF de artigo, compara com os nomes dos docentes
    for article_pdf in articles_pdfs:
        article_text = extract_text_from_pdf(os.path.join(articles_folder, article_pdf))
        if article_text:
            for nome_principal, info in docentes_info.items():
                for nome_citacao in info["nomes_citacoes"]:
                    if nome_citacao in article_text:
                        info["citacoes"] += 1
                        info["arquivos"].append(article_pdf)

    # Montando os resultados
    for nome_principal, info in docentes_info.items():
        results.append({
            'Docente': nome_principal,
            'Situação': f'{info["citacoes"]} citação(ões)' if info["citacoes"] > 0 else 'Ok',
            'Arquivos': ', '.join(info['arquivos']) if info['arquivos'] else 'Nenhum'
        })

    return results

def save_results_to_excel(results, output_file):
    
    df = pd.DataFrame(results)
    
    # Formatação da planilha
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultados')
        worksheet = writer.sheets['Resultados']
        
        # Ajuste de largura das colunas
        for col in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

# Interface de seleção de pastas
def select_folder(prompt):
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    folder = askdirectory(title=prompt)
    return folder

# Seleção de pastas pelo usuário
articles_folder = select_folder("Selecione a pasta com os artigos")
docentes_folder = select_folder("Selecione a pasta com os docentes")

# Comparar referências com docentes
results = compare_references_with_docentes(articles_folder, docentes_folder)

# Salvando os resultados em Excel
output_file = 'resultados_comparacao_docentes.xlsx'
save_results_to_excel(results, output_file)

print(f'Resultados salvos em {output_file}')
