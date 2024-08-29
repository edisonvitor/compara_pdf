import os
import re
import fitz
import pandas as pd
from collections import defaultdict
from tkinter import Tk
from tkinter.filedialog import askdirectory
from tqdm import tqdm

def extract_text_from_pdf(pdf_path):
    """Extrai o texto de um PDF usando PyMuPDF."""
    try:
        with fitz.open(pdf_path) as pdf_document:
            text = ""
            for page in pdf_document:
                text += page.get_text()
        return text
    except Exception as e:
        print(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""

def extrair_nome_principal(texto):
    """Extrai o nome principal do docente ou candidato."""
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

def extrair_contextos_secoes(texto, nome_citacoes, padrao_secao, padrao_fim_secao):
    """Extrai contextos dentro de seções específicas."""
    contextos = defaultdict(list)
    secoes = re.findall(rf'{padrao_secao}.*?{padrao_fim_secao}', texto, re.DOTALL)
    for secao in secoes:
        for nome in nome_citacoes:
            if nome in secao:
                contextos[nome].append(secao.strip())
    return contextos

def compare_references_with_docentes(candidates_folder, docentes_folder):
    """Compara referências bibliográficas dos candidatos com os nomes dos docentes."""
    results = []
    docentes_info = {}

    # Lista todos os PDFs nas duas pastas
    candidates_pdfs = [f for f in os.listdir(candidates_folder) if f.endswith('.pdf')]
    docentes_pdfs = [f for f in os.listdir(docentes_folder) if f.endswith('.pdf')]

    # Extração dos nomes dos docentes dos PDFs da banca
    print("\nProcessando os docentes...")
    for docente_pdf in tqdm(docentes_pdfs, desc="Docentes processados"):
        docente_text = extract_text_from_pdf(os.path.join(docentes_folder, docente_pdf))
        if docente_text:
            nome_principal = extrair_nome_principal(docente_text)
            nomes_citacoes = extrair_nomes_citacoes(docente_text)
            if nome_principal and nomes_citacoes:
                docentes_info[nome_principal] = {
                    "nomes_citacoes": set(nomes_citacoes),
                    "citacoes": defaultdict(lambda: defaultdict(list))  # Lista de contextos para cada candidato
                }

    # Para cada PDF de candidato, compara com os nomes dos docentes
    print("\nProcessando os candidatos...")
    for candidate_pdf in tqdm(candidates_pdfs, desc="Candidatos processados"):
        candidate_text = extract_text_from_pdf(os.path.join(candidates_folder, candidate_pdf))
        if candidate_text:
            nome_candidato = extrair_nome_principal(candidate_text)
            if nome_candidato:
                for nome_principal, info in docentes_info.items():
                    contextos = extrair_contextos_secoes(
                        candidate_text,
                        info["nomes_citacoes"],
                        padrao_secao=r'Artigos completos publicados em periódicos',
                        padrao_fim_secao=r'\d+\.\s|Referências|Bibliografia'
                    )
                    for nome_citacao, ctx in contextos.items():
                        info["citacoes"][nome_candidato][nome_citacao].extend(ctx)

    # Montando os resultados
    all_candidatos = set()
    for info in docentes_info.values():
        all_candidatos.update(info["citacoes"].keys())

    for nome_principal, info in docentes_info.items():
        total_citacoes = sum(len(citacoes) for citacoes in info["citacoes"].values())
        row = {
            'Docente': nome_principal,
            'Situação': f'{total_citacoes} citação(ões)' if total_citacoes > 0 else 'Ok'
        }
        for candidato in all_candidatos:
            citacoes = info["citacoes"].get(candidato, {})
            row[candidato] = ' | '.join([f'{var}: {", ".join(ctx)}' for var, ctx in citacoes.items()]) if citacoes else ''
        results.append(row)

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

def select_folder(prompt):
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    folder = askdirectory(title=prompt)
    return folder

# Seleção de pastas pelo usuário
candidates_folder = select_folder("Selecione a pasta com os candidatos")
docentes_folder = select_folder("Selecione a pasta com os docentes")

# Comparar referências com docentes
results = compare_references_with_docentes(candidates_folder, docentes_folder)

# Salvando os resultados em Excel
output_file = 'resultados_comparacao_docentes.xlsx'
save_results_to_excel(results, output_file)

print(f'\nResultados salvos em {output_file}')
