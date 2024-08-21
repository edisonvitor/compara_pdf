import PyPDF2
import fitz

def extract_text_from_pdf(pdf_path):
    """Extrai o texto de um PDF usando PyMuPDF.

    Args:
        pdf_path (str): Caminho para o arquivo PDF.

    Returns:
        str: Texto extraído ou uma string vazia se não for possível extrair texto.
    """
    try:
        with fitz.open(pdf_path) as pdf_document:
            text = ""
            for page in pdf_document:
                text += page.get_text()
            # Salvando o texto em um arquivo TXT
            text_file_path = "texto_extraido.txt"  # Substitua pelo caminho desejado
            with open(text_file_path, "w", encoding="utf-8") as text_file:
                text_file.write(text)
            print(f"Texto salvo em {text_file_path}")
        return text
    except Exception as e:
        print(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""  # Retorna uma string vazia em caso de erro
    
extract_text_from_pdf('PROFESSOR - Diego Silva Batista.pdf')