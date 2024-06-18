import fitz 
from docx import Document
import os
import tkinter as tk
from tkinter import filedialog

def extract_info_from_pdf(pdf_path):
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"No such file: '{pdf_path}'")
    
    doc = fitz.open(pdf_path)
    extracted_text = ""
    
    # Extrair texto de todas as páginas
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        extracted_text += page.get_text()
    
    return extracted_text

def filter_info(text):
    # Palavras-chave para procurar no texto
    keywords = ["Veículo", "Placa", "Ano", "Cor", "Lote", "Nº motor", "Câmbio", "Chassi", "Potência e cc", "Etiqueta"]
    filtered_info = []

    # Dividir o texto em linhas
    lines = text.splitlines()

    # Procurar pelas palavras-chave nas linhas
    for line in lines:
        for keyword in keywords:
            if keyword.lower() in line.lower():  # Comparar em minúsculas para ignorar maiúsculas/minúsculas
                filtered_info.append(line.strip())
                break  # Parar de procurar por esta palavra-chave após encontrar

    return "\n".join(filtered_info)

def save_to_word(text, word_path):
    doc = Document()
    doc.add_heading('Informações Extraídas do PDF', 0)
    doc.add_paragraph(text)
    doc.save(word_path)

def main(pdf_path, word_path):
    try:
        extracted_text = extract_info_from_pdf(pdf_path)
        filtered_text = filter_info(extracted_text)
        save_to_word(filtered_text, word_path)
        print(f"Informações extraídas e salvas em {word_path}")
    except FileNotFoundError as e:
        print(e)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    return file_path

def select_word_file():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Arquivos Word", "*.docx")])
    return file_path

if __name__ == "__main__":
    try:
        pdf_path = select_pdf_file()
        if not pdf_path:
            raise ValueError("Nenhum arquivo PDF selecionado.")
        
        word_path = select_word_file()
        if not word_path:
            raise ValueError("Nenhum local de destino selecionado para o arquivo Word.")

        main(pdf_path, word_path)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")