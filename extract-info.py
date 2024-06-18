import fitz  # PyMuPDF
from docx import Document
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Text
import re

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
    # Palavras-chave para procurar no texto e suas expressões regulares associadas
    patterns = {
        "Veículo": re.compile(r'Modelo\s+([\w\/\s\-]+)'),
        "Placa": re.compile(r'Placa\s+(\w+)'),
        "Ano": re.compile(r'Ano de Fabricação\s+(\d{4})'),
        "Cor": re.compile(r'Cor\s+([\w]+)'),
        "Lote": re.compile(r'Lote\s+(\d+)'),  # Placeholder, não parece existir no exemplo dado
        "Nº motor": re.compile(r'Nº Motor\s+([\w\-]+)'),
        "Câmbio": re.compile(r'Nº Câmbio\s+([\w\-]+)'),  # Placeholder, não parece existir no exemplo dado
        "Chassi": re.compile(r'Chassi\s+([\w]+)'),
        "Potência e cc": re.compile(r'Potência \(cv\)\s+(\d+)\s+Cilindradas \(cc\)\s+(\d+)'),
        "Etiqueta": re.compile(r'Etiqueta\s+(\w+)')  # Placeholder, não parece existir no exemplo dado
    }

    filtered_info = {}
    
    for key, pattern in patterns.items():
        match = pattern.search(text)
        if match:
            if key == "Potência e cc":
                filtered_info[key] = f"{match.group(1)} cv, {match.group(2)} cc"
            else:
                filtered_info[key] = match.group(1)
        else:
            filtered_info[key] = "Informação não encontrada"

    return "\n".join([f"{k}: {v}" for k, v in filtered_info.items()])

def save_to_word(text, word_path):
    doc = Document()
    doc.add_heading('Informações Extraídas do PDF', 0)
    
    for line in text.split('\n'):
        parts = line.split(": ")
        if len(parts) == 2:
            key, value = parts
            doc.add_heading(key, level=1)
            doc.add_paragraph(value)
        
    doc.save(word_path)

def process_pdf():
    try:
        pdf_path = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if not pdf_path:
            raise ValueError("Nenhum arquivo PDF selecionado.")

        extracted_text = extract_info_from_pdf(pdf_path)
        filtered_text = filter_info(extracted_text)
        
        text_display.delete(1.0, tk.END)
        text_display.insert(tk.END, filtered_text)
        
        save_button.config(state=tk.NORMAL)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def save_result():
    try:
        word_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Arquivos Word", "*.docx")])
        if not word_path:
            raise ValueError("Nenhum local de destino selecionado para o arquivo Word.")
        
        filtered_text = text_display.get(1.0, tk.END)
        save_to_word(filtered_text.strip(), word_path)
        messagebox.showinfo("Sucesso", f"Informações salvas em {word_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def main():
    global text_display, save_button

    root = tk.Tk()
    root.title("Extrator de Informações de PDF")

    frame = tk.Frame(root)
    frame.pack(pady=10)

    select_button = tk.Button(frame, text="Selecionar PDF", command=process_pdf)
    select_button.grid(row=0, column=0, padx=10)

    save_button = tk.Button(frame, text="Salvar como Word", command=save_result, state=tk.DISABLED)
    save_button.grid(row=0, column=1, padx=10)
    
    text_display = Text(root, wrap='word', width=80, height=20)
    text_display.pack(padx=10, pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()