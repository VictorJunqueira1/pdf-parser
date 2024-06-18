import fitz  # PyMuPDF
from docx import Document
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Text

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
    keywords = {
        "Veículo": "",
        "Placa": "",
        "Ano": "",
        "Cor": "",
        "Lote": "",  # Presente no exemplo fornecido
        "Nº motor": "",
        "Câmbio": "",
        "Chassi": "",
        "Potência e cc": "",
        "Etiqueta": ""  # Presente no exemplo fornecido
    }

    # Dividir o texto em linhas
    lines = text.splitlines()

    # Procurar pelas palavras-chave nas linhas
    for line in lines:
        for keyword in keywords.keys():
            if keyword.lower() in line.lower():  # Comparar em minúsculas para ignorar maiúsculas/minúsculas
                keywords[keyword] = line.strip()
                break  # Parar de procurar por esta palavra-chave após encontrar

    filtered_info = "\n".join([f"{k}: {v}" for k, v in keywords.items() if v])
    return filtered_info

def save_to_word(text, word_path):
    doc = Document()
    doc.add_heading('Informações Extraídas do PDF', 0)
    doc.add_paragraph(text)
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
        save_to_word(filtered_text, word_path)
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
