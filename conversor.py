import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
from pdf2docx import Converter
import os

def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename(filetypes=[("Arquivos suportados", "*.txt;*.pdf")])
    entrada_arquivo.delete(0, tk.END)
    entrada_arquivo.insert(0, arquivo_path)
    
    # Atualiza as opções com base no tipo de arquivo
    atualizar_opcoes(os.path.splitext(arquivo_path)[1].lower())

def atualizar_opcoes(extensao):
    menu_formato["menu"].delete(0, "end")  # Limpa o menu de opções
    
    if extensao == ".txt":
        formatos = ["xlsx", "csv", "json"]
    elif extensao == ".pdf":
        formatos = ["docx", "xlsx"]
    else:
        formatos = []
    
    if formatos:
        formato_var.set(formatos[0])  # Define a primeira opção como padrão
        for formato in formatos:
            menu_formato["menu"].add_command(label=formato, command=lambda f=formato: formato_var.set(f))
    else:
        formato_var.set("")

def converter():
    arquivo_path = entrada_arquivo.get()
    formato = formato_var.get()

    if not arquivo_path:
        messagebox.showerror("Erro", "Selecione um arquivo primeiro!")
        return
    
    extensao = os.path.splitext(arquivo_path)[1].lower()

    try:
        novo_nome = os.path.splitext(arquivo_path)[0] + "." + formato
        
        if extensao == ".txt":
            df = pd.read_csv(arquivo_path, delimiter="\t", encoding="utf-8")
            if formato == "xlsx":
                df.to_excel(novo_nome, index=False, engine="openpyxl")
            elif formato == "csv":
                df.to_csv(novo_nome, index=False)
            elif formato == "json":
                df.to_json(novo_nome, orient="records", indent=4)
        
        elif extensao == ".pdf":
            if formato == "docx":
                cv = Converter(arquivo_path)
                cv.convert(novo_nome, start=0, end=None)
                cv.close()
            elif formato == "xlsx":
                with pdfplumber.open(arquivo_path) as pdf:
                    tabelas = []
                    for pagina in pdf.pages:
                        for tabela in pagina.extract_tables():
                            tabelas.extend(tabela)
                    if tabelas:
                        df = pd.DataFrame(tabelas)
                        df.to_excel(novo_nome, index=False, engine="openpyxl")
                    else:
                        messagebox.showerror("Erro", "Nenhuma tabela encontrada no PDF!")
                        return
        
        messagebox.showinfo("Sucesso", f"Arquivo salvo como {novo_nome}")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão: {e}")

# Criando a interface gráfica
root = tk.Tk()
root.title("Conversor de Arquivos")

# Layout
tk.Label(root, text="Arquivo:").grid(row=0, column=0, padx=10, pady=10)
entrada_arquivo = tk.Entry(root, width=50)
entrada_arquivo.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=selecionar_arquivo).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Converter para:").grid(row=1, column=0, padx=10, pady=10)
formato_var = tk.StringVar(value="")
menu_formato = tk.OptionMenu(root, formato_var, "")
menu_formato.grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Converter", command=converter, bg="green", fg="white").grid(row=2, column=1, pady=20)

# Iniciar a interface
root.mainloop()