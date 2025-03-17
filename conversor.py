import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
    entrada_arquivo.delete(0, tk.END)
    entrada_arquivo.insert(0, arquivo_path)

def converter():
    arquivo_path = entrada_arquivo.get()
    formato = formato_var.get()
    
    if not arquivo_path:
        messagebox.showerror("Erro", "Selecione um arquivo primeiro!")
        return

    try:
        df = pd.read_csv(arquivo_path, delimiter="\t", encoding="utf-8")
        novo_nome = os.path.splitext(arquivo_path)[0] + "." + formato
        
        if formato == "xlsx":
            df.to_excel(novo_nome, index=False, engine="openpyxl")
        elif formato == "csv":
            df.to_csv(novo_nome, index=False)
        elif formato == "json":
            df.to_json(novo_nome, orient="records", indent=4)
        else:
            messagebox.showerror("Erro", "Formato não suportado!")
            return
        
        messagebox.showinfo("Sucesso", f"Arquivo salvo como {novo_nome}")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão: {e}")

# Criando a interface gráfica
root = tk.Tk()
root.title("Conversor de Arquivos TXT")

# Layout
tk.Label(root, text="Arquivo TXT:").grid(row=0, column=0, padx=10, pady=10)
entrada_arquivo = tk.Entry(root, width=50)
entrada_arquivo.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=selecionar_arquivo).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Converter para:").grid(row=1, column=0, padx=10, pady=10)
formatos_disponiveis = ["xlsx", "csv", "json"]
formato_var = tk.StringVar(value="xlsx")
tk.OptionMenu(root, formato_var, *formatos_disponiveis).grid(row=1, column=1, padx=10, pady=10)

tk.Button(root, text="Converter", command=converter, bg="green", fg="white").grid(row=2, column=1, pady=20)

# Iniciar a interface
root.mainloop()
