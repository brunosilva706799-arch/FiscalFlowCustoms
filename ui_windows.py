# =============================================================================
# --- ARQUIVO: ui_windows.py ---
# (Contém as classes das Janelas de Ferramentas e Utilitários - Toplevels)
# =============================================================================

import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import ttkbootstrap as ttk

class SettingsWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Configurações Gerais", master=controller)
        self.controller = controller
        self.bind("<Escape>", lambda e: self.destroy())
        self.geometry("700x400")
        self.transient(controller)
        self.grab_set()

        self.xml_path_var = tk.StringVar(value=self.controller.default_xml_path)
        self.output_path_var = tk.StringVar(value=self.controller.default_output_path)
        self.filename_pattern_var = tk.StringVar(value=self.controller.output_filename_pattern)
        self.confirm_on_exit_var = tk.BooleanVar(value=self.controller.confirm_on_exit)
        self.ask_to_open_excel_var = tk.BooleanVar(value=self.controller.ask_to_open_excel)
        
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        path_frame = ttk.Labelframe(main_frame, text=" Pastas e Arquivos Padrão ", padding=15)
        path_frame.pack(fill="x")
        path_frame.columnconfigure(1, weight=1)
        
        ttk.Label(path_frame, text="Pasta para ler XMLs:").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        xml_entry_frame = ttk.Frame(path_frame)
        xml_entry_frame.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        xml_entry_frame.columnconfigure(0, weight=1)
        ttk.Entry(xml_entry_frame, textvariable=self.xml_path_var).grid(row=0, column=0, sticky='ew')
        ttk.Button(xml_entry_frame, text="Procurar...", command=self.select_xml_path).grid(row=0, column=1, padx=(5,0))
        
        ttk.Label(path_frame, text="Pasta para salvar Excel:").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        output_entry_frame = ttk.Frame(path_frame)
        output_entry_frame.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        output_entry_frame.columnconfigure(0, weight=1)
        ttk.Entry(output_entry_frame, textvariable=self.output_path_var).grid(row=0, column=0, sticky='ew')
        ttk.Button(output_entry_frame, text="Procurar...", command=self.select_output_path).grid(row=0, column=1, padx=(5,0))

        ttk.Label(path_frame, text="Padrão para nome do arquivo:").grid(row=2, column=0, sticky='w', pady=5, padx=5)
        ttk.Entry(path_frame, textvariable=self.filename_pattern_var, width=60).grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(path_frame, text="Use {data} para a data atual (Ex: AAAA-MM-DD)", bootstyle="secondary").grid(row=3, column=1, sticky='w', padx=5)
        
        prefs_frame = ttk.Labelframe(main_frame, text=" Comportamento do Programa ", padding=15)
        prefs_frame.pack(fill="x", pady=20)
        
        ttk.Checkbutton(prefs_frame, text="Perguntar antes de sair do programa", variable=self.confirm_on_exit_var).pack(anchor='w', padx=5)
        ttk.Checkbutton(prefs_frame, text="Perguntar para abrir o Excel após salvar", variable=self.ask_to_open_excel_var).pack(anchor='w', padx=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Salvar e Fechar", command=self.save_and_close, bootstyle="primary").pack(side='left', padx=10)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side='left')

    def select_xml_path(self):
        folder_path = filedialog.askdirectory(title="Selecione a Pasta Padrão para XMLs")
        if folder_path:
            self.xml_path_var.set(folder_path)
            
    def select_output_path(self):
        folder_path = filedialog.askdirectory(title="Selecione a Pasta Padrão para Salvar Planilhas")
        if folder_path:
            self.output_path_var.set(folder_path)

    def save_and_close(self):
        self.controller.default_xml_path = self.xml_path_var.get()
        self.controller.default_output_path = self.output_path_var.get()
        self.controller.output_filename_pattern = self.filename_pattern_var.get()
        self.controller.confirm_on_exit = self.confirm_on_exit_var.get()
        self.controller.ask_to_open_excel = self.ask_to_open_excel_var.get()
        
        self.controller.save_config()
        messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!", parent=self)
        self.destroy()

class ValidatorWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Validador de CNPJ/CPF", master=controller)
        self.bind("<Escape>", lambda e: self.destroy())
        self.geometry("400x200")
        self.transient(controller)
        self.grab_set()
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        
        self.doc_var = tk.StringVar()
        self.result_var = tk.StringVar(value="Aguardando número...")
        
        ttk.Label(main_frame, text="Digite o CNPJ ou CPF (apenas números):").pack(pady=5)
        entry = ttk.Entry(main_frame, textvariable=self.doc_var, width=30)
        entry.pack(pady=5)
        entry.focus_set()
        entry.bind("<Return>", lambda e: self.validate())
        
        ttk.Button(main_frame, text="Validar", command=self.validate, bootstyle="primary").pack(pady=10)
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        self.result_label = ttk.Label(main_frame, textvariable=self.result_var, font=("Helvetica", 10, "bold"))
        self.result_label.pack(pady=5)

    def validate(self):
        doc = re.sub(r'[^0-9]', '', self.doc_var.get())
        if len(doc) == 11 and self._validate_cpf(doc):
            self.result_var.set("CPF Válido!")
            self.result_label.config(bootstyle="success")
        elif len(doc) == 14 and self._validate_cnpj(doc):
            self.result_var.set("CNPJ Válido!")
            self.result_label.config(bootstyle="success")
        else:
            self.result_var.set("Número Inválido!")
            self.result_label.config(bootstyle="danger")

    def _validate_cpf(self, cpf):
        if len(cpf) != 11 or len(set(cpf)) == 1: return False
        for i in range(9, 11):
            value = sum((int(cpf[num]) * ((i + 1) - num) for num in range(0, i)))
            digit = ((value * 10) % 11) % 10
            if str(digit) != cpf[i]: return False
        return True

    def _validate_cnpj(self, cnpj):
        if len(cnpj) != 14 or len(set(cnpj)) == 1: return False
        for i in range(12, 14):
            if i == 12: weights = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
            else: weights = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
            value = sum(int(cnpj[j]) * weights[j] for j in range(i))
            digit = 11 - (value % 11)
            if digit >= 10: digit = 0
            if str(digit) != cnpj[i]: return False
        return True

class KeyParserWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Analisador de Chave de Acesso NFe", master=controller)
        self.bind("<Escape>", lambda e: self.destroy())
        self.geometry("500x400")
        self.transient(controller)
        self.grab_set()
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        self.key_var = tk.StringVar()
        ttk.Label(main_frame, text="Cole a Chave de Acesso (44 dígitos):").grid(row=0, column=0, sticky='w')
        entry_frame = ttk.Frame(main_frame)
        entry_frame.grid(row=1, column=0, sticky='ew', pady=5)
        entry_frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(entry_frame, textvariable=self.key_var)
        entry.grid(row=0, column=0, sticky='ew', padx=(0,10))
        entry.focus_set()
        entry.bind("<Return>", lambda e: self.parse_key())
        
        ttk.Button(entry_frame, text="Analisar", command=self.parse_key, bootstyle="primary").grid(row=0, column=1)
        
        self.text_area = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=10, font=("Courier New", 10))
        self.text_area.grid(row=2, column=0, sticky="nsew", pady=10)
        self.text_area.insert(tk.END, "Aguardando chave para análise...")
        self.text_area.config(state='disabled')
        
    def parse_key(self):
        key = re.sub(r'[^0-9]', '', self.key_var.get())
        self.text_area.config(state='normal')
        self.text_area.delete('1.0', tk.END)
        if len(key) != 44:
            self.text_area.insert(tk.END, "Erro: A chave de acesso deve conter 44 dígitos.")
            self.text_area.config(state='disabled')
            return
        
        parts = {
            "UF do Emitente (cUF)": key[0:2],
            "Ano e Mês de Emissão (AAMM)": key[2:6],
            "CNPJ do Emitente": key[6:20],
            "Modelo (mod)": key[20:22],
            "Série (serie)": key[22:25],
            "Número da NF (nNF)": key[25:34],
            "Forma de Emissão (tpEmis)": key[34:35],
            "Código Numérico (cNF)": key[35:43],
            "Dígito Verificador (cDV)": key[43:44]
        }
        
        result_text = "--- Análise da Chave de Acesso ---\n\n"
        for label, value in parts.items():
            result_text += f"{label:<30}: {value}\n"
        
        self.text_area.insert(tk.END, result_text)
        self.text_area.config(state='disabled')

# =============================================================================
# --- ARQUIVO: ui_windows.py: FIM ---
# =============================================================================