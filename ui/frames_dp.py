# =============================================================================
# --- ARQUIVO: ui/frames_dp.py ---
# (Versão final com as duas telas de lançamento 100% funcionais)
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttk
from .dialogs_dp import (CompanyManagementWindow, EmployeeManagementWindow, 
                         PayrollCodeManagementWindow, EmployeeDialog, PayrollCodeDialog)
import dp_logic
import threading

# --- TELA 1: MENU PRINCIPAL DO MÓDULO DEPARTAMENTO PESSOAL ---
class DPMainFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        btn_voltar = ttk.Button(top_frame, text="< Voltar ao Menu Principal", command=lambda: self.controller.show_frame("HomeFrame"))
        btn_voltar.grid(row=0, column=0, sticky="w")
        
        title = ttk.Label(top_frame, text="Departamento Pessoal", font=("-size 18 -weight bold"))
        title.grid(row=0, column=1, pady=(0, 10))

        main_container = ttk.Frame(self, padding=50)
        main_container.grid(row=1, column=0, sticky="nsew")
        main_container.columnconfigure(0, weight=1)

        style = self.controller.app_style
        style.configure('Large.TButton', font=('Helvetica', 12))

        btn_lancamentos = ttk.Button(main_container, text="Lançamentos em Folha", style='Large.TButton', 
                                     command=lambda: self.controller.pulse_and_navigate(btn_lancamentos, "DPLançamentosToolFrame"))
        btn_lancamentos.pack(fill="x", ipady=15, pady=5)
        
        ttk.Button(main_container, text="(Futura Ferramenta DP)", style='Large.TButton', state="disabled").pack(fill="x", ipady=15, pady=5)


# --- TELA 2: FERRAMENTA DE LANÇAMENTOS EM FOLHA ---
class DPLançamentosToolFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        btn_voltar = ttk.Button(top_frame, text="< Voltar ao Menu DP", command=lambda: self.controller.show_frame("DPMainFrame"))
        btn_voltar.grid(row=0, column=0, sticky="w")
        
        title = ttk.Label(top_frame, text="Ferramenta de Lançamentos em Folha", font=("-size 18 -weight bold"))
        title.grid(row=0, column=1, pady=(0, 10))

        main_container = ttk.Frame(self, padding=20)
        main_container.grid(row=1, column=0, sticky="nsew")
        main_container.columnconfigure(0, weight=1)

        config_frame = ttk.Labelframe(main_container, text=" Configurações ", padding=20)
        config_frame.pack(fill="x", pady=(10, 20))
        config_frame.columnconfigure(0, weight=1)
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(2, weight=1)
        
        style = self.controller.app_style
        style.configure('Large.TButton', font=('Helvetica', 12))

        btn_empresas = ttk.Button(config_frame, text="Gerenciar Empresas", style='Large.TButton', command=self.open_company_management)
        btn_empresas.grid(row=0, column=0, sticky="ew", padx=5, ipady=10)

        btn_colaboradores = ttk.Button(config_frame, text="Gerenciar Colaboradores", style='Large.TButton', command=self.open_employee_management)
        btn_colaboradores.grid(row=0, column=1, sticky="ew", padx=5, ipady=10)
        
        btn_rubricas = ttk.Button(config_frame, text="Gerenciar Rubricas", style='Large.TButton', command=self.open_payroll_code_management)
        btn_rubricas.grid(row=0, column=2, sticky="ew", padx=5, ipady=10)

        operacional_frame = ttk.Labelframe(main_container, text=" Modos de Lançamento ", padding=20)
        operacional_frame.pack(fill="x", pady=10)
        operacional_frame.columnconfigure(0, weight=1)
        operacional_frame.columnconfigure(1, weight=1)

        btn_lanc_colab = ttk.Button(operacional_frame, text="Lançamento por Colaborador", style='Large.TButton', 
                                    command=lambda: self.controller.pulse_and_navigate(btn_lanc_colab, "DPLancamentoColabFrame"))
        btn_lanc_colab.grid(row=0, column=0, sticky="ew", padx=5, ipady=10)

        btn_lanc_rubrica = ttk.Button(operacional_frame, text="Lançamento por Rubrica", style='Large.TButton',
                                      command=lambda: self.controller.pulse_and_navigate(btn_lanc_rubrica, "DPLancamentoRubricaFrame"))
        btn_lanc_rubrica.grid(row=0, column=1, sticky="ew", padx=5, ipady=10)

    def open_company_management(self):
        if not hasattr(self, 'company_win') or not self.company_win.winfo_exists():
            self.company_win = CompanyManagementWindow(self.controller, self)
        else:
            self.company_win.lift()

    def open_employee_management(self):
        if not hasattr(self, 'employee_win') or not self.employee_win.winfo_exists():
            self.employee_win = EmployeeManagementWindow(self.controller, self)
        else:
            self.employee_win.lift()

    def open_payroll_code_management(self):
        if not hasattr(self, 'payroll_code_win') or not self.payroll_code_win.winfo_exists():
            self.payroll_code_win = PayrollCodeManagementWindow(self.controller, self)
        else:
            self.payroll_code_win.lift()


# --- TELA 3: TELA DE LANÇAMENTO POR COLABORADOR ---
class DPLancamentoColabFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        
        self.companies_map = {}
        self.all_employees = []
        self.all_payroll_codes = []
        self.launches_in_session = []
        self.selected_employee = None
        self.selected_payroll_code = None
        self.monthly_hours = 220.0
        
        self.create_widgets()
        self.bind_events()

    def create_widgets(self):
        header_frame = ttk.Labelframe(self, text=" Contexto do Lançamento ", padding=15)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.columnconfigure(1, weight=1)

        ttk.Label(header_frame, text="Empresa:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.company_combo = ttk.Combobox(header_frame, state="readonly")
        self.company_combo.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        ttk.Label(header_frame, text="Competência (MM/AAAA):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.competence_entry = ttk.Entry(header_frame)
        self.competence_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(header_frame, text="Tipo de Cálculo:").grid(row=1, column=2, sticky="w", padx=20, pady=5)
        self.calc_type_combo = ttk.Combobox(header_frame, state="readonly", values=["11 - Folha Mensal", "41 - Adiantamento"])
        self.calc_type_combo.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        self.calc_type_combo.set("11 - Folha Mensal")
        
        entry_frame = ttk.Labelframe(self, text=" Lançamento Individual ", padding=15)
        entry_frame.grid(row=1, column=0, sticky="ew", padx=10)
        entry_frame.columnconfigure(1, weight=1)
        entry_frame.columnconfigure(3, weight=1)

        ttk.Label(entry_frame, text="Cód. Colaborador (F2-Buscar, F7-Novo):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.employee_code_entry = ttk.Entry(entry_frame)
        self.employee_code_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        self.employee_name_label = ttk.Label(entry_frame, text="Nome do Colaborador...", bootstyle="secondary", width=40)
        self.employee_name_label.grid(row=0, column=2, columnspan=2, sticky="w", padx=5, pady=5)

        ttk.Label(entry_frame, text="Cód. Rubrica (F2-Buscar, F7-Novo):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.payroll_code_entry = ttk.Entry(entry_frame)
        self.payroll_code_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        self.payroll_code_name_label = ttk.Label(entry_frame, text="Nome da Rubrica...", bootstyle="secondary", width=40)
        self.payroll_code_name_label.grid(row=1, column=2, columnspan=2, sticky="w", padx=5, pady=5)

        self.value_label = ttk.Label(entry_frame, text="Valor:")
        self.value_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.value_var = tk.StringVar()
        self.value_entry = ttk.Entry(entry_frame, textvariable=self.value_var)
        self.value_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        launch_actions_frame = ttk.Frame(entry_frame)
        launch_actions_frame.grid(row=2, column=2, columnspan=2, sticky="ew")

        self.calculate_btn = ttk.Button(launch_actions_frame, text="Calcular", command=self.on_calculate_click, bootstyle="info-outline", state="disabled")
        self.calculate_btn.pack(side="left", padx=(5,10))

        self.add_launch_btn = ttk.Button(launch_actions_frame, text="Adicionar Lançamento", command=self.add_launch, bootstyle="success")
        self.add_launch_btn.pack(side="left")

        self.calculation_memo_label = ttk.Label(launch_actions_frame, text="", bootstyle="info")
        self.calculation_memo_label.pack(side="left", padx=15, fill="x", expand=True)

        tree_frame = ttk.Frame(self)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        
        cols = ('Cód. Colab.', 'Nome Colaborador', 'Cód. Rubrica', 'Nome Rubrica', 'Valor Lançado', 'Valor Calculado (R$)')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings')
        self.tree.pack(side="left", expand=True, fill="both")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        
        for col in cols: self.tree.heading(col, text=col)
        self.tree.column('Cód. Colab.', width=100, stretch=False, anchor="center")
        self.tree.column('Nome Colaborador', width=250)
        self.tree.column('Cód. Rubrica', width=100, stretch=False, anchor="center")
        self.tree.column('Nome Rubrica', width=250)
        self.tree.column('Valor Lançado', width=120, anchor="e")
        self.tree.column('Valor Calculado (R$)', width=150, anchor="e")

        action_frame = ttk.Frame(self, padding=10)
        action_frame.grid(row=3, column=0, sticky="ew")
        
        ttk.Button(action_frame, text="< Voltar", command=lambda: self.controller.show_frame("DPLançamentosToolFrame")).pack(side="left")
        
        self.generate_btn = ttk.Button(action_frame, text="Gerar Arquivo TXT", command=self.generate_file, bootstyle="primary")
        self.generate_btn.pack(side="right")
        ttk.Button(action_frame, text="Limpar Lançamentos", command=self.clear_launches, bootstyle="danger-outline").pack(side="right", padx=10)
    
    def bind_events(self):
        self.bind("<Visibility>", lambda e: self.on_frame_activated())
        self.company_combo.bind("<<ComboboxSelected>>", self.on_company_select)
        self.employee_code_entry.bind("<FocusOut>", self.validate_employee)
        self.employee_code_entry.bind("<Return>", self.validate_employee)
        self.payroll_code_entry.bind("<FocusOut>", self.validate_payroll_code)
        self.payroll_code_entry.bind("<Return>", self.validate_payroll_code)
        
        self.employee_code_entry.bind("<F2>", lambda e: self.open_management_window("employee"))
        self.employee_code_entry.bind("<F7>", lambda e: self.open_add_dialog("employee"))
        self.payroll_code_entry.bind("<F2>", lambda e: self.open_management_window("payroll_code"))
        self.payroll_code_entry.bind("<F7>", lambda e: self.open_add_dialog("payroll_code"))

    def on_frame_activated(self):
        self.load_initial_data()

    def load_initial_data(self):
        companies, error = self.controller.get_companies()
        if error: return
        self.companies_map = {c['name']: c for c in companies}
        self.company_combo['values'] = [c['name'] for c in companies]

        codes, error = self.controller.get_payroll_codes()
        if error: return
        self.all_payroll_codes = codes

    def on_company_select(self, event=None):
        self.employee_code_entry.delete(0, 'end')
        self.reset_employee_info()
        
        company_name = self.company_combo.get()
        company_id = self.companies_map[company_name]['id']
        threading.Thread(target=self._load_employees_thread, args=(company_id,), daemon=True).start()

    def _load_employees_thread(self, company_id):
        employees, error = self.controller.get_employees(company_id)
        if error: return
        self.all_employees = employees

    def validate_employee(self, event=None):
        code = self.employee_code_entry.get()
        if not code: 
            self.reset_employee_info()
            return
        
        found = next((emp for emp in self.all_employees if emp['employee_code'] == code), None)
        if found:
            self.employee_name_label.config(text=found['full_name'], bootstyle="primary")
            self.selected_employee = found
        else:
            self.employee_name_label.config(text="Colaborador não encontrado!", bootstyle="danger")
            self.selected_employee = None
        self.reset_payroll_code_info()
    
    def validate_payroll_code(self, event=None):
        code = self.payroll_code_entry.get()
        if not code:
            self.reset_payroll_code_info()
            return

        found = next((p_code for p_code in self.all_payroll_codes if p_code['code'] == code), None)
        if found:
            self.payroll_code_name_label.config(text=found['name'], bootstyle="primary")
            self.value_label.config(text=f"{found['value_type']}:")
            self.selected_payroll_code = found
            if found.get('calculation_base') != 'Valor Informado':
                self.calculate_btn.config(state="normal")
            else:
                self.calculate_btn.config(state="disabled")
        else:
            self.payroll_code_name_label.config(text="Rubrica não encontrada!", bootstyle="danger")
            self.selected_payroll_code = None
            self.calculate_btn.config(state="disabled")
        self.calculation_memo_label.config(text="")

    def on_calculate_click(self):
        if not self.selected_employee or not self.selected_payroll_code:
            messagebox.showwarning("Atenção", "Selecione um colaborador e uma rubrica válidos.", parent=self)
            return
        
        try:
            quantity = float(self.value_var.get().replace(',', '.'))
        except (ValueError, TypeError):
            messagebox.showerror("Valor Inválido", "A quantidade a ser calculada deve ser um número.", parent=self)
            return
        
        salary = self.selected_employee.get('salary', 0.0)
        base = self.selected_payroll_code.get('calculation_base')
        factor = self.selected_payroll_code.get('calculation_factor')
        
        calculated_value, memo = dp_logic.calculate_payroll_value(salary, base, factor, quantity, self.monthly_hours)
        
        self.calculation_memo_label.config(text=memo)

    def add_launch(self):
        if not all([self.company_combo.get(), self.competence_entry.get(), self.selected_employee, self.selected_payroll_code, self.value_entry.get()]):
            messagebox.showwarning("Campos Incompletos", "Todos os campos devem ser preenchidos e validados antes de adicionar.", parent=self)
            return

        quantity_to_launch = self.value_var.get().replace(",", ".")
        calculated_value_formatted = "N/A"
        
        if self.selected_payroll_code.get('calculation_base') != 'Valor Informado':
            try:
                quantity = float(quantity_to_launch)
                salary = self.selected_employee.get('salary', 0.0)
                base = self.selected_payroll_code.get('calculation_base')
                factor = self.selected_payroll_code.get('calculation_factor')
                calculated_value, _ = dp_logic.calculate_payroll_value(salary, base, factor, quantity, self.monthly_hours)
                calculated_value_formatted = f"R$ {calculated_value:.2f}"
            except (ValueError, TypeError):
                messagebox.showerror("Erro", "Valor inválido para cálculo.", parent=self)
                return

        launch_data = {
            'employee_code': self.selected_employee['employee_code'],
            'employee_name': self.selected_employee['full_name'],
            'payroll_code': self.selected_payroll_code,
            'value': quantity_to_launch 
        }
        self.launches_in_session.append(launch_data)
        
        self.tree.insert('', 'end', values=(
            launch_data['employee_code'],
            launch_data['employee_name'],
            launch_data['payroll_code']['code'],
            launch_data['payroll_code']['name'],
            quantity_to_launch,
            calculated_value_formatted
        ))
        
        self.payroll_code_entry.delete(0, 'end')
        self.value_entry.delete(0, 'end')
        self.reset_payroll_code_info()
        self.payroll_code_entry.focus_set()

    def clear_launches(self):
        if not self.launches_in_session: return
        if messagebox.askyesno("Confirmar", "Deseja realmente limpar todos os lançamentos da grade?", parent=self):
            for i in self.tree.get_children(): self.tree.delete(i)
            self.launches_in_session.clear()
            
    def generate_file(self):
        if not self.launches_in_session:
            messagebox.showwarning("Aviso", "Não há lançamentos para gerar o arquivo.", parent=self)
            return

        company_name = self.company_combo.get()
        company_data = self.companies_map.get(company_name)
        competence = self.competence_entry.get()
        calc_type_str = self.calc_type_combo.get()
        
        if not all([company_data, competence, calc_type_str]):
            messagebox.showwarning("Campos Incompletos", "Empresa, Competência e Tipo de Cálculo devem ser preenchidos.", parent=self)
            return
            
        company_code = company_data['code']
        calc_type = calc_type_str.split(" ")[0]

        filepath = filedialog.asksaveasfilename(
            title="Salvar Arquivo de Importação",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not filepath: return

        threading.Thread(target=self._generate_file_thread, args=(company_code, calc_type, competence, self.launches_in_session, filepath), daemon=True).start()

    def _generate_file_thread(self, company_code, calc_type, competence, launches, path):
        success, message = dp_logic.generate_import_file(company_code, calc_type, competence, launches, path)
        if success:
            self.after(0, lambda: messagebox.showinfo("Sucesso", message, parent=self))
        else:
            self.after(0, lambda: messagebox.showerror("Erro", message, parent=self))
    
    def reset_employee_info(self):
        self.employee_name_label.config(text="Nome do Colaborador...", bootstyle="secondary")
        self.selected_employee = None
        self.reset_payroll_code_info()

    def reset_payroll_code_info(self):
        self.payroll_code_name_label.config(text="Nome da Rubrica...", bootstyle="secondary")
        self.value_label.config(text="Valor:")
        self.selected_payroll_code = None
        self.calculate_btn.config(state="disabled")
        self.calculation_memo_label.config(text="")

    def open_management_window(self, item_type):
        company_id = self.companies_map.get(self.company_combo.get(), {}).get('id')
        if item_type == "employee":
            if not company_id: messagebox.showwarning("Aviso", "Selecione uma empresa primeiro.", parent=self); return
            win = EmployeeManagementWindow(self.controller, self)
        elif item_type == "payroll_code":
            win = PayrollCodeManagementWindow(self.controller, self)
        
        if win:
            self.wait_window(win)
            if item_type == "employee": 
                self.controller.invalidate_cache('employees', sub_key=company_id)
                self.on_company_select()
            if item_type == "payroll_code": 
                self.controller.invalidate_cache('payroll_codes')
                self.load_initial_data()

    def open_add_dialog(self, item_type):
        company_id = self.companies_map.get(self.company_combo.get(), {}).get('id')
        if item_type == "employee":
            if not company_id: messagebox.showwarning("Aviso", "Selecione uma empresa primeiro.", parent=self); return
            win = EmployeeDialog(self.controller, self, company_id=company_id)
        elif item_type == "payroll_code":
            win = PayrollCodeDialog(self.controller, self)

        if win:
            self.wait_window(win)
            if item_type == "employee":
                self.controller.invalidate_cache('employees', sub_key=company_id)
                self.on_company_select()
            if item_type == "payroll_code":
                self.controller.invalidate_cache('payroll_codes')
                self.load_initial_data()


# --- TELA 4: TELA DE LANÇAMENTO POR RUBRICA ---
class DPLancamentoRubricaFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        
        self.companies_map = {}
        self.all_employees = []
        self.all_payroll_codes = []
        self.launches_in_session = []
        self.selected_company_id = None
        self.selected_payroll_code = None
        self.selected_employee = None
        self.monthly_hours = 220.0

        self.create_widgets()
        self.bind_events()
    
    def create_widgets(self):
        header_frame = ttk.Labelframe(self, text=" Contexto do Lançamento ", padding=15)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.columnconfigure(1, weight=1)

        ttk.Label(header_frame, text="Empresa:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.company_combo = ttk.Combobox(header_frame, state="readonly")
        self.company_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(header_frame, text="Competência (MM/AAAA):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.competence_entry = ttk.Entry(header_frame)
        self.competence_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(header_frame, text="Tipo de Cálculo:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.calc_type_combo = ttk.Combobox(header_frame, state="readonly", values=["11 - Folha Mensal", "41 - Adiantamento"])
        self.calc_type_combo.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        self.calc_type_combo.set("11 - Folha Mensal")

        ttk.Label(header_frame, text="Rubrica (F2-Buscar, F7-Novo):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.payroll_code_combo = ttk.Combobox(header_frame, state="readonly")
        self.payroll_code_combo.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        
        entry_frame = ttk.Labelframe(self, text=" Lançamentos em Massa ", padding=15)
        entry_frame.grid(row=1, column=0, sticky="ew", padx=10)
        entry_frame.columnconfigure(1, weight=1)
        
        ttk.Label(entry_frame, text="Cód. Colaborador (F2/F7):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.employee_code_entry = ttk.Entry(entry_frame)
        self.employee_code_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.employee_name_label = ttk.Label(entry_frame, text="Nome...", bootstyle="secondary")
        self.employee_name_label.grid(row=0, column=2, sticky="w", padx=5, pady=5, columnspan=2)

        self.value_label = ttk.Label(entry_frame, text="Valor:")
        self.value_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.value_entry = ttk.Entry(entry_frame)
        self.value_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        self.add_launch_btn = ttk.Button(entry_frame, text="Adicionar", command=self.add_launch)
        self.add_launch_btn.grid(row=1, column=2, padx=5, pady=5)
        
        self.calculation_memo_label = ttk.Label(entry_frame, text="", bootstyle="info")
        self.calculation_memo_label.grid(row=1, column=3, sticky="w", padx=10)

        tree_frame = ttk.Frame(self)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        
        cols = ('Cód. Colab.', 'Nome Colaborador', 'Valor Lançado', 'Valor Calculado (R$)')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings')
        self.tree.pack(side="left", expand=True, fill="both")
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        
        for col in cols: self.tree.heading(col, text=col)
        self.tree.column('Cód. Colab.', width=100, stretch=False, anchor="center")
        self.tree.column('Nome Colaborador', width=300)
        self.tree.column('Valor Lançado', width=120, anchor="e")
        self.tree.column('Valor Calculado (R$)', width=150, anchor="e")

        action_frame = ttk.Frame(self, padding=10)
        action_frame.grid(row=3, column=0, sticky="ew")
        
        ttk.Button(action_frame, text="< Voltar", command=lambda: self.controller.show_frame("DPLançamentosToolFrame")).pack(side="left")
        
        self.generate_btn = ttk.Button(action_frame, text="Gerar Arquivo TXT", command=self.generate_file)
        self.generate_btn.pack(side="right")
        ttk.Button(action_frame, text="Limpar Lançamentos", command=self.clear_launches, bootstyle="danger-outline").pack(side="right", padx=10)

    def bind_events(self):
        self.bind("<Visibility>", lambda e: self.on_frame_activated())
        self.company_combo.bind("<<ComboboxSelected>>", self.on_company_select)
        self.payroll_code_combo.bind("<<ComboboxSelected>>", self.on_payroll_code_select)
        self.employee_code_entry.bind("<FocusOut>", self.validate_employee)
        self.employee_code_entry.bind("<Return>", lambda e: self.value_entry.focus_set())
        self.value_entry.bind("<Return>", lambda e: self.add_launch())
        
        self.employee_code_entry.bind("<F2>", lambda e: self.open_management_window("employee"))
        self.employee_code_entry.bind("<F7>", lambda e: self.open_add_dialog("employee"))

    def on_frame_activated(self):
        self.load_initial_data()

    def load_initial_data(self):
        companies, error = self.controller.get_companies()
        if error: return
        self.companies_map = {c['name']: c for c in companies}
        self.company_combo['values'] = [c['name'] for c in companies]

        codes, error = self.controller.get_payroll_codes()
        if error: return
        self.all_payroll_codes = codes
        self.payroll_code_combo['values'] = [f"{c['code']} - {c['name']}" for c in codes]
        
    def on_company_select(self, event=None):
        company_name = self.company_combo.get()
        company_id = self.companies_map[company_name]['id']
        self.selected_company_id = company_id
        threading.Thread(target=self._load_employees_thread, args=(company_id,), daemon=True).start()

    def _load_employees_thread(self, company_id):
        employees, error = self.controller.get_employees(company_id)
        if error: return
        self.all_employees = employees

    def on_payroll_code_select(self, event=None):
        selection = self.payroll_code_combo.get()
        if not selection: return
        code = selection.split(" ")[0]
        found = next((p_code for p_code in self.all_payroll_codes if p_code['code'] == code), None)
        self.selected_payroll_code = found
        if found:
            self.value_label.config(text=f"{found['value_type']}:")
        
    def validate_employee(self, event=None):
        code = self.employee_code_entry.get()
        if not code:
            self.employee_name_label.config(text="Nome...", bootstyle="secondary")
            self.selected_employee = None
            return
        
        found = next((emp for emp in self.all_employees if emp['employee_code'] == code), None)
        if found:
            self.employee_name_label.config(text=found['full_name'], bootstyle="primary")
            self.selected_employee = found
        else:
            self.employee_name_label.config(text="Não encontrado!", bootstyle="danger")
            self.selected_employee = None

    def add_launch(self):
        if not all([self.company_combo.get(), self.competence_entry.get(), self.selected_employee, self.selected_payroll_code, self.value_entry.get()]):
            messagebox.showwarning("Campos Incompletos", "Todos os campos de contexto, colaborador e valor devem ser preenchidos.", parent=self)
            return

        quantity_to_launch = self.value_entry.get().replace(",", ".")
        calculated_value_formatted = "N/A"
        
        if self.selected_payroll_code.get('calculation_base') != 'Valor Informado':
            try:
                quantity = float(quantity_to_launch)
                salary = self.selected_employee.get('salary', 0.0)
                base = self.selected_payroll_code.get('calculation_base')
                factor = self.selected_payroll_code.get('calculation_factor')
                calculated_value, _ = dp_logic.calculate_payroll_value(salary, base, factor, quantity, self.monthly_hours)
                calculated_value_formatted = f"R$ {calculated_value:.2f}"
            except (ValueError, TypeError):
                messagebox.showerror("Erro", "Valor inválido para cálculo.", parent=self)
                return

        launch_data = {
            'employee_code': self.selected_employee['employee_code'],
            'employee_name': self.selected_employee['full_name'],
            'payroll_code': self.selected_payroll_code,
            'value': quantity_to_launch
        }
        self.launches_in_session.append(launch_data)
        
        self.tree.insert('', 'end', values=(
            launch_data['employee_code'],
            launch_data['employee_name'],
            quantity_to_launch,
            calculated_value_formatted
        ))
        
        self.employee_code_entry.delete(0, 'end')
        self.value_entry.delete(0, 'end')
        self.employee_code_entry.focus_set()
        self.employee_name_label.config(text="Nome...", bootstyle="secondary")
        self.selected_employee = None

    def clear_launches(self):
        if not self.launches_in_session: return
        if messagebox.askyesno("Confirmar", "Deseja realmente limpar todos os lançamentos da grade?", parent=self):
            for i in self.tree.get_children(): self.tree.delete(i)
            self.launches_in_session.clear()
            
    def generate_file(self):
        if not self.launches_in_session:
            messagebox.showwarning("Aviso", "Não há lançamentos para gerar o arquivo.", parent=self)
            return

        company_name = self.company_combo.get()
        company_data = self.companies_map.get(company_name)
        competence = self.competence_entry.get()
        calc_type_str = self.calc_type_combo.get()
        
        if not all([company_data, competence, calc_type_str, self.selected_payroll_code]):
            messagebox.showwarning("Campos Incompletos", "Empresa, Competência, Tipo de Cálculo e Rubrica devem ser preenchidos.", parent=self)
            return
            
        company_code = company_data['code']
        calc_type = calc_type_str.split(" ")[0]

        filepath = filedialog.asksaveasfilename(
            title="Salvar Arquivo de Importação",
            defaultextension=".txt",
            filetypes=[("Arquivo de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not filepath: return

        threading.Thread(target=self._generate_file_thread, args=(company_code, calc_type, competence, self.launches_in_session, filepath), daemon=True).start()

    def _generate_file_thread(self, company_code, calc_type, competence, launches, path):
        success, message = dp_logic.generate_import_file(company_code, calc_type, competence, launches, path)
        if success:
            self.after(0, lambda: messagebox.showinfo("Sucesso", message, parent=self))
        else:
            self.after(0, lambda: messagebox.showerror("Erro", message, parent=self))

    def open_management_window(self, item_type):
        # Apenas a janela de colaboradores pode ser aberta a partir daqui
        if item_type == "employee":
            if not self.selected_company_id:
                messagebox.showwarning("Aviso", "Selecione uma empresa primeiro.", parent=self)
                return
            win = EmployeeManagementWindow(self.controller, self)
            if win:
                self.wait_window(win)
                self.controller.invalidate_cache('employees', sub_key=self.selected_company_id)
                self.on_company_select()

    def open_add_dialog(self, item_type):
        if item_type == "employee":
            if not self.selected_company_id:
                messagebox.showwarning("Aviso", "Selecione uma empresa primeiro.", parent=self)
                return
            win = EmployeeDialog(self.controller, self, company_id=self.selected_company_id)
            if win:
                self.wait_window(win)
                self.controller.invalidate_cache('employees', sub_key=self.selected_company_id)
                self.on_company_select()