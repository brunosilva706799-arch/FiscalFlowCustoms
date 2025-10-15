# =============================================================================
# --- ARQUIVO: ui/dialogs_dp.py ---
# (Atualizado para usar o cache central em todas as janelas)
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttk
import dp_logic # Importa nossa nova lógica de DP
import threading

class CompanyDialog(ttk.Toplevel):
    """Janela de formulário para Adicionar ou Editar uma Empresa."""
    def __init__(self, controller, parent_window, company_to_edit=None):
        title = "Editar Empresa" if company_to_edit else "Adicionar Nova Empresa"
        super().__init__(title=title, master=parent_window)
        self.controller = controller
        self.parent_window = parent_window
        self.company_to_edit = company_to_edit
        
        self.transient(parent_window); self.grab_set()

        self.code_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.cnpj_var = tk.StringVar()
        
        self.create_widgets()

        if self.company_to_edit:
            self.load_company_data()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Código da Empresa:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.code_var).grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(main_frame, text="Razão Social:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.name_var).grid(row=1, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(main_frame, text="CNPJ:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.cnpj_var).grid(row=2, column=1, sticky="ew", pady=5, padx=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(20,0))
        ttk.Button(btn_frame, text="Salvar", command=self.save_company, bootstyle="primary").pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def load_company_data(self):
        self.code_var.set(self.company_to_edit['code'])
        self.name_var.set(self.company_to_edit['name'])
        self.cnpj_var.set(self.company_to_edit['cnpj'])

    def save_company(self):
        code = self.code_var.get()
        name = self.name_var.get()
        cnpj = self.cnpj_var.get()

        if not all([code, name, cnpj]):
            messagebox.showerror("Erro de Validação", "Todos os campos são obrigatórios.", parent=self)
            return

        if self.company_to_edit:
            result = dp_logic.update_company(self.company_to_edit['id'], code, name, cnpj)
        else:
            result = dp_logic.add_company(code, name, cnpj)
        
        if "sucesso" in result:
            messagebox.showinfo("Sucesso", result, parent=self.parent_window)
            self.controller.invalidate_cache('companies')
            self.parent_window.refresh_company_list()
            self.destroy()
        else:
            messagebox.showerror("Erro", result, parent=self)


class CompanyManagementWindow(ttk.Toplevel):
    """Janela principal para Gerenciar Empresas (listar, adicionar, editar, remover)."""
    def __init__(self, controller, parent_frame):
        super().__init__(title="Gerenciar Empresas", master=controller)
        self.controller = controller
        self.parent_frame = parent_frame
        self.minsize(700, 400); self.grab_set()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")

        cols = ('ID', 'Código', 'Razão Social', 'CNPJ')
        self.tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        self.pack_propagate(False) 
        self.tree.pack(expand=True, fill="both")

        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=80, stretch=False)
        self.tree.heading('Razão Social', text='Razão Social')
        self.tree.heading('CNPJ', text='CNPJ')
        self.tree.column('ID', width=0, stretch=False)

        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0))
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Adicionar Nova", command=self.add_new_company, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Editar Selecionada", command=self.edit_selected_company).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Remover Selecionada", command=self.delete_selected_company, bootstyle="danger").pack(side="left")
        ttk.Button(btn_frame, text="Atualizar Lista", command=lambda: self.refresh_company_list(force=True), bootstyle="secondary-outline").pack(side="right")

        self.refresh_company_list()

    def refresh_company_list(self, force=False):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        companies, error = self.controller.get_companies(force_refresh=force)
        if error: return
        
        for company in companies:
            self.tree.insert('', 'end', values=(company['id'], company['code'], company['name'], company['cnpj']))

    def add_new_company(self):
        CompanyDialog(self.controller, self)

    def edit_selected_company(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma empresa na lista para editar.", parent=self)
            return
        
        item_values = self.tree.item(selected_item, 'values')
        company_data = {
            'id': item_values[0],
            'code': item_values[1],
            'name': item_values[2],
            'cnpj': item_values[3]
        }
        CompanyDialog(self.controller, self, company_to_edit=company_data)

    def delete_selected_company(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma empresa na lista para remover.", parent=self)
            return
            
        item_values = self.tree.item(selected_item, 'values')
        company_id = item_values[0]
        company_name = item_values[2]

        if messagebox.askyesno("Confirmar Remoção", f"Tem certeza que deseja remover a empresa '{company_name}'?", parent=self):
            result = dp_logic.delete_company(company_id)
            if "sucesso" in result:
                messagebox.showinfo("Sucesso", result, parent=self)
                self.controller.invalidate_cache('companies')
                self.refresh_company_list()
            else:
                messagebox.showerror("Erro", result, parent=self)


class EmployeeDialog(ttk.Toplevel):
    """Janela de formulário para Adicionar ou Editar um Colaborador."""
    def __init__(self, controller, parent_window, company_id, employee_to_edit=None):
        title = "Editar Colaborador" if employee_to_edit else "Adicionar Novo Colaborador"
        super().__init__(title=title, master=parent_window)
        
        self.controller = controller
        self.parent_window = parent_window
        self.company_id = company_id
        self.employee_to_edit = employee_to_edit
        
        self.transient(parent_window); self.grab_set()

        self.code_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.salary_var = tk.StringVar()
        
        self.create_widgets()

        if self.employee_to_edit:
            self.load_employee_data()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Código (Folha):").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.code_var).grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(main_frame, text="Nome Completo:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.name_var).grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        ttk.Label(main_frame, text="Salário Contratual:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        self.salary_entry = ttk.Entry(main_frame, textvariable=self.salary_var)
        self.salary_entry.grid(row=2, column=1, sticky="ew", pady=5, padx=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(20,0))
        ttk.Button(btn_frame, text="Salvar", command=self.save_employee, bootstyle="primary").pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def load_employee_data(self):
        self.code_var.set(self.employee_to_edit['employee_code'])
        self.name_var.set(self.employee_to_edit['full_name'])
        self.salary_var.set(f"{self.employee_to_edit.get('salary', 0.0):.2f}")

    def save_employee(self):
        code = self.code_var.get()
        name = self.name_var.get()
        salary_str = self.salary_var.get().replace(",", ".")

        if not all([code, name, salary_str]):
            messagebox.showerror("Erro de Validação", "Todos os campos são obrigatórios.", parent=self)
            return
        
        try:
            salary_float = float(salary_str)
        except ValueError:
            messagebox.showerror("Erro de Validação", "O valor do salário deve ser um número válido.", parent=self)
            return

        if self.employee_to_edit:
            result = dp_logic.update_employee(self.employee_to_edit['id'], code, name, salary_float, self.company_id)
        else:
            result = dp_logic.add_employee(self.company_id, code, name, salary_float)
        
        if "sucesso" in result:
            messagebox.showinfo("Sucesso", result, parent=self.parent_window)
            self.controller.invalidate_cache('employees', sub_key=self.company_id)
            self.parent_window.refresh_employee_list()
            self.destroy()
        else:
            messagebox.showerror("Erro", result, parent=self)


class EmployeeImportPreviewDialog(ttk.Toplevel):
    def __init__(self, controller, parent_window, company_id, employee_data):
        super().__init__(title="Pré-visualização da Importação", master=parent_window)
        self.controller = controller
        self.parent_window = parent_window
        self.company_id = company_id
        self.employee_data = employee_data
        
        self.minsize(800, 500)
        self.transient(parent_window); self.grab_set()

        self.create_widgets()
        self.populate_tree()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")

        instructions = "Abaixo estão os colaboradores encontrados. Dê um duplo clique para editar. Registros 'Ignorados' não serão importados."
        ttk.Label(main_frame, text=instructions, wraplength=780, justify="left").pack(fill="x", pady=(0, 10))

        cols = ('Código', 'Nome', 'Salário', 'Status')
        self.tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        self.tree.pack(expand=True, fill="both")

        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=100, stretch=False)
        self.tree.heading('Nome', text='Nome Completo')
        self.tree.heading('Salário', text='Salário')
        self.tree.column('Salário', width=100, stretch=False, anchor="e")
        self.tree.heading('Status', text='Status')
        self.tree.column('Status', width=150, stretch=False, anchor="center")
        
        self.tree.tag_configure("Novo", foreground="green")
        self.tree.tag_configure("Ignorado", foreground="gray")
        self.tree.bind("<Double-1>", self.on_double_click)

        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0))
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="Excluir Linha Selecionada", command=self.delete_selected_row, bootstyle="danger-outline").pack(side="left")
        
        self.confirm_btn = ttk.Button(btn_frame, text="Confirmar e Importar", command=self.confirm_import, bootstyle="success")
        self.confirm_btn.pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def populate_tree(self):
        for item in self.employee_data:
            tag = item['status'].split(" ")[0]
            salary_formatted = f"{item.get('salary', 0.0):.2f}"
            self.tree.insert('', 'end', values=(item['employee_code'], item['full_name'], salary_formatted, item['status']), tags=(tag,))

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return

        selected_item = self.tree.focus()
        column = self.tree.identify_column(event.x)
        column_index = int(column.replace("#", "")) - 1

        if column_index > 2: return

        x, y, width, height = self.tree.bbox(selected_item, column)
        
        entry_var = tk.StringVar()
        entry = ttk.Entry(self.tree, textvariable=entry_var)
        entry.place(x=x, y=y, width=width, height=height)
        
        current_value = self.tree.item(selected_item, "values")[column_index]
        entry_var.set(current_value)
        entry.focus_set()
        entry.selection_range(0, 'end')

        def on_focus_out(event):
            self.update_tree_value(selected_item, column_index, entry_var.get())
            entry.destroy()

        def on_enter_key(event):
            self.update_tree_value(selected_item, column_index, entry_var.get())
            entry.destroy()

        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", on_enter_key)

    def update_tree_value(self, item, column_index, new_value):
        current_values = list(self.tree.item(item, "values"))
        current_values[column_index] = new_value
        self.tree.item(item, values=tuple(current_values))
    
    def delete_selected_row(self):
        selected_item = self.tree.focus()
        if selected_item:
            self.tree.delete(selected_item)

    def confirm_import(self):
        final_data_to_import = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, 'values')
            try:
                salary = float(str(values[2]).replace(",", "."))
            except:
                salary = 0.0
            final_data_to_import.append({
                'employee_code': values[0],
                'full_name': values[1],
                'salary': salary,
                'status': values[3]
            })

        self.confirm_btn.config(state="disabled", text="Importando...")
        self.update_idletasks()

        threading.Thread(target=self._confirm_import_thread, args=(self.company_id, final_data_to_import), daemon=True).start()

    def _confirm_import_thread(self, company_id, data):
        success, message = dp_logic.save_imported_employees(company_id, data)
        self.after(0, self.on_import_complete, success, message)

    def on_import_complete(self, success, message):
        self.confirm_btn.config(state="normal", text="Confirmar e Importar")
        if success:
            messagebox.showinfo("Sucesso", message, parent=self.parent_window)
            self.controller.invalidate_cache('employees', sub_key=self.company_id)
            self.parent_window.refresh_employee_list()
            self.destroy()
        else:
            messagebox.showerror("Erro", message, parent=self)


class EmployeeManagementWindow(ttk.Toplevel):
    def __init__(self, controller, parent_frame):
        super().__init__(title="Gerenciar Colaboradores", master=controller)
        self.controller = controller
        self.parent_frame = parent_frame
        self.minsize(800, 500); self.grab_set()
        
        self.selected_company_id = tk.StringVar()
        self.companies_map = {}
        self.delete_mode = False
        self.selection_state = {}
        self.all_employees_list = [] # Cache local para a busca

        self.create_widgets()

    def create_widgets(self):
        self.main_frame = ttk.Frame(self, padding=10)
        self.main_frame.pack(expand=True, fill="both")
        
        company_frame = ttk.Frame(self.main_frame)
        company_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(company_frame, text="Selecione a Empresa:", font="-weight bold").pack(side="left", padx=(0, 10))
        
        self.company_combo = ttk.Combobox(company_frame, state="readonly", width=50)
        self.company_combo.pack(side="left", fill="x", expand=True)
        self.populate_company_combo()
        self.company_combo.bind("<<ComboboxSelected>>", self.on_company_select)
        
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree_frame.pack(expand=True, fill="both")
        self.setup_treeview_normal_mode()

        self.normal_buttons_frame = ttk.Frame(self.main_frame, padding=(0,10,0,0))
        self.normal_buttons_frame.pack(fill="x")
        self.add_btn = ttk.Button(self.normal_buttons_frame, text="Adicionar Novo", command=self.add_new_employee, bootstyle="success", state="disabled")
        self.add_btn.pack(side="left")
        self.edit_btn = ttk.Button(self.normal_buttons_frame, text="Editar Selecionado", command=self.edit_selected_employee, state="disabled")
        self.edit_btn.pack(side="left", padx=10)
        self.delete_btn = ttk.Button(self.normal_buttons_frame, text="Remover", command=self.enter_delete_mode, bootstyle="danger", state="disabled")
        self.delete_btn.pack(side="left")
        self.import_btn = ttk.Button(self.normal_buttons_frame, text="Importar de Planilha...", command=self.import_from_file, bootstyle="info", state="disabled")
        self.import_btn.pack(side="left", padx=10)

        self.delete_buttons_frame = ttk.Frame(self.main_frame, padding=(0,10,0,0))
        ttk.Button(self.delete_buttons_frame, text="Selecionar Todos", command=self.select_all, bootstyle="secondary-outline").pack(side="left")
        ttk.Button(self.delete_buttons_frame, text="Desmarcar Todos", command=self.deselect_all, bootstyle="secondary-outline").pack(side="left", padx=5)
        ttk.Button(self.delete_buttons_frame, text="Inverter Seleção", command=self.invert_selection, bootstyle="secondary-outline").pack(side="left")
        self.confirm_delete_btn = ttk.Button(self.delete_buttons_frame, text="Confirmar Exclusão", command=self.confirm_delete, bootstyle="success")
        self.confirm_delete_btn.pack(side="right")
        ttk.Button(self.delete_buttons_frame, text="Cancelar", command=self.exit_delete_mode, bootstyle="danger-outline").pack(side="right", padx=10)

    def setup_treeview_normal_mode(self):
        if hasattr(self, 'tree'): self.tree.destroy()
        cols = ('ID', 'Código', 'Nome Completo', 'Salário')
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        self.tree.pack(expand=True, fill="both")
        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=100, stretch=False)
        self.tree.heading('Nome Completo', text='Nome Completo')
        self.tree.heading('Salário', text='Salário')
        self.tree.column('Salário', width=120, stretch=False, anchor="e")
        self.tree.column('ID', width=0, stretch=False)
        self.tree.unbind("<Button-1>")

    def setup_treeview_delete_mode(self):
        if hasattr(self, 'tree'): self.tree.destroy()
        cols = ('ID', 'Código', 'Nome Completo', 'Salário')
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show="tree headings")
        self.tree.pack(expand=True, fill="both")
        
        self.tree.column("#0", width=40, stretch=False, anchor='center')
        self.tree.heading("#0", text="Sel.")
        
        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=100, stretch=False)
        self.tree.heading('Nome Completo', text='Nome Completo')
        self.tree.heading('Salário', text='Salário')
        self.tree.column('Salário', width=120, stretch=False, anchor="e")
        self.tree.column('ID', width=0, stretch=False)
        
        style = ttk.Style()
        selected_bg = style.colors.selectbg
        self.tree.tag_configure("selected", background=selected_bg)
        
        self.tree.bind("<Button-1>", self.toggle_checkbox)

    def enter_delete_mode(self):
        if not self.selected_company_id.get(): return
        self.delete_mode = True
        self.normal_buttons_frame.pack_forget()
        self.delete_buttons_frame.pack(fill="x")
        self.setup_treeview_delete_mode()
        self.refresh_employee_list(force=True)

    def exit_delete_mode(self):
        self.delete_mode = False
        self.delete_buttons_frame.pack_forget()
        self.normal_buttons_frame.pack(fill="x")
        self.setup_treeview_normal_mode()
        self.refresh_employee_list()

    def toggle_checkbox(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id or self.tree.identify_column(event.x) != "#0":
            return
        
        self.selection_state[item_id] = not self.selection_state.get(item_id, False)
        self.update_checkbox(item_id)
        
    def update_checkbox(self, item_id):
        if self.selection_state.get(item_id):
            self.tree.item(item_id, text="☑", tags=('selected',))
        else:
            self.tree.item(item_id, text="☐", tags=())

    def select_all(self):
        for item_id in self.tree.get_children():
            self.selection_state[item_id] = True
            self.update_checkbox(item_id)
            
    def deselect_all(self):
        for item_id in self.tree.get_children():
            self.selection_state[item_id] = False
            self.update_checkbox(item_id)
            
    def invert_selection(self):
        for item_id in self.tree.get_children():
            self.selection_state[item_id] = not self.selection_state.get(item_id, False)
            self.update_checkbox(item_id)

    def confirm_delete(self):
        selected_ids = [self.tree.item(item_id, "values")[0] for item_id, selected in self.selection_state.items() if selected]
        
        if not selected_ids:
            messagebox.showwarning("Nenhuma Seleção", "Nenhum colaborador foi selecionado para exclusão.", parent=self)
            return

        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja remover os {len(selected_ids)} colaboradores selecionados?", parent=self):
            threading.Thread(target=self._confirm_delete_thread, args=(selected_ids,), daemon=True).start()

    def _confirm_delete_thread(self, ids_to_delete):
        success, message = dp_logic.delete_multiple_employees(ids_to_delete)
        self.after(0, self.on_delete_complete, success, message)

    def on_delete_complete(self, success, message):
        if success:
            messagebox.showinfo("Sucesso", message, parent=self)
            self.controller.invalidate_cache('employees', sub_key=self.selected_company_id.get())
            self.exit_delete_mode()
        else:
            messagebox.showerror("Erro", message, parent=self)

    def populate_company_combo(self):
        companies, error = self.controller.get_companies()
        if error: return
        self.companies_map = {c['name']: c['id'] for c in companies}
        self.company_combo['values'] = [c['name'] for c in companies]

    def on_company_select(self, event=None):
        if self.delete_mode: self.exit_delete_mode()
        self.refresh_employee_list()
        self.toggle_buttons("normal")

    def toggle_buttons(self, state):
        self.add_btn.config(state=state)
        self.edit_btn.config(state=state)
        self.delete_btn.config(state=state)
        self.import_btn.config(state=state)

    def refresh_employee_list(self, force=False):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        company_name = self.company_combo.get()
        if not company_name: return
        company_id = self.companies_map.get(company_name)
        if not company_id: return

        self.selected_company_id.set(company_id)
        
        employees, error = self.controller.get_employees(company_id, force_refresh=force)
        if error: return
        
        self.all_employees_list = employees
        self.selection_state.clear()
        for emp in self.all_employees_list:
            salary_formatted = f"R$ {emp.get('salary', 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            item_id = self.tree.insert('', 'end', values=(emp['id'], emp['employee_code'], emp['full_name'], salary_formatted))
            if self.delete_mode:
                self.selection_state[item_id] = False
                self.update_checkbox(item_id)

    def add_new_employee(self):
        company_id = self.selected_company_id.get()
        EmployeeDialog(self.controller, self, company_id=company_id)

    def edit_selected_employee(self):
        selected_item = self.tree.focus()
        if not selected_item: messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um colaborador para editar.", parent=self); return
        item_values = self.tree.item(selected_item, 'values')
        
        emp_id = item_values[0]
        emp_data_raw = next((e for e in self.all_employees_list if e['id'] == emp_id), None)
        
        if emp_data_raw:
            EmployeeDialog(self.controller, self, company_id=self.selected_company_id.get(), employee_to_edit=emp_data_raw)

    def import_from_file(self):
        company_id = self.selected_company_id.get()
        filepath = filedialog.askopenfilename(title="Selecione a planilha de colaboradores", filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
        if not filepath: return
        
        success, data_or_error = dp_logic.read_employees_from_file(company_id, filepath)

        if success:
            if not data_or_error: messagebox.showinfo("Aviso", "Nenhum colaborador novo encontrado na planilha.", parent=self); return
            EmployeeImportPreviewDialog(self.controller, self, company_id, data_or_error)
        else:
            messagebox.showerror("Erro na Leitura", data_or_error, parent=self)


# --- JANELAS PARA GERENCIAMENTO DE RUBRICAS ---

class PayrollCodeDialog(ttk.Toplevel):
    def __init__(self, controller, parent_window, code_to_edit=None):
        title = "Editar Rubrica" if code_to_edit else "Adicionar Nova Rubrica"
        super().__init__(title=title, master=parent_window)
        
        self.controller = controller
        self.parent_window = parent_window
        self.code_to_edit = code_to_edit
        
        self.transient(parent_window); self.grab_set()

        self.code_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.value_type_var = tk.StringVar()
        self.calc_base_var = tk.StringVar()
        self.calc_factor_var = tk.StringVar(value="1.0")
        
        self.create_widgets()

        if self.code_to_edit:
            self.load_code_data()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Código:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.code_var).grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(main_frame, text="Nome da Rubrica:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.name_var).grid(row=1, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(main_frame, text="Unidade de Lançamento:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        value_types = ['Valor', 'Horas', 'Dias', 'Percentual', 'Informativo']
        type_combo = ttk.Combobox(main_frame, textvariable=self.value_type_var, values=value_types, state="readonly")
        type_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        type_combo.set('Valor')

        ttk.Label(main_frame, text="Base de Cálculo:").grid(row=3, column=0, sticky="w", pady=5, padx=5)
        base_types = ["Valor Informado", "Baseado no Salário-Hora", "Percentual sobre o Salário"]
        base_combo = ttk.Combobox(main_frame, textvariable=self.calc_base_var, values=base_types, state="readonly")
        base_combo.grid(row=3, column=1, sticky="ew", pady=5, padx=5)
        base_combo.set("Valor Informado")

        ttk.Label(main_frame, text="Fator de Cálculo (ex: 1.5):").grid(row=4, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.calc_factor_var).grid(row=4, column=1, sticky="ew", pady=5, padx=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(20,0))
        ttk.Button(btn_frame, text="Salvar", command=self.save_code, bootstyle="primary").pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def load_code_data(self):
        self.code_var.set(self.code_to_edit['code'])
        self.name_var.set(self.code_to_edit['name'])
        self.value_type_var.set(self.code_to_edit['value_type'])
        self.calc_base_var.set(self.code_to_edit.get('calculation_base', 'Valor Informado'))
        self.calc_factor_var.set(self.code_to_edit.get('calculation_factor', 1.0))

    def save_code(self):
        code = self.code_var.get()
        name = self.name_var.get()
        value_type = self.value_type_var.get()
        calc_base = self.calc_base_var.get()
        calc_factor_str = self.calc_factor_var.get().replace(",", ".")

        if not all([code, name, value_type, calc_base, calc_factor_str]):
            messagebox.showerror("Erro de Validação", "Todos os campos são obrigatórios.", parent=self)
            return

        try:
            calc_factor_float = float(calc_factor_str)
        except ValueError:
            messagebox.showerror("Erro de Validação", "O Fator de Cálculo deve ser um número válido.", parent=self)
            return

        if self.code_to_edit:
            result = dp_logic.update_payroll_code(self.code_to_edit['id'], code, name, value_type, calc_base, calc_factor_float)
        else:
            result = dp_logic.add_payroll_code(code, name, value_type, calc_base, calc_factor_float)
        
        if "sucesso" in result:
            messagebox.showinfo("Sucesso", result, parent=self.parent_window)
            self.controller.invalidate_cache('payroll_codes')
            self.parent_window.refresh_payroll_code_list()
            self.destroy()
        else:
            messagebox.showerror("Erro", result, parent=self)


class PayrollCodeImportPreviewDialog(ttk.Toplevel):
    """Janela de pré-visualização para importação de Rubricas."""
    def __init__(self, controller, parent_window, payroll_code_data):
        super().__init__(title="Pré-visualização da Importação de Rubricas", master=parent_window)
        self.controller = controller
        self.parent_window = parent_window
        self.payroll_code_data = payroll_code_data
        
        self.minsize(800, 500)
        self.transient(parent_window); self.grab_set()

        self.create_widgets()
        self.populate_tree()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")

        instructions = "Abaixo estão as rubricas encontradas. Dê um duplo clique para editar. Registros 'Ignorados' não serão importados."
        ttk.Label(main_frame, text=instructions, wraplength=780, justify="left").pack(fill="x", pady=(0, 10))

        cols = ('Código', 'Nome', 'Unidade', 'Status')
        self.tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        self.tree.pack(expand=True, fill="both")

        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=80, stretch=False)
        self.tree.heading('Nome', text='Nome da Rubrica')
        self.tree.heading('Unidade', text='Unidade')
        self.tree.column('Unidade', width=120, stretch=False)
        self.tree.heading('Status', text='Status')
        self.tree.column('Status', width=150, stretch=False, anchor="center")
        
        self.tree.tag_configure("Novo", foreground="green")
        self.tree.tag_configure("Ignorado", foreground="gray")
        self.tree.bind("<Double-1>", self.on_double_click)

        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0))
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="Excluir Linha Selecionada", command=self.delete_selected_row, bootstyle="danger-outline").pack(side="left")
        
        self.confirm_btn = ttk.Button(btn_frame, text="Confirmar e Importar", command=self.confirm_import, bootstyle="success")
        self.confirm_btn.pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def populate_tree(self):
        for item in self.payroll_code_data:
            tag = item['status'].split(" ")[0]
            self.tree.insert('', 'end', values=(item['code'], item['name'], item['value_type'], item['status']), tags=(tag,))

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return

        selected_item = self.tree.focus()
        column = self.tree.identify_column(event.x)
        column_index = int(column.replace("#", "")) - 1

        if column_index > 2: return

        x, y, width, height = self.tree.bbox(selected_item, column)
        
        entry_var = tk.StringVar()
        entry = ttk.Entry(self.tree, textvariable=entry_var)
        entry.place(x=x, y=y, width=width, height=height)
        
        current_value = self.tree.item(selected_item, "values")[column_index]
        entry_var.set(current_value)
        entry.focus_set()
        entry.selection_range(0, 'end')

        def on_focus_out(event):
            self.update_tree_value(selected_item, column_index, entry_var.get())
            entry.destroy()

        def on_enter_key(event):
            self.update_tree_value(selected_item, column_index, entry_var.get())
            entry.destroy()

        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", on_enter_key)

    def update_tree_value(self, item, column_index, new_value):
        current_values = list(self.tree.item(item, "values"))
        current_values[column_index] = new_value
        self.tree.item(item, values=tuple(current_values))
    
    def delete_selected_row(self):
        selected_item = self.tree.focus()
        if selected_item:
            self.tree.delete(selected_item)

    def confirm_import(self):
        final_data_to_import = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, 'values')
            final_data_to_import.append({
                'code': values[0],
                'name': values[1],
                'value_type': values[2],
                'status': values[3]
            })

        self.confirm_btn.config(state="disabled", text="Importando...")
        self.update_idletasks()

        threading.Thread(target=self._confirm_import_thread, args=(final_data_to_import,), daemon=True).start()

    def _confirm_import_thread(self, data):
        success, message = dp_logic.save_imported_payroll_codes(data)
        self.after(0, self.on_import_complete, success, message)

    def on_import_complete(self, success, message):
        self.confirm_btn.config(state="normal", text="Confirmar e Importar")
        if success:
            messagebox.showinfo("Sucesso", message, parent=self.parent_window)
            self.controller.invalidate_cache('payroll_codes')
            self.parent_window.refresh_payroll_code_list()
            self.destroy()
        else:
            messagebox.showerror("Erro", message, parent=self)


class PayrollCodeManagementWindow(ttk.Toplevel):
    """Janela principal para Gerenciar Rubricas."""
    def __init__(self, controller, parent_frame):
        super().__init__(title="Gerenciar Rubricas", master=controller)
        self.controller = controller
        self.parent_frame = parent_frame
        self.minsize(800, 500); self.grab_set()
        
        self.all_codes_list = []
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_list)

        self.create_widgets()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")
        
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(search_frame, text="Pesquisar:").pack(side="left", padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True)
        
        tree_container = ttk.Frame(main_frame)
        tree_container.pack(expand=True, fill="both")

        cols = ('ID', 'Código', 'Nome', 'Tipo de Valor', 'Base de Cálculo', 'Fator')
        self.tree = ttk.Treeview(tree_container, columns=cols, show='headings')
        
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", expand=True, fill="both")

        self.tree.heading('Código', text='Código')
        self.tree.column('Código', width=80, stretch=False, anchor="center")
        self.tree.heading('Nome', text='Nome da Rubrica')
        self.tree.heading('Tipo de Valor', text='Unidade')
        self.tree.column('Tipo de Valor', width=100, stretch=False, anchor="center")
        self.tree.heading('Base de Cálculo', text='Base de Cálculo')
        self.tree.column('Base de Cálculo', width=180)
        self.tree.heading('Fator', text='Fator')
        self.tree.column('Fator', width=80, stretch=False, anchor="e")
        self.tree.column('ID', width=0, stretch=False)

        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0))
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Adicionar Nova", command=self.add_new_code, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Editar Selecionada", command=self.edit_selected_code).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Remover Selecionada", command=self.delete_selected_code, bootstyle="danger").pack(side="left")
        ttk.Button(btn_frame, text="Importar de Planilha...", command=self.import_from_file, bootstyle="info").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Atualizar Lista", command=lambda: self.refresh_payroll_code_list(force=True), bootstyle="secondary-outline").pack(side="right")
        
        self.refresh_payroll_code_list()

    def refresh_payroll_code_list(self, force=False):
        codes, error = self.controller.get_payroll_codes(force_refresh=force)
        if error: return
        self.all_codes_list = codes
        self.filter_list()

    def populate_tree(self, data_list):
        for i in self.tree.get_children(): self.tree.delete(i)
        for code in data_list:
            self.tree.insert('', 'end', values=(
                code['id'], 
                code['code'], 
                code['name'], 
                code['value_type'],
                code.get('calculation_base', 'N/A'),
                f"{code.get('calculation_factor', 1.0):.2f}"
            ))
            
    def filter_list(self, *args):
        search_term = self.search_var.get().lower()
        if not search_term:
            self.populate_tree(self.all_codes_list)
            return
        
        filtered_list = []
        for code in self.all_codes_list:
            if search_term in str(code['code']).lower() or search_term in str(code['name']).lower():
                filtered_list.append(code)
        
        self.populate_tree(filtered_list)

    def add_new_code(self):
        PayrollCodeDialog(self.controller, self)

    def edit_selected_code(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma rubrica para editar.", parent=self)
            return
        
        item_values = self.tree.item(selected_item, 'values')
        
        code_id = item_values[0]
        code_data = next((c for c in self.all_codes_list if c['id'] == code_id), None)

        if code_data:
            PayrollCodeDialog(self.controller, self, code_to_edit=code_data)

    def delete_selected_code(self):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma rubrica para remover.", parent=self)
            return
        
        item_values = self.tree.item(selected_item, 'values')
        code_id, code_name = item_values[0], item_values[2]

        if messagebox.askyesno("Confirmar Remoção", f"Tem certeza que deseja remover a rubrica '{code_name}'?", parent=self):
            result = dp_logic.delete_payroll_code(code_id)
            if "sucesso" in result:
                messagebox.showinfo("Sucesso", result, parent=self)
                self.controller.invalidate_cache('payroll_codes')
                self.refresh_payroll_code_list()
            else:
                messagebox.showerror("Erro", result, parent=self)
    
    def import_from_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione a planilha de rubricas",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not filepath: return
        
        success, data_or_error = dp_logic.read_payroll_codes_from_file(filepath)

        if success:
            if not data_or_error:
                messagebox.showinfo("Aviso", "Nenhuma rubrica nova encontrada na planilha.", parent=self)
                return
            PayrollCodeImportPreviewDialog(self.controller, self, data_or_error)
        else:
            messagebox.showerror("Erro na Leitura", data_or_error, parent=self)