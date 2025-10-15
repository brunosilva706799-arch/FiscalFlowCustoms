# =============================================================================
# --- ARQUIVO: ui/dialogs_user.py (VERSÃO COMPLETA E CORRIGIDA) ---
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox
import auth_logic
import threading
import os
import re
import webbrowser

def is_valid_email(email):
    """Verifica se o formato do e-mail é válido."""
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email)

class ChangePasswordDialog(ttk.Toplevel):
    def __init__(self, controller, user_id):
        super().__init__(title="Alterar Minha Senha", master=controller)
        self.controller = controller
        self.user_id = user_id
        self.transient(controller); self.grab_set()
        self.old_password_var = tk.StringVar(); self.new_password_var = tk.StringVar()
        self.confirm_password_var = tk.StringVar(); self.show_passwords_var = tk.BooleanVar(value=False)
        self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20); main_frame.pack(expand=True, fill="both")
        ttk.Label(main_frame, text="Senha Atual:").pack(fill="x")
        self.old_pass_entry = ttk.Entry(main_frame, textvariable=self.old_password_var, show="*"); self.old_pass_entry.pack(fill="x", pady=(0, 10)); self.old_pass_entry.focus_set()
        ttk.Label(main_frame, text="Nova Senha:").pack(fill="x")
        self.new_pass_entry = ttk.Entry(main_frame, textvariable=self.new_password_var, show="*"); self.new_pass_entry.pack(fill="x", pady=(0, 10))
        ttk.Label(main_frame, text="Confirmar Nova Senha:").pack(fill="x")
        self.confirm_pass_entry = ttk.Entry(main_frame, textvariable=self.confirm_password_var, show="*"); self.confirm_pass_entry.pack(fill="x", pady=(0, 5))
        ttk.Checkbutton(main_frame, text="Mostrar Senhas", variable=self.show_passwords_var, command=self._toggle_passwords_visibility).pack(anchor="w", pady=(0, 15))
        btn_frame = ttk.Frame(main_frame); btn_frame.pack(fill="x")
        self.save_button = ttk.Button(btn_frame, text="Salvar Nova Senha", command=self.save_password, bootstyle="primary"); self.save_button.pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)
    def _toggle_passwords_visibility(self):
        show_char = "" if self.show_passwords_var.get() else "*"; self.old_pass_entry.config(show=show_char)
        self.new_pass_entry.config(show=show_char); self.confirm_pass_entry.config(show=show_char)
    def save_password(self):
        old_pass, new_pass, confirm_pass = self.old_password_var.get(), self.new_password_var.get(), self.confirm_password_var.get()
        if not all([old_pass, new_pass, confirm_pass]): messagebox.showerror("Erro", "Todos os campos são obrigatórios.", parent=self); return
        if new_pass != confirm_pass: messagebox.showerror("Erro", "As novas senhas não coincidem.", parent=self); return
        self.save_button.config(state="disabled"); self.config(cursor="watch"); self.update_idletasks()
        try:
            result = auth_logic.change_password(self.user_id, old_pass, new_pass)
            if "sucesso" in result: messagebox.showinfo("Sucesso", result, parent=self.controller); self.destroy()
            else: messagebox.showerror("Erro", result, parent=self)
        finally:
            self.save_button.config(state="normal"); self.config(cursor="")
            
class RequestAccessDialog(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Solicitar Acesso", master=controller)
        self.controller = controller
        self.geometry("400x350"); self.transient(controller); self.grab_set()
        self.full_name_var = tk.StringVar(); self.email_var = tk.StringVar(); self.sector_var = tk.StringVar()
        self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20); main_frame.pack(expand=True, fill="both")
        ttk.Label(main_frame, text="Nome Completo:").pack(fill="x"); ttk.Entry(main_frame, textvariable=self.full_name_var).pack(fill="x", pady=(0, 10))
        ttk.Label(main_frame, text="E-mail Corporativo:").pack(fill="x"); ttk.Entry(main_frame, textvariable=self.email_var).pack(fill="x", pady=(0, 10))
        ttk.Label(main_frame, text="Setor:").pack(fill="x")
        all_sectors, error = self.controller.get_sectors()
        if error: sector_names = ["Erro ao carregar setores"]
        else: sector_names = [s['name'] for s in all_sectors] if all_sectors else ["Nenhum setor cadastrado"]
        sector_combo = ttk.Combobox(main_frame, textvariable=self.sector_var, values=sector_names, state="readonly"); sector_combo.pack(fill="x", pady=(0, 20))
        if sector_names and "Erro" not in sector_names[0] and "Nenhum" not in sector_names[0]: sector_combo.set(sector_names[0])
        btn_frame = ttk.Frame(main_frame); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Enviar Solicitação", command=self.submit_request, bootstyle="primary").pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)
    def submit_request(self):
        full_name, email, sector = self.full_name_var.get(), self.email_var.get(), self.sector_var.get()
        if not full_name.strip() or not email.strip() or not sector.strip(): messagebox.showwarning("Campos Vazios", "Todos os campos são obrigatórios.", parent=self); return
        if "Erro" in sector or "Nenhum" in sector: messagebox.showwarning("Setor Inválido", "Não há um setor válido. Contate um administrador.", parent=self); return
        if not is_valid_email(email): messagebox.showerror("E-mail Inválido", "Insira um e-mail válido.", parent=self); return
        result = auth_logic.create_access_request(full_name, email, sector)
        messagebox.showinfo("Solicitação de Acesso", result, parent=self.controller)
        if "sucesso" in result: self.destroy()

class RequestsWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Solicitações de Acesso", master=controller)
        self.controller = controller; self.minsize(700, 400); self.grab_set(); self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        cols = ('ID', 'Data', 'Nome Completo', 'E-mail', 'Setor'); self.requests_tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        for col in cols: self.requests_tree.heading(col, text=col)
        self.requests_tree.column('ID', width=0, stretch=False); self.requests_tree.column('Data', width=120, stretch=False)
        self.requests_tree.column('Nome Completo', width=200); self.requests_tree.column('E-mail', width=200)
        self.requests_tree.column('Setor', width=100); self.requests_tree.pack(expand=True, fill="both")
        btn_frame = ttk.Frame(main_frame, padding=(0, 10)); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Aprovar e Criar Usuário", command=self.approve_request, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Rejeitar Solicitação", command=self.reject_request, bootstyle="danger").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Atualizar Lista", command=self.refresh_requests_list, bootstyle="secondary-outline").pack(side="right")
        self.refresh_requests_list()
    def refresh_requests_list(self):
        for i in self.requests_tree.get_children(): self.requests_tree.delete(i)
        requests, error = auth_logic.get_pending_requests()
        if error: messagebox.showerror("Erro", error, parent=self); return
        for req in requests: self.requests_tree.insert('', 'end', values=(req['id'], req['request_date'], req['full_name'], req['email'], req['sector']))
        title = "Solicitações de Acesso";
        if requests: title += f" ({len(requests)})"
        self.title(title)
    def approve_request(self):
        selected = self.requests_tree.focus()
        if not selected: messagebox.showwarning("Seleção", "Selecione uma solicitação.", parent=self); return
        req_data = self.requests_tree.item(selected, 'values')
        UserDialog(self.controller, self, request_to_approve={'id': req_data[0], 'username': req_data[3].split('@')[0], 'email': req_data[3]})
    def reject_request(self):
        selected = self.requests_tree.focus()
        if not selected: messagebox.showwarning("Seleção", "Selecione uma solicitação.", parent=self); return
        req_id, req_name = self.requests_tree.item(selected, 'values')[0:3:2]
        if messagebox.askyesno("Confirmar", f"Rejeitar a solicitação de '{req_name}'?", parent=self):
            _, message = auth_logic.update_request_status(req_id, "rejeitado")
            messagebox.showinfo("Resultado", message, parent=self); self.refresh_requests_list()

class UserManagementWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Gerenciar Usuários", master=controller)
        self.controller = controller; self.minsize(700, 400); self.grab_set(); self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        cols = ('ID', 'Usuário', 'E-mail', 'Nível', 'Acesso Códigos'); self.user_tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        for col in cols: self.user_tree.heading(col, text=col)
        self.user_tree.column('ID', width=0, stretch=False); self.user_tree.column('Usuário', width=150)
        self.user_tree.column('E-mail', width=200); self.user_tree.column('Nível', width=100, anchor='center')
        self.user_tree.column('Acesso Códigos', width=120, anchor='center'); self.user_tree.pack(expand=True, fill="both"); self.refresh_user_list()
        btn_frame = ttk.Frame(main_frame, padding=(0, 10)); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Adicionar Usuário", command=self.add_user, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Editar Usuário", command=self.edit_user).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Remover Usuário", command=self.delete_user, bootstyle="danger").pack(side="left")
        ttk.Button(btn_frame, text="Atualizar Lista", command=self.refresh_user_list, bootstyle="secondary-outline").pack(side="right")
    def refresh_user_list(self):
        for i in self.user_tree.get_children(): self.user_tree.delete(i)
        self.all_users_cache, error = auth_logic.get_all_users()
        if error: messagebox.showerror("Erro", error, parent=self); return
        access_map = {'Nenhum': 'Nenhum', 'Consulta': 'Consulta', 'Total': 'Total'}
        for user in self.all_users_cache:
            access_text = access_map.get(user.get('acesso_codigos_cliente', 'Nenhum'), 'N/A')
            self.user_tree.insert('', 'end', values=(user['id'], user['username'], user['email'], user['level'], access_text))
    def add_user(self):
        UserDialog(self.controller, self)
    def edit_user(self):
        selected = self.user_tree.focus()
        if not selected: messagebox.showwarning("Seleção", "Selecione um usuário.", parent=self); return
        user_id = self.user_tree.item(selected, 'values')[0]
        user_dict = next((u for u in self.all_users_cache if u['id'] == user_id), None)
        if user_dict: UserDialog(self.controller, self, user_to_edit=user_dict)
    def delete_user(self):
        selected = self.user_tree.focus()
        if not selected: messagebox.showwarning("Seleção", "Selecione um usuário.", parent=self); return
        user_id, username = self.user_tree.item(selected, 'values')[0:2]
        if messagebox.askyesno("Confirmar", f"Remover o usuário '{username}'?", parent=self):
            result = auth_logic.delete_user(user_id); messagebox.showinfo("Resultado", result, parent=self); self.refresh_user_list()

class UserDialog(ttk.Toplevel):
    def __init__(self, controller, parent_admin_panel, user_to_edit=None, request_to_approve=None):
        self.user_to_edit, self.request_to_approve = user_to_edit, request_to_approve
        title = "Editar Usuário" if user_to_edit else "Adicionar Novo Usuário"
        super().__init__(title=title, master=parent_admin_panel)
        self.controller = controller; self.parent_admin_panel = parent_admin_panel
        self.transient(parent_admin_panel); self.grab_set()
        self.username_var = tk.StringVar(); self.email_var = tk.StringVar(); self.level_var = tk.StringVar(value="Padrão")
        self.client_code_access_var = tk.StringVar(value="Nenhum Acesso"); self.sector_vars = {}
        self.create_widgets()
        if self.user_to_edit: self.load_user_data()
        if self.request_to_approve: self.username_var.set(self.request_to_approve['username']); self.email_var.set(self.request_to_approve['email'])
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20); main_frame.pack(expand=True, fill="both")
        form_frame = ttk.Labelframe(main_frame, text=" Dados do Usuário ", padding=15); form_frame.pack(fill="x")
        form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="Usuário:").grid(row=0, column=0, sticky='w', pady=5); ttk.Entry(form_frame, textvariable=self.username_var).grid(row=0, column=1, sticky='ew', pady=5)
        ttk.Label(form_frame, text="E-mail:").grid(row=1, column=0, sticky='w', pady=5); ttk.Entry(form_frame, textvariable=self.email_var).grid(row=1, column=1, sticky='ew', pady=5)
        ttk.Label(form_frame, text="Nível de Acesso:").grid(row=2, column=0, sticky='w', pady=5); ttk.Combobox(form_frame, textvariable=self.level_var, values=['Padrão', 'Admin', 'T.I.'], state="readonly").grid(row=2, column=1, sticky='ew', pady=5)
        ttk.Label(form_frame, text="Acesso Ferramenta Códigos:").grid(row=3, column=0, sticky='w', pady=5)
        access_values = ['Nenhum Acesso', 'Apenas Consulta', 'Acesso Total']
        ttk.Combobox(form_frame, textvariable=self.client_code_access_var, values=access_values, state="readonly").grid(row=3, column=1, sticky='ew', pady=5)
        if not self.user_to_edit: ttk.Label(form_frame, text="A senha será definida pelo usuário no primeiro acesso.", bootstyle="info").grid(row=4, column=0, columnspan=2, sticky='w', pady=10)
        sectors_frame = ttk.Labelframe(main_frame, text=" Associar a Setores ", padding=15); sectors_frame.pack(fill="x", pady=20)
        all_sectors, error = self.controller.get_sectors()
        if error: ttk.Label(sectors_frame, text="Não foi possível carregar setores.", bootstyle="danger").pack()
        else:
            for sector in all_sectors:
                var = tk.BooleanVar(); ttk.Checkbutton(sectors_frame, text=sector['name'], variable=var).pack(anchor='w'); self.sector_vars[sector['id']] = var
        btn_frame = ttk.Frame(main_frame); btn_frame.pack(fill="x")
        self.save_btn = ttk.Button(btn_frame, text="Salvar", command=self.save_user, bootstyle="primary"); self.save_btn.pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)
    def load_user_data(self):
        self.username_var.set(self.user_to_edit['username']); self.email_var.set(self.user_to_edit['email']); self.level_var.set(self.user_to_edit['level'])
        db_access_level = self.user_to_edit.get('acesso_codigos_cliente', 'Nenhum')
        ui_access_map = {'Nenhum': 'Nenhum Acesso', 'Consulta': 'Apenas Consulta', 'Total': 'Acesso Total'}
        self.client_code_access_var.set(ui_access_map.get(db_access_level, 'Nenhum Acesso'))
        user_sectors = auth_logic.get_user_sectors(self.user_to_edit['id'])
        for sector_id, var in self.sector_vars.items():
            if sector_id in user_sectors: var.set(True)
    def save_user(self):
        username, email, level = self.username_var.get(), self.email_var.get(), self.level_var.get()
        if not is_valid_email(email): messagebox.showerror("E-mail Inválido", "Insira um e-mail válido.", parent=self); return
        selected_sector_ids = [sid for sid, var in self.sector_vars.items() if var.get()]
        ui_access_level = self.client_code_access_var.get()
        db_access_map = {'Nenhum Acesso': 'Nenhum', 'Apenas Consulta': 'Consulta', 'Acesso Total': 'Total'}
        db_access_level = db_access_map.get(ui_access_level)
        self.save_btn.config(state="disabled", text="Salvando..."); self.config(cursor="watch"); self.update_idletasks()
        threading.Thread(target=self._save_user_in_thread, args=(username, email, level, selected_sector_ids, db_access_level), daemon=True).start()
    def _save_user_in_thread(self, username, email, level, sector_ids, access_level):
        if self.user_to_edit:
            result = auth_logic.update_user(self.user_to_edit['id'], username, email, level, sector_ids, client_code_access=access_level)
            self.after(0, self.on_save_complete, result)
        else:
            user_id, message = auth_logic.add_user(username, email, level, sector_ids, client_code_access=access_level)
            if user_id:
                if self.request_to_approve: auth_logic.update_request_status(self.request_to_approve['id'], 'aprovado')
                token = auth_logic.generate_password_setup_token(user_id)
                _, email_message = auth_logic.send_creation_email(email, username, token)
                final_message = f"{message}\n\n{email_message}"
                self.after(0, self.on_save_complete, final_message)
            else: self.after(0, self.on_save_complete, message)
    def on_save_complete(self, result_message):
        self.config(cursor=""); self.save_btn.config(state="normal", text="Salvar")
        messagebox.showinfo("Resultado", result_message, parent=self.parent_admin_panel)
        if "sucesso" in result_message or "enviado" in result_message:
            if hasattr(self.parent_admin_panel, 'refresh_user_list'): self.parent_admin_panel.refresh_user_list()
            if hasattr(self.parent_admin_panel, 'refresh_requests_list'): self.parent_admin_panel.refresh_requests_list()
            self.destroy()

class SectorManagementWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Gerenciar Setores", master=controller); self.controller = controller; self.minsize(400, 300); self.grab_set(); self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        cols = ('ID', 'Nome do Setor'); self.sector_tree = ttk.Treeview(main_frame, columns=cols, show='headings')
        self.sector_tree.heading('Nome do Setor', text='Nome do Setor'); self.sector_tree.column('ID', width=0, stretch=False)
        self.sector_tree.pack(expand=True, fill="both"); self.refresh_sector_list()
        btn_frame = ttk.Frame(main_frame, padding=(0, 10)); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Adicionar Setor", command=self.add_sector, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Remover Setor", command=self.delete_sector, bootstyle="danger").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Atualizar Lista", command=lambda: self.refresh_sector_list(force=True), bootstyle="secondary-outline").pack(side="right")
    def refresh_sector_list(self, force=False):
        for i in self.sector_tree.get_children(): self.sector_tree.delete(i)
        sectors, error = self.controller.get_sectors(force_refresh=force)
        if error: return
        for sector in sectors: self.sector_tree.insert('', 'end', values=(sector['id'], sector['name']))
    def add_sector(self):
        new_name = simpledialog.askstring("Novo Setor", "Nome do novo setor:", parent=self)
        if new_name:
            result = auth_logic.add_sector(new_name); messagebox.showinfo("Resultado", result, parent=self)
            if "sucesso" in result: self.controller.invalidate_cache('sectors'); self.refresh_sector_list()
    def delete_sector(self):
        selected = self.sector_tree.focus()
        if not selected: messagebox.showwarning("Seleção", "Selecione um setor.", parent=self); return
        s_id, s_name = self.sector_tree.item(selected, 'values')
        if messagebox.askyesno("Confirmar", f"Remover o setor '{s_name}'?", parent=self):
            result = auth_logic.delete_sector(s_id); messagebox.showinfo("Resultado", result, parent=self)
            if "sucesso" in result: self.controller.invalidate_cache('sectors'); self.refresh_sector_list()

class TemplateEditorWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Editar Templates de E-mail", master=controller); self.controller = controller; self.minsize(700, 400); self.grab_set()
        self.selected_template_identifier = None; self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(1, weight=1); main_frame.rowconfigure(0, weight=1)
        left_pane = ttk.Frame(main_frame, padding=(0, 0, 10, 0)); left_pane.grid(row=0, column=0, sticky="ns")
        cols = ('Identifier', 'Nome do Template'); self.template_tree = ttk.Treeview(left_pane, columns=cols, show='headings', height=10)
        self.template_tree.heading('Nome do Template', text='Template'); self.template_tree.column('Identifier', width=0, stretch=False)
        self.template_tree.pack(expand=True, fill="y"); self.template_tree.bind('<<TreeviewSelect>>', self.on_template_select)
        right_pane = ttk.Frame(main_frame); right_pane.grid(row=0, column=1, sticky="nsew")
        info_frame = ttk.Labelframe(right_pane, text=" Informações do Template ", padding=15); info_frame.pack(fill="x")
        self.subject_label = ttk.Label(info_frame, text="Assunto:", font="-weight bold"); self.subject_label.pack(fill="x")
        self.variables_label = ttk.Label(info_frame, text="Variáveis disponíveis:", bootstyle="info", padding=(0,5,0,0)); self.variables_label.pack(fill="x")
        action_frame = ttk.Frame(right_pane, padding=(0,20)); action_frame.pack(fill="x")
        ttk.Button(action_frame, text="Exportar para Editar...", command=self.export_template).pack(side="left")
        ttk.Button(action_frame, text="Importar Template Editado...", command=self.import_template, bootstyle="primary").pack(side="left", padx=10)
        self.refresh_template_list()
    def refresh_template_list(self):
        self.on_template_select(None)
        for i in self.template_tree.get_children(): self.template_tree.delete(i)
        for t in auth_logic.get_all_email_templates(): self.template_tree.insert('', 'end', values=(t['identifier'], t['name']))
    def on_template_select(self, event):
        selected = self.template_tree.focus()
        if not selected: self.selected_template_identifier=None; self.subject_label.config(text="Assunto:"); self.variables_label.config(text="Variáveis:"); return
        identifier = self.template_tree.item(selected, 'values')[0]; self.selected_template_identifier = identifier
        template_data = auth_logic.get_email_template(identifier)
        if not template_data: return
        self.subject_label.config(text=f"Assunto: {template_data['subject']}")
        if 'support_reply' in identifier: vars_text = "{username}, {ticket_id}, {subject}, {attendant_name}"
        elif 'support_closed' in identifier: vars_text = "{username}, {ticket_id}, {subject}"
        elif 'password_reset' in identifier or 'account_creation' in identifier: vars_text = "{username}, {token}"
        else: vars_text = "N/A"
        self.variables_label.config(text=f"Variáveis disponíveis: {vars_text}")
    def export_template(self):
        if not self.selected_template_identifier: messagebox.showwarning("Seleção", "Selecione um template.", parent=self); return
        template_data = auth_logic.get_email_template(self.selected_template_identifier)
        if not template_data: return
        fp = filedialog.asksaveasfilename(title="Salvar", initialfile=f"{self.selected_template_identifier}.html", defaultextension=".html", filetypes=[("HTML", "*.html")])
        if not fp: return
        try:
            with open(fp, 'w', encoding='utf-8') as f: f.write(template_data['body_html'])
            if messagebox.askyesno("Sucesso", f"Template exportado.\nDeseja abrir?", parent=self): os.startfile(fp)
        except Exception as e: messagebox.showerror("Erro", f"Não foi possível exportar.\n\nErro: {e}", parent=self)
    def import_template(self):
        if not self.selected_template_identifier: messagebox.showwarning("Seleção", "Selecione um template.", parent=self); return
        fp = filedialog.askopenfilename(title="Importar", filetypes=[("HTML", "*.html")])
        if not fp: return
        try:
            with open(fp, 'r', encoding='utf-8') as f: content = f.read()
            new_subject = auth_logic.get_email_template(self.selected_template_identifier)['subject']
            result = auth_logic.update_email_template(self.selected_template_identifier, new_subject, content)
            messagebox.showinfo("Resultado", result, parent=self); self.refresh_template_list()
        except Exception as e: messagebox.showerror("Erro", f"Não foi possível importar.\n\nErro: {e}", parent=self)
        
class CommunicationWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Enviar Comunicado Global", master=controller); self.controller = controller; self.minsize(600, 500); self.grab_set(); self.create_widgets()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        main_frame.columnconfigure(0, weight=1); main_frame.rowconfigure(2, weight=1)
        dest_frame = ttk.Labelframe(main_frame, text=" Destinatários ", padding=15); dest_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.recipient_type = tk.StringVar(value="Todos")
        ttk.Radiobutton(dest_frame, text="Todos os Usuários", variable=self.recipient_type, value="Todos", command=self.update_recipient_options).pack(anchor="w")
        ttk.Radiobutton(dest_frame, text="Selecionar por Setor", variable=self.recipient_type, value="Setor", command=self.update_recipient_options).pack(anchor="w")
        self.sector_combo = ttk.Combobox(dest_frame, state="disabled"); self.sector_combo.pack(fill="x", padx=(20,0), pady=5)
        msg_frame = ttk.Labelframe(main_frame, text=" Mensagem ", padding=15); msg_frame.grid(row=1, column=0, sticky="ew")
        msg_frame.columnconfigure(1, weight=1)
        ttk.Label(msg_frame, text="Assunto:").grid(row=0, column=0, sticky="w", padx=(0,10))
        self.comm_subject_entry = ttk.Entry(msg_frame); self.comm_subject_entry.grid(row=0, column=1, sticky="ew")
        self.comm_body_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=10); self.comm_body_text.grid(row=2, column=0, sticky="nsew", pady=10)
        self.send_comm_btn = ttk.Button(main_frame, text="Enviar Comunicado", command=self.send_communication, bootstyle="primary")
        self.send_comm_btn.grid(row=3, column=0, sticky="e", pady=(10,0))
    def update_recipient_options(self):
        if self.recipient_type.get() == "Setor":
            self.sector_combo.config(state="readonly"); sectors, error = self.controller.get_sectors()
            if error: return
            self.sector_combo['values'] = [s['name'] for s in sectors]
            if sectors: self.sector_combo.set(sectors[0]['name'])
        else: self.sector_combo.config(state="disabled"); self.sector_combo.set("")
    def send_communication(self):
        rec_type, subj, body = self.recipient_type.get(), self.comm_subject_entry.get(), self.comm_body_text.get("1.0", "end-1c")
        if not (subj and body): messagebox.showwarning("Campos Vazios", "Assunto e corpo são obrigatórios.", parent=self); return
        rec_list = []
        if rec_type == "Todos": emails, err = auth_logic.get_all_user_emails()
        elif rec_type == "Setor":
            s_name = self.sector_combo.get()
            if not s_name: messagebox.showwarning("Seleção", "Selecione um setor.", parent=self); return
            all_sectors, err = self.controller.get_sectors()
            if err: return
            s_id = next((s['id'] for s in all_sectors if s['name'] == s_name), None)
            if s_id: emails, err = auth_logic.get_user_emails_by_sector(s_id)
        if err: messagebox.showerror("Erro", err, parent=self); return
        rec_list = emails
        if not rec_list: messagebox.showinfo("Aviso", "Nenhum destinatário encontrado.", parent=self); return
        self.send_comm_btn.config(state="disabled", text="Enviando...")
        threading.Thread(target=self._send_comm_in_thread, args=(rec_list, subj, body), daemon=True).start()
    def _send_comm_in_thread(self, rec_list, subj, body):
        stat, msg = auth_logic.send_communication_email(rec_list, subj, body); self.after(0, self.show_comm_status, stat, msg)
    def show_comm_status(self, status, message):
        if status:
            messagebox.showinfo("Envio", message, parent=self)
            self.comm_subject_entry.delete(0, tk.END); self.comm_body_text.delete("1.0", tk.END)
        else: messagebox.showerror("Erro", message, parent=self)
        self.send_comm_btn.config(state="normal", text="Enviar Comunicado")