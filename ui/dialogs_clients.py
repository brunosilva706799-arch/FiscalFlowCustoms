# ============================================================================
# --- ARQUIVO: ui/dialogs_clients.py ---
# (Interface para o Gerenciador de Códigos de Cliente e Relatório)
# ============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttk
import threading
import pytz
from datetime import datetime
import client_logic 
import report_logic 

class ClientCodeManagerWindow(ttk.Toplevel):
    """
    Janela para gerenciar (criar, consultar, editar, excluir) os códigos de cliente.
    """
    def __init__(self, controller, parent):
        super().__init__(title="Gerenciador de Códigos de Cliente", master=parent)
        self.controller = controller
        
        self.all_clients_cache = []

        self.geometry("900x500")
        self.minsize(800, 400)
        
        # self.transient(parent) # Removido para permitir minimizar/maximizar
        # self.grab_set()       # Removido para permitir minimizar/maximizar
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', self.filter_list)
        self.name_var = tk.StringVar()
        self.code_var = tk.StringVar()
        self.select_all_var = tk.BooleanVar()
        
        self.create_widgets()
        self.load_clients()
        self._apply_permissions()

    def create_widgets(self):
        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(expand=True, fill="both", padx=10, pady=10)

        list_frame = ttk.Frame(main_pane, padding=5)
        list_frame.rowconfigure(4, weight=1)
        list_frame.columnconfigure(0, weight=1)
        
        ttk.Label(list_frame, text="Buscar Cliente:", font=("-size 10 -weight bold")).grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(list_frame, textvariable=self.search_var)
        search_entry.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        self.import_btn = ttk.Button(list_frame, text="Importar Planilha...", command=self.import_from_file, bootstyle="info")
        self.import_btn.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        tree_frame = ttk.Frame(list_frame)
        tree_frame.grid(row=4, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        columns = ("name", "code")
        self.client_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.client_tree.heading("name", text="Nome do Cliente")
        self.client_tree.heading("code", text="Código")
        self.client_tree.column("name", width=250)
        self.client_tree.column("code", width=120, anchor='center')

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.client_tree.yview)
        self.client_tree.configure(yscrollcommand=scrollbar.set)
        
        self.client_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.client_tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        # --- [NOVO] Adiciona o bind para o clique direito ---
        self.client_tree.bind("<Button-3>", self.show_context_menu)
        
        select_all_check = ttk.Checkbutton(list_frame, variable=self.select_all_var, 
                                           text="Selecionar Todos / Limpar Seleção", 
                                           command=self.toggle_select_all)
        select_all_check.grid(row=3, column=0, sticky="w", pady=5)
        
        main_pane.add(list_frame, weight=2)

        self.form_frame = ttk.Labelframe(main_pane, text=" Gerenciar Cliente ", padding=15)
        self.form_frame.rowconfigure(4, weight=1)
        self.form_frame.columnconfigure(0, weight=1)
        self.form_frame.columnconfigure(1, weight=0)

        ttk.Label(self.form_frame, text="Nome do Cliente:").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))
        self.name_entry = ttk.Entry(self.form_frame, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=0, columnspan=2, sticky="ew")

        ttk.Label(self.form_frame, text="Código Gerado:").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 5))
        self.code_entry = ttk.Entry(self.form_frame, textvariable=self.code_var, state="readonly")
        self.code_entry.grid(row=3, column=0, sticky="ew")

        self.copy_btn = ttk.Button(self.form_frame, text="📋", command=self.copy_code_to_clipboard, bootstyle="secondary-outline", width=2)
        self.copy_btn.grid(row=3, column=1, sticky="w", padx=(5,0))

        btn_container = ttk.Frame(self.form_frame)
        btn_container.grid(row=5, column=0, columnspan=2, sticky="sew", pady=(20, 0))
        btn_container.columnconfigure((0,1), weight=1)

        self.save_new_btn = ttk.Button(btn_container, text="Salvar Novo", command=self.save_new, bootstyle="success")
        self.save_new_btn.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,5))
        
        self.save_edit_btn = ttk.Button(btn_container, text="Salvar Edição", command=self.save_edit, state="disabled")
        self.save_edit_btn.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0,5))
        
        self.delete_btn = ttk.Button(btn_container, text="Excluir Selecionado(s)", command=self.delete_selected, bootstyle="danger", state="disabled")
        self.delete_btn.grid(row=2, column=0, sticky="ew", padx=(0,2))
        
        self.clear_btn = ttk.Button(btn_container, text="Limpar Formulário", command=self.clear_form, bootstyle="secondary-outline")
        self.clear_btn.grid(row=2, column=1, sticky="ew", padx=(2,0))
        
        main_pane.add(self.form_frame, weight=1)
        
        # --- [NOVO] Cria o menu de contexto ---
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Excluir Selecionado(s)", command=self.delete_selected)

    # --- [NOVA FUNÇÃO] Para mostrar o menu de clique-direito ---
    def show_context_menu(self, event):
        iid = self.client_tree.identify_row(event.y)
        if iid:
            if iid not in self.client_tree.selection():
                self.client_tree.selection_set(iid)
        
        access_level = self.controller.current_user.get('acesso_codigos_cliente', 'Nenhum')
        if self.client_tree.selection() and access_level == 'Total':
            self.context_menu.entryconfig("Excluir Selecionado(s)", state="normal")
        else:
            self.context_menu.entryconfig("Excluir Selecionado(s)", state="disabled")
            
        self.context_menu.post(event.x_root, event.y_root)

    def _apply_permissions(self):
        access_level = self.controller.current_user.get('acesso_codigos_cliente', 'Nenhum')
        if access_level == 'Consulta':
            self.title("Gerenciador de Códigos de Cliente (Modo Consulta)")
            self.form_frame.config(text=" Consultar Cliente ")
            self.import_btn.config(state="disabled"); self.name_entry.config(state="readonly")
            self.save_new_btn.config(state="disabled"); self.save_edit_btn.config(state="disabled")
            self.delete_btn.config(state="disabled")

    def copy_code_to_clipboard(self):
        code = self.code_var.get()
        if not code: return
        self.clipboard_clear(); self.clipboard_append(code)
        original_text = self.copy_btn.cget("text")
        self.copy_btn.config(text="✓", bootstyle="success")
        self.after(1500, lambda: self.copy_btn.config(text=original_text, bootstyle="secondary-outline"))

    def load_clients(self):
        # ... (código da função inalterado) ...
        self.status_label = ttk.Label(self, text="Carregando clientes...", bootstyle="secondary")
        self.status_label.pack(side="bottom", fill="x", padx=10, pady=(0, 5))
        threading.Thread(target=self._load_clients_thread, daemon=True).start()
    def _load_clients_thread(self):
        clients, error = client_logic.get_all_clients()
        if error: self.after(0, messagebox.showerror, "Erro", error, {"parent": self})
        else: self.all_clients_cache = clients; self.after(0, self.filter_list)
        self.after(0, self.status_label.destroy)
    def _populate_treeview(self, clients_to_display):
        self.client_tree.delete(*self.client_tree.get_children())
        for client in clients_to_display:
            self.client_tree.insert("", "end", iid=client['id'], values=(client.get('name'), client.get('code')))
    def filter_list(self, *args):
        search_term = self.search_var.get().lower()
        filtered = self.all_clients_cache if not search_term else [c for c in self.all_clients_cache if search_term in c.get('name', '').lower()]
        self._populate_treeview(filtered)
    def on_tree_select(self, event=None):
        items = self.client_tree.selection()
        access_level = self.controller.current_user.get('acesso_codigos_cliente', 'Nenhum')
        if len(items) == 1:
            if access_level == 'Total':
                self.form_frame.config(text=" Editar Cliente "); self.name_entry.config(state="normal")
                self.save_new_btn.config(state="disabled"); self.save_edit_btn.config(state="normal"); self.delete_btn.config(state="normal", bootstyle="danger-outline")
            data = next((c for c in self.all_clients_cache if c['id'] == items[0]), None)
            if data: self.name_var.set(data.get('name', '')); self.code_var.set(data.get('code', ''))
        elif len(items) > 1:
            self.clear_form(clear_selection=False); self.form_frame.config(text=f" {len(items)} clientes selecionados ")
            self.name_entry.config(state="disabled"); self.save_new_btn.config(state="disabled"); self.save_edit_btn.config(state="disabled")
            if access_level == 'Total': self.delete_btn.config(state="normal", bootstyle="danger")
        else: self.clear_form()
    def clear_form(self, clear_selection=True):
        self.name_var.set(""); self.code_var.set(""); 
        self.form_frame.config(text=" Gerenciar Cliente ")
        if clear_selection and self.client_tree.selection(): self.client_tree.selection_remove(self.client_tree.selection())
        access_level = self.controller.current_user.get('acesso_codigos_cliente', 'Nenhum')
        if access_level == 'Total':
            self.name_entry.config(state="normal")
            self.save_new_btn.config(state="normal"); self.save_edit_btn.config(state="disabled"); self.delete_btn.config(state="disabled", bootstyle="danger-outline")
    def toggle_select_all(self):
        if self.select_all_var.get(): self.client_tree.selection_set(self.client_tree.get_children())
        else: self.client_tree.selection_remove(self.client_tree.selection())
    def save_new(self):
        name = self.name_var.get()
        if not name.strip(): messagebox.showwarning("Campo Vazio", "O nome do cliente é obrigatório.", parent=self); return
        self.save_new_btn.config(state="disabled", text="Salvando...")
        threading.Thread(target=self._save_new_thread, args=(name,), daemon=True).start()
    def _save_new_thread(self, name):
        _, msg = client_logic.create_client_code(name); self.after(0, self.on_save_complete, msg is not None, msg)
    def save_edit(self):
        name = self.name_var.get(); sel = self.client_tree.selection()
        if not sel or not name.strip(): return
        self.save_edit_btn.config(state="disabled", text="Atualizando...")
        threading.Thread(target=self._save_edit_thread, args=(sel[0], name), daemon=True).start()
    def _save_edit_thread(self, cid, name):
        ok, msg = client_logic.update_client(cid, name); self.after(0, self.on_save_complete, ok, msg)
    def on_save_complete(self, success, message):
        self.save_new_btn.config(state="normal", text="Salvar Novo"); self.save_edit_btn.config(text="Salvar Edição")
        if success: messagebox.showinfo("Sucesso", message, parent=self); self.load_clients(); self.clear_form()
        else: messagebox.showerror("Erro", message, parent=self); self.clear_form()
    def delete_selected(self):
        sel_ids = self.client_tree.selection()
        if not sel_ids: return
        if not messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir os {len(sel_ids)} cliente(s) selecionado(s)?", parent=self, icon='warning'): return
        self.delete_btn.config(state="disabled", text="Excluindo...")
        threading.Thread(target=self._delete_thread, args=(sel_ids,), daemon=True).start()
    def _delete_thread(self, cids):
        ok, msg = client_logic.delete_clients_batch(cids); self.after(0, self.on_delete_complete, ok, msg)
    def on_delete_complete(self, success, message):
        self.delete_btn.config(text="Excluir Selecionado(s)")
        if success: messagebox.showinfo("Sucesso", message, parent=self); self.load_clients(); self.clear_form()
        else: messagebox.showerror("Erro", message, parent=self); self.clear_form()
    def import_from_file(self):
        fp = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Excel", "*.xlsx")])
        if not fp: return
        self.import_status_label = ttk.Label(self, text="Importando...", bootstyle="info")
        self.import_status_label.pack(side="bottom", fill="x", padx=10, pady=(0, 5)); self.update_idletasks()
        threading.Thread(target=self._import_thread, args=(fp,), daemon=True).start()
    def _import_thread(self, fp):
        imp, skip, err = client_logic.import_clients_from_xlsx(fp); self.after(0, self.on_import_complete, imp, skip, err)
    def on_import_complete(self, imported, skipped, error):
        self.import_status_label.destroy()
        if error: messagebox.showerror("Erro", error, parent=self)
        else: messagebox.showinfo("Sucesso", f"Importação Concluída!\n\nImportados: {imported}\nIgnorados: {skipped}", parent=self); self.load_clients()

# ... (O restante do arquivo, com ClientCodeReportWindow e ReportPreviewWindow, permanece o mesmo) ...
class ClientCodeReportWindow(ttk.Toplevel):
    def __init__(self, controller, parent):
        super().__init__(title="Relatório de Códigos de Cliente", master=parent); self.controller = controller; self.geometry("500x300"); self.minsize(450, 250)
        self.transient(parent); self.grab_set(); self.total_clients_var = tk.StringVar(value="Carregando..."); self.last_added_var = tk.StringVar(value="Carregando...")
        self.create_widgets(); self.load_summary_data()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20); main_frame.pack(expand=True, fill="both")
        title_label = ttk.Label(main_frame, text="Gerador de Relatórios", font=("-size 14 -weight bold")); title_label.pack(pady=(0, 20))
        info_frame = ttk.Labelframe(main_frame, text=" Resumo da Base de Dados ", padding=15); info_frame.pack(fill="x", expand=True)
        ttk.Label(info_frame, text="Total de Clientes Cadastrados:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(info_frame, textvariable=self.total_clients_var, font="-weight bold").grid(row=0, column=1, sticky="w", padx=10)
        ttk.Label(info_frame, text="Último Cadastro Realizado em:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(info_frame, textvariable=self.last_added_var, font="-weight bold").grid(row=1, column=1, sticky="w", padx=10)
        self.generate_btn = ttk.Button(main_frame, text="Gerar Relatório", bootstyle="primary", command=self.open_preview_window, state="disabled")
        self.generate_btn.pack(pady=(20, 0), ipady=5, fill="x")
    def load_summary_data(self):
        threading.Thread(target=self._load_summary_thread, daemon=True).start()
    def _load_summary_thread(self):
        summary, error = client_logic.get_clients_summary(); self.after(0, self.on_summary_loaded, summary, error)
    def on_summary_loaded(self, summary, error):
        if error:
            self.total_clients_var.set("Erro"); self.last_added_var.set("Erro"); messagebox.showerror("Erro", f"Não foi possível carregar o resumo:\n{error}", parent=self); return
        self.total_clients_var.set(str(summary.get('total_clients', 0)))
        last_date = summary.get('last_added_date')
        if isinstance(last_date, datetime):
            brasilia_tz = pytz.timezone('America/Sao_Paulo'); local_time = last_date.astimezone(brasilia_tz)
            self.last_added_var.set(local_time.strftime("%d/%m/%Y às %H:%M"))
        else: self.last_added_var.set("Nenhum cliente cadastrado")
        self.generate_btn.config(state="normal")
    def open_preview_window(self):
        ReportPreviewWindow(self.controller, self)
class ReportPreviewWindow(ttk.Toplevel):
    def __init__(self, controller, parent):
        super().__init__(title="Pré-visualização do Relatório", master=parent)
        self.controller = controller; self.all_clients_cache = []; self.geometry("800x600"); self.minsize(600, 400); self.transient(parent); self.grab_set()
        self.search_var = tk.StringVar(); self.search_var.trace_add('write', self.filter_list)
        self.create_widgets(); self.load_clients()
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        main_frame.rowconfigure(1, weight=1); main_frame.columnconfigure(1, weight=1)
        toolbar = ttk.Frame(main_frame, padding=(0, 0, 10, 0)); toolbar.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 10))
        ttk.Label(toolbar, text="Exportar", font="-weight bold").pack(pady=(0,10))
        ttk.Button(toolbar, text="📄 PDF", command=self.export_pdf, bootstyle="danger").pack(fill="x", pady=2)
        ttk.Button(toolbar, text="📄 Word", command=self.export_word, bootstyle="info").pack(fill="x", pady=2)
        ttk.Button(toolbar, text="📄 Excel", command=self.export_excel, bootstyle="success").pack(fill="x", pady=2)
        controls_frame = ttk.Frame(main_frame); controls_frame.grid(row=0, column=1, sticky="ew", pady=(0, 10)); controls_frame.columnconfigure(1, weight=1)
        ttk.Label(controls_frame, text="Buscar:").pack(side="left"); ttk.Entry(controls_frame, textvariable=self.search_var).pack(side="left", fill="x", expand=True, padx=5)
        self.count_label = ttk.Label(controls_frame, text="Clientes: 0"); self.count_label.pack(side="left", padx=5)
        tree_frame = ttk.Frame(main_frame); tree_frame.grid(row=1, column=1, sticky="nsew"); tree_frame.rowconfigure(0, weight=1); tree_frame.columnconfigure(0, weight=1)
        self.client_tree = ttk.Treeview(tree_frame, columns=("name", "code"), show="headings")
        self.client_tree.heading("name", text="Nome do Cliente"); self.client_tree.heading("code", text="Código")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.client_tree.yview); self.client_tree.configure(yscrollcommand=scrollbar.set)
        self.client_tree.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
    def load_clients(self):
        threading.Thread(target=self._load_clients_thread, daemon=True).start()
    def _load_clients_thread(self):
        clients, error = client_logic.get_all_clients()
        if error: self.after(0, messagebox.showerror, "Erro", error, {"parent": self}); return
        self.all_clients_cache = clients; self.after(0, self.filter_list)
    def _populate_treeview(self, clients_to_display):
        self.client_tree.delete(*self.client_tree.get_children())
        for client in clients_to_display:
            self.client_tree.insert("", "end", iid=client['id'], values=(client.get('name'), client.get('code')))
        self.count_label.config(text=f"Clientes: {len(clients_to_display)}")
    def filter_list(self, *args):
        search_term = self.search_var.get().lower()
        filtered = self.all_clients_cache if not search_term else [c for c in self.all_clients_cache if search_term in c.get('name', '').lower()]
        self._populate_treeview(filtered)
    def _get_visible_clients(self):
        visible_ids = self.client_tree.get_children()
        return [c for c in self.all_clients_cache if c['id'] in visible_ids]
    def export_excel(self): report_logic.export_clients_to_excel(self._get_visible_clients(), self)
    def export_pdf(self): report_logic.export_clients_to_pdf(self._get_visible_clients(), self)
    def export_word(self): report_logic.export_clients_to_word(self._get_visible_clients(), self)