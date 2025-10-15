# =============================================================================
# --- ARQUIVO: ui/frames_support.py ---
# (Versão final com todas as melhorias do sistema de suporte)
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttk
from .dialogs_support import NewTicketDialog, TicketChatWindow
import support_logic
from datetime import datetime
import pytz
from PIL import Image, ImageTk

# --- TELA 1: TELA DE ESCOLHA DO TIPO DE SUPORTE (COM CARTÕES) ---
class SupportChoiceFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        
        self.support_options = [
            {
                "title": "Suporte do Sistema (Fiscal Flow)",
                "description": "Para reportar erros, sugerir melhorias ou tirar dúvidas sobre o funcionamento do programa.",
                "action": lambda: self.controller.show_support_tickets("developer"),
                "bootstyle": "primary"
            },
            {
                "title": "Suporte de T.I. (Infraestrutura)",
                "description": "Para solicitar ajuda com impressoras, rede, acesso a sistemas ou problemas no computador.",
                "action": lambda: self.controller.show_support_tickets("it"),
                "bootstyle": "info"
            }
        ]
        
        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        btn_voltar = ttk.Button(top_frame, text="< Voltar ao Menu Principal", command=lambda: self.controller.show_frame("HomeFrame"))
        btn_voltar.grid(row=0, column=0, sticky="w")
        
        title = ttk.Label(top_frame, text="Central de Suporte", font=("-size 18 -weight bold"))
        title.grid(row=0, column=1, pady=(0, 10))

        main_container = ttk.Frame(self, padding=(20, 10))
        main_container.grid(row=1, column=0, sticky="nsew")
        
        max_cols = 3
        for i, option in enumerate(self.support_options):
            row = i // max_cols
            col = i % max_cols
            main_container.columnconfigure(col, weight=1)

            card = ttk.Labelframe(main_container, text=option["title"], padding=20)
            card.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
            card.rowconfigure(0, weight=1)
            card.columnconfigure(0, weight=1)

            ttk.Label(card, text=option["description"], wraplength=300, justify="left").pack(expand=True, fill="both", pady=(0, 20))
            
            btn_acessar = ttk.Button(card, text="Acessar Canal", bootstyle=option["bootstyle"], command=option["action"])
            btn_acessar.pack(fill="x", ipady=5)


# --- TELA 2: TELA DE LISTAGEM DE CHAMADOS DO USUÁRIO ---
class UserTicketsFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.ticket_type = None
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.tickets_on_display = []
        self.brasilia_tz = pytz.timezone('America/Sao_Paulo')
        self._icons = {}
        self.load_icons()
        self.create_widgets()

    def load_icons(self):
        icon_data = {"Red": "red.png", "Yellow": "yellow.png", "Green": "green.png", "Gray": "gray.png"}
        for status, filename in icon_data.items():
            try:
                path = self.controller.resource_path(f'assets/icons/{filename}')
                img = Image.open(path).resize((16, 16), Image.LANCZOS)
                self._icons[status] = ImageTk.PhotoImage(img)
            except Exception as e:
                print(f"Erro ao carregar ícone {filename}: {e}")
                self._icons[status] = None

    def create_widgets(self):
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        btn_voltar = ttk.Button(top_frame, text="< Voltar", command=lambda: self.controller.show_frame("SupportChoiceFrame"))
        btn_voltar.grid(row=0, column=0, sticky="w")
        
        self.title_label = ttk.Label(top_frame, text="Meus Chamados", font=("-size 18 -weight bold"))
        self.title_label.grid(row=0, column=1, pady=(0, 10))

        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        tree_container = ttk.Frame(main_frame)
        tree_container.pack(expand=True, fill="both")

        self.tree = ttk.Treeview(tree_container, show='tree headings')
        cols = ('Assunto', 'Status', 'Última Atualização')
        self.tree['columns'] = cols
        
        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", expand=True, fill="both")

        self.tree.column("#0", width=50, stretch=False, anchor='center')
        # --- [ALTERADO] Adicionado título à coluna de ícones ---
        self.tree.heading("#0", text="Status")

        self.tree.heading('Assunto', text='Assunto')
        self.tree.heading('Status', text='Status')
        self.tree.column('Status', width=120, stretch=False, anchor="center")
        self.tree.heading('Última Atualização', text='Última Atualização')
        self.tree.column('Última Atualização', width=150, stretch=False, anchor="center")

        strikethrough_font = ttk.font.Font(family="Helvetica", size=10, overstrike=True)
        self.tree.tag_configure("Fechado", font=strikethrough_font, foreground="gray")

        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0))
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Abrir Novo Chamado", command=self.open_new_ticket_dialog, bootstyle="success").pack(side="left")
        ttk.Button(btn_frame, text="Ver Chamado Selecionado", command=self.open_selected_ticket).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Atualizar Lista", command=lambda: self.refresh_tickets_list(force=True), bootstyle="secondary-outline").pack(side="right")

    def on_frame_activated(self):
        self.refresh_tickets_list()

    def set_ticket_type(self, ticket_type):
        self.ticket_type = ticket_type
        title_text = "Suporte do Sistema" if ticket_type == "developer" else "Suporte de T.I."
        self.title_label.config(text=f"Meus Chamados - {title_text}")

    def refresh_tickets_list(self, force=False):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        if not self.ticket_type or not self.controller.current_user: return
        user_id = self.controller.current_user['id']
        tickets, error = support_logic.get_tickets_for_user(user_id, self.ticket_type)
        self.tickets_on_display = tickets
        if error: messagebox.showerror("Erro", error, parent=self); return
        
        for ticket in self.tickets_on_display:
            last_update = ticket['last_updated_at']
            last_update_str = last_update.astimezone(self.brasilia_tz).strftime("%d/%m/%Y %H:%M") if isinstance(last_update, datetime) else "Carregando..."
            
            status_icon = self._icons.get(ticket.get('flag_color', 'Gray')) if self.ticket_type == 'it' else None
            tags_to_apply = ('Fechado',) if ticket.get('status') == 'Fechado' else ()
            
            options = { "iid": ticket['id'], "values": [ ticket['subject'], ticket['status'], last_update_str ], "tags": tags_to_apply }
            if status_icon: options["image"] = status_icon
            
            self.tree.insert('', 'end', **options)

    def open_new_ticket_dialog(self):
        NewTicketDialog(self.controller, self, self.ticket_type)

    def open_selected_ticket(self):
        selected_item = self.tree.focus()
        if not selected_item: messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um chamado na lista.", parent=self); return
        ticket_id = selected_item
        ticket_data = next((t for t in self.tickets_on_display if t['id'] == ticket_id), None)
        if ticket_data:
            ticket_data['ticket_type'] = self.ticket_type
            TicketChatWindow(self.controller, self, ticket_data)

# --- TELA 3: TELA DE LISTAGEM DE CHAMADOS PARA ADMINISTRADORES ---
class AdminTicketsFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.ticket_type = None
        self.columnconfigure(0, weight=1); self.rowconfigure(1, weight=1)
        self.tickets_on_display = []; self.all_tickets = []
        self.brasilia_tz = pytz.timezone('America/Sao_Paulo')
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_list)
        self._icons = {}; self.load_icons()
        self.create_widgets()

    def load_icons(self):
        icon_data = {"Red": "red.png", "Yellow": "yellow.png", "Green": "green.png", "Gray": "gray.png"}
        for status, filename in icon_data.items():
            try:
                path = self.controller.resource_path(f'assets/icons/{filename}')
                img = Image.open(path).resize((16, 16), Image.LANCZOS)
                self._icons[status] = ImageTk.PhotoImage(img)
            except Exception as e:
                print(f"Erro ao carregar ícone {filename}: {e}"); self._icons[status] = None

    def create_widgets(self):
        top_frame = ttk.Frame(self, padding=(10, 10, 10, 0)); top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)
        ttk.Button(top_frame, text="< Voltar", command=lambda: self.controller.show_frame("SupportChoiceFrame")).grid(row=0, column=0, sticky="w")
        self.title_label = ttk.Label(top_frame, text="Painel de Atendimento", font=("-size 18 -weight bold")); self.title_label.grid(row=0, column=1, pady=(0, 10))

        main_frame = ttk.Frame(self, padding=10); main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1); main_frame.rowconfigure(1, weight=1)
        
        filter_frame = ttk.Frame(main_frame); filter_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(filter_frame, text="Filtrar por Solicitante:").pack(side="left", padx=(0, 5))
        ttk.Entry(filter_frame, textvariable=self.search_var).pack(side="left", fill="x", expand=True)

        tree_container = ttk.Frame(main_frame); tree_container.pack(expand=True, fill="both")
        
        self.tree = ttk.Treeview(tree_container, show="tree headings")
        cols = ('Solicitante', 'Assunto', 'Status', 'Última Atualização')
        self.tree['columns'] = cols

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y"); self.tree.pack(side="left", expand=True, fill="both")

        self.tree.column("#0", width=50, stretch=False, anchor='center')
        # --- [ALTERADO] Adicionado título à coluna de ícones ---
        self.tree.heading("#0", text="Status")

        self.tree.heading('Solicitante', text='Solicitante'); self.tree.column('Solicitante', width=150, stretch=False)
        self.tree.heading('Assunto', text='Assunto')
        self.tree.heading('Status', text='Status'); self.tree.column('Status', width=120, stretch=False, anchor="center")
        self.tree.heading('Última Atualização', text='Última Atualização'); self.tree.column('Última Atualização', width=150, stretch=False, anchor="center")

        strikethrough_font = ttk.font.Font(family="Helvetica", size=10, overstrike=True)
        self.tree.tag_configure("Fechado", font=strikethrough_font, foreground="gray")

        bottom_frame = ttk.Frame(main_frame); bottom_frame.pack(fill="x", pady=(10,0))
        self.legend_frame = ttk.Frame(bottom_frame)
        btn_frame = ttk.Frame(bottom_frame); btn_frame.pack(side="right")
        ttk.Button(btn_frame, text="Ver Chamado", command=self.open_selected_ticket, bootstyle="primary").pack(side="left")
        ttk.Button(btn_frame, text="Excluir Chamado", command=self.delete_selected_ticket, bootstyle="danger").pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Atualizar Lista", command=lambda: self.refresh_tickets_list(force=True), bootstyle="secondary-outline").pack(side="right")

    def on_frame_activated(self):
        self.refresh_tickets_list()

    def set_ticket_type(self, ticket_type):
        self.ticket_type = ticket_type
        title_text = "Suporte do Sistema" if ticket_type == "developer" else "Suporte de T.I."
        self.title_label.config(text=f"Painel de Atendimento - {title_text}")

    def refresh_tickets_list(self, force=False):
        if not self.ticket_type or not self.controller.current_user: return
        tickets, error = support_logic.get_all_tickets(self.ticket_type)
        self.all_tickets = tickets
        if error: messagebox.showerror("Erro", error, parent=self); return
        self.filter_list()

    def populate_tree(self, tickets_to_show):
        for i in self.tree.get_children(): self.tree.delete(i)
        self.tickets_on_display = tickets_to_show
        
        for ticket in self.tickets_on_display:
            last_update = ticket['last_updated_at']
            last_update_str = last_update.astimezone(self.brasilia_tz).strftime("%d/%m/%Y %H:%M") if isinstance(last_update, datetime) else "Carregando..."
            
            status_icon = self._icons.get(ticket.get('flag_color', 'Gray')) if self.ticket_type == 'it' else None
            tags_to_apply = ('Fechado',) if ticket.get('status') == 'Fechado' else ()
            
            options = {
                "iid": ticket['id'],
                "values": [
                    ticket.get('username', 'N/A'),
                    ticket.get('subject', 'N/A'),
                    ticket.get('status', 'N/A'),
                    last_update_str
                ],
                "tags": tags_to_apply
            }
            
            if status_icon:
                options["image"] = status_icon

            self.tree.insert('', 'end', **options)

    def filter_list(self, *args):
        search_term = self.search_var.get().lower()
        if not search_term: self.populate_tree(self.all_tickets); return
        filtered_list = [t for t in self.all_tickets if search_term in str(t.get('username')).lower()]
        self.populate_tree(filtered_list)

    def open_selected_ticket(self):
        selected_item = self.tree.focus()
        if not selected_item: messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um chamado na lista.", parent=self); return
        ticket_id = selected_item
        ticket_data = next((t for t in self.tickets_on_display if t['id'] == ticket_id), None)
        if ticket_data:
            ticket_data['ticket_type'] = self.ticket_type
            TicketChatWindow(self.controller, self, ticket_data)
            
    def delete_selected_ticket(self):
        selected_item = self.tree.focus()
        if not selected_item: messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um chamado para remover.", parent=self); return
        ticket_id = selected_item
        
        ticket_subject = self.tree.item(selected_item, 'values')[1]
        
        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja remover o chamado '{ticket_subject}'?\n\nEsta ação não pode ser desfeita.", parent=self):
            success, message = support_logic.delete_ticket(ticket_id)
            if success: messagebox.showinfo("Sucesso", message, parent=self); self.refresh_tickets_list(force=True)
            else: messagebox.showerror("Erro", message, parent=self)