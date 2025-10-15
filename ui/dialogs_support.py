# =============================================================================
# --- ARQUIVO: ui/dialogs_support.py ---
# (Corrigido para usar link de download direto do Google Drive)
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import ttkbootstrap as ttk
import support_logic
import drive_logic
import threading
from datetime import datetime
import pytz
import uuid
import os
import webbrowser

# --- Importações para a pré-visualização de imagens ---
import io
try:
    import requests
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror(
        "Bibliotecas Faltando",
        "As bibliotecas 'requests' e 'Pillow' são necessárias para a pré-visualização de imagens.\n"
        "Por favor, instale-as executando no terminal:\n"
        "pip install requests pillow"
    )

class NewTicketDialog(ttk.Toplevel):
    """Janela de formulário para criar um novo ticket de suporte."""
    def __init__(self, controller, parent_window, ticket_type):
        super().__init__(title="Abrir Novo Chamado de Suporte", master=parent_window)
        self.controller = controller
        self.parent_window = parent_window
        self.ticket_type = ticket_type
        self.minsize(500, 400)
        self.transient(parent_window); self.grab_set()
        self.subject_var = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        main_frame.rowconfigure(3, weight=1); main_frame.columnconfigure(0, weight=1)
        ttk.Label(main_frame, text="Assunto:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Entry(main_frame, textvariable=self.subject_var).grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(main_frame, text="Descreva sua solicitação:").grid(row=2, column=0, sticky="w", pady=(0, 5))
        self.message_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=10)
        self.message_text.grid(row=3, column=0, sticky="nsew"); self.message_text.focus_set()
        btn_frame = ttk.Frame(main_frame, padding=(0, 15, 0, 0)); btn_frame.grid(row=4, column=0, sticky="e")
        self.send_btn = ttk.Button(btn_frame, text="Enviar Chamado", command=self.submit_ticket, bootstyle="primary")
        self.send_btn.pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def submit_ticket(self):
        subject = self.subject_var.get()
        first_message = self.message_text.get("1.0", "end-1c")
        if not all([subject, first_message]):
            messagebox.showerror("Campos Vazios", "O assunto e a mensagem são obrigatórios.", parent=self); return
        self.send_btn.config(state="disabled", text="Enviando..."); self.update_idletasks()
        user_info = self.controller.current_user
        threading.Thread(target=self._submit_ticket_thread, args=(user_info['id'], user_info['username'], subject, first_message, self.ticket_type), daemon=True).start()
    
    def _submit_ticket_thread(self, user_id, username, subject, first_message, ticket_type):
        ticket_id, message = support_logic.create_ticket(user_id, username, subject, first_message, ticket_type)
        self.after(0, self.on_submit_complete, ticket_id is not None, message)

    def on_submit_complete(self, success, message):
        self.send_btn.config(state="normal", text="Enviar Chamado")
        if success:
            messagebox.showinfo("Sucesso", message, parent=self.parent_window)
            if hasattr(self.parent_window, 'refresh_tickets_list'):
                self.parent_window.refresh_tickets_list(force=True)
            self.destroy()
        else:
            messagebox.showerror("Erro", message, parent=self)

class TicketChatWindow(ttk.Toplevel):
    def __init__(self, controller, parent_window, ticket_data):
        super().__init__(title=f"Chamado: {ticket_data.get('subject')}", master=parent_window)
        self.controller = controller
        self.parent_window = parent_window
        self.ticket_data = ticket_data
        self.brasilia_tz = pytz.timezone('America/Sao_Paulo')
        self.image_references = []
        self.flag_map = {"Red": "Urgente", "Yellow": "Média", "Green": "Baixa", "Gray": "Resolvido"}
        self.flag_map_reverse = {v: k for k, v in self.flag_map.items()}
        self.selected_flag_color = self.ticket_data.get('flag_color', 'Gray')
        self.initial_status = self.ticket_data.get('status', 'Aberto')
        self.selected_category_id = self.ticket_data.get('category_id')
        self.selected_category_name = self.ticket_data.get('category_name', 'Não definida')
        self.minsize(700, 550)
        self.transient(parent_window); self.grab_set()
        self.create_widgets()
        self.load_messages()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        main_frame.rowconfigure(1, weight=1); main_frame.columnconfigure(0, weight=1)
        
        current_user_level = self.controller.current_user.get('level')
        is_attendant = current_user_level in ['Admin', 'Desenvolvedor', 'T.I.']

        if is_attendant:
            self.create_admin_panel(main_frame)

        chat_container = ttk.Frame(main_frame); chat_container.grid(row=1, column=0, sticky="nsew", pady=5)
        chat_container.rowconfigure(0, weight=1); chat_container.columnconfigure(0, weight=1)
        self.chat_canvas = tk.Canvas(chat_container, bd=0, highlightthickness=0); self.chat_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(chat_container, orient="vertical", command=self.chat_canvas.yview); scrollbar.grid(row=0, column=1, sticky="ns")
        self.chat_canvas.configure(yscrollcommand=scrollbar.set)
        self.chat_frame = ttk.Frame(self.chat_canvas, padding=(10, 0, 20, 0))
        self.canvas_frame_id = self.chat_canvas.create_window((0, 0), window=self.chat_frame, anchor="nw")
        self.chat_frame.bind("<Configure>", self.on_frame_configure)
        self.chat_canvas.bind("<Configure>", self.on_canvas_configure)
        
        send_frame = ttk.Frame(main_frame); send_frame.grid(row=2, column=0, sticky="ew", pady=(5,0)); send_frame.columnconfigure(0, weight=1)
        self.attach_btn = ttk.Button(send_frame, text="📎", command=self.attach_file, bootstyle="secondary", width=2); self.attach_btn.pack(side="left", padx=(0, 5))
        self.message_entry = ttk.Entry(send_frame, font=("Helvetica", 10)); self.message_entry.pack(side="left", expand=True, fill="x"); self.message_entry.focus_set()
        self.message_entry.bind("<Return>", self.send_message)
        self.message_entry.bind("<Control-v>", self.paste_from_clipboard)
        self.send_btn = ttk.Button(send_frame, text="Enviar", command=self.send_message, bootstyle="primary"); self.send_btn.pack(side="left", padx=(10, 0))

    def on_frame_configure(self, event=None):
        self.chat_canvas.configure(scrollregion=self.chat_canvas.bbox("all"))
    def on_canvas_configure(self, event=None):
        self.chat_canvas.itemconfig(self.canvas_frame_id, width=event.width)
    def scroll_to_bottom(self):
        self.chat_canvas.update_idletasks(); self.chat_canvas.yview_moveto(1.0)

    def create_admin_panel(self, parent):
        admin_panel = ttk.Labelframe(parent, text=" Painel do Atendente ", padding=10)
        admin_panel.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        admin_panel.columnconfigure(3, weight=1)
        status_frame = ttk.Frame(admin_panel); status_frame.grid(row=0, column=0, sticky='w', padx=(0,15))
        ttk.Label(status_frame, text="Status:").pack(side="left", padx=(0,5))
        status_options = ["Aberto", "Em Andamento", "Aguardando Resposta", "Fechado"]
        self.status_combo = ttk.Combobox(status_frame, values=status_options, state="readonly", width=18)
        self.status_combo.set(self.initial_status)
        self.status_combo.pack(side="left")
        self.status_combo.bind("<<ComboboxSelected>>", self.on_change)
        
        # --- [MODIFICADO] A condição "if self.ticket_data.get('ticket_type') == 'it':" foi REMOVIDA daqui ---
        flag_frame = ttk.Frame(admin_panel); flag_frame.grid(row=0, column=1, sticky='w', padx=(0,15))
        ttk.Label(flag_frame, text="Classificar:").pack(side="left", padx=(0, 5))
        self.icon_preview_label = ttk.Label(flag_frame); self.icon_preview_label.pack(side="left", padx=(0, 5))
        self.flag_combo_var = tk.StringVar()
        self.flag_combo = ttk.Combobox(flag_frame, textvariable=self.flag_combo_var, values=list(self.flag_map.values()), state="readonly", width=12)
        self.flag_combo.pack(side="left")
        self.flag_combo.bind("<<ComboboxSelected>>", self.on_flag_combo_select)
        self.update_flag_display()
        # --- Fim da Modificação ---
        
        category_frame = ttk.Frame(admin_panel); category_frame.grid(row=0, column=2, sticky='w')
        ttk.Label(category_frame, text="Categoria:").pack(side="left", padx=(0,5))
        self.category_label = ttk.Label(category_frame, text=self.selected_category_name, width=20, anchor='w')
        self.category_label.pack(side="left")
        ttk.Button(category_frame, text="Alterar", command=self.open_category_selector, bootstyle="outline-secondary").pack(side="left", padx=5)
        self.save_admin_btn = ttk.Button(admin_panel, text="Salvar Alterações", command=self.save_admin_changes, state="disabled")
        self.save_admin_btn.grid(row=0, column=3, sticky='e')

    def open_category_selector(self):
        SelectCategoryDialog(self, self.controller, self.ticket_data['ticket_type'], self.on_category_selected)
    def on_category_selected(self, category_id, category_name):
        self.selected_category_id = category_id; self.selected_category_name = category_name
        self.category_label.config(text=category_name); self.on_change()
    def on_change(self, event=None):
        self.save_admin_btn.config(state="normal")
    def update_flag_display(self):
        if hasattr(self.parent_window, '_icons'):
            icon = self.parent_window._icons.get(self.selected_flag_color); self.icon_preview_label.config(image=icon)
        display_text = self.flag_map.get(self.selected_flag_color, "Resolvido"); self.flag_combo_var.set(display_text)
    def on_flag_combo_select(self, event=None):
        selected_text = self.flag_combo_var.get()
        self.selected_flag_color = self.flag_map_reverse.get(selected_text, "Gray")
        self.update_flag_display(); self.on_change()

    def save_admin_changes(self):
        new_status = self.status_combo.get(); new_color = self.selected_flag_color; ticket_id = self.ticket_data.get('id')
        self.save_admin_btn.config(state="disabled", text="Salvando..."); self.update_idletasks()
        threading.Thread(target=self._save_admin_thread, args=(ticket_id, new_status, new_color, self.selected_category_id, self.selected_category_name), daemon=True).start()
    def _save_admin_thread(self, ticket_id, new_status, new_color, cat_id, cat_name):
        success, message = support_logic.update_ticket_details(ticket_id, new_status, new_color, cat_id, cat_name)
        self.after(0, self.on_admin_save_complete, success, message)
    def on_admin_save_complete(self, success, message):
        self.save_admin_btn.config(text="Salvar Alterações")
        if success:
            if hasattr(self.parent_window, 'refresh_tickets_list'): self.parent_window.refresh_tickets_list(force=True)
            self.initial_status = self.status_combo.get()
            self.ticket_data['flag_color'] = self.selected_flag_color
            self.ticket_data['category_id'] = self.selected_category_id
            self.ticket_data['category_name'] = self.selected_category_name
        else:
            messagebox.showerror("Erro", message, parent=self); self.save_admin_btn.config(state="normal")

    def load_messages(self):
        for widget in self.chat_frame.winfo_children(): widget.destroy()
        ticket_id = self.ticket_data.get('id')
        messages, error = support_logic.get_messages_for_ticket(ticket_id)
        if error: messagebox.showerror("Erro", error, parent=self); return
        for msg in messages: self.display_message(msg)
        self.scroll_to_bottom()

    def display_message(self, msg_data):
        is_user_message = msg_data.get('sender_name') == self.controller.current_user['username']
        msg_container = ttk.Frame(self.chat_frame)
        sender = msg_data.get('sender_name', 'Desconhecido'); timestamp = msg_data.get('timestamp')
        if isinstance(timestamp, datetime): time_str = timestamp.astimezone(self.brasilia_tz).strftime("%d/%m/%Y %H:%M")
        else: time_str = "enviando..."
        header_text = f"{sender} - {time_str}"
        ttk.Label(msg_container, text=header_text, font=("Helvetica", 9, "bold")).pack(anchor="w", padx=5, pady=(5,0))
        text = msg_data.get('text', '')
        if text: ttk.Label(msg_container, text=text, wraplength=500, justify="left", font=("Helvetica", 10)).pack(anchor="w", padx=5, pady=2)
        attachment_url = msg_data.get('attachment_url')
        if attachment_url:
            filename = msg_data.get('attachment_filename', 'anexo')
            is_image = any(filename.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif'])
            if is_image:
                placeholder = ttk.Label(msg_container, text="🖼️ Carregando imagem...", style="secondary.TLabel"); placeholder.pack(anchor="w", padx=5, pady=5)
                threading.Thread(target=self.load_image_from_url, args=(attachment_url, placeholder), daemon=True).start()
            else:
                link_label = ttk.Label(msg_container, text=f"📎 {filename}", style="info.TLabel", cursor="hand2"); link_label.pack(anchor="w", padx=5, pady=5)
                link_label.bind("<Button-1>", lambda e, url=attachment_url: webbrowser.open(url))
        msg_container.pack(fill="x", pady=2, padx=5, anchor="w" if not is_user_message else "e")

    def load_image_from_url(self, url, placeholder_widget):
        try:
            file_id = url.split('/d/')[1].split('/')[0]
            direct_download_url = f'https://drive.google.com/uc?export=download&id={file_id}'
            response = requests.get(direct_download_url, timeout=15); response.raise_for_status()
            image_data = io.BytesIO(response.content); pil_image = Image.open(image_data)
            pil_image.thumbnail((250, 250)); photo_image = ImageTk.PhotoImage(pil_image)
            self.image_references.append(photo_image)
            self.after(0, self.update_image_widget, placeholder_widget, photo_image, url)
        except Exception as e:
            print(f"Erro ao carregar imagem: {e}")
            self.after(0, placeholder_widget.config, {"text": "❌ Falha ao carregar imagem"})

    def update_image_widget(self, placeholder, image, url):
        placeholder.config(image=image, text="", cursor="hand2")
        placeholder.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))

    def attach_file(self):
        filepath = filedialog.askopenfilename()
        if not filepath: return
        self.set_sending_state(True, "Enviando anexo...")
        threading.Thread(target=self._upload_thread, args=(filepath,), daemon=True).start()

    def paste_from_clipboard(self, event=None):
        try: image = Image.open(io.BytesIO(self.clipboard_get(type='PNG')))
        except Exception: image = None
        if image:
            self.set_sending_state(True, "Enviando imagem...")
            temp_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", "FiscalFlow")
            os.makedirs(temp_dir, exist_ok=True)
            temp_filepath = os.path.join(temp_dir, f"paste_{uuid.uuid4()}.png")
            image.save(temp_filepath, 'PNG')
            threading.Thread(target=self._upload_thread, args=(temp_filepath, True), daemon=True).start()
            return "break"
        
    def _upload_thread(self, filepath, delete_after=False):
        original_filename = os.path.basename(filepath)
        url, _, error = drive_logic.upload_attachment(filepath, original_filename)
        if delete_after:
            try: os.remove(filepath)
            except Exception as e: print(f"Não foi possível remover arquivo temporário: {e}")
        self.after(0, self.on_upload_complete, url, original_filename, error)

    def on_upload_complete(self, url, filename, error):
        if error: messagebox.showerror("Erro de Upload", error, parent=self); self.set_sending_state(False); return
        self.send_message(attachment_info={'url': url, 'filename': filename})

    def send_message(self, event=None, attachment_info=None):
        text = self.message_entry.get()
        if not text.strip() and not attachment_info: return
        self.set_sending_state(True)
        ticket_id = self.ticket_data.get('id'); user_info = self.controller.current_user
        url = attachment_info.get('url') if attachment_info else None
        filename = attachment_info.get('filename') if attachment_info else None
        threading.Thread(target=self._send_message_thread, args=(ticket_id, user_info['id'], user_info['username'], user_info['level'], text, url, filename), daemon=True).start()

    def _send_message_thread(self, ticket_id, user_id, username, user_level, text, url, filename):
        success, message = support_logic.add_message_to_ticket(ticket_id, user_id, username, user_level, text, url, filename)
        self.after(0, self.on_message_sent, success, message)

    def on_message_sent(self, success, message):
        self.set_sending_state(False)
        if success:
            self.message_entry.delete(0, 'end'); self.load_messages()
            if hasattr(self.parent_window, 'refresh_tickets_list'): self.parent_window.refresh_tickets_list(force=True)
        else: messagebox.showerror("Erro", message, parent=self)

    def set_sending_state(self, is_sending, custom_text=None):
        state = "disabled" if is_sending else "normal"
        text = "Enviando..." if is_sending and not custom_text else "Enviar"
        if custom_text: text = custom_text
        self.send_btn.config(state=state, text=text); self.attach_btn.config(state=state); self.message_entry.config(state=state)

class SelectCategoryDialog(ttk.Toplevel):
    def __init__(self, parent, controller, ticket_type, callback):
        super().__init__(title="Selecionar Categoria", master=parent)
        self.parent = parent; self.controller = controller; self.ticket_type = ticket_type; self.callback = callback
        self.all_categories = []; self.minsize(400, 300); self.transient(parent); self.grab_set()
        self.create_widgets(); self.load_categories()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        main_frame.rowconfigure(1, weight=1); main_frame.columnconfigure(0, weight=1)
        search_frame = ttk.Frame(main_frame); search_frame.grid(row=0, column=0, sticky='ew', pady=(0,5))
        ttk.Label(search_frame, text="Pesquisar:").pack(side='left')
        self.search_var = tk.StringVar(); self.search_var.trace_add('write', self.filter_list)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side='left', fill='x', expand=True, padx=5)
        self.listbox = tk.Listbox(main_frame, selectmode='browse'); self.listbox.grid(row=1, column=0, sticky='nsew')
        self.listbox.bind('<Double-1>', self.on_select)
        btn_frame = ttk.Frame(main_frame, padding=(0,10,0,0)); btn_frame.grid(row=2, column=0, sticky='e')
        ttk.Button(btn_frame, text="Confirmar", command=self.on_select, bootstyle='primary').pack(side='right')
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(side='right', padx=10)
    
    def load_categories(self):
        self.all_categories, error = support_logic.get_categories_for_type(self.ticket_type)
        if error: messagebox.showerror("Erro", error, parent=self)
        self.filter_list()

    def filter_list(self, *args):
        self.listbox.delete(0, 'end')
        search_term = self.search_var.get().strip(); search_lower = search_term.lower()
        filtered_list = [cat for cat in self.all_categories if search_lower in cat['name'].lower()]
        if search_term and not any(cat['name'].lower() == search_lower for cat in self.all_categories):
            self.listbox.insert('end', f"Adicionar nova categoria: '{search_term}'")
        for category in filtered_list: self.listbox.insert('end', category['name'])

    def on_select(self, event=None):
        selection = self.listbox.curselection()
        if not selection: return
        selected_text = self.listbox.get(selection[0])
        if selected_text.startswith("Adicionar nova categoria:"):
            new_name = self.search_var.get().strip()
            new_id, message = support_logic.add_category(new_name, self.ticket_type)
            if new_id: self.callback(new_id, new_name); self.destroy()
            else: messagebox.showerror("Erro", message, parent=self)
        else:
            category = next((cat for cat in self.all_categories if cat['name'] == selected_text), None)
            if category: self.callback(category['id'], category['name']); self.destroy()