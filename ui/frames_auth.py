# ===================================================================================
# --- ARQUIVO: ui/frames_auth.py ---
# (Contém as telas de Login, Autenticação e Boas-vindas)
# ===================================================================================

import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext, simpledialog
from collections import defaultdict
from PIL import Image, ImageTk, ImageEnhance
import ttkbootstrap as ttk
from datetime import datetime
import logging
import threading

import auth_logic

# --- TELA DE LOGIN ---
class LoginFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(0, weight=1)
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.show_password_var = tk.BooleanVar(value=False)
        self.create_widgets()

    def create_widgets(self):
        main_container = ttk.Frame(self, padding=40)
        main_container.grid(row=0, column=0)
        title = ttk.Label(main_container, text="Fiscal Flow - Acesso", font=("-size 20 -weight bold"))
        title.pack(pady=(0, 30))
        form_frame = ttk.Frame(main_container)
        form_frame.pack(fill="x")
        user_label = ttk.Label(form_frame, text="Usuário:")
        user_label.pack(fill="x")
        user_entry = ttk.Entry(form_frame, textvariable=self.username_var, width=40)
        user_entry.pack(fill="x", pady=(5, 15))
        user_entry.focus_set()
        pass_label = ttk.Label(form_frame, text="Senha:")
        pass_label.pack(fill="x")
        self.pass_entry = ttk.Entry(form_frame, textvariable=self.password_var, show="*", width=40)
        self.pass_entry.pack(fill="x", pady=5)
        
        show_pass_check = ttk.Checkbutton(form_frame, text="Mostrar Senha", variable=self.show_password_var, command=self._toggle_password_visibility)
        show_pass_check.pack(anchor="w", pady=(5,0))
        
        user_entry.bind("<Return>", lambda event: self.attempt_login())
        self.pass_entry.bind("<Return>", lambda event: self.attempt_login())
        
        button_container = ttk.Frame(main_container)
        button_container.pack(fill="x", pady=(30, 10))
        button_container.columnconfigure(0, weight=1)
        
        self.login_button = ttk.Button(button_container, text="Entrar", command=self.attempt_login, bootstyle="primary")
        self.login_button.grid(row=0, column=0, sticky="ew", pady=5)
        
        links_frame = ttk.Frame(button_container)
        links_frame.grid(row=1, column=0, pady=5)
        links_frame.columnconfigure(0, weight=1)
        links_frame.columnconfigure(1, weight=1)

        setup_password_button = ttk.Button(links_frame, text="Primeiro Acesso", command=lambda: self.controller.show_frame("SetPasswordFrame"), bootstyle="link")
        setup_password_button.grid(row=0, column=0, sticky="w")

        forgot_password_button = ttk.Button(links_frame, text="Esqueci minha senha", command=self.forgot_password, bootstyle="link")
        forgot_password_button.grid(row=0, column=1, sticky="e")
        
        request_access_button = ttk.Button(button_container, text="Solicitar Acesso", command=lambda: self.controller.open_request_access_dialog(), bootstyle="link")
        request_access_button.grid(row=2, column=0, pady=5)
        
        self.status_label = ttk.Label(main_container, text="", bootstyle="danger")
        self.status_label.pack(pady=(10, 0))

    def _toggle_password_visibility(self):
        self.pass_entry.config(show="" if self.show_password_var.get() else "*")

    def attempt_login(self):
        self.status_label.config(text="")
        username, password = self.username_var.get(), self.password_var.get()
        if not username or not password:
            self.status_label.config(text="Usuário e senha são obrigatórios.")
            return

        self.login_button.config(state="disabled")
        self.status_label.config(text="Verificando...", bootstyle="info")
        self.controller.config(cursor="watch")
        self.controller.update_idletasks()

        try:
            user_data = auth_logic.verify_user(username, password)
            if user_data:
                self.username_var.set(""), self.password_var.set("")
                self.controller.on_login_success(user_data)
            else:
                self.status_label.config(text="Usuário ou senha inválidos.", bootstyle="danger")
                self.password_var.set("")
        except Exception as e:
            self.status_label.config(text=f"Erro de conexão: {e}", bootstyle="danger")
            logging.error(f"Erro de conexão no login: {e}")
        finally:
            self.login_button.config(state="normal")
            self.controller.config(cursor="")

    def forgot_password(self):
        email = simpledialog.askstring("Recuperar Senha", "Digite o seu e-mail de cadastro:", parent=self)
        if not email:
            return

        self.status_label.config(text="Enviando e-mail de recuperação...", bootstyle="info")
        self.controller.config(cursor="watch")
        self.controller.update_idletasks()
        
        threading.Thread(target=self._forgot_password_thread, args=(email,), daemon=True).start()

    def _forgot_password_thread(self, email):
        success, message = auth_logic.request_password_reset(email)
        self.after(0, self.on_forgot_password_complete, success, message, email)

    def on_forgot_password_complete(self, success, message, email):
        self.controller.config(cursor="")
        if success:
            messagebox.showinfo("E-mail Enviado", message, parent=self.controller)
            self.controller.frames["ResetPasswordFrame"].prepare_for_email(email)
            self.controller.show_frame("ResetPasswordFrame")
        else:
            messagebox.showerror("Erro", message, parent=self.controller)
        self.status_label.config(text="")

# --- TELA DE DEFINIR SENHA ---
class SetPasswordFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(0, weight=1)
        self.username_var, self.token_var = tk.StringVar(), tk.StringVar()
        self.new_password_var, self.confirm_password_var = tk.StringVar(), tk.StringVar()
        self.show_passwords_var = tk.BooleanVar(value=False)
        self.create_widgets()

    def create_widgets(self):
        main_container = ttk.Frame(self, padding=40)
        main_container.grid(row=0, column=0)
        title = ttk.Label(main_container, text="Configurar Senha", font=("-size 20 -weight bold"))
        title.pack(pady=(0, 30))
        form_frame = ttk.Frame(main_container)
        form_frame.pack(fill="x")
        ttk.Label(form_frame, text="Usuário:").pack(fill="x")
        ttk.Entry(form_frame, textvariable=self.username_var, width=40).pack(fill="x", pady=(5, 15))
        ttk.Label(form_frame, text="Código de Verificação (recebido por e-mail):").pack(fill="x")
        ttk.Entry(form_frame, textvariable=self.token_var, width=40).pack(fill="x", pady=(5, 15))
        ttk.Label(form_frame, text="Nova Senha:").pack(fill="x")
        self.new_pass_entry = ttk.Entry(form_frame, textvariable=self.new_password_var, show="*", width=40)
        self.new_pass_entry.pack(fill="x", pady=(5, 15))
        ttk.Label(form_frame, text="Confirmar Nova Senha:").pack(fill="x")
        self.confirm_pass_entry = ttk.Entry(form_frame, textvariable=self.confirm_password_var, show="*", width=40)
        self.confirm_pass_entry.pack(fill="x", pady=5)
        
        show_pass_check = ttk.Checkbutton(form_frame, text="Mostrar Senhas", variable=self.show_passwords_var, command=self._toggle_passwords_visibility)
        show_pass_check.pack(anchor="w", pady=(5,0))

        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=(30, 10))
        self.set_pass_button = ttk.Button(button_frame, text="Definir Senha", command=self.set_password, bootstyle="primary")
        self.set_pass_button.pack(side="right")
        back_button = ttk.Button(button_frame, text="Voltar para o Login", command=lambda: self.controller.show_frame("LoginFrame"), bootstyle="secondary")
        back_button.pack(side="left")
        self.status_label = ttk.Label(main_container, text="", bootstyle="danger")
        self.status_label.pack(pady=(10, 0))

    def _toggle_passwords_visibility(self):
        show_char = "" if self.show_passwords_var.get() else "*"
        self.new_pass_entry.config(show=show_char)
        self.confirm_pass_entry.config(show=show_char)

    def set_password(self):
        self.status_label.config(text="")
        username, token, new_password, confirm_password = self.username_var.get(), self.token_var.get(), self.new_password_var.get(), self.confirm_password_var.get()
        if not all([username, token, new_password, confirm_password]):
            self.status_label.config(text="Todos os campos são obrigatórios.")
            return
        if new_password != confirm_password:
            self.status_label.config(text="As senhas não coincidem.")
            return

        self.set_pass_button.config(state="disabled")
        self.controller.config(cursor="watch")
        self.update_idletasks()

        try:
            result = auth_logic.set_password_with_token(username, token, new_password)
            if "sucesso" in result:
                messagebox.showinfo("Sucesso", result)
                self.username_var.set(""), self.token_var.set("")
                self.new_password_var.set(""), self.confirm_password_var.set("")
                self.controller.show_frame("LoginFrame")
            else:
                self.status_label.config(text=result)
        except Exception as e:
            self.status_label.config(text="Erro de sistema. Verifique o console para detalhes.")
            print(f"ERRO CRÍTICO em set_password: {e}")
        finally:
            self.set_pass_button.config(state="normal")
            self.controller.config(cursor="")

# --- TELA DE REDEFINIR SENHA ---
class ResetPasswordFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(0, weight=1)
        
        self.email_var = tk.StringVar()
        self.token_var = tk.StringVar()
        self.new_password_var = tk.StringVar()
        self.confirm_password_var = tk.StringVar()
        self.show_passwords_var = tk.BooleanVar(value=False)
        
        self.create_widgets()

    def create_widgets(self):
        main_container = ttk.Frame(self, padding=40)
        main_container.grid(row=0, column=0)
        title = ttk.Label(main_container, text="Redefinir Senha", font=("-size 20 -weight bold"))
        title.pack(pady=(0, 30))
        form_frame = ttk.Frame(main_container)
        form_frame.pack(fill="x")

        ttk.Label(form_frame, text="E-mail de Cadastro:").pack(fill="x")
        self.email_entry = ttk.Entry(form_frame, textvariable=self.email_var, width=40)
        self.email_entry.pack(fill="x", pady=(5, 15))

        ttk.Label(form_frame, text="Código de Verificação (recebido por e-mail):").pack(fill="x")
        self.token_entry = ttk.Entry(form_frame, textvariable=self.token_var, width=40)
        self.token_entry.pack(fill="x", pady=(5, 15))

        ttk.Label(form_frame, text="Nova Senha:").pack(fill="x")
        self.new_pass_entry = ttk.Entry(form_frame, textvariable=self.new_password_var, show="*", width=40, state="disabled")
        self.new_pass_entry.pack(fill="x", pady=(5, 15))

        ttk.Label(form_frame, text="Confirmar Nova Senha:").pack(fill="x")
        self.confirm_pass_entry = ttk.Entry(form_frame, textvariable=self.confirm_password_var, show="*", width=40, state="disabled")
        self.confirm_pass_entry.pack(fill="x", pady=5)
        
        show_pass_check = ttk.Checkbutton(form_frame, text="Mostrar Senhas", variable=self.show_passwords_var, command=self._toggle_passwords_visibility)
        show_pass_check.pack(anchor="w", pady=(5,0))

        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=(30, 10))
        
        self.reset_pass_button = ttk.Button(button_frame, text="Redefinir Senha", command=self.reset_password, bootstyle="primary", state="disabled")
        self.reset_pass_button.pack(side="right")
        
        back_button = ttk.Button(button_frame, text="Voltar para o Login", command=lambda: self.controller.show_frame("LoginFrame"), bootstyle="secondary")
        back_button.pack(side="left")
        
        self.status_label = ttk.Label(main_container, text="", bootstyle="danger")
        self.status_label.pack(pady=(10, 0))

        self.token_var.trace_add("write", self.on_token_change)

    def on_token_change(self, *args):
        if len(self.token_var.get()) == 6:
            self.new_pass_entry.config(state="normal")
            self.confirm_pass_entry.config(state="normal")
            self.reset_pass_button.config(state="normal")
        else:
            self.new_pass_entry.config(state="disabled")
            self.confirm_pass_entry.config(state="disabled")
            self.reset_pass_button.config(state="disabled")
    
    def prepare_for_email(self, email):
        self.email_var.set(email)
        self.token_var.set("")
        self.new_password_var.set("")
        self.confirm_password_var.set("")
        self.status_label.config(text="")
        self.token_entry.focus_set()

    def _toggle_passwords_visibility(self):
        show_char = "" if self.show_passwords_var.get() else "*"
        self.new_pass_entry.config(show=show_char)
        self.confirm_pass_entry.config(show=show_char)

    def reset_password(self):
        self.status_label.config(text="")
        email = self.email_var.get()
        token = self.token_var.get()
        new_password = self.new_password_var.get()
        confirm_password = self.confirm_password_var.get()

        if not all([email, token, new_password, confirm_password]):
            self.status_label.config(text="Todos os campos são obrigatórios.")
            return
        if new_password != confirm_password:
            self.status_label.config(text="As senhas não coincidem.")
            return

        self.reset_pass_button.config(state="disabled")
        self.controller.config(cursor="watch")
        self.update_idletasks()
        
        threading.Thread(target=self._reset_password_thread, args=(email, token, new_password), daemon=True).start()

    def _reset_password_thread(self, email, token, new_password):
        result = auth_logic.reset_password_with_token(email, token, new_password)
        self.after(0, self.on_reset_complete, result)

    def on_reset_complete(self, result):
        self.controller.config(cursor="")
        self.reset_pass_button.config(state="normal")
        
        if "sucesso" in result:
            messagebox.showinfo("Sucesso", result, parent=self.controller)
            self.controller.show_frame("LoginFrame")
        else:
            self.status_label.config(text=result)

class WhatsNewFrame(ttk.Frame):
    def __init__(self, parent, controller, current_version, release_notes):
        super().__init__(parent)
        self.controller = controller
        self.columnconfigure(0, weight=1); self.rowconfigure(1, weight=1)
        ttk.Label(self, text=f"Novidades da Versão {current_version}", font=("Helvetica", 22, "bold")).grid(row=0, column=0, pady=20)
        text_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=10, width=50, font=("Helvetica", 10))
        text_area.grid(row=1, column=0, sticky="nsew", padx=50, pady=10)
        text_area.insert(tk.INSERT, release_notes)
        text_area.config(state='disabled')
        continue_button = ttk.Button(self, text="Entendido, continuar para o programa", command=self.close_and_continue, bootstyle="primary")
        continue_button.grid(row=2, column=0, pady=20)

    def close_and_continue(self):
        self.controller.save_config(); self.controller.show_main_app()