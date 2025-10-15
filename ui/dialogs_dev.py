# =============================================================================
# --- ARQUIVO: ui/dialogs_dev.py ---
# (Contém as janelas de diálogo para o desenvolvedor)
# =============================================================================

import tkinter as tk
from tkinter import ttk, scrolledtext
import ttkbootstrap as ttk
import threading
import test_logic # Importa nossa nova lógica de teste

class TestRunnerWindow(ttk.Toplevel):
    def __init__(self, controller):
        super().__init__(title="Auto-teste do Sistema", master=controller)
        self.controller = controller
        self.minsize(600, 400)
        self.grab_set()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Área de texto para os resultados
        self.results_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, state="disabled")
        self.results_text.grid(row=0, column=0, sticky="nsew", columnspan=2)

        # Configuração das tags de cor
        self.results_text.tag_config("OK", foreground="green")
        self.results_text.tag_config("FALHOU", foreground="red")

        # Botões
        self.run_button = ttk.Button(main_frame, text="Iniciar Testes", command=self.run_tests, bootstyle="primary")
        self.run_button.grid(row=1, column=0, pady=(10,0), sticky="w")
        
        close_button = ttk.Button(main_frame, text="Fechar", command=self.destroy, bootstyle="secondary")
        close_button.grid(row=1, column=1, pady=(10,0), sticky="e")

    def run_tests(self):
        """Inicia a suíte de testes em uma thread separada."""
        self.run_button.config(state="disabled", text="Testando...")
        
        # Limpa os resultados anteriores
        self.results_text.config(state="normal")
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert('1.0', "[INICIANDO AUTO-TESTE DO SISTEMA...]\n\n")
        self.results_text.config(state="disabled")

        # Roda os testes em segundo plano para não congelar a interface
        threading.Thread(target=self._run_tests_in_thread, daemon=True).start()

    def _run_tests_in_thread(self):
        """Função que a thread executa."""
        results = test_logic.run_all_tests()
        # Envia os resultados de volta para a thread principal da GUI
        self.after(0, self.display_results, results)

    def display_results(self, results):
        """Exibe os resultados formatados na área de texto."""
        self.results_text.config(state="normal")
        
        failures = 0
        for test_name, (status, message) in results.items():
            self.results_text.insert(tk.END, f"- Testando {test_name}... ")
            self.results_text.insert(tk.END, f"[{status}]\n", status) # Usa a tag de cor
            if status == "FALHOU":
                failures += 1
                self.results_text.insert(tk.END, f"  Detalhe: {message}\n\n")
        
        self.results_text.insert(tk.END, "\n[TESTE CONCLUÍDO]\n")
        if failures > 0:
            summary = f"Resultado: {failures} teste(s) falharam."
            self.results_text.insert(tk.END, summary, "FALHOU")
        else:
            summary = "Resultado: Todos os testes passaram com sucesso!"
            self.results_text.insert(tk.END, summary, "OK")

        self.results_text.config(state="disabled")
        self.run_button.config(state="normal", text="Iniciar Testes Novamente")