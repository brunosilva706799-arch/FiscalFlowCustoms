# ============================================================================
# --- ARQUIVO: ui/dialogs_flow.py ---
# (Interface para o Gerenciador de Códigos de Cliente e Relatório)
# ============================================================================

import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttk
import os
import subprocess
from collections import defaultdict
from ttkbootstrap.scrolled import ScrolledFrame

class UpdateDownloadWindow(ttk.Toplevel):
    def __init__(self, controller, version_to_download):
        super().__init__(title="Atualização em Andamento", master=controller)
        self.controller = controller
        self.geometry("500x200")
        self.transient(controller); self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both"); main_frame.columnconfigure(0, weight=1)
        self.title_var = tk.StringVar(value=f"Baixando Versão {version_to_download}...")
        ttk.Label(main_frame, textvariable=self.title_var, font=("Helvetica", 14, "bold")).grid(row=0, column=0, pady=(0, 10))
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate', length=300)
        self.progress_bar.grid(row=1, column=0, pady=5, sticky='ew')
        self.status_var = tk.StringVar(value="Conectando ao servidor...")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=2, column=0, pady=5)
        self.download_finished = False; self.download_path = ""

    def update_progress(self, downloaded_bytes, total_bytes):
        percentage = (downloaded_bytes / total_bytes) * 100 if total_bytes > 0 else 0
        self.progress_bar['value'] = percentage
        downloaded_mb, total_mb = downloaded_bytes / (1024 * 1024), total_bytes / (1024 * 1024)
        self.status_var.set(f"Baixado: {downloaded_mb:.2f} MB de {total_mb:.2f} MB ({percentage:.0f}%)")
        
    def on_download_complete(self, final_path):
        self.download_finished = True; self.download_path = final_path
        self.title_var.set("Download Concluído!"); self.status_var.set(f"Arquivo salvo em: {final_path}")
        ttk.Button(self.winfo_children()[0], text="Instalar e Reiniciar", command=self.install_and_close, bootstyle="success").grid(row=3, column=0, pady=20)

    def on_download_error(self, error_message):
        self.destroy(); messagebox.showerror("Erro no Download", f"Não foi possível baixar a atualização.\n\nErro: {error_message}", parent=self.controller)

    def on_closing(self):
        if not self.download_finished:
            if messagebox.askyesno("Cancelar?", "Deseja realmente cancelar o download da atualização?", parent=self): self.destroy()
        else: self.destroy()

    def install_and_close(self):
        try:
            subprocess.Popen([self.download_path]); self.controller.destroy()
        except Exception as e:
            messagebox.showerror("Erro ao Instalar", f"Não foi possível iniciar o instalador.\n\nErro: {e}", parent=self.controller)

class DashboardWindow(ttk.Toplevel):
    def __init__(self, controller, dashboard_data):
        super().__init__(title="Resultado da Extração", master=controller)
        self.controller = controller
        self.dashboard_data = dashboard_data
        self.entry_widgets = {}
        self.minsize(700, 500); self.grab_set()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        resumo_frame = ttk.Labelframe(main_frame, text=" Resumo Geral da Extração ", padding=15)
        resumo_frame.pack(fill="x", pady=(0, 10), side="top")
        resumo_data = self.dashboard_data.get('resumo_geral', {})
        
        ttk.Label(resumo_frame, text="Total de Arquivos Processados:").grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=resumo_data.get('total_notas', 0), font="-weight bold").grid(row=0, column=2, sticky="w")
        
        ttk.Label(resumo_frame, text="Notas de Entrada:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=resumo_data.get('notas_entrada', 0), font="-weight bold").grid(row=1, column=1, sticky="w")
        ttk.Label(resumo_frame, text="Notas de Saída:").grid(row=1, column=2, sticky="w", padx=20, pady=2)
        ttk.Label(resumo_frame, text=resumo_data.get('notas_saida', 0), font="-weight bold").grid(row=1, column=3, sticky="w")
        
        ttk.Label(resumo_frame, text="Notas Canceladas:").grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=resumo_data.get('notas_canceladas', 0), font="-weight bold", bootstyle="warning").grid(row=2, column=2, sticky="w")
        
        ttk.Label(resumo_frame, text="Valor Total (Entrada):").grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        valor_total_entrada = f"R$ {resumo_data.get('valor_total_entrada', 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ttk.Label(resumo_frame, text=valor_total_entrada, font="-weight bold", bootstyle="success").grid(row=3, column=2, sticky="w", columnspan=2)
        
        ttk.Label(resumo_frame, text="Valor Total (Saída):").grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        valor_total_saida = f"R$ {resumo_data.get('valor_total_saida', 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ttk.Label(resumo_frame, text=valor_total_saida, font="-weight bold", bootstyle="info").grid(row=4, column=2, sticky="w", columnspan=2)
        
        ttk.Separator(resumo_frame, orient='horizontal').grid(row=5, column=0, columnspan=4, pady=10, sticky='ew')
        
        def format_currency(value): return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        ttk.Label(resumo_frame, text="Imposto", font="-weight bold").grid(row=6, column=0, sticky="w", padx=5)
        ttk.Label(resumo_frame, text="Valor (Entrada)", font="-weight bold").grid(row=6, column=2, sticky="w", padx=20)
        ttk.Label(resumo_frame, text="Valor (Saída)", font="-weight bold").grid(row=6, column=3, sticky="w", padx=5)
        
        impostos_data = resumo_data.get('impostos', {})
        impostos_entrada = impostos_data.get('Entrada', defaultdict(float))
        impostos_saida = impostos_data.get('Saída', defaultdict(float))
        
        ttk.Label(resumo_frame, text="Total ICMS:").grid(row=7, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=format_currency(impostos_entrada['vICMS']), font="-weight bold").grid(row=7, column=2, sticky="w", padx=20)
        ttk.Label(resumo_frame, text=format_currency(impostos_saida['vICMS']), font="-weight bold").grid(row=7, column=3, sticky="w", padx=5)
        
        ttk.Label(resumo_frame, text="Total IPI:").grid(row=8, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=format_currency(impostos_entrada['vIPI']), font="-weight bold").grid(row=8, column=2, sticky="w", padx=20)
        ttk.Label(resumo_frame, text=format_currency(impostos_saida['vIPI']), font="-weight bold").grid(row=8, column=3, sticky="w", padx=5)
        
        ttk.Label(resumo_frame, text="Total PIS:").grid(row=9, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=format_currency(impostos_entrada['vPIS']), font="-weight bold").grid(row=9, column=2, sticky="w", padx=20)
        ttk.Label(resumo_frame, text=format_currency(impostos_saida['vPIS']), font="-weight bold").grid(row=9, column=3, sticky="w", padx=5)
        
        ttk.Label(resumo_frame, text="Total COFINS:").grid(row=10, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Label(resumo_frame, text=format_currency(impostos_entrada['vCOFINS']), font="-weight bold").grid(row=10, column=2, sticky="w", padx=20)
        ttk.Label(resumo_frame, text=format_currency(impostos_saida['vCOFINS']), font="-weight bold").grid(row=10, column=3, sticky="w", padx=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))
        button_frame.columnconfigure(0, weight=1); button_frame.columnconfigure(1, weight=0); button_frame.columnconfigure(2, weight=1)
        ttk.Button(button_frame, text="Salvar Planilha Básica", command=self.save_basic).grid(row=0, column=0, sticky="w", padx=5)
        ttk.Button(button_frame, text="Ver Pré-visualização Detalhada", command=self.open_preview, bootstyle="info-outline").grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Salvar Planilha Completa", command=self.confirm_and_save, bootstyle="primary").grid(row=0, column=2, sticky="e", padx=5)
        
        scrollable_area = ScrolledFrame(main_frame, autohide=True)
        scrollable_area.pack(fill="both", expand=True, side="top")
        processos_pendentes = self.dashboard_data.get('processos_pendentes', {})
        if processos_pendentes:
            acoes_frame = ttk.Labelframe(scrollable_area, text=" Ações Pendentes ", padding=15)
            acoes_frame.pack(fill="x", expand=True, padx=5, pady=5)
            instrucao_texto = "Para processos com múltiplas NFs de Saída, informe o total esperado de notas para validar a emissão das notas de serviço."
            instrucao_label = ttk.Label(acoes_frame, text=instrucao_texto, wraplength=550, justify="left", bootstyle="secondary")
            instrucao_label.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))
            ttk.Label(acoes_frame, text="Processo", font="-weight bold").grid(row=1, column=0, padx=5, pady=5)
            ttk.Label(acoes_frame, text="Notas Encontradas", font="-weight bold").grid(row=1, column=1, padx=5, pady=5)
            ttk.Label(acoes_frame, text="Total Esperado (informar)", font="-weight bold").grid(row=1, column=2, padx=5, pady=5)
            for i, (processo, contagem) in enumerate(processos_pendentes.items(), start=2):
                ttk.Label(acoes_frame, text=processo).grid(row=i, column=0, padx=5, pady=3, sticky="w")
                ttk.Label(acoes_frame, text=str(contagem)).grid(row=i, column=1, padx=5, pady=3)
                entry = ttk.Entry(acoes_frame, width=10)
                entry.grid(row=i, column=2, padx=5, pady=3)
                self.entry_widgets[processo] = (entry, contagem)

    def _collect_and_validate_counts(self):
        contagens_confirmadas = {}
        for processo, (entry, contagem_encontrada) in self.entry_widgets.items():
            valor_str = entry.get().strip()
            if not valor_str: messagebox.showerror("Erro", f"O campo 'Total Esperado' para o processo '{processo}' não pode estar vazio.", parent=self); return None
            try:
                valor_int = int(valor_str)
                if valor_int < contagem_encontrada: messagebox.showerror("Erro", f"O valor esperado ({valor_int}) para '{processo}' não pode ser menor que o encontrado ({contagem_encontrada}).", parent=self); return None
                contagens_confirmadas[processo] = {'encontrado': contagem_encontrada, 'esperado': valor_int}
            except ValueError: messagebox.showerror("Erro", f"Insira um número inteiro válido para o processo '{processo}'.", parent=self); return None
        return contagens_confirmadas

    def save_basic(self):
        nfe_tool_frame = self.controller.frames['NFeToolFrame']; nfe_tool_frame.salvar_dados_basicos(); self.destroy()
    def confirm_and_save(self):
        contagens = self._collect_and_validate_counts()
        if contagens is None: return
        nfe_tool_frame = self.controller.frames['NFeToolFrame']; nfe_tool_frame.salvar_planilha_completa(contagens); self.destroy()
    def open_preview(self):
        contagens = self._collect_and_validate_counts()
        if contagens is None: return
        nfe_tool_frame = self.controller.frames['NFeToolFrame']; nfe_tool_frame.show_preview_window(contagens)
        
class PreviewWindow(ttk.Toplevel):
    def __init__(self, controller, all_extracted_data, contagens_confirmadas):
        super().__init__(title="Pré-visualização da Planilha", master=controller)
        self.geometry("1200x700"); self.grab_set()
        
        self.all_extracted_data = all_extracted_data
        self.contagens_confirmadas = contagens_confirmadas
        self.active_tree = None
        
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(expand=True, fill="both")
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(expand=True, fill="both")
        
        self.build_tabs()

        close_button = ttk.Button(main_frame, text="Fechar Pré-visualização", command=self.destroy, bootstyle="secondary")
        close_button.pack(side="bottom", pady=(10,0))

    def build_tabs(self):
        for i in self.notebook.tabs():
            self.notebook.forget(i)

        dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend = self._process_data()
        cabecalhos = self._get_headers()

        self.create_tab("Entrada", cabecalhos['completo'], dados_entrada, add_context_menu=True)
        self.create_tab("Saída", cabecalhos['completo'], dados_saida, add_context_menu=True)
        self.create_tab("Serviços Autorizados", cabecalhos['servico'], dados_servico_auth)
        self.create_tab("Serviços Pendentes", cabecalhos['pendentes'], dados_servico_pend)

    def _process_data(self):
        dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend = [], [], [], []
        processos_saida = defaultdict(list)
        for dados in self.all_extracted_data:
            if dados['tipo_nota'] == 'Saída':
                processos_saida[dados.get('processo_normalizado', 'N/A')].append(dados)
        
        for dados in self.all_extracted_data:
            if dados['tipo_nota'] == "Saída": dados_saida.append(dados)
            else: dados_entrada.append(dados)

            if dados.get('valor_servico_trading', 0.0) > 0.0 and dados['tipo_nota'] == 'Entrada':
                processo = dados.get('processo_normalizado', 'N/A')
                nfs_encontradas = len(processos_saida.get(processo, []))
                contagem_info = self.contagens_confirmadas.get(processo)
                esperada = contagem_info.get('esperado') if contagem_info else None
                
                auth = (esperada is not None and nfs_encontradas == esperada) or (nfs_encontradas > 0 and processo not in self.contagens_confirmadas)
                
                if auth: dados_servico_auth.append(dados)
                else:
                    dados['qtde_encontrada'] = nfs_encontradas
                    dados['qtde_esperada'] = esperada
                    dados['qtde_a_emitir'] = esperada - nfs_encontradas if esperada is not None else 'N/A'
                    dados_servico_pend.append(dados)
        return dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend

    def _get_headers(self):
        # --- [MODIFICADO] Lista de cabeçalhos agora está completa ---
        cabecalho_completo = [
            'Arquivo', 'Número da NF', 'Status da NF', 'Data de Emissão', 'Cliente', 
            'CNPJ/CPF Dest.', 'Nome do Processo', 'CFOP', 'Valor Total dos Produtos', 'Valor II', 
            'Valor ICMS', 'Valor IPI', 'Valor PIS', 'Valor COFINS', 'Outras Despesas', 
            'Valor AFRMM', 'Frete Nacional', 'Valor Total'
        ]
        return {
            'completo': cabecalho_completo,
            'servico': ['Número da NF', 'Cliente', 'Data de Emissão', 'Nome do Processo', 'Valor Serviço Trading'],
            'pendentes': ['Número da NF', 'Cliente', 'Data de Emissão', 'Nome do Processo', 'Valor Serviço Trading', 'Qtde Encontrada', 'Qtde Esperada', 'Qtde a Emitir']
        }

    def create_tab(self, sheet_name, headers, data_rows, add_context_menu=False):
        if not data_rows: return
        tab = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(tab, text=f"{sheet_name} ({len(data_rows)})")
        tree_frame = ttk.Frame(tab); tree_frame.pack(expand=True, fill="both")
        tree = ttk.Treeview(tree_frame, columns=headers, show='headings')
        tree.pack(side="left", expand=True, fill="both")
        
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        ysb.pack(side="right", fill="y"); tree.configure(yscrollcommand=ysb.set)
        xsb = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview)
        xsb.pack(side="bottom", fill="x"); tree.configure(xscrollcommand=xsb.set)

        for header in headers:
            # --- [MODIFICADO] Ancora valores numéricos à direita para melhor visualização ---
            anchor = 'e' if 'Valor' in header or 'Frete' in header else 'w'
            tree.heading(header, text=header, command=lambda h=header, t=tree: self.sort_column(t, h, False))
            tree.column(header, width=150, anchor=anchor, stretch=False)
        
        tree.tag_configure('cancelled_row', foreground='gray', font='-slant italic')
        self.populate_tree(tree, headers, data_rows)

        if add_context_menu:
            self.context_menu = tk.Menu(self, tearoff=0)
            self.context_menu.add_command(label="Marcar como Cancelada", command=self.mark_as_cancelled)
            tree.bind("<Button-3>", lambda e, t=tree: self.show_context_menu(e, t))

    def populate_tree(self, tree, headers, data_rows):
        tree.delete(*tree.get_children())
        for data_dict in data_rows:
            values = [self._format_value(data_dict, header) for header in headers]
            iid = data_dict.get('nome_arquivo')
            tags = ('cancelled_row',) if 'Cancelada' in str(data_dict.get('status_nf', '')) else ()
            tree.insert('', 'end', values=values, iid=iid, tags=tags)
    
    def _format_value(self, data_dict, header):
        key = self._get_key_from_header(header)
        value = data_dict.get(key)
        
        if header in ['Valor ICMS', 'Valor IPI', 'Valor PIS', 'Valor COFINS']:
            value = float(data_dict.get('impostos', {}).get(key, '0.00'))

        if isinstance(value, (int, float)):
            return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return value if value is not None else ''
        
    def _get_key_from_header(self, header):
        # --- [MODIFICADO] Mapeamento atualizado com todas as chaves ---
        mapping = {
            'Arquivo': 'nome_arquivo', 'Número da NF': 'numero_nf', 'Status da NF': 'status_nf',
            'Data de Emissão': 'data_emissao', 'Cliente': 'nome_cliente', 'CNPJ/CPF Dest.': 'cnpj_cpf_destinatario',
            'Nome do Processo': 'nome_processo', 'CFOP': 'cfop_nf', 'Valor Total dos Produtos': 'valor_total_produtos',
            'Valor II': 'vII', 'Valor ICMS': 'vICMS', 'Valor IPI': 'vIPI', 'Valor PIS': 'vPIS',
            'Valor COFINS': 'vCOFINS', 'Outras Despesas': 'vOutras', 'Valor AFRMM': 'vAFRMM',
            'Frete Nacional': 'valor_frete_nacional', 'Valor Total': 'valor_total_nf', 
            'Valor Serviço Trading': 'valor_servico_trading', 'Qtde Encontrada': 'qtde_encontrada',
            'Qtde Esperada': 'qtde_esperada', 'Qtde a Emitir': 'qtde_a_emitir'
        }
        return mapping.get(header)

    def show_context_menu(self, event, tree):
        self.active_tree = tree
        selection_id = tree.identify_row(event.y)
        if not selection_id: return
        
        tree.selection_set(selection_id)
        item_data = next((d for d in self.all_extracted_data if d.get('nome_arquivo') == selection_id), None)
        
        if item_data and item_data.get('status_nf') == 'Autorizada':
            self.context_menu.entryconfig("Marcar como Cancelada", state="normal")
        else:
            self.context_menu.entryconfig("Marcar como Cancelada", state="disabled")
            
        self.context_menu.post(event.x_root, event.y_root)

    def mark_as_cancelled(self):
        if not self.active_tree: return
        selection_id = self.active_tree.focus()
        if not selection_id: return
        
        for data_dict in self.all_extracted_data:
            if data_dict.get('nome_arquivo') == selection_id:
                data_dict['status_nf'] = 'Cancelada (Manual)'
                data_dict['valor_total_nf'] = 0.0
                data_dict['valor_total_produtos'] = 0.0
                data_dict['valor_frete_nacional'] = 0.0
                data_dict['vII'] = 0.0
                data_dict['vAFRMM'] = 0.0
                data_dict['vOutras'] = 0.0
                data_dict['impostos'] = defaultdict(lambda: '0.00')
                break
        
        self.build_tabs()

    def sort_column(self, tree, col, reverse):
        key = self._get_key_from_header(col)
        # --- [MODIFICADO] Lógica de ordenação aprimorada para funcionar com dados brutos ---
        is_numeric = 'Valor' in col or 'Qtde' in col or 'Frete' in col or 'Número' in col

        try:
            data_list = []
            for iid in tree.get_children(''):
                item_data = next((d for d in self.all_extracted_data if d.get('nome_arquivo') == iid), None)
                if not item_data: continue

                val = item_data.get(key, 0.0 if is_numeric else '')
                if 'imposto' in col.lower(): val = float(item_data.get('impostos', {}).get(key, '0.00'))
                
                if is_numeric and not isinstance(val, (int, float)):
                    val = 0.0
                
                data_list.append((val, iid))

            data_list.sort(key=lambda t: t[0], reverse=reverse)
            for index, (val, iid) in enumerate(data_list): tree.move(iid, '', index)
            tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))
        except Exception as e:
            print(f"Erro ao ordenar coluna: {e}")