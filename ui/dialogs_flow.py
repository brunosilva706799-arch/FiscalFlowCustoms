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
from ttkbootstrap.tableview import Tableview
import core_logic

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
        self.resumo_labels = {} 
        self.minsize(700, 500); self.grab_set()
        self.create_widgets()
        self.update_dashboard_display()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")
        
        self.resumo_frame = ttk.Labelframe(main_frame, text=" Resumo Geral da Extração ", padding=15)
        self.resumo_frame.pack(fill="x", pady=(0, 10), side="top")
        
        labels_info = {
            'total_notas': {"text": "Total de Arquivos Processados:", "row": 0, "col": 0, "cs": 2},
            'notas_entrada': {"text": "Notas de Entrada:", "row": 1, "col": 0},
            'notas_saida': {"text": "Notas de Saída:", "row": 1, "col": 2, "padx": 20},
            'notas_canceladas': {"text": "Notas Canceladas:", "row": 2, "col": 0, "cs": 2, "bootstyle": "warning"},
            'valor_total_entrada': {"text": "Valor Total (Entrada):", "row": 3, "col": 0, "cs": 2, "bootstyle": "success"},
            'valor_total_saida': {"text": "Valor Total (Saída):", "row": 4, "col": 0, "cs": 2, "bootstyle": "info"},
        }
        for key, info in labels_info.items():
            ttk.Label(self.resumo_frame, text=info["text"]).grid(row=info["row"], column=info.get("col", 0), columnspan=info.get("cs", 1), sticky="w", padx=info.get("padx", 5), pady=2)
            var = tk.StringVar()
            label = ttk.Label(self.resumo_frame, textvariable=var, font="-weight bold", bootstyle=info.get("bootstyle", "default"))
            label.grid(row=info["row"], column=info.get("col", 0) + info.get("cs", 1), sticky="w", padx=5, columnspan=2)
            self.resumo_labels[key] = var

        ttk.Separator(self.resumo_frame, orient='horizontal').grid(row=5, column=0, columnspan=4, pady=10, sticky='ew')
        
        ttk.Label(self.resumo_frame, text="Imposto", font="-weight bold").grid(row=6, column=0, sticky="w", padx=5)
        ttk.Label(self.resumo_frame, text="Valor (Entrada)", font="-weight bold").grid(row=6, column=2, sticky="w", padx=20)
        ttk.Label(self.resumo_frame, text="Valor (Saída)", font="-weight bold").grid(row=6, column=3, sticky="w", padx=5)

        impostos_info = {
            'vICMS': {"text": "Total ICMS:", "row": 7},
            'vIPI': {"text": "Total IPI:", "row": 8},
            'vPIS': {"text": "Total PIS:", "row": 9},
            'vCOFINS': {"text": "Total COFINS:", "row": 10},
        }
        self.resumo_labels['impostos'] = {'Entrada': {}, 'Saída': {}}
        for key, info in impostos_info.items():
            ttk.Label(self.resumo_frame, text=info["text"]).grid(row=info["row"], column=0, columnspan=2, sticky="w", padx=5, pady=2)
            var_ent = tk.StringVar()
            ttk.Label(self.resumo_frame, textvariable=var_ent, font="-weight bold").grid(row=info["row"], column=2, sticky="w", padx=20)
            self.resumo_labels['impostos']['Entrada'][key] = var_ent
            var_sai = tk.StringVar()
            ttk.Label(self.resumo_frame, textvariable=var_sai, font="-weight bold").grid(row=info["row"], column=3, sticky="w", padx=5)
            self.resumo_labels['impostos']['Saída'][key] = var_sai

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))
        button_frame.columnconfigure([0, 1, 2, 3], weight=1)
        
        ttk.Button(button_frame, text="Salvar Planilha Básica", command=self.save_basic).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(button_frame, text="Editar / Visualizar Dados", command=self.open_preview, bootstyle="info-outline").grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(button_frame, text="Recalcular", command=self.recalculate_dashboard, bootstyle="secondary").grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Button(button_frame, text="Salvar Planilha Completa", command=self.confirm_and_save, bootstyle="primary").grid(row=0, column=3, sticky="ew", padx=2)
        
        scrollable_area = ScrolledFrame(main_frame, autohide=True)
        scrollable_area.pack(fill="both", expand=True, side="top")
        processos_pendentes = self.dashboard_data.get('processos_pendentes', {})
        if processos_pendentes:
            acoes_frame = ttk.Labelframe(scrollable_area, text=" Ações Pendentes ", padding=15)
            acoes_frame.pack(fill="x", expand=True, padx=5, pady=5)
            instrucao_label = ttk.Label(acoes_frame, text="Para processos com múltiplas NFs de Saída, informe o total esperado de notas.", wraplength=550, justify="left", bootstyle="secondary")
            instrucao_label.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))
            ttk.Label(acoes_frame, text="Processo", font="-weight bold").grid(row=1, column=0, padx=5, pady=5)
            ttk.Label(acoes_frame, text="Notas Encontradas", font="-weight bold").grid(row=1, column=1, padx=5, pady=5)
            ttk.Label(acoes_frame, text="Total Esperado (informar)", font="-weight bold").grid(row=1, column=2, padx=5, pady=5)
            for i, (processo, contagem) in enumerate(processos_pendentes.items(), start=2):
                ttk.Label(acoes_frame, text=processo).grid(row=i, column=0, padx=5, pady=3, sticky="w")
                ttk.Label(acoes_frame, text=str(contagem)).grid(row=i, column=1, padx=5, pady=3)
                entry = ttk.Entry(acoes_frame, width=10); entry.grid(row=i, column=2, padx=5, pady=3)
                self.entry_widgets[processo] = (entry, contagem)

    def update_dashboard_display(self):
        resumo_data = self.dashboard_data.get('resumo_geral', {})
        def format_currency(value): return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        self.resumo_labels['total_notas'].set(resumo_data.get('total_notas', 0))
        self.resumo_labels['notas_entrada'].set(resumo_data.get('notas_entrada', 0))
        self.resumo_labels['notas_saida'].set(resumo_data.get('notas_saida', 0))
        self.resumo_labels['notas_canceladas'].set(resumo_data.get('notas_canceladas', 0))
        self.resumo_labels['valor_total_entrada'].set(format_currency(resumo_data.get('valor_total_entrada', 0.0)))
        self.resumo_labels['valor_total_saida'].set(format_currency(resumo_data.get('valor_total_saida', 0.0)))
        
        impostos_data = resumo_data.get('impostos', {})
        impostos_entrada = impostos_data.get('Entrada', defaultdict(float))
        impostos_saida = impostos_data.get('Saída', defaultdict(float))

        for imp_key in ['vICMS', 'vIPI', 'vPIS', 'vCOFINS']:
            self.resumo_labels['impostos']['Entrada'][imp_key].set(format_currency(impostos_entrada[imp_key]))
            self.resumo_labels['impostos']['Saída'][imp_key].set(format_currency(impostos_saida[imp_key]))

    def recalculate_dashboard(self):
        nfe_tool_frame = self.controller.frames['NFeToolFrame']
        self.dashboard_data = core_logic.calcular_dados_dashboard(nfe_tool_frame.dados_extraidos_em_memoria)
        self.update_dashboard_display()
        messagebox.showinfo("Recálculo", "Os totais do dashboard foram atualizados com sucesso.", parent=self)

    def _collect_and_validate_counts(self):
        contagens_confirmadas = {}
        for processo, (entry, contagem_encontrada) in self.entry_widgets.items():
            valor_str = entry.get().strip()
            if not valor_str: messagebox.showerror("Erro", f"O campo 'Total Esperado' para '{processo}' não pode estar vazio.", parent=self); return None
            try:
                valor_int = int(valor_str)
                if valor_int < contagem_encontrada: messagebox.showerror("Erro", f"O valor esperado ({valor_int}) para '{processo}' não pode ser menor que o encontrado ({contagem_encontrada}).", parent=self); return None
                contagens_confirmadas[processo] = {'encontrado': contagem_encontrada, 'esperado': valor_int}
            except ValueError: messagebox.showerror("Erro", f"Insira um número inteiro válido para '{processo}'.", parent=self); return None
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
        nfe_tool_frame = self.controller.frames['NFeToolFrame']
        PreviewWindow(self.controller, self, nfe_tool_frame.dados_extraidos_em_memoria, contagens)
        
class StatusActionDialog(ttk.Toplevel):
    def __init__(self, parent, selected_count):
        super().__init__(title="Alterar Status", master=parent)
        self.result = None
        self.transient(parent); self.grab_set()
        main_frame = ttk.Frame(self, padding=20); main_frame.pack(expand=True, fill="both")
        ttk.Label(main_frame, text=f"Qual ação aplicar às {selected_count} notas selecionadas?", wraplength=300).pack(pady=(0, 15))
        self.action_var = tk.StringVar(value="cancel")
        ttk.Radiobutton(main_frame, text="Marcar como Cancelada (Manual)", variable=self.action_var, value="cancel").pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(main_frame, text="Reverter para Status Original", variable=self.action_var, value="revert").pack(anchor="w", padx=10, pady=2)
        btn_frame = ttk.Frame(main_frame, padding=(0, 20, 0, 0)); btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Confirmar", command=self.on_confirm, bootstyle="primary").pack(side="right")
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancel).pack(side="right", padx=10)
    def on_confirm(self):
        self.result = self.action_var.get(); self.destroy()
    def on_cancel(self):
        self.result = None; self.destroy()

class PreviewWindow(ttk.Toplevel):
    def __init__(self, controller, parent_dashboard, all_extracted_data, contagens_confirmadas):
        super().__init__(title="Editar e Visualizar Dados", master=parent_dashboard)
        self.controller = controller
        self.parent_dashboard = parent_dashboard
        self.all_extracted_data = all_extracted_data
        self.contagens_confirmadas = contagens_confirmadas
        self.treeviews = {}
        self.checked_items = {}
        self.geometry("1200x700"); self.grab_set()
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=10); main_frame.pack(expand=True, fill="both")
        toolbar = ttk.Frame(main_frame); toolbar.pack(fill="x", pady=(0, 5))
        self.action_btn = ttk.Button(toolbar, text="Alterar Status dos Selecionados", command=self.open_status_action_dialog, state="disabled")
        self.action_btn.pack(side="left")
        self.notebook = ttk.Notebook(main_frame); self.notebook.pack(expand=True, fill="both")
        self.build_tabs()
        close_button = ttk.Button(main_frame, text="Salvar Alterações e Voltar", command=self.save_and_close, bootstyle="primary")
        close_button.pack(side="bottom", pady=(10,0))
    
    def save_and_close(self):
        self.parent_dashboard.recalculate_dashboard(); self.destroy()

    def build_tabs(self):
        for i in self.notebook.tabs(): self.notebook.forget(i)
        self.treeviews.clear()
        self.checked_items.clear()
        dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend = self._process_data()
        cabecalhos = self._get_headers()
        self.create_tab("Entrada", cabecalhos['completo'], dados_entrada)
        self.create_tab("Saída", cabecalhos['completo'], dados_saida)
        self.create_tab("Serviços Autorizados", cabecalhos['servico'], dados_servico_auth)
        self.create_tab("Serviços Pendentes", cabecalhos['pendentes'], dados_servico_pend)
        self.update_action_button_state()

    def _process_data(self):
        dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend = [], [], [], []
        processos_saida = defaultdict(list)
        for dados in self.all_extracted_data:
            if dados['tipo_nota'] == 'Saída' and 'Cancelada' not in dados.get('status_nf', ''):
                processos_saida[dados.get('processo_normalizado', 'N/A')].append(dados)
        for dados in self.all_extracted_data:
            if dados['tipo_nota'] == "Saída": dados_saida.append(dados)
            elif dados['tipo_nota'] == "Entrada": dados_entrada.append(dados)
            if dados.get('valor_servico_trading', 0.0) > 0.0 and dados['tipo_nota'] == 'Entrada' and 'Cancelada' not in dados.get('status_nf', ''):
                processo = dados.get('processo_normalizado', 'N/A')
                nfs_encontradas = len(processos_saida.get(processo, []))
                contagem_info = self.contagens_confirmadas.get(processo)
                esperada = contagem_info.get('esperado') if contagem_info else None
                auth = False
                if processo in self.contagens_confirmadas:
                    if esperada is not None and nfs_encontradas == esperada: auth = True
                elif nfs_encontradas > 0: auth = True
                
                if auth: dados_servico_auth.append(dados)
                else:
                    dados['qtde_encontrada'] = nfs_encontradas; dados['qtde_esperada'] = esperada
                    dados['qtde_a_emitir'] = esperada - nfs_encontradas if esperada is not None else 'N/A'
                    dados_servico_pend.append(dados)
        return dados_entrada, dados_saida, dados_servico_auth, dados_servico_pend

    def _get_headers(self):
        cabecalho_completo = ['Arquivo', 'Número da NF', 'Status da NF', 'Data de Emissão', 'Cliente', 'CNPJ/CPF Dest.', 'Nome do Processo', 'Valor Serviço Trading', 'CFOP', 'Valor Total dos Produtos', 'Valor II', 'Valor ICMS', 'Valor IPI', 'Valor PIS', 'Valor COFINS', 'Outras Despesas', 'Valor AFRMM', 'Frete Nacional', 'Valor Total']
        return {
            'completo': ['select'] + cabecalho_completo,
            'servico': ['select'] + ['Número da NF', 'Cliente', 'Data de Emissão', 'Nome do Processo', 'Valor Serviço Trading'],
            'pendentes': ['select'] + ['Número da NF', 'Cliente', 'Data de Emissão', 'Nome do Processo', 'Valor Serviço Trading', 'Qtde Encontrada', 'Qtde Esperada', 'Qtde a Emitir']
        }

    def create_tab(self, sheet_name, headers, data_rows):
        if not data_rows: return
        tab = ttk.Frame(self.notebook, padding=5); self.notebook.add(tab, text=f"{sheet_name} ({len(data_rows)})")
        tree_frame = ttk.Frame(tab); tree_frame.pack(expand=True, fill="both")
        tree = ttk.Treeview(tree_frame, columns=headers, show='headings'); tree.pack(side="left", expand=True, fill="both")
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview); ysb.pack(side="right", fill="y"); tree.configure(yscrollcommand=ysb.set)
        xsb = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview); xsb.pack(side="bottom", fill="x"); tree.configure(xscrollcommand=xsb.set)

        for header in headers:
            anchor = 'e' if 'Valor' in header or 'Frete' in header else 'w'
            tree.heading(header, text=header, command=lambda h=header, t=tree: self.sort_column(t, h, False))
            tree.column(header, width=150, anchor=anchor, stretch=False)
        
        tree.column('select', width=40, stretch=False, anchor='center')
        tree.heading('select', text='[ ]')
        
        tree.tag_configure('cancelled_row', foreground='gray', font='-slant italic')
        self.populate_tree(tree, headers, data_rows)
        self.treeviews[sheet_name] = tree
        tree.bind('<Button-1>', self.toggle_check)

    def populate_tree(self, tree, headers, data_rows):
        tree.delete(*tree.get_children())
        if not hasattr(self, 'checked_items'): self.checked_items = {}
        for data_dict in data_rows:
            iid = data_dict.get('nome_arquivo')
            self.checked_items[iid] = self.checked_items.get(iid, False)
            checkbox = '[x]' if self.checked_items[iid] else '[ ]'
            values = [checkbox] + [self._format_value(data_dict, h) for h in headers[1:]]
            tags = ('cancelled_row',) if 'Cancelada' in str(data_dict.get('status_nf', '')) else ()
            tree.insert('', 'end', values=values, iid=iid, tags=tags)
    
    def toggle_check(self, event):
        tree = event.widget; iid = tree.identify_row(event.y)
        if not iid: return
        self.checked_items[iid] = not self.checked_items.get(iid, False)
        current_values = list(tree.item(iid, 'values'))
        current_values[0] = '[x]' if self.checked_items[iid] else '[ ]'
        tree.item(iid, values=current_values)
        self.update_action_button_state()
        
    def _format_value(self, data_dict, header):
        key = self._get_key_from_header(header); value = data_dict.get(key)
        if header in ['Valor ICMS', 'Valor IPI', 'Valor PIS', 'Valor COFINS']: value = float(data_dict.get('impostos', {}).get(key, '0.00'))
        if isinstance(value, (int, float)):
            return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return value if value is not None else ''
        
    def _get_key_from_header(self, header):
        mapping = { 'Arquivo': 'nome_arquivo', 'Número da NF': 'numero_nf', 'Status da NF': 'status_nf', 'Data de Emissão': 'data_emissao', 'Cliente': 'nome_cliente', 'CNPJ/CPF Dest.': 'cnpj_cpf_destinatario', 'Nome do Processo': 'nome_processo', 'CFOP': 'cfop_nf', 'Valor Total dos Produtos': 'valor_total_produtos', 'Valor II': 'vII', 'Valor ICMS': 'vICMS', 'Valor IPI': 'vIPI', 'Valor PIS': 'vPIS', 'Valor COFINS': 'vCOFINS', 'Outras Despesas': 'vOutras', 'Valor AFRMM': 'vAFRMM', 'Frete Nacional': 'valor_frete_nacional', 'Valor Total': 'valor_total_nf', 'Valor Serviço Trading': 'valor_servico_trading', 'Qtde Encontrada': 'qtde_encontrada', 'Qtde Esperada': 'qtde_esperada', 'Qtde a Emitir': 'qtde_a_emitir' }
        return mapping.get(header)

    def update_action_button_state(self, event=None):
        any_checked = any(self.checked_items.values())
        self.action_btn.config(state="normal" if any_checked else "disabled")

    def open_status_action_dialog(self):
        checked_iids = [iid for iid, checked in self.checked_items.items() if checked]
        if not checked_iids: return
        dialog = StatusActionDialog(self, len(checked_iids))
        self.wait_window(dialog)
        if dialog.result:
            self.apply_batch_status_change(checked_iids, dialog.result)
    
    def apply_batch_status_change(self, iids, action):
        for iid in iids:
            for i, data_dict in enumerate(self.all_extracted_data):
                if data_dict.get('nome_arquivo') == iid:
                    if action == 'cancel':
                        self.all_extracted_data[i]['status_nf'] = 'Cancelada (Manual)'
                        self.all_extracted_data[i]['valor_total_nf'] = 0.0
                        self.all_extracted_data[i]['valor_total_produtos'] = 0.0
                        self.all_extracted_data[i]['valor_frete_nacional'] = 0.0
                        self.all_extracted_data[i]['vII'] = 0.0
                        self.all_extracted_data[i]['vAFRMM'] = 0.0
                        self.all_extracted_data[i]['vOutras'] = 0.0
                        self.all_extracted_data[i]['impostos'] = defaultdict(lambda: '0.00')
                    elif action == 'revert':
                        if '_dados_originais' in data_dict:
                            self.all_extracted_data[i] = data_dict['_dados_originais'].copy()
                    break
        self.build_tabs()
    
    def sort_column(self, tree, col, reverse):
        key = self._get_key_from_header(col)
        is_numeric = 'Valor' in col or 'Qtde' in col or 'Frete' in col or 'Número' in col
        try:
            data_list = []
            for iid in tree.get_children(''):
                item_data = next((d for d in self.all_extracted_data if d.get('nome_arquivo') == iid), None)
                if not item_data: continue
                val = item_data.get(key, 0.0 if is_numeric else '')
                if 'imposto' in col.lower(): val = float(item_data.get('impostos', {}).get(key, '0.00'))
                if is_numeric and not isinstance(val, (int, float)): val = 0.0
                data_list.append((val, iid))

            data_list.sort(key=lambda t: t[0], reverse=reverse)
            for index, (val, iid) in enumerate(data_list):
                tree.move(iid, '', index)
            tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))
        except Exception as e:
            print(f"Erro ao ordenar coluna: {e}")