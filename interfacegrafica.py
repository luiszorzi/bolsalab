import tkinter as tk
from tkinter import messagebox, ttk
import pyvisa
import threading
import time
import csv
import datetime
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import os
import matplotlib.animation as animation
import numpy as np
from functools import partial

class JanelaControleCombinado(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Controle de Sequência de Testes")
        self.geometry("1200x850")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.font_titulo = ("Segoe UI", 12, "bold")
        self.font_corpo = ("Segoe UI", 11)
        self.font_summary_title = ("Segoe UI", 10, "bold")
        self.font_summary_subtitle = ("Segoe UI", 9, "bold")
        self.font_summary_detail = ("Segoe UI", 9)
        
        style = ttk.Style(self)
        style.configure("TLabelFrame.Label", font=self.font_titulo)

        self.selections = {'fonte': True, 'carga': True, 'multimetro': True}
        
        self.rm = pyvisa.ResourceManager()
        self.instruments = {}
        self.entries = {}
        
        self.unified_etapas = []
        self.editing_etapa_idx = None

        self.animation = None
        self.plot_window = None
        self.historico_timestamps = []
        self.historico_tensao = []
        self.historico_corrente = []

        # --- Frames Principais ---
        top_frame = tk.Frame(self)
        top_frame.pack(padx=10, pady=10, fill='x')
        
        control_frame = tk.Frame(self)
        control_frame.pack(pady=10, fill='x', padx=10)
        
        main_content_frame = tk.Frame(self)
        main_content_frame.pack(padx=10, pady=10, fill='both', expand=True)

        editor_panel = ttk.LabelFrame(main_content_frame, text="Configuração da Etapa")
        editor_panel.pack(side=tk.LEFT, fill='y', padx=(0, 10), anchor='n')
        self.editor_frame = editor_panel
        
        sequence_panel = ttk.LabelFrame(main_content_frame, text="Sequência de Testes")
        sequence_panel.pack(side=tk.LEFT, fill='both', expand=True)
        
        canvas = tk.Canvas(sequence_panel)
        scrollbar = ttk.Scrollbar(sequence_panel, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        config_ui_frame = ttk.LabelFrame(top_frame, text="Configurações Gerais e Conexão")
        config_ui_frame.pack(fill='x')

        self._create_connection_ui(config_ui_frame, 'fonte', 'Fonte (PWS4305)', "USB0::0x0699::0x0392::C010658::INSTR")
        self._create_connection_ui(config_ui_frame, 'carga', 'Carga Eletrônica', "USB0::0x05E6::0x2380::802436052757810021::INSTR")
        self._create_connection_ui(config_ui_frame, 'multimetro', 'Multímetro', "TCPIP0::172.30.248.100::3490::SOCKET")

        global_settings_frame = tk.Frame(config_ui_frame)
        global_settings_frame.pack(pady=5, padx=5, fill='x')

        tk.Label(global_settings_frame, text="Nome do Arquivo:", font=self.font_corpo).pack(side=tk.LEFT, padx=(0,5))
        entry_csv_name = tk.Entry(global_settings_frame, width=20, font=self.font_corpo)
        entry_csv_name.insert(0, "medicoes")
        entry_csv_name.pack(side=tk.LEFT)
        self.entries['multimetro_csv_name'] = entry_csv_name
        tk.Label(global_settings_frame, text=".csv", font=self.font_corpo).pack(side=tk.LEFT, padx=10)
        
        ### ALTERAÇÃO ### Adicionado campo para número de ciclos
        tk.Label(global_settings_frame, text="Número de Ciclos:", font=self.font_corpo).pack(side=tk.LEFT)
        self.cycles_entry = tk.Entry(global_settings_frame, width=5, font=self.font_corpo)
        self.cycles_entry.insert(0, "1")
        self.cycles_entry.pack(side=tk.LEFT, padx=5)

        self.plot_var = tk.BooleanVar(value=True)
        tk.Checkbutton(global_settings_frame, text="Gerar gráficos ao vivo", variable=self.plot_var, font=self.font_corpo).pack(side=tk.LEFT, padx=20)
        
        self.btn_conectar = tk.Button(control_frame, text="Conectar Equipamentos", command=self.conectar_todos, font=self.font_corpo)
        self.btn_conectar.pack(side=tk.LEFT, padx=5)
        self.btn_iniciar = tk.Button(control_frame, text="Iniciar Sequência", command=self.iniciar_sequencia, state=tk.DISABLED, font=self.font_corpo)
        self.btn_iniciar.pack(side=tk.LEFT, padx=5)
        self.btn_abrir_grafico = tk.Button(control_frame, text="Abrir Gráfico em Tempo Real", command=self.abrir_grafico_realtime, state=tk.DISABLED, font=self.font_corpo)
        self.btn_abrir_grafico.pack(side=tk.LEFT, padx=5)
        
        botoes_etapa_frame = tk.Frame(control_frame)
        botoes_etapa_frame.pack(side=tk.RIGHT, padx=5)
        tk.Button(botoes_etapa_frame, text="Salvar Etapa", command=self.add_or_update_etapa, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_etapa_frame, text="Remover Última Etapa", command=self.remover_ultima_etapa, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        
        status_frame = tk.Frame(self)
        status_frame.pack(fill='x', padx=10, pady=5, side=tk.BOTTOM)
        self.label_status_geral = tk.Label(status_frame, text="Status: Aguardando conexão...", font=("Segoe UI", 12))
        self.label_status_geral.pack(pady=5)
        
        self._create_editor_ui()
        self._clear_editor()

    def _create_editor_ui(self):
        etapa_frame = self.editor_frame
        self.editor_widgets = {}

        activation_frame = tk.Frame(etapa_frame)
        activation_frame.pack(pady=5, padx=5, fill='x')
        tk.Label(activation_frame, text="Equipamentos Ativos:", font=self.font_corpo).pack(side=tk.LEFT, padx=(0, 10))
        
        fonte_active_var = tk.BooleanVar(value=False)
        multi_active_var = tk.BooleanVar(value=False)
        carga_active_var = tk.BooleanVar(value=False)

        tk.Checkbutton(activation_frame, text="Fonte", variable=fonte_active_var, font=self.font_corpo).pack(side=tk.LEFT)
        tk.Checkbutton(activation_frame, text="Multímetro", variable=multi_active_var, font=self.font_corpo).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(activation_frame, text="Carga", variable=carga_active_var, font=self.font_corpo).pack(side=tk.LEFT)
        
        self.editor_widgets['fonte_active_var'] = fonte_active_var
        self.editor_widgets['multi_active_var'] = multi_active_var
        self.editor_widgets['carga_active_var'] = carga_active_var
        
        fonte_config_frame = ttk.LabelFrame(etapa_frame, text="Config. Fonte")
        stop_conditions_frame = ttk.LabelFrame(etapa_frame, text="Condições de Parada")
        carga_config_frame = ttk.LabelFrame(etapa_frame, text="Config. Carga")
        multi_config_frame = ttk.LabelFrame(etapa_frame, text="Config. Multímetro")
        
        fc_frame = tk.Frame(fonte_config_frame); fc_frame.pack(padx=5, pady=5)
        tk.Label(fc_frame, text="Tensão (V):", font=self.font_corpo).pack(side=tk.LEFT)
        entry_v = tk.Entry(fc_frame, width=7, font=self.font_corpo); entry_v.insert(0, "10.0"); entry_v.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(fc_frame, text="Corrente (A):", font=self.font_corpo).pack(side=tk.LEFT)
        entry_i = tk.Entry(fc_frame, width=7, font=self.font_corpo); entry_i.insert(0, "1.0"); entry_i.pack(side=tk.LEFT)
        self.editor_widgets.update({'entry_v': entry_v, 'entry_i': entry_i})

        cc_frame = tk.Frame(carga_config_frame); cc_frame.pack(padx=5, pady=5)
        var_modo = tk.StringVar(value="Resistência Constante (CR)")
        modos_com_cv = ["Corrente Constante (CC)", "Tensão Constante (CV)", "Potência Constante (CP)", "Resistência Constante (CR)"]
        modo_menu = ttk.Combobox(cc_frame, textvariable=var_modo, values=modos_com_cv, width=25, state="readonly", font=self.font_corpo); modo_menu.pack(side=tk.LEFT, padx=5)
        entry_val = tk.Entry(cc_frame, width=8, font=self.font_corpo); entry_val.insert(0, "100"); entry_val.pack(side=tk.LEFT, padx=5)
        self.editor_widgets.update({'var_modo': var_modo, 'entry_val': entry_val})

        mc_frame = tk.Frame(multi_config_frame); mc_frame.pack(padx=5, pady=5)
        volt_meas_var = tk.BooleanVar(value=False)
        curr_meas_var = tk.BooleanVar(value=False)
        tk.Checkbutton(mc_frame, text="Medir Tensão", variable=volt_meas_var, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(mc_frame, text="Medir Corrente", variable=curr_meas_var, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Label(mc_frame, text="Intervalo (s):", font=self.font_corpo).pack(side=tk.LEFT, padx=(10,5))
        entry_intervalo = tk.Entry(mc_frame, width=8, font=self.font_corpo); entry_intervalo.insert(0, "1.0"); entry_intervalo.pack(side=tk.LEFT)
        self.editor_widgets.update({'volt_meas_var': volt_meas_var, 'curr_meas_var': curr_meas_var, 'entry_intervalo': entry_intervalo})

        duracao_frame = tk.Frame(stop_conditions_frame); duracao_frame.pack(side=tk.LEFT, padx=10, pady=5)
        duracao_check_var = tk.BooleanVar(value=True)
        duracao_check = tk.Checkbutton(duracao_frame, text="Duração (s):", variable=duracao_check_var, font=self.font_corpo); duracao_check.pack(side=tk.LEFT)
        entry_duracao = tk.Entry(duracao_frame, width=7, font=self.font_corpo); entry_duracao.insert(0, "10.0"); entry_duracao.pack(side=tk.LEFT)
        
        volt_trigger_frame = tk.Frame(stop_conditions_frame); volt_trigger_frame.pack(side=tk.LEFT, padx=10, pady=5)
        volt_check_var = tk.BooleanVar(value=False)
        volt_check = tk.Checkbutton(volt_trigger_frame, text="Tensão:", variable=volt_check_var, font=self.font_corpo); volt_check.pack(side=tk.LEFT)
        volt_cond_var = tk.StringVar(value='<=')
        volt_cond_menu = ttk.Combobox(volt_trigger_frame, textvariable=volt_cond_var, values=['>=', '<='], width=3, state='disabled', font=self.font_corpo); volt_cond_menu.pack(side=tk.LEFT)
        entry_vt = tk.Entry(volt_trigger_frame, width=7, state='disabled', font=self.font_corpo); entry_vt.insert(0, "9.5"); entry_vt.pack(side=tk.LEFT)

        curr_trigger_frame = tk.Frame(stop_conditions_frame); curr_trigger_frame.pack(side=tk.LEFT, padx=10, pady=5)
        curr_check_var = tk.BooleanVar(value=False)
        curr_check = tk.Checkbutton(curr_trigger_frame, text="Corrente:", variable=curr_check_var, font=self.font_corpo); curr_check.pack(side=tk.LEFT)
        curr_cond_var = tk.StringVar(value='>=')
        curr_cond_menu = ttk.Combobox(curr_trigger_frame, textvariable=curr_cond_var, values=['>=', '<='], width=3, state='disabled', font=self.font_corpo); curr_cond_menu.pack(side=tk.LEFT)
        entry_ct = tk.Entry(curr_trigger_frame, width=7, state='disabled', font=self.font_corpo); entry_ct.insert(0, "100"); entry_ct.pack(side=tk.LEFT)
        curr_unit_var = tk.StringVar(value='mA')
        curr_unit_menu = ttk.Combobox(curr_trigger_frame, textvariable=curr_unit_var, values=['A', 'mA'], width=3, state='disabled', font=self.font_corpo); curr_unit_menu.pack(side=tk.LEFT, padx=(2,0))
        
        self.editor_widgets.update({
            'duracao_check_var': duracao_check_var, 'entry_duracao': entry_duracao,
            'volt_check_var': volt_check_var, 'volt_cond_var': volt_cond_var, 'entry_vt': entry_vt,
            'curr_check_var': curr_check_var, 'curr_cond_var': curr_cond_var, 'entry_ct': entry_ct, 'curr_unit_var': curr_unit_var
        })
        
        def update_ui_visibility_and_state(*args):
            fonte_config_frame.pack_forget()
            stop_conditions_frame.pack_forget()
            carga_config_frame.pack_forget()
            multi_config_frame.pack_forget()

            pack_options = {'pady': 5, 'padx': 5, 'fill': 'x'}
            
            if fonte_active_var.get(): fonte_config_frame.pack(**pack_options)
            stop_conditions_frame.pack(**pack_options)
            if carga_active_var.get(): carga_config_frame.pack(**pack_options)
            if multi_active_var.get(): multi_config_frame.pack(**pack_options)
            
            modos_sem_cv = ["Corrente Constante (CC)", "Potência Constante (CP)", "Resistência Constante (CR)"]
            if fonte_active_var.get() and carga_active_var.get():
                modo_menu['values'] = modos_sem_cv
                if var_modo.get() == "Tensão Constante (CV)": var_modo.set("Resistência Constante (CR)")
            else:
                modo_menu['values'] = modos_com_cv
            trigger_state = 'normal' if multi_active_var.get() else 'disabled'
            volt_check.config(state=trigger_state)
            curr_check.config(state=trigger_state)
            dummy_false_var = tk.BooleanVar(value=False)
            self._toggle_widgets_state(duracao_check_var, [entry_duracao])
            self._toggle_widgets_state(volt_check_var if trigger_state == 'normal' else dummy_false_var, [volt_cond_menu, entry_vt])
            self._toggle_widgets_state(curr_check_var if trigger_state == 'normal' else dummy_false_var, [curr_cond_menu, entry_ct, curr_unit_menu])

        fonte_active_var.trace_add("write", update_ui_visibility_and_state)
        carga_active_var.trace_add("write", update_ui_visibility_and_state)
        multi_active_var.trace_add("write", update_ui_visibility_and_state)
        duracao_check.config(command=update_ui_visibility_and_state)
        volt_check.config(command=update_ui_visibility_and_state)
        curr_check.config(command=update_ui_visibility_and_state)
        self.update_editor_visibility = update_ui_visibility_and_state

    def _clear_editor(self):
        self.editing_etapa_idx = None
        self.editor_frame.config(text="Configuração da Etapa: NOVA")
        self.editor_widgets['fonte_active_var'].set(False)
        self.editor_widgets['multi_active_var'].set(False)
        self.editor_widgets['carga_active_var'].set(False)
        self.editor_widgets['duracao_check_var'].set(True)
        self.editor_widgets['entry_duracao'].delete(0, tk.END); self.editor_widgets['entry_duracao'].insert(0, "10.0")
        self.editor_widgets['volt_check_var'].set(False)
        self.editor_widgets['curr_check_var'].set(False)
        self.editor_widgets['entry_v'].delete(0, tk.END); self.editor_widgets['entry_v'].insert(0, "10.0")
        self.editor_widgets['entry_i'].delete(0, tk.END); self.editor_widgets['entry_i'].insert(0, "1.0")
        self.editor_widgets['var_modo'].set("Resistência Constante (CR)")
        self.editor_widgets['entry_val'].delete(0, tk.END); self.editor_widgets['entry_val'].insert(0, "100")
        self.editor_widgets['volt_meas_var'].set(False)
        self.editor_widgets['curr_meas_var'].set(False)
        self.editor_widgets['entry_intervalo'].delete(0, tk.END); self.editor_widgets['entry_intervalo'].insert(0, "1.0")
        self.update_editor_visibility()
        self._update_sequence_display()

    def _load_etapa_to_editor(self, etapa_idx, event=None):
        if etapa_idx >= len(self.unified_etapas): return
        self.editing_etapa_idx = etapa_idx
        self.editor_frame.config(text=f"Configuração da Etapa: {etapa_idx + 1}")
        data = self.unified_etapas[etapa_idx]
        for key, value in data.items():
            if key in self.editor_widgets:
                widget_var = self.editor_widgets[key]
                if isinstance(widget_var, (tk.BooleanVar, tk.StringVar)):
                    widget_var.set(value)
                elif isinstance(widget_var, tk.Entry):
                    widget_var.delete(0, tk.END)
                    widget_var.insert(0, str(value))
        self.update_editor_visibility()
        self._update_sequence_display()

    def add_or_update_etapa(self):
        etapa_data = {}
        for key, widget in self.editor_widgets.items():
            value = widget.get() if hasattr(widget, 'get') else None
            etapa_data[key] = value
        if self.editing_etapa_idx is None:
            self.unified_etapas.append(etapa_data)
        else:
            self.unified_etapas[self.editing_etapa_idx] = etapa_data
        self._clear_editor()
    
    def remover_ultima_etapa(self):
        if not self.unified_etapas:
            messagebox.showwarning("Aviso", "Não há etapas para remover.")
            return
        if self.editing_etapa_idx == len(self.unified_etapas) - 1:
            self._clear_editor()
        self.unified_etapas.pop()
        self._update_sequence_display()

    def _update_sequence_display(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        for i, etapa_data in enumerate(self.unified_etapas):
            summary_frame = tk.Frame(self.scrollable_frame, relief="raised", borderwidth=2, pady=5, padx=5)
            summary_frame.pack(fill='x', padx=5, pady=3)
            bg_color = "lightblue" if i == self.editing_etapa_idx else summary_frame.cget('bg')
            summary_frame.config(bg=bg_color)
            
            active_equipment = []
            if etapa_data.get('fonte_active_var'): active_equipment.append("Fonte")
            if etapa_data.get('carga_active_var'): active_equipment.append("Carga")
            if etapa_data.get('multimetro_active_var'): active_equipment.append("Multímetro")
            title_text = f"Etapa {i+1}: {' / '.join(active_equipment) if active_equipment else 'Nenhum equipamento ativo'}"
            tk.Label(summary_frame, text=title_text, font=self.font_summary_title, bg=bg_color).pack(anchor='w')

            ttk.Separator(summary_frame, orient='horizontal').pack(fill='x', pady=4)
            
            if etapa_data.get('fonte_active_var'):
                tk.Label(summary_frame, text="Config. Fonte", font=self.font_summary_subtitle, bg=bg_color).pack(anchor='w')
                v = etapa_data.get('entry_v', 'N/A')
                i_val = etapa_data.get('entry_i', 'N/A')
                fonte_details = f"  • Tensão: {v}V / Corrente: {i_val}A"
                tk.Label(summary_frame, text=fonte_details, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')
            
            tk.Label(summary_frame, text="Condições de Parada", font=self.font_summary_subtitle, bg=bg_color).pack(anchor='w')
            if etapa_data.get('duracao_check_var'):
                dur_text = f"  • Duração: {etapa_data.get('entry_duracao', 'N/A')}s"
                tk.Label(summary_frame, text=dur_text, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')
            if etapa_data.get('volt_check_var'):
                cond = etapa_data.get('volt_cond_var', '')
                val = etapa_data.get('entry_vt', 'N/A')
                volt_trigger_text = f"  • Gatilho Tensão: {cond} {val}V"
                tk.Label(summary_frame, text=volt_trigger_text, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')
            if etapa_data.get('curr_check_var'):
                cond = etapa_data.get('curr_cond_var', '')
                val = etapa_data.get('entry_ct', 'N/A')
                unit = etapa_data.get('curr_unit_var', '')
                curr_trigger_text = f"  • Gatilho Corrente: {cond} {val}{unit}"
                tk.Label(summary_frame, text=curr_trigger_text, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')

            if etapa_data.get('carga_active_var'):
                tk.Label(summary_frame, text="Config. Carga", font=self.font_summary_subtitle, bg=bg_color).pack(anchor='w')
                modo = etapa_data.get('var_modo', 'N/A').split('(')[0].strip()
                val = etapa_data.get('entry_val', 'N/A')
                carga_details = f"  • Modo: {modo} = {val}"
                tk.Label(summary_frame, text=carga_details, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')

            if etapa_data.get('multimetro_active_var'):
                tk.Label(summary_frame, text="Config. Multímetro", font=self.font_summary_subtitle, bg=bg_color).pack(anchor='w')
                medidas = []
                if etapa_data.get('volt_meas_var'): medidas.append("Tensão")
                if etapa_data.get('curr_meas_var'): medidas.append("Corrente")
                med_text = ", ".join(medidas) if medidas else "Nenhuma"
                
                multi_text1 = f"  • Medições: {med_text}"
                multi_text2 = f"  • Intervalo: {etapa_data.get('entry_intervalo', 'N/A')}s"
                tk.Label(summary_frame, text=multi_text1, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')
                tk.Label(summary_frame, text=multi_text2, font=self.font_summary_detail, bg=bg_color).pack(anchor='w')

            click_handler = partial(self._load_etapa_to_editor, i)
            summary_frame.bind("<Button-1>", click_handler)
            for child in summary_frame.winfo_children():
                if isinstance(child, (tk.Label, tk.Frame)):
                    child.bind("<Button-1>", click_handler)

    def _create_connection_ui(self, parent, key, text, default_addr):
        frame = tk.Frame(parent)
        frame.pack(pady=5, padx=5, fill='x')
        tk.Label(frame, text=f"{text} VISA:", font=self.font_corpo).pack(side=tk.LEFT, padx=(0,5))
        entry_addr = tk.Entry(frame, width=40, font=self.font_corpo)
        entry_addr.insert(0, default_addr)
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True, fill='x')
        tk.Button(frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr), font=self.font_corpo).pack(side=tk.LEFT)
        self.entries[f'{key}_addr'] = entry_addr
        
    def _toggle_widgets_state(self, active_var, widgets_list):
        state = 'normal' if active_var.get() else 'disabled'
        for w in widgets_list:
            if isinstance(w, ttk.Combobox):
                w.config(state='readonly' if active_var.get() else 'disabled')
            else:
                w.config(state=state)
    
    def buscar_enderecos(self, entry_widget):
        try:
            recursos = self.rm.list_resources()
            if not recursos:
                messagebox.showinfo("Nenhum dispositivo", "Nenhum dispositivo VISA encontrado.")
                return
            janela_lista = tk.Toplevel(self)
            janela_lista.title("Dispositivos encontrados")
            janela_lista.geometry("420x300")
            janela_lista.grab_set()
            tk.Label(janela_lista, text="Selecione um endereço:", font=("Arial", 12)).pack(pady=5)
            lista = tk.Listbox(janela_lista, width=60, font=self.font_corpo)
            lista.pack(pady=5, padx=5, fill='both', expand=True)
            recursos_filtrados = [r for r in recursos if r.startswith('USB') or r.startswith('TCPIP')]
            for r in recursos_filtrados:
                lista.insert(tk.END, r) 
            def selecionar():
                if lista.curselection(): 
                    selecionado = lista.get(tk.ACTIVE)
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, selecionado)
                janela_lista.destroy()
            tk.Button(janela_lista, text="Selecionar", command=selecionar, font=self.font_corpo).pack(pady=5)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dispositivos:\n{e}")

    def conectar_todos(self):
        self.instruments = {}
        log = []
        try:
            if self.selections['fonte']:
                inst = self.rm.open_resource(self.entries['fonte_addr'].get())
                inst.timeout = 5000
                self.instruments['fonte'] = inst
                log.append(f"Fonte Conectada: {inst.query('*IDN?').strip()}")
            if self.selections['multimetro']:
                inst = self.rm.open_resource(self.entries['multimetro_addr'].get())
                inst.timeout = 5000
                inst.write_termination = '\n'
                inst.read_termination = '\n'
                self.instruments['multimetro'] = inst
                log.append(f"Multímetro Conectado: {inst.query('*IDN?').strip()}")
            if self.selections['carga']:
                inst = self.rm.open_resource(self.entries['carga_addr'].get())
                inst.timeout = 5000
                self.instruments['carga'] = inst
                log.append(f"Carga Conectada: {inst.query('*IDN?').strip()}")
            messagebox.showinfo("Conexão Bem-Sucedida", "\n".join(log))
            self.btn_iniciar.config(state=tk.NORMAL)
            self.label_status_geral.config(text="Status: Conectado. Pronto para iniciar.")
        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Falha ao conectar com um ou mais dispositivos:\n{e}")
            self.btn_iniciar.config(state=tk.DISABLED)
            self.label_status_geral.config(text="Status: Falha na conexão.")

    def iniciar_sequencia(self):
        self.historico_timestamps.clear()
        self.historico_tensao.clear()
        self.historico_corrente.clear()
        self.btn_abrir_grafico.config(state=tk.DISABLED, text="Abrir Gráfico em Tempo Real")
        if self.plot_window: self._on_plot_close()
        self.volt_a_medir = any(etapa.get('multi_active_var') and etapa.get('volt_meas_var') for etapa in self.unified_etapas)
        self.curr_a_medir = any(etapa.get('multi_active_var') and etapa.get('curr_meas_var') for etapa in self.unified_etapas)
        threading.Thread(target=self.executar_sequencia, daemon=True).start()

    ### ALTERAÇÃO ### Função de execução agora inclui o loop de ciclos.
    def executar_sequencia(self):
        should_plot = self.plot_var.get() and (self.volt_a_medir or self.curr_a_medir)
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_conectar.config(state=tk.DISABLED)
        if should_plot: self.btn_abrir_grafico.config(state=tk.NORMAL)
        
        try:
            num_cycles = int(self.cycles_entry.get())
            if num_cycles < 1:
                raise ValueError("O número de ciclos deve ser pelo menos 1.")
        except ValueError as e:
            self.after(0, lambda: messagebox.showerror("Valor Inválido", f"Número de ciclos inválido: {e}"))
            self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
            return

        csv_filename = ""
        try:
            base_name = self.entries['multimetro_csv_name'].get().strip()
            if not base_name:
                self.after(0, lambda: messagebox.showwarning("Nome Inválido", "O nome do arquivo CSV não pode estar em branco."))
                self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                return
            
            documentos_dir = os.path.join(os.path.expanduser('~'), 'Documents')
            os.makedirs(documentos_dir, exist_ok=True)
            csv_filename = os.path.join(documentos_dir, f"{base_name}.csv")
            
            ### ALTERAÇÃO ### Adicionada coluna 'Ciclo' ao CSV
            csv_header = ["Timestamp", "Ciclo", "Etapa", "Tensao_Fonte_Set", "Corrente_Fonte_Set", "Modo_Carga", "Valor_Carga_Set", "Tensao_Multimetro", "Corrente_Multimetro"]
            
            with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(csv_header)
                if not self.unified_etapas:
                    self.after(0, lambda: messagebox.showinfo("Aviso", "Nenhuma etapa configurada."))
                    return

                self.current_volt_range = 100 

                # --- Loop de Ciclos ---
                for cycle_num in range(1, num_cycles + 1):
                    for i, etapa_config in enumerate(self.unified_etapas):
                        etapa_num = i + 1
                        self.label_status_geral.config(text=f"Ciclo {cycle_num}/{num_cycles} - Configurando Etapa {etapa_num}...")
                        
                        fonte = self.instruments.get('fonte')
                        carga = self.instruments.get('carga')
                        multimetro = self.instruments.get('multimetro')
                        
                        v_set, i_set, modo_carga_set, valor_carga_set = "N/A", "N/A", "N/A", "N/A"
                        fonte_ativa = etapa_config.get('fonte_active_var')
                        carga_ativa = etapa_config.get('carga_active_var')
                        multi_ativo = etapa_config.get('multi_active_var')

                        # Configura a fonte (ativa ou desativa)
                        if fonte_ativa and fonte:
                            v_set = float(etapa_config.get('entry_v')); i_set = float(etapa_config.get('entry_i'))
                            fonte.write(f"VOLT {v_set}"); fonte.write(f"CURR {i_set}"); fonte.write("OUTP ON")
                        elif fonte:
                            fonte.write("OUTP OFF")

                        # Configura a carga (ativa ou desativa)
                        if carga_ativa and carga:
                            modo_carga_set = etapa_config.get('var_modo'); valor_carga_set = etapa_config.get('entry_val')
                            sigla = modo_carga_set.split('(')[1].replace(')', '')
                            valor_carga_num = float(valor_carga_set) if valor_carga_set else 0
                            cmd_map = {"CC": "CURR", "CV": "VOLT", "CR": "RES", "CP": "POW"}
                            if sigla in cmd_map:
                                carga.write(f"FUNC {cmd_map[sigla]}"); carga.write(f"{cmd_map[sigla]} {valor_carga_num}")
                            carga.write("INPUT ON")
                        elif carga:
                            carga.write("INPUT OFF")
                        
                        time.sleep(0.5)
                        self.label_status_geral.config(text=f"Ciclo {cycle_num}/{num_cycles} - Executando Etapa {etapa_num}...")
                        
                        parar_por_duracao = etapa_config.get('duracao_check_var')
                        etapa_duracao = float(etapa_config.get('entry_duracao')) if parar_por_duracao else float('inf')
                        start_time = time.time()

                        if multi_ativo and multimetro:
                            intervalo = float(etapa_config.get('entry_intervalo'))
                            medir_v = etapa_config.get('volt_meas_var')
                            medir_i = etapa_config.get('curr_meas_var')
                            
                            while time.time() - start_time < etapa_duracao:
                                loop_start_time = time.time()
                                tensao_str, corrente_str, v_num, i_num = "N/A", "N/A", np.nan, np.nan
                                try:
                                    if medir_v:
                                        multimetro.write(f"CONF:VOLT:DC {self.current_volt_range}"); multimetro.write("INIT")
                                        v_num = float(multimetro.query("FETCH?").strip().split(',')[0])
                                        tensao_str = f"{v_num:.5f}"
                                        self.current_volt_range = 10 if v_num < 10 else 100
                                    if medir_i:
                                        multimetro.write("CONF:CURR:DC 10"); multimetro.write("INIT")
                                        i_num = float(multimetro.query("FETCH?").strip().split(',')[0])
                                        corrente_str = f"{i_num:.5f}"
                                except Exception as e:
                                    print(f"Erro no multímetro: {e}"); tensao_str = corrente_str = "ERRO"
                                
                                self.historico_timestamps.append(datetime.datetime.now())
                                self.historico_tensao.append(v_num if medir_v else np.nan)
                                self.historico_corrente.append(i_num if medir_i else np.nan)
                                
                                ### ALTERAÇÃO ### Adicionado `cycle_num` ao registro do CSV
                                csv_writer.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], cycle_num, etapa_num, v_set, i_set, modo_carga_set, valor_carga_set, tensao_str, corrente_str])
                                csv_file.flush()
                                self.label_status_geral.config(text=f"Ciclo {cycle_num}/{num_cycles} Etapa {etapa_num}: V={tensao_str} | I={corrente_str}")
                                
                                stop_reason = self.check_stop_conditions(etapa_config, v_num, i_num)
                                if stop_reason:
                                    self.label_status_geral.config(text=f"Ciclo {cycle_num}/{num_cycles} Etapa {etapa_num}: {stop_reason}!")
                                    break
                                
                                elapsed_in_loop = time.time() - loop_start_time
                                sleep_duration = intervalo - elapsed_in_loop
                                if sleep_duration > 0: time.sleep(sleep_duration)
                        else:
                            csv_writer.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], cycle_num, etapa_num, v_set, i_set, modo_carga_set, valor_carga_set, "N/A", "N/A"])
                            csv_file.flush()
                            if parar_por_duracao: time.sleep(etapa_duracao)

                # Desliga os equipamentos ao final de cada etapa do ciclo
                if fonte: fonte.write("OUTP OFF")
                if carga: carga.write("INPUT OFF")
                self.label_status_geral.config(text=f"Ciclo {cycle_num}/{num_cycles} Etapa {etapa_num} concluída.")
                time.sleep(0.5)

            final_message = f"Sequência finalizada após {num_cycles} ciclo(s). Dados salvos em '{os.path.basename(csv_filename)}'."
            self.label_status_geral.config(text=f"Status: {final_message}")
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Erro na Sequência", f"Ocorreu um erro durante a execução:\n{e}"))
            self.label_status_geral.config(text="Status: Erro na sequência.")
        finally:
            try:
                if 'fonte' in self.instruments and self.instruments.get('fonte'): self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments and self.instruments.get('carga'): self.instruments['carga'].write("INPUT OFF")
            except Exception: pass
            self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
            if self.historico_timestamps and (self.volt_a_medir or self.curr_a_medir): 
                self.btn_abrir_grafico.config(state=tk.NORMAL, text="Ver Gráfico do Resultado")

    def check_stop_conditions(self, config, v_num, i_num):
        try:
            if config.get('volt_check_var') and not np.isnan(v_num):
                alvo = float(config.get('entry_vt')); cond = config.get('volt_cond_var')
                if (cond == '>=' and v_num >= alvo) or (cond == '<=' and v_num <= alvo): return "Gatilho de Tensão Atingido"
            if config.get('curr_check_var') and not np.isnan(i_num):
                alvo = float(config.get('entry_ct')); unit = config.get('curr_unit_var'); cond = config.get('curr_cond_var')
                alvo_a = alvo / 1000.0 if unit == 'mA' else alvo
                if (cond == '>=' and i_num >= alvo_a) or (cond == '<=' and i_num <= alvo_a): return "Gatilho de Corrente Atingido"
        except (ValueError, TypeError) as e: print(f"Erro ao avaliar gatilho: {e}")
        return None

    def abrir_grafico_realtime(self):
        if self.plot_window and self.plot_window.winfo_exists():
            self.plot_window.lift()
        else:
            self._setup_realtime_plot()

    def _setup_realtime_plot(self):
        self.plot_window = tk.Toplevel(self)
        self.plot_window.title("Gráfico em Tempo Real")
        self.plot_window.geometry("1000x800")
        self.plot_window.protocol("WM_DELETE_WINDOW", self._on_plot_close)
        self.fig = Figure(figsize=(10, 8), dpi=100)
        
        self.plot_timestamps = self.historico_timestamps.copy()
        self.plot_tensao_data = self.historico_tensao.copy()
        self.plot_corrente_data = self.historico_corrente.copy()
        
        num_plots = (1 if self.volt_a_medir else 0) + (1 if self.curr_a_medir else 0)
        if num_plots == 0:
            self.plot_window.destroy(); messagebox.showinfo("Gráfico", "Nenhuma medição selecionada."); return
        
        plot_index = 1
        self.ax1, self.ax2, self.line1, self.line2 = None, None, None, None
        if self.volt_a_medir:
            self.ax1 = self.fig.add_subplot(num_plots, 1, plot_index)
            self.line1, = self.ax1.plot(self.plot_timestamps, self.plot_tensao_data, marker='.', linestyle='-', label='Tensão (V)')
            self.ax1.set_title("Medições de Tensão vs Tempo"); self.ax1.set_ylabel("Tensão (V)"); self.ax1.grid(True); self.ax1.legend()
            plot_index += 1
        if self.curr_a_medir:
            self.ax2 = self.fig.add_subplot(num_plots, 1, plot_index)
            self.line2, = self.ax2.plot(self.plot_timestamps, self.plot_corrente_data, marker='.', linestyle='-', color='r', label='Corrente (A)')
            self.ax2.set_title("Medições de Corrente vs Tempo"); self.ax2.set_ylabel("Corrente (A)"); self.ax2.grid(True); self.ax2.legend()
        
        self.fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(self.fig, master=self.plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        toolbar = NavigationToolbar2Tk(canvas, self.plot_window)
        toolbar.update()
        self.animation = animation.FuncAnimation(self.fig, self._update_plot, interval=500, blit=False, cache_frame_data=False)

    def _update_plot(self, frame):
        try:
            if len(self.historico_timestamps) > len(self.plot_timestamps):
                self.plot_timestamps = self.historico_timestamps.copy()
                self.plot_tensao_data = self.historico_tensao.copy()
                self.plot_corrente_data = self.historico_corrente.copy()
                if self.ax1:
                    self.line1.set_data(self.plot_timestamps, self.plot_tensao_data)
                    self.ax1.relim(); self.ax1.autoscale_view()
                if self.ax2:
                    self.line2.set_data(self.plot_timestamps, self.plot_corrente_data)
                    self.ax2.relim(); self.ax2.autoscale_view()
                if self.plot_window: self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"Erro ao atualizar o gráfico: {e}")

    def _on_plot_close(self):
        if self.animation:
            self.animation.event_source.stop()
            self.animation = None
        if self.plot_window:
            self.plot_window.destroy()
            self.plot_window = None

    def on_close(self):
        try:
            self._on_plot_close()
            for key in ['fonte', 'carga', 'multimetro']:
                if key in self.instruments and self.instruments[key]:
                    self.instruments[key].close()
        except Exception as e:
            print(f"Erro ao fechar conexões: {e}")
        finally:
            self.destroy()

if __name__ == "__main__":
    app = JanelaControleCombinado()
    app.mainloop()