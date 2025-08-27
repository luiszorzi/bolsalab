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
import numpy as np # Necessário para np.nan

# JANELA PRINCIPAL DE SELEÇÃO
class JanelaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Equipamentos")
        self.geometry("350x250")

        self.use_fonte = tk.BooleanVar(value=False)
        self.use_multimetro = tk.BooleanVar(value=False)
        self.use_carga = tk.BooleanVar(value=False)

        tk.Label(self, text="Escolher equipamentos:", font=("Arial", 16)).pack(pady=20)

        tk.Checkbutton(self, text="Fonte de Alimentação", variable=self.use_fonte, font=("Arial", 12)).pack(anchor='w', padx=40)
        tk.Checkbutton(self, text="Multímetro", variable=self.use_multimetro, font=("Arial", 12)).pack(anchor='w', padx=40)
        tk.Checkbutton(self, text="Carga Eletrônica", variable=self.use_carga, font=("Arial", 12)).pack(anchor='w', padx=40)

        btn_avancar = tk.Button(self, text="Avançar →", width=25, command=self.abrir_controle_combinado)
        btn_avancar.pack(pady=25)

    def abrir_controle_combinado(self):
        selections = {
            'fonte': self.use_fonte.get(),
            'multimetro': self.use_multimetro.get(),
            'carga': self.use_carga.get()
        }

        if not any(selections.values()):
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione pelo menos um equipamento para continuar.")
            return

        self.withdraw()
        control_window = JanelaControleCombinado(self, selections)
        control_window.grab_set()


# JANELA DE CONTROLE COMBINADO
class JanelaControleCombinado(tk.Toplevel):
    def __init__(self, master, selections):
        super().__init__(master)
        self.title("Controle de Equipamentos")
        self.geometry("880x920") 
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.font_titulo = ("Segoe UI", 12, "bold")
        self.font_corpo = ("Segoe UI", 11)
        
        style = ttk.Style(self)
        style.configure("TLabelFrame.Label", font=self.font_titulo)

        self.selections = selections
        self.rm = pyvisa.ResourceManager()
        self.instruments = {}
        self.entries = {}
        self.etapas = {'fonte': [], 'carga': []}
        self.labels = {}
        self.frames = {}
        
        self.fonte_e_carga_juntas = self.selections['fonte'] and self.selections['carga']
        
        self.animation = None
        self.plot_window = None

        # Listas para armazenar todo o histórico da medição atual
        self.historico_timestamps = []
        self.historico_tensao = []
        self.historico_corrente = []

        main_frame = tk.Frame(self)
        main_frame.pack(padx=10, pady=10, fill='both', expand=True)

        if self.selections['fonte']:
            self._create_fonte_ui(main_frame)
        if self.selections['multimetro']:
            self._create_multimetro_ui(main_frame)
        if self.selections['carga']:
            self._create_carga_ui(main_frame)

        control_frame = tk.Frame(self)
        control_frame.pack(pady=10, fill='x', padx=10)
        
        self.btn_conectar = tk.Button(control_frame, text="Conectar Equipamentos", command=self.conectar_todos, font=self.font_corpo)
        self.btn_conectar.pack(pady=5)

        self.btn_iniciar = tk.Button(control_frame, text="Iniciar Sequência", command=self.iniciar_sequencia, state=tk.DISABLED, font=self.font_corpo)
        self.btn_iniciar.pack(pady=5)
        
        self.btn_abrir_grafico = tk.Button(control_frame, text="Abrir Gráfico em Tempo Real", command=self.abrir_grafico_realtime, state=tk.DISABLED, font=self.font_corpo)
        self.btn_abrir_grafico.pack(pady=5)

        self.label_status_geral = tk.Label(control_frame, text="Status: Aguardando conexão...", font=("Segoe UI", 12))
        self.label_status_geral.pack(pady=10)

        self.btn_voltar = tk.Button(control_frame, text="← Voltar", command=self.on_close, font=self.font_corpo)
        self.btn_voltar.pack(side=tk.BOTTOM, pady=10)

    def _create_fonte_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Fonte de Alimentação (PWS4305)")
        frame.pack(pady=10, padx=5, fill='x')
        self.frames['fonte'] = frame

        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:", font=self.font_corpo).pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40, font=self.font_corpo)
        entry_addr.insert(0, "USB0::0x0699::0x0392::C010658::INSTR")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr), font=self.font_corpo).pack(side=tk.LEFT)
        self.entries['fonte_addr'] = entry_addr

        self.frames['fonte_etapas'] = tk.Frame(frame)
        self.frames['fonte_etapas'].pack(pady=5, padx=5, fill='x')

        botoes_etapa_frame = tk.Frame(frame)
        botoes_etapa_frame.pack(pady=5)
        tk.Button(botoes_etapa_frame, text="+ Adicionar Etapa", command=self.adicionar_etapa_fonte, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_etapa_frame, text="- Remover Etapa", command=self.remover_etapa_fonte, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        
        self.adicionar_etapa_fonte()

    def _create_multimetro_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Multímetro")
        frame.pack(pady=10, padx=5, fill='x')
        
        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:", font=self.font_corpo).pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40, font=self.font_corpo)
        entry_addr.insert(0, "TCPIP0::172.30.248.100::3490::SOCKET")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr), font=self.font_corpo).pack(side=tk.LEFT)
        self.entries['multimetro_addr'] = entry_addr
        
        config_frame = tk.Frame(frame)
        config_frame.pack(pady=5, padx=5, fill='x')
        
        self.volt_meas_var = tk.BooleanVar(value=True)
        self.curr_meas_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(config_frame, text="Medir Tensão", variable=self.volt_meas_var, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(config_frame, text="Medir Corrente", variable=self.curr_meas_var, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(config_frame, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10, pady=5)
        
        tk.Label(config_frame, text="Intervalo (s):", font=self.font_corpo).pack(side=tk.LEFT, padx=(0,5))
        entry_intervalo = tk.Entry(config_frame, width=8, font=self.font_corpo)
        entry_intervalo.insert(0, "1.0")
        entry_intervalo.pack(side=tk.LEFT)
        self.entries['multimetro_intervalo'] = entry_intervalo
        
        tk.Label(config_frame, text="Nome do Arquivo:", font=self.font_corpo).pack(side=tk.LEFT, padx=(10,5))
        
        filename_frame = tk.Frame(config_frame)
        filename_frame.pack(side=tk.LEFT, expand=True, fill='x')

        entry_csv_name = tk.Entry(filename_frame, width=20, font=self.font_corpo)
        entry_csv_name.insert(0, "medicoes")
        entry_csv_name.pack(side=tk.LEFT)
        self.entries['multimetro_csv_name'] = entry_csv_name

        tk.Label(filename_frame, text=".csv", font=self.font_corpo).pack(side=tk.LEFT)

        plot_frame = tk.Frame(frame)
        plot_frame.pack(pady=10, padx=5, fill='x')
        
        self.plot_var = tk.BooleanVar(value=True)
        plot_check = tk.Checkbutton(plot_frame, text="Gerar gráficos", variable=self.plot_var, font=self.font_corpo)
        plot_check.pack(anchor='w')

    def _create_carga_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Carga Eletrônica")
        frame.pack(pady=10, padx=5, fill='x')
        self.frames['carga'] = frame
        
        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:", font=self.font_corpo).pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40, font=self.font_corpo)
        entry_addr.insert(0, "USB0::0x05E6::0x2380::802436052757810021::INSTR")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr), font=self.font_corpo).pack(side=tk.LEFT)
        self.entries['carga_addr'] = entry_addr
        
        self.frames['carga_etapas'] = tk.Frame(frame)
        self.frames['carga_etapas'].pack(pady=5, padx=5)
        
        botoes_etapa_frame = tk.Frame(frame)
        botoes_etapa_frame.pack(pady=5)
        tk.Button(botoes_etapa_frame, text="+ Adicionar Etapa", command=self.adicionar_etapa_carga, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_etapa_frame, text="- Remover Etapa", command=self.remover_etapa_carga, font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        
        self.adicionar_etapa_carga()

    def _toggle_entry_state(self, check_var, *widgets):
        state = tk.NORMAL if check_var.get() else tk.DISABLED
        for widget in widgets:
            widget.config(state=state)

    def adicionar_etapa_fonte(self):
        etapa_idx = len(self.etapas['fonte'])
        etapa_frame = ttk.LabelFrame(self.frames['fonte_etapas'], text=f"Etapa {etapa_idx + 1}")
        etapa_frame.pack(pady=5, padx=5, fill='x')
        
        config_frame = tk.Frame(etapa_frame)
        config_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(config_frame, text="Tensão (V):", font=self.font_corpo).pack(side=tk.LEFT)
        entry_v = tk.Entry(config_frame, width=7, font=self.font_corpo)
        entry_v.insert(0, "10.0")
        entry_v.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(config_frame, text="Corrente (A):", font=self.font_corpo).pack(side=tk.LEFT)
        entry_i = tk.Entry(config_frame, width=7, font=self.font_corpo)
        entry_i.insert(0, "1.0")
        entry_i.pack(side=tk.LEFT)
        
        trigger_frame = tk.Frame(etapa_frame)
        trigger_frame.pack(pady=5, fill='x', padx=5)
        
        time_check_var = tk.BooleanVar(value=True)
        volt_check_var = tk.BooleanVar(value=False)
        curr_check_var = tk.BooleanVar(value=False)

        time_check = tk.Checkbutton(trigger_frame, text="Duração (s):", variable=time_check_var, font=self.font_corpo)
        time_check.pack(side=tk.LEFT)
        entry_t = tk.Entry(trigger_frame, width=7, font=self.font_corpo)
        entry_t.insert(0, "5")
        entry_t.pack(side=tk.LEFT)
        
        volt_check = tk.Checkbutton(trigger_frame, text="Tensão:", variable=volt_check_var, font=self.font_corpo)
        volt_check.pack(side=tk.LEFT, padx=(15, 0))
        volt_cond_var = tk.StringVar(value='>=')
        volt_cond_menu = ttk.Combobox(trigger_frame, textvariable=volt_cond_var, values=['>=', '<='], width=3, state=tk.DISABLED, font=self.font_corpo)
        volt_cond_menu.pack(side=tk.LEFT)
        entry_vt = tk.Entry(trigger_frame, width=7, state=tk.DISABLED, font=self.font_corpo)
        entry_vt.insert(0, "9.5")
        entry_vt.pack(side=tk.LEFT)
        
        curr_check = tk.Checkbutton(trigger_frame, text="Corrente:", variable=curr_check_var, font=self.font_corpo)
        curr_check.pack(side=tk.LEFT, padx=(15, 0))
        curr_cond_var = tk.StringVar(value='>=')
        curr_cond_menu = ttk.Combobox(trigger_frame, textvariable=curr_cond_var, values=['>=', '<='], width=3, state=tk.DISABLED, font=self.font_corpo)
        curr_cond_menu.pack(side=tk.LEFT)
        entry_ct = tk.Entry(trigger_frame, width=7, state=tk.DISABLED, font=self.font_corpo)
        entry_ct.insert(0, "100")
        entry_ct.pack(side=tk.LEFT)

        curr_unit_var = tk.StringVar(value='mA')
        curr_unit_menu = ttk.Combobox(trigger_frame, textvariable=curr_unit_var, values=['A', 'mA'], width=3, state=tk.DISABLED, font=self.font_corpo)
        curr_unit_menu.pack(side=tk.LEFT, padx=(2,0))
        
        def update_trigger_selection(selected_var, other_var1, other_var2):
            if selected_var.get():
                other_var1.set(False)
                other_var2.set(False)
            
            self._toggle_entry_state(time_check_var, entry_t)
            self._toggle_entry_state(volt_check_var, volt_cond_menu, entry_vt)
            self._toggle_entry_state(curr_check_var, curr_cond_menu, entry_ct, curr_unit_menu)

        time_check.config(command=lambda: update_trigger_selection(time_check_var, volt_check_var, curr_check_var))
        volt_check.config(command=lambda: update_trigger_selection(volt_check_var, time_check_var, curr_check_var))
        curr_check.config(command=lambda: update_trigger_selection(curr_check_var, time_check_var, volt_check_var))
        
        widgets = {
            'frame': etapa_frame, 'v_entry': entry_v, 'i_entry': entry_i,
            'time_check_var': time_check_var, 'time_entry': entry_t,
            'volt_check_var': volt_check_var, 'volt_cond_var': volt_cond_var, 'volt_target_entry': entry_vt,
            'curr_check_var': curr_check_var, 'curr_cond_var': curr_cond_var, 'curr_target_entry': entry_ct,
            'curr_unit_var': curr_unit_var,
        }
        self.etapas['fonte'].append(widgets)

    def remover_etapa_fonte(self):
        if not self.etapas['fonte']:
            messagebox.showwarning("Aviso", "Nenhuma etapa da fonte para remover.")
            return
        widgets_a_remover = self.etapas['fonte'].pop()
        widgets_a_remover['frame'].destroy()
        for i, widgets in enumerate(self.etapas['fonte']):
            widgets['frame'].config(text=f"Etapa {i + 1}")

    def adicionar_etapa_carga(self,):
        frame = tk.Frame(self.frames['carga_etapas'])
        frame.pack(pady=2, fill='x')
        
        tk.Label(frame, text=f"Etapa {len(self.etapas['carga']) + 1}:", font=self.font_corpo).pack(side=tk.LEFT, padx=5)
        var_modo = tk.StringVar(value="Resistência Constante (CR)")
        
        if self.fonte_e_carga_juntas:
            modos = ["Corrente Constante (CC)", "Potência Constante (CP)", "Resistência Constante (CR)"]
        else:
            modos = ["Corrente Constante (CC)", "Tensão Constante (CV)", "Potência Constante (CP)", "Resistência Constante (CR)"]
        
        modo_menu = ttk.Combobox(frame, textvariable=var_modo, values=modos, width=25, state="readonly", font=self.font_corpo)
        modo_menu.pack(side=tk.LEFT, padx=5)
        
        entry_val = tk.Entry(frame, width=8, font=self.font_corpo)
        entry_val.insert(0, "100")
        entry_val.pack(side=tk.LEFT, padx=5)
        
        self.etapas['carga'].append((frame, var_modo, entry_val))

    def remover_etapa_carga(self):
        if not self.etapas['carga']:
            messagebox.showwarning("Aviso", "Nenhuma etapa da carga para remover.")
            return
        frame_a_remover, _, _ = self.etapas['carga'].pop()
        frame_a_remover.destroy()
        for i, (frame, _, _) in enumerate(self.etapas['carga']):
            frame.winfo_children()[0].config(text=f"Etapa {i + 1}:")

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
                else:
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

    def abrir_grafico_realtime(self):
        if self.plot_window:
            self.plot_window.lift()
            return

        self.csv_header = ["Timestamp", "Etapa", "Tensao_Fonte", "Corrente_Fonte", "Modo_Carga", "Valor_Carga"]
        if self.volt_meas_var.get(): self.csv_header.append("Tensao_Multimetro")
        if self.curr_meas_var.get(): self.csv_header.append("Corrente_Multimetro")
        
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

        tensao_disponivel = 'Tensao_Multimetro' in self.csv_header
        corrente_disponivel = 'Corrente_Multimetro' in self.csv_header
        
        num_plots = int(tensao_disponivel) + int(corrente_disponivel)
        if num_plots == 0: return 
        plot_index = 1

        self.ax1, self.ax2 = None, None
        
        if tensao_disponivel:
            self.ax1 = self.fig.add_subplot(num_plots, 1, plot_index)
            self.line1, = self.ax1.plot(self.plot_timestamps, self.plot_tensao_data, marker='.', linestyle='-', label='Tensão (V)')
            self.ax1.set_title("Medições de Tensão vs. Tempo")
            self.ax1.set_ylabel("Tensão (V)")
            self.ax1.grid(True)
            self.ax1.legend()
            plot_index += 1

        if corrente_disponivel:
            self.ax2 = self.fig.add_subplot(num_plots, 1, plot_index)
            self.line2, = self.ax2.plot(self.plot_timestamps, self.plot_corrente_data, marker='.', linestyle='-', color='r', label='Corrente (A)')
            self.ax2.set_title("Medições de Corrente vs. Tempo")
            self.ax2.set_ylabel("Corrente (A)")
            self.ax2.grid(True)
            self.ax2.legend()
        
        self.fig.tight_layout(pad=3.0)
        
        canvas = FigureCanvasTkAgg(self.fig, master=self.plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        toolbar = NavigationToolbar2Tk(canvas, self.plot_window)
        toolbar.update()
        
        self.animation = animation.FuncAnimation(self.fig, self._update_plot, interval=500, blit=False, cache_frame_data=False)

    def _update_plot(self, frame):
        try:
            pontos_ja_plotados = len(self.plot_timestamps)
            
            if len(self.historico_timestamps) > pontos_ja_plotados:
                novos_timestamps = self.historico_timestamps[pontos_ja_plotados:]
                novas_tensoes = self.historico_tensao[pontos_ja_plotados:]
                novas_correntes = self.historico_corrente[pontos_ja_plotados:]

                self.plot_timestamps.extend(novos_timestamps)
                self.plot_tensao_data.extend(novas_tensoes)
                self.plot_corrente_data.extend(novas_correntes)
                
                if self.ax1:
                    self.line1.set_data(self.plot_timestamps, self.plot_tensao_data)
                    self.ax1.relim()
                    self.ax1.autoscale_view()

                if self.ax2:
                    self.line2.set_data(self.plot_timestamps, self.plot_corrente_data)
                    self.ax2.relim()
                    self.ax2.autoscale_view()
            
                if self.plot_window:
                    self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"Erro ao atualizar o gráfico: {e}")

    def _on_plot_close(self):
        if self.animation:
            self.animation.event_source.stop()
            self.animation = None
        if self.plot_window:
            self.plot_window.destroy()
            self.plot_window = None

    def iniciar_sequencia(self):
        # Limpa o histórico antes de cada nova sequência
        self.historico_timestamps.clear()
        self.historico_tensao.clear()
        self.historico_corrente.clear()

        # Reseta o estao e texto do botão do gráficdo
        self.btn_abrir_grafico.config(state=tk.DISABLED, text="Abrir Gráfico em Tempo Real")
        
        # Fecha a janela do gráfico anterior, se estiver aberta
        if self.plot_window:
            self._on_plot_close()

        threading.Thread(target=self.executar_sequencia, daemon=True).start()

    def executar_sequencia(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_conectar.config(state=tk.DISABLED)
        if self.selections['multimetro'] and self.plot_var.get():
            self.btn_abrir_grafico.config(state=tk.NORMAL)
        
        csv_file = None
        csv_writer = None
        csv_filename = ""
        
        try:
            intervalo_medicao = 1.0
            medir_tensao = False
            medir_corrente = False

            if self.selections['multimetro']:
                intervalo_medicao = float(self.entries['multimetro_intervalo'].get()) 
                
                base_name = self.entries['multimetro_csv_name'].get().strip()
                if not base_name:
                    messagebox.showwarning("Nome Inválido", "O nome do arquivo CSV não pode estar em branco.")
                    self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                    return
                
                documentos_dir = os.path.join(os.path.expanduser('~'), 'Documents')
                os.makedirs(documentos_dir, exist_ok=True)
                csv_filename = os.path.join(documentos_dir, f"{base_name}.csv")

                medir_tensao = self.volt_meas_var.get()
                medir_corrente = self.curr_meas_var.get()

                if not medir_tensao and not medir_corrente:
                    messagebox.showwarning("Configuração Inválida", "Nenhum modo de medição selecionado.")
                    self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                    return

            if self.selections['fonte']:
                for etapa_widgets in self.etapas['fonte']:
                    if etapa_widgets['volt_check_var'].get() and not medir_tensao:
                        messagebox.showwarning("Configuração Inválida", "Gatilho de Tensão exige medição de Tensão.")
                        self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                        return
                    if etapa_widgets['curr_check_var'].get() and not medir_corrente:
                        messagebox.showwarning("Configuração Inválida", "Gatilho de Corrente exige medição de Corrente.")
                        self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                        return
            
            if self.selections['multimetro']:
                csv_header = ["Timestamp", "Etapa", "Tensao_Fonte", "Corrente_Fonte", "Modo_Carga", "Valor_Carga"]
                if medir_tensao: csv_header.append("Tensao_Multimetro")
                if medir_corrente: csv_header.append("Corrente_Multimetro")
                
                csv_file = open(csv_filename, 'w', newline='', encoding='utf-8')
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(csv_header)
                csv_file.flush()
                os.fsync(csv_file.fileno())
                
            num_etapas_fonte = len(self.etapas['fonte']) if self.selections['fonte'] else 0
            num_etapas_carga = len(self.etapas['carga']) if self.selections['carga'] else 0
            num_etapas_geral = max(num_etapas_fonte, num_etapas_carga)

            if num_etapas_geral == 0:
                messagebox.showinfo("Aviso", "Nenhuma etapa configurada.")
                self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                return

            self.current_volt_range = 100 
            
            for i in range(num_etapas_geral):
                etapa_num = i + 1
                fonte = self.instruments.get('fonte')
                carga = self.instruments.get('carga')
                multimetro = self.instruments.get('multimetro')
                
                self.label_status_geral.config(text=f"Configurando Etapa {etapa_num}...")
                
                v_set, i_set, modo_carga_set, valor_carga_set = "N/A", "N/A", "N/A", "N/A"

                if i < num_etapas_fonte:
                    etapa_widgets_fonte = self.etapas['fonte'][i]
                    v_set = float(etapa_widgets_fonte['v_entry'].get())
                    i_set = float(etapa_widgets_fonte['i_entry'].get())
                    if fonte:
                        fonte.write(f"VOLT {v_set}")
                        fonte.write(f"CURR {i_set}")
                
                if i < num_etapas_carga:
                    _, var_modo, entry_val = self.etapas['carga'][i]
                    modo_carga_set = var_modo.get()
                    valor_carga_set = entry_val.get()
                    if carga:
                        sigla = modo_carga_set.split('(')[1].replace(')', '')
                        valor_carga_num = float(valor_carga_set) if valor_carga_set else 0
                        cmd_map = {"CC": ("CURR", f"CURR {valor_carga_num}"), "CV": ("VOLT", f"VOLT {valor_carga_num}"), 
                                    "CR": ("RES", f"RES {valor_carga_num}"), "CP": ("POW", f"POW {valor_carga_num}")}
                        if sigla in cmd_map:
                            carga.write(f"FUNC {cmd_map[sigla][0]}")
                            carga.write(cmd_map[sigla][1])
                
                if fonte: fonte.write("OUTP ON")
                time.sleep(0.2) 
                if carga: carga.write("INPUT ON")
                
                self.label_status_geral.config(text=f"Estabilizando Etapa {etapa_num}")

                start_time = time.time()
                last_v, last_i = "N/A", "N/A"

                while True:
                    current_elapsed_time = time.time() - start_time
                    loop_start_time = time.time()
                    
                    stop_condition_met = False
                    
                    if self.selections['fonte'] and i < num_etapas_fonte:
                        etapa_widgets_fonte = self.etapas['fonte'][i]
                        try:
                            if etapa_widgets_fonte['time_check_var'].get():
                                duracao_fonte = float(etapa_widgets_fonte['time_entry'].get())
                                if current_elapsed_time >= duracao_fonte:
                                    stop_condition_met = True
                        except (ValueError, TypeError): pass
            
                    if stop_condition_met: break

                    tensao_multi_str, corrente_multi_str = "N/A", "N/A"
                    valor_numerico_v, valor_numerico_i = None, None

                    if multimetro and self.selections['multimetro']:
                        try:
                            if medir_tensao:
                                multimetro.write(f"CONF:VOLT:DC {self.current_volt_range}")
                                multimetro.write("INIT")
                                valor_bruto_v = multimetro.query("FETCH?").strip().split(',')[0]
                                valor_numerico_v = float(valor_bruto_v)
                                tensao_multi_str = f"{valor_numerico_v:.5f}"
                            
                                if valor_numerico_v < 10.0:
                                    self.current_volt_range = 10 
                                else:
                                    self.current_volt_range = 100
                                
                            if medir_corrente:
                                multimetro.write("CONF:CURR:DC 10")
                                multimetro.write("INIT")
                                valor_bruto_i = multimetro.query("FETCH?").strip().split(',')[0]
                                valor_numerico_i = float(valor_bruto_i)
                                corrente_multi_str = f"{valor_numerico_i:.5f}"
                        except Exception as e:
                            print(f"Erro de comunicação com o multímetro: {e}")
                            if medir_tensao: tensao_multi_str = "ERRO"
                            if medir_corrente: corrente_multi_str = "ERRO"
                    
                    if self.plot_var.get() and self.selections['multimetro']:
                        self.historico_timestamps.append(datetime.datetime.now())
                        
                        if medir_tensao:
                            self.historico_tensao.append(valor_numerico_v if valor_numerico_v is not None else np.nan)
                        if medir_corrente:
                            self.historico_corrente.append(valor_numerico_i if valor_numerico_i is not None else np.nan)

                    if tensao_multi_str not in ["N/A", "ERRO"]: last_v = tensao_multi_str
                    if corrente_multi_str not in ["N/A", "ERRO"]: last_i = corrente_multi_str
                    
                    ui_status = f"Etapa {etapa_num}: "
                    if medir_tensao: ui_status += f"V: {last_v}V "
                    if medir_corrente: ui_status += f"I: {last_i}A "
                    self.label_status_geral.config(text=ui_status)

                    if csv_writer:
                        log_row = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], etapa_num, v_set, i_set, modo_carga_set, valor_carga_set]
                        if medir_tensao: log_row.append(tensao_multi_str)
                        if medir_corrente: log_row.append(corrente_multi_str)
                        
                        csv_writer.writerow(log_row)
                        csv_file.flush()
                        os.fsync(csv_file.fileno())

                    if self.selections.get('fonte') and self.selections.get('multimetro') and i < num_etapas_fonte:
                        etapa_widgets_fonte = self.etapas['fonte'][i]
                        try:
                            if etapa_widgets_fonte['volt_check_var'].get() and tensao_multi_str not in ["N/A", "ERRO"]:
                                alvo, cond = float(etapa_widgets_fonte['volt_target_entry'].get()), etapa_widgets_fonte['volt_cond_var'].get()
                                if (cond == '>=' and float(tensao_multi_str) >= alvo) or (cond == '<=' and float(tensao_multi_str) <= alvo):
                                    break 
                            if etapa_widgets_fonte['curr_check_var'].get() and corrente_multi_str not in ["N/A", "ERRO"]:
                                alvo_str, unit, cond = etapa_widgets_fonte['curr_target_entry'].get(), etapa_widgets_fonte['curr_unit_var'].get(), etapa_widgets_fonte['curr_cond_var'].get()
                                alvo_numerico = float(alvo_str)
                                alvo_a = alvo_numerico / 1000.0 if unit == 'mA' else alvo_numerico
                                if (cond == '>=' and float(corrente_multi_str) >= alvo_a) or \
                                    (cond == '<=' and float(corrente_multi_str) <= alvo_a):
                                    break
                        except (ValueError, TypeError): pass

                    if self.selections.get('fonte') and i < num_etapas_fonte:
                        etapa_widgets_fonte = self.etapas['fonte'][i]
                        if not etapa_widgets_fonte['time_check_var'].get() and not etapa_widgets_fonte['volt_check_var'].get() and not etapa_widgets_fonte['curr_check_var'].get():
                            messagebox.showwarning("Loop Infinito", f"Etapa {etapa_num} da fonte não possui condição de parada. A etapa será pulada.")
                            break 
                    
                    elapsed_in_loop = time.time() - loop_start_time
                    sleep_duration = intervalo_medicao - elapsed_in_loop
                    if sleep_duration > 0:
                        time.sleep(sleep_duration)

                if fonte: fonte.write("OUTP OFF")
                if carga: carga.write("INPUT OFF")
                
                self.label_status_geral.config(text=f"Etapa {etapa_num} concluída.")
                time.sleep(0.5)

            final_message = "Sequência finalizada."
            if self.selections['multimetro'] and csv_filename:
                final_message += f" Dados salvos em sua pasta de Documentos ('{os.path.basename(csv_filename)}')."
            self.label_status_geral.config(text=f"Status: {final_message}")
    
        except Exception as e:
            messagebox.showerror("Erro na Sequência", f"Ocorreu um erro durante a execução:\n{e}")
            self.label_status_geral.config(text="Status: Erro na sequência.")
        finally:
            if csv_file:
                csv_file.close()
            
            try:
                if 'fonte' in self.instruments and self.instruments.get('fonte'): self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments and self.instruments.get('carga'): self.instruments['carga'].write("INPUT OFF")
            except Exception: pass
            
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_conectar.config(state=tk.NORMAL)
            
            if self.historico_timestamps:
                self.btn_abrir_grafico.config(state=tk.NORMAL, text="Ver Gráfico do Resultado")
            else:
                self.btn_abrir_grafico.config(state=tk.DISABLED)

    def on_close(self):
        try:
            if self.plot_window:
                self._on_plot_close()

            if 'fonte' in self.instruments and self.instruments.get('fonte'): self.instruments['fonte'].close()
            if 'carga' in self.instruments and self.instruments.get('carga'): self.instruments['carga'].close()
            if 'multimetro' in self.instruments and self.instruments.get('multimetro'): self.instruments['multimetro'].close()
        except Exception as e:
            print(f"Erro ao fechar conexões: {e}")
        finally:
            self.master.deiconify()
            self.destroy()

if __name__ == "__main__":
    app = JanelaPrincipal()
    app.mainloop()