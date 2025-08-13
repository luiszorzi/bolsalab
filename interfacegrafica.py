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
            
            # Filtra a lista para mostrar apenas os recursos que interessam
            recursos_filtrados = [r for r in recursos if r.startswith('USB') or r.startswith('TCPIP')]

            for r in recursos_filtrados:
                lista.insert(tk.END, r) # Adiciona apenas os recursos filtrados
                
            def selecionar():
                if lista.curselection(): # Evita erro se nada for selecionado
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

    def iniciar_sequencia(self):
        threading.Thread(target=self.executar_sequencia, daemon=True).start()

    def executar_sequencia(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_conectar.config(state=tk.DISABLED)
        
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
                
                try:
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                except NameError:
                    script_dir = os.getcwd()
                
                csv_filename = os.path.join(script_dir, f"{base_name}.csv")

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
                if medir_tensao:
                    csv_header.append("Tensao_Multimetro")
                if medir_corrente:
                    csv_header.append("Corrente_Multimetro")
                
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
                    if multimetro and self.selections['multimetro']:
                        try:
                            if medir_tensao:
                                multimetro.write("CONF:VOLT:DC 100")
                                multimetro.write("INIT")
                                valor_bruto_v = multimetro.query("FETCH?").strip().split(',')[0]
                                valor_numerico_v = float(valor_bruto_v)
                                tensao_multi_str = f"{valor_numerico_v:.6f}"
                            if medir_corrente:
                                multimetro.write("CONF:CURR:DC 10")
                                multimetro.write("INIT")
                                valor_bruto_i = multimetro.query("FETCH?").strip().split(',')[0]
                                valor_numerico_i = float(valor_bruto_i)
                                corrente_multi_str = f"{valor_numerico_i:.6f}"
                        except Exception as e:
                            print(f"Erro de comunicação com o multímetro: {e}")
                            if medir_tensao: tensao_multi_str = "ERRO"
                            if medir_corrente: corrente_multi_str = "ERRO"
                    
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

                    if self.selections['fonte'] and self.selections['multimetro'] and i < num_etapas_fonte:
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

                    if self.selections['fonte'] and i < num_etapas_fonte:
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
                final_message += f" Dados salvos em '{os.path.basename(csv_filename)}'."
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
            
            if self.selections['multimetro'] and self.plot_var.get() and csv_filename:
                self.after(0, self.plotar_graficos, csv_filename)

    def plotar_graficos(self, csv_filename):
        try:
            data = pd.read_csv(csv_filename)
            if data.empty:
                messagebox.showwarning("Gráfico", "O arquivo de dados está vazio.")
                return
        except Exception as e:
            messagebox.showerror("Gráfico", f"Erro ao ler arquivo de dados:\n{e}")
            return
            
        plot_window = tk.Toplevel(self)
        plot_window.title(f"Gráficos - {os.path.basename(csv_filename)}")
        plot_window.geometry("1000x800")
        plot_window.grab_set()

        fig = Figure(figsize=(10, 8), dpi=100)
        
        data['Timestamp'] = pd.to_datetime(data['Timestamp'])

        tensao_disponivel = 'Tensao_Multimetro' in data.columns and self.volt_meas_var.get()
        corrente_disponivel = 'Corrente_Multimetro' in data.columns and self.curr_meas_var.get()

        if not tensao_disponivel and not corrente_disponivel:
            messagebox.showinfo("Gráfico", "Não há dados de Tensão ou Corrente para plotar.")
            plot_window.destroy()
            return

        if tensao_disponivel:
            ax1 = fig.add_subplot(211 if corrente_disponivel else 111)
            data['Tensao_Multimetro'] = pd.to_numeric(data['Tensao_Multimetro'], errors='coerce')
            ax1.plot(data['Timestamp'], data['Tensao_Multimetro'], marker='.', linestyle='-', label='Tensão (V)')
            ax1.set_title("Medições de Tensão vs. Tempo")
            ax1.set_ylabel("Tensão (V)")
            ax1.grid(True)
            ax1.legend()
            
        if corrente_disponivel:
            ax2 = fig.add_subplot(212 if tensao_disponivel else 111)
            data['Corrente_Multimetro'] = pd.to_numeric(data['Corrente_Multimetro'], errors='coerce')
            ax2.plot(data['Timestamp'], data['Corrente_Multimetro'], marker='.', linestyle='-', color='r', label='Corrente (A)')
            ax2.set_title("Medições de Corrente vs. Tempo")
            ax2.set_ylabel("Corrente (A)")
            ax2.grid(True)
            ax2.legend()

        fig.tight_layout(pad=3.0) 

        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        
        toolbar = NavigationToolbar2Tk(canvas, plot_window)
        toolbar.update()
        
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def on_close(self):
        try:
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