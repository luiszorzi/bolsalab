import tkinter as tk
from tkinter import messagebox, ttk
import pyvisa
import threading
import time
import csv
import datetime

# JANELA PRINCIPAL DE SELEÇÃO
class MainWindow(tk.Tk):
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
        control_window = CombinedControlWindow(self, selections)
        control_window.grab_set()

# JANELA DE CONTROLE COMBINADO
class CombinedControlWindow(tk.Toplevel):
    def __init__(self, master, selections):
        super().__init__(master)
        self.title("Controle de Equipamentos")
        self.geometry("800x850")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.selections = selections
        self.rm = pyvisa.ResourceManager()
        self.instruments = {}
        self.entries = {}
        self.etapas = {'fonte': [], 'carga': []}
        self.labels = {}
        self.frames = {}

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
        
        self.btn_conectar = tk.Button(control_frame, text="Conectar Equipamentos", command=self.conectar_todos)
        self.btn_conectar.pack(pady=5)

        self.btn_iniciar = tk.Button(control_frame, text="Iniciar Sequência", command=self.iniciar_sequencia, state=tk.DISABLED)
        self.btn_iniciar.pack(pady=5)

        self.label_status_geral = tk.Label(control_frame, text="Status: Aguardando conexão...", font=("Arial", 12))
        self.label_status_geral.pack(pady=10)

        self.btn_voltar = tk.Button(control_frame, text="← Voltar", command=self.on_close)
        self.btn_voltar.pack(side=tk.BOTTOM, pady=10)

    def _create_fonte_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Fonte de Alimentação (PWS4305)")
        frame.pack(pady=10, padx=5, fill='x')
        self.frames['fonte'] = frame

        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:").pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40)
        entry_addr.insert(0, "USB0::0x0699::0x0392::C010658::INSTR")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr)).pack(side=tk.LEFT)
        self.entries['fonte_addr'] = entry_addr

        self.frames['fonte_etapas'] = tk.Frame(frame)
        self.frames['fonte_etapas'].pack(pady=5, padx=5, fill='x')

        botoes_etapa_frame = tk.Frame(frame)
        botoes_etapa_frame.pack(pady=5)
        tk.Button(botoes_etapa_frame, text="+ Adicionar Etapa", command=self.adicionar_etapa_fonte).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_etapa_frame, text="- Remover Etapa", command=self.remover_etapa_fonte).pack(side=tk.LEFT, padx=5)
        
        self.adicionar_etapa_fonte()

    def _create_multimetro_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Multímetro")
        frame.pack(pady=10, padx=5, fill='x')
        
        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:").pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40)
        entry_addr.insert(0, "TCPIP0::172.30.248.100::3490::SOCKET")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr)).pack(side=tk.LEFT)
        self.entries['multimetro_addr'] = entry_addr
        
        config_frame = tk.Frame(frame)
        config_frame.pack(pady=5, padx=5, fill='x')
        
        self.volt_meas_var = tk.BooleanVar(value=True)
        self.curr_meas_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(config_frame, text="Medir Tensão", variable=self.volt_meas_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(config_frame, text="Medir Corrente", variable=self.curr_meas_var).pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(config_frame, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10, pady=5)
        
        tk.Label(config_frame, text="Intervalo (s):").pack(side=tk.LEFT, padx=(0,5))
        entry_intervalo = tk.Entry(config_frame, width=8)
        entry_intervalo.insert(0, "1.0")
        entry_intervalo.pack(side=tk.LEFT)
        self.entries['multimetro_intervalo'] = entry_intervalo
        
        tk.Label(config_frame, text="Nome do Arquivo:").pack(side=tk.LEFT, padx=(10,5))
        entry_csv = tk.Entry(config_frame, width=25)
        entry_csv.insert(0, "medicoes.csv")
        entry_csv.pack(side=tk.LEFT, expand=True)
        self.entries['multimetro_csv'] = entry_csv

    def _create_carga_ui(self, parent):
        frame = ttk.LabelFrame(parent, text="Carga Eletrônica")
        frame.pack(pady=10, padx=5, fill='x')
        self.frames['carga'] = frame
        
        addr_frame = tk.Frame(frame)
        addr_frame.pack(pady=5, fill='x', padx=5)
        tk.Label(addr_frame, text="Endereço VISA:").pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40)
        entry_addr.insert(0, "USB0::0x05E6::0x2380::802436052757810021::INSTR")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        tk.Button(addr_frame, text="Buscar", command=lambda: self.buscar_enderecos(entry_addr)).pack(side=tk.LEFT)
        self.entries['carga_addr'] = entry_addr
        
        self.frames['carga_etapas'] = tk.Frame(frame)
        self.frames['carga_etapas'].pack(pady=5, padx=5)
        
        botoes_etapa_frame = tk.Frame(frame)
        botoes_etapa_frame.pack(pady=5)
        tk.Button(botoes_etapa_frame, text="+ Adicionar Etapa", command=self.adicionar_etapa_carga).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_etapa_frame, text="- Remover Etapa", command=self.remover_etapa_carga).pack(side=tk.LEFT, padx=5)
        
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
        tk.Label(config_frame, text="Tensão (V):").pack(side=tk.LEFT)
        entry_v = tk.Entry(config_frame, width=7)
        entry_v.insert(0, "10.0")
        entry_v.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(config_frame, text="Corrente (A):").pack(side=tk.LEFT)
        entry_i = tk.Entry(config_frame, width=7)
        entry_i.insert(0, "1.0")
        entry_i.pack(side=tk.LEFT)
        
        trigger_frame = tk.Frame(etapa_frame)
        trigger_frame.pack(pady=5, fill='x', padx=5)
        
        time_check_var = tk.BooleanVar(value=True)
        time_check = tk.Checkbutton(trigger_frame, text="Duração (s):", variable=time_check_var)
        time_check.pack(side=tk.LEFT)
        entry_t = tk.Entry(trigger_frame, width=7)
        entry_t.insert(0, "5")
        entry_t.pack(side=tk.LEFT)
        
        volt_check_var = tk.BooleanVar(value=False)
        volt_check = tk.Checkbutton(trigger_frame, text="Tensão:", variable=volt_check_var)
        volt_check.pack(side=tk.LEFT, padx=(15, 0))
        volt_cond_var = tk.StringVar(value='<=')
        volt_cond_menu = ttk.Combobox(trigger_frame, textvariable=volt_cond_var, values=['>=', '<='], width=3, state=tk.DISABLED)
        volt_cond_menu.pack(side=tk.LEFT)
        entry_vt = tk.Entry(trigger_frame, width=7, state=tk.DISABLED)
        entry_vt.insert(0, "9.5")
        entry_vt.pack(side=tk.LEFT)
        
        curr_check_var = tk.BooleanVar(value=False)
        curr_check = tk.Checkbutton(trigger_frame, text="Corrente:", variable=curr_check_var)
        curr_check.pack(side=tk.LEFT, padx=(15, 0))
        curr_cond_var = tk.StringVar(value='>=')
        curr_cond_menu = ttk.Combobox(trigger_frame, textvariable=curr_cond_var, values=['>=', '<='], width=3, state=tk.DISABLED)
        curr_cond_menu.pack(side=tk.LEFT)
        entry_ct = tk.Entry(trigger_frame, width=7, state=tk.DISABLED)
        entry_ct.insert(0, "0.1")
        entry_ct.pack(side=tk.LEFT)
        
        time_check.config(command=lambda: self._toggle_entry_state(time_check_var, entry_t))
        volt_check.config(command=lambda: self._toggle_entry_state(volt_check_var, volt_cond_menu, entry_vt))
        curr_check.config(command=lambda: self._toggle_entry_state(curr_check_var, curr_cond_menu, entry_ct))
        
        widgets = {
            'frame': etapa_frame, 'v_entry': entry_v, 'i_entry': entry_i,
            'time_check_var': time_check_var, 'time_entry': entry_t,
            'volt_check_var': volt_check_var, 'volt_cond_var': volt_cond_var, 'volt_target_entry': entry_vt,
            'curr_check_var': curr_check_var, 'curr_cond_var': curr_cond_var, 'curr_target_entry': entry_ct,
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
        frame.pack(pady=2)
        
        tk.Label(frame, text=f"Etapa {len(self.etapas['carga']) + 1}:").pack(side=tk.LEFT, padx=5)
        var_modo = tk.StringVar(value="Resistência Constante (CR)")
        modos = ["Corrente Constante (CC)", "Tensão Constante (CV)", "Potência Constante (CP)", "Resistência Constante (CR)"]
        modo_menu = ttk.Combobox(frame, textvariable=var_modo, values=modos, width=25, state="readonly")
        modo_menu.pack(side=tk.LEFT, padx=5)
        
        entry_val = tk.Entry(frame, width=8)
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
            lista = tk.Listbox(janela_lista, width=60)
            lista.pack(pady=5, padx=5, fill='both', expand=True)
            
            for r in recursos:
                lista.insert(tk.END, r)
                
            def selecionar():
                selecionado = lista.get(tk.ACTIVE)
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, selecionado)
                janela_lista.destroy()
                
            tk.Button(janela_lista, text="Selecionar", command=selecionar).pack(pady=5)
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
        
        try:
            # SETUP INICIAL
            intervalo_medicao = float(self.entries['multimetro_intervalo'].get())
            csv_filename = self.entries['multimetro_csv'].get()
            medir_tensao = self.volt_meas_var.get()
            medir_corrente = self.curr_meas_var.get()

            # VALIDAÇÕES
            if self.selections['multimetro'] and not medir_tensao and not medir_corrente:
                messagebox.showwarning("Configuração Inválida", "Nenhum modo de medição (Tensão/Corrente) selecionado para o multímetro.")
                self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                return

            for etapa_widgets in self.etapas['fonte']:
                if etapa_widgets['volt_check_var'].get() and not medir_tensao:
                    messagebox.showwarning("Configuração Inválida", "Gatilho por 'Tensão' exige que 'Medir Tensão' no multímetro esteja selecionado.")
                    self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                    return
                
                if etapa_widgets['curr_check_var'].get() and not medir_corrente:
                    messagebox.showwarning("Configuração Inválida", "Gatilho por 'Corrente' exige que 'Medir Corrente' no multímetro esteja selecionado.")
                    self.btn_iniciar.config(state=tk.NORMAL); self.btn_conectar.config(state=tk.NORMAL)
                    return

            # PREPARAÇÃO DO CSV
            csv_header = ["Timestamp", "Etapa", "Tensao_Fonte_Config_V", "Corrente_Fonte_Config_A", "Modo_Carga_Config", "Valor_Carga_Config"]
            if medir_tensao:
                csv_header.append("Tensao_Multimetro_V")
            if medir_corrente:
                csv_header.append("Corrente_Multimetro_A")
            
            with open(csv_filename, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(csv_header)

            # LÓGICA PRINCIPAL
            num_etapas_fonte = len(self.etapas['fonte'])
            num_etapas_carga = len(self.etapas['carga'])
            num_etapas_geral = max(num_etapas_fonte, num_etapas_carga)

            v_set, i_set, modo_carga_set, valor_carga_set = "N/A", "N/A", "N/A", "N/A"

            for i in range(num_etapas_geral):
                etapa_num = i + 1
                fonte = self.instruments.get('fonte')
                carga = self.instruments.get('carga')
                multimetro = self.instruments.get('multimetro')
                
                self.label_status_geral.config(text=f"Configurando Etapa {etapa_num}...")
                
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
                
                if fonte:
                    fonte.write("OUTP ON")
                if carga:
                    carga.write("INPUT ON")
                
                start_time = time.time()
                last_v, last_i = "N/A", "N/A"

                while True:
                    current_elapsed_time = time.time() - start_time
                    loop_start_time = time.time()
                    
                    stop_condition_met = False
                    
                    if not stop_condition_met and i < num_etapas_fonte:
                        etapa_widgets_fonte = self.etapas['fonte'][i]
                        try:
                            if etapa_widgets_fonte['time_check_var'].get():
                                duracao_fonte = float(etapa_widgets_fonte['time_entry'].get())
                                if current_elapsed_time >= duracao_fonte:
                                    stop_condition_met = True
                        except (ValueError, TypeError):
                            pass
                    
                    if stop_condition_met:
                        break

                    tensao_multi_str, corrente_multi_str = "N/A", "N/A"
                    if multimetro:
                        try:
                            if medir_tensao:
                                multimetro.write("CONF:VOLT:DC")
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

                    log_row = [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3], etapa_num, v_set, i_set, modo_carga_set, valor_carga_set]
                    if medir_tensao: log_row.append(tensao_multi_str)
                    if medir_corrente: log_row.append(corrente_multi_str)
                    with open(csv_filename, 'a', newline='', encoding='utf-8') as f:
                        csv.writer(f).writerow(log_row)

                    if i < num_etapas_fonte:
                        etapa_widgets_fonte = self.etapas['fonte'][i]
                        try:
                            if etapa_widgets_fonte['volt_check_var'].get() and tensao_multi_str not in ["N/A", "ERRO"]:
                                alvo, cond = float(etapa_widgets_fonte['volt_target_entry'].get()), etapa_widgets_fonte['volt_cond_var'].get()
                                if (cond == '>=' and float(tensao_multi_str) >= alvo) or (cond == '<=' and float(tensao_multi_str) <= alvo):
                                    break 
                            if etapa_widgets_fonte['curr_check_var'].get() and corrente_multi_str not in ["N/A", "ERRO"]:
                                alvo, cond = float(etapa_widgets_fonte['curr_target_entry'].get()), etapa_widgets_fonte['curr_cond_var'].get()
                                if (cond == '>=' and float(corrente_multi_str) >= alvo) or (cond == '<=' and float(corrente_multi_str) <= alvo):
                                    break
                        except (ValueError, TypeError):
                            pass
                    
                    elapsed_in_loop = time.time() - loop_start_time
                    sleep_duration = intervalo_medicao - elapsed_in_loop
                    if sleep_duration > 0:
                        time.sleep(sleep_duration)

                if fonte: fonte.write("OUTP OFF")
                if carga: carga.write("INPUT OFF")
                
                self.label_status_geral.config(text=f"Etapa {etapa_num} concluída.")
                time.sleep(0.5)

            self.label_status_geral.config(text=f"Status: Sequência finalizada. Dados salvos em '{csv_filename}'.")
        
        except Exception as e:
            messagebox.showerror("Erro na Sequência", f"Ocorreu um erro durante a execução:\n{e}")
            self.label_status_geral.config(text="Status: Erro na sequência.")
        finally:
            try:
                if 'fonte' in self.instruments and self.instruments.get('fonte'): self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments and self.instruments.get('carga'): self.instruments['carga'].write("INPUT OFF")
            except Exception:
                pass
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_conectar.config(state=tk.NORMAL)

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
    app = MainWindow()
    app.mainloop()