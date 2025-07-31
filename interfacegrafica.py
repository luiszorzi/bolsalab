import tkinter as tk
from tkinter import messagebox, ttk
import pyvisa
import threading
import time
import csv
import datetime

# JANELA PRINCIPAL DE SELEÇÃO (Sem alterações)
class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Equipamentos")
        self.geometry("350x250")

        self.use_fonte = tk.BooleanVar()
        self.use_multimetro = tk.BooleanVar()
        self.use_carga = tk.BooleanVar()

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
        self.geometry("650x800") # Aumentei um pouco a janela
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

    # MÉTODOS DE CRIAÇÃO DE UI
    
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
        self.frames['fonte_etapas'].pack(pady=5, padx=5)
        
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

        # Intervalo e Nome do Arquivo CSV
        config_frame = tk.Frame(frame)
        config_frame.pack(pady=5, padx=5, fill='x')
        
        tk.Label(config_frame, text="Intervalo (s):").pack(side=tk.LEFT, padx=(0,5))
        entry_intervalo = tk.Entry(config_frame, width=8)
        entry_intervalo.insert(0, "0.5")
        entry_intervalo.pack(side=tk.LEFT)
        self.entries['multimetro_intervalo'] = entry_intervalo

        tk.Label(config_frame, text="Nome do Arquivo:").pack(side=tk.LEFT, padx=(10,5))
        entry_csv = tk.Entry(config_frame, width=25)
        entry_csv.insert(0, "medicoes.csv")
        entry_csv.pack(side=tk.LEFT, expand=True)
        self.entries['multimetro_csv'] = entry_csv
        
        # Frame para mostrar a leitura
        leitura_frame = tk.Frame(frame)
        leitura_frame.pack(pady=5, padx=5, fill='x')
        label_leitura = tk.Label(leitura_frame, text="Última Leitura: -- V", font=("Arial", 11, "italic"))
        label_leitura.pack()
        self.labels['multimetro_leitura'] = label_leitura
    
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

    # MÉTODOS PARA ADICIONAR/REMOVER ETAPAS 
    def adicionar_etapa_fonte(self):    
        frame = tk.Frame(self.frames['fonte_etapas'])
        frame.pack(pady=2)
        tk.Label(frame, text=f"Etapa {len(self.etapas['fonte']) + 1}:").pack(side=tk.LEFT, padx=5)
        
        tk.Label(frame, text="Tensão (V):").pack(side=tk.LEFT)
        entry_v = tk.Entry(frame, width=7)
        entry_v.insert(0, "12.0")
        entry_v.pack(side=tk.LEFT)
        
        tk.Label(frame, text="Corrente (A):").pack(side=tk.LEFT, padx=(5,0))
        entry_i = tk.Entry(frame, width=7)
        entry_i.insert(0, "1.0")
        entry_i.pack(side=tk.LEFT)

        tk.Label(frame, text="Duração (s):").pack(side=tk.LEFT, padx=(5,0))
        entry_t = tk.Entry(frame, width=5)
        entry_t.insert(0, "5")
        entry_t.pack(side=tk.LEFT)
        
        self.etapas['fonte'].append((frame, entry_v, entry_i, entry_t))

    def remover_etapa_fonte(self):
        if not self.etapas['fonte']:
            messagebox.showwarning("Aviso", "Nenhuma etapa da fonte para remover.")
            return
        
        frame_a_remover, _, _, _ = self.etapas['fonte'].pop()
        frame_a_remover.destroy()

        for i, (frame, _, _, _) in enumerate(self.etapas['fonte']):
            label_etapa = frame.winfo_children()[0]
            label_etapa.config(text=f"Etapa {i + 1}:")

    def adicionar_etapa_carga(self):
        frame = tk.Frame(self.frames['carga_etapas'])
        frame.pack(pady=2)
        
        tk.Label(frame, text=f"Etapa {len(self.etapas['carga']) + 1}:").pack(side=tk.LEFT, padx=5)

        var_modo = tk.StringVar(value="Corrente Constante (CC)")
        modos = ["Corrente Constante (CC)", "Tensão Constante (CV)", "Potência Constante (CP)", "Resistência Constante (CR)"]
        modo_menu = ttk.Combobox(frame, textvariable=var_modo, values=modos, width=25, state="readonly")
        modo_menu.pack(side=tk.LEFT, padx=5)

        entry_val = tk.Entry(frame, width=8)
        entry_val.pack(side=tk.LEFT, padx=5)

        tk.Label(frame, text="Duração (s):").pack(side=tk.LEFT)
        entry_tempo = tk.Entry(frame, width=5)
        entry_tempo.insert(0, "5")
        entry_tempo.pack(side=tk.LEFT, padx=5)

        self.etapas['carga'].append((frame, var_modo, entry_val, entry_tempo))

    def remover_etapa_carga(self):
        if not self.etapas['carga']:
            messagebox.showwarning("Aviso", "Nenhuma etapa da carga para remover.")
            return

        frame_a_remover, _, _, _ = self.etapas['carga'].pop()
        frame_a_remover.destroy()

        for i, (frame, _, _, _) in enumerate(self.etapas['carga']):
            label_etapa = frame.winfo_children()[0]
            label_etapa.config(text=f"Etapa {i + 1}:")

    # MÉTODOS DE CONTROLE E COMUNICAÇÃO VISA
    def buscar_enderecos(self, entry_widget):
        try:
            recursos = self.rm.list_resources()
            if not recursos:
                messagebox.showinfo("Nenhum dispositivo", "Nenhum dispositivo VISA encontrado.")
                return

            janela_lista = tk.Toplevel(self)
            janela_lista.title("Dispositivos encontrados")
            janela_lista.geometry("420x300")
            janela_lista.grab_set() # Garante que esta janela fique em foco
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
                addr = self.entries['fonte_addr'].get()
                inst = self.rm.open_resource(addr)
                inst.timeout = 5000
                self.instruments['fonte'] = inst
                log.append(f"Fonte Conectada: {inst.query('*IDN?').strip()}")
            
            if self.selections['multimetro']:
                # Agora usa pyvisa para o multímetro
                addr = self.entries['multimetro_addr'].get()
                inst = self.rm.open_resource(addr)
                inst.timeout = 5000
                inst.write_termination = '\n'
                inst.read_termination = '\n'
                self.instruments['multimetro'] = inst
                log.append(f"Multímetro Conectado: {inst.query('*IDN?').strip()}")
                
            if self.selections['carga']:
                addr = self.entries['carga_addr'].get()
                inst = self.rm.open_resource(addr)
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
            intervalo_medicao = float(self.entries['multimetro_intervalo'].get())
            csv_filename = self.entries['multimetro_csv'].get()

            with open(csv_filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Etapa", "Tensao_Fonte_V", "Corrente_Fonte_A", "Tensao_Multimetro_V"])
            
            if 'multimetro' in self.instruments:
                # Configura o multímetro uma vez fora do loop
                self.instruments['multimetro'].write("CONF:VOLT:DC")
                self.instruments['multimetro'].write("VOLT:DC:NPLC 1") 

            num_etapas = 0
            if 'fonte' in self.instruments and self.etapas['fonte']:
                num_etapas = max(num_etapas, len(self.etapas['fonte']))
            if 'carga' in self.instruments and self.etapas['carga']:
                num_etapas = max(num_etapas, len(self.etapas['carga']))
            
            if num_etapas == 0:
                messagebox.showinfo("Aviso", "Nenhuma etapa configurada para fonte ou carga.")
                self.btn_iniciar.config(state=tk.NORMAL)
                self.btn_conectar.config(state=tk.NORMAL)
                return
            
            for i in range(num_etapas):
                self.label_status_geral.config(text=f"Executando Etapa {i+1} de {num_etapas}...")
                tensao_fonte_config, corrente_fonte_config, duracao_etapa = "N/A", "N/A", 0
                
                if 'fonte' in self.instruments and i < len(self.etapas['fonte']):
                    _, entry_v, entry_i, entry_t = self.etapas['fonte'][i]
                    v, i_limit, t = float(entry_v.get()), float(entry_i.get()), float(entry_t.get())
                    tensao_fonte_config, corrente_fonte_config, duracao_etapa = v, i_limit, max(duracao_etapa, t)
                    self.instruments['fonte'].write(f"VOLT {v}")
                    self.instruments['fonte'].write(f"CURR {i_limit}")
                    self.instruments['fonte'].write("OUTP ON")
                
                if duracao_etapa <= 0: continue
                
                num_leituras = int(duracao_etapa / intervalo_medicao) or 1

                for j in range(num_leituras):
                    leitura_status = f"Etapa {i+1}: Medindo... ({j+1}/{num_leituras})"
                    self.label_status_geral.config(text=leitura_status)
                    
                    valor_tensao_csv = "ERRO" 

                    if 'multimetro' in self.instruments:
                        try:
                            self.instruments['multimetro'].write("INIT")
                            time.sleep(0.1) 
                            valor_lido_str = self.instruments['multimetro'].query("FETCH?").strip()
                            valor_lido_float = float(valor_lido_str.split(',')[0]) # Pega o primeiro valor caso retorne múltiplos
                            self.labels['multimetro_leitura'].config(text=f"Última Leitura: {valor_lido_float:.4f} V")
                            valor_tensao_csv = f"{valor_lido_float:.4f}"

                        except Exception as e:
                            self.labels['multimetro_leitura'].config(text=f"Erro na leitura: {e}")
                            valor_tensao_csv = "ERRO_LEITURA"
                    
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    with open(csv_filename, 'a', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        writer.writerow([timestamp, i + 1, tensao_fonte_config, corrente_fonte_config, valor_tensao_csv])
                    
                    tempo_restante_intervalo = intervalo_medicao - 0.1
                    if tempo_restante_intervalo > 0:
                        time.sleep(tempo_restante_intervalo)

                if 'fonte' in self.instruments: self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments: self.instruments['carga'].write("INPUT OFF")
                time.sleep(0.5)

            self.label_status_geral.config(text=f"Status: Sequência finalizada. Dados salvos em '{csv_filename}'.")
            
        except Exception as e:
            messagebox.showerror("Erro na Sequência", f"Ocorreu um erro durante a execução:\n{e}")
            self.label_status_geral.config(text="Status: Erro na sequência.")
        finally:
            try:
                if 'fonte' in self.instruments and self.instruments.get('fonte').session: self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments and self.instruments.get('carga').session: self.instruments['carga'].write("INPUT OFF")
            except Exception as e:
                print(f"Aviso: não foi possível garantir o desligamento das saídas. Erro: {e}")
            
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_conectar.config(state=tk.NORMAL)

    def on_close(self):
        try:
            if 'fonte' in self.instruments and self.instruments['fonte']:
                self.instruments['fonte'].write("OUTP OFF")
                self.instruments['fonte'].close()
            if 'carga' in self.instruments and self.instruments['carga']:
                self.instruments['carga'].write("INPUT OFF")
                self.instruments['carga'].close()
            if 'multimetro' in self.instruments and self.instruments['multimetro']:
                self.instruments['multimetro'].close()
        except Exception as e:
            print(f"Erro ao desligar/desconectar equipamentos ao fechar: {e}")
        finally:
            self.master.deiconify()
            self.destroy()

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()