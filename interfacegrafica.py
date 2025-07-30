import tkinter as tk
from tkinter import messagebox, ttk
import pyvisa
import threading
import time
import socket
import csv


# Classe para comunicação TCP socket com o multímetro Fluke 
class FlukeSocket:
    def __init__(self, ip, port=3490, timeout=5):
        self.ip = ip
        self.port = port
        self.timeout = timeout
        self.sock = None
    
    def connect(self):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.settimeout(self.timeout)
        self.sock.connect((self.ip, self.port))
    
    def disconnect(self):
        if self.sock:
            self.sock.close()
            self.sock = None
    
    def send_command(self, cmd):
        if not self.sock:
            raise RuntimeError("Socket não conectado")
        self.sock.sendall(cmd.encode() + b'\n')
    
    def receive_response(self, buffer_size=1024):
        if not self.sock:
            raise RuntimeError("Socket não conectado")
        resp = self.sock.recv(buffer_size)
        return resp.decode().strip()
    
    def query(self, cmd):
        self.send_command(cmd)
        time.sleep(0.2)  
        return self.receive_response()



# JANELA PRINCIPAL DE SELEÇÃO

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


# JANELA DE CONTROLE COMBINADO (para controlar todos os equipamentos selecionados)

class CombinedControlWindow(tk.Toplevel):
    
    def __init__(self, master, selections):
        super().__init__(master)
        self.title("Controle de Equipamentos")
        self.geometry("600x750")
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
        entry_addr.insert(0, "USB0::0x0699::0x0408::C000012::INSTR")
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
        tk.Label(addr_frame, text="Endereço IP:").pack(side=tk.LEFT)
        entry_addr = tk.Entry(addr_frame, width=40)
        entry_addr.insert(0, "172.30.248.100")
        entry_addr.pack(side=tk.LEFT, padx=5, expand=True)
        self.entries['multimetro_ip'] = entry_addr

        tk.Label(frame, text="Intervalo (s):").pack(side=tk.LEFT, padx=(10, 0))
        entry_intervalo = tk.Entry(frame, width=5)
        entry_intervalo.insert(0, "1.0")
        entry_intervalo.pack(side=tk.LEFT, padx=(0, 10))
        self.entries['multimetro_intervalo'] = entry_intervalo

        tk.Label(frame, text="Arquivo CSV:").pack(side=tk.LEFT)
        entry_csv = tk.Entry(frame, width=20)
        entry_csv.insert(0, "medicoes.csv")
        entry_csv.pack(side=tk.LEFT, padx=5)
        self.entries['multimetro_csv'] = entry_csv


        # LINHAS ADICIONADAS PARA CORRIGIR O ERRO
        leitura_frame = tk.Frame(frame)
        leitura_frame.pack(pady=5, padx=5, fill='x')

        # Cria o label que mostrará a tensão
        label_leitura = tk.Label(leitura_frame, text="Última Leitura: -- V", font=("Arial", 11, "italic"))
        label_leitura.pack()

        # Guarda a referência do label no dicionário para ser acessado depois
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

    # MÉTODOS PARA ADICIONAR/REMOVER ETAPAS DINAMICAMENTE

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

        tk.Label(frame, text="Tempo (s):").pack(side=tk.LEFT, padx=(5,0))
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

        tk.Label(frame, text="Tempo (s):").pack(side=tk.LEFT)
        entry_tempo = tk.Entry(frame, width=5)
        entry_tempo.insert(0, "2")
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
            tk.Label(janela_lista, text="Selecione um endereço:", font=("Arial", 12)).pack(pady=5)
            lista = tk.Listbox(janela_lista, width=60)
            lista.pack(pady=5)
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
                inst.timeout = 5000 # Timeout de 5 segundos, mais razoável
                self.instruments['fonte'] = inst
                log.append(f"Fonte Conectada: {inst.query('*IDN?')}")
            
            if self.selections['multimetro']:
                ip = self.entries['multimetro_ip'].get()
                fluke = FlukeSocket(ip, timeout=15)
                fluke.connect()
                self.instruments['multimetro'] = fluke
                try:
                    idn = fluke.query('*IDN?')
                    log.append(f"Multímetro Conectado: {idn}")
                except Exception:
                    log.append("Multímetro conectado")

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
            num_etapas = 0
            if 'fonte' in self.instruments:
                num_etapas = max(num_etapas, len(self.etapas['fonte']))
            if 'carga' in self.instruments:
                num_etapas = max(num_etapas, len(self.etapas['carga']))
            
            if num_etapas == 0:
                # Se só tiver multímetro, faz uma leitura simples
                if 'multimetro' in self.instruments:
                    try:
                        valor = self.instruments['multimetro'].query("MEAS:VOLT:DC?")
                        valor = valor.strip()
                        if not valor:
                            raise ValueError("Resposta vazia do multímetro")
                        self.labels['multimetro_leitura'].config(text=f"Tensão medida: {valor} V")
                        self.label_status_geral.config(text="Status: Leitura do multímetro concluída.")
                    except Exception as e:
                        self.labels['multimetro_leitura'].config(text=f"Erro na leitura: {e}")
                        print(f"Erro na leitura do multímetro: {e}")
                return

            for i in range(num_etapas):
                self.label_status_geral.config(text=f"Executando Etapa {i+1} de {num_etapas}...")

                tempo_de_espera = 0
                
                # LÓGICA DA FONTE 
                if 'fonte' in self.instruments and i < len(self.etapas['fonte']):
                    _, entry_v, entry_i, entry_t = self.etapas['fonte'][i]
                    v = float(entry_v.get())
                    i_limit = float(entry_i.get())
                    tempo_de_espera = max(tempo_de_espera, float(entry_t.get()))
                    self.instruments['fonte'].write(f"VOLT {v}")
                    self.instruments['fonte'].write(f"CURR {i_limit}")
                    self.instruments['fonte'].write("OUTP ON")
                
                # LÓGICA DA CARGA 
                if 'carga' in self.instruments and i < len(self.etapas['carga']):
                    _, var_modo, entry_val, entry_tempo_carga = self.etapas['carga'][i]
                    tempo_de_espera = max(tempo_de_espera, float(entry_tempo_carga.get()))
                    modo_completo = var_modo.get()
                    sigla = modo_completo.split('(')[1].replace(')', '')
                    try:
                        valor = float(entry_val.get())
                    except ValueError:
                        valor = 0
                    cmd_map = {
                        "CC": ("FUNC CURR", f"CURR {valor}"),
                        "CV": ("FUNC VOLT", f"VOLT {valor}"),
                        "CP": ("FUNC POW", f"POW {valor}"),
                        "CR": ("FUNC RES", f"RES {valor}")
                    }
                    if sigla in cmd_map:
                        self.instruments['carga'].write("*RST")
                        self.instruments['carga'].write(cmd_map[sigla][0])
                        self.instruments['carga'].write(cmd_map[sigla][1])
                        self.instruments['carga'].write("INPUT ON")

                # INÍCIO DO CONTADOR 
                inicio = time.time()
                tempo_passado = 0

                # Pegando intervalo e nome do arquivo CSV
                leitura_intervalo_str = self.entries.get('multimetro_intervalo', None)
                csv_filename_str = self.entries.get('multimetro_csv', None)

                # Garantindo que os valores estão definidos e convertidos corretamente
                try:
                    leitura_intervalo = float(leitura_intervalo_str.get()) if leitura_intervalo_str else 1.0
                except Exception:
                    leitura_intervalo = 1.0
                
                try:
                    csv_filename = csv_filename_str.get() if csv_filename_str else "medicoes.csv"
                except Exception:
                    csv_filename = "medicoes.csv"

                # Criação do CSV e escrita das leituras do multímetro durante a etapa
                import csv  # caso não esteja importado no início do arquivo
                
                try:
                    # Se o arquivo não existir, escreve cabeçalho, senão adiciona
                    arquivo_existe = False
                    import os
                    if os.path.isfile(csv_filename):
                        arquivo_existe = True

                    with open(csv_filename, mode='a', newline='') as file:
                        writer = csv.writer(file)
                        if not arquivo_existe:
                            writer.writerow(['Timestamp (s)', 'Etapa', 'Tensão (V)'])
                        
                        etapa_num = i + 1
                        while tempo_passado < tempo_de_espera:
                            timestamp = round(time.time() - inicio, 3)
                            try:
                                valor = self.instruments['multimetro'].query("MEAS:VOLT:DC?")
                                valor = valor.strip()
                                if not valor:
                                    raise ValueError("Resposta vazia do multímetro")
                                self.labels['multimetro_leitura'].config(text=f"Tensão medida: {valor} V")
                                writer.writerow([timestamp, etapa_num, valor])
                            except Exception as e:
                                self.labels['multimetro_leitura'].config(text=f"Erro na leitura: {e}")
                                print(f"Erro na leitura do multímetro na etapa {etapa_num}: {e}")
                                writer.writerow([timestamp, etapa_num, "ERRO"])

                            time.sleep(leitura_intervalo)
                            tempo_passado = time.time() - inicio
                except Exception as e:
                    messagebox.showerror("Erro ao salvar CSV", f"Erro ao salvar medições no CSV:\n{e}")

                # AJUSTE PARA CUMPRIR O TEMPO EXATO DA ETAPA 
                tempo_passado = time.time() - inicio
                tempo_restante = tempo_de_espera - tempo_passado
                if tempo_restante > 0:
                    time.sleep(tempo_restante)

                # DESLIGA SAÍDAS 
                if 'fonte' in self.instruments:
                    self.instruments['fonte'].write("OUTP OFF")
                if 'carga' in self.instruments:
                    self.instruments['carga'].write("INPUT OFF")
                time.sleep(0.5)

            self.label_status_geral.config(text="Status: Sequência finalizada. Saídas desligadas.")
            
        except Exception as e:
            messagebox.showerror("Erro na Sequência", f"Ocorreu um erro durante a execução:\n{e}")
            self.label_status_geral.config(text="Status: Erro na sequência.")
        finally:
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_conectar.config(state=tk.NORMAL)


    def on_close(self):
        try:
            if 'fonte' in self.instruments and self.instruments['fonte']:
                self.instruments['fonte'].write("OUTP OFF")
            if 'carga' in self.instruments and self.instruments['carga']:
                self.instruments['carga'].write("INPUT OFF")
            if 'multimetro' in self.instruments and self.instruments['multimetro']:
                self.instruments['multimetro'].disconnect()
        except Exception as e:
            print(f"Erro ao desligar saídas na hora de fechar: {e}")
        finally:
            self.master.deiconify()
            self.destroy()

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()