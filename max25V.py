import pyvisa
import time
import csv
import matplotlib.pyplot as plt
import pandas as pd

class Fluke8846A:
    """Classe para controlar o multímetro Fluke 8846A."""
    def __init__(self, ip_address="172.30.248.100", port=3490):
        # Usar '@py' para o backend pyvisa-py
        self.rm = pyvisa.ResourceManager('@py')
        self.device = None
        self.ip = ip_address
        self.port = port

    def connect(self):
        try:
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip}::{self.port}::SOCKET")
            self.device.timeout = 5000
            self.device.write_termination = '\n'
            self.device.read_termination = '\n'
            print("Multímetro conectado:", self.get_id())
            return True
        except Exception as e:
            print("Erro ao conectar ao multímetro:", e)
            return False

    def get_id(self):
        return self.device.query("*IDN?").strip()

    def configure_voltage(self):
        self.device.write("CONF:VOLT:DC")
        # Ajusta a faixa para cobrir de 5 a 20V
        self.device.write("VOLT:DC:RANGE 20")
        # NPLC 1 é um bom equilíbrio entre velocidade e precisão
        self.device.write(f"VOLT:DC:NPLC 1")
        self.device.write("TRIG:SOUR IMM")
        self.device.write("SAMP:COUN 1")

    def read_voltage(self):
        self.device.write("INIT")
        return float(self.device.query("FETCH?").strip())

    def close(self):
        if self.device:
            self.device.close()
            print("Conexão com multímetro encerrada.")

# --- Funções para manipulação do arquivo CSV ---

def inicializar_csv(nome_arquivo):
    """Cria o arquivo CSV e escreve o cabeçalho."""
    with open(nome_arquivo, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Tempo_s", "Setpoint_Fonte_V", "Corrente_Fonte_A", "Tensao_Capacitor_V"])

def salvar_csv(nome_arquivo, tempo, setpoint_fonte, corrente_fonte, tensao_capacitor):
    """Adiciona uma linha de dados ao arquivo CSV."""
    with open(nome_arquivo, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([
            f"{tempo:.2f}",
            f"{setpoint_fonte:.2f}",
            f"{corrente_fonte:.6f}",
            f"{tensao_capacitor:.6f}"
        ])

def main():
    print("Iniciando experimento do circuito RC\n")

    # --- Conexão com os Equipamentos ---
    try:
        rm = pyvisa.ResourceManager()
        fonte = rm.open_resource('USB0::0x0699::0x0392::C010658::INSTR')
        print("Fonte conectada:", fonte.query("*IDN?").strip())
    except Exception as e:
        print("Erro ao conectar à fonte USB:", e)
        return

    multimetro = Fluke8846A(ip_address="172.30.248.100", port=3490)
    if not multimetro.connect():
        fonte.close()
        return
    multimetro.configure_voltage()

    # --- Configurações do Experimento ---
    nome_csv = "dados_circuito_rc.csv"
    inicializar_csv(nome_csv)
    
    DURACAO_TOTAL = 10  # segundos
    INTERVALO = 0.1    # segundos (10 medições por segundo)
    
    setpoint_atual = 0.0

    try:
        fonte.write("OUTP ON")
        fonte.write("VOLT 0")
        time.sleep(1)
        
        tempo_inicial = time.time()
        
        # --- Loop Principal do Experimento ---
        while True:
            tempo_decorrido = time.time() - tempo_inicial
            if tempo_decorrido > DURACAO_TOTAL:
                break

            # Define o setpoint da fonte com base no tempo
            if tempo_decorrido < 3:
                novo_setpoint = 10.0
            elif tempo_decorrido < 6:
                novo_setpoint = 20.0
            else:
                novo_setpoint = 5.0
            
            # Aplica a nova tensão apenas quando ela mudar
            if novo_setpoint != setpoint_atual:
                setpoint_atual = novo_setpoint
                fonte.write(f"VOLT {setpoint_atual}")
                print(f"\n--- Mudando setpoint para {setpoint_atual}V em t={tempo_decorrido:.1f}s ---\n")

            # Realiza as medições
            corrente_fonte = float(fonte.query("MEAS:CURR?"))
            tensao_capacitor = multimetro.read_voltage()

            # Salva e imprime os dados
            salvar_csv(nome_csv, tempo_decorrido, setpoint_atual, corrente_fonte, tensao_capacitor)
            print(f"t={tempo_decorrido:5.2f}s | Setpoint={setpoint_atual:4.1f}V | "
                  f"I_fonte={corrente_fonte*1000:6.3f}mA | V_cap={tensao_capacitor:6.3f}V")
            
            time.sleep(INTERVALO)

    finally:
        # --- Desligamento Seguro ---
        fonte.write("VOLT 0")
        fonte.write("OUTP OFF")
        multimetro.close()
        fonte.close()
        print("\nConexões encerradas e equipamentos desligados.")

    # --- Geração do Gráfico ---
    df = pd.read_csv(nome_csv)
    fig, ax1 = plt.subplots(figsize=(12, 6))

    ax1.set_xlabel("Tempo (s)")
    ax1.set_ylabel("Tensão (V)", color="black")
    ax1.plot(df["Tempo_s"], df["Setpoint_Fonte_V"], label="Setpoint da Fonte (V)", color="blue", linestyle="--")
    ax1.plot(df["Tempo_s"], df["Tensao_Capacitor_V"], label="Tensão no Capacitor (V)", color="red", marker=".", markersize=3, linestyle="-")
    ax1.tick_params(axis='y', labelcolor="black")
    ax1.grid(True, which='both', linestyle=':')

    # Eixo secundário para a corrente
    ax2 = ax1.twinx()
    ax2.set_ylabel("Corrente (mA)", color="green")
    ax2.plot(df["Tempo_s"], df["Corrente_Fonte_A"] * 1000, label="Corrente da Fonte (mA)", color="green", linestyle="--", alpha=0.7)
    ax2.tick_params(axis='y', labelcolor="green")

    fig.suptitle("Curva de Carga e Descarga do Circuito RC", fontsize=16)
    fig.legend(loc="upper right", bbox_to_anchor=(1,1), bbox_transform=ax1.transAxes)
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    plt.show()

if __name__ == "__main__":
    main()