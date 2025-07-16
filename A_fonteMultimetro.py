import pyvisa
import time
import csv
import matplotlib.pyplot as plt
import pandas as pd

# === Classe para o multímetro FLUKE 8846A ===
class Fluke8846A:
    def __init__(self, ip_address="172.30.248.100", port=3490):
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

    def configure_current(self, nplc=1):
        self.device.write("CONF:CURR:DC")
        self.device.write("CURR:DC:RANGE 10")
        self.device.write(f"CURR:DC:NPLC {nplc}")
        self.device.write("TRIG:SOUR IMM")
        self.device.write("SAMP:COUN 1")
        time.sleep(0.5)

    def read_current(self):
        self.device.write("INIT")
        return float(self.device.query("FETCH?").split(',')[0])

    def close(self):
        if self.device:
            self.device.close()
            print("Conexão com multímetro encerrada")

# === Funções auxiliares ===
def inicializar_csv(nome_arquivo):
    with open(nome_arquivo, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Corrente_Fonte_A", "Corrente_Multimetro_A"])

def salvar_csv(nome_arquivo, corrente_fonte, corrente_multimetro):
    with open(nome_arquivo, 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([f"{corrente_fonte:.6f}", f"{corrente_multimetro:.6f}"])

# === Função principal ===
def main():
    print("Comparação de Corrente - Fonte vs Multímetro\n")

    estagios = [
        {"tensao": 15.0, "tempo": 20},
        {"tensao": 30.0, "tempo": 20},
        {"tensao": 10.0, "tempo": 20}
    ]

    rm = pyvisa.ResourceManager()
    recursos = rm.list_resources()
    if not recursos:
        print("Nenhum dispositivo encontrado via VISA.")
        return

    print("Dispositivos encontrados:", recursos)
    fonte = rm.open_resource(recursos[0])
    print("Fonte conectada:", fonte.query("*IDN?").strip())

    multimetro = Fluke8846A()
    if not multimetro.connect():
        return
    multimetro.configure_current()

    nome_csv = "corrente_comparada.csv"
    inicializar_csv(nome_csv)

    try:
        for idx, estagio in enumerate(estagios, 1):
            tensao = estagio["tensao"]
            tempo_total = estagio["tempo"]

            fonte.write(f"VOLT {tensao}")
            fonte.write("OUTP ON")

            print(f"\n[{idx}] Aplicando {tensao:.1f} V por {tempo_total} segundos...")

            inicio = time.time()
            while time.time() - inicio < tempo_total:
                try:
                    corrente_fonte = float(fonte.query("MEAS:CURR?"))
                    corrente_multimetro = multimetro.read_current()
                except Exception as e:
                    print(f"Erro na leitura:", e)
                    corrente_fonte = 0.0
                    corrente_multimetro = 0.0

                print(f"Fonte: {corrente_fonte:.6f} A | Multímetro: {corrente_multimetro:.6f} A")
                salvar_csv(nome_csv, corrente_fonte, corrente_multimetro)
                time.sleep(1.0)  # <-- 1 medição por segundo

            fonte.write("OUTP OFF")
            print(f"[{idx}] Estágio concluído.")

    finally:
        multimetro.close()
        fonte.close()
        print("\nMedições finalizadas e conexões encerradas.")

        # Geração do gráfico
        df = pd.read_csv(nome_csv)
        plt.figure(figsize=(10, 6))

        plt.subplot(2, 1, 1)
        plt.plot(df.index, df["Corrente_Fonte_A"], label="Fonte", color='blue')
        plt.ylabel("Corrente (A)")
        plt.title("Corrente da Fonte")
        plt.grid(True)
        plt.xlabel("Ordem da Leitura")

        plt.subplot(2, 1, 2)
        plt.plot(df.index, df["Corrente_Multimetro_A"], label="Multímetro", color='green')
        plt.ylabel("Corrente (A)")
        plt.title("Corrente do Multímetro (10A)")
        plt.grid(True)
        plt.xlabel("Ordem da Leitura")

        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    main()
