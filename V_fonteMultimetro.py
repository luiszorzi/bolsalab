import pyvisa
import time
import csv
import matplotlib.pyplot as plt
import pandas as pd

# === Parâmetros da medição ===
DURACAO_ESTAGIO = 20  # segundos por estágio
INTERVALO_MEDICAO = 1.0  # 1 medição por segundo
ESTAGIOS = [
    {"tensao": 15.0, "corrente": 2.0},
    {"tensao": 30.0, "corrente": 2.0},
    {"tensao": 10.0, "corrente": 2.0}
]
CSV_FILE = "comparacao_tensao.csv"

# === Inicializa CSV ===
with open(CSV_FILE, mode='w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(["Tensao_Fonte_V", "Tensao_Multimetro_V"])

# === Inicializa VISA ===
rm = pyvisa.ResourceManager()

FONTE_ADDR = "USB0::0x0699::0x0392::C010658::INSTR"
MULTIMETRO_ADDR = "TCPIP0::172.30.248.100::3490::SOCKET"

try:
    fonte = rm.open_resource(FONTE_ADDR)
    print("Fonte conectada:", fonte.query("*IDN?").strip())

    multimetro = rm.open_resource(MULTIMETRO_ADDR)
    multimetro.timeout = 5000
    multimetro.write_termination = '\n'
    multimetro.read_termination = '\n'
    print("Multímetro conectado:", multimetro.query("*IDN?").strip())

    # Configura o multímetro para medir tensão DC
    multimetro.write("CONF:VOLT:DC")
    multimetro.write("VOLT:DC:NPLC 10")
    multimetro.write("TRIG:SOUR IMM")
    multimetro.write("SAMP:COUN 1")
    time.sleep(0.5)

    for i, estagio in enumerate(ESTAGIOS, 1):
        V = estagio["tensao"]
        I = estagio["corrente"]

        print(f"\n[{i}] Aplicando {V} V por {DURACAO_ESTAGIO} segundos...")

        fonte.write(f"VOLT {V}")
        fonte.write(f"CURR {I}")
        fonte.write("OUTP ON")

        t0 = time.time()
        while time.time() - t0 < DURACAO_ESTAGIO:
            try:
                tensao_fonte = V
                multimetro.write("INIT")
                tensao_medida = float(multimetro.query("FETCH?").split(',')[0])

                with open(CSV_FILE, mode='a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([f"{tensao_fonte:.2f}", f"{tensao_medida:.4f}"])

                print(f"Fonte: {tensao_fonte:.2f} V | Multímetro: {tensao_medida:.4f} V")
            except Exception as e:
                print("Erro durante leitura:", e)

            time.sleep(INTERVALO_MEDICAO)

        fonte.write("OUTP OFF")
        print(f"[{i}] Estágio concluído.")

except Exception as e:
    print("Erro durante a execução:", e)

finally:
    try:
        fonte.close()
        multimetro.close()
    except:
        pass
    print("\nMedições finalizadas e conexões encerradas.")

# === Geração do gráfico ===
try:
    df = pd.read_csv(CSV_FILE)

    plt.figure(figsize=(10, 6))

    plt.subplot(2, 1, 1)
    plt.plot(df.index, df["Tensao_Fonte_V"], label="Tensão da Fonte", color='blue')
    plt.ylabel("Tensão Fonte (V)")
    plt.title("Tensão da Fonte")
    plt.grid(True)

    plt.subplot(2, 1, 2)
    plt.plot(df.index, df["Tensao_Multimetro_V"], label="Tensão Multímetro", color='red')
    plt.ylabel("Tensão Multímetro (V)")
    plt.xlabel("Ordem da Leitura")
    plt.title("Tensão do Multímetro")
    plt.grid(True)

    plt.tight_layout()
    plt.show()

except Exception as e:
    print("Erro ao gerar gráfico:", e)
