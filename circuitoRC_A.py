import time
import pyvisa
import matplotlib.pyplot as plt
import csv
from datetime import datetime  # Adicionado para capturar o horário

RECURSO_FONTE = "USB0::0x0699::0x0392::C010658::INSTR"
IP_FLUKE = "172.30.248.100"
PORTA_FLUKE = 3490

INTERVALO = 0.001   
R = 10000.0 

try:
    rm = pyvisa.ResourceManager()

    fonte = rm.open_resource(RECURSO_FONTE)
    _ = fonte.query("*IDN?")

    fluke = rm.open_resource(f"TCPIP0::{IP_FLUKE}::{PORTA_FLUKE}::SOCKET")
    fluke.timeout = 5000
    fluke.write_termination = '\n'
    fluke.read_termination = '\n'

    fluke.write("CONF:CURR:DC 10")
    fluke.write("CURR:DC:NPLC 10")  
    fluke.write("TRIG:SOUR IMM")
    fluke.write("SAMP:COUN 1")

except Exception as e:
    print(f"Erro na conexão: {e}")
    exit()

def ler_corrente_A():
    fluke.write("INIT")
    leitura = fluke.query("FETCH?")
    try:
        valor_str = leitura.strip().replace(',', '.')
        return float(valor_str)
    except Exception as e:
        print(f"Erro na leitura '{leitura}': {e}")
        return 0.0

def aplicar_voltagem_ate_corrente_limiar(voltagem_alvo, limite_A, condicao):
    print(f"\nAplicando {voltagem_alvo:.1f} V até corrente atingir {limite_A:.8f} A...")
    fonte.write(f"VOLT {voltagem_alvo}")
    
    while True:
        t = time.time() - inicio_global
        corrente = ler_corrente_A()
        horario = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Captura do horário atual

        tempos.append(t)
        correntes.append(corrente)
        horarios.append(horario)

        print(f"[{horario}] | [{t:.3f}s] Corrente: {corrente:.8f} A")

        if condicao(corrente, limite_A):
            print(f"Limiar de corrente {limite_A:.8f} A atingido.")
            break

        time.sleep(INTERVALO)

# Listas de dados
tempos = []
correntes = []
horarios = []  # Lista de horários

# Início
print("\nIniciando experimento RC com leitura apenas da corrente")
fonte.write("OUTP ON")
fonte.write("VOLT 0")
time.sleep(0.5)

inicio_global = time.time()

# Etapas do experimento
aplicar_voltagem_ate_corrente_limiar(voltagem_alvo=10, limite_A=0.00001, condicao=lambda i, lim: i <= lim)
aplicar_voltagem_ate_corrente_limiar(voltagem_alvo=20, limite_A=0.00001, condicao=lambda i, lim: i <= lim)
aplicar_voltagem_ate_corrente_limiar(voltagem_alvo=5,  limite_A=-0.00001, condicao=lambda i, lim: i >= lim)

# Encerramento
fonte.write("VOLT 0")
fonte.write("OUTP OFF")
fonte.close()
fluke.close()
print("Equipamentos desligados.")

# Salvar dados com horário
nome_arquivo = "corrente_experimento_com_horario.csv"
with open(nome_arquivo, mode="w", newline='') as f:
    writer = csv.writer(f)
    writer.writerow(["Data e Hora", "Tempo (s)", "Corrente (mA)"])
    for horario, t, i in zip(horarios, tempos, correntes):
        writer.writerow([horario, t, i * 1000])
print(f"Dados salvos em '{nome_arquivo}'")

# Gráfico em mA
print("Gerando gráfico")
correntes_mA = [i * 1000 for i in correntes]
plt.figure(figsize=(12, 6))
plt.plot(tempos, correntes_mA, label="Corrente no Circuito (mA)", color="red", linewidth=2)
plt.ylim(min(correntes_mA) * 1.2, max(correntes_mA) * 1.2)
plt.title("Corrente no Circuito RC")
plt.xlabel("Tempo (s)")
plt.ylabel("Corrente (mA)")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
