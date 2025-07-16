import time
import pyvisa
import matplotlib.pyplot as plt
import csv
from datetime import datetime  # <-- IMPORTADO

RECURSO_FONTE = "USB0::0x0699::0x0392::C010658::INSTR"
IP_FLUKE = "172.30.248.100"
PORTA_FLUKE = 3490

INTERVALO = 0.1  
R = 10000.0        
C = 100e-6       

try:
    rm = pyvisa.ResourceManager()

    fonte = rm.open_resource(RECURSO_FONTE)
    _ = fonte.query("*IDN?")

    fluke = rm.open_resource(f"TCPIP0::{IP_FLUKE}::{PORTA_FLUKE}::SOCKET")
    fluke.timeout = 3000
    fluke.write_termination = '\n'
    fluke.read_termination = '\n'

    fluke.write("CONF:VOLT:DC")
    fluke.write("VOLT:DC:NPLC 10")
    fluke.write("TRIG:SOUR IMM")
    fluke.write("SAMP:COUN 1")

except Exception as e:
    print(f"Erro na conexão: {e}")
    exit()

def ler_tensao():
    fluke.write("INIT")
    leitura = fluke.query("FETCH?")
    try:
        return float(leitura.strip().split(',')[0])
    except:
        print(f"Erro na leitura: '{leitura}'")
        return 0.0

def aplicar_voltagem_ate_limiar(voltagem_alvo, limite, condicao):
    print(f"\nAplicando {voltagem_alvo:.1f} V até atingir {limite:.2f} V...")
    fonte.write(f"VOLT {voltagem_alvo}")
    
    while True:
        t = time.time() - inicio_global
        v_cap = ler_tensao()
        horario = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        tempos.append(t)
        tensoes_capacitor.append(v_cap)
        setpoints_fonte.append(voltagem_alvo)
        horarios.append(horario)

        print(f"[{horario}] | [{t:.3f}s] | Fonte: {voltagem_alvo:.1f} V | Vcap: {v_cap:.4f} V")

        if condicao(v_cap, limite):
            print(f"Limiar de {limite:.2f} V atingido.")
            break

        time.sleep(INTERVALO)

# Listas de dados
tempos = []
tensoes_capacitor = []
setpoints_fonte = []
horarios = []

# Início
print("\nIniciando experimento RC")
fonte.write("OUTP ON")
fonte.write("VOLT 0")
time.sleep(0.5)

v_inicial = ler_tensao()
print(f"Tensão inicial do capacitor: {v_inicial:.6f} V")

inicio_global = time.time()

# Etapas do experimento
aplicar_voltagem_ate_limiar(voltagem_alvo=10, limite=9.5, condicao=lambda v, lim: v >= lim)
aplicar_voltagem_ate_limiar(voltagem_alvo=20, limite=19.5, condicao=lambda v, lim: v >= lim)
aplicar_voltagem_ate_limiar(voltagem_alvo=5,  limite=5.5, condicao=lambda v, lim: v <= lim)

# Encerramento
print("\nFinalizando experimento...")
fonte.write("VOLT 0")
fonte.write("OUTP OFF")
fonte.close()
fluke.close()
print("Equipamentos desligados.")

# Salvar dados
nome_arquivo = "dados_rc_evento_com_horario.csv"
with open(nome_arquivo, mode="w", newline='') as f:
    writer = csv.writer(f)
    writer.writerow(["Data e Hora", "Tempo (s)", "Tensão Capacitor (V)", "Tensão Fonte (V)"])
    for horario, t, vcap, vfonte in zip(horarios, tempos, tensoes_capacitor, setpoints_fonte):
        writer.writerow([horario, t, vcap, vfonte])

print(f"Dados salvos em '{nome_arquivo}'")

# Gráfico
plt.figure(figsize=(12, 6))
plt.plot(tempos, tensoes_capacitor, label="Tensão no Capacitor (Medida)", color="green", linewidth=2)
plt.title("Tensão no Circuito RC")
plt.xlabel("Tempo (s)")
plt.ylabel("Tensão (V)")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
