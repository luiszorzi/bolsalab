import time
import pyvisa
import matplotlib.pyplot as plt
import csv
from datetime import datetime 

# Configurações dos equipamentos
RECURSO_FONTE = "USB0::0x0699::0x0392::C010658::INSTR"
IP_FLUKE = "172.30.248.100"
PORTA_FLUKE = 3490

INTERVALO = 1  # escolher o intervalo das medicoes
R = 5000.0       
C = 1000e-6        

try:
    rm = pyvisa.ResourceManager()

    # Conecta na fonte
    fonte = rm.open_resource(RECURSO_FONTE)
    _ = fonte.query("*IDN?")
    fonte.write("OUTP ON")  

    # Conecta no multímetro
    fluke = rm.open_resource(f"TCPIP0::{IP_FLUKE}::{PORTA_FLUKE}::SOCKET")
    fluke.timeout = 5000
    fluke.write_termination = '\n'
    fluke.read_termination = '\n'

    # Configura Fluke para leitura rápida
    fluke.write("CONF:VOLT:DC")
    fluke.write("VOLT:DC:NPLC 10")
    fluke.write("TRIG:SOUR IMM")
    fluke.write("SAMP:COUN 1")

except Exception as e:
    print(f"Erro na conexão: {e}")
    exit()

def ler_corrente_e_tensao():
    try:
        
        fluke.write("CONF:CURR:DC 10")
        fluke.write("CURR:DC:NPLC 10")
        fluke.write("INIT")
        leitura_i = fluke.query("FETCH?")
        i = float(leitura_i.strip().replace(',', '.'))

        fluke.write("CONF:VOLT:DC")
        fluke.write("INIT")
        leitura_v = fluke.query("FETCH?")
        v = float(leitura_v.strip().split(',')[0])

        return i, v
    except Exception as e:
        print(f"Erro na leitura: {e}")
        return 0.0, 0.0

def aplicar_voltagem_ate_limiar(voltagem_alvo, limite_tensao, condicao, next_time):
    print(f"\nAplicando {voltagem_alvo:.1f} V até atingir {limite_tensao:.2f} V...")
    fonte.write(f"VOLT {voltagem_alvo}")

    while True:
        now = time.time()
        if now < next_time:
            time.sleep(next_time - now)

        t = time.time() - inicio_global
        corrente, v_cap = ler_corrente_e_tensao()

        horario = datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
        horarios.append(horario)

        tempos.append(t)
        tensoes_capacitor.append(v_cap)
        correntes.append(corrente)
        setpoints_fonte.append(voltagem_alvo)

        print(f"[{horario}] | [{t:.3f}s] | Fonte: | {voltagem_alvo:.1f} V | Corrente: {corrente*1000:.4f} mA | Vcap: {v_cap:.4f} V")

        if condicao(v_cap, limite_tensao):
            print(f"Limiar de {limite_tensao:.2f} V atingido.")
            break

        next_time += INTERVALO

    return next_time

# Listas para dados
tempos = []
tensoes_capacitor = []
correntes = []
setpoints_fonte = []
horarios = []  
corrente_inicial, v_inicial = ler_corrente_e_tensao()
print(f"Tensão inicial do capacitor: {v_inicial:.6f} V")

inicio_global = time.time()
next_time = time.time()

# Experimento com controle de intervalo contínuo
next_time = aplicar_voltagem_ate_limiar(10, 9.5, lambda v, lim: v >= lim, next_time)
next_time = aplicar_voltagem_ate_limiar(20, 19.5, lambda v, lim: v >= lim, next_time)
next_time = aplicar_voltagem_ate_limiar(5,  5.5,  lambda v, lim: v <= lim, next_time)

print("\nFinalizando experimento...")
fonte.write("VOLT 0")
fonte.write("OUTP OFF")
fonte.close()
fluke.close()
print("Equipamentos desligados.")

arquivo = "dados_completos_rc.csv"
with open(arquivo, mode="w", newline='') as f:
    writer = csv.writer(f)
    writer.writerow(["Data e Hora", "Tempo (s)", "Tensão Capacitor (V)", "Corrente (mA)", "Tensão Fonte (V)"])  
    for horario, t, vcap, corrente, vfonte in zip(horarios, tempos, tensoes_capacitor, correntes, setpoints_fonte):
        writer.writerow([horario, t, vcap, corrente * 1000, vfonte])

print(f"Dados salvos em '{arquivo}'")

plt.figure(figsize=(14, 6))

plt.subplot(2, 1, 1)
plt.plot(tempos, tensoes_capacitor, label="Tensão no Capacitor", color="green")
plt.ylabel("Tensão (V)")
plt.legend()
plt.grid(True)

plt.subplot(2, 1, 2)
correntes_mA = [i * 1000 for i in correntes]
plt.plot(tempos, correntes_mA, label="Corrente no Circuito", color="red")
plt.xlabel("Tempo (s)")
plt.ylabel("Corrente (mA)")
plt.grid(True)
plt.legend()

plt.tight_layout()
plt.show()
