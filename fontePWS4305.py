import pyvisa
import time
import csv

# Cria o gerenciador de recursos VISA
rm = pyvisa.ResourceManager()

# Lista os instrumentos conectados
print("Instrumentos detectados:")
print(rm.list_resources())

# Usa o primeiro instrumento da lista (ajuste se necessário)
instrument_address = rm.list_resources()[0]
inst = rm.open_resource(instrument_address)

# Identificação do instrumento
print("Instrumento conectado:")
print(inst.query("*IDN?"))

# Configura a fonte
inst.write("VOLT 3")       # Define tensão para 5 V
inst.write("CURR 1")       # Define corrente máxima para 1 A
inst.write("OUTP ON")      # Liga a saída da fonte

# Cria e abre o arquivo CSV
with open("dados_pws4305.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["Tensao (V)", "Corrente (A)"])

    # Coleta 12 medições com intervalo de 5 segundos
    for i in range(12):
        voltage = inst.query("MEAS:VOLT?").strip()  # Limpa o valor
        current = inst.query("MEAS:CURR?").strip()  # Limpa o valor

        # Remover qualquer valor extra (como 'CV' ou 'RMT', s aparecer)
        voltage = voltage.replace("CV", "").strip()
        current = current.replace("RMT", "").strip()

        writer.writerow([voltage, current])
        print(f"{voltage} V - {current} A")
        time.sleep(5)

# Desliga a saída
inst.write("OUTP OFF")
print("Aquisição finalizada. Saída desligada.")
