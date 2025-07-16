import pyvisa
import time
import csv

# Inicia VISA e conecta ao primeiro instrumento disponível
rm = pyvisa.ResourceManager()
inst = rm.open_resource(rm.list_resources()[0])

print("Conectado a:", inst.query("*IDN?").strip())

# Liga a saída da fonte
inst.write("OUTP ON")

# Cria arquivo CSV para salvar os dados
with open("medicoes.csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Tensao (V)", "Corrente (A)"])

    for _ in range(12):  # 12 medições (1 por 5 segundos = 1 minuto)
        v = inst.query("MEAS:VOLT?").strip()
        a = inst.query("MEAS:CURR?").strip()
        writer.writerow([v, a])
        print(f"{v} V | {a} A")
        time.sleep(5)

# Desliga saída
inst.write("OUTP OFF")
print("Fim das medições.")
