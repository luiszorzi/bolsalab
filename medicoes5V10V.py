import pyvisa
import time
import csv

rm = pyvisa.ResourceManager()
inst = rm.open_resource(rm.list_resources()[0])

with open(r"C:\Users\luisp\Downloads\projeto\dados_medicoes.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["Tensao (V)", "Corrente (A)"])

    intervalo = 5

    # Liga saída, configura 5V, espera estabilizar
    inst.write("OUTP OFF")
    inst.write("VOLT 5")
    inst.write("CURR 1")
    inst.write("OUTP ON")
    time.sleep(2)  # espera estabilizar

    # Fase 1: 6 medições de 5V
    for _ in range(6):
        tensao = inst.query("MEAS:VOLT?").strip()
        corrente = inst.query("MEAS:CURR?").strip()
        writer.writerow([tensao, corrente])
        print(f"{tensao} V  -  {corrente} A")
        time.sleep(intervalo)

    # Muda para 10V, espera estabilizar
    inst.write("VOLT 10")
    time.sleep(2)

    # Fase 2: 6 medições de 10V
    for _ in range(6):
        tensao = inst.query("MEAS:VOLT?").strip()
        corrente = inst.query("MEAS:CURR?").strip()
        writer.writerow([tensao, corrente])
        print(f"{tensao} V  -  {corrente} A")
        time.sleep(intervalo)

    inst.write("OUTP OFF")

print("Medições finalizadas. Fonte desligada.")
