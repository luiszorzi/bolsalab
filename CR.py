import pyvisa
import csv
import time

def conectar_dispositivo(endereco):
    rm = pyvisa.ResourceManager()
    instr = rm.open_resource(endereco)
    instr.write("*IDN?")
    idn = instr.read()
    print("Instrumento conectado:", idn.strip())
    return instr

def configurar_resistencia(instr, resistencia_ohms):
    instr.write("FUNC RES")
    instr.write(f"RES {resistencia_ohms}")
    print(f"Resistência configurada para {resistencia_ohms} ohms")

def ativar_entrada(instr):
    instr.write("INP ON")
    print("Entrada ligada")

def desativar_entrada(instr):
    instr.write("INP OFF")
    print("Entrada desligada")

def medir(instr):
    tensao = float(instr.query("MEAS:VOLT?"))
    corrente = float(instr.query("MEAS:CURR?"))
    potencia = float(instr.query("MEAS:POW?"))
    return tensao, corrente, potencia

def salvar_csv(dados, caminho_csv):
    with open(caminho_csv, mode='w', newline='') as arquivo:
        writer = csv.writer(arquivo)
        writer.writerow(["Tensao (V)", "Corrente (A)", "Potencia (W)"])
        writer.writerows(dados)

def main():
    endereco = 'USB0::0x05E6::0x2380::802436052757810021::INSTR'
    caminho_csv = 'C:\\Users\\luisp\\Downloads\\projeto\\RC.csv'
    resistencia = 100  # Ohms

    instr = conectar_dispositivo(endereco)
    configurar_resistencia(instr, resistencia)
    ativar_entrada(instr)

    dados = []
    for _ in range(30):
        tensao, corrente, potencia = medir(instr)
        print(f"Tensão: {tensao:.3f} V | Corrente: {corrente:.3f} A | Potência: {potencia:.3f} W")
        dados.append([tensao, corrente, potencia])
        time.sleep(1)

    desativar_entrada(instr)
    salvar_csv(dados, caminho_csv)
    print(f"Dados salvos em {caminho_csv}")

if __name__ == "__main__":
    main()
