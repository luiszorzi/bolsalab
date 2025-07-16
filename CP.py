import pyvisa
import time
import csv
import os

def conectar_dispositivo(endereco_usb):
    rm = pyvisa.ResourceManager()
    instr = rm.open_resource(endereco_usb)
    instr.timeout = 5000  # timeout em ms
    idn = instr.query("*IDN?").strip()
    print("Instrumento conectado:", idn)
    return instr

def configurar_potencia_constante(instr, potencia_watts):
    instr.write("*RST")  # Resetar o instrumento
    instr.write("FUNC POW")  # Modo potência constante
    instr.write(f"POW {potencia_watts}")  # Configurar potência
    instr.write("INPUT ON")  # Liga a carga
    time.sleep(0.1)  # Pequena espera para garantir que liga
    status_saida = instr.query("INPUT?").strip()
    print(f"Potência configurada para {potencia_watts} W | Saída ligada? {status_saida}")

def ler_medidas(instr):
    tensao = float(instr.query("MEAS:VOLT?"))
    corrente = float(instr.query("MEAS:CURR?"))
    potencia = float(instr.query("MEAS:POW?"))
    return tensao, corrente, potencia

def salvar_csv(nome_arquivo, dados, modo='a'):
    cabecalho = ['Tensão(V)', 'Corrente(A)', 'Potência(W)']
    arquivo_existe = os.path.isfile(nome_arquivo)
    with open(nome_arquivo, modo, newline='') as csvfile:
        writer = csv.writer(csvfile)
        if not arquivo_existe or modo == 'w':
            writer.writerow(cabecalho)
        writer.writerow(dados)

def main():
    endereco_usb = 'USB0::0x05E6::0x2380::802436052757810021::INSTR'
    arquivo_csv = r'C:\Users\luisp\Downloads\projeto\CP.csv'

    # Apaga arquivo CSV no início da execução para limpar dados antigos
    if os.path.exists(arquivo_csv):
        os.remove(arquivo_csv)

    instr = conectar_dispositivo(endereco_usb)
    configurar_potencia_constante(instr, 10)  # 10 W em modo potência constante

    tempo_total = 30  # segundos
    intervalo = 1    # intervalo de 1 segundo entre leituras

    print(f"Iniciando medições por {tempo_total} segundos...")

    inicio = time.time()
    while (time.time() - inicio) < tempo_total:
        v, i, p = ler_medidas(instr)
        print(f"Tensão: {v:.3f} V | Corrente: {i:.3f} A | Potência: {p:.3f} W")
        salvar_csv(arquivo_csv, [f"{v:.3f}", f"{i:.3f}", f"{p:.3f}"])
        time.sleep(intervalo)

    instr.write("INPUT OFF")  # Desliga a carga ao final
    print("Medição finalizada e carga desligada.")

if __name__ == "__main__":
    main()
