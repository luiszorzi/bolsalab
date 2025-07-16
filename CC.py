import pyvisa
import time
import csv
import os

def conectar_dispositivo(endereco_usb):
    rm = pyvisa.ResourceManager()
    instr = rm.open_resource(endereco_usb)
    instr.timeout = 5000  # timeout em ms
    print("Instrumento conectado:", instr.query("*IDN?").strip())
    return instr

def configurar_corrente_constante(instr, corrente_amps):
    instr.write("*RST")  # Resetar o instrumento
    instr.write("FUNC CURR")  # Modo corrente constante
    instr.write(f"CURR {corrente_amps}")  # Configurar corrente
    instr.write("INPUT ON")  # Liga a carga
    print(f"Corrente configurada para {corrente_amps} A")

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
    endereco_usb = 'USB0::0x05E6::0x2380::802436052757810021::INSTR'  # Ajuste conforme seu instrumento
    arquivo_csv = r'C:\Users\luisp\Downloads\projeto\dados_carga.csv'

    # Apaga arquivo CSV no início da execução para limpar dados antigos
    if os.path.exists(arquivo_csv):
        os.remove(arquivo_csv)

    instr = conectar_dispositivo(endereco_usb)
    configurar_corrente_constante(instr, 1)  # 1A

    tempo_total = 30  # segundos
    intervalo = 1    # 1 segundo

    print(f"Iniciando medições por {tempo_total} segundos...")

    inicio = time.time()
    while (time.time() - inicio) < tempo_total:
        v, i, p = ler_medidas(instr)
        print(f"Tensão: {v:.3f} V | Corrente: {i:.3f} A | Potência: {p:.3f} W")
        salvar_csv(arquivo_csv, [f"{v:.3f}", f"{i:.3f}", f"{p:.3f}"])
        time.sleep(intervalo)

    instr.write("INPUT OFF")  # Desliga a carga no final
    print("Medição finalizada e carga desligada.")

if __name__ == "__main__":
    main()
