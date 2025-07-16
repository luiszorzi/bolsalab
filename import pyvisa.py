import pyvisa
import time

# === Configurações da fonte ===
tensao_desejada = 5       # Volts
corrente_limite = 0.001  #  100ma

# === Inicializa comunicação com a fonte ===
rm = pyvisa.ResourceManager()
inst = rm.open_resource(rm.list_resources()[0])

print("Instrumento conectado:")
print(inst.query("*IDN?").strip())

# === Configura a fonte ===
inst.write("*RST")  # Reseta a fonte
inst.write(f"VOLT {tensao_desejada}")
inst.write(f"CURR {corrente_limite}")
inst.write("OUTP ON")
time.sleep(1)  # Espera estabilizar

# === Medições ===
tensao_real = float(inst.query("MEAS:VOLT?").strip())
corrente_real = float(inst.query("MEAS:CURR?").strip())

# === Estima resistência medida ===
resistencia_calculada = tensao_real / corrente_real if corrente_real > 0 else None

# === Lê status bruto da fonte ===
status_bruto = inst.query("STAT:OPER:COND?").strip()
status_int = int(status_bruto)

# === Interpreta o modo de operação ===
if (status_int & 2048) != 0:
    modo_operacao = "CV (Tensão constante)"
else:
    modo_operacao = "CC (Corrente constante ou outro)"

# === Resultados ===
print(f"\nTensão medida: {tensao_real:.5f} V")
print(f"Corrente medida: {corrente_real:.5f} A")
print(f"Resistência estimada: {resistencia_calculada:.2f} Ω" if resistencia_calculada else "Corrente zero, resistência não estimada")
print(f"Modo de operação da fonte: {modo_operacao}")

# === Desliga saída ===
inst.write("OUTP OFF")
print("Saída desligada.")
