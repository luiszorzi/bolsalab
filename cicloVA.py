import pyvisa
import time

# === Lista de estágios: tensão (V), corrente (A), tempo ativo (s) ===
estagios = [
    {"tensao": 5.0, "corrente": 0.1,  "tempo": 3},
    {"tensao": 3.3, "corrente": 0.05, "tempo": 3},
    {"tensao": 2.0, "corrente": 0.02, "tempo": 3},
    {"tensao": 1.5, "corrente": 0.01, "tempo": 3},
    {"tensao": 10.0, "corrente": 2, "tempo": 3},
    {"tensao": 2.5, "corrente": 0.3, "tempo": 3}
]

tempo_espera_entre_estagios = 2  # segundos de espera entre ciclos

# === Inicializa comunicação com a fonte ===
rm = pyvisa.ResourceManager()
inst = rm.open_resource(rm.list_resources()[0])

print("Instrumento conectado:")
print(inst.query("*IDN?").strip())

# === Executa cada estágio ===
for i, estagio in enumerate(estagios, 1):
    tensao = estagio["tensao"]
    corrente = estagio["corrente"]
    tempo = estagio["tempo"]

    inst.write(f"VOLT {tensao}")
    inst.write(f"CURR {corrente}")
    inst.write("OUTP ON")

    print(f"\n[{i}] Aplicando {tensao} V / {corrente} A por {tempo} segundos")
    time.sleep(tempo)

    inst.write("OUTP OFF")
    print(f"[{i}] Saída desligada. Aguardando {tempo_espera_entre_estagios} segundos...")
    time.sleep(tempo_espera_entre_estagios)

# === Finaliza ===
print("\nSequência finalizada. Fonte desligada.")
