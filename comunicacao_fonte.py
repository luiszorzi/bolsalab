import pyvisa  # importa pyvisa

rm = pyvisa.ResourceManager()  # gerenciador VISA
inst = rm.open_resource(rm.list_resources()[0])  # abre 1º instrumento
print(inst.query("*IDN?"))  # imprime identificação
