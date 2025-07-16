import tkinter as tk
from tkinter import ttk, messagebox
import random
import pyvisa

class FonteAlimentacaoWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Interface - Fonte de Alimentação (PWS4305)")
        self.geometry("400x300")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.rm = pyvisa.ResourceManager()
        self.instrumento = None

        tk.Label(self, text="Fonte Tektronix PWS4305", font=("Arial", 16)).pack(pady=10)

        frame_endereco = tk.Frame(self)
        frame_endereco.pack(pady=5)
        tk.Label(frame_endereco, text="Endereço VISA:").pack(side=tk.LEFT)
        self.entry_endereco = tk.Entry(frame_endereco, width=30)
        self.entry_endereco.pack(side=tk.LEFT)
        self.entry_endereco.insert(0, "USB0::0x0699::0x0408::C000012::INSTR")

        self.btn_conectar = tk.Button(self, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(pady=10)

        frame_tensao = tk.Frame(self)
        frame_tensao.pack(pady=5)
        tk.Label(frame_tensao, text="Tensão (V):").pack(side=tk.LEFT)
        self.entry_tensao = tk.Entry(frame_tensao, width=10)
        self.entry_tensao.pack(side=tk.LEFT)
        self.entry_tensao.insert(0, "12.0")

        frame_corrente = tk.Frame(self)
        frame_corrente.pack(pady=5)
        tk.Label(frame_corrente, text="Corrente máx (A):").pack(side=tk.LEFT)
        self.entry_corrente = tk.Entry(frame_corrente, width=10)
        self.entry_corrente.pack(side=tk.LEFT)
        self.entry_corrente.insert(0, "2.0")

        self.btn_aplicar = tk.Button(self, text="Aplicar", state=tk.DISABLED, command=self.aplicar)
        self.btn_aplicar.pack(pady=15)

        self.label_tensao_saida = tk.Label(self, text="Tensão aplicada: --- V", font=("Arial", 12))
        self.label_corrente_saida = tk.Label(self, text="Corrente medida: --- A", font=("Arial", 12))
        self.label_tensao_saida.pack(pady=2)
        self.label_corrente_saida.pack(pady=2)

    def conectar(self):
        endereco = self.entry_endereco.get()
        try:
            self.instrumento = self.rm.open_resource(endereco)
            self.instrumento.timeout = 2000
            idn = self.instrumento.query("*IDN?")
            messagebox.showinfo("Conectado", f"Instrumento identificado:\n{idn}")
            self.btn_aplicar.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na conexão:\n{e}")

    def aplicar(self):
        try:
            tensao = float(self.entry_tensao.get())
            corrente = float(self.entry_corrente.get())
        except ValueError:
            messagebox.showerror("Erro", "Insira valores válidos.")
            return

        try:
            self.instrumento.write(f"VOLT {tensao}")
            self.instrumento.write(f"CURR {corrente}")
            self.instrumento.write("OUTP ON")

            tensao_real = float(self.instrumento.query("MEAS:VOLT?"))
            corrente_real = float(self.instrumento.query("MEAS:CURR?"))

            self.label_tensao_saida.config(text=f"Tensão aplicada: {tensao_real:.2f} V")
            self.label_corrente_saida.config(text=f"Corrente medida: {corrente_real:.3f} A")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao aplicar comandos:\n{e}")

    def on_close(self):
        try:
            if self.instrumento:
                self.instrumento.write("OUTP OFF")
                print("Fonte desligada ao fechar a janela.")
        except Exception as e:
            print("Erro ao desligar a fonte:", e)
        self.destroy()


# Exemplo para testar isoladamente a janela:
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Oculta janela principal
    FonteAlimentacaoWindow(root)
    root.mainloop()
