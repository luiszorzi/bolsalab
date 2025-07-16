import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import time
import threading
import random  # Substitua pelas leituras reais dos instrumentos

class InterfaceMedicao:
    def __init__(self, root):
        self.root = root
        self.root.title("Monitoramento dos Equipamentos")
        self.root.geometry("1000x600")

        # Dados
        self.tempo = []
        self.tensao_multimetro = []
        self.tensao_fonte = []
        self.corrente_carga = []
        self.coletando = False
        self.start_time = None

        # Frame gráfico
        self.fig, self.ax = plt.subplots(3, 1, figsize=(8, 6), sharex=True)
        self.fig.subplots_adjust(hspace=0.5)
        self.canvas = FigureCanvasTkAgg(self.fig, master=root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Títulos
        self.ax[0].set_title("Tensão (Multímetro Fluke 8846A)")
        self.ax[1].set_title("Tensão (Fonte PWS4305)")
        self.ax[2].set_title("Corrente (Carga Keithley 2380)")

        for a in self.ax:
            a.set_ylabel("Valor")
        self.ax[2].set_xlabel("Tempo (s)")

        # Linhas dos gráficos
        self.line_multimetro, = self.ax[0].plot([], [], label="Tensão Multímetro", color='blue')
        self.line_fonte, = self.ax[1].plot([], [], label="Tensão Fonte", color='green')
        self.line_carga, = self.ax[2].plot([], [], label="Corrente Carga", color='red')

        for a in self.ax:
            a.legend()

        # Botões
        frame = ttk.Frame(root)
        frame.pack(pady=10)

        self.btn_start = ttk.Button(frame, text="Iniciar", command=self.iniciar)
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_stop = ttk.Button(frame, text="Parar", command=self.parar, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=10)

    def iniciar(self):
        self.coletando = True
        self.start_time = time.time()
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        threading.Thread(target=self.coletar_dados).start()

    def parar(self):
        self.coletando = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)

    def coletar_dados(self):
        while self.coletando:
            t = round(time.time() - self.start_time, 1)

            # Simulações (substitua pelos comandos SCPI reais usando pyvisa)
            tensao_multimetro = round(12 + random.uniform(-0.2, 0.2), 2)
            tensao_fonte = round(12 + random.uniform(-0.1, 0.1), 2)
            corrente_carga = round(1.5 + random.uniform(-0.2, 0.2), 2)

            # Armazenar dados
            self.tempo.append(t)
            self.tensao_multimetro.append(tensao_multimetro)
            self.tensao_fonte.append(tensao_fonte)
            self.corrente_carga.append(corrente_carga)

            # Atualizar gráficos
            self.line_multimetro.set_data(self.tempo, self.tensao_multimetro)
            self.line_fonte.set_data(self.tempo, self.tensao_fonte)
            self.line_carga.set_data(self.tempo, self.corrente_carga)

            for a in self.ax:
                a.relim()
                a.autoscale_view()

            self.canvas.draw()
            time.sleep(1)

# Inicialização da interface
if __name__ == "__main__":
    root = tk.Tk()
    app = InterfaceMedicao(root)
    root.mainloop()
