import tkinter as tk
from tkinter import ttk, messagebox
import random

# ---------- Janela da CARGA ELETRÔNICA ----------
class CargaEletronicaWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Interface - Carga Eletrônica")
        self.geometry("400x300")

        tk.Label(self, text="Carga Eletrônica", font=("Arial", 16)).pack(pady=10)

        frame_modo = tk.Frame(self)
        frame_modo.pack(pady=5)
        tk.Label(frame_modo, text="Modo:").pack(side=tk.LEFT)
        self.modo_var = tk.StringVar()
        self.combo_modo = ttk.Combobox(frame_modo, textvariable=self.modo_var, state="readonly")
        self.combo_modo['values'] = ['Corrente Constante (CC)', 'Tensão Constante (CV)',
                                     'Potência Constante (CP)', 'Resistência Constante (CR)']
        self.combo_modo.current(0)
        self.combo_modo.pack(side=tk.LEFT)
        self.combo_modo.bind("<<ComboboxSelected>>", self.atualizar_unidade)

        frame_valor = tk.Frame(self)
        frame_valor.pack(pady=5)
        tk.Label(frame_valor, text="Valor:").pack(side=tk.LEFT)
        self.entry_valor = tk.Entry(frame_valor, width=10)
        self.entry_valor.pack(side=tk.LEFT)
        self.label_unidade = tk.Label(frame_valor, text="A")
        self.label_unidade.pack(side=tk.LEFT)

        self.btn_aplicar = tk.Button(self, text="Aplicar", command=self.aplicar)
        self.btn_aplicar.pack(pady=15)

        self.label_v = tk.Label(self, text="Tensão: --- V", font=("Arial", 12))
        self.label_i = tk.Label(self, text="Corrente: --- A", font=("Arial", 12))
        self.label_p = tk.Label(self, text="Potência: --- W", font=("Arial", 12))
        self.label_v.pack(pady=2)
        self.label_i.pack(pady=2)
        self.label_p.pack(pady=2)

    def atualizar_unidade(self, event=None):
        modo = self.modo_var.get()
        if "Corrente" in modo:
            self.label_unidade.config(text="A")
        elif "Tensão" in modo:
            self.label_unidade.config(text="V")
        elif "Potência" in modo:
            self.label_unidade.config(text="W")
        elif "Resistência" in modo:
            self.label_unidade.config(text="Ω")

    def aplicar(self):
        try:
            valor = float(self.entry_valor.get())
        except ValueError:
            messagebox.showerror("Erro", "Insira um valor numérico válido.")
            return

        modo = self.modo_var.get()
        tensao = round(random.uniform(10.0, 13.0), 2)

        if "Corrente" in modo:
            corrente = valor
        elif "Tensão" in modo:
            corrente = round(random.uniform(0.5, 2.0), 2)
            tensao = valor
        elif "Potência" in modo:
            corrente = round(valor / tensao, 2)
        elif "Resistência" in modo:
            corrente = round(tensao / valor, 2)
        else:
            corrente = 0

        potencia = round(tensao * corrente, 2)

        self.label_v.config(text=f"Tensão: {tensao} V")
        self.label_i.config(text=f"Corrente: {corrente} A")
        self.label_p.config(text=f"Potência: {potencia} W")


# ---------- Janela do MULTÍMETRO ----------
class MultimetroWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Interface - Multímetro")
        self.geometry("400x250")

        tk.Label(self, text="Interface do Multímetro", font=("Arial", 14)).pack(pady=10)

        frame_ip = tk.Frame(self)
        frame_ip.pack(pady=5)
        tk.Label(frame_ip, text="IP:").pack(side=tk.LEFT)
        self.entry_ip = tk.Entry(frame_ip)
        self.entry_ip.pack(side=tk.LEFT)
        self.entry_ip.insert(0, "192.168.0.10")

        frame_porta = tk.Frame(self)
        frame_porta.pack(pady=5)
        tk.Label(frame_porta, text="Porta:").pack(side=tk.LEFT)
        self.entry_porta = tk.Entry(frame_porta)
        self.entry_porta.pack(side=tk.LEFT)
        self.entry_porta.insert(0, "3490")

        self.btn_conectar = tk.Button(self, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(pady=10)

        self.btn_ler = tk.Button(self, text="Ler Tensão", command=self.ler_tensao, state=tk.DISABLED)
        self.btn_ler.pack(pady=10)

        self.label_resultado = tk.Label(self, text="Tensão: --- V", font=("Arial", 16))
        self.label_resultado.pack(pady=10)

    def conectar(self):
        ip = self.entry_ip.get()
        porta = self.entry_porta.get()
        if ip and porta.isdigit():
            messagebox.showinfo("Conectado", f"Conectado ao Multímetro em {ip}:{porta}")
            self.btn_ler.config(state=tk.NORMAL)
        else:
            messagebox.showerror("Erro", "IP ou Porta inválidos.")

    def ler_tensao(self):
        tensao = round(random.uniform(11.5, 12.5), 3)
        self.label_resultado.config(text=f"Tensão: {tensao} V")


# ---------- Janela da FONTE DE ALIMENTAÇÃO ----------
class FonteAlimentacaoWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Interface - Fonte de Alimentação")
        self.geometry("400x250")

        tk.Label(self, text="Fonte de Alimentação", font=("Arial", 16)).pack(pady=10)

        frame_tensao = tk.Frame(self)
        frame_tensao.pack(pady=5)
        tk.Label(frame_tensao, text="Tensão (V):").pack(side=tk.LEFT)
        self.entry_tensao = tk.Entry(frame_tensao, width=10)
        self.entry_tensao.pack(side=tk.LEFT)
        self.entry_tensao.insert(0, "12.0")

        frame_corrente = tk.Frame(self)
        frame_corrente.pack(pady=5)
        tk.Label(frame_corrente, text="Limite Corrente (A):").pack(side=tk.LEFT)
        self.entry_corrente = tk.Entry(frame_corrente, width=10)
        self.entry_corrente.pack(side=tk.LEFT)
        self.entry_corrente.insert(0, "2.0")

        self.btn_aplicar = tk.Button(self, text="Aplicar", command=self.aplicar)
        self.btn_aplicar.pack(pady=15)

        self.label_tensao_saida = tk.Label(self, text="Tensão aplicada: --- V", font=("Arial", 12))
        self.label_corrente_saida = tk.Label(self, text="Corrente medida: --- A", font=("Arial", 12))
        self.label_tensao_saida.pack(pady=2)
        self.label_corrente_saida.pack(pady=2)

    def aplicar(self):
        try:
            tensao = float(self.entry_tensao.get())
            corrente_limite = float(self.entry_corrente.get())
        except ValueError:
            messagebox.showerror("Erro", "Insira valores numéricos válidos.")
            return

        # Simulação da corrente medida (pode ser aleatória até o limite)
        corrente_medida = round(random.uniform(0, corrente_limite), 2)

        self.label_tensao_saida.config(text=f"Tensão aplicada: {tensao} V")
        self.label_corrente_saida.config(text=f"Corrente medida: {corrente_medida} A")


# ---------- Janela principal ----------
class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Equipamento")
        self.geometry("350x300")

        tk.Label(self, text="Escolha o Equipamento", font=("Arial", 20)).pack(pady=20)

        btn_multimetro = tk.Button(self, text="Multímetro", width=25, command=self.abrir_multimetro)
        btn_fonte = tk.Button(self, text="Fonte de Alimentação", width=25, command=self.abrir_fonte)
        btn_carga = tk.Button(self, text="Carga Eletrônica", width=25, command=self.abrir_carga)

        btn_multimetro.pack(pady=10)
        btn_fonte.pack(pady=10)
        btn_carga.pack(pady=10)

    def abrir_multimetro(self):
        MultimetroWindow(self)

    def abrir_fonte(self):
        FonteAlimentacaoWindow(self)

    def abrir_carga(self):
        CargaEletronicaWindow(self)


# ---------- Executar o programa ----------
if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
