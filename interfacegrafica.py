import tkinter as tk
from tkinter import messagebox
import pyvisa
import threading
import time

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Equipamento")
        self.geometry("320x220")

        tk.Label(self, text="Escolha o equipamento:", font=("Arial", 14)).pack(pady=20)

        btn_fonte = tk.Button(self, text="Fonte de Alimentação", width=25, command=self.abrir_fonte)
        btn_fonte.pack(pady=5)

        btn_multimetro = tk.Button(self, text="Multímetro", width=25, command=self.abrir_multimetro)
        btn_multimetro.pack(pady=5)

        btn_carga = tk.Button(self, text="Carga Eletrônica", width=25, command=self.abrir_carga)
        btn_carga.pack(pady=5)

    def abrir_fonte(self):
        self.withdraw()
        fonte = FonteAlimentacaoWindow(self)
        fonte.grab_set()

    def abrir_multimetro(self):
        self.withdraw()
        mult = MultimetroWindow(self)
        mult.grab_set()

    def abrir_carga(self):
        self.withdraw()
        carga = CargaEletronicaWindow(self)
        carga.grab_set()

# -------------------- Fonte de Alimentação --------------------

class FonteAlimentacaoWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Fonte de Alimentação (PWS4305)")
        self.geometry("520x620")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.rm = pyvisa.ResourceManager()
        self.instrumento = None
        self.etapas = []

        tk.Label(self, text="Fonte Tektronix PWS4305", font=("Arial", 16)).pack(pady=10)

        frame_endereco = tk.Frame(self)
        frame_endereco.pack(pady=5)
        tk.Label(frame_endereco, text="Endereço VISA:").pack(side=tk.LEFT)
        self.entry_endereco = tk.Entry(frame_endereco, width=35)
        self.entry_endereco.pack(side=tk.LEFT)
        self.entry_endereco.insert(0, "USB0::0x0699::0x0408::C000012::INSTR")

        self.btn_buscar = tk.Button(frame_endereco, text="Buscar", command=self.buscar_enderecos)
        self.btn_buscar.pack(side=tk.LEFT, padx=5)

        self.btn_conectar = tk.Button(self, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(pady=10)

        self.frame_etapas = tk.Frame(self)
        self.frame_etapas.pack(pady=10)

        self.adicionar_etapa()

        self.btn_mais_etapa = tk.Button(self, text="+ Adicionar Etapa", command=self.adicionar_etapa, state=tk.DISABLED)
        self.btn_mais_etapa.pack(pady=5)

        self.btn_iniciar = tk.Button(self, text="Iniciar Etapas", command=self.iniciar_etapas, state=tk.DISABLED)
        self.btn_iniciar.pack(pady=10)

        self.label_status = tk.Label(self, text="Status: ---", font=("Arial", 12))
        self.label_status.pack(pady=10)

        self.btn_voltar = tk.Button(self, text="← Voltar", command=self.on_close)
        self.btn_voltar.pack(side=tk.BOTTOM, pady=10)

    def buscar_enderecos(self):
        try:
            recursos = self.rm.list_resources()
            if not recursos:
                messagebox.showinfo("Nenhum dispositivo", "Nenhum dispositivo VISA encontrado.")
                return

            janela_lista = tk.Toplevel(self)
            janela_lista.title("Dispositivos encontrados")
            janela_lista.geometry("420x300")
            tk.Label(janela_lista, text="Selecione um endereço:", font=("Arial", 12)).pack(pady=5)
            lista = tk.Listbox(janela_lista, width=60)
            lista.pack(pady=5)
            for r in recursos:
                lista.insert(tk.END, r)

            def selecionar():
                selecionado = lista.get(tk.ACTIVE)
                self.entry_endereco.delete(0, tk.END)
                self.entry_endereco.insert(0, selecionado)
                janela_lista.destroy()

            tk.Button(janela_lista, text="Selecionar", command=selecionar).pack(pady=5)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dispositivos:\n{e}")

    def conectar(self):
        endereco = self.entry_endereco.get()
        try:
            self.instrumento = self.rm.open_resource(endereco)
            self.instrumento.timeout = 3000
            idn = self.instrumento.query("*IDN?")
            messagebox.showinfo("Conectado", f"Instrumento identificado:\n{idn}")
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_mais_etapa.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na conexão:\n{e}")

    def adicionar_etapa(self):
        frame = tk.Frame(self.frame_etapas)
        frame.pack(pady=5)

        tk.Label(frame, text="Tensão (V):").pack(side=tk.LEFT)
        entry_v = tk.Entry(frame, width=7)
        entry_v.pack(side=tk.LEFT)
        entry_v.insert(0, "12.0")

        tk.Label(frame, text="Tempo (s):").pack(side=tk.LEFT)
        entry_t = tk.Entry(frame, width=5)
        entry_t.pack(side=tk.LEFT)
        entry_t.insert(0, "2")

        self.etapas.append((entry_v, entry_t))

    def iniciar_etapas(self):
        threading.Thread(target=self.executar_etapas, daemon=True).start()

    def executar_etapas(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        try:
            for i, (entry_v, entry_t) in enumerate(self.etapas):
                v = float(entry_v.get())
                t = float(entry_t.get())

                self.label_status.config(text=f"Aplicando {v} V por {t} s (Etapa {i+1})")
                self.instrumento.write(f"VOLT {v}")
                self.instrumento.write("OUTP ON")

                time.sleep(t)

                self.instrumento.write("OUTP OFF")
                time.sleep(0.5)

            self.label_status.config(text="Sequência finalizada. Saída desligada.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante as etapas:\n{e}")
        finally:
            self.btn_iniciar.config(state=tk.NORMAL)

    def on_close(self):
        try:
            if self.instrumento:
                self.instrumento.write("OUTP OFF")
        except Exception as e:
            print("Erro ao desligar a fonte:", e)
        self.master.deiconify()
        self.destroy()

# -------------------- Multímetro --------------------

class MultimetroWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Multímetro")
        self.geometry("460x320")
        self.protocol("WM_DELETE_WINDOW", self.voltar)

        self.rm = pyvisa.ResourceManager()
        self.instrumento = None

        tk.Label(self, text="Multímetro", font=("Arial", 16)).pack(pady=10)

        frame = tk.Frame(self)
        frame.pack(pady=5)

        tk.Label(frame, text="Endereço VISA:").pack(side=tk.LEFT)
        self.entry_endereco = tk.Entry(frame, width=35)
        self.entry_endereco.pack(side=tk.LEFT)
        self.entry_endereco.insert(0, "TCPIP::192.168.0.10::INSTR")

        self.btn_buscar = tk.Button(frame, text="Buscar", command=self.buscar_enderecos)
        self.btn_buscar.pack(side=tk.LEFT, padx=5)

        tk.Button(self, text="Conectar", command=self.conectar).pack(pady=10)
        self.label_medida = tk.Label(self, text="Tensão medida: ---", font=("Arial", 12))
        self.label_medida.pack(pady=10)

        self.btn_medida = tk.Button(self, text="Ler Medição", command=self.ler_medida, state=tk.DISABLED)
        self.btn_medida.pack(pady=10)

        tk.Button(self, text="← Voltar", command=self.voltar).pack(side=tk.BOTTOM, pady=10)

    def buscar_enderecos(self):
        try:
            recursos = self.rm.list_resources()
            if not recursos:
                messagebox.showinfo("Nenhum dispositivo", "Nenhum dispositivo VISA encontrado.")
                return

            janela_lista = tk.Toplevel(self)
            janela_lista.title("Dispositivos encontrados")
            janela_lista.geometry("420x300")
            tk.Label(janela_lista, text="Selecione um endereço:", font=("Arial", 12)).pack(pady=5)
            lista = tk.Listbox(janela_lista, width=60)
            lista.pack(pady=5)
            for r in recursos:
                lista.insert(tk.END, r)

            def selecionar():
                selecionado = lista.get(tk.ACTIVE)
                self.entry_endereco.delete(0, tk.END)
                self.entry_endereco.insert(0, selecionado)
                janela_lista.destroy()

            tk.Button(janela_lista, text="Selecionar", command=selecionar).pack(pady=5)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dispositivos:\n{e}")

    def conectar(self):
        endereco = self.entry_endereco.get()
        try:
            self.instrumento = self.rm.open_resource(endereco)
            self.instrumento.timeout = 3000
            idn = self.instrumento.query("*IDN?")
            messagebox.showinfo("Conectado", f"Instrumento identificado:\n{idn}")
            self.btn_medida.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na conexão:\n{e}")

    def ler_medida(self):
        try:
            valor = self.instrumento.query("MEAS:VOLT:DC?").strip()
            self.label_medida.config(text=f"Tensão medida: {valor} V")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na leitura:\n{e}")

    def voltar(self):
        self.master.deiconify()
        self.destroy()

# -------------------- Carga Eletrônica --------------------

class CargaEletronicaWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Carga Eletrônica")
        self.geometry("520x660")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.rm = pyvisa.ResourceManager()
        self.instrumento = None
        self.etapas = []

        tk.Label(self, text="Carga Eletrônica", font=("Arial", 16)).pack(pady=10)

        # Endereço VISA
        frame_endereco = tk.Frame(self)
        frame_endereco.pack(pady=5)
        tk.Label(frame_endereco, text="Endereço VISA:").pack(side=tk.LEFT)
        self.entry_endereco = tk.Entry(frame_endereco, width=40)
        self.entry_endereco.pack(side=tk.LEFT)
        self.entry_endereco.insert(0, "USB0::0x05E6::0x2380::802436052757810021::INSTR")

        self.btn_buscar = tk.Button(frame_endereco, text="Buscar", command=self.buscar_enderecos)
        self.btn_buscar.pack(side=tk.LEFT, padx=5)

        self.btn_conectar = tk.Button(self, text="Conectar", command=self.conectar)
        self.btn_conectar.pack(pady=10)

        # Modo operação
        frame_modo = tk.Frame(self)
        frame_modo.pack(pady=5)
        tk.Label(frame_modo, text="Modo de Operação:").pack(side=tk.LEFT)

        self.modo_var = tk.StringVar(value="Tensão Constante")
        modos = ["Tensão Constante", "Corrente Constante", "Potência Constante", "Resistência Constante"]
        self.modo_menu = tk.OptionMenu(frame_modo, self.modo_var, *modos, command=self.modo_alterado)
        self.modo_menu.pack(side=tk.LEFT)

        # Etapas
        self.frame_etapas = tk.Frame(self)
        self.frame_etapas.pack(pady=10)

        self.btn_mais_etapa = tk.Button(self, text="+ Adicionar Etapa", command=self.adicionar_etapa, state=tk.DISABLED)
        self.btn_mais_etapa.pack(pady=5)

        self.btn_iniciar = tk.Button(self, text="Iniciar Etapas", command=self.iniciar_etapas, state=tk.DISABLED)
        self.btn_iniciar.pack(pady=10)

        self.label_status = tk.Label(self, text="Status: ---", font=("Arial", 12))
        self.label_status.pack(pady=10)

        self.btn_voltar = tk.Button(self, text="← Voltar", command=self.on_close)
        self.btn_voltar.pack(side=tk.BOTTOM, pady=10)

        # Inicializa com uma etapa
        self.adicionar_etapa()

    def buscar_enderecos(self):
        try:
            recursos = self.rm.list_resources()
            if not recursos:
                messagebox.showinfo("Nenhum dispositivo", "Nenhum dispositivo VISA encontrado.")
                return

            janela_lista = tk.Toplevel(self)
            janela_lista.title("Dispositivos encontrados")
            janela_lista.geometry("420x300")
            tk.Label(janela_lista, text="Selecione um endereço:", font=("Arial", 12)).pack(pady=5)
            lista = tk.Listbox(janela_lista, width=60)
            lista.pack(pady=5)
            for r in recursos:
                lista.insert(tk.END, r)

            def selecionar():
                selecionado = lista.get(tk.ACTIVE)
                self.entry_endereco.delete(0, tk.END)
                self.entry_endereco.insert(0, selecionado)
                janela_lista.destroy()

            tk.Button(janela_lista, text="Selecionar", command=selecionar).pack(pady=5)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dispositivos:\n{e}")

    def conectar(self):
        endereco = self.entry_endereco.get()
        try:
            self.instrumento = self.rm.open_resource(endereco)
            self.instrumento.timeout = 3000
            idn = self.instrumento.query("*IDN?")
            messagebox.showinfo("Conectado", f"Instrumento identificado:\n{idn}")
            self.btn_mais_etapa.config(state=tk.NORMAL)
            self.btn_iniciar.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na conexão:\n{e}")

    def modo_alterado(self, _):
        # Limpa etapas e cria uma nova para o modo atual
        for w in self.frame_etapas.winfo_children():
            w.destroy()
        self.etapas.clear()
        self.adicionar_etapa()

    def adicionar_etapa(self):
        frame = tk.Frame(self.frame_etapas)
        frame.pack(pady=5)

        modo = self.modo_var.get()

        if modo == "Resistência Constante":
            tk.Label(frame, text="Resistência (Ω):").pack(side=tk.LEFT)
            entry_val = tk.Entry(frame, width=8)
            entry_val.pack(side=tk.LEFT)
            entry_val.insert(0, "100")

        else:
            unidade = {
                "Tensão Constante": "Tensão (V):",
                "Corrente Constante": "Corrente (A):",
                "Potência Constante": "Potência (W):"
            }
            tk.Label(frame, text=unidade.get(modo, "Valor:")).pack(side=tk.LEFT)
            entry_val = tk.Entry(frame, width=8)
            entry_val.pack(side=tk.LEFT)
            if modo == "Tensão Constante":
                entry_val.insert(0, "12.0")
            elif modo == "Corrente Constante":
                entry_val.insert(0, "1.0")
            elif modo == "Potência Constante":
                entry_val.insert(0, "10.0")

        tk.Label(frame, text="Tempo (s):").pack(side=tk.LEFT)
        entry_tempo = tk.Entry(frame, width=5)
        entry_tempo.pack(side=tk.LEFT)
        entry_tempo.insert(0, "2")

        self.etapas.append((entry_val, entry_tempo))

    def iniciar_etapas(self):
        threading.Thread(target=self.executar_etapas, daemon=True).start()

    def executar_etapas(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_mais_etapa.config(state=tk.DISABLED)
        try:
            modo = self.modo_var.get()
            for i, (entry_val, entry_tempo) in enumerate(self.etapas):
                valor = float(entry_val.get())
                tempo_s = float(entry_tempo.get())

                self.label_status.config(text=f"{modo}: {valor} por {tempo_s}s (Etapa {i+1})")

                if modo == "Tensão Constante":
                    self.instrumento.write("*RST")
                    self.instrumento.write("FUNC VOLT")
                    self.instrumento.write(f"VOLT {valor}")
                    self.instrumento.write("INPUT ON")

                elif modo == "Corrente Constante":
                    self.instrumento.write("*RST")
                    self.instrumento.write("FUNC CURR")
                    self.instrumento.write(f"CURR {valor}")
                    self.instrumento.write("INPUT ON")

                elif modo == "Potência Constante":
                    self.instrumento.write("*RST")
                    self.instrumento.write("FUNC POW")
                    self.instrumento.write(f"POW {valor}")
                    self.instrumento.write("INPUT ON")

                elif modo == "Resistência Constante":
                    self.instrumento.write("*RST")
                    self.instrumento.write("FUNC RES")
                    self.instrumento.write(f"RES {valor}")
                    self.instrumento.write("INPUT ON")

                else:
                    raise ValueError("Modo desconhecido")

                time.sleep(tempo_s)

                self.instrumento.write("INPUT OFF")
                time.sleep(0.5)

            self.label_status.config(text="Sequência finalizada. Saída desligada.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante as etapas:\n{e}")
        finally:
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_mais_etapa.config(state=tk.NORMAL)

    def on_close(self):
        try:
            if self.instrumento:
                self.instrumento.write("INPUT OFF")
        except Exception as e:
            print("Erro ao desligar a carga:", e)
        self.master.deiconify()
        self.destroy()

# --------------- Execução principal ----------------

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
