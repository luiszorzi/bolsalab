import tkinter as tk
from tkinter import ttk, messagebox

# Credenciais fixas (você pode futuramente ler de um banco de dados)
USUARIO_CORRETO = "admin"
SENHA_CORRETA = "1234"

# Função que será chamada após login bem-sucedido
def abrir_interface_principal():
    janela_principal = tk.Toplevel()
    janela_principal.title("Interface Principal")
    janela_principal.geometry("600x400")

    label = ttk.Label(janela_principal, text="Bem-vindo à Interface Principal!", font=("Arial", 14))
    label.pack(pady=20)

# Verificação de login
def verificar_login():
    usuario = entry_usuario.get()
    senha = entry_senha.get()

    if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
        root.destroy()  # Fecha janela de login
        abrir_interface_principal()
    else:
        messagebox.showerror("Erro", "Usuário ou senha incorretos.")

# Janela de login
root = tk.Tk()
root.title("Login")
root.geometry("300x200")

frame_login = ttk.Frame(root, padding=20)
frame_login.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame_login, text="Usuário:").pack(pady=5)
entry_usuario = ttk.Entry(frame_login)
entry_usuario.pack()

ttk.Label(frame_login, text="Senha:").pack(pady=5)
entry_senha = ttk.Entry(frame_login, show="*")
entry_senha.pack()

btn_login = ttk.Button(frame_login, text="Entrar", command=verificar_login)
btn_login.pack(pady=15)

root.mainloop()
