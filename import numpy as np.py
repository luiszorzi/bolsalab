import numpy as np
import matplotlib.pyplot as plt
import control as ctrl

# Tempo de simulação
t = np.linspace(0, 0.1, 1000)  # 0 a 100 ms

# Sistema RC de 2ª ordem aproximado (com base no gráfico enviado)
# Forma: G(s) = 1 / (0.001s^2 + 0.09s + 1)
num = [1]
den = [0.001, 0.09, 1]
G = ctrl.TransferFunction(num, den)

# Controladores
Kp = 30
Ki = 500
Kd = 0.001

# Controlador P
C_P = ctrl.TransferFunction([Kp], [1])
T_P = ctrl.feedback(C_P * G, 1)

# Controlador PI
C_PI = ctrl.TransferFunction([Kp, Ki], [1, 0])
T_PI = ctrl.feedback(C_PI * G, 1)

# Controlador PD
C_PD = ctrl.TransferFunction([Kd, Kp], [1])
T_PD = ctrl.feedback(C_PD * G, 1)

# Controlador PID
C_PID = ctrl.TransferFunction([Kd, Kp, Ki], [1, 0])
T_PID = ctrl.feedback(C_PID * G, 1)

# Sistema sem controle
T_open = ctrl.step_response(G, t)
T_resp_P = ctrl.step_response(T_P, t)
T_resp_PI = ctrl.step_response(T_PI, t)
T_resp_PD = ctrl.step_response(T_PD, t)
T_resp_PID = ctrl.step_response(T_PID, t)

# Plotando os resultados
plt.figure(figsize=(12, 7))
plt.plot(T_open.time, T_open.outputs, label='Sem Controle (Sistema)', linewidth=2)
plt.plot(T_resp_P.time, T_resp_P.outputs, label='Controlador P', linestyle='--')
plt.plot(T_resp_PI.time, T_resp_PI.outputs, label='Controlador PI', linestyle='-.')
plt.plot(T_resp_PD.time, T_resp_PD.outputs, label='Controlador PD', linestyle=':')
plt.plot(T_resp_PID.time, T_resp_PID.outputs, label='Controlador PID', linestyle='-', alpha=0.7)

# Linha de referência de 98%
plt.axhline(0.98, color='red', linestyle='--', linewidth=1)

plt.title("Comparação da Resposta ao Degrau\nSistema RC de 2ª Ordem com Controladores", fontsize=14)
plt.xlabel("Tempo (s)")
plt.ylabel("Tensão (normalizada)")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
