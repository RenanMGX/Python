import os
import time
import socket

#Cria uma lista para armazenar os nomes dos computadores
computers = []
nome_maquina = "digite o nome da maquina"

#Popula a lista com os nomes dos computadores de 
for i in range(1, 251):
    computers.append(nome_maquina + str(i))

#Laço infinito para verificar o status de conexão dos computadores
while True:
# Cria uma lista para armazenar os status de conexão dos computadores
    statuses = []
# Verifica o status de conexão de cada computador e adiciona a lista de statuses
    for computer in computers:
        try:
            # Tenta se conectar ao computador
            socket.create_connection((computer, 135), timeout=1)
            # Se conseguir, adiciona "Online" a lista de statuses
            statuses.append("Online")
        except:
            # Se não conseguir, adiciona "Offline" a lista de statuses
            statuses.append("Offline")

# Exibe os nomes dos computadores e seus respectivos statuses
for i in range(len(computers)):
    print(computers[i] + ": " + statuses[i])

# Aguarda 1 segundo antes de verificar novamente
time.sleep(1)
import tkinter as tk
from tkinter import ttk

#Função para atualizar a lista de computadores e statuses na interface gráfica
def update_list():
    for i in range(len(computers)):
        try:
            # Tenta se conectar ao computador
            socket.create_connection((computers[i], 135), timeout=1)
            # Se conseguir, adiciona "Online" a lista de statuses
            statuses[i] = "Online"
        except:
            # Se não conseguir, adiciona "Offline" a lista de statuses
            statuses[i] = "Offline"
            # Limpa a lista antes de adicionar os novos valores
            listbox.delete(0, tk.END)
    # Adiciona os nomes dos computadores e seus respectivos statuses na lista
    for i in range(len(computers)):
        listbox.insert(tk.END, computers[i] + ": " + statuses[i])

# Agenda a próxima atualização da lista
root.after(1000, update_list)
#Cria a janela principal
root = tk.Tk()
root.title("Computer Status")

#Cria a lista para exibir os nomes dos computadores e statuses
listbox = tk.Listbox(root)
listbox.pack(fill=tk.BOTH, expand=True)

#Cria o botão para encerrar o script
end_button = ttk.Button(root, text="End", command=root.destroy)
end_button.pack()

#Inicia a atualização da lista
root.after(1000, update_list)

#Inicia a execução da interface gráfica
root.mainloop()

