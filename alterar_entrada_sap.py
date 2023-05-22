import os
import re
try:
    import psutil
    import tkinter as tk
    from tkinter import messagebox
except:
    os.system("pip install psutil")
    os.system("pip install tkinter")
    import psutil
    from tkinter import messagebox
    import tkinter as tk


try:
    # Verificar se o processo SAP GUI está aberto
    for proc in psutil.process_iter():
        if proc.name() == "saplogon.exe":
            proc.kill()
    # listar todos os usuários do computador
    users = os.listdir("C:\\Users\\")
    # percorrer cada usuário
    for user in users:
        filepath = f"C:\\Users\\{user}\\AppData\\Roaming\\SAP\\Common\\SAPUILandscape.xml"
        # verificar se o arquivo existe
        if os.path.isfile(filepath):
            # abrir arquivo
            with open(filepath, "r") as f:
                data = f.read()
            # procurar e substituir
            servidor = [["Servidor_antigo", "Servidor_novo"],]
            for serv in servidor:
                data = re.sub(f'server="{serv[0]}:3200"', f'server="{serv[1]}:3200"', data)
            # abrir arquivo novamente para escrever
            with open(filepath, "w") as f:
                f.write(data)
    # messagebox.showinfo("Informação", "O script foi concluído com sucesso!")

    def popup_completed():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Concluído", "Entrada SAP alterada com sucesso!")

# Chame a função ao final do seu script
    popup_completed()
except:
    messagebox.showerror("Erro", "Ocorreu um erro durante a execução do script.")

