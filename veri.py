import os
import platform
import psutil
import socket
import subprocess
import openpyxl
from time import sleep

nome_maquina = "digite o nome da maquina"

# adiciona os dados na planilha do excel
def add_to_excel(hostname, username):
    # verifica se o arquivo já existe, se sim adiciona os dados a planilha existente, senão cria uma nova
    if os.path.isfile("computer_user.xlsx"):
        wb = openpyxl.load_workbook("computer_user.xlsx")
        sheet = wb.active
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Computer Name", "User Name"])
    sheet.append([hostname, username])
    wb.save("computer_user.xlsx")
    print("Dados adicionados à planilha do Excel com sucesso.")
    wb.close()

contador = 1
maximo = 300
while contador <= 250:
    sleep(3)
    print("\n")
    # Verificando se o computador está online
    if len(str(contador)) == 1:
        hostname = f"{nome_maquina}00" + str(contador)
    elif len(str(contador)) == 2:
        hostname = f"{nome_maquina}0" + str(contador)
    else:
        hostname = f"{nome_maquina}" + str(contador)
    # if socket.gethostbyname(hostname):
    #     print(f"{hostname} está online.")
    response = os.system("ping -c 1 " + hostname)

    if response == 0:
        # print(hostname, 'está online.')
        print()
    else:
        contador += 1
        with open("computer_info.txt", "a") as file:
            file.write("Nome do Computador: " + hostname +"; " + "*Computador está Offline \n")
        add_to_excel(hostname, "*Computador está Offline")
        # print(hostname, 'está offline.')
        continue

    # Extraindo as informações do computador
    try:
        username = subprocess.check_output("WMIC /NODE:" + hostname + " COMPUTERSYSTEM GET USERNAME", shell=True)
    except subprocess.CalledProcessError as errors:
        print(f"Error retrieving the username from {hostname}")
        with open("computer_info.txt", "a") as file:
            file.write("Nome do Computador: " + hostname +"; " + "*O servidor RCP está desativado \n")
        add_to_excel(hostname, "*O servidor RCP está desativado")
        contador += 1
        continue
    username = str(username)
    username = username.split(f"{nome_maquina.upper()}\\")
    print(username)
    if len(username) == 1:
        username = "vazio"
    else:
        username = username[1][1:]
        username = username.split(" ")
        username = str(username[0])
        # print(username)
    # Salvando as informações em um arquivo .txt
    with open("computer_info.txt", "a") as file:
        file.write("Nome do Computador: " + hostname +"; " + username + "\n")
    add_to_excel(hostname, username)
    if contador >= maximo:
        break
    contador += 1

    
