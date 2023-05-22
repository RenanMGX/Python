import os
import subprocess
import threading
import queue
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.app import App
from kivy.core.window import Window
import tkinter as tk
from tkinter import ttk
import shutil
import time
import winreg
# for i in range(2):
#     try:
#         from kivy.uix.boxlayout import BoxLayout
#         from kivy.uix.button import Button
#         from kivy.uix.checkbox import CheckBox
#         from kivy.uix.label import Label
#         from kivy.uix.scrollview import ScrollView
#         from kivy.app import App
#         from kivy.core.window import Window
#         import tkinter as tk
#         from tkinter import ttk
#         import shutil
#         import time
#         import winreg
#     except:
#         os.system("pip install pip --upgrade")
#         os.system("pip install app")
#         os.system("pip install kivy")
#         os.system("pip install winreg")
#         os.system("pip install tk")
#         os.system("pip install shutil")
#         os.system("pip install kivy[base] kivy_examples --pre --extra-index-url https://kivy.org/downloads/simple/")
#função para instalação de algum programa utilizando subprocess.popen
def install_popen(path):
    install = subprocess.Popen(path, shell=True)
    while install.poll() is None:
        time.sleep(1)
def intall_run(path):
    command = None
    if type(path) == list:
        command = [path[0], path[1]]
    else:
        command = path
    try:
        result = subprocess.run(command, shell=True)
        veri_program(result)
    except PermissionError:
        print("sem permissão")
#verificar instalação
def veri_program(result):
    if result.returncode == 0:
        print("Instalação concluída com sucesso!")
        # Caminho para o arquivo de execução do programa
        path_program = r"C:\Programas\program.exe"
        # Executando o programa
        subprocess.run([path_program], capture_output=True)
    else:
        print("A instalação falhou com código de erro:", result.returncode)
# Função para a instalação do winrar
def padronizar():
    original_file = "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\readerdc64_br_l_cra_mdr_install.exe"
    copy_file = "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\adobe_reader.exe"
    shutil.copy(original_file, copy_file)
    path = ["\\\\patrimar089\\e$\\Programas\\Outros\\winrar\\winrar-x64-620br.exe",
            "\\\\patrimar089\\e$\\Programas\\Outros\\adobe_reader\\adobe_reader.exe",
            "\\\\patrimar089\\e$\\Programas\\Outros\\Chrome\\ChromeSetup.exe"
    ]
    install_popen(path[2])
    intall_run([path[0], "/S"])
    intall_run([path[1], "/quiet"])
    Window.restore()
# Função para a instalação do sap_770
def sap_770():
    path = "\\\\patrimar089\\e$\\\Programas\\\Outros\\\SAP_770\\SetupAll.exe"
    intall_run(path)
    Window.restore()
# Função para a instalação do open_vpn
def open_vpn():
    path = "\\\\patrimar089\\e$\\Programas\\Outros\\VPN\\Open VPN\\OpenVPN-2.5.1-I601-amd64.msi"
    intall_run([path, "/passive"])
    original_file = "\\\\patrimar089\\e$\\Programas\\Outros\\Open_VPN\\nao_mexer\\Patrimar.ovpn"
    copy_file = "C:\\Users\\Public\\Downloads\\Patrimar.ovpn"
    shutil.copy(original_file, copy_file)
    copy_file2 = "C:\Program Files\OpenVPN\config\Patrimar.ovpn"
    command = f'powershell -Command "Start-Process cmd -ArgumentList \'/C, copy {original_file} {copy_file2} /y\' -Verb runAs"'
    subprocess.run(command, shell=True)
    shutil.copy(copy_file, copy_file2)
    Window.restore()
# Função para a instalação do office_2016
def office_2016():
    path = ["\\\\patrimar089\\e$\\Programas\\Outros\\Office\\2016 office\\Office - 2016 - Standard\\setup.exe",
            "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\2016 office\\Atualização outlook 2016\\outlook2016-kb4011123-fullfile-x64-glb.exe"
    ]
    intall_run(path[0])
    intall_run([path[1], "/quiet"])
    Window.restore()
# Função para a instalação do office_365
def office_365():
    path = "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\365 on-line\\Office365 64bit Setup.exe"
    intall_run(path)
    Window.restore()
# Função para a instalação do visualizador_sketchup   
def visualizador_sketchup():
    path = "\\\\patrimar089\\e$\\Programas\\Outros\\Sketchup 2022\\Sketchup Visualizador\\SketchUpViewer-2022-0-354-126.exe"
    intall_run([path, "/silent"])
    Window.restore()
def drive_hp_g8_250():
    path = ["\\\\patrimar089\\e$\\Programas\\Outros\\drive_hp_g8_250\\audio.exe",
            "\\\\patrimar089\\e$\\Programas\\Outros\\drive_hp_g8_250\\video.exe"
    ]
    intall_run(path[0])
    intall_run(path[0])
    Window.restore()
def project():
    path = "\\\\patrimar089\\e$\\Programas\\Outros\\Office\\MS Project\\MSProject_ptbr_64bits.exe"
    intall_run(path)
    Window.restore()
def adap_wifi():
    path = "\\\\patrimar089\\e$\\Programas\\Outros\\Drivers\\Adaptador Wifi\\DWA-131_E1_V5.11b03\\Setup.exe"
    intall_run(path)
    Window.restore()





#variavel com o nome das opçoes
items_lista = [
"Instalar Programas Padrões",
"Driver Adaptador WIFI", 
"SAP 770", 
"Open VPN", 
"Office 2016", 
"Office 365", 
"Visualizador Sketchup 2022", 
"Drive Video e Audio do HP G8 250 *Carregamento Demorado*",
"Project",
] 
#interface em Kivy    
class InstallerApp(App):
    def build(self):
        root = BoxLayout(orientation='vertical')
        self.items = items_lista
        self.checkboxes = {}
        scroll = ScrollView()
        container = BoxLayout(orientation='vertical', size_hint_y=None)
        container.bind(minimum_height=container.setter('height'))
        for item in self.items:
            item_box = BoxLayout(size_hint_y=None, height=30, padding=5)
            label = Label(text=item, size_hint_x=0.8)
            checkbox = CheckBox(size_hint_x=0.2)
            self.checkboxes[item] = checkbox
            item_box.add_widget(label)
            item_box.add_widget(checkbox)
            container.add_widget(item_box)
        scroll.add_widget(container)
        root.add_widget(scroll)
        install_button = Button(text="Instalar", size_hint_y=None, height=40, on_press=self.exec)
        root.add_widget(install_button)
        return root
    def exec(self, instance):
        # Window.minimize()
        self.install()
        # Window.restore()
    def install(self):
        for item, checkbox in self.checkboxes.items():
            if checkbox.active:
                Window.minimize()
                if item == "Instalar Programas Padrões":
                    print("Programas Padrões sendo executado")
                    padronizar()
                elif item == "SAP 770":
                    print("SAP 770 sendo executado")
                    sap_770()
                elif item == "Open VPN":
                    open_vpn()
                    print("Open VPN sendo executado")
                elif item == "Office 2016":
                    office_2016()
                    print("Office 2016 sendo executado")
                elif item == "Office 365":
                    office_365()
                    print("Office 365 sendo executado")
                elif item == "Visualizador Sketchup":
                    visualizador_sketchup()
                    print("Visualizador Sketchup sendo executado")
                elif item == "Drive Video e Audio do HP G8 250 *Carregamento Demorado*":
                    drive_hp_g8_250()
                    print("Drive Video e Audio do HP G8 250 sendo executado")
                elif item == "Project":
                    project()
                    print("Project sendo executado")
                elif item == "Driver Adaptador WIFI":
                    adap_wifi()
                    print("Driver Adaptador WIFI sendo executado")

if __name__ == '__main__':
    InstallerApp().run()
