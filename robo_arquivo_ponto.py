import psutil
import pyautogui
import subprocess
import time

# # Encerra o processo GerenciadorRepWin.exe
# for proc in psutil.process_iter():
#     if proc.name() == "GerenciadorRepWin.exe":
#         proc.kill()

# # Abre o programa GerenciadorRepWin.exe
# pyautogui.press('win')
# pyautogui.typewrite('Gerenciador Rep')
# pyautogui.press('enter')

#executar =  subprocess.Popen("C:\\Program Files (x86)\\Topdata\\Gerenciador Inner Rep\\GerenciadorRepWin.exe")

# Espera 10 segundos para o programa abrir completamente
time.sleep(5)

# Clica na aba "Geral"
general_button_pos = pyautogui.locateOnScreen('botao_geral.png')  # substitua pelo arquivo de imagem do botão "Geral"
pyautogui.click(general_button_pos)

# Espera 2 segundos para a aba "Geral" ser selecionada
time.sleep(2)

# Clica no botão "Enviar configurações"
send_button_pos = pyautogui.locateOnScreen('botao_enviar_config.png')  # substitua pelo arquivo de imagem do botão "Enviar configurações"
pyautogui.click(send_button_pos)

time.sleep(2)

#clica no botão "comunicar"
send_button_pos = pyautogui.locateOnScreen('comunicar.png')  # substitua pelo arquivo de imagem do botão "Enviar configurações"
pyautogui.click(send_button_pos)

time.sleep(2)
# preenche os dados do cpf
pyautogui.write("111.111.111-11")

time.sleep(2)

#clica no botão "OK"
send_button_pos = pyautogui.locateOnScreen('ok.png')  # substitua pelo arquivo de imagem do botão "Enviar configurações"
pyautogui.click(send_button_pos)

