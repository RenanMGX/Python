import os
import getpass
import re
from fuzzywuzzy import fuzz
import shutil

#função que identifica e retorna o arquivo .htm mais proximo do nome do usuario
def pegar_caminho_htm(pasta):
    # Obtém o nome de usuário do sistema operacional
    user = getpass.getuser()
    # Caminho da pasta de assinaturas do Outlook
    path = f'C:\\Users\\{user}\\AppData\\Roaming\\Microsoft\\{pasta}\\'
    #user = "thiago.silva"
    #path = f'\\\\patrimar148\\c$\\Users\\{user}\\AppData\\Roaming\\Microsoft\\Assinaturas'
    # Lista os arquivos na pasta de assinaturas
    files = os.listdir(path)
    # Procura os arquivos .htm com nome parecido com o usuário
    file_list = []
    for file in files:
        if file.endswith('.htm'):
            ratio = fuzz.ratio(file.lower(), user.lower())
            if ratio > 15:  # altere o valor aqui para ajustar a sensibilidade
                file_list.append(os.path.join(path, file))
    return file_list
# alterar tamanho dentro do arquivo .htm
def alterar_htm(caminho):
    # Abrir o arquivo HTML
    with open(caminho, "r") as f:
        content = f.read()
    # Encontrar a tag v:shape e alterar o valor da tag style
    content = re.sub(r'<v:shape\s+[^>]*style\s*=\s*["\']([^"\']*)["\'][^>]*>',
                 r'<v:shape style="width:330pt; height:140pt">', content)
    # Escrever o conteúdo modificado em um novo arquivo
    with open(caminho, "w") as f:
        f.write(content)
# copia o arquivo para a pasta
def copiar_arquivo(destino_path):
    # caminho de origem dos arquivos de imagem
    origem_path = '\\\\server011\\NETLOGON\\ASSINATURA\\'
    # listar todos os arquivos no caminho de origem
    arquivos = os.listdir(origem_path)
    # encontrar o arquivo de imagem
    imagem = None
    for arquivo in arquivos:
        if arquivo.endswith('.png') or arquivo.endswith('.jpg'):
            imagem = arquivo
            break
    # copiar a imagem para o novo local com nome "image001.png"
    if imagem:
        origem = os.path.join(origem_path, imagem)
        destino_png = os.path.join(destino_path, 'image001.png')
        shutil.copy(origem, destino_png)
    # copiar a imagem para o novo local com nome "image001.jpg"
    if imagem:
        origem = os.path.join(origem_path, imagem)
        destino_jpg = os.path.join(destino_path, 'image001.jpg')
        shutil.copy(origem, destino_jpg)
# Imprime a assinatura
pastas = ["Signatures","Assinaturas"]
caminho_htm = []
for pasta in pastas:
    try:
        pasta_temporario = pegar_caminho_htm(pasta)
    except:
        continue
    for caminho in pasta_temporario:
        caminho_htm.append(caminho)
for assinatura in caminho_htm:
    caminho_arquivo = assinatura[:-4] + "_arquivos\\"
    alterar_htm(assinatura)
    copiar_arquivo(caminho_arquivo)
