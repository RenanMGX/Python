import os
import shutil
import datetime
# Definir a configuração de localização para 'C'
#locale.setlocale(locale.LC_ALL, 'C')

# Agora tente definir a configuração de localização para a que você precisa
#locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# pasta com os arquivos
source_folder = r"\\server008\G\ARQ_PATRIMAR\WORK\PONTO\ultimos_arquivos"

# cria uma pasta nova com a data atual
today = datetime.date.today()
one_day = datetime.timedelta(days=1)
yesterday = today - one_day
mes = yesterday.strftime("%m - %B")
mes = str(mes).title()

if mes == "03 - Marã§O":
    mes = "03 - Março"
elif mes == "01 - January":
    mes = "01 - Janeiro"
elif mes == "02 - February":
    mes = "02 - Fevereiro"
elif mes == "03 - March":
    mes = "03 - Março"
elif mes == "04 - April":
    mes = "04 - Abril"
elif mes == "05 - May":
    mes = "05 - Maio"
elif mes == "06 - June":
    mes = "06 - Junho"
elif mes == "07 - July":
    mes = "07 - Julho"
elif mes == "08 - August":
    mes = "08 - Agosto"
elif mes == "09 - September":
    mes = "09 - Setembro"
elif mes == "10 - October":
    mes = "10 - Outubro"
elif mes == "11 - November":
    mes = "11 - Novembro"
elif mes == "12 - December":
    mes = "12 - Dezembro"

new_folder = os.path.join(r"\\server008\G\ARQ_PATRIMAR\WORK\PONTO\Escritório Central", str(yesterday.year), mes, f"dia {yesterday.day}")

ano_atual = datetime.datetime.now().year

# Loop para cada mês do ano

os.makedirs(new_folder, exist_ok=True)
# lista os arquivos dentro da pasta e copia-os para a nova pasta
for filename in os.listdir(source_folder):
    filepath = os.path.join(source_folder, filename)
    if os.path.isfile(filepath):
        shutil.copy(filepath, os.path.join(new_folder, filename))
        os.remove(filepath)

print("Arquivos copiados e originais apagados com sucesso!")
