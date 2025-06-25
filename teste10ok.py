import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tempfile
import shutil
import os


# Caminho para a planilha
arquivo_planilha = "/home/felipesouza/Documents/projetoChmadosSua/chamados.xlsx"
df = pd.read_excel(arquivo_planilha)

# Configurar opções do navegador Edge
edge_options = Options()
edge_options.add_argument("--start-maximized")

user_data_dir = tempfile.mkdtemp()
if os.path.exists(user_data_dir):
    shutil.rmtree(user_data_dir)
edge_options.add_argument(f"--user-data-dir={user_data_dir}")

navegador = webdriver.Edge(
    service=Service(EdgeChromiumDriverManager().install()), options=edge_options
)

# Acessar o site e fazer login
navegador.get("https://sua.riobranco.ac.gov.br/index.php?noAUTO=1")
login, senha = 'felipe.pinto', 'felypekratos2'
navegador.find_element(By.XPATH, '//*[@id="login_name"]').send_keys(login)
navegador.find_element(By.XPATH, '//*[@id="login_password"]').send_keys(senha)
navegador.find_element(By.XPATH, '//*[@id="boxlogin"]/form/p[5]/input').click()
time.sleep(2)

# Loop pelos chamados
for index, row in df.iterrows():
    navegador.get("https://sua.riobranco.ac.gov.br/front/ticket.form.php")
    time.sleep(3)

    titulo = row['TITULO']
    descricao = row['DESCRICAO']
    hora = row['HORA']
    categoria = row['CATEGORIA']
    atribuido = row['ATRIBUIDO']
    localizacao = row['LOCALIZACAO']
    unidade = row['UNIDADE']
    hora_formatada = pd.to_datetime(hora).strftime("%d-%m-%Y %H:%M:%S")

    # Título
    navegador.find_element(
        By.XPATH, '//*[@id="mainformtable4"]/tbody/tr[1]/td/input'
    ).send_keys(titulo + ' - ' + unidade)
    time.sleep(2)

    # Descrição (TinyMCE)
    iframe = WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, '//iframe[contains(@id,"ifr")]'))
    )
    navegador.switch_to.frame(iframe)
    editor = WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.ID, 'tinymce'))
    )
    editor.clear()
    editor.send_keys(descricao)
    navegador.switch_to.default_content()

    # Categoria (Select2)
    try:
        cat = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//span[contains(@id,"select2-dropdown_itilcategories_id") '
                           'and contains(@class,"select2-selection__rendered")]')
            )
        )
        cat.click()
        inp = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]'))
        )
        inp.send_keys(categoria)
        time.sleep(1)
        inp.send_keys(Keys.RETURN)
        print(f"[{index}] Categoria: {categoria}")
    except Exception as e:
        print(f"[{index}] Erro Categoria:", e)

    # Atribuído (Select2)
    try:
        att = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//span[contains(@id,"select2-dropdown__users_id_assign") '
                           'and contains(@class,"select2-selection__rendered")]')
            )
        )
        att.click()
        inp = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]'))
        )
        inp.send_keys(atribuido)
        time.sleep(1)
        inp.send_keys(Keys.RETURN)
        print(f"[{index}] Atribuído: {atribuido}")
    except Exception as e:
        print(f"[{index}] Erro Atribuído:", e)

    # Localização (Select2)
    try:
        loc = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//span[contains(@id,"select2-dropdown_locations_id") '
                           'and contains(@class,"select2-selection__rendered")]')
            )
        )
        loc.click()
        inp = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]'))
        )
        inp.send_keys(localizacao)
        time.sleep(1)
        inp.send_keys(Keys.RETURN)
        print(f"[{index}] Localização: {localizacao}")
    except Exception as e:
        print(f"[{index}] Erro Localização:", e)

    # Data/Hora (flatpickr)
    cal = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@class="input-button"]'))
    )
    cal.click()

    campo = WebDriverWait(navegador, 10).until(
        EC.visibility_of_element_located(
            (By.XPATH, '//input[@class="no-wrap flatpickr form-control input active"]')
        )
    )
    campo.clear()
    campo.send_keys(hora_formatada)
    campo.send_keys(Keys.RETURN)

    try:
        WebDriverWait(navegador, 10).until(
            EC.invisibility_of_element_located(
                (By.XPATH, '//div[contains(@class,"flatpickr-calendar")]')
            )
        )
    except:
        navegador.find_element(By.XPATH, '//body').click()

    time.sleep(2)

    #clicar em enviar após preencher dados do chamado
    navegador.find_element(By.XPATH, '//*[@id="tabsbody"]/div[2]/button').click()



# Finalização
input("Pressione Enter para sair...")
navegador.quit()
