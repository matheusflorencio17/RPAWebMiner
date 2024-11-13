import time
import logging
import pandas as pd
import os
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import requests
 
# Instanciando o arquivo .env para que possa ser feito o preenchimento dos dados sem alteração no código fonte
load_dotenv()

ITEM_NAME = os.getenv("ITEM_NAME")
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")
SEND_EMAIL = os.getenv("SEND_EMAIL")

# Configuração do arquivo para que seja feito o laço condicional e armazene no log
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("log.txt", mode='w')
    ]
)

# Função para verificar a disponibilidade do site
def verificando_site(url, max_retries=3):
    logging.info("Começando a verificar se o site está no ar.")
    for attempt in range(max_retries):
        try:
            response = requests.get(url)
            if response.status_code == 200:
                logging.info(f"Tentativa {attempt + 1}: Site {url} está disponível com status 200.")
                return True
            else:
                logging.warning(f"Tentativa {attempt + 1}: Site {url} respondeu com código {response.status_code}.")
        except requests.RequestException as e:
            logging.warning(f"Tentativa {attempt + 1}: Erro ao acessar {url}. Erro: {e}")
        time.sleep(2)

    logging.error(f"Site {url} fora do ar após {max_retries} tentativas.")
    return False

# Função para inicializar o navegador
def inicializando_nav():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-dev-shm-usage")  # Evita o uso excessivo de memória no Docker

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

# Função para realizar a busca de um item no site da Magazine Luiza
def pesq_item(driver):
    driver.get("https://www.magazineluiza.com.br/")

    logging.info("Site está em funcionamento!")
    try:
        item_buscar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#search-container > div > form > input"))
        )

        item_buscar.click()
        time.sleep(5)

        logging.info("Encontrada barra de pesquisa")

        item_buscar.send_keys(ITEM_NAME)
        item_buscar.send_keys(Keys.RETURN)

        time.sleep(5)

        logging.info(f"Pesquisa feita pelo item: {ITEM_NAME}")
        time.sleep(10)

    except Exception as e:
        logging.error(f"Erro ao tentar realizar a pesquisa: {e}")


def info_produto_capturar(driver):
    produtos_list = []
    total_produtos = 0
    produtos_sem_avaliacao = 0
    pagina_atual = 1
    max_paginas = 2

    while True:
        logging.info(f"Capturando produtos da página {pagina_atual}...")

        try:
            # Encontrar os produtos na página
            produtos = driver.find_elements(By.CSS_SELECTOR, "a.sc-fHjqPf.eXlKzg.sc-fvwjDU.gftEET")
            if produtos:
                print(f"Total de {len(produtos)} produtos encontrados na página {pagina_atual}:")
                for idx, produto in enumerate(produtos, start=1):
                    try:
                        produto_nome = produto.find_element(By.CSS_SELECTOR, "h2.sc-cvalOF").text.strip()

                        try:
                            qtd_aval_texto = produto.find_element(By.CSS_SELECTOR, "span.sc-fUkmAC.geJyjP").text.strip()
                            match = re.search(r"\((\d+)\)", qtd_aval_texto)
                            if match:
                                qtd_aval = int(match.group(1))
                            else:
                                qtd_aval = 0

                        except:
                            qtd_aval = 0

                        produto_url = produto.get_attribute("href")

                        produtos_list.append({
                            "PRODUTO": produto_nome,
                            "QTD_AVAL": qtd_aval,
                            "URL": produto_url
                        })

                        total_produtos += 1
                        if qtd_aval == 0:
                            produtos_sem_avaliacao += 1

                        print(f"{idx}. {produto_nome} - Avaliações: {qtd_aval} - URL: {produto_url}")
                    except Exception as e:
                        print(f"Erro ao capturar informações do produto {idx}: {e}")
            else:
                print(f"Nenhum produto encontrado na página {pagina_atual}.")
                break

            """ Foi Feita construção de um laço condicional utilizado para verificação da existencia de novas paginas,
             com base no limite em que fiz o sette acima """
            try:
                # Tente encontrar o botão "Próxima Página"
                next_button = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Go to next page']")
                if next_button and "disabled" not in next_button.get_attribute("class"):
                    next_button.click()
                    pagina_atual += 1
                    time.sleep(15)
                else:
                    print(f"Não há mais páginas para carregar. Fim da coleta.")
                    break
            except Exception as e:
                print(f"Erro ao tentar encontrar o botão da próxima página: {e}")
                break

            if pagina_atual > max_paginas:
                print("Limite de páginas atingido.")
                break

        except Exception as e:
            print(f"Erro ao tentar encontrar os produtos: {e}")
            break

    logging.info(f"Total de produtos capturados: {total_produtos}")
    logging.info(f"Produtos sem avaliação: {produtos_sem_avaliacao}")

    return produtos_list


def envia_email(com_arquivo):
    print("Tentando realizar o envio do e-mail")

    assunto = "Relatório Notebooks"
    corpo = "Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza."

    msg = MIMEMultipart()
    msg['From'] = USER_EMAIL
    msg['To'] = SEND_EMAIL
    msg['Subject'] = assunto

    msg.attach(MIMEText(corpo, 'plain'))

    anexo = com_arquivo
    with open(anexo, 'rb') as anexo_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(anexo)}')
        msg.attach(part)

    # Conectando ao servidor SMTP do Gmail e enviando o e-mail
    try:
        # Usando o servidor SMTP do Gmail
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(USER_EMAIL, USER_PASSWORD)
        texto = msg.as_string()
        servidor.sendmail(USER_EMAIL, SEND_EMAIL, texto)
        servidor.quit()
        print(f"E-mail enviado com sucesso para {SEND_EMAIL}!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def save_to_excel(produtos):
    df = pd.DataFrame(produtos)

    df = df[df['QTD_AVAL'] > 0]

    # Convertendo as avaliações para inteiros para facilitar a comparação
    df['QTD_AVAL'] = pd.to_numeric(df['QTD_AVAL'], errors='coerce')

    # Fazendo a criação e comparação para filtrar produtos melhores e piores
    melhores = df[df['QTD_AVAL'] >= 100]
    piores = df[df['QTD_AVAL'] < 100]

    output_dir = 'Output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    file_path = os.path.join(output_dir, 'Notebooks.xlsx')

    try:
        # Salvando os dados no Excel com múltiplas abas
        with pd.ExcelWriter(file_path) as writer:
            melhores.to_excel(writer, sheet_name="Melhores", index=False)
            piores.to_excel(writer, sheet_name="Piores", index=False)

        logging.info(f"Arquivo salvo em {file_path}")
        return file_path  
        
    except PermissionError:
        print(f"Erro de permissão ao salvar o arquivo {file_path}. Verifique se o arquivo está aberto ou se você tem permissões para escrever.")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

def main():
    driver = inicializando_nav()

    try:
        pesq_item(driver)
        
        produtos = info_produto_capturar(driver)
        
        if produtos:
            file_path = save_to_excel(produtos)
            if file_path:
                envia_email(file_path)
        else:
            logging.warning("Nenhum produto foi encontrado.")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
