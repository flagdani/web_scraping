# cndestadual.py

import os
import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import logging
import pytesseract
import fitz
import shutil
from PIL import Image
import re
import random

# Configuração do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'path/to/tesseract.exe' # Tesseract para OCR
os.environ['TESSDATA_PREFIX'] = r'path/to/tessdata'  # Tessdata

download_directory = os.path.join(os.path.expanduser("~"), "Downloads")

negativas_dir = r'path/to/cndestadual/pr_negativas'
positivas_efeito_negativas_dir = r'path/to/cndestadual/pr_positivas_efeito_negativas'

def configurar_logging():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def carregar_dados(planilha):
    df = pd.read_excel(planilha, dtype={'CNPJ': str})  # Garantir que os CNPJs sejam lidos como strings
    if 'CND_ESTADUAL' not in df.columns:
        df['CND_ESTADUAL'] = ''
    else:
        df['CND_ESTADUAL'] = df['CND_ESTADUAL'].astype(str)
    return df

def iniciar_navegador():
    options = uc.ChromeOptions()
    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.default_directory": download_directory,
        "safebrowsing.enabled": "false",
        "profile.default_content_setting_values.automatic_downloads": 1,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.disable_download_protection": True
    }
    options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(options=options)
    return driver

class ProcessadorPDF:
    @staticmethod
    def pixmap_para_pil(pix):
        mode = "RGBA" if pix.alpha else "RGB"
        size = (pix.width, pix.height)
        return Image.frombytes(mode, size, pix.samples)

    def extrair_texto_de_pdf(self, caminho_pdf):
        texto = ""
        with fitz.open(caminho_pdf) as doc:
            for page_number, page in enumerate(doc):
                pix = page.get_pixmap()
                img = self.pixmap_para_pil(pix)
                texto += pytesseract.image_to_string(img, lang='por')
        return texto

    def verificar_status_pdf(self, caminho_pdf):
        texto = self.extrair_texto_de_pdf(caminho_pdf).upper()
        match_com_efeitos = re.search(r"COM\s*(?:E|F|E?T?F?E?I?T?O?S?)?\s*DE\s*NEGATIVA", texto)

        if "POSITIVA" in texto:
            if match_com_efeitos:
                status = "Positiva com Efeitos de Negativa"
            else:
                status = "Positiva"
        elif "NEGATIVA" in texto:
            status = "Negativa"
        else:
            status = "Texto não reconhecido"

        return status

# Função para acessar o site e clicar em "Emitir"
def acessar_site(driver):
    logging.info("Acessando site inicial")
    driver.get("https://www.fazenda.pr.gov.br/servicos/Mais-buscados/Certidoes/Emitir-Certidao-Negativa-Receita-Estadual-kZrX5gol")

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        logging.info("Página inicial carregada com sucesso.")

        for _ in range(3):  # Tentar clicar no botão até 3 vezes
            try:
                emitir_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH,
                                                "//a[contains(@class, 'btn') and contains(@class, 'btn-primary') and text()='Emitir']"))
                )
                emitir_button.click()
                logging.info("Botão 'Emitir' clicado com sucesso.")
                break
            except (TimeoutException, NoSuchElementException) as e:
                logging.error(f"Falha ao clicar no botão 'Emitir'. Tentando novamente... - {e}")
                time.sleep(5)
                driver.refresh()  # Recarregar a página e tentar novamente
                continue

        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "EmissaoCnpj")))
        logging.info("Nova aba aberta com sucesso.")

    except Exception as e:
        logging.error(f"Erro ao acessar o site inicial - {e}")
        raise

# Função para preencher o formulário e submeter
def preencher_formulario(driver, cnpj):
    try:
        logging.info(f"Preenchendo formulário para o CNPJ: {cnpj}")

        time.sleep(random.uniform(2, 5))  # Aguardar um pouco mais para garantir que todos os scripts do site sejam carregados
        cnpj_input = driver.find_element(By.ID, "EmissaoCnpj")
        cnpj_input.clear()
        cnpj_input.send_keys(cnpj)
        time.sleep(random.uniform(1, 3))  # Aguardar um pouco antes de clicar no botão

        submit_button = driver.find_element(By.ID, "submitBtn")
        submit_button.click()

    except Exception as e:
        logging.error(f"Erro ao preencher o formulário para o CNPJ: {cnpj} - {e}")
        raise

# Função para processar o resultado
def processar_resultado(driver, cnpj, row_index):
    processador_pdf = ProcessadorPDF()
    tentativas = 0
    while tentativas < 3:
        try:
            success_alert = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "alert-success")))
            pdf_link = success_alert.find_element(By.LINK_TEXT, "CLIQUE AQUI")
            pdf_link.click()
            time.sleep(random.uniform(15, 20))  # Tempo aumentado para garantir o download

            pdf_filename = encontrar_arquivo_pdf(download_directory)
            if pdf_filename:
                status = processador_pdf.verificar_status_pdf(pdf_filename)
                cnpj_extraido = extrair_cnpj_do_texto(pdf_filename)

                if cnpj_extraido and cnpj_extraido != cnpj:
                    logging.warning(
                        f"CNPJ extraído ({cnpj_extraido}) é diferente do esperado ({cnpj}). Usando CNPJ extraído.")

                cnpj_usado_para_nome = cnpj_extraido if cnpj_extraido else cnpj

                if status == 'Negativa':
                    novo_nome = os.path.join(negativas_dir, f"{cnpj_usado_para_nome}_negativa.pdf")
                    resultado = "OK, Negativa"
                elif status == 'Positiva com Efeitos de Negativa':
                    novo_nome = os.path.join(positivas_efeito_negativas_dir,
                                             f"{cnpj_usado_para_nome}_positiva_efeito_negativa.pdf")
                    resultado = "OK, Positivas efeito negativas"
                else:
                    novo_nome = os.path.join(download_directory, f"{cnpj_usado_para_nome}_texto_nao_reconhecido.pdf")
                    resultado = "Texto não reconhecido"

                if os.path.exists(novo_nome):
                    os.remove(novo_nome)
                shutil.move(pdf_filename, novo_nome)
                logging.info(f"Arquivo movido e renomeado para: {novo_nome}")
                return resultado
            else:
                logging.error(f"Erro ao localizar o PDF para o CNPJ: {cnpj}")
                return "Erro no download do PDF"
        except (NoSuchElementException, TimeoutException, WebDriverException) as e:
            tentativas += 1
            logging.error(f"Tentativa {tentativas} falhou para o CNPJ: {cnpj} - {e}. Recarregando a página.")
            driver.refresh()
            time.sleep(random.uniform(3, 5))  # Aguarde alguns segundos antes de tentar novamente

    try:
        error_alert = driver.find_element(By.CLASS_NAME, "alert-danger")
        mensagem_erro = error_alert.text.strip()
        if "As informações disponíveis não permitem a emissão de Certidão Automática para o requerente." in mensagem_erro:
            logging.warning(f"Erro na emissão para o CNPJ: {cnpj} - {mensagem_erro}")
            return "S/CND"
    except NoSuchElementException:
        logging.error(f"Erro desconhecido para o CNPJ: {cnpj} após várias tentativas")

    return "Erro desconhecido após várias tentativas"

# Função para encontrar o arquivo PDF recém-baixado
def encontrar_arquivo_pdf(diretorio):
    # Definir um tempo limite total para a busca
    tempo_limite = 30
    tempo_passado = 0
    while tempo_passado < tempo_limite:
        arquivos = os.listdir(diretorio)
        arquivos_pdf = [os.path.join(diretorio, arquivo) for arquivo in arquivos if arquivo.endswith(".pdf")]
        if arquivos_pdf:
            # Encontrar o arquivo PDF mais recente na pasta
            pdf_mais_recente = max(arquivos_pdf, key=os.path.getctime)
            return pdf_mais_recente
        time.sleep(1)
        tempo_passado += 1
    return None

# Função para extrair o CNPJ do texto extraído do PDF
def extrair_cnpj_do_texto(caminho_pdf):
    processador = ProcessadorPDF()
    texto = processador.extrair_texto_de_pdf(caminho_pdf).upper()
    match = re.search(r'CNPJMF[:\s]*([0-9./-]+)', texto)
    if match:
        cnpj = re.sub(r'\D', '', match.group(1))
        return cnpj if len(cnpj) == 14 else None
    return None

# Função principal
def main():
    configurar_logging()
    planilha = r'path/to/excel_file.xlsx'  # Caminho da planilha de consulta e salvamento
    dirs = {
        'downloads': os.path.join(os.path.expanduser('~'), 'Downloads'),
        'pdfs': r'path/to/cndcuritiba/pdfs',
        'negativos': r'path/to/cndcuritiba/negativos',
        'positivos': r'path/to/cndcuritiba/positivos',
        'positivas_efeito_negativas': r'path/to/cndcuritiba/positivas_efeito_negativas'
    }
    for dir_path in dirs.values():
        os.makedirs(dir_path, exist_ok=True)

    cnpjs = carregar_cnpjs(planilha)

    # solucionador = SolucionadorRecaptcha('your_captcha_api_key')  # Comentado: Integre uma solução de CAPTCHA
    solucionador = None  # Placeholder
    processador_pdf = ProcessadorPDFCuritiba()
    gerenciador_planilha = GerenciadorPlanilhaCuritiba(planilha)
    navegador_web = NavegadorWebCuritiba(cnpjs, dirs, solucionador, processador_pdf, gerenciador_planilha)

    chrome_options = navegador_web.configurar_chrome()
    driver = uc.Chrome(options=chrome_options)

    for cnpj in cnpjs:
        try:
            navegador_web.acessar_site(driver, cnpj)
        except Exception as e:
            logging.error(f"Ocorreu um erro ao processar o CNPJ {cnpj} após tentativa com tempos aumentados: {e}")
            gerenciador_planilha.atualizar_planilha(cnpj, f"Erro: {e}")
        time.sleep(30)

    driver.quit()
    gerenciador_planilha.salvar_planilha()


if __name__ == "__main__":
    main()
