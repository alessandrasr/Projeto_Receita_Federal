from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import subprocess
import time
import pandas as pd
import os

# === CONFIGURAÇÕES GLOBAIS ===
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
URL = "https://servicos.receita.fazenda.gov.br/servicos/cpf/consultasituacao/consultapublica.asp"
EXCEL_INPUT = "receita.xlsx"
EXCEL_OUTPUT = "Resultados_receita.xlsx"

# === FUNÇÕES ===

def iniciar_chrome():
    """Inicia o Chrome com depuração remota."""
    command = f'"{CHROME_PATH}" --remote-debugging-port=9222 --user-data-dir="C:\\temp\\selenium"'
    subprocess.Popen(command)
    chrome_options = Options()
    chrome_options.debugger_address = "127.0.0.1:9222"
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def carregar_planilha(caminho):
    """Carrega e trata a planilha de entrada."""
    df = pd.read_excel(caminho)
    df['CPF'] = df['CPF'].astype(str).str.zfill(11)
    df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y')
    return df

def preencher_captcha(driver, wait):
    """Tenta interagir com o hCaptcha até duas vezes."""
    for tentativa in range(2):
        try:
            # Entrar no iframe do captcha
            iframe = wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//iframe[contains(@src, "hcaptcha.com")]')))
            checkbox = wait.until(EC.element_to_be_clickable((By.ID, "checkbox")))
            checkbox.click()
            driver.switch_to.default_content()
            return True  # sucesso
        except Exception as e:
            print(f"⚠️ Tentativa {tentativa + 1} - Erro ao clicar no hCaptcha: {e}")
            driver.switch_to.default_content()
            time.sleep(2)
    return False  # falhou mesmo após 2 tentativas


def preencher_formulario(driver, wait, cpf, data):
    """Preenche o formulário com CPF e data de nascimento."""
    campo_cpf = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtCPF"]')))
    campo_cpf.clear()
    campo_cpf.send_keys(cpf)

    campo_data = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataNascimento"]')))
    campo_data.clear()
    campo_data.send_keys(data)

    time.sleep(2)
    botao_consultar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_submit"]')))
    botao_consultar.click()

def extrair_nome(driver, wait):
    """Extrai o nome da Receita Federal se possível."""
    try:
        nome_elemento = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Nome:')]/b")))
        return nome_elemento.text.strip()
    except:
        print("O elemento nome não existe ou não foi encontrado")
        return "Não encontrado"

def salvar_resultados(resultados, caminho_saida):
    """Salva os resultados em Excel."""
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel(caminho_saida, index=False)
    os.startfile(caminho_saida)
    print("Resultados salvos com sucesso")

def executar_consulta():
    """Fluxo principal do processo."""
    driver = iniciar_chrome()
    wait = WebDriverWait(driver, 30)
    actions = ActionChains(driver)

    input("Depois de abrir a página do Chrome, pressione Enter...")
    driver.get(URL)

    df = carregar_planilha(EXCEL_INPUT)
    resultados = []

    for i, linha in enumerate(df.itertuples(index=False)):
        cpf = linha.CPF
        data = linha.DATA.strftime('%d/%m/%Y')
        print(f"Consultando CPF (Linha {i+1}): {cpf}")


        time.sleep(2)
        preencher_captcha(driver, wait)
        preencher_formulario(driver, wait, cpf, data)

        driver.implicitly_wait(10)
        nome = extrair_nome(driver, wait)

        resultados.append({
            "CPF": cpf,
            "DATA": data,
            "Nome na Receita Federal": nome
        })

        driver.back()
        driver.refresh()
        time.sleep(2)

    salvar_resultados(resultados, EXCEL_OUTPUT)
    driver.quit()

# === EXECUÇÃO PRINCIPAL ===
if __name__ == "__main__":
    executar_consulta()
    