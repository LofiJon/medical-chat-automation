from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import pandas as pd
import time
import os

# Carrega variáveis do .env
load_dotenv()

# Variáveis de ambiente
CHROMEDRIVER_PATH = os.path.join(os.getcwd(), "chromedriver.exe")
EXCEL_FILE = os.getenv("EXCEL_FILE")
LOGIN_URL = os.getenv("LOGIN_URL")
USERNAME = os.getenv("USER")
PASSWORD = os.getenv("PASS")
WAIT_TIMEOUT = 10
print(USERNAME, PASSWORD, LOGIN_URL, EXCEL_FILE, CHROMEDRIVER_PATH)
def read_questions_from_excel(file_path):
    df = pd.read_excel(file_path, skiprows=2)
    df.columns = ['risk_color', 'weight', 'range', 'question']
    return df

def setup_driver(chromedriver_path):
    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    return driver

def login(driver, wait, username, password):
    driver.get(LOGIN_URL)
    username_input = wait.until(EC.presence_of_element_located((
        By.XPATH, "//label[contains(text(), 'Usuário')]/following-sibling::div//input"
    )))
    password_input = wait.until(EC.presence_of_element_located((
        By.XPATH, "//label[contains(text(), 'Senha')]/following-sibling::div//input"
    )))
    username_input.clear()
    username_input.send_keys(username)
    password_input.clear()
    password_input.send_keys(password)
    login_button = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(text(), 'LOGIN')]"
    )))
    login_button.click()

def create_form(driver, wait):
    create_form_button = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(text(), 'Criar Formulário')]"
    )))
    create_form_button.click()

def click_botao_por_texto(driver, texto):
    botoes = driver.find_elements(By.TAG_NAME, "button")
    for botao in botoes:
        if texto in botao.get_attribute("textContent"):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
                time.sleep(0.3)
                botao.click()
                return True
            except Exception as e:
                print(f"⚠️ Erro ao clicar no botão '{texto}': {e}")
                return False
    print(f"❌ Botão com texto '{texto}' não encontrado.")
    return False

def fill_and_add_question(driver, wait, question, weight, primeira_pergunta):
    question_input = wait.until(EC.presence_of_element_located((
        By.XPATH, "//label[contains(text(), 'Pergunta')]/following-sibling::div//input"
    )))
    question_input.send_keys(Keys.CONTROL + "a", Keys.DELETE)
    question_input.send_keys(question)

    weight_input = driver.find_element(
        By.XPATH, "//label[contains(text(), 'Peso da Pergunta')]/following-sibling::div//input"
    )
    weight_input.send_keys(Keys.CONTROL + "a", Keys.DELETE)
    time.sleep(0.2)
    weight_input.send_keys(str(weight))
    weight_input.send_keys(Keys.TAB)

    observation_input = driver.find_element(
        By.XPATH, "//label[contains(text(), 'Observação')]/following-sibling::div//input"
    )
    observation_input.clear()

    if primeira_pergunta:
        for _ in range(2):
            if not click_botao_por_texto(driver, "Adicionar Alternativa"):
                print("❌ Não foi possível clicar no botão 'Adicionar Alternativa'")
                return
            wait.until(EC.presence_of_element_located((
                By.XPATH, "//label[contains(text(), 'Texto da Alternativa')]/following-sibling::div//input"
            )))
        time.sleep(0.4)

    alternativas = driver.find_elements(By.XPATH, "//label[contains(text(), 'Texto da Alternativa')]/following-sibling::div//input")
    pesos = driver.find_elements(By.XPATH, "//label[contains(text(), 'Peso')]/following-sibling::div//input")

    if len(alternativas) < 2 or len(pesos) < 2:
        print("❌ Menos de duas alternativas disponíveis")
        return

    alternativas[-2].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    alternativas[-2].send_keys("Sim")
    pesos[-2].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    pesos[-2].send_keys("100")

    alternativas[-1].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    alternativas[-1].send_keys("Não")
    pesos[-1].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    pesos[-1].send_keys("0")

    if not click_botao_por_texto(driver, "Adicionar Pergunta"):
        print("❌ Não foi possível clicar no botão 'Adicionar Pergunta'")
        return

def main():
    df = read_questions_from_excel(EXCEL_FILE)
    driver = setup_driver(CHROMEDRIVER_PATH)
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    primeira_pergunta = True

    try:
        login(driver, wait, USERNAME, PASSWORD)
        create_form(driver, wait)

        for _, row in df.iterrows():
            question = row['question']
            peso = 0
            raw_peso = row['weight']
            if pd.notna(raw_peso):
                try:
                    peso = int(str(raw_peso).strip().replace('%', '').replace(',', '').strip())
                except ValueError:
                    print(f"⚠️ Peso inválido: {raw_peso} — pergunta pulada.")
                    continue

            fill_and_add_question(driver, wait, question, peso, primeira_pergunta)
            primeira_pergunta = False
            print(f"✅ Pergunta adicionada: {question[:60]}...")
            time.sleep(1.2)

    finally:
        time.sleep(2)
        driver.quit()

if __name__ == "__main__":
    main()
