import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

load_dotenv()

LOGIN_URL = os.getenv("URL")
USERNAME = os.getenv("USER")
PASSWORD = os.getenv("PASS")
PLANILHA = os.getenv("EXCEL_FILE")

def start_driver():
    """Initialize and configure the Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    return webdriver.Chrome(options=options)

def login(driver, wait, username, password):
    """Perform login and navigate to the form creation page."""
    driver.get(LOGIN_URL)
    time.sleep(1)

    wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "input")))
    inputs = driver.find_elements(By.TAG_NAME, "input")
    if len(inputs) < 2:
        raise Exception("Login fields not found.")

    inputs[0].clear()
    inputs[0].send_keys(username)
    inputs[1].clear()
    inputs[1].send_keys(password)

    login_button = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(translate(., 'login', 'LOGIN'), 'LOGIN')]"
    )))
    login_button.click()

    create_form_btn = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[normalize-space()='Criar Formulário']"
    )))
    create_form_btn.click()

def fill_form(driver, wait, questions):
    """Fill the form with questions and weights."""
    for question, weight in questions:
        print("🔄 Starting to fill the question...")

        # Validation before proceeding
        if not question or not question.strip() or weight < 1:
            print(f"⚠️ Skipping invalid question: '{question}' (weight={weight})")
            continue

        fill_by_index(driver, wait, 0, question)
        fill_by_index(driver, wait, 1, str(weight))

        # Add alternatives
        for _ in range(2):
            click_button_by_text(driver, "Adicionar Alternativa")
            time.sleep(0.5)

        # Wait for input fields to render
        wait.until(lambda d: len(d.find_elements(By.XPATH, "//input[@type='text']")) >= 4)
        wait.until(lambda d: len(d.find_elements(By.XPATH, "//input[@type='number']")) >= 3)

        alternatives = driver.find_elements(By.XPATH, "//input[@type='text']")[-2:]
        weights = driver.find_elements(By.XPATH, "//input[@type='number']")[-2:]

        if len(alternatives) < 2 or len(weights) < 2:
            print("⚠️ Could not find fields for alternatives.")
            continue
        alternatives[0].clear()
        alternatives[0].send_keys("Sim")
        alternatives[1].clear()
        alternatives[1].send_keys("Não")

        clear_and_type(weights[0], 100)
        clear_and_type(weights[1], 0)
        click_button_by_text(driver, "Adicionar Pergunta")
        print(f"✅ Question added: {question[:60]}...")

def fill_by_index(driver, wait, index, value):
    """
    Fill input field by its index, ensuring the value is set correctly.
    Args:
        driver: Selenium WebDriver instance.
        wait: WebDriverWait instance.
        index: Index of the input field to fill.
        value: Value to input.
    Raises:
        Exception: If the field with the given index is not found.
    """
    print(f"📝 Filling field #{index} with value: '{value}'")
    try:
        fields = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//input[@type='text' or @type='number']")))
    except Exception as e:
        raise Exception(f"⚠️ Unable to locate input fields: {e}")

    if index >= len(fields):
        raise Exception(f"⚠️ Field with index {index} not found.")

    field = fields[index]
    scroll_into_view(driver, field)

    # Try clicking the field with retries
    for attempt in range(3):
        try:
            field.click()
            break
        except Exception:
            print(f"⚠️ Attempt {attempt+1}: click failed. Retrying...")
            time.sleep(0.4)
            scroll_into_view(driver, field)
    else:
        raise Exception(f"⚠️ Could not click field #{index} after retries.")

    # Clear and type the value robustly
    driver.execute_script("arguments[0].value = '';", field)
    time.sleep(0.2)
    field.send_keys(Keys.CONTROL, 'a')
    field.send_keys(Keys.DELETE)
    time.sleep(0.1)
    field.send_keys(str(value))
    time.sleep(0.3)

    # Confirm the value is set, retry if necessary
    for _ in range(3):
        if field.get_attribute("value").strip() == str(value):
            return
        print("🔁 Correcting value in the field, value not confirmed...")
        driver.execute_script("arguments[0].value = '';", field)
        field.send_keys(str(value))
        time.sleep(0.3)

    final_val = field.get_attribute("value").strip()
    if final_val != str(value):
        print(f"⚠️ Final value in field #{index} is '{final_val}', expected: '{value}'")

def scroll_into_view(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    time.sleep(0.3)

def click_button_by_text(driver, text):
    """Click a button by its visible text."""
    try:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//button[contains(., '{text}')]"))
        )
        scroll_into_view(driver, button)
        button.click()
        return True
    except Exception as e:
        print(f"⚠️ Button '{text}' not found. Exception: {e}")
        return False

def clear_and_type(field, value):
    """Força a limpeza com JS + digitação robusta."""
    driver = field.parent  # pega o WebDriver do elemento
    driver.execute_script("arguments[0].value = '';", field)
    time.sleep(0.2)

    field.click()
    field.send_keys(Keys.CONTROL, 'a')
    field.send_keys(Keys.DELETE)
    time.sleep(0.2)
    field.send_keys(str(value))
    time.sleep(0.3)

    # Confirma se o valor está certo, senão tenta de novo
    for _ in range(2):
        current = field.get_attribute("value").strip()
        if current == str(value):
            return
        print(f"🔁 Tentando corrigir valor '{current}' para '{value}'")
        driver.execute_script("arguments[0].value = '';", field)
        time.sleep(0.2)
        field.send_keys(str(value))
        time.sleep(0.2)

    final = field.get_attribute("value").strip()
    if final != str(value):
        print(f"⚠️ Valor final no campo ainda é '{final}', esperado: '{value}'")

def scroll_into_view(driver, element):
    """Scroll the element into view."""
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    time.sleep(0.3)

def load_questions():
    """Load questions and weights from the Excel file."""
    df = pd.read_excel(PLANILHA, skiprows=1)
    questions = []
    for _, row in df.iterrows():
        question = str(row.iloc[3])
        try:
            weight = int(row.iloc[1])
            if weight < 1 or weight > 4:
                raise ValueError
        except Exception:
            print(f"⚠️ Invalid weight: '{row.iloc[1]}', using 1.")
            weight = 1
        questions.append((question, weight))
    return questions

def main():
    """Main execution function."""
    driver = start_driver()
    wait = WebDriverWait(driver, 15)

    try:
        login(driver, wait, USERNAME, PASSWORD)
        questions = load_questions()
        fill_form(driver, wait, questions)
    except Exception as e:
        print("❌ Error:", e)
    finally:
        input("Press Enter to close...")
        driver.quit()

if __name__ == "__main__":
    main()
