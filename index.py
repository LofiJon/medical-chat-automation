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
    """Fill input field by its index."""
    print(f"📝 Filling field #{index} with value: '{value}'")
    fields = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, "//input[@type='text' or @type='number']")))
    if index >= len(fields):
        raise Exception(f"⚠️ Field with index {index} not found.")
    field = fields[index]
    field.clear()
    field.send_keys(value)

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
    """Clear the field and type the new value."""
    field.click()
    field.send_keys(Keys.CONTROL, "a")
    field.send_keys(Keys.DELETE)
    field.send_keys(str(value))

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
