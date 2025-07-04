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
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# Load environment variables from .env file
load_dotenv()

# Environment variables
CHROMEDRIVER_PATH = os.path.join(os.getcwd(), "chromedriver.exe")
EXCEL_FILE = os.getenv("EXCEL_FILE")
LOGIN_URL = os.getenv("LOGIN_URL")
USERNAME = os.getenv("USER")
PASSWORD = os.getenv("PASS")
WAIT_TIMEOUT = 10

logging.info(f"Loaded environment variables: USERNAME={USERNAME}, LOGIN_URL={LOGIN_URL}, EXCEL_FILE={EXCEL_FILE}, CHROMEDRIVER_PATH={CHROMEDRIVER_PATH}")

def read_questions_from_excel(file_path):
    """
    Reads questions from an Excel file, skipping the first two rows.
    Returns a DataFrame with standardized column names.
    """
    try:
        df = pd.read_excel(file_path, skiprows=2)
        df.columns = ['risk_color', 'weight', 'range', 'question']
        logging.info(f"Loaded {len(df)} questions from Excel.")
        return df
    except Exception as e:
        logging.error(f"Failed to read Excel file: {e}")
        raise

def setup_driver(chromedriver_path):
    """
    Sets up and returns a Selenium Chrome WebDriver.
    """
    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    logging.info("WebDriver initialized and window maximized.")
    return driver

def login(driver, wait, username, password):
    """
    Logs into the web application using provided credentials.
    """
    logging.info("Navigating to login page.")
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
    logging.info("Login submitted.")

def create_form(driver, wait):
    """
    Clicks the button to create a new form.
    """
    create_form_button = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(text(), 'Criar Formulário')]"
    )))
    create_form_button.click()
    logging.info("Clicked 'Create Form' button.")

def click_button_by_text(driver, text):
    """
    Clicks a button containing the specified text.
    Returns True if successful, False otherwise.
    """
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for button in buttons:
        if text in button.get_attribute("textContent"):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                time.sleep(0.3)
                button.click()
                logging.info(f"Clicked button with text '{text}'.")
                return True
            except Exception as e:
                logging.warning(f"Error clicking button '{text}': {e}")
                return False
    logging.warning(f"Button with text '{text}' not found.")
    return False

def fill_and_add_question(driver, wait, question, weight, is_first_question):
    """
    Fills in the question form and adds the question to the form.
    Handles alternatives for the first question.
    """
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

    if is_first_question:
        # Add two alternatives for the first question
        for _ in range(2):
            if not click_button_by_text(driver, "Adicionar Alternativa"):
                logging.error("Could not click 'Add Alternative' button.")
                return
            wait.until(EC.presence_of_element_located((
                By.XPATH, "//label[contains(text(), 'Texto da Alternativa')]/following-sibling::div//input"
            )))
        time.sleep(0.4)

    alternatives = driver.find_elements(By.XPATH, "//label[contains(text(), 'Texto da Alternativa')]/following-sibling::div//input")
    weights = driver.find_elements(By.XPATH, "//label[contains(text(), 'Peso')]/following-sibling::div//input")

    if len(alternatives) < 2 or len(weights) < 2:
        logging.error("Less than two alternatives available.")
        return

    # Fill alternatives: "Yes" and "No"
    alternatives[-2].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    alternatives[-2].send_keys("Sim")
    weights[-2].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    weights[-2].send_keys("100")

    alternatives[-1].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    alternatives[-1].send_keys("Não")
    weights[-1].send_keys(Keys.CONTROL + "a", Keys.DELETE)
    weights[-1].send_keys("0")

    if not click_button_by_text(driver, "Adicionar Pergunta"):
        logging.error("Could not click 'Add Question' button.")
        return

def main():
    """
    Main execution function.
    Reads questions, logs in, creates form, and adds questions.
    """
    try:
        df = read_questions_from_excel(EXCEL_FILE)
        driver = setup_driver(CHROMEDRIVER_PATH)
        wait = WebDriverWait(driver, WAIT_TIMEOUT)
        is_first_question = True

        try:
            login(driver, wait, USERNAME, PASSWORD)
            create_form(driver, wait)

            for _, row in df.iterrows():
                question = row['question']
                weight = 0
                raw_weight = row['weight']
                if pd.notna(raw_weight):
                    try:
                        weight = int(str(raw_weight).strip().replace('%', '').replace(',', '').strip())
                    except ValueError:
                        logging.warning(f"Invalid weight: {raw_weight} — question skipped.")
                        continue

                fill_and_add_question(driver, wait, question, weight, is_first_question)
                is_first_question = False
                logging.info(f"Question added: {question[:60]}...")
                time.sleep(1.2)

        finally:
            time.sleep(2)
            driver.quit()
            logging.info("WebDriver closed.")

    except Exception as e:
        logging.error(f"Fatal error: {e}")

if __name__ == "__main__":
    main()
