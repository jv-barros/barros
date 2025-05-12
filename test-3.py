import logging
import os
import time
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

API_KEY = '062c4984134e20c6fe8fa73a9c7bc5ac'
CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
EXCEL_FILE = 'data.xlsx'
CAPTCHA_SAVE_PATH = "captchas"

os.makedirs(CAPTCHA_SAVE_PATH, exist_ok=True)

def catch_grid_captcha_info(driver, captcha_iframe_xpath):
    """Extract CAPTCHA details and save the image."""
    try:
        # Switch to CAPTCHA iframe
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, captcha_iframe_xpath))
        )
        driver.switch_to.frame(iframe)

        elements_captcha = driver.find_elements(By.XPATH, "/html/body/div/div[1]/div/div/div[2]/div")
        elements = len(elements_captcha)
        logging.info(f"Found {elements} elements in CAPTCHA grid.")

        header_text = None
        for xpath in [
            "/html/body/div/div[1]/div/div/div[1]/div[1]/div[1]/h2/span",
            "/html/body/div/div[1]/div/div/div[1]/div[1]/div[1]/p"
        ]:
            try:
                header_text = driver.find_element(By.XPATH, xpath).text.strip()
                if header_text:
                    break
            except Exception:
                continue
        
        if not header_text:
            raise ValueError("Unable to extract CAPTCHA instructions.")

        if "nenhuma" in header_text or "nenhum" in header_text:
            raise ValueError("Unsupported CAPTCHA format.")

        # Capture a screenshot
        captcha_element = driver.find_element(By.XPATH, "/html/body")
        captcha_image_path = os.path.join(CAPTCHA_SAVE_PATH, "captcha.png")
        captcha_element.screenshot(captcha_image_path)

        if not os.path.exists(captcha_image_path):
            raise FileNotFoundError(f"Captcha image not saved: {captcha_image_path}")

        logging.info(f"CAPTCHA image successfully saved at {captcha_image_path}")

        return header_text, captcha_image_path, "/html/body/div/div[1]/div/div/div[2]", elements

    except Exception as err:
        logging.error(f"Error extracting CAPTCHA info: {err}")
        return None

def solve_captcha_images(image_paths):
    """Upload CAPTCHA images and get solutions."""
    try:
        logging.info("Starting CAPTCHA-solving process.")
        captcha_ids = []

        for image_path in image_paths:
            if not os.path.exists(image_path):
                logging.error(f"Image file not found: {image_path}")
                return None 

            try:
                with open(image_path, "rb") as img_file:
                    response = requests.post(
                        "http://2captcha.com/in.php",
                        files={"file": ("captcha.png", img_file)},
                        data={"key": API_KEY, "method": "post"}
                    ).text

                if response.startswith("OK|"):
                    captcha_id = response.split("|")[1]
                    captcha_ids.append(captcha_id)
                    logging.info(f"Uploaded {image_path}, CAPTCHA ID: {captcha_id}")
                else:
                    logging.error(f"Error uploading {image_path}: {response}")
                    return None  
            except requests.RequestException as e:
                logging.error(f"Request error for {image_path}: {str(e)}")
                return None  

        solutions = []
        for captcha_id in captcha_ids:
            for attempt in range(12):  
                try:
                    time.sleep(5)  
                    solution_response = requests.get(
                        f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
                    ).text

                    if solution_response == "CAPCHA_NOT_READY":
                        logging.info(f"CAPTCHA ID {captcha_id} not ready. Retrying...")
                        continue
                    elif solution_response.startswith("OK|"):
                        indices = solution_response.split("|")[1].split(',')
                        solutions.extend([int(idx) for idx in indices if idx.isdigit()])
                        logging.info(f"Solution for CAPTCHA ID {captcha_id}: {indices}")
                        break  
                    else:
                        logging.error(f"Error fetching solution for CAPTCHA ID {captcha_id}: {solution_response}")
                        return None  
                except requests.RequestException as e:
                    logging.error(f"Request error for CAPTCHA ID {captcha_id}: {str(e)}")
                    return None  

        return solutions if solutions else None

    except Exception as e:
        logging.exception("Unexpected error solving CAPTCHA images.")
        return None

# Selenium setup
logging.info(f"Loading credentials from {EXCEL_FILE}.")
df = pd.read_excel(EXCEL_FILE)
credentials = df[['CPF', 'Password']]

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

try:
    logging.info("Navigating to Meu INSS login page.")
    driver.get("https://www.meu.inss.gov.br/#/login")

    wait = WebDriverWait(driver, 10)
    gov_br = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[3]/div[1]/main[2]/div/div[1]/div/form/button")))
    gov_br.click()
    logging.info("Clicked on 'Entrar com gov.br'.")

    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[1])

    cpf_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/input")))

    for index, row in credentials.iterrows():
        cpf, password = row['CPF'], row['Password']
        logging.info(f"Processing CPF: {cpf}")

        cpf_field.clear()
        cpf_field.send_keys(cpf)

        continue_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/div/button")
        continue_button.click()

        captcha_iframe_xpath = "/html/body/div[3]/div[1]/iframe"

        captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath)
        if not captcha_info:
            continue

        header_text, captcha_image_path, grid_xpath, elements = captcha_info
        captcha_solution = solve_captcha_images([captcha_image_path])
        if not captcha_solution:
            continue

        captcha_positions = [int(digit) for digit in str(captcha_solution[0])]

        for captcha_number in captcha_positions:
            captcha_cells = driver.find_elements(By.XPATH, f"{grid_xpath}/div[{captcha_number}]")
            for cell in captcha_cells:
                cell.click()

        submit_button = driver.find_element(By.XPATH, "/html/body/div/div[3]/div[3]")
        submit_button.click()

        time.sleep(2)
        driver.switch_to.default_content()

        logging.info("Attempting second CAPTCHA resolution.")
        captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath)
        if captcha_info:
            logging.info("Successfully located the second CAPTCHA.")

except Exception as e:
    logging.exception(f"An error occurred: {e}")
finally:
    driver.quit()
