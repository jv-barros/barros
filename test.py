import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import requests
import os
import base64
import io
from PIL import Image

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

API_KEY = '062c4984134e20c6fe8fa73a9c7bc5ac'

def catch_grid_captcha_info(driver, captcha_iframe_xpath, save_screenshot_path):
    try:
        # Ensure the save directory exists
        os.makedirs(save_screenshot_path, exist_ok=True)

        # Take a screenshot of the full page
        screenshot_base64 = driver.get_screenshot_as_base64()
        screenshot = Image.open(io.BytesIO(base64.b64decode(screenshot_base64)))
        screenshot_path = os.path.join(save_screenshot_path, "iframe_screenshot.png")
        screenshot.save(screenshot_path)

        # Switch to the CAPTCHA iframe
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, captcha_iframe_xpath))
        )
        driver.switch_to.frame(iframe)

        # Locate CAPTCHA elements
        elements_captcha = driver.find_elements(By.XPATH, "/html/body/div/div[1]/div/div/div[2]/div[not(contains(@class, 'printonly'))]")

        elements = len(elements_captcha)

        # Extract CAPTCHA instructions
        try:
            header_text = driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div/div[1]/div[1]/div[1]/h2/span").text
            if not header_text:
                header_text = driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div/div[1]/div[1]/div[1]/p").text
        except Exception:
            raise ValueError("Unable to extract CAPTCHA instructions.")
        
        header_text = header_text.replace("\n", " ")

        if "nenhuma" in header_text or "nenhum" in header_text or not header_text:
            raise ValueError("Esse modelo de captcha não é possível tratar")

        # Get the CAPTCHA grid element location and size
        captcha_body_element = driver.find_element(By.XPATH, "/html/body/div/div[1]/div/div/div[2]")
        location = captcha_body_element.location
        size = captcha_body_element.size

        # Crop the screenshot to include only the CAPTCHA grid
        left = int(location["x"])
        top = int(location["y"])
        right = left + int(size["width"])
        bottom = top + int(size["height"])

        captcha_image = screenshot.crop((left, top, right, bottom))
        captcha_image_path = os.path.join(save_screenshot_path, "captcha_image.png")
        captcha_image.save(captcha_image_path)

        # Return extracted information
        return [
            header_text,
            captcha_image_path,
            "/html/body/div/div[1]/div/div/div[2]",
            elements
            # columns_captcha_table,
        ]
    except Exception as err:
        print(f"Erro ao extrair informações do captcha: {err}")
        raise

def solve_captcha_images(image_paths):
    """Send images to CAPTCHA-solving service and return indices of correct options."""
    try:
        logging.info("Starting CAPTCHA-solving process.")
        captcha_ids = []
        for image_path in image_paths:
            with open(image_path, "rb") as img_file:
                response = requests.post(
                    "http://2captcha.com/in.php",
                    files={"file": ("captcha.png", img_file)},
                    data={"key": API_KEY, "method": "post"}
                ).text
            if 'OK|' in response:
                captcha_ids.append(response.split('|')[1])
                logging.info(f"Image {image_path} uploaded successfully. CAPTCHA ID: {captcha_ids[-1]}")
            else:
                logging.error(f"Error uploading image {image_path}: {response}")

        solutions = []
        for captcha_id in captcha_ids:
            for _ in range(30):
                time.sleep(5)
                solution_response = requests.get(
                    f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
                ).text
                if 'CAPCHA_NOT_READY' in solution_response:
                    logging.info(f"CAPTCHA ID {captcha_id} not ready. Retrying...")
                    continue
                elif 'OK|' in solution_response:
                    solutions.append(solution_response.split('|')[1])
                    logging.info(f"Solution for CAPTCHA ID {captcha_id}: {solutions[-1]}")
                    break
                else:
                    logging.error(f"Error fetching solution for CAPTCHA ID {captcha_id}: {solution_response}")
        indices = [int(idx) - 1 for idx in solutions if idx.isdigit()]
        return indices
    except Exception as e:
        logging.exception("Error solving CAPTCHA images.")
        return None

# Load data from Excel file
excel_file = 'data.xlsx'
logging.info("Loading Excel file: %s", excel_file)
df = pd.read_excel(excel_file)
credentials = df[['CPF', 'Password']]

# Configure Selenium WebDriver
service = Service('/usr/local/bin/chromedriver')
driver = webdriver.Chrome(service=service)

try:
    logging.info("Navigating to Meu INSS login page.")
    driver.get("https://www.meu.inss.gov.br/#/login")

    wait = WebDriverWait(driver, 10)
    gov_br = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[3]/div[1]/main[2]/div/div[1]/div/form/button")
    gov_br.click()
    logging.info("Clicked on 'Entrar com gov.br'.")

    time.sleep(2)
    new_window_handle = [handle for handle in driver.window_handles if handle != driver.current_window_handle][0]
    driver.switch_to.window(new_window_handle)
    logging.info("Switched to new window: %s", driver.title)

    cpf_field = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/input")

    for index, row in credentials.iterrows():
        cpf = row['CPF']
        password = row['Password']
        logging.info("Processing credentials for CPF: %s", cpf)

        cpf_field.clear()
        cpf_field.send_keys(cpf)
        logging.info("Entered CPF.")

        continue_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/div/button")
        continue_button.click()
        logging.info("Clicked 'Continuar'.")

        time.sleep(4)
        save_screenshot_path = "captchas"
        captcha_iframe_xpath = "/html/body/div[3]/div[1]/iframe"

        try:
            # test improve function in here
            captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath, save_screenshot_path)
            header_text, captcha_image_path, grid_xpath, elements = captcha_info
            logging.info(f"CAPTCHA Instructions: {header_text}")

            captcha_solution = solve_captcha_images([captcha_image_path])
            if not captcha_solution:
                logging.error("Failed to resolve CAPTCHA.")
                continue

            # Click resolved CAPTCHA cells
            # change this part for click on correct cells 
            # before, the code use a position of rows and columns to find correct paths 
            # now, we locate the elements in grid and your positions 
            # change the logic of clicking, please 
            for idx in captcha_solution:
                row = idx // elements + 1
                col = idx % elements + 1
                cell_xpath = f"{grid_xpath}/tr[{row}]/td[{col}]"
                driver.find_element(By.XPATH, cell_xpath).click()
                logging.info(f"Clicked CAPTCHA cell at row {row}, column {col}.")

            submit_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[3]/div/button")
            submit_button.click()
            logging.info("Submitted CAPTCHA form.")

            password_field = driver.find_element(By.ID, "txtPassword")
            password_field.clear()
            password_field.send_keys(password)
            logging.info("Entered password.")

            login_button = driver.find_element(By.ID, "btnLogin")
            login_button.click()
            logging.info("Clicked login button.")

            time.sleep(5)
            driver.get("https://www.meu.inss.gov.br/#/login")
            logging.info("Navigated back to login page.")
        except Exception as e:
            logging.error(f"Error handling CAPTCHA: {e}")
finally:
    driver.quit()
    logging.info("Browser closed.")
