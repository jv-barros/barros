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

        # Upload each image and retrieve CAPTCHA IDs
        for image_path in image_paths:
            with open(image_path, "rb") as img_file:
                response = requests.post(
                    "http://2captcha.com/in.php",
                    files={"file": ("captcha.png", img_file)},
                    data={"key": API_KEY, "method": "post"}
                ).text
            
            if 'OK|' in response:
                captcha_id = response.split('|')[1]
                captcha_ids.append(captcha_id)
                logging.info(f"Image {image_path} uploaded successfully. CAPTCHA ID: {captcha_id}")
            else:
                logging.error(f"Error uploading image {image_path}: {response}")

        solutions = []
        for captcha_id in captcha_ids:
            wait_time = 5  # Initial wait time
            for _ in range(10):  # Try up to 10 times (50 seconds max)
                time.sleep(wait_time)
                solution_response = requests.get(
                    f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
                ).text
                
                if 'CAPCHA_NOT_READY' in solution_response:
                    logging.info(f"CAPTCHA ID {captcha_id} not ready. Retrying...")
                    wait_time = min(wait_time * 2, 20)  # Exponential backoff (max 20s)
                    continue
                elif 'OK|' in solution_response:
                    indices = solution_response.split('|')[1].split(',')
                    solutions.extend([int(idx) - 1 for idx in indices if idx.isdigit()])
                    logging.info(f"Solution for CAPTCHA ID {captcha_id}: {indices}")
                    break
                else:
                    logging.error(f"Error fetching solution for CAPTCHA ID {captcha_id}: {solution_response}")
                    break

        return solutions if solutions else None
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
            captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath, save_screenshot_path)
            header_text, captcha_image_path, grid_xpath, elements = captcha_info
            logging.info(f"CAPTCHA Instructions: {header_text}")

            captcha_solution = solve_captcha_images([captcha_image_path])
            if not captcha_solution:
                logging.error("Failed to resolve CAPTCHA.")
                continue

            # Find grid items dynamically (assuming images in CAPTCHA)
            captcha_cells = driver.find_elements(By.XPATH, f"{grid_xpath}//img")

            if not captcha_cells:
                logging.error("No CAPTCHA cells found. Cannot proceed.")
                continue

            # Click the correct cells based on the resolved indices
            for idx in captcha_solution:
                if 0 <= idx < len(captcha_cells):
                    driver.execute_script("arguments[0].click();", captcha_cells[idx])
                    logging.info(f"Clicked CAPTCHA cell at index {idx}.")
                else:
                    logging.warning(f"Invalid index {idx} for CAPTCHA grid.")

            # Submit the CAPTCHA selection
            submit_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[3]/div/button")
            submit_button.click()
            logging.info("Submitted CAPTCHA form.")

            # Handle password entry and login
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
