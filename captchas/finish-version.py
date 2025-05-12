import logging
import os
import time
import base64
import io
import requests
import pandas as pd
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configure logging for debugging and execution tracking
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

API_KEY = '062c4984134e20c6fe8fa73a9c7bc5ac'  # API key for CAPTCHA-solving service
CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
EXCEL_FILE = 'data.xlsx'
CAPTCHA_SAVE_PATH = "captchas"

# Ensure the CAPTCHA directory exists
os.makedirs(CAPTCHA_SAVE_PATH, exist_ok=True)

def catch_grid_captcha_info(driver, captcha_iframe_xpath):
    """Extracts CAPTCHA details and ensures the image is correctly saved."""
    try:
        # Switch to CAPTCHA iframe
        iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, captcha_iframe_xpath))
        )
        driver.switch_to.frame(iframe)

        # Locate CAPTCHA grid elements
        elements_captcha = driver.find_elements(By.XPATH, "/html/body/div/div[1]/div/div/div[2]/div")
        elements = len(elements_captcha)
        logging.info(f"Found {elements} elements in CAPTCHA grid.")

        # Extract CAPTCHA instructions
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

        header_text = header_text.replace("\n", " ")

        if "nenhuma" in header_text or "nenhum" in header_text:
            raise ValueError("Unsupported CAPTCHA format.")

        # Capture a screenshot of the CAPTCHA iframe
        captcha_element = driver.find_element(By.XPATH, "/html/body")
        captcha_image_path = os.path.join(CAPTCHA_SAVE_PATH, "captcha.png")
        captcha_element.screenshot(captcha_image_path)

        # ✅ Check if the image was actually saved
        if not os.path.exists(captcha_image_path):
            raise FileNotFoundError(f"Captcha image not saved: {captcha_image_path}")

        logging.info(f"CAPTCHA image successfully saved at {captcha_image_path}")

        return header_text, captcha_image_path, "/html/body/div/div[1]/div/div/div[2]", elements

    except Exception as err:
        logging.error(f"Error extracting CAPTCHA info: {err}")
        return None

def solve_captcha_images(image_paths):
    """Uploads CAPTCHA images to a solving service and retrieves the correct selection indices."""
    try:
        logging.info("Starting CAPTCHA-solving process.")
        captcha_ids = []

        for image_path in image_paths:
            if not os.path.exists(image_path):
                logging.error(f"Image file not found: {image_path}")
                return None  # Exit if image is missing

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
            for attempt in range(12):  # Increased retries from 10 to 12 (total 60s)
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


# Load credentials from an Excel file
logging.info(f"Loading credentials from {EXCEL_FILE}.")
df = pd.read_excel(EXCEL_FILE)
credentials = df[['CPF', 'Password']]

# Set up Selenium WebDriver
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
    logging.info(f"Switched to new window: {driver.title}")

    cpf_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/input")))

    for index, row in credentials.iterrows():
        cpf, password = row['CPF'], row['Password']
        logging.info(f"Processing CPF: {cpf}")

        cpf_field.clear()
        cpf_field.send_keys(cpf)

        continue_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/div/button")
        continue_button.click()
        logging.info("Clicked 'Continuar'.")

        time.sleep(10)

        captcha_iframe_xpath = "/html/body/div[3]/div[1]/iframe"
        # while (captcha_iframe_xpath) call process_captcha()

        # Function to process CAPTCHA and return result
        def process_captcha():
            captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath)
            if not captcha_info:
                logging.error("Failed to retrieve CAPTCHA details.")
                return None

            header_text, captcha_image_path, grid_xpath, elements = captcha_info
            logging.info(f"CAPTCHA Instructions: {header_text}")

            captcha_solution = solve_captcha_images([captcha_image_path])
            if not captcha_solution:
                logging.error("Failed to resolve CAPTCHA.")
                return None

            # The captcha_solution is a list containing the positions of the correct cells to click.
            # Convert the integer solution into a list of digits
            captcha_positions = [int(digit) for digit in str(captcha_solution[0])]

            for captcha_number in captcha_positions:
                # Locate the corresponding CAPTCHA cell using the extracted position.
                captcha_cells = driver.find_elements(By.XPATH, f"{grid_xpath}/div[{captcha_number}]")
                # Click on the identified CAPTCHA cell.
                for cell in captcha_cells:
                    cell.click()
                    logging.info(f"Clicked CAPTCHA cell at index {captcha_number}.")

            # Locate and click the submit button to verify the CAPTCHA solution.
            submit_button = driver.find_element(By.XPATH, "/html/body/div/div[3]/div[3]")
            time.sleep(1)
            submit_button.click()
            logging.info("Submitted CAPTCHA.")



        while captcha_iframe_xpath:
            # Solve the first CAPTCHA
            process_captcha()

            # Handle second CAPTCHA, same logic
            try:
                logging.info("Handling second CAPTCHA...")
                driver.switch_to.default_content()
                process_captcha()
            except Exception as e:
                logging.error(f"Error handling second CAPTCHA: {e}")
            

finally:
    logging.info("Test login finally.")
    time.sleep(2)
    driver.quit()
    logging.info("Browser closed.")
