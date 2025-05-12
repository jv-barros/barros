import os
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image
import base64
import io

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
        captcha_rows = driver.find_elements(By.XPATH, "/html/body/div/div/div[2]/div[2]/div/table/tbody/tr[not(contains(@class, 'printonly'))]")
        captcha_columns = driver.find_elements(By.XPATH, "//*[@id='rc-imageselect-target']/table/tbody/tr[1]/td[not(contains(@class, 'printonly'))]")
        rows_captcha_table = len(captcha_rows)
        columns_captcha_table = len(captcha_columns)

        # Extract CAPTCHA instructions
        try:
            header_text = driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[1]/div[1]/div").text
            if not header_text:
                header_text = driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[1]/div[1]/div[2]").text
        except Exception:
            raise ValueError("Unable to extract CAPTCHA instructions.")
        
        header_text = header_text.replace("\n", " ")

        if "nenhuma" in header_text or "nenhum" in header_text or not header_text:
            raise ValueError("Esse modelo de captcha não é possível tratar")

        # Get the CAPTCHA grid element location and size
        captcha_body_element = driver.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div")
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
            "/html/body/div/div/div[2]/div[2]/div/table",
            rows_captcha_table,
            columns_captcha_table,
        ]
    except Exception as err:
        print(f"Erro ao extrair informações do captcha: {err}")
        raise

# the proccess is not complete here 
def main(driver, captcha_iframe_xpath, save_screenshot_path):
    try:
        # Wait for CAPTCHA iframe to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, captcha_iframe_xpath))
        )

        # Extract CAPTCHA information
        captcha_info = catch_grid_captcha_info(driver, captcha_iframe_xpath, save_screenshot_path)
        header_text, captcha_image_path, table_xpath, rows, cols = captcha_info

        print(f"CAPTCHA Header: {header_text}")
        print(f"CAPTCHA Image Path: {captcha_image_path}")
        print(f"CAPTCHA Grid Dimensions: {rows}x{cols}")

        # Example of interacting with CAPTCHA grid
        for row in range(rows):
            for col in range(cols):
                cell_xpath = f"{table_xpath}/tbody/tr[{row + 1}]/td[{col + 1}]"
                cell = driver.find_element(By.XPATH, cell_xpath)
                ActionChains(driver).move_to_element(cell).click().perform()

        # Optionally submit the CAPTCHA
        submit_button_xpath = "//button[contains(text(), 'Verificar')]"
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, submit_button_xpath))
        ).click()

        print("CAPTCHA solved and submitted.")

    except TimeoutException:
        print("Timeout while waiting for CAPTCHA iframe.")
    except Exception as e:
        print(f"An error occurred: {e}")
