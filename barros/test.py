from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from selenium.webdriver.common.by import By
import requests
from PIL import Image
from io import BytesIO
import time

API_KEY = '062c4984134e20c6fe8fa73a9c7bc5ac'

def solve_captcha_images(image_paths):
    """Send images to CAPTCHA-solving service and return indices of correct options."""
    try:
        # Upload all images to the CAPTCHA service
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
            else:
                print(f"Error uploading image {image_path}: {response}")

        # Poll for the solutions
        solutions = []
        for captcha_id in captcha_ids:
            for _ in range(30):  # Retry up to 30 times
                time.sleep(5)
                solution_response = requests.get(
                    f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
                ).text
                if 'CAPCHA_NOT_READY' in solution_response:
                    continue
                elif 'OK|' in solution_response:
                    solutions.append(solution_response.split('|')[1])
                    break
                else:
                    print(f"Error fetching solution for captcha ID {captcha_id}: {solution_response}")
        
        # Parse solutions into indices (example: ["1", "3"] -> [0, 2])
        indices = [int(idx) - 1 for idx in solutions if idx.isdigit()]
        return indices

    except Exception as e:
        print(f"Error solving CAPTCHA images: {e}")
        return None



# Load data from Excel file
excel_file = 'data.xlsx'  # Replace with your Excel file's path
df = pd.read_excel(excel_file)

# Ensure the file has columns 'CPF' and 'Password' for login credentials
credentials = df[['CPF', 'Password']]

# Configure Selenium WebDriver
service = Service('/usr/local/bin/chromedriver')  # Update with your ChromeDriver path
driver = webdriver.Chrome(service=service)

try:
    # Navigate to Meu INSS login page
    driver.get("https://www.meu.inss.gov.br/#/login")

    # Wait for the page to load and locate CPF field
    wait = WebDriverWait(driver, 10)


    # Add click in "Entrar com gov.br"
    gov_br = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[3]/div[1]/main[2]/div/div[1]/div/form/button")
    gov_br.click()



    # test change window 

    # Wait for the new window to open (adjust sleep duration if necessary)
    time.sleep(2)  # Replace with WebDriverWait for a better approach if needed

    # Switch to the newly opened window
    new_window_handle = [handle for handle in driver.window_handles if handle != driver.current_window_handle][0]
    driver.switch_to.window(new_window_handle)

    # Optionally, print the title of the new window to confirm the switch
    print("Switched to new window with title:", driver.title)
   


    cpf_field = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/input")

    for index, row in credentials.iterrows():
        cpf = row['CPF']
        password = row['Password']

        # Input CPF
        cpf_field.clear()
        cpf_field.send_keys(cpf)
        # Insert correctly cpf 

    try:
        # Click "Continuar"
        continue_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[1]/div[2]/div/button")
        continue_button.click()

        # Wait for CAPTCHA to load
        time.sleep(4)

        # Locate the CAPTCHA container and find all image elements within it
        captcha_container_xpath = "/html/body"
        captcha_container = driver.find_element(By.XPATH, captcha_container_xpath)

        # the container is finded, but the images inside the grid not
        # the images aren't finded 
        # make the necessary changes for fix it 


        captcha_images = captcha_container.find_elements(By.CLASS_NAME, "task")

        if not captcha_images:
            raise ValueError("No CAPTCHA images found.")

        # Save each CAPTCHA image locally
        image_paths = []
        for idx, img in enumerate(captcha_images):
            image_path = f"captcha_image_{idx}.png"
            with open(image_path, "wb") as file:
                file.write(img.screenshot_as_png)
            image_paths.append(image_path)

        # Send images to CAPTCHA-solving service
        captcha_solution = solve_captcha_images(image_paths)  # Define this function as shown earlier

        if not captcha_solution:
            raise ValueError("Failed to resolve CAPTCHA.")

        # Click the correct images based on the solution
        for idx in captcha_solution:
            if idx < len(captcha_images):  # Ensure valid index
                captcha_images[idx].click()
            else:
                print(f"Index {idx} out of range for CAPTCHA images.")

        # Submit the form after selecting images
        submit_button = driver.find_element(By.XPATH, "/html/body/div[1]/main/form/div[3]/div/button")  # Update XPath
        submit_button.click()

    except Exception as e:
        print(f"An error occurred during CAPTCHA handling: {e}")



        # Input password
        password_field = driver.find_element(By.ID, "txtPassword")
        password_field.clear()
        password_field.send_keys(password)

        # Click login button
        login_button = driver.find_element(By.ID, "btnLogin")
        login_button.click()

        # Add a delay to handle login and potential CAPTCHA
        time.sleep(5)

        # Optional: Check for successful login and perform additional actions
        # Example: Logout or handle errors

        # Navigate back to login page (if testing multiple accounts)
        driver.get("https://www.meu.inss.gov.br/#/login")
        cpf_field = wait.until(EC.presence_of_element_located((By.ID, "txtCPF")))

finally:
    # Close the browser
    driver.quit()
