import subprocess
import logging
import time

# Configure logging
logging.basicConfig(filename='script_runner.log', level=logging.INFO, format='%(asctime)s - %(message)s')

def run_script(script_path):
    logging.info(f"Running script: {script_path}")
    process = subprocess.Popen(['python', script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    if process.returncode == 0:
        logging.info(f"Script {script_path} executed successfully.")
    else:
        logging.error(f"Error running script {script_path}: {stderr.decode('utf-8')}")

if __name__ == "__main__":
    # Replace these paths with the paths to your scripts
    script1_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\ctcc_download_upload_files.py"
    # script2_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\ctccXmenu.py"
    script2_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\spqVuca.py"
    script3_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\spqManeZigpay.py"
    script4_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\spqVeiChicoZigpay.py"
    # script5_path = r"C:\\Users\\lucas\\OneDrive\\automacoes\\sendEmail.py"

    # Run first script
    run_script(script1_path)
    time.sleep(100)  # Delay for 10 seconds

    # Run second script if the first script executed successfully
    run_script(script2_path)
    time.sleep(100)  # Delay for 10 seconds

    # Run third script if the second script executed successfully
    run_script(script3_path)
    time.sleep(100)  # Delay for 10 seconds

    # Run fourth script if the third script executed successfully
    run_script(script4_path)
    time.sleep(50)  # Delay for 10 seconds

    # # Run fifth script if the fourth script executed successfully
    # run_script(script5_path)
    # time.sleep(50)  # Delay for 10 seconds

    # # Run sixth script if the fifth script executed successfully
    # run_script(script6_path)

