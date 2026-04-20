from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def wait_for_cookie(driver, cookie_name, timeout=10, poll_frequency=0.5):
    start_time = time.time()
    while time.time() - start_time < timeout:
        cookie = driver.get_cookie(cookie_name)
        if cookie and "value" in cookie:
            return cookie["value"]
        time.sleep(poll_frequency)
    raise TimeoutException(f"Cookie '{cookie_name}' not found within {timeout} seconds")


def GetNovaCookies(orchestrator_connection: OrchestratorConnection):
    orchestrator_connection.log_info("Running GetNovaCookies")

    KMDNovaRobotLogin = orchestrator_connection.get_credential("KMDNovaRobotLogin")
    NovaUserName = KMDNovaRobotLogin.username
    NovaPassword = KMDNovaRobotLogin.password

    orchestrator_connection.log_info("Initializing selenium")

    app_data_path = os.getenv("LOCALAPPDATA")
    chrome_user_data_path = os.path.join(app_data_path, "Google", "Chrome", "User Data")

    chrome_options = Options()
    # chrome_options.add_argument("--headless=new")
    chrome_options.add_argument(f"--user-data-dir={chrome_user_data_path}")
    chrome_options.add_argument("--window-size=1920,900")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("force-device-scale-factor=0.5")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--profile-directory=Default")
    # chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--remote-debugging-pipe")

    driver = webdriver.Chrome(options=chrome_options)

    try:
        orchestrator_connection.log_info("Opening NOVA")
        driver.get("https://kmdnova.dk")
        driver.maximize_window()

        wait = WebDriverWait(driver, 10)

        try:
            orchestrator_connection.log_info("STEP 1: Waiting for username field")
            wait.until(
                EC.visibility_of_element_located((By.NAME, "UserInfo.Username"))
            ).send_keys(NovaUserName)
            orchestrator_connection.log_info("STEP 1 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 1 (username field): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 2: Waiting for password field")
            wait.until(
                EC.visibility_of_element_located((By.NAME, "UserInfo.Password"))
            ).send_keys(NovaPassword)
            orchestrator_connection.log_info("STEP 2 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 2 (password field): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 3: Waiting for login button")
            wait.until(
                EC.element_to_be_clickable((By.ID, "logonBtn"))
            ).click()
            orchestrator_connection.log_info("STEP 3 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 3 (login button): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 4: Getting all cookies")
            cookies_list = driver.get_cookies()
            orchestrator_connection.log_info(f"STEP 4 OK - cookies: {[c['name'] for c in cookies_list]}")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 4 (get_cookies): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 5: Waiting for __RequestVerificationToken")
            out_verification_token = wait_for_cookie(driver, "__RequestVerificationToken", timeout=60)
            orchestrator_connection.log_info("STEP 5 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 5 (__RequestVerificationToken): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 6: Waiting for KMDLogonWebSessionHandler")
            out_kmd_logon_web_session_handler = wait_for_cookie(driver, "KMDLogonWebSessionHandler", timeout=60)
            orchestrator_connection.log_info("STEP 6 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 6 (KMDLogonWebSessionHandler): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 7: Finding token element")
            elements = driver.find_elements(By.XPATH, "/html/body/input[1]")
            out_request_verification_token = None

            if elements:
                element = elements[0]
                out_request_verification_token = element.get_attribute("ncg-request-verification-token")
                orchestrator_connection.log_info("STEP 7 OK - token element found")
            else:
                orchestrator_connection.log_error("FAILED at STEP 7: Verification token element not found")
                raise Exception("Verification token element not found")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 7 (DOM token): {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 8: Updating credential Kmd_verification_token")
            orchestrator_connection.update_credential(
                "Kmd_verification_token",
                "Verification_token",
                out_verification_token
            )
            orchestrator_connection.log_info("STEP 8 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 8: {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 9: Updating credential Kmd_logon_web_session_handler")
            orchestrator_connection.update_credential(
                "Kmd_logon_web_session_handler",
                "Logon_web_session_handler",
                out_kmd_logon_web_session_handler
            )
            orchestrator_connection.log_info("STEP 9 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 9: {str(e)}")
            raise

        try:
            orchestrator_connection.log_info("STEP 10: Updating credential KMD_request_verification_token")
            orchestrator_connection.update_credential(
                "KMD_request_verification_token",
                "Request_verification_token",
                out_request_verification_token
            )
            orchestrator_connection.log_info("STEP 10 OK")
        except Exception as e:
            orchestrator_connection.log_error(f"FAILED at STEP 10: {str(e)}")
            raise

    finally:
        driver.quit()
