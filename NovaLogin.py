from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import sys
import pathlib
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def GetNovaCookies(orchestrator_connection: OrchestratorConnection):
    orchestrator_connection.log_info("Running GetNovaCookies")
    # Fetch credentials
    KMDNovaRobotLogin = orchestrator_connection.get_credential("KMDNovaRobotLogin")
    NovaUserName = KMDNovaRobotLogin.username
    NovaPassword = KMDNovaRobotLogin.password

    orchestrator_connection.log_info("Initializing selenium")

    app_data_path = os.getenv('LOCALAPPDATA')
    chrome_user_data_path = os.path.join(app_data_path, "Google", "Chrome", "User Data")

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument(f"--user-data-dir={chrome_user_data_path}")
    chrome_options.add_argument("--window-size=1920,900")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("force-device-scale-factor=0.5")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--profile-directory=Default")
    #chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument('--remote-debugging-pipe')
    driver = webdriver.Chrome(options=chrome_options)

    orchestrator_connection.log_info("Opening NOVA")
    print("Creating WebDriver...")
    driver.get("https://cap-wsswlbs-wm3q2021.kmd.dk/KMD.YH.KMDLogonWEB.NET/AspSson.aspx?KmdLogon_sApplCallback=https://cap-awswlbs-wm3q2021.kmd.dk/KMDNovaESDH/forside&KMDLogon_sProtocol=tcpip&KMDLogon_sApplPrefix=--&KMDLogon_sOrigin=ApplAsp&ExtraData=true")
    driver.maximize_window()

    wait = WebDriverWait(driver, 10)

    # Log in
    wait.until(EC.visibility_of_element_located((By.NAME, "UserInfo.Username"))).send_keys(NovaUserName)
    wait.until(EC.visibility_of_element_located((By.NAME, "UserInfo.Password"))).send_keys(NovaPassword)
    wait.until(EC.element_to_be_clickable((By.ID, "logonBtn"))).click()

    orchestrator_connection.log_info("Getting cookies")
    # Get Cookies
    cookies_list = driver.get_cookies()

    def get_cookie_value(cookies, name):
        for cookie in cookies:
            if cookie['name'] == name:
                return cookie['value']
        return None

    out_verification_token = get_cookie_value(cookies_list, "__RequestVerificationToken")
    out_kmd_logon_web_session_handler = get_cookie_value(cookies_list, "KMDLogonWebSessionHandler")

    elements = driver.find_elements(By.XPATH, "/html/body/input[1]")
    out_request_verification_token = None

    if elements:
        element = elements[0]
        out_request_verification_token = element.get_attribute("ncg-request-verification-token")
    else:
        print("Verification token element not found.")

    driver.quit()

    # Post tokens to OpenOrchestrator
    orchestrator_connection.update_credential("Kmd_verification_token", "Verification_token", out_verification_token)
    orchestrator_connection.update_credential("Kmd_logon_web_session_handler", "Logon_web_session_handler", out_kmd_logon_web_session_handler)
    orchestrator_connection.update_credential("KMD_request_verification_token", "Request_verification_token", out_request_verification_token)


