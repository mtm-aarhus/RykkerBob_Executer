
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
def GetNovaCookies(orchestrator_connection: OrchestratorConnection):
    KMDNovaRobotLogin = orchestrator_connection.get_credential("KMDNovaRobotLogin")
    NovaUserName = KMDNovaRobotLogin.username
    NovaPassword = KMDNovaRobotLogin.password
    import os
    import sys
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    import pathlib

    def get_selenium_manager_path():
        try:
            selenium_path = pathlib.Path(sys.modules['selenium'].__file__).parent
            selenium_folder_index = selenium_path.parts.index('selenium')
            if selenium_folder_index != -1 and selenium_folder_index < len(selenium_path.parts) - 1:
                version_folder = selenium_path.parts[selenium_folder_index + 1]
                version_folder_path = pathlib.Path(*selenium_path.parts[:selenium_folder_index + 2])
                for file in version_folder_path.rglob("*manager*.exe"):
                    os.environ["SE_MANAGER_PATH"] = str(file)
                    return file
        except Exception as e:
            print(f"Error finding Selenium Manager Path: {e}")
            return None

    def main():
        try:
            get_selenium_manager_path()
            
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
            chrome_options.add_argument("--remote-debugging-port=9222")
            
            driver = webdriver.Chrome(options=chrome_options)
            
            print("Creating WebDriver...")
            print("Navigating to URL...")
            driver.get("https://cap-wsswlbs-wm3q2021.kmd.dk/KMD.YH.KMDLogonWEB.NET/AspSson.aspx?KmdLogon_sApplCallback=https://cap-awswlbs-wm3q2021.kmd.dk/KMDNovaESDH/forside&KMDLogon_sProtocol=tcpip&KMDLogon_sApplPrefix=--&KMDLogon_sOrigin=ApplAsp&ExtraData=true")
            print("Maximizing window...")
            driver.maximize_window()
            
            wait = WebDriverWait(driver, 10)
            
            # Log in
            wait.until(EC.visibility_of_element_located((By.NAME, "UserInfo.Username"))).send_keys(NovaUserName)
            wait.until(EC.visibility_of_element_located((By.NAME, "UserInfo.Password"))).send_keys(NovaPassword)
            wait.until(EC.element_to_be_clickable((By.ID, "logonBtn"))).click()
            print("Der er logget ind på hjemmesiden")
            
            # Get Cookies
            cookies_list = driver.get_cookies()
            
            def get_cookie_value(cookies, name):
                for cookie in cookies:
                    if cookie['name'] == name:
                        return cookie['value']
                return None
            
            out_verification_token = get_cookie_value(cookies_list, "__RequestVerificationToken_L0tNRE5vdmFFU0RI0")
            out_kmd_logon_web_session_handler = get_cookie_value(cookies_list, "KMDLogonWebSessionHandler")
            
            
            elements = driver.find_elements(By.XPATH, "/html/body/input[1]")
            
            out_request_verification_token = None
            
            if elements:
                element = elements[0]
                out_request_verification_token = element.get_attribute("ncg-request-verification-token")
            else:
                print("Element not found")
            
            driver.quit()    

            #Post tokes to OpenOrchestrator:
            orchestrator_connection.update_credential("Kmd_verification_token", "Verification_token", out_verification_token)
            orchestrator_connection.update_credential("Kmd_logon_web_session_handler", "Logon_web_session_handler", out_kmd_logon_web_session_handler)
            orchestrator_connection.update_credential("KMD_request_verification_token", "Request_verification_token", out_request_verification_token)

        except Exception as e:
            print(e)
            raise

    main()