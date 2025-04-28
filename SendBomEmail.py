from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import requests
import json
import time
def invoke_SendBomEmail(Arguments_SendBomEmail,orchestrator_connection: OrchestratorConnection): 
    #Initialize variables
    Sagsnummer = Arguments_SendBomEmail.get("in_Sagsnummer")
    BomNumber = Arguments_SendBomEmail.get("in_BomNumber")
    BomCaseId = Arguments_SendBomEmail.get("in_BomCaseId")
    CaseAddress = Arguments_SendBomEmail.get("in_CaseAddress")
    BomCaseType = Arguments_SendBomEmail.get("in_BomCaseType")
    Kommunenummer = Arguments_SendBomEmail.get("in_Kommunenummer")
    BomCaseTypeCode = Arguments_SendBomEmail.get("in_BomCaseTypeCode")
    CadastralNumber = Arguments_SendBomEmail.get("in_CadastralNumber")
    BomCasePhaseCode = Arguments_SendBomEmail.get("in_bomCasePhaseCode")
    BomCaseStateCode = Arguments_SendBomEmail.get("in_bomCaseStateCode")
    StreetName = Arguments_SendBomEmail.get("in_StreetName")
    HouseNumber = Arguments_SendBomEmail.get("in_HouseNumber")
    Tidspunkt = Arguments_SendBomEmail.get("in_Tidspunkt")
    Dato = Arguments_SendBomEmail.get("in_Dato")
    EmailText = Arguments_SendBomEmail.get("in_EmailText")
    Title = Arguments_SendBomEmail.get("in_Title")


    print(BomCaseType)
    print(Sagsnummer)
    KMD_verification_token = orchestrator_connection.get_credential("Kmd_verification_token")
    Verification_token = KMD_verification_token.password
    
    Kmd_logon_web_session_handler = orchestrator_connection.get_credential("Kmd_logon_web_session_handler")
    Logon_web_session_handler = Kmd_logon_web_session_handler.password
    
    KMD_request_verification_token = orchestrator_connection.get_credential("KMD_request_verification_token")
    Request_verification_token = KMD_request_verification_token.password

    if not StreetName or StreetName is None:
        Address = CaseAddress.split(",")[0].strip()

    else:
        Address = f"{StreetName} {HouseNumber}"

    url = "https://cap-awswlbs-wm3q2021.kmd.dk/KMDNovaESDH/api/ServiceRelayer/BomCase/SendReplyToApplicant"

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://cap-awswlbs-wm3q2021.kmd.dk",
        "RequestVerificationToken": Request_verification_token,
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36 Edg/117.0.2045.31",
        "X-Requested-With": "XMLHttpRequest",
        "sec-ch-ua": '"Microsoft Edge";v="117", "Not;A=Brand";v="8", "Chromium";v="117"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"'
    }

    cookies = {
        "kmdNovaIndstillingerCurrent": "MTM-Byggeri",
        "__RequestVerificationToken_L0tNRE5vdmFFU0RI0": Verification_token,
        "KMDLogonWebSessionHandler": Logon_web_session_handler
    }

    body = {
        "BomCaseId": BomCaseId,
        "CaseNumber": Sagsnummer,
        "PartyDetails": {
            "Name": CaseAddress,
            "Property": {"MunicipalityNumber": Kommunenummer},
            "PropertyMunicipalityNumber": Kommunenummer
        },
        "BomCaseType": {
            "Title": BomCaseType,
            "Code": BomCaseTypeCode
        },
        "BomCaseNumber": BomNumber,
        "CadastralNumber": CadastralNumber,
        "Title": Title,
        "MessageToApplicant": EmailText,
        "NbcInitiative": "Applicant",
        "BomCaseStateCode": BomCaseStateCode,
        "BomPhaseCode": BomCasePhaseCode,
        "Deadline": f"{Dato}T{Tidspunkt}Z",
        "CaseworkerContactInfo": {
            "Email": None,
            "Name": None,
        }
    }
    
    # Convert to JSON string
    body_str = json.dumps(body, ensure_ascii=False)
    # Send request
    response = requests.post(url, headers=headers, cookies=cookies, data=body_str.encode("utf-8"))
    print(f"Netværkskald har status: {response.status_code}")
    time.sleep(10)
    return{
        "out_text": "Bom besked sendt"
        }

