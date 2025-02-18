from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke_GetCaseInfoAndCheckCaseState(Arguments_GetCaseInfoAndCheckCaseState): 
    #imports: 
    import requests
    import uuid
    
    #Initialize variables
    caseUuid = Arguments_GetCaseInfoAndCheckCaseState.get("in_caseUuid")
    caseworkerPersonId= Arguments_GetCaseInfoAndCheckCaseState.get("in_caseworkerPersonId")
    ListOfErrorMessages = Arguments_GetCaseInfoAndCheckCaseState.get("in_ListOfErrorMessages")
    ListOfFailedCases = Arguments_GetCaseInfoAndCheckCaseState.get("in_ListOfFailedCases")
    KMDNovaURL = Arguments_GetCaseInfoAndCheckCaseState.get("in_KMDNovaURL")
    Sagsnummer = Arguments_GetCaseInfoAndCheckCaseState.get("in_Sagsnummer")
    Token = Arguments_GetCaseInfoAndCheckCaseState.get("in_Token")
    


    # Functioner: 
    def check_bom_case(data):
        # Iterate through cases
        for case in data.get("cases", []):
            building_case = case.get("buildingCase", {})
            bom_case_attributes = building_case.get("bomCaseAttributes")
            
            if bom_case_attributes is None:
                out_IsBomCase = False
                print("Sagen er ikke tilknyttet BOM - derfor assignes out_IsBomCase til:", out_IsBomCase)
            else:
                out_IsBomCase = True
                print("Sagen er tilknyttet BOM - derfor assignes out_IsBomCase til:", out_IsBomCase)
            
            return out_IsBomCase

    # --- Henter sagsinfor ---- 
        # ---- Henter Sagsnummer og Sagsbeskrivelse ---- 
    TransactionID = str(uuid.uuid4())

    # Construct the JSON payload
    payload = {
        "common": {
            "transactionId": TransactionID
        },
        "paging": {
            "startRow": 1,
            "numberOfRows": 100
        },
        "caseAttributes": {
            "userFriendlyCaseNumber": Sagsnummer
        },
        "states": {
            "states": [
                {"progressState": "Udfoert"},
                {"progressState": "Bestilt"},
                {"progressState": "Afgjort"},
                {"progressState": "Oplyst"},
                {"progressState": "Opstaaet"}
            ]
        },
        "caseGetOutput": {
            "state": {
                "progressState": True,
                "title": True,
                "activeCode": True,
                "fromDate": True
            },
            "caseAttributes": {
                "userFriendlyCaseNumber": True
            },
            "caseworker": {
                "kspIdentity": {
                    "novaUserId": True,
                    "racfId": True,
                    "fullName": True
                }
            },
            "buildingCase": {
                "buildingCaseAttributes": {
                    "applicationStatusDates": {
                        "decisionDate": True,
                        "applicationCompletedDate": True,
                        "closeDate": True,
                        "closingReason": True
                    }
                },
                "propertyInformation": {
                    "caseAddress": True,
                    "streetName": True,
                    "houseNumber": True,
                    "esrPropertyNumber": True,
                    "municipalityNumber": True,
                    "cadastralNumbers": {
                        "cadastralLetters": True,
                        "cadastralNumber": True
                    }
                },
                "bomCaseAttributes": {
                    "bomCaseId": True,
                    "bomNumber": True,
                    "bomCaseTypeDescription": True,
                    "bomCaseTypeCode": True,
                    "bomCaseStateCode": True,
                    "bomCasePhaseCode": True
                }
            }
        }
    }

    # Define headers
    headers = {
        "Authorization": f"Bearer {Token}",
        "Content-Type": "application/json"
    }

    # Define the API endpoint
    url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"
    # Make the HTTP request
    try:
        response = requests.put(url, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error for non-2xx responses
        data = response.json()
        print("Success:", response.status_code, response.text)
    except Exception as e:
        raise Exception("Kunne ikke hente sagsinfo:", str(e))

    out_IsBomCase = check_bom_case(data)

    if out_IsBomCase:
        out_AfgørelsesDato = data["cases"][0]["buildingCase"]["buildingCaseAttributes"]["applicationStatusDates"]["decisionDate"]
        out_CaseAddress = data["cases"][0]["buildingCase"]["propertyInformation"]["caseAddress"]
        Out_StreetName = data["cases"][0]["buildingCase"]["propertyInformation"]["streetName"].ToString
        Out_HouseNumber = data["cases"][0]["buildingCase"]["propertyInformation"]["houseNumber"].ToString
        out_BomCaseId = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseId"].ToString
        out_BomNumber = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomNumber"].ToString
        out_BomCaseType = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseTypeDescription"].ToString
        out_Kommunenummer = data["cases"][0]["common"]["municipalityNumber"].ToString
        out_BomCaseTypeCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseTypeCode"].ToString
        cadastralLetters = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralLetters"].ToString
        cadastralNumber = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralNumber"].ToString
        out_CadastralNumber = cadastralLetters+cadastralNumber
        out_bomCasePhaseCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCasePhaseCode"].ToString
        out_bomCaseStateCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseStateCode"].ToString
