from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke_GetCaseInfoAndCheckCaseState(Arguments_GetCaseInfoAndCheckCaseState): 
    #imports: 
    import requests
    import uuid
    from datetime import datetime
    import Datastore
    
    #Initialize variables
    caseUuid = Arguments_GetCaseInfoAndCheckCaseState.get("in_caseUuid")
    KMDNovaURL = Arguments_GetCaseInfoAndCheckCaseState.get("in_KMDNovaURL")
    Sagsnummer = Arguments_GetCaseInfoAndCheckCaseState.get("in_Sagsnummer")
    Token = Arguments_GetCaseInfoAndCheckCaseState.get("in_Token")
    racfId = Arguments_GetCaseInfoAndCheckCaseState.get("in_racfId")
    fullName = Arguments_GetCaseInfoAndCheckCaseState.get("in_fullName")
    Out_MissingData =False

    # Functioner: 
    def check_bom_case(data):
        for case in data.get("cases", []):
            out_IsBomCase = bool(case.get("buildingCase", {}).get("bomCaseAttributes", {}).get("bomCaseId"))
            print("out_IsBomCase set to:", out_IsBomCase)
            return out_IsBomCase

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
        print("Success:", response.status_code)
    except Exception as e:
        raise Exception("Kunne ikke hente sagsinfo:", str(e))

    out_IsBomCase = check_bom_case(data)
    NumberOfRows = data["pagingInformation"]["numberOfRows"]

    if out_IsBomCase:
        try: 
            out_AfgørelsesDato = data["cases"][0]["buildingCase"]["buildingCaseAttributes"]["applicationStatusDates"]["decisionDate"]
            out_CaseAddress = data["cases"][0]["buildingCase"]["propertyInformation"]["caseAddress"]
            Out_StreetName = data["cases"][0]["buildingCase"]["propertyInformation"]["streetName"]
            Out_HouseNumber = data["cases"][0]["buildingCase"]["propertyInformation"]["houseNumber"]
            out_BomCaseId = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseId"]
            out_BomNumber = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomNumber"]
            out_BomCaseType = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseTypeDescription"]
            out_Kommunenummer = data["cases"][0]["common"]["municipalityNumber"]
            out_BomCaseTypeCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseTypeCode"]
            cadastralLetters = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralLetters"]
            cadastralNumber = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralNumber"]
            out_CadastralNumber = cadastralLetters+str(cadastralNumber)
            out_bomCasePhaseCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCasePhaseCode"]
            out_bomCaseStateCode = data["cases"][0]["buildingCase"]["bomCaseAttributes"]["bomCaseStateCode"]

            print(f"Casedradralnumber: {out_CadastralNumber}")
        except (KeyError, IndexError, TypeError) as e:
            import traceback

            # Get the most recent traceback
            tb = traceback.extract_tb(e.__traceback__)

            # Find the variable name that caused the error
            error_var = None
            for frame in tb:
                if "out_" in frame.line:
                    error_var = frame.line.strip()
                    break

            # Format error message
            if error_var:
                error_message = f"Processen fejlede ved variablen: {error_var} med fejlen: {type(e).__name__} - {e}"
            else:
                error_message = f"Processen fejlede: {type(e).__name__} - {e}"
        
            # If the error message contains "out_", extract relevant part
            if "out_" in error_message:
                try:
                    data = Datastore.load_data()
                    split_str = error_var.split("_")[1]
                    refined_str = split_str.split(" =")[0]
                    TransactionID = str(uuid.uuid4())
                    UuidMissingData = str(uuid.uuid4())
                    Aktivitetsnavn = "Nyt materiale"
                    BeskrivelseManglerData = f"Mangler følgende datakilde: {refined_str}"
                    StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")

                    url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"

                    payload = {
                        "common": {
                            "transactionId": TransactionID,
                            "uuid": UuidMissingData
                        },
                        "caseUuid": caseUuid,
                        "title": Aktivitetsnavn,
                        "description": BeskrivelseManglerData,
                        "caseworker": {
                            "kspIdentity": {
                                "racfId": racfId,
                                "fullName": fullName,
                            }
                        },
                        "startDate": StartDato,
                        "TaskTypeName": "Aktivitet",
                        "statusCode": "S"
                    }
                    headers = {
                        "Authorization": f"Bearer {Token}",
                        "Content-Type": "application/json"
                    }

                    response = requests.post(url, json=payload, headers=headers)
                    print(f"API Response: {response.status_code}")

                except Exception as api_error:
                    print(f"Error occurred during API call: {api_error}")
                Out_MissingData =True
                if Sagsnummer not in data["ListOfFailedCases"]:
                    data["ListOfFailedCases"].append(Sagsnummer)
                    data["ListOfErrorMessages"].append(refined_str)
                    Datastore.save_data(data)
                else:
                    print(f"Sagsnummer {Sagsnummer} is already registered as failed. Skipping append.")

            else:
                Out_MissingData =False
    
    else:
        out_AfgørelsesDato = data["cases"][0]["buildingCase"]["buildingCaseAttributes"]["applicationStatusDates"]["decisionDate"]
        out_CaseAddress = data["cases"][0]["buildingCase"]["propertyInformation"]["caseAddress"]
        out_Kommunenummer = data["cases"][0]["common"]["municipalityNumber"]
        cadastralLetters = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralLetters"]
        cadastralNumber = data["cases"][0]["buildingCase"]["propertyInformation"]["cadastralNumbers"][0]["cadastralNumber"]
        out_CadastralNumber = cadastralLetters+str(cadastralNumber)


    if NumberOfRows > 0: 
        print("Sagen er fortsat åben")

    else: 
        
        raise Exception("Sagen er lukket, fortsætter til næste QueueItem")
    
    result =  {
    "out_AfgørelsesDato": out_AfgørelsesDato,
    "out_CadastralNumber":out_CadastralNumber,
    "out_CaseAddress":out_CaseAddress,
    "out_IsBomCase": out_IsBomCase,
    "out_Kommunenummer": out_Kommunenummer
    }

    if out_IsBomCase:
        result.update({
            "out_BomCaseId": out_BomCaseId,
            "Out_HouseNumber": Out_HouseNumber,
            "out_bomCasePhaseCode": out_bomCasePhaseCode,
            "out_bomCaseStateCode": out_bomCaseStateCode,
            "out_BomCaseType": out_BomCaseType,
            "out_BomCaseTypeCode": out_BomCaseTypeCode,
            "out_BomNumber": out_BomNumber,
            "Out_MissingData":Out_MissingData,
            "Out_StreetName": Out_StreetName
        })

    return result
