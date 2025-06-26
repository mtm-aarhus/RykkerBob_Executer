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
        print("Entered BOM case block")
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
            print(f"CaseAddress: {out_CaseAddress}")
            print(f"cadastralLetters: {cadastralLetters}")
            print(f"Number: {cadastralNumber}")

        except (KeyError, IndexError, TypeError) as e:
            print("❌ Failed to extract BOM case data:", e)
            # Get variable that caused error from traceback
            import traceback
            tb = traceback.extract_tb(e.__traceback__)
            error_var = None
            for frame in tb:
                if "out_" in frame.line:
                    error_var = frame.line.strip()
                    break

            if error_var:
                try:
                    split_str = error_var.split("_")[1]  # e.g. "CadastralNumber = ..."
                    refined_str = split_str.split(" =")[0]  # -> "CadastralNumber"
                except Exception:
                    refined_str = "UkendtData"
            else:
                refined_str = "UkendtData"

            # Fallback API call to report missing data
            try:
                data_store = Datastore.load_data()
                TransactionID = str(uuid.uuid4())
                UuidMissingData = str(uuid.uuid4())
                StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")

                payload = {
                    "common": {
                        "transactionId": TransactionID,
                        "uuid": UuidMissingData
                    },
                    "caseUuid": caseUuid,
                    "title": "Nyt materiale",
                    "description": f"Mangler følgende datakilde: {refined_str}",
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

                task_url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"
                task_response = requests.post(task_url, json=payload, headers=headers)
                print(f"📨 Missing data API call made, status: {task_response.status_code}")

                Out_MissingData = True
                if Sagsnummer not in data_store["ListOfFailedCases"]:
                    data_store["ListOfFailedCases"].append(Sagsnummer)
                    data_store["ListOfErrorMessages"].append(refined_str)
                    Datastore.save_data(data_store)
                else:
                    print(f"Sagsnummer {Sagsnummer} already registered")

            except Exception as api_error:
                print("⚠️ Error during fallback API call:", api_error)

            # 🚫 Stop further execution
            raise Exception(f"Missing data field '{refined_str}' could not be extracted. Process stopped.")
    
    else:
        print("Entered NON-BOM case block")
        try:
            case = data["cases"][0]
            bc = case["buildingCase"]
            pi = bc["propertyInformation"]
            bca = bc["buildingCaseAttributes"]

            out_AfgørelsesDato = bca["applicationStatusDates"]["decisionDate"]
            out_CaseAddress = pi["caseAddress"]
            out_Kommunenummer = case["common"]["municipalityNumber"]
            cadastral = pi["cadastralNumbers"][0]
            cadastralLetters = cadastral["cadastralLetters"]
            cadastralNumber = cadastral["cadastralNumber"]
            out_CadastralNumber = cadastralLetters + str(cadastralNumber)
            print(f"Extracted NON-BOM cadastral number: {out_CadastralNumber}")
        except Exception as e:
            print("Error extracting NON-BOM case data:", e)


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
