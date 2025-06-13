from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def invoke_GetCaseInfoAndCheckCaseState(Arguments_GetCaseInfoAndCheckCaseState): 
    import requests
    import uuid
    from datetime import datetime
    import Datastore
    import traceback

    caseUuid = Arguments_GetCaseInfoAndCheckCaseState.get("in_caseUuid")
    KMDNovaURL = Arguments_GetCaseInfoAndCheckCaseState.get("in_KMDNovaURL")
    Sagsnummer = Arguments_GetCaseInfoAndCheckCaseState.get("in_Sagsnummer")
    Token = Arguments_GetCaseInfoAndCheckCaseState.get("in_Token")
    racfId = Arguments_GetCaseInfoAndCheckCaseState.get("in_racfId")
    fullName = Arguments_GetCaseInfoAndCheckCaseState.get("in_fullName")
    Out_MissingData = False

    out_IsBomCase = False

    def check_bom_case(data):
        for case in data.get("cases", []):
            return bool(case.get("buildingCase", {}).get("bomCaseAttributes", {}).get("bomCaseId"))
        return False

    TransactionID = str(uuid.uuid4())

    payload = {
        "common": {"transactionId": TransactionID},
        "paging": {"startRow": 1, "numberOfRows": 100},
        "caseAttributes": {"userFriendlyCaseNumber": Sagsnummer},
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
            "state": {"progressState": True, "title": True, "activeCode": True, "fromDate": True},
            "caseAttributes": {"userFriendlyCaseNumber": True},
            "caseworker": {
                "kspIdentity": {"novaUserId": True, "racfId": True, "fullName": True}
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

    headers = {
        "Authorization": f"Bearer {Token}",
        "Content-Type": "application/json"
    }

    url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"
    try:
        response = requests.put(url, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()
        print("Success:", response.status_code)
        print(data)
    except Exception as e:
        raise Exception("Kunne ikke hente sagsinfo:", str(e))

    out_IsBomCase = check_bom_case(data)
    NumberOfRows = data["pagingInformation"].get("numberOfRows", 0)

    try:
        case = data["cases"][0]
        building_case = case["buildingCase"]
        property_info = building_case["propertyInformation"]
        bom_attrs = building_case["bomCaseAttributes"]
        status_dates = building_case["buildingCaseAttributes"]["applicationStatusDates"]
        cadastral = property_info["cadastralNumbers"][0]

        if "cadastralNumber" not in cadastral:
            raise Exception("Missing required field: cadastralNumber")

        out_AfgørelsesDato = status_dates["decisionDate"]
        out_CaseAddress = property_info["caseAddress"]
        Out_StreetName = property_info["streetName"]
        Out_HouseNumber = property_info["houseNumber"]
        out_BomCaseId = bom_attrs["bomCaseId"]
        out_BomNumber = bom_attrs["bomNumber"]
        out_BomCaseType = bom_attrs["bomCaseTypeDescription"]
        out_Kommunenummer = case["common"]["municipalityNumber"]
        out_BomCaseTypeCode = bom_attrs["bomCaseTypeCode"]
        cadastralLetters = cadastral["cadastralLetters"]
        cadastralNumber = cadastral["cadastralNumber"]
        out_CadastralNumber = cadastralLetters + str(cadastralNumber)
        out_bomCasePhaseCode = bom_attrs["bomCasePhaseCode"]
        out_bomCaseStateCode = bom_attrs["bomCaseStateCode"]

    except (KeyError, IndexError, TypeError, Exception) as e:
        tb = traceback.extract_tb(e.__traceback__)
        error_var = None
        for frame in tb:
            if "out_" in frame.line:
                error_var = frame.line.strip()
                break

        if error_var:
            error_message = f"Processen fejlede ved variablen: {error_var} med fejlen: {type(e).__name__} - {e}"
        else:
            error_message = f"Processen fejlede: {type(e).__name__} - {e}"

        if "out_" in error_message:
            try:
                data_store = Datastore.load_data()
                split_str = error_var.split("_")[1]
                refined_str = split_str.split(" =")[0]
                TransactionID = str(uuid.uuid4())
                UuidMissingData = str(uuid.uuid4())
                Aktivitetsnavn = "Nyt materiale"
                BeskrivelseManglerData = f"Mangler følgende datakilde: {refined_str}"
                StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")

                url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"
                payload = {
                    "common": {"transactionId": TransactionID, "uuid": UuidMissingData},
                    "caseUuid": caseUuid,
                    "title": Aktivitetsnavn,
                    "description": BeskrivelseManglerData,
                    "caseworker": {"kspIdentity": {"racfId": racfId, "fullName": fullName}},
                    "startDate": StartDato,
                    "TaskTypeName": "Aktivitet",
                    "statusCode": "S"
                }

                response = requests.post(url, json=payload, headers=headers)
                print(f"API Response: {response.status_code}")

                Out_MissingData = True
                if Sagsnummer not in data_store["ListOfFailedCases"]:
                    data_store["ListOfFailedCases"].append(Sagsnummer)
                    data_store["ListOfErrorMessages"].append(refined_str)
                    Datastore.save_data(data_store)
                else:
                    print(f"Sagsnummer {Sagsnummer} is already registered as failed. Skipping append.")

            except Exception as api_error:
                print(f"Error occurred during API call: {api_error}")
        else:
            Out_MissingData = False

    if NumberOfRows > 0: 
        print("Sagen er fortsat åben")
    else: 
        raise Exception("Sagen er lukket, fortsætter til næste QueueItem")

    result = {
        "out_AfgørelsesDato": out_AfgørelsesDato,
        "out_CadastralNumber": out_CadastralNumber,
        "out_CaseAddress": out_CaseAddress,
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
            "Out_MissingData": Out_MissingData,
            "Out_StreetName": Out_StreetName
        })

    return result
