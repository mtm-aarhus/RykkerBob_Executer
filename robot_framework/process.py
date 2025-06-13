"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
#Import Packages:
import os
import requests
from datetime import datetime, timedelta, date, timezone
import locale
from dateutil.relativedelta import relativedelta
from dateutil import parser
import uuid

#Import scripts and functions
from GetKmdAcessToken import GetKMDToken
from NovaLogin import GetNovaCookies
import GetCaseInfoAndCheckCaseState
import SendBomEmail
import CheckIfEmailSent 
import SendDigitalPost
from SendSMTPMail import send_email 
import Datastore
import json

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    #   ---- Henter Assets ----
    KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
    KMD_access_token = GetKMDToken(orchestrator_connection)
    
    # ---- Henter datastore ----
    data = Datastore.load_data()
 
   
    # ---- Henter Kø-elementer ----
    queue = json.loads(queue_element.data)
    Sagsnummer = queue.get("caseNumber")
    Taskuuid = queue.get("taskUuid")
    caseUuid = queue.get("caseUuid")
    TaskStartDate = queue.get("taskStartDate")
    TaskDeadline = queue.get("taskDeadline")
    fullName = queue.get("fullName")
    racfId = queue.get("racfId")
    RykkerNummer = queue.get("RykkerNummer")


    orchestrator_connection.log_info(f"Sagsnummer: {Sagsnummer}")
    orchestrator_connection.log_info(f"Taskuuid: {Taskuuid}")
    orchestrator_connection.log_info(f"caseUuid: {caseUuid}")
    orchestrator_connection.log_info(f"TaskStartDate: {TaskStartDate}")
    orchestrator_connection.log_info(f"TaskDeadline: {TaskDeadline}")
    orchestrator_connection.log_info(f"Sagsbehandler: {fullName}")
    orchestrator_connection.log_info(f"RykkerNummer: {RykkerNummer}")

    # ----- Run GetCaseInfoAndCheckCaseState -----
    Arguments_GetCaseInfoAndCheckCaseState = {
        "in_caseUuid": caseUuid,
        "in_KMDNovaURL": KMDNovaURL,
        "in_Sagsnummer": Sagsnummer,
        "in_Token": KMD_access_token,
        "in_fullName": fullName,
        "in_racfId": racfId
    }

    GetCaseInfoAndCheckCaseState_Output_arguments = GetCaseInfoAndCheckCaseState.invoke_GetCaseInfoAndCheckCaseState(Arguments_GetCaseInfoAndCheckCaseState)
    Afgørelsesdato = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_AfgørelsesDato")
    BomCaseId = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_BomCaseId")
    BomCasePhaseCode = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_bomCasePhaseCode")
    BomCaseStateCode = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_bomCaseStateCode")
    BomCaseType = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_BomCaseType")
    BomCaseTypeCode = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_BomCaseTypeCode")
    BomNumber = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_BomNumber")
    CadastralNumber = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_CadastralNumber")
    CaseAddress = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_CaseAddress")
    HouseNumber = GetCaseInfoAndCheckCaseState_Output_arguments.get("Out_HouseNumber")
    IsBomCase = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_IsBomCase")
    Kommunenummer = GetCaseInfoAndCheckCaseState_Output_arguments.get("out_Kommunenummer")
    MissingData = GetCaseInfoAndCheckCaseState_Output_arguments.get("Out_MissingData")
    StreetName = GetCaseInfoAndCheckCaseState_Output_arguments.get("Out_StreetName")
    if MissingData: 
        print("Sagen mangler data, og tilføjes derfor til listen over fejlede sager.")

    else: 
        if IsBomCase: 
            print("Sagen er tilknyttet BOM, sender derfor besked til ansøger og ejer")
            #Predefiune variables: 
            Deadline = None
            EmailText = None
            Title = None
            Description = None
            StrDeadline = None
            DigitalPostSendt = None
            BeskrivelseTilEjer = None
            Tidspunkt = None
            Dato = None

            def case_1():
                AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
                AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
                FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                Dato, Tidspunkt = FormatDate.split(" ")
                StrNewDeadline = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%Y-%m-%dT") + "00:00:00"
                EmailText = ("Vi har endnu ikke modtaget besked om, at dit projekt er sat i gang.\n\n"
                            "Tilladelsen bortfalder, hvis arbejdet ikke er sat i gang inden et år fra tilladelsens dato. "
                            "Dette skal meddeles via Byg og Miljø.\n\n\nMed venlig hilsen\n\nByggeri")
                Title = "1. Rykkerskrivelse - Projektet er ikke påbegyndt"
                Description = "Rykkerskrivelse udført af robot"
                StrDeadline = (AfgørelsesDato + timedelta(days=44*7)).strftime("%Y-%m-%d")
                DigitalPostSendt = True
                return locals()  
            def case_2():
                AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
                AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
                FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                Dato, Tidspunkt = FormatDate.split(" ")
                StrNewDeadline = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%Y-%m-%dT") + "00:00:00"
                EmailText = ("Vi har endnu ikke modtaget besked om, at dit projekt er sat i gang.\n\n"
                "Tilladelsen bortfalder, hvis arbejdet ikke er sat i gang inden et år fra tilladelsens dato. "
                "Dette skal meddeles via Byg og Miljø.\n\n\nMed venlig hilsen\n\nByggeri")
                Title = "2. Rykkerskrivelse - Projektet er ikke påbegyndt"
                Description = "2. Rykkerskrivelse udført af robot"
                StrDeadline = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d")
                DigitalPostSendt = False
                BeskrivelseTilEjer = "Mangler at udsende 1. orientering til ejer"
                return locals()
            def case_3():
                AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
                AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
                FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                Dato, Tidspunkt = FormatDate.split(" ")
                StrNewDeadline = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%Y-%m-%dT") + "00:00:00"
                locale.setlocale(locale.LC_TIME, "da_DK.UTF-8")
                DayDate = AfgørelsesDato.strftime("%d. %B %Y")
                EmailText = (
                    f"Den {DayDate} har du fået byggetilladelse til ovennævnte byggeri.\n\n"
                    f"Tilladelsen blev bl.a. givet på vilkår af, at vi skulle have besked, når byggeriet blev påbegyndt. "
                    f"Da vi ikke har hørt fra dig siden, vil vi gerne have oplyst, om sagen er opgivet.\n\n"
                    f"En byggetilladelse bortfalder, hvis byggearbejdet ikke er påbegyndt inden 1 år fra tilladelsens dato, "
                    f"hvilket du er blevet gjort bekendt med i byggetilladelsen. Denne frist er fastsat i byggelovens § 16, stk. 10.\n\n"
                    f"Der er nu gået et år siden byggetilladelsen blev givet.\n\n"
                    f"Hvis vi ikke hører fra dig inden 14 dage, betragter vi sagen som opgivet, og byggetilladelsen vil blive annulleret.\n\n"
                    f"Der skal søges om en ny byggetilladelse, hvis byggeriet senere ønskes gennemført.\n\n"
                    f"Såfremt byggeriet er påbegyndt, skal du være opmærksom på, at der kan være andre krav i byggetilladelsen, som skal fremsendes.\n\n"
                    f"Med venlig hilsen\n\n"
                    f"Byggeri"
                )
                Title = "3. Rykkerskrivelse - Projektet er ikke påbegyndt"
                Description = "Henlæg - 3. rykker er udført af robot"
                StrDeadline = (datetime.now() + timedelta(days=14)).strftime("%Y-%m-%dT00:00:00+00:00")
                #StrDeadline = (AfgørelsesDato + timedelta(days=14)).strftime("%Y-%m-%d")
                DigitalPostSendt = False
                BeskrivelseTilEjer = "Mangler at udsende 2. orientering til ejer"
                return locals()

            # Dictionary-based switch
            switch = {
                1: case_1,
                2: case_2,
                3: case_3
            }

            assigned_variables = switch.get(RykkerNummer)()  
            if assigned_variables is None:
                raise ValueError(f"Ugyldigt RykkerNummer: {RykkerNummer}")

            # Extract all needed variables with .get() to avoid KeyErrors
            EmailText = assigned_variables.get("EmailText")
            Title = assigned_variables.get("Title")
            Description = assigned_variables.get("Description")
            StrDeadline = assigned_variables.get("StrDeadline")
            DigitalPostSendt = assigned_variables.get("DigitalPostSendt")
            BeskrivelseTilEjer = assigned_variables.get("BeskrivelseTilEjer")
            Tidspunkt = assigned_variables.get("Tidspunkt")
            Dato = assigned_variables.get("Dato")
            StrNewDeadline = assigned_variables.get("StrNewDeadline")

            # ----- Run Send BomEmail -----
            Arguments_SendBomEmail = {
                "in_Sagsnummer": Sagsnummer,
                "in_BomNumber": BomNumber,
                "in_BomCaseId": BomCaseId,
                "in_CaseAddress": CaseAddress,
                "in_BomCaseType": BomCaseType,
                "in_Kommunenummer": Kommunenummer,
                "in_BomCaseTypeCode": BomCaseTypeCode,
                "in_CadastralNumber": CadastralNumber,
                "in_bomCasePhaseCode": BomCasePhaseCode,
                "in_bomCaseStateCode": BomCaseStateCode,
                "in_StreetName": StreetName,
                "in_HouseNumber": HouseNumber,
                "in_Tidspunkt": Tidspunkt,
                "in_Dato": Dato,
                "in_EmailText": EmailText,
                "in_Title": Title
            }


            SendBomEmail_Output_arguments = SendBomEmail.invoke_SendBomEmail(Arguments_SendBomEmail,orchestrator_connection)
            Text = SendBomEmail_Output_arguments.get("out_text")
            print(Text)

                # ----- Run CheckIfEmailSent -----
            Arguments_CheckIfEmailSent = {
                "in_Sagsnummer": Sagsnummer,
                "in_Token": KMD_access_token,
                "in_Title": Title,
                "in_NovaAPIURL": KMDNovaURL,
                
            }
            CheckIfEmailSent_Output_arguments = CheckIfEmailSent.invoke_CheckIfEmailSent(Arguments_CheckIfEmailSent,orchestrator_connection)
            out_DocumentSendt = CheckIfEmailSent_Output_arguments.get("out_DocumentSendt")

    
            if out_DocumentSendt and (RykkerNummer == 2 or RykkerNummer == 3):
                try:
                    # ----- Run SendDigitalPost -----
                    Arguments_SendDigitalPost = {
                        "in_Afgørelsesdato": Afgørelsesdato,
                        "in_Sagsnummer": Sagsnummer,
                        "in_Dato": Dato,
                        "in_NovaAPIURL": KMDNovaURL,
                        "in_RykkerNummer": RykkerNummer,
                        "in_Token": KMD_access_token,
                        "in_BeskrivelseTilEjer": BeskrivelseTilEjer,
                        "in_fullName": fullName,
                        "in_racfId": racfId
                    }
                    
                    SendDigitalPost_Output_arguments = SendDigitalPost.invoke_SendDigitalPost(Arguments_SendDigitalPost,orchestrator_connection)
                    out_DigitaltPostSendt = SendDigitalPost_Output_arguments.get("out_DigitaltPostSendt")
                    print(out_DigitaltPostSendt)
                except Exception as e:
                    print(f"Robotten fejlede: {e}. Mail sendes til udvikler")
                    
                    # Define email details
                    sender = "RykkerBob<rpamtm001@aarhus.dk>" 
                    subject = "Robot fejlede i udsendelse af digital post"
                    body = f"""Kære Udvikler,<br><br>
                    Robotten fejlede. Følgende sagsnummer fik ikke sendt digital post: {Sagsnummer} <br><br>
                    {BeskrivelseTilEjer} via digital post! <br><br>
                    Med venlig hilsen<br><br>
                    Teknik & Miljø<br><br>
                    Digitalisering<br><br>
                    Aarhus Kommune
                    """

                    smtp_server = "smtp.adm.aarhuskommune.dk"   
                    smtp_port = 25               

                    # Call the send_email function
                    send_email(
                        receiver="Gujc@aarhus.dk",
                        sender=sender,
                        subject=subject,
                        body=body,
                        smtp_server=smtp_server,
                        smtp_port=smtp_port,
                        html_body=True
                    )

                    # Opdaterer sagen med nyt materiale
                    try:
            
                        StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")
                        url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"
                        
                        payload = {
                            "common": {
                                "transactionId": str(uuid.uuid4()),
                                "uuid": str(uuid.uuid4())
                            },
                            "caseUuid": caseUuid,
                            "title": "Nyt materiale",
                            "description": f"{BeskrivelseTilEjer}",
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
                            "Authorization": f"Bearer {KMD_access_token}",
                            "Content-Type": "application/json"
                        }

                        response = requests.post(url, json=payload, headers=headers)
                        print(f"API Response: {response.status_code}")

                    except Exception as api_error:
                        print(f"Error occurred during API call: {api_error}")
            else:
                print(f"Der udsendes ikke digital post da rykker nummer er: {RykkerNummer}")
                out_DigitaltPostSendt = False
            

            if out_DocumentSendt:
                #Opdaterer sagsbeskrivelse"
                TaskTitel = "17. Afventer påbegyndelse"
                taskType = "Aktivitet"
                responsibleOrgUnitId = "0c89d77b-c86f-460f-9eaf-d238e4f451ed"
                TaskStartDate = parser.parse(TaskStartDate)
                transformed_Startdate = TaskStartDate.strftime("%Y-%m-%dT00:00:00")
                input_date = datetime.strptime(Dato, "%Y-%m-%d").date()
    

                if RykkerNummer == 1: 
                    if input_date <= date.today():
                        print("Datoen er overskredet eller er i dag - sæt fristen til d.d plus 4 uger")
                        StrDeadline = (datetime.now() + timedelta(days=4*7)).strftime("%Y-%m-%dT00:00:00+00:00")
                    else: 
                        if datetime.strptime(StrDeadline, "%Y-%m-%d").date() <= date.today():
                            print("Fristdato er før i dag - sætter den nye fristdato til d.d + 4 uger")        
                            StrDeadline = (datetime.now() + timedelta(days=4*7)).strftime("%Y-%m-%dT00:00:00+00:00")
                        else: 
                            print("Fristdato er i fremtiden")
                            
                else: 
                    print("Det er ikke første rykker")

                url = f"{KMDNovaURL}/Task/Update?api-version=2.0-Case"
                try:
                    payload = {
                        "common": {
                            "transactionId": str(uuid.uuid4())
                        },
                        "uuid": Taskuuid,
                        "caseUuid": caseUuid,
                        "title": TaskTitel,
                        "description": Description,
                        "caseworker": {
                            "kspIdentity": {
                            "racfId": racfId,
                            "fullName": fullName,
                            }
                        },
                        "responsibleOrgUnitId": responsibleOrgUnitId,
                        "deadline": StrDeadline,
                        "startDate": transformed_Startdate, 
                        "taskType": taskType
                    }
                    headers = {
                        "Authorization": f"Bearer {KMD_access_token}",
                        "Content-Type": "application/json"
                    }
                    response = requests.put(url, json=payload, headers=headers)
                    print(f"API Response: {response.text}")
                    print(f"API staus: {response.status_code}")

                except Exception as api_error:
                    print(f"Error occurred during API call: {api_error}")    

                # Tilføjer sagsnummeret til listen over kørte sager: 
                if Sagsnummer not in data["ListOfProcessedItems"]:
                    data["ListOfProcessedItems"].append(Sagsnummer)
                    Datastore.save_data(data)
                else:
                    print(f"Sagsnummer {Sagsnummer} already processed, skipping append.")

            else:       
                print("Dokumentet blev ikke sendt korrekt - Sagen opdateres ikke")    

                
        else: 

            if RykkerNummer == 1: 
                print("Sagen er ikke tilknyttet BOM - Sender mail og opdaterer sag ift. manuel håndtering")
                StartDato = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S+00:00")
                BeskrivelseIkkeBOM = "Ikke BOM-sag – kræver manuel håndtering"
                url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"
                try:
                    payload = {
                        "common": {
                            "transactionId": str(uuid.uuid4()),
                            "uuid": str(uuid.uuid4())
                        },
                        "caseUuid": caseUuid,
                        "title": "Nyt materiale",
                        "description": BeskrivelseIkkeBOM,
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
                        "Authorization": f"Bearer {KMD_access_token}",
                        "Content-Type": "application/json"
                    }

                    response = requests.post(url, json=payload, headers=headers)
                    print(f"API Response: {response.text}")
                    print(f"API staus: {response.status_code}")

                except Exception as api_error:
                    print(f"Error occurred during API call: {api_error}") 


            #Opdaterer task: 
                Title = "17. Afventer påbegyndelse"
                description = "Ikke BOM sag – Kræver manuel håndtering"
                url = f"{KMDNovaURL}/Task/Update?api-version=2.0-Case"
                TaskStartDate = parser.parse(TaskStartDate)
                transformed_Startdate = TaskStartDate.strftime("%Y-%m-%dT00:00:00")
                TaskDeadline = parser.parse(TaskDeadline)
                transformed_Deadline = TaskDeadline.strftime("%Y-%m-%dT00:00:00+00:00")


                try:
                    payload = {
                        "common": {
                            "transactionId": str(uuid.uuid4())
                        },
                        "uuid": Taskuuid,
                        "caseUuid": caseUuid,
                        "title": Title,
                        "description": description,
                        "caseworker": {
                            "kspIdentity": {
                                "racfId": racfId,
                                "fullName": fullName,
                            }
                        },
                        "deadline": transformed_Deadline,
                        "startDate": TaskStartDate,
                        "taskType": "Aktivitet"
                        }
                    
                    headers = {
                        "Authorization": f"Bearer {KMD_access_token}",
                        "Content-Type": "application/json"
                    }
                    print(payload)
                    response = requests.put(url, json=payload, headers=headers)
                    print(f"API Response: {response.text}")
                    print(f"API staus: {response.status_code}")

                except Exception as api_error:
                    print(f"Error occurred during API call: {api_error}") 
                

                                # Define email details
                    sender = "RykkerBob<rpamtm001@aarhus.dk>" 
                    subject = f"Sagsnummer: {Sagsnummer} mangler BOM-tilknytning"
                    body = f"""Kære sagsbehandler er ikke tilknyttet BOM <br><br>
                    Følgende sagsnummer: {Sagsnummer} <br><br>
                    Derfor kræves det at sagen manuelt håndteres <br><br>
                    Med venlig hilsen<br><br>
                    Teknik & Miljø<br><br>
                    Digitalisering<br><br>
                    Aarhus Kommune
                    """
                    smtp_server = "smtp.adm.aarhuskommune.dk"   
                    smtp_port = 25               

                    # Call the send_email function
                    send_email(
                        receiver="Gujc@aarhus.dk",
                        sender=sender,
                        subject=subject,
                        body=body,
                        smtp_server=smtp_server,
                        smtp_port=smtp_port,
                        html_body=True
                    )


