from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import uuid
import json
import requests
from datetime import datetime
from SendSMTPMail import send_email 
import pyodbc  
import pandas as pd
import os
from docx import Document
import getpass

def invoke_SendDigitalPost(Arguments_SendDigitalPost,orchestrator_connection: OrchestratorConnection): 
    Out_NewDate = datetime.now().strftime("%d-%m-%Y")
    def update_word_template(
        EjerNavn,
        EjerAdresse,
        CaseTitle,
        CaseAdress,
        Out_NewDate,
        CaseNumber,
        input_filename
        ):
        # Map placeholders to actual values
        replacements = {
            "<<sagEjer1Navn>>": EjerNavn,
            "<<sagEjer1Adresse>>": EjerAdresse,
            "<<sagsNavn>>": CaseTitle,
            "<<sagsAdresse>>": CaseAdress,
            "klik her for at angive en dato.": Out_NewDate,
            "«Sagsnummer»": CaseNumber,
        }

        # Load the document
        doc = Document(input_filename)

        # Replace placeholders in paragraphs
        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    for run in para.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, val)

        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, val in replacements.items():
                        if key in cell.text:
                            cell.text = cell.text.replace(key, val)

        # Build save path
        username = getpass.getuser()
        save_path = os.path.join("C:\\Users", username, "Downloads", input_filename)

        # Save the updated document
        doc.save(save_path)
        print(f"Document saved to: {save_path}")
    
    
    
    #Initialize variables
    Afgørelsesdato = Arguments_SendDigitalPost.get("in_Afgørelsesdato")
    Beskrivelse = Arguments_SendDigitalPost.get("in_Beskrivelse")
    BeskrivelseTilEjer = Arguments_SendDigitalPost.get("in_BeskrivelseTilEjer")
    Sagsnummer = Arguments_SendDigitalPost.get("in_Sagsnummer")
    caseworkerPersonId = Arguments_SendDigitalPost.get("in_caseworkerPersonId")
    Dato = Arguments_SendDigitalPost.get("in_Dato")
    KMDNovaURL = Arguments_SendDigitalPost.get("in_NovaAPIURL")
    RykkerNummer = Arguments_SendDigitalPost.get("in_RykkerNummer")
    Token = Arguments_SendDigitalPost.get("in_Token")
    
    out_DigitaltPostSendt = False
    AarhusKommuneCVR = "55133018"
    IsCvrAarhusKommune = False

    # Henter yderligere sags information: 

    TransactionID = str(uuid.uuid4())

    Payload = {
        "common": {
            "transactionId": TransactionID
        },
        "paging": {
            "startRow": 1,
            "numberOfRows": 500
        },
        "caseAttributes": {
            "userFriendlyCaseNumber": Sagsnummer
        },
        "caseGetOutput": {
            "caseAttributes": {
                "title": True
            },
            "caseParty": {
                "index": True,
                "identificationType": True,
                "identification": True,
                "partyRole": True,
                "partyRoleName": True,
                "participantRole": True,
                "name": True,
                "participantContactInformation": True,
                "participantRemark": True
            },
            "buildingCase": {
                "propertyInformation": {
                    "caseAddress": True
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
        response = requests.put(url, headers=headers, json=Payload)
        response.raise_for_status()  # Raise an error for non-2xx responses
        data = response.json()
        print("Success:", response.status_code)
    except Exception as e:
        raise Exception("Kunne ikke hente sagsinfo:", str(e))


    # Extract required variables
    for case in data['cases']:
        CaseUuid = case['common']['uuid']
        CaseTitle = case['caseAttributes']['title']
        CaseAddress = case['buildingCase']['propertyInformation']['caseAddress']
        
        identificationType = None
        identification = None

        # Iterate through case parties
        for party in case['caseParties']:
            if "PRI" in party.get('partyRole', '') and "Primær" in party.get('participantRole', ''):
                identificationType = party.get('identificationType')
                identification = party.get('identification')

        # Print the extracted values
        print(f"CaseUuid: {CaseUuid}")
        print(f"CaseTitle: {CaseTitle}")
        print(f"CaseAddress: {CaseAddress}")
        
        if identificationType and identification:
            print(f"IdentificationType: {identificationType}")
            print(f"Identification: {identification}")


    # Get Ejer Oplysninger
    conn_str = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=srvsql29;DATABASE=LOIS;Trusted_Connection=yes"
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Read SQL query from file
    with open('BFENummer.sql', 'r', encoding='utf-8') as file:
        sql_query = file.read()

    sql_query = sql_query.replace("'Ejendomsnummer'", f"'{identification}'")

    # Execute query and fetch results
    cursor.execute(sql_query)
    columns = [column[0] for column in cursor.description]  # Get column names
    rows = cursor.fetchall()

    # Convert to a pandas DataFrame (acting as a DataTable)
    df = pd.DataFrame.from_records(rows, columns=columns)

    # Close database connection
    cursor.close()
    conn.close()

    # Check if DataFrame is empty
    if df.empty:
        raise Exception("Det er hverken et CPR eller CVR-nummer")  # Raise exception if no rows

    print("Der findes en sag med dette BFE-nummer")  # If rows exist

    # Initialize output variables
    CVR = None
    EjerAdresse = None
    EjerNavn = None
    BoolCVR = False
    BoolCPR = False
    BoolBeskyttet = False
    

    # Iterate through DataFrame rows
    for index, row in df.iterrows():
        if "Virksomhed" in str(row["EjerType"]): #Check for EjerType
            print("Det er en virksomhed")
            if "Hovedejer" in str(row["EjerStatusKode_T"]):  # Check if Hovedejer
                CVR = str(row["EjendeVirksomhedCVRnr"])
                EjerAdresse = str(row["EjersAdresseDanmark"])
                EjerNavn = str(row["NavnJusteret"])
                BoolCVR = True

                # Check for Beskyttelse (NULL check)
                if pd.isnull(row["Beskyttelse"]):
                    print("Ejeren har ingen adressebeskyttelse")
                    BoolBeskyttet = False
                else:
                    try:
                        if int(row["Beskyttelse"]) == 1:
                            print("Ejeren har adressebeskyttelse")
                            BoolBeskyttet = True
                        else:
                            print("Ejeren har ingen adressebeskyttelse")
                            BoolBeskyttet = False
                    except ValueError:
                        print("Invalid data format for Beskyttelse")

                print(f"CVR: {CVR}")
                print(f"EjerNavn: {EjerNavn}")
                print(f"EjerAdresse: {EjerAdresse}")
                # Break loop if Hovedejer found
                break

        else:
            if "Person" in str(row["EjerType"]):
                print("Ejer er en person")
                if "Hovedejer" in str(row["EjerStatusKode_T"]):  # Check if Hovedejer
                    CPR = str(row["EjendePersonPersonNR"])
                    EjerAdresse = str(row["EjersAdresseDanmark"])
                    EjerNavn = str(row["NavnJusteret"])
                    BoolCPR = True

                    # Check for Beskyttelse (NULL check)
                    if pd.isnull(row["Beskyttelse"]):
                        print("Ejeren har ingen adressebeskyttelse")
                        BoolBeskyttet = False
                    else:
                        try:
                            if int(row["Beskyttelse"]) == 1:
                                print("Ejeren har adressebeskyttelse")
                                BoolBeskyttet = True
                            else:
                                print("Ejeren har ingen adressebeskyttelse")
                                BoolBeskyttet = False
                        except ValueError:
                            print("Invalid data format for Beskyttelse")

                    print(f"EjerNavn: {EjerNavn}")
                    print(f"EjerAdresse: {EjerAdresse}")
                    # Break loop if Hovedejer found
                    break

    
    if BoolCVR: 
        if AarhusKommuneCVR in CVR: 
            print("Aarhus Kommune er ejer på sagen - Sender fejlmail til Byggeri")
            #Sender error email til byggeri: 
            sender = "Rykkerbob<rpamtm001@aarhus.dk>" # Replace with actual sender
            subject = f"Sagsnummer: {Sagsnummer} har Aarhus Kommune som ejer"
            
            body = f"""Kære sagsbehandler, <br><br>
            Følgende sagsnummer: {Sagsnummer} har Aarhus Kommune som ejer, og mangler at få udsendt: {RykkerNummer}. rykker via Digital Post til ejeren.<br><br>
            Den nye frist dato er: {Dato}<br><br>
            I bedes derfor manuelt udsende rykkeren til ejeren via Digital Post<br><br>
            Med venlig hilsen<br><br>
            Teknik og Miljø<br><br>
            Digitalisering<br><br>
            Aarhus Kommune.
            """
            
            smtp_server = "smtp.adm.aarhuskommune.dk"   # Replace with your SMTP server
            smtp_port = 25                    # Replace with your SMTP port

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

            try:
                TransactionID = str(uuid.uuid4())
                Uuid= str(uuid.uuid4())
                Aktivitetsnavn = "Nyt materiale"
    
                StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")

                url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"

                payload = {
                    "common": {
                        "transactionId": TransactionID,
                        "uuid": Uuid
                    },
                    "caseUuid": CaseUuid,
                    "title": Aktivitetsnavn,
                    "description": f"{BeskrivelseTilEjer} - Aarhus Kommune er ejer",
                    #"caseworkerPersonId": caseworkerPersonId,
                    "caseworker": {
                        "losIdentity": {
                            "novaUnitId": "0c89d77b-c86f-460f-9eaf-d238e4f451ed",
                            "administrativeUnitId": 70528,
                            "fullName": "Plan og Byggeri",
                            "userKey": "2GBYGSAG"
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
    
    if not EjerAdresse or EjerAdresse.strip() == "":
        print("Der findes ingen adresse på ejeren. Derfor bliver der ikke udsendt digital post")

        try:
            TransactionID = str(uuid.uuid4())
            Uuid= str(uuid.uuid4())
            Aktivitetsnavn = "Nyt materiale"

            StartDato = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")

            url = f"{KMDNovaURL}/Task/Import?api-version=2.0-Case"

            payload = {
                "common": {
                    "transactionId": TransactionID,
                    "uuid": Uuid
                },
                "caseUuid": CaseUuid,
                "title": Aktivitetsnavn,
                "description": f"{BeskrivelseTilEjer} - Ejer mangler adresse",
                #"caseworkerPersonId": caseworkerPersonId,
                "caseworker": {
                    "losIdentity": {
                        "novaUnitId": "0c89d77b-c86f-460f-9eaf-d238e4f451ed",
                        "administrativeUnitId": 70528,
                        "fullName": "Plan og Byggeri",
                        "userKey": "2GBYGSAG"
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



    else:
        if IsCvrAarhusKommune == False: 
            try: 
                if RykkerNummer == 2: 
                    print("Sender 2. rykker")
                    update_word_template(
                        EjerNavn,
                        EjerAdresse,
                        CaseTitle,
                        CaseAddress,
                        Out_NewDate,
                        Sagsnummer,
                        input_filename="1. Orientering til ejer vedr. rykker for paabegyndelse.docx"
                    )

                else: 
                    if RykkerNummer == 3: 
                        print("Sender 3. rykker")
                        update_word_template(
                            EjerNavn,
                            EjerAdresse,
                            CaseTitle,
                            CaseAddress,
                            Out_NewDate,
                            Sagsnummer,
                            input_filename="2. Orientering til ejer vedr. rykker for paabegyndelse.docx"
                        )
                    else: 
                        raise ValueError("Det er hverken 2. eller 3. rykker som er blevet oprettet")
                    
            except: 
                print("Sletter lokal fil")
    return {
        "out_DigitaltPostSendt": out_DigitaltPostSendt,
        "Out_NewDate":Out_NewDate
    }