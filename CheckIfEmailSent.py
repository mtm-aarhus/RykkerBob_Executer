from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import uuid
import json
import requests
from datetime import datetime
from SendSMTPMail import send_email 

def invoke_CheckIfEmailSent(Arguments_CheckIfEmailSent,orchestrator_connection: OrchestratorConnection): 
    #Initialize variables
    Sagsnummer = Arguments_CheckIfEmailSent.get("in_Sagsnummer")
    Token = Arguments_CheckIfEmailSent.get("in_Token")
    in_Title = Arguments_CheckIfEmailSent.get("in_Title")
    KMDNovaURL= Arguments_CheckIfEmailSent.get("in_NovaAPIURL")

    TransactionID = str(uuid.uuid4())

    # Construct the JSON payload
    payload = {
    "common": {"transactionId": TransactionID},
    "paging": {"startRow": 1, "numberOfRows": 1},
    "caseNumber": Sagsnummer,
    "getOutput": {
        "caseworker": True,
        "documentDate": True,
        "title": True,
        "description": True,
        "fileExtension": True,
        "approved": True,
        "acceptReceived": True,
        }
    }


    # Define headers
    headers = {
        "Authorization": f"Bearer {Token}",
        "Content-Type": "application/json"
    }

    # Define the API endpoint
    url = f"{KMDNovaURL}/Document/GetList?api-version=2.0-Case"
    # Make the HTTP request
    try:
        response = requests.put(url, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error for non-2xx responses
        data = response.json()
        print("Success:", response.status_code)
    except Exception as e:
        raise Exception("Kunne ikke hente sagsinfo:", str(e))
    
    #Assigner variable
    Caseworker = "AZMTM01"
    Godkendt = "True"
    DateTime = datetime.now().strftime("%Y-%m-%dT00:00:00")
   
    # Initialize variables
    DateMatches = False
    caseworkerMatches = False
    TitleMatches = False
    approvedMatches = False 

    # Iterate through the "documents"
    for document in data.get("documents", []):
        document_date = document.get("documentDate")

        # Check if documentDate is missing or empty
        if not document_date:
            DateMatches = False
            break
        # Check if the formatted documentDate matches DateTime
        if document_date == DateTime:
            DateMatches = True
            print("DateMatches:", DateMatches)
        else:
            DateMatches = False
            print("DateMatches:", DateMatches)

        racf_id = document.get("caseworker", {}).get("kspIdentity", {}).get("racfId")
    
        if not racf_id:
            caseworkerMatches = False
            break

        # Check if Caseworker matches:
        if racf_id == Caseworker:
            caseworkerMatches = True
            print("Caseworker:", caseworkerMatches)
        else:
            caseworkerMatches = False
            print("Caseworker:", caseworkerMatches)


        Title = document.get("title")
        if not Title:
            TitleMatches = False
            break

        # Check if Caseworker matches:
        if in_Title in Title:
            TitleMatches = True
            print("Title:", TitleMatches)
        else:
            TitleMatches = False
            print("Title:", TitleMatches)

            

        Approved = str(document.get("approved"))
        if not Title:
            approvedMatches = False
            break

        # Check if Caseworker matches:
        if Godkendt == Approved:
            approvedMatches = True
            print("Approved:", approvedMatches)
        else:
            approvedMatches = False
            print("Approved:", approvedMatches)

    if approvedMatches and caseworkerMatches and DateMatches and TitleMatches:
        #orchestrator_connection.log_info("Dokumentet er sendt til modtageren, alt er godt")
        print("Dokumentet er sendt til modtageren, alt er godt")
        out_DocumentSendt = True

    else:
        #orchestrator_connection.log_info("Dokumentet er ikke blevet oploadet korrekt")
        print("Robotten fejlede... Mail sendes til udvikler")
        out_DocumentSendt = False

        # Define email details
        sender = "Rykkerbob<rpamtm001@aarhus.dk>" # Replace with actual sender
        subject = "Robot fejlede i rykkerskrivelse"
        body = f"""Kære Udvikler, <br><br>
        Robotten fejlede. Følgende sagsnummer fik ikke sendt en mail korrekt: {Sagsnummer}<br><br>
        Kontroller venligst at rykkeren blev udsendt korrekt.<br><br>
        Med venlig hilsen<br><br>
        Teknik og Miljø<br><br>
        Digitalisering<br><br>
        Aarhus Kommune.
        """
        smtp_server = "smtp.adm.aarhuskommune.dk"   # Replace with your SMTP server
        smtp_port = 25                    # Replace with your SMTP port

        # Call the send_email function
        send_email(
            receiver="rpamtm001@aarhus.dk",
            sender=sender,
            subject=subject,
            body=body,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )

    return{
        "out_DocumentSendt":  out_DocumentSendt
    }
