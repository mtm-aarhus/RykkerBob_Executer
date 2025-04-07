#Import Packages:
import os
import requests
from datetime import datetime, timedelta
import locale
from dateutil.relativedelta import relativedelta
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

#Import scripts and functions
from GetKmdAcessToken import GetKMDToken
from NovaLogin import GetNovaCookies
import GetCaseInfoAndCheckCaseState
import SendBomEmail
import CheckIfEmailSent 
import SendDigitalPost
#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
KMD_access_token = GetKMDToken(orchestrator_connection)
#GetNovaCookies(orchestrator_connection)



# ---- Henter Kø-elementer ----
Sagsnummer = "S2021-456011"
uuid = "a56f6298-0a23-408e-bd26-e01808907c28"
caseUuid = "9c60ce1c-5f57-44ab-b805-44800017000c"
TastStartDate = "2025-02-18T13:23:10.9487697+01:00"
TaskDeadline = "2025-02-18T01:00:00+01:00"
caseworkerPersonId = "43ed2c49-a62f-4dde-84e1-428b0061328a"
RykkerNummer = 3


# ----- Run GetCaseInfoAndCheckCaseState -----
Arguments_GetCaseInfoAndCheckCaseState = {
    "in_caseUuid": caseUuid,
    "in_caseworkerPersonId": caseworkerPersonId,
    "in_ListOfErrorMessages": [],
    "in_ListOfFailedCases": [],
    "in_KMDNovaURL": KMDNovaURL,
    "in_Sagsnummer": Sagsnummer,
    "in_Token": KMD_access_token
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
ListOfErrorMessages = GetCaseInfoAndCheckCaseState_Output_arguments.get("ListOfFailedCases")
ListOfFailedCases = GetCaseInfoAndCheckCaseState_Output_arguments.get("ListOfFailedCases")
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
        NewDeadline = None
        DigitalPostSendt = None
        BeskrivelseTilEjer = None
        Tidspunkt = None
        Dato = None

        def case_1():
            # One-liner date formatting & calculations
            AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
            AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
            FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            # Extracting date and time separately
            Dato, Tidspunkt = FormatDate.split(" ")
            # Concatenating date and time as Deadline
            Deadline = f"{Dato}{Tidspunkt}"
            EmailText = ("Vi har endnu ikke modtaget besked om, at dit projekt er sat i gang.\n\n"
                        "Tilladelsen bortfalder, hvis arbejdet ikke er sat i gang inden et år fra tilladelsens dato. "
                        "Dette skal meddeles via Byg og Miljø.\n\n\nMed venlig hilsen\n\nByggeri")
            Title = "1. Rykkerskrivelse - Projektet er ikke påbegyndt"
            Description = "Rykkerskrivelse udført af robot"
            # New deadline calculation in one line
            NewDeadline = (AfgørelsesDato + timedelta(days=44*7)).strftime("%Y-%m-%d")
            DigitalPostSendt = True
            return locals()  # Return all variables as a dictionary
        def case_2():
            AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
            AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
            FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            Dato, Tidspunkt = FormatDate.split(" ")
            Deadline = f"{Dato} {Tidspunkt}"
            EmailText = ("Vi har endnu ikke modtaget besked om, at dit projekt er sat i gang.\n\n"
            "Tilladelsen bortfalder, hvis arbejdet ikke er sat i gang inden et år fra tilladelsens dato. "
            "Dette skal meddeles via Byg og Miljø.\n\n\nMed venlig hilsen\n\nByggeri")
            Title = "2. Rykkerskrivelse - Projektet er ikke påbegyndt"
            Description = "2. Rykkerskrivelse udført af robot"
            strDeadline = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d")
            DigitalPostSendt = False
            BeskrivelseTilEjer = "Mangler at udsende 1. orientering til ejer"
            return locals()
        def case_3():
            # One-liner date formatting & calculations
            AfgørelsesDato = datetime.strptime(Afgørelsesdato,"%Y-%m-%dT%H:%M:%S")
            AfgørelsesDatoFormateret = (AfgørelsesDato + relativedelta(years=1, days=-1)).strftime("%m/%d/%Y")
            FormatDate = datetime.strptime(AfgørelsesDatoFormateret, "%m/%d/%Y").strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            # Extracting date and time separately
            Dato, Tidspunkt = FormatDate.split(" ")
            # Concatenating date and time as Deadline
            Deadline = f"{Dato}{Tidspunkt}"
            # Set Danish locale for month formatting
        
            # Set locale to Danish for month names
            locale.setlocale(locale.LC_TIME, "da_DK.UTF-8")

            # Format the date in the desired format: "dd. MMMM yyyy"
            DayDate = AfgørelsesDato.strftime("%d. %B %Y")

            # Email text with proper newlines
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
            # Title
            Title = "3. Rykkerskrivelse - Projektet er ikke påbegyndt"
            Description = "Henlæg - 3. rykker er udført af robot"
            NewDeadline = (AfgørelsesDato + timedelta(days=14)).strftime("%Y-%m-%d")
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
        globals().update(assigned_variables)

        # ----- Run Send BomEmail -----
        # Arguments_SendBomEmail = {
        #     "in_Sagsnummer": Sagsnummer,
        #     "in_BomNumber": BomNumber,
        #     "in_BomCaseId": BomCaseId,
        #     "in_CaseAddress": CaseAddress,
        #     "in_BomCaseType": BomCaseType,
        #     "in_Kommunenummer": Kommunenummer,
        #     "in_BomCaseTypeCode": BomCaseTypeCode,
        #     "in_CadastralNumber": CadastralNumber,
        #     "in_bomCasePhaseCode": BomCasePhaseCode,
        #     "in_bomCaseStateCode": BomCaseStateCode,
        #     "in_StreetName": StreetName,
        #     "in_HouseNumber": HouseNumber,
        #     "in_Tidspunkt": Tidspunkt,
        #     "in_Dato": Dato,
        #     "in_EmailText": EmailText,
        #     "in_Title": Title
        # }
        # SendBomEmail_Output_arguments = SendBomEmail.invoke_SendBomEmail(Arguments_SendBomEmail,orchestrator_connection)
        # Text = SendBomEmail_Output_arguments.get("out_text")
        # print(Text)

            # ----- Run CheckIfEmailSent -----
        Arguments_CheckIfEmailSent = {
            "in_Sagsnummer": Sagsnummer,
            "in_Token": KMD_access_token,
            "in_Title": Title,
            "in_NovaAPIURL": KMDNovaURL,
            
        }
        CheckIfEmailSent_Output_arguments = CheckIfEmailSent.invoke_CheckIfEmailSent(Arguments_CheckIfEmailSent,orchestrator_connection)
        out_DocumentSendt = CheckIfEmailSent_Output_arguments.get("out_DocumentSendt")
        print(out_DocumentSendt)

        if out_DocumentSendt and (RykkerNummer == 2 or RykkerNummer == 3):
            # ----- Run SendDigitalPost -----
            Arguments_SendDigitalPost = {
                "in_Afgørelsesdato": Afgørelsesdato,
                "in_Beskrivelse": Description,
                "in_Sagsnummer": Sagsnummer,
                "in_caseworkerPersonId": caseworkerPersonId,
                "in_Dato": Dato,
                "in_NovaAPIURL": KMDNovaURL,
                "in_RykkerNummer": RykkerNummer,
                "in_Token": KMD_access_token,
                "in_BeskrivelseTilEjer": BeskrivelseTilEjer
            }
            SendDigitalPost_Output_arguments = SendDigitalPost.invoke_SendDigitalPost(Arguments_SendDigitalPost,orchestrator_connection)
            out_DigitaltPostSendt = SendDigitalPost_Output_arguments.get("out_DigitaltPostSendt")
            print(out_DigitaltPostSendt)
        
        else:
            print(f"Der udsendes ikke digital post da rykker nummer er: {RykkerNummer}")


    else: 
        print("Sagen er ikke tilknyttet BOM")

        if RykkerNummer == 1: 
            print("Sagen er ikke tilknyttet BOM - Sender mail og opdaterer sag ift. manuel håndtering")
