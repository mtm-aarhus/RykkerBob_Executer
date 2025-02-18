import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from GetKmdAcessToken import GetKMDToken
import requests
from NovaLogin import GetNovaCookies
import GetCaseInfoAndCheckCaseState

#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
KMD_access_token = GetKMDToken(orchestrator_connection)
#verification_token, session_handler, request_token = GetNovaCookies(orchestrator_connection)



# ---- Henter Kø-elementer ----
Sagsnummer = "S2021-456011"
uuid = "a56f6298-0a23-408e-bd26-e01808907c28"
caseUuid = "9c60ce1c-5f57-44ab-b805-44800017000c"
TastStartDate = "2025-02-18T13:23:10.9487697+01:00"
TaskDeadline = "2025-02-18T01:00:00+01:00"
caseworkerPersonId = "43ed2c49-a62f-4dde-84e1-428b0061328a"
RykkerNummer = 1


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
print("Test")