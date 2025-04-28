"""This module handles resetting the state of the computer so the robot can work with a clean slate."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import Datastore
from SendSMTPMail import send_email 

def reset(orchestrator_connection: OrchestratorConnection) -> None:
    """Clean up, close/kill all programs and start them again. """
    orchestrator_connection.log_trace("Resetting.")
    clean_up(orchestrator_connection)
    close_all(orchestrator_connection)
    kill_all(orchestrator_connection)
    open_all(orchestrator_connection)

def Send_Finish_mail(orchestrator_connection: OrchestratorConnection) -> None:
    orchestrator_connection.log_trace("Sender mail til sagsbehandler")
    data = Datastore.load_data()
    processed_items = data["ListOfProcessedItems"]
    failed_cases = data["ListOfFailedCases"]
    error_messages = data["ListOfErrorMessages"]

    if processed_items:
        print("Sender email overblik over kørte sagsnumre")

        # Create FinalString: each case number on a new line (HTML <br>)
        FinalString = "<br>".join(processed_items) + "<br><br>"

        # Build BodyMail
        BodyMail = f"""Kære Sagsbehandler,<br><br>
            "Robotten har udsendt rykkere til følgende sager:<br><br>
            {FinalString}
            Kontroller venligst at robotten har udført sit arbejde korrekt<br><br>
            Med venlig hilsen<br><br>
            Teknik og Miljø<br><br>
            Digitalisering<br><br>
            Aarhus Kommune."""

        # Define email details
        sender = "RykkerBob<rpamtm001@aarhus.dk>" 
        subject = "RykkerRobot sagsnummer oversigt"

        smtp_server = "smtp.adm.aarhuskommune.dk"   
        smtp_port = 25               

        # Call the send_email function
        send_email(
            receiver="Gujc@aarhus.dk",
            sender=sender,
            subject=subject,
            body=BodyMail,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )
    
    else:
        print("Ingen sager er blevet behandlet uden at fejle")
        
    # ----- Failed Cases Mail -----
    if failed_cases:
        print("Der er fejlede sager")

        ErrorMessageString = ""

        # Combine failed cases with their corresponding error messages
        for case_number, error_msg in zip(failed_cases, error_messages):
            ErrorMessageString += f"{case_number} - mangler følgende datakilde: {error_msg}<br><br>"

        BodyMailErrors = f"""Kære Sagsbehandler,<br><br>
            Kære Sagsbehandler,<br><br>
            Robotten mangler data ved følgende sager:<br><br>
            {ErrorMessageString}
            Kontroller venligst at sagen har data.<br><br>
            Med venlig hilsen<br><br>
            Teknik og Miljø<br><br>
            Digitalisering<br><br>
            Aarhus Kommune."""


        # Define email details
        sender = "RykkerBob<rpamtm001@aarhus.dk>" 
        subject = "RykkerRobot sagsnummer oversigt"

        smtp_server = "smtp.adm.aarhuskommune.dk"   
        smtp_port = 25               

        # Call the send_email function
        send_email(
            receiver="Gujc@aarhus.dk",
            sender=sender,
            subject=subject,
            body=BodyMailErrors,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )

    else:
        print("Ingen sager er fejlet.")


def clean_up(orchestrator_connection: OrchestratorConnection) -> None:
    """Do any cleanup needed to leave a blank slate."""
    orchestrator_connection.log_trace("Doing cleanup.")


def close_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Gracefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Closing all applications.")


def kill_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Forcefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Killing all applications.")


def open_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Open all programs used by the robot."""
    orchestrator_connection.log_trace("Opening all applications.")
