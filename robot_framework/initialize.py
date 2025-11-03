"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from NovaLogin import GetNovaCookies
import Datastore
import platform
import socket

def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    machine_name = platform.node() or socket.gethostname()
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace(f"Initializing, running on {machine_name}")
    

    Datastore.save_data(Datastore.default_data.copy())
    orchestrator_connection.log_trace("Importing cookies")
    
    GetNovaCookies(orchestrator_connection)