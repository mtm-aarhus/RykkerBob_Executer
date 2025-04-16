"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")
    import Datastore

    Datastore.out_ListOfProcessedItems.clear()
    Datastore.out_ListOfErrorMessages.clear()
    Datastore.out_ListOfFailedCases.clear()
