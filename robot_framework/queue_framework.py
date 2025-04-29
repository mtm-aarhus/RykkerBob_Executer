"""This module is the primary module of the robot framework. It collects the functionality of the rest of the framework."""

# This module is not meant to exist next to linear_framework.py in production:
# pylint: disable=duplicate-code

import sys

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus

from robot_framework import initialize
from robot_framework import reset
from robot_framework.exceptions import handle_error, BusinessError, log_exception
from robot_framework import process
from robot_framework import config
import time  # Optional: for adding delay between retries


def main():
    """The entry point for the framework. Should be called as the first thing when running the robot."""
    orchestrator_connection = OrchestratorConnection.create_connection_from_args()
    sys.excepthook = log_exception(orchestrator_connection)

    orchestrator_connection.log_trace("Robot Framework started.")
    initialize.initialize(orchestrator_connection)

    queue_element = None
    error_count = 0
    task_count = 0
    max_retries = 1  # Number of retries
    # Retry loop
    for _ in range(config.MAX_RETRY_COUNT):
        try:
            reset.reset(orchestrator_connection)

            # Queue loop
            while task_count < config.MAX_TASK_COUNT:
                task_count += 1
                queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

                if not queue_element:
                    orchestrator_connection.log_info("Queue empty.")
                    break  # Break queue loop

                try:
                    for attempt in range(1, max_retries + 1):
                        try:
                            process.process(orchestrator_connection, queue_element)
                            print("Operation succeeded!")
                            break  # Exit the loop if successful
                        except Exception as e:
                            print(f"Attempt {attempt} failed: {e}")
                            if attempt < max_retries:
                                time.sleep(1)  # Optional: delay before retrying
                                reset.reset(orchestrator_connection)
                                print("Retrying...")
                            else:
                                print("Operation failed after 2 attempts.")
                                raise
                    
                    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)

                except BusinessError as error:
                    handle_error("Business Error", error, queue_element, orchestrator_connection)

            break  # Break retry loop

        # We actually want to catch all exceptions possible here.
        # pylint: disable-next = broad-exception-caught
        except Exception as error:
            error_count += 1
            handle_error(f"Process Error #{error_count}", error, queue_element, orchestrator_connection)

    reset.Send_Finish_mail(orchestrator_connection)
    reset.clean_up(orchestrator_connection)
    reset.close_all(orchestrator_connection)
    reset.kill_all(orchestrator_connection)

    if config.FAIL_ROBOT_ON_TOO_MANY_ERRORS and error_count == config.MAX_RETRY_COUNT:
        raise RuntimeError("Process failed too many times.")
