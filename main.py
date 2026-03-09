import asyncio
from datetime import datetime
import logging
import sys
import argparse
import os

from automation_server_client import (
    AutomationServer,
    Workqueue,
    WorkItemError,
    Credential,
    WorkItemStatus,
)
from kmd_nexus_client import NexusClientManager
from odk_tools.tracking import Tracker
from process.config import load_excel_mapping, get_excel_mapping

proces_navn = "Tildeling af sygeplejeartikler til terminale borgere"
nexus: NexusClientManager
tracker: Tracker


def terminal_check(borger: dict):
    visning = nexus.borgere.hent_visning(borger)
    aktiviteter = nexus.nexus_client.get(
            visning["_links"]["patientActivities"]["href"]
        ).json()
    aktiviteter = [ref for ref in aktiviteter if ref.get("patientActivityType") == "formData"]
    terminal_erklæring_form = next(
        (
            ref
            for ref in aktiviteter
            if ref["formDefinition"]["title"] == "Terminalerklæring"
            and ref["formDataStatus"] != "DELETED"
            and (ref.get("additionalInformation") or [{}])[0].get("value") == "Ja"
        ),
        None,
    )
    if not terminal_erklæring_form:
        return True, "Borger har ingen terminalerklæring", None

    terminal_erklæring = nexus.hent_fra_reference(terminal_erklæring_form)
    terminal_dato = terminal_erklæring.get("observationTimestamp")
    if terminal_dato:
        terminal_dato = datetime.strptime(terminal_dato, "%Y-%m-%dT%H:%M:%S.%f%z").strftime("%d-%m-%Y")
    return False, "", terminal_dato

def plejehjem_check(borger: dict):
    logger = logging.getLogger(__name__)
    borgers_orgs = nexus.organisationer.hent_organisationer_for_borger(borger)
    plejehjemsliste = {row["Plejehjem og bosteder"] for row in regler if row.get("Plejehjem og bosteder")}
    match = next((org for org in borgers_orgs if org["organization"]["name"] in plejehjemsliste), None)
    return (True, "Borger bor på plejehjem eller bosted") if match else (False, "")

async def populate_queue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from populate workqueue!")
    sygeplejehjælpemidler_org = nexus.organisationer.hent_organisation_ved_navn(
        "Sygeplejehjælpemidler"
    )
    aktivitetsliste = nexus.aktivitetslister.hent_aktivitetsliste(
        navn="Opgaver - 6 mdr tilbage til 6 mdr frem",
        organisation=sygeplejehjælpemidler_org,
        medarbejder=None,
    )
    # Behold alle items hvor description = "Sygeplejehjælpemidler, Robotbruger Odense" og assignmentDescription indeholder §26:
    opgaver = [
        item
        for item in aktivitetsliste
        if item.get("description") == "Sygeplejehjælpemidler, Robotbruger Odense"
        and "§26" in (item.get("assignmentDescription") or "")
    ]

    for opgave in opgaver:
        workqueue.add_item(
            data={
                "opgave_id": opgave["id"],
                "cpr": opgave["patients"][0]["patientIdentifier"]["identifier"],
                "emne": opgave["name"],
                "skema_id": opgave["children"][0]["id"],
            },
            reference=opgave["patients"][0]["patientIdentifier"]["identifier"]
            + " / "
            + str(opgave["id"]),
        )


async def process_workqueue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from process workqueue!")
    opret_opgave_til_personalet = False
    besked_til_personalet = ""
    terminal_dato = None

    for item in workqueue:
        with item:
            data = item.data  # Item data deserialized from json as dict
            try:
                # Process the item here
                borger = nexus.borgere.hent_borger(data["cpr"])
                # Hvis borger ikke bor i Odense, så skal der oprettes opgave til personalet
                if borger["primaryAddress"]["administrativeAreaCode"] != "461":
                    besked_til_personalet = "Borger bor ikke i Odense"
                    opret_opgave_til_personalet = True

                if opret_opgave_til_personalet == False:
                    opret_opgave_til_personalet, besked_til_personalet, terminal_dato = terminal_check(borger)

                if opret_opgave_til_personalet == False:
                    opret_opgave_til_personalet, besked_til_personalet = plejehjem_check(borger)
                pass
            except Exception as e:
                logger.error(f"Error processing item: {data}. Error: {e}")
                item.fail(str(e))


if __name__ == "__main__":
    ats = AutomationServer.from_environment()
    workqueue = ats.workqueue()

    # Initialize external systems for automation here..
    nexus_credential = Credential.get_credential("KMD Nexus - produktion")
    nexus_database_credential = Credential.get_credential("KMD Nexus - database")
    tracking_credential = Credential.get_credential("Odense SQL Server")

    tracker = Tracker(
        username=tracking_credential.username, password=tracking_credential.password
    )

    nexus = NexusClientManager(
        client_id=nexus_credential.username,
        client_secret=nexus_credential.password,
        instance=nexus_credential.data["instance"],
        timeout=60,
    )

    # Parse command line arguments
    parser = argparse.ArgumentParser(description=proces_navn)
    parser.add_argument(
        "--excel-file",
        default=os.environ.get("EXCEL_MAPPING_PATH"),
        help="Path to the Excel file containing mapping data (default: ./Regelsæt.xlsx)",
    )

    parser.add_argument(
        "--queue",
        action="store_true",
        help="Populate the queue with test data and exit",
    )
    parser.add_argument("--prio", default=False)
    args = parser.parse_args()

    # Validate Excel files exists (skip validation for Windows paths on Linux)
    def is_windows_path(path: str) -> bool:
        """Check if path is a Windows path (has drive letter or UNC path)"""
        return (
            (len(path) > 1 and path[1] == ":")
            or path.startswith("\\\\")
            or path.startswith("//")
        )

    # Load excel mapping data once on startup (only if files exist on current system)
    if os.path.isfile(args.excel_file):
        load_excel_mapping(args.excel_file)
    elif not is_windows_path(args.excel_file):
        raise FileNotFoundError(f"Excel file not found: {args.excel_file}")

    regler = get_excel_mapping().get(
        "Ark1", []
    )  # intet "l" til sidst. Excel kan ikke have så mange karakterer.

    # Queue management
    if "--queue" in sys.argv:
        workqueue.clear_workqueue(WorkItemStatus.NEW)
        asyncio.run(populate_queue(workqueue))
        exit(0)

    # Process workqueue
    asyncio.run(process_workqueue(workqueue))
