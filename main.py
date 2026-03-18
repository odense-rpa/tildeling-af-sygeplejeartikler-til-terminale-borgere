import asyncio
from datetime import datetime
import json
import logging
import sys
import argparse
import os
from pathlib import Path
import httpx
import sbsip

from automation_server_client import (
    AutomationServer,
    Workqueue,
    Credential,
    WorkItemStatus,
)
from kmd_nexus_client import NexusClientManager
from kmd_nexus_client.tree_helpers import filter_by_path

from odk_tools.word import generate_docx_binary
from odk_tools.tracking import Tracker
from process.config import load_excel_mapping, get_excel_mapping

proces_navn = "Tildeling af sygeplejeartikler til terminale borger (§122)"
nexus: NexusClientManager
tracker: Tracker

indsatsparagraffer = {
    "§ 83 stk. 1 nr 1",
    "§ 83 stk. 1 nr 2",
    "§ 83 stk. 2 nr 1",
    "§ 83 stk. 2 nr 2",
    "§ 119",
    "§ 94",
    "§ 95",
    "§ 138",
    "§ 26",
    "§ 20",
    "§9 stk. 2",
    "§ 23",
}


def terminalcheck(borger: dict):
    visning = nexus.borgere.hent_visning(borger)
    aktiviteter = nexus.nexus_client.get(
        visning["_links"]["patientActivities"]["href"]
    ).json()
    aktiviteter = [
        ref for ref in aktiviteter if ref.get("patientActivityType") == "formData"
    ]
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
        terminal_dato = datetime.strptime(
            terminal_dato, "%Y-%m-%dT%H:%M:%S.%f%z"
        ).strftime("%d-%m-%Y")
    return False, "", terminal_dato


def plejehjemscheck(borger: dict):
    borgers_orgs = nexus.organisationer.hent_organisationer_for_borger(borger)
    plejehjemsliste = {
        row["Plejehjem og bosteder"]
        for row in regler
        if row.get("Plejehjem og bosteder")
    }
    match = next(
        (org for org in borgers_orgs if org["organization"]["name"] in plejehjemsliste),
        None,
    )
    return (True, "Borger bor på plejehjem eller bosted") if match else (False, "")


def indsatscheck(borger: dict):
    visning = nexus.borgere.hent_visning(borger)
    borgers_indsatsreferencer = nexus.borgere.hent_referencer(visning)
    filtrerede_indsats_referencer = filter_by_path(
        borgers_indsatsreferencer,
        path_pattern="/Sundhedsfagligt grundforløb/*/Indsatser/*",
        active_pathways_only=False,
    ) + filter_by_path(
        borgers_indsatsreferencer,
        path_pattern="/Ældre og sundhedsfagligt grundforløb/*/Indsatser/*",
        active_pathways_only=False,
    )
    for indsats_ref in filtrerede_indsats_referencer:
        indsats = nexus.hent_fra_reference(indsats_ref)
        værdier = nexus.indsatser.hent_indsats_elementer(indsats)
        if (
            værdier.get("paragraph", {}).get("paragraph", {}).get("section", "")
            in indsatsparagraffer
        ):
            return True, "Borger har aktiv indsats"
    return False, ""


def opret_opgave_til_personalet(borger: dict, data: dict, besked: str):
    skemareferencer = nexus.skemaer.hent_skemareferencer(borger)
    skemareference = next(
        (ref for ref in skemareferencer if ref["id"] == data["skema_id"]), None
    )
    skema = nexus.hent_fra_reference(skemareference)
    nexus.opgaver.opret_opgave(
        objekt=skema,
        opgave_type="Tværfagligt samarbejde",
        titel=f"§26 afvist: {besked}",
        ansvarlig_organisation="Sygeplejehjælpemidler",
        ansvarlig_medarbejder=None,
        start_dato=datetime.now().strftime("%d-%m-%Y"),
        forfald_dato=datetime.now().strftime("%d-%m-%Y"),
        beskrivelse="",
    )


def opret_forløb(borger: dict):
    # Opret forløb til afgørelse og helhedspleje
    nexus.forløb.opret_forløb(
        borger=borger,
        grundforløb_navn="Ældre og sundhedsfagligt grundforløb",
        forløb_navn="Sag SOFF: Afgørelse - Lov om social service",
    )
    nexus.forløb.opret_forløb(
        borger=borger,
        grundforløb_navn="Ældre og sundhedsfagligt grundforløb",
        forløb_navn="Sag SOFF: Helhedspleje",
    )
    # Hent det oprettede afgørelsesforløb
    visning = nexus.borgere.hent_visning(borger)
    borgers_referencer = nexus.borgere.hent_referencer(visning)
    filtreret_forløb = filter_by_path(
        borgers_referencer,
        path_pattern="/Ældre og sundhedsfagligt grundforløb/Sag SOFF: Afgørelse - Lov om social service",
        active_pathways_only=False,
    )
    afgørelses_forløb = next(
        (
            forløb
            for forløb in filtreret_forløb
            if forløb["name"] == "Sag SOFF: Afgørelse - Lov om social service"
        ),
        None,
    )
    if not afgørelses_forløb:
        raise Exception(
            "Kunne ikke finde det oprettede afgørelsesforløb i borgerens referencer"
        )
    afgørelses_forløb = nexus.hent_fra_reference(afgørelses_forløb)

    return afgørelses_forløb


def send_brev_til_borger(
    borger: dict, data: dict, terminal_dato: str, afgørelses_forløb: dict
):
    brevfelter = {
        "GADE": borger["primaryAddress"]["addressLine1"],
        "POSTNR": borger["primaryAddress"]["postalCode"],
        "BY": borger["primaryAddress"]["postalDistrict"],
        "BORGERNAVN": borger["fullName"],
        "DAGSDATO": datetime.now().strftime("%d-%m-%Y"),
        "TERMINALDATO": terminal_dato,
        "FORSLAG": data["emne"],
        "CPR": data["cpr"],
    }

    # Konverter Word skabelon til PDF ved at sende den til en ekstern render service (odknet) sammen med brevfelter som data. Gem den resulterende PDF i output.pdf og send den som digital post til borgeren via SBSip
    with open(
        "input/Tildeling af sygeplejeartikler til terminale borgere.docx", "rb"
    ) as f:
        response = httpx.post(
            "http://rpa-ats.odknet.dk:8331/render",
            files={
                "file": ("Tildeling af sygeplejeartikler til terminale borgere.docx", f)
            },
            data={"fields": json.dumps(brevfelter)},
        )

    pdf_path = Path("Tildeling af sygeplejeartikler til terminale borgere (§26).pdf")
    pdf_path.write_bytes(response.content)

    sbsip.send_digital_post(
        cpr=data["cpr"],
        overskrift="Tildeling af sygeplejeartikler til terminale borgere (§26)",
        beskrivelse="Tildeling af sygeplejeartikler til terminale borgere (§26)",
        vedhæftet_fil=pdf_path,
    )

    # Upload dokument til nexus
    dokument = nexus.forløb.opret_dokument(
        borger=borger,
        forløb=afgørelses_forløb,
        fil=pdf_path.read_bytes(),
        filnavn="Bevilling tilskud sygeplejeartikler.pdf",
        titel="Bevilling tilskud sygeplejeartikler",
        noter=None,
        modtaget=datetime.now(),
    )
    if not dokument:
        pdf_path.unlink()  # Slet den genererede PDF, da den ikke kunne uploades til nexus
        raise Exception("Kunne ikke oprette dokument i nexus for den genererede PDF")
    

    # Tilføj tag til dokumentet i nexus
    tags = nexus.nexus_client.get(
        dokument["_links"]["availableTags"]["href"]
    ).json()
    tag = next((tag for tag in tags if tag["name"] == "ÆL § 26"), None)
    dokument["tags"] = dokument.get("tags", []) + [tag]
    nexus.nexus_client.put(dokument["_links"]["self"]["href"], json=dokument)

    pdf_path.unlink()


def opret_sagsnotat(borger: dict, terminal_dato: str, data: dict):
    skema_data = {
        "Emne": "Sygeplejeartikler § 26",
        "Tekst": (
            f"Der er på {datetime.now().strftime('%d-%m-%Y')} sendt bevilling til borger på "
            f"{data['emne']} jf. Ældreloven §26.\n"
            f"Borger er terminalerklæret pr. {terminal_dato} og opfylder kriterierne for bevilling."
        ),
    }

    skema = nexus.skemaer.opret_komplet_skema(
        borger=borger,
        skematype_navn="Sagsnotat - NY",
        handling_navn="Låst",
        data=skema_data,
        grundforløb="Ældre og sundhedsfagligt grundforløb",
        forløb="Sag SOFF: Helhedspleje",
    )
    if not skema:
        raise Exception("Kunne ikke oprette skema for sagsnotat")


def opret_indsats(borger: dict):

    visning = nexus.borgere.hent_visning(borger)
    borgers_referencer = nexus.borgere.hent_referencer(visning)
    filtrede_indsats_referencer = filter_by_path(
        borgers_referencer,
        path_pattern="/Ældre og sundhedsfagligt grundforløb/*/Indsatser/*",
        active_pathways_only=False,
    )
    ønsket_indsats_reference = next(
        (
            ref
            for ref in filtrede_indsats_referencer
            if ref["name"] == "Sygeplejeartikler - ÆL § 26"
        ),
        None,
    )

    if ønsket_indsats_reference:
        return

    nexus.indsatser.opret_indsats(
        borger=borger,
        grundforløb="Ældre og sundhedsfagligt grundforløb",
        forløb="Sag SOFF: Helhedspleje",
        indsats="Sygeplejeartikler - ÆL § 26",
        felter={
            "workflowApprovedDate": datetime.today(),
            "billingStartDate": datetime.today(),
            "entryDate": datetime.today(),
            "orderedDate": datetime.today(),
            "workflowRequestedDate": datetime.today(),
        },
        leverandør="Sygeplejehjælpemidler",
        oprettelsesform="Ansøg, Bevilg, Bestil",
    )


def tilføj_organisation(borger: dict):
    borgers_orgs = nexus.organisationer.hent_organisationer_for_borger(borger)
    if any(
        org["organization"]["name"] == "Sygeplejehjælpemidler" for org in borgers_orgs
    ):
        return

    sygeplejehjælpemidler_org = nexus.organisationer.hent_organisation_ved_navn(
        "Sygeplejehjælpemidler"
    )
    nexus.organisationer.tilføj_borger_til_organisation(
        borger, sygeplejehjælpemidler_org
    )


def afslut_opgave(data: dict):
    sygeplejehjælpemidler_org = nexus.organisationer.hent_organisation_ved_navn(
        "Sygeplejehjælpemidler"
    )
    aktivitetsliste = nexus.aktivitetslister.hent_aktivitetsliste(
        navn="Opgaver - 6 mdr tilbage til 6 mdr frem",
        organisation=sygeplejehjælpemidler_org,
        medarbejder=None,
    )
    opgave = next(
        (opgave for opgave in aktivitetsliste if opgave["id"] == data["opgave_id"]),
        None,
    )
    if not opgave:
        raise Exception(
            f"Kunne ikke finde opgave med id {data['opgave_id']} for at afslutte den"
        )
    opgave = nexus.hent_fra_reference(opgave)
    nexus.opgaver.luk_opgave(opgave)


async def populate_queue(workqueue: Workqueue):

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
        cpr = opgave["patients"][0]["patientIdentifier"]["identifier"]
        workqueue.add_item(
            data={
                "opgave_id": opgave["id"],
                "cpr": cpr,
                "emne": opgave["name"],
                "skema_id": opgave["children"][0]["id"],
            },
            reference=f"{cpr} / {opgave['id']}",
        )


async def process_workqueue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    for item in workqueue:
        with item:
            opgave_til_personalet = False
            besked_til_personalet = ""
            terminal_dato = "01-01-1970"
            afgørelses_forløb = None

            data = item.data  # Item data deserialized from json as dict
            try:
                # Process the item here
                borger = nexus.borgere.hent_borger(data["cpr"])
                # Hvis borger ikke bor i Odense, så skal der oprettes opgave til personalet
                if borger["primaryAddress"]["administrativeAreaCode"] != "461":
                    besked_til_personalet = "Borger bor ikke i Odense"
                    opgave_til_personalet = True

                # Check om terimnalerklæring
                if not opgave_til_personalet:
                    (
                        opgave_til_personalet,
                        besked_til_personalet,
                        terminal_dato,
                    ) = terminalcheck(borger)
                # Check om borger bor på plejehjem eller bosted
                if not opgave_til_personalet:
                    opgave_til_personalet, besked_til_personalet = plejehjemscheck(
                        borger
                    )
                # Check om borger har aktiv indsats under relevante paragraffer
                if not opgave_til_personalet:
                    opgave_til_personalet, besked_til_personalet = indsatscheck(borger)
                # Hvis et af checksne returerner True, så opret en opgave til personalet med beskeden og spring resten af behandlingen over
                if opgave_til_personalet:
                    opret_opgave_til_personalet(borger, data, besked_til_personalet)
                    tracker.track_partial_task(
                        process=proces_navn,
                    )
                    continue  # Skip resten af behandlingen og gå videre til næste item i køen

                # Opret forløb til afgørelse og helhedspleje
                afgørelses_forløb = opret_forløb(borger)
                # Generer og send brev til borger
                send_brev_til_borger(borger, data, terminal_dato, afgørelses_forløb)
                # Opret sagsnotat i nexus
                opret_sagsnotat(borger, terminal_dato, data)
                # Opret indsats i nexus
                opret_indsats(borger)
                # Tilføj organisation "Sygeplejehjælpemidler"
                tilføj_organisation(borger)
                # Afslut opgave
                afslut_opgave(data)

                tracker.track_task(process_name=proces_navn)

            except Exception as e:
                logger.error(f"Error processing item: {data}. Error: {e}")
                item.fail(str(e))


if __name__ == "__main__":
    ats = AutomationServer.from_environment()
    workqueue = ats.workqueue()

    # Initialize external systems for automation here..
    nexus_credential = Credential.get_credential("KMD Nexus - produktion")
    tracking_credential = Credential.get_credential("Odense SQL Server")
    SBSip_credential = Credential.get_credential("SBSip - produktion")

    tracker = Tracker(
        username=tracking_credential.username, password=tracking_credential.password
    )

    nexus = NexusClientManager(
        client_id=nexus_credential.username,
        client_secret=nexus_credential.password,
        instance=nexus_credential.data["instance"],
        timeout=60,
    )

    sbsip.start_sbsip(
        brugernavn=SBSip_credential.username,
        adgangskode=SBSip_credential.password,
    )

    # Parse command line arguments
    parser = argparse.ArgumentParser(description=proces_navn)
    parser.add_argument(
        "--excel-file",
        default=os.environ.get("EXCEL_MAPPING_PATH"),
        help="Path to the Excel file containing mapping data (default: ./Regelsæt.xlsx)",
    )

    parser.add_argument(
        "--word-template",
        default=os.environ.get("LETTER_TEMPLATE_PATH"),
        help="Path to the Word template for letter generation (default: ./Tildeling af sygeplejeartikler til terminale borgere.docx)",
    )

    parser.add_argument(
        "--queue",
        action="store_true",
        help="Populate the queue with test data and exit",
    )
    args = parser.parse_args()


    # Queue management
    if "--queue" in sys.argv:
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

        workqueue.clear_workqueue(WorkItemStatus.NEW)
        asyncio.run(populate_queue(workqueue))
        exit(0)



    # Process workqueue
    asyncio.run(process_workqueue(workqueue))
