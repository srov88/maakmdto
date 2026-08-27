#!/usr/bin/env python3

"""
Gemaakt door Thijs Vorstenburg en Ronald Koenis.

Het script is gemaakt om een MDTO-SIP te maken van 2 raadsvergaderingsbestanden die op een bronmap staan.
zodat het is in te lezen in bv MAIS-Flexis.

Input bestaat uit
- de YML-bestanden van de vergadering, dat vormt de basis
- een excel met informatie over gescande documenten die via een id zijn gelinkt aan de YML.
- de scans zelf.

Gebruik:
    python lees_yml_bronnen.py --bronmap ./project_sources

De bronmap kan ook via de omgevingsvariabele YML_BRONMAP worden ingesteld.

Prompt:
Maak een python script dat alle .yml-bestanden leest (gebruik de voorbeeldbestanden uit de bronnen) van een in te stellen bronlocatie, inclusief all submappen. 
Voor elk yml-bestand:
Toon van de meeting de eigenschappen meeting id, naam (staat in attributes met id 1), locatie (staat in attribute met id 50) en de datum en starttijd (staat in start_date).
En voor elk agenda-item:
Het id, het agendapunt (staat in de title_prefix), de naam (staat in attributes met id 1), de startseconde (staat in start_offset) en eindseconde (staat in end_offset)
Doe dit niet voor agenda-items
En voor elk document van het agenda-item:
Het id, de titel, de wijzigingsdatum, de categorie (staat in de eerste types met een ingevulde value) en de bestandsnaam.

"""

# We gebruiken deze andere python-modules:
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Any

import yaml
import mdto
from mdto.gegevensgroepen import *


def verwijder_ongeldige_tekens(tekst: str) -> tuple[str, int]:
    """Verwijder control-tekens die volgens de YAML-specificatie niet zijn toegestaan."""
    geldige_tekst = "".join(
        teken
        for teken in tekst
        if teken in "\t\n\r"
        or 0x20 <= ord(teken) <= 0x7E
        or ord(teken) >= 0xA0
    )
    return geldige_tekst, len(tekst) - len(geldige_tekst)


def lees_yml_bestanden(bronmap: Path) -> list[tuple[Path, Any]]:
    """Lees alle .yml-bestanden onder *bronmap*, alfabetisch gesorteerd."""
    if not bronmap.exists():
        raise FileNotFoundError(f"Bronmap bestaat niet: {bronmap}")
    if not bronmap.is_dir():
        raise NotADirectoryError(f"Bronlocatie is geen map: {bronmap}")

    resultaten: list[tuple[Path, Any]] = []
    for bestand in sorted(bronmap.rglob("*.yml")):
        try:
            # utf-8-sig verwerkt zowel gewone UTF-8 als de BOM in de voorbeelden.
            with bestand.open("r", encoding="utf-8-sig") as invoer:
                ruwe_tekst = invoer.read()
            yaml_tekst, aantal_verwijderd = verwijder_ongeldige_tekens(ruwe_tekst)
            if aantal_verwijderd:
                print(
                    f"Waarschuwing: {bestand.name}: {aantal_verwijderd} "
                    "ongeldig control-teken verwijderd",
                    file=sys.stderr,
                )
            inhoud = yaml.safe_load(yaml_tekst)
        except (OSError, UnicodeError, yaml.YAMLError) as fout:
            print(f"Waarschuwing: {bestand.name} kon niet worden gelezen: {fout}", file=sys.stderr)
            continue

        resultaten.append((bestand, inhoud))

    return resultaten


def waarde_met_id(items: Any, gezocht_id: int, sleutel: str = "value") -> Any:
    """Haal *sleutel* op uit het eerste item met het opgegeven id."""
    if not isinstance(items, list):
        return None
    for item in items:
        if isinstance(item, dict) and item.get("id") == gezocht_id:
            return item.get(sleutel)
    return None


def agenda_attributen(agenda_item: dict[str, Any]) -> Any:
    """Geef de attributen uit type_data van een agenda-item terug."""
    type_data = agenda_item.get("type_data")
    return type_data.get("attributes") if isinstance(type_data, dict) else None


def eerste_bestandsnaam(document: dict[str, Any]) -> Any:
    """Geef de eerste aanwezige bestandsnaam uit de documentversies terug."""
    versions = document.get("versions")
    if not isinstance(versions, list):
        return None
    for version in versions:
        if isinstance(version, dict) and version.get("file_name"):
            return version["file_name"]
    return None


def heeft_waarde(waarde: Any) -> bool:
    """Bepaal of een veld een niet-lege waarde bevat."""
    return waarde is not None and str(waarde).strip() != ""


def eerste_gevulde_waarde(items: Any) -> Any:
    """Geef de eerste niet-lege value uit een lijst met mappings terug."""
    if not isinstance(items, list):
        return None
    for item in items:
        if isinstance(item, dict) and heeft_waarde(item.get("value")):
            return item["value"]
    return None


def agenda_item_kop(agenda_item: dict[str, Any]) -> tuple[Any, Any]:
    """Geef het agendapunt en de naam van een agenda-item terug."""
    type_data = agenda_item.get("type_data")
    agendapunt = type_data.get("title_prefix") if isinstance(type_data, dict) else None
    naam = waarde_met_id(agenda_attributen(agenda_item), 1)
    return agendapunt, naam


def iter_agenda_items(items: Any, niveau: int = 0):
    """Doorloop alleen benoemde agenda-items, inclusief benoemde subitems."""
    if not isinstance(items, list):
        return
    for item in items:
        if not isinstance(item, dict):
            continue
        agendapunt, naam = agenda_item_kop(item)
        tonen = heeft_waarde(agendapunt) or heeft_waarde(naam)
        if tonen:
            yield item, niveau
        # Geldige kinderen van een overgeslagen leeg item blijven zichtbaar.
        kindniveau = niveau + 1 if tonen else niveau
        yield from iter_agenda_items(item.get("agenda_items"), kindniveau)


def toon_document(document: dict[str, Any], inspringing: str) -> None:
    """Toon de gevraagde eigenschappen van één document."""
    categorie = eerste_gevulde_waarde(document.get("types")) or "-"
    bestandsnaam = eerste_bestandsnaam(document) or "-"
    print(f"{inspringing}Document")
    print(f"{inspringing}  id:              {document.get('id', '-')}")
    print(f"{inspringing}  titel:           {document.get('title', '-')}")
    print(f"{inspringing}  wijzigingsdatum: {document.get('last_modified', '-')}")
    print(f"{inspringing}  categorie:       {categorie}")
    print(f"{inspringing}  bestandsnaam:    {bestandsnaam}")


def toon_agenda_item(agenda_item: dict[str, Any], niveau: int) -> None:
    """Toon een agenda-item en alle direct gekoppelde documenten."""
    inspringing = "  " * (niveau + 1)
    agendapunt, naam = agenda_item_kop(agenda_item)
    agendapunt = agendapunt if heeft_waarde(agendapunt) else "-"
    naam = naam if heeft_waarde(naam) else "-"
    startseconde = agenda_item.get("start_offset")
    eindseconde = agenda_item.get("end_offset")
    startseconde = "-" if startseconde in (None, "") else startseconde
    eindseconde = "-" if eindseconde in (None, "") else eindseconde

    print(f"{inspringing}Agenda-item")
    print(f"{inspringing}  id:            {agenda_item.get('id', '-')}")
    print(f"{inspringing}  agendapunt:    {agendapunt}")
    print(f"{inspringing}  naam:          {naam}")
    print(f"{inspringing}  startseconde:  {startseconde}")
    print(f"{inspringing}  eindseconde:   {eindseconde}")

    documenten = agenda_item.get("documents")
    if isinstance(documenten, list):
        for document in documenten:
            if isinstance(document, dict):
                toon_document(document, inspringing + "  ")


def toon_meeting(bestand: Path, inhoud: Any) -> None:
    """Toon meeting, agenda-items en documenten uit een Notubiz-export."""
    if not isinstance(inhoud, dict):
        print(f"- {bestand.name}: YAML-hoofdstructuur is geen mapping")
        return

    naam = waarde_met_id(inhoud.get("attributes"), 1) or "-"
    locatie = waarde_met_id(inhoud.get("attributes"), 50) or "-"
    planning = inhoud.get("plannings") or []
    start_date = planning[0].get("start_date") if planning else None
    datum, starttijd = "-", "-"
    if start_date:
        delen = str(start_date).split(maxsplit=1)
        datum = delen[0]
        starttijd = delen[1] if len(delen) == 2 else "-"

    print(f"\nBestand: {bestand}")
    print("Meeting")
    print(f"  meeting id: {inhoud.get('id', '-')}")
    print(f"  naam:       {naam}")
    print(f"  locatie:    {locatie}")
    print(f"  datum:      {datum}")
    print(f"  starttijd:  {starttijd}")

    for agenda_item, niveau in iter_agenda_items(inhoud.get("agenda_items")):
        toon_agenda_item(agenda_item, niveau)


def verwerk_bronmap(bronmap: Path) -> int:
    """Lees de bronmap, toon de inhoudssamenvatting en geef een exitcode terug."""
    try:
        documenten = lees_yml_bestanden(bronmap)
    except (FileNotFoundError, NotADirectoryError) as fout:
        print(f"Fout: {fout}", file=sys.stderr)
        return 2

    if not documenten:
        print(f"Geen leesbare .yml-bestanden gevonden in {bronmap}")
        return 1

    print(f"{len(documenten)} .yml-bestand(en) gelezen uit {bronmap.resolve()}:")
    for bestand, inhoud in documenten:
        toon_meeting(bestand.relative_to(bronmap), inhoud)
    return 0


def parse_args() -> argparse.Namespace:
    standaardmap = os.environ.get("YML_BRONMAP", "project_sources")
    parser = argparse.ArgumentParser(
        description=(
            "Lees alle .yml-bestanden uit een bronmap en de submappen en toon "
            "de meetinggegevens."
        )
    )
    parser.add_argument(
        "--bronmap",
        type=Path,
        default=Path(standaardmap),
        help=f"Hoofdmap met .yml-bestanden en submappen (standaard: {standaardmap!r})",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return verwerk_bronmap(args.bronmap)


if __name__ == "__main__":
    raise SystemExit(main())
