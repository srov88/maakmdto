#!/usr/bin/env python3
"""Schrijf documenten uit NotuBiz-meetingexports naar één CSV-bestand.

t.b.v. pilot gemeente Stichtse Vecht NotuBiz --> RANU MAIS-Flexis.
Volledig AI-gegenereerd 

Je gebruikt dit om een aanvulbare excel te genereren, die te gebruiken is in de maakMDTOuitYML.py

Installatie:
    python -m pip install PyYAML

Gebruik:
    python toon_documenten.py
    python toon_documenten.py --bronmap "C:\\Exports"
    python toon_documenten.py --bronmap ./exports --uitvoer documenten.csv
    python toon_documenten.py --bronmap ./exports --recursief
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterator

try:
    import yaml
except ImportError:
    sys.exit(
        "PyYAML ontbreekt. Installeer dit pakket met: "
        "python -m pip install PyYAML"
    )


# Pas deze map aan als je het script zonder --bronmap wilt starten.
BRONMAP = Path(r"./bronmap")

# Pas deze bestandsnaam aan als je het script zonder --uitvoer wilt starten.
UITVOERBESTAND = Path("documenten.csv")

# NotuBiz-exports kunnen C0/C1-stuurtekens bevatten die niet geldig zijn in YAML.
ONGELDIGE_TEKENS = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]"
)


def lees_yaml(pad: Path) -> dict[str, Any]:
    """Lees een YAML-export en verwijder ongeldige stuurtekens."""
    tekst = pad.read_text(encoding="utf-8-sig")
    tekst = ONGELDIGE_TEKENS.sub("", tekst)
    gegevens = yaml.safe_load(tekst)
    if not isinstance(gegevens, dict):
        raise ValueError("de YAML bevat geen object op het hoogste niveau")
    return gegevens


def als_lijst(waarde: Any) -> list[Any]:
    """NotuBiz gebruikt voor lege verzamelingen soms {} in plaats van []."""
    if isinstance(waarde, list):
        return waarde
    if isinstance(waarde, dict):
        return list(waarde.values())
    return []


def tekstwaarde(waarde: Any) -> str:
    if waarde is None:
        return ""
    if isinstance(waarde, (datetime, date)):
        return waarde.isoformat(sep=" ")
    return str(waarde)


def meeting_startdatum(meeting: dict[str, Any]) -> str:
    """Neem de eerste ingevulde start_date uit plannings."""
    for planning in als_lijst(meeting.get("plannings")):
        if isinstance(planning, dict) and planning.get("start_date") is not None:
            return tekstwaarde(planning["start_date"])
    return ""


def bestandsnaam(document: dict[str, Any]) -> str:
    """Geef de bestandsnaam van de actuele documentversie terug."""
    versies = [
        versie
        for versie in als_lijst(document.get("versions"))
        if isinstance(versie, dict) and versie.get("file_name")
    ]
    if not versies:
        return tekstwaarde(document.get("file_name"))

    actuele_versie = document.get("version")
    for versie in versies:
        if actuele_versie is not None and str(versie.get("id")) == str(actuele_versie):
            return tekstwaarde(versie["file_name"])

    # Als het actuele versienummer ontbreekt, is de laatste versie de beste keuze.
    return tekstwaarde(versies[-1]["file_name"])


def documenten_uit_agenda(
    agenda_items: Any,
    overgenomen_prefix: str = "",
) -> Iterator[tuple[str, str, str, str]]:
    """Lever prefix, document-id, titel en bestandsnaam uit alle agendapunten."""
    for item in als_lijst(agenda_items):
        if not isinstance(item, dict):
            continue

        type_data = item.get("type_data")
        eigen_prefix = (
            type_data.get("title_prefix") if isinstance(type_data, dict) else None
        )
        prefix = tekstwaarde(eigen_prefix) if eigen_prefix not in (None, "") else overgenomen_prefix

        for document in als_lijst(item.get("documents")):
            if not isinstance(document, dict):
                continue
            yield (
                prefix,
                tekstwaarde(document.get("id")),
                tekstwaarde(document.get("title")),
                bestandsnaam(document),
            )

        yield from documenten_uit_agenda(item.get("agenda_items"), prefix)


def yaml_bestanden(bronmap: Path, recursief: bool) -> list[Path]:
    patronen = ("*.yml", "*.yaml")
    bestanden: list[Path] = []
    for patroon in patronen:
        bestanden.extend(
            bronmap.rglob(patroon) if recursief else bronmap.glob(patroon)
        )
    return sorted(set(bestanden), key=lambda pad: str(pad).casefold())


def veilige_kolom(waarde: str) -> str:
    """Houd iedere documentregel op één uitvoerregel."""
    return waarde.replace("\t", " ").replace("\r", " ").replace("\n", " ")


def argumenten() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Schrijf documenten uit alle NotuBiz-YAML-bestanden in een bronmap "
            "naar één CSV-bestand."
        )
    )
    parser.add_argument(
        "--bronmap",
        type=Path,
        default=BRONMAP,
        help=f"map met YAML-bestanden (standaard: {BRONMAP})",
    )
    parser.add_argument(
        "--recursief",
        action="store_true",
        help="doorzoek ook onderliggende mappen",
    )
    parser.add_argument(
        "--uitvoer",
        type=Path,
        default=UITVOERBESTAND,
        help=f"te schrijven CSV-bestand (standaard: {UITVOERBESTAND})",
    )
    return parser.parse_args()


def main() -> int:
    args = argumenten()
    bronmap = args.bronmap.expanduser()
    if not bronmap.is_dir():
        print(f"Fout: bronmap bestaat niet of is geen map: {bronmap}", file=sys.stderr)
        return 2

    bestanden = yaml_bestanden(bronmap, args.recursief)
    if not bestanden:
        print(f"Geen .yml- of .yaml-bestanden gevonden in: {bronmap}", file=sys.stderr)
        return 1

    rijen: list[tuple[str, str, str, str, str]] = []
    aantal_fouten = 0

    for yaml_pad in bestanden:
        try:
            meeting = lees_yaml(yaml_pad)
            startdatum = meeting_startdatum(meeting)
            for prefix, document_id, titel, naam in documenten_uit_agenda(
                meeting.get("agenda_items")
            ):
                rijen.append(
                    tuple(
                        veilige_kolom(kolom)
                        for kolom in (startdatum, prefix, document_id, titel, naam)
                    )
                )
        except (OSError, UnicodeError, yaml.YAMLError, ValueError) as fout:
            print(f"Waarschuwing: {yaml_pad}: {fout}", file=sys.stderr)
            aantal_fouten += 1

    uitvoerpad = args.uitvoer.expanduser()
    try:
        uitvoerpad.parent.mkdir(parents=True, exist_ok=True)
        with uitvoerpad.open("w", encoding="utf-8-sig", newline="") as csv_bestand:
            schrijver = csv.writer(csv_bestand, delimiter=";", lineterminator="\n")
            schrijver.writerow(
                ("start_date", "agendapunt", "document_id", "titel", "bestandsnaam")
            )
            schrijver.writerows(rijen)
    except OSError as fout:
        print(f"Fout: CSV-bestand kan niet worden geschreven: {fout}", file=sys.stderr)
        return 2

    print(
        f"{len(rijen)} document(en) uit {len(bestanden)} YAML-bestand(en) "
        f"geschreven naar: {uitvoerpad}"
    )
    return 1 if aantal_fouten else 0


if __name__ == "__main__":
    raise SystemExit(main())
