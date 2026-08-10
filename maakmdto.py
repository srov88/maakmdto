# Gemaakt door Thijs Vorstenburg en Ronald Koenis
import mdto
from mdto.gegevensgroepen import *
from pathlib import Path
import shutil 
import sys
import pandas #om te installeren: python3 -m pip install pandas en: python3 -m pip install openpyxl
from config import bronMap, doelMap, excelmap

# Als je nog geen excel hebt, is dit het statement om alle bestanden tonen,
# hiermee kan je de excel, kolom BestandsnaamInclPad voeden.
# for bestand in bronMap.rglob("*"):
#       if bestand.is_file():
#           relatieve_naam = bestand.relative_to(bronMap)
#           print(relatieve_naam)

# inlezen begeleidende excel:
excel = Path(excelmap / "Bestandsbeschrijvingen.xlsx")
if not excel.exists(): 
    print(f"Bestand '{excel}' ontbreekt op {excelmap}.")
    sys.exit("voortijdig einde")

# Eerste werkblad inlezen
df = pandas.read_excel(excel, sheet_name=0)
vereiste_kolommen = ["BestandsnaamInclPad", "id-uitgever", "vergadering.id", "vergadering.naam", "vergaderdatum", "doc.id", "doc.naam", "doc.classificatie", "doc.vaststellingsdatum"]
ontbrekende_kolommen = [
    kolom for kolom in vereiste_kolommen if kolom not in df.columns
    ]
if ontbrekende_kolommen: 
    print(f"Excel gevonden, maar er ontbreken kolommen: {', '.join(ontbrekende_kolommen)}")
    sys.exit("voortijdig einde")

print("Excel gevonden en alle vereiste kolommen zijn aanwezig.")

#de doelmap maken of legen: 
doelMap.mkdir(parents=True, exist_ok=True)
if bronMap != doelMap:
    for item in doelMap.iterdir():
        if item.is_dir(): 
            shutil.rmtree(item)
        else: 
            item.unlink()
else:
    print("bronmap gelijk aan doelmap, wordt nog niet ondersteund")
    sys.exit("voortijdig einde")

# Maak een set met alle relatieve paden in BronMap, nodig voor de controle verderop.
bestanden = {
    Path(p.relative_to(bronMap))
    for p in bronMap.rglob("*")
    if p.is_file() 
    }

#Eerst de dingen klaarzetten die voor elke regel gelijk zijn:
archiefvormer = VerwijzingGegevens(
    verwijzingNaam="gemeente Stichtse Vecht",
    verwijzingIdentificatie=IdentificatieGegevens("gm1904","TOOI register gemeenten compleet")
    ) 

#waardering is altijd blijvend te bewaren
waardering = BegripGegevens(
    begripLabel="Blijvend te bewaren",
    begripCode="B",
    begripBegrippenlijst=VerwijzingGegevens("Begrippenlijst Waarderingen MDTO")
    ) 
informatiecategorie = BegripGegevens(
    begripLabel="Agenda, verslag en besluitenlijst van bestuurlijke besluitvorming - Verwerkt",
    begripCode="19.1.6",
    begripBegrippenlijst=VerwijzingGegevens("Selectielijst gemeenten en intergemeentelijke organen 2017")
    )
informatiecategorieMotie = BegripGegevens(
    begripLabel="Adhesiebetuiging en/of motie - Ingewilligd",
    begripCode="6.1.4",
    begripBegrippenlijst=VerwijzingGegevens("Selectielijst gemeenten en intergemeentelijke organen 2017")
    )
# vooralsnog Geen beperkingen, TODO: aanpassen formatting en toevoegen andere attributen 
beperkingType = BegripGegevens("Geen beperking", VerwijzingGegevens("Begrippenlijst BeperkingGebruikTypeLijst MDTO"))
beperkingGebruik = BeperkingGebruikGegevens(beperkingGebruikType=beperkingType)

# OK, all set, here we go. We gaan voor elke rij uit de excel objecten aanmaken:
vergaderingId=""
for index, rij in df.iterrows():
    BestandsnaamInclPad = Path(str(rij["BestandsnaamInclPad"]).strip())

    if BestandsnaamInclPad in bestanden: 
        print(f"Verwerken rij {index + 2}: {BestandsnaamInclPad}") 
    else:
        print(f"FOUT: Rij {index + 2}: bestand '{BestandsnaamInclPad}' is niet gevonden in {bronMap}")
        continue

    #Een kopie van het bestand naar doelmap zetten (aangepast zodat het het met een try ook werkt)
    bestand = bronMap / BestandsnaamInclPad
    doelbestand = doelMap / BestandsnaamInclPad
    doelbestand.parent.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2(bestand, doelbestand)
        print("Bestand gekopieerd.")
    except Exception as fout:
        print(f"FOUT bij kopieren {bestand}: {type(fout).__name__}: {fout}")

    try:
        # eerst een vergaderobject aanmaken als we een nieuwe vergadering aantreffen
        # Voor differentiatie vergaderobject genoemd, maar het is een informatieobject
        # TODO agendapunten aanmaken als informatieobject
        if vergaderingId != str(rij["id-uitgever"]).strip() + "." + str(rij["vergadering.id"]).strip():
            vergaderingId = str(rij["id-uitgever"]).strip() + "." + str(rij["vergadering.id"]).strip()
            vergaderingNaam = str(rij["vergadering.naam"]).strip()
            print(f"Nieuwe vergadering {vergaderingNaam}")
            
            # maak identificatiekenmerk element
            vergaderingInformatieobjectIdentificatieGegegevens = IdentificatieGegevens(vergaderingId, "gemeente Stichtse Vecht")
            vergaderdatum = rij["vergaderdatum"]
            # TODO: de aanvangstijd ergens aan toevoegen.
            if pandas.notna(vergaderdatum):
                datum_string = vergaderdatum.strftime("%Y-%m-%d")
            else:    
                print(f"FOUT: de vergaderdatum is niet gevonden")
                sys.exit("voortijdig einde")
            vergaderingDekkingInTijd = DekkingInTijdGegevens(
                dekkingInTijdType = BegripGegevens("vergaderdatum", VerwijzingGegevens("Begrippenlijst TODO")),
                dekkingInTijdBegindatum = datum_string
                )
            
            #optionele parameters komen in optargs:
            optargs = {}

            # maak vergaderobject op basis van deze gegevens
            informatieobject = Informatieobject (
                identificatie = vergaderingInformatieobjectIdentificatieGegegevens,
                naam = vergaderingNaam,
                waardering = waardering,
                informatiecategorie = informatiecategorie,
                archiefvormer = archiefvormer,
                taal = "nl",
                dekkingInTijd = vergaderingDekkingInTijd,
                beperkingGebruik = beperkingGebruik,
                aggregatieniveau = BegripGegevens("Dossier", VerwijzingGegevens("Begrippenlijst Aggregatieniveaus MDTO")),
                **optargs
                )

            # vergaderobject opslaan (naam moet gelijk zijn aan de bovenliggende map)
            # Het script gaat nu uit van de doelmap als de map die de naam van de vergadering bevat.
            vergadering_map = BestandsnaamInclPad.parent

            # Bestand staat direct in de bronmap: de doelmap is de vergadering. In de SIP zit maar één vergadering
            if vergadering_map == Path(""):
                vergadering_naam = doelMap.name
                uitvoermap = doelMap

            # Bestand staat in een submap: die submap is de vergadering. In de SIP zitten meerdere vergaderingen
            else:
                vergadering_naam = vergadering_map.name
                uitvoermap = doelMap / vergadering_map

            uitvoerbestand = uitvoermap / f"{vergadering_naam}.mdto.xml"
            uitvoerbestand.parent.mkdir(parents=True, exist_ok=True)
            informatieobject.save(uitvoerbestand)
            print(f"- vergaderobject aangemaakt: {uitvoerbestand}")

        # per bestand maken we eerst het informatieobject aan:
        id = str(rij["id-uitgever"]).strip() + "." + str(rij["doc.id"]).strip()
        bestandInformatieobjectIdentificatieGegegevens = IdentificatieGegevens(id, "gemeente Stichtse Vecht")

        #optionele parameters komen in optargs:
        optargs = {}

        # de vaststelling als event meegeven UNDER CONSTRUCTION:
        vaststellingsdatum = rij["doc.vaststellingsdatum"]
        if pandas.notna(vaststellingsdatum):
            datum_string = vaststellingsdatum.strftime("%Y-%m-%d" + "T23:00:00")
            #print (datum_string)
        
        # maak informatieobject voor het bestand:
        informatieobject = Informatieobject (
            identificatie = bestandInformatieobjectIdentificatieGegegevens, 
            naam = str(rij["doc.naam"]).strip(),
            waardering = waardering,
            informatiecategorie = informatiecategorieMotie if str(rij["doc.classificatie"]).strip().lower() == "motie"  else informatiecategorie,
            archiefvormer = archiefvormer,
            taal = "nl",
            beperkingGebruik = beperkingGebruik, 
            aggregatieniveau = BegripGegevens("Archiefstuk", VerwijzingGegevens("Begrippenlijst Aggregatieniveaus MDTO")),
            classificatie = BegripGegevens(str(rij["doc.classificatie"]).strip(), VerwijzingGegevens("Begrippenlijst TODO")),
            isOnderdeelVan = VerwijzingGegevens( 
                verwijzingNaam = vergaderingNaam,
                verwijzingIdentificatie = vergaderingInformatieobjectIdentificatieGegegevens 
                ), 
            **optargs 
            ) 

        # informatieobject opslaan (de .rsplit(".", 1)[0] haalt de extensie weg): 
        # uitvoerbestand = doelMap / (BestandsnaamInclPad.rsplit(".", 1)[0] + ".mdto.xml") --> rsplit vervangen omdat string gaf problemen met windows versus Mac 
        uitvoerbestand = doelMap / BestandsnaamInclPad.with_suffix(".mdto.xml") 
        uitvoerbestand.parent.mkdir(parents=True, exist_ok=True)
        informatieobject.save(uitvoerbestand) 
        print(f"- Informatieobject aangemaakt: {uitvoerbestand}") 

        # en daarna de xml voor het bestand zelf: 
        representatie = VerwijzingGegevens( 
            verwijzingNaam=informatieobject.naam, 
            verwijzingIdentificatie=bestandInformatieobjectIdentificatieGegegevens 
            ) 

        # Mimetype (True) geeft error met aanmaken van bepaalde files, PRONOM (false) heeft daar geen last van. Sowieso is Pronom een betere identificatie dan mimetype 
        mdto_bestand = Bestand.from_file(
            file_or_filename=bestand,
            isRepresentatieVan=representatie, 
            use_mimetype=False, 
            ) 
        
        uitvoerbestand = ( doelMap 
                            / BestandsnaamInclPad.parent 
                            / f"{BestandsnaamInclPad.stem}.bestand.mdto.xml" 
                            ) 
        uitvoerbestand.parent.mkdir(parents=True, exist_ok=True)
        mdto_bestand.save(uitvoerbestand) 
        print(f"- Bestandsobject aangemaakt: {uitvoerbestand}") 

    except Exception as fout: 
        print(f"FOUT bij {bestand}: {type(fout).__name__}: {fout}") 

    # Een except die wat meer info geeft, was nodig voor het troubleshooten met mimetype.
    # Statement kan er in een later stadium weer uit
    # except Exception:
    #     import traceback
    #     traceback.print_exc()

print("Klaar.")
# Einde
