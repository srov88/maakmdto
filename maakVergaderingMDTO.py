import mdto
from mdto.gegevensgroepen import *
from pathlib import Path
import shutil
import sys
import pandas as pd #om te installeren: python3 -m pip install pandas en: python3 -m pip install openpyxl
 
bronMap = Path(r"input") 
doelMap = Path(r"output")

# Als je nog geen excel hebt, is dit het statement om alle bestanden tonen, hiermee kan je de excel, kolom BestandsnaamInclPad voeden.
#for bestand in bronMap.rglob("*"):
#    if bestand.is_file():
#        relatieve_naam = bestand.relative_to(bronMap)
#        print(relatieve_naam)

# inlezen begeleidende excel:
excel = Path(bronMap / "Bestandsbeschrijvingen.xlsx")
if not excel.exists():
    print(f"Bestand '{excel}' ontbreekt op {bronMap}.")
    sys.exit("voortijdig einde")
    
# Eerste werkblad inlezen
df = pd.read_excel(excel, sheet_name=0)
vereiste_kolommen = ["BestandsnaamInclPad", "id-uitgever", "vergadering.id", "vergadering.naam", "doc.id", "doc.naam", "doc.vaststellingsdatum"]
ontbrekende_kolommen = [
    kolom for kolom in vereiste_kolommen
    if kolom not in df.columns
]
if ontbrekende_kolommen:
    print(f"Excel gevonden, maar er ontbreken kolommen: {', '.join(ontbrekende_kolommen)}")
    sys.exit("voortijdig einde")

print("Excel gevonden en alle vereiste kolommen zijn aanwezig.")

#de doelmap maken of legen: 
doelMap.mkdir(parents=True, exist_ok=True)
if bronMap != doelMap :
    for item in doelMap.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()
else :
    print("bronmap gelijk aan doelmap, wordt nog niet ondersteund")
    sys.exit("voortijdig einde")
    
# Maak een set met alle relatieve paden in BronMap, nodig voor de controle verderop.
bestanden = {
    str(p.relative_to(bronMap))
    for p in bronMap.rglob("*")
    if p.is_file()
}

#Eerst de dingen klaarzetten die voor elke regel gelijk zijn:
archiefvormer = VerwijzingGegevens(
    verwijzingNaam="gemeente Stichtse Vecht",
    verwijzingIdentificatie=IdentificatieGegevens("gm1904","TOOI register gemeenten compleet")
    )
#waardering is altijd blijvend te bewaren
waardering = BegripGegevens(begripLabel="Blijvend te bewaren",
            begripCode="B",
            begripBegrippenlijst=VerwijzingGegevens("Begrippenlijst Waarderingen MDTO"))
# vooralsnog: beperkingGebruik kent nooit beperkingen:
beperkingType = BegripGegevens("Geen beperking", VerwijzingGegevens("Begrippenlijst BeperkingGebruikTypeLijst MDTO"))
beperkingGebruik = BeperkingGebruikGegevens(beperkingGebruikType=beperkingType)

# OK, all set, here we go. We gaan voor elke rij uit de excel objecten aanmaken:
vergaderingId=""
for index, rij in df.iterrows():
    BestandsnaamInclPad = str(rij["BestandsnaamInclPad"]).strip()

    if BestandsnaamInclPad in bestanden:
        print(f"Verwerken rij {index + 2}: {BestandsnaamInclPad}")
    else:
        print(f"Rij {index + 2}: bestand '{BestandsnaamInclPad}' is niet gevonden in {bronMap}")
        continue
        
    #Een kopie van het bestand naar doelmap zetten (lukt niet binnen een 'try')
    bestand = Path(bronMap / BestandsnaamInclPad)
    doelbestand = Path (doelMap / BestandsnaamInclPad)
    doelbestand.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(bestand, doelbestand)

    try:
        #eerst een vergaderings.informatieobject aanmaken ... als we een nieuwe vergadering aantreffen:
        if vergaderingId !=  str(rij["id-uitgever"]).strip() + "." + str(rij["vergadering.id"]).strip():
            vergaderingId = str(rij["id-uitgever"]).strip() + "." + str(rij["vergadering.id"]).strip()
            vergaderingNaam = str(rij["vergadering.naam"]).strip()
            print(f"Nieuwe vergadering {vergaderingNaam}")
            # maak identificatiekenmerk element
            vergaderingInformatieobjectIdentificatieGegegevens = IdentificatieGegevens(vergaderingId, "gemeente Stichtse Vecht")
        
            #optionele parameters komen in optargs:
            optargs = {}
                
            # maak informatieobject op basis van deze gegevens
            informatieobject = Informatieobject (
                identificatie = vergaderingInformatieobjectIdentificatieGegegevens,
                naam = vergaderingNaam,
                waardering = waardering,
                archiefvormer = archiefvormer,
                beperkingGebruik = beperkingGebruik,
                aggregatieniveau = BegripGegevens("Dossier", VerwijzingGegevens("Begrippenlijst Aggregatieniveaus MDTO")),
                **optargs
                )
     
            # informatieobject opslaan (naam moet gelijk zijn aan de submap):
            uitvoerbestand = doelMap / (BestandsnaamInclPad.split("/", 1)[0] + "/" + BestandsnaamInclPad.split("/", 1)[0] + ".mdto.xml")
            informatieobject.save(uitvoerbestand)
            print(f"- Informatieobject aangemaakt: {uitvoerbestand}")
            
        # per bestand maken we eerst het informatieobject aan:
        id = str(rij["id-uitgever"]).strip() + "." + str(rij["doc.id"]).strip()
        bestandInformatieobjectIdentificatieGegegevens = IdentificatieGegevens(id, "gemeente Stichtse Vecht")
    
        #optionele parameters komen in optargs:
        optargs = {}
        
        #FIXME dit geeft index-fouten
 #       vaststellingsdatum = df.loc[df["BestandsnaamInclPad"] == bestand.name, "doc.vaststellingsdatum"].iloc[0]
#        if pd.notna(vaststellingsdatum):
#            datum_string = vaststellingsdatum.strftime("%Y-%m-%d" + "T23:00:00")
#            print (datum_string)
        
        # maak informatieobject voor het bestand:
        informatieobject = Informatieobject (
            identificatie = bestandInformatieobjectIdentificatieGegegevens,
            naam = str(rij["doc.naam"]).strip(),
            waardering = waardering,
            archiefvormer = archiefvormer,
            beperkingGebruik = beperkingGebruik,
            isOnderdeelVan = VerwijzingGegevens(
                verwijzingNaam = vergaderingNaam,
                verwijzingIdentificatie = vergaderingInformatieobjectIdentificatieGegegevens
                ),
            **optargs
            )
     
        # informatieobject opslaan (de .rsplit(".", 1)[0] haalt de extensie weg):
        uitvoerbestand = doelMap / (BestandsnaamInclPad.rsplit(".", 1)[0] + ".mdto.xml")
        informatieobject.save(uitvoerbestand)
        print(f"- Informatieobject aangemaakt: {uitvoerbestand}")

        # en daarna de xml voor het bestand zelf:
        representatie = VerwijzingGegevens(
            verwijzingNaam=informatieobject.naam,
            verwijzingIdentificatie=bestandInformatieobjectIdentificatieGegegevens
            )

        mdto_bestand = Bestand.from_file(
            file_or_filename=bestand,
            isRepresentatieVan=representatie,
            use_mimetype=True,
        )
 
        uitvoerbestand = doelMap / (BestandsnaamInclPad + ".bestand.mdto.xml")
        mdto_bestand.save(uitvoerbestand)
        print(f"- Bestandsobject aangemaakt: {uitvoerbestand}")

    except Exception as fout:
        print(f"FOUT bij {bestand}: {type(fout).__name__}: {fout}")
 
print("Klaar.")
