"""
lag_demo_data.py — Generer realistiske norske testdata for demo
Kjoer: python lag_demo_data.py
Krever: pandas, openpyxl
"""

import pandas as pd
import random
from datetime import date, timedelta


random.seed(42)

# Norske kundenavn
KUNDER_PRIVAT = [
    "Hansen Familie", "Andersen Familie", "Nilsen Husholdning",
    "Berg Privat", "Karlsen Hjem", "Larsen Familie", "Eriksen Husholdning",
    "Johansen Privat", "Olsen Familie", "Pedersen Hjem",
    "Thorsen Familie", "Haugen Husholdning", "Dahl Privat",
    "Moen Familie", "Bakke Husholdning"
]

KUNDER_BEDRIFT = [
    "Hansen Borettslag", "Kiwi Majorstuen", "Rema Grorud",
    "Andersen Eiendom AS", "Berg Utleie AS", "Oslo Boligforvaltning",
    "Skovveien Sameie", "Groruddalen Borettslag", "Frogner Eiendom",
    "Nordstrand Sameie", "Bjerke Naeringseiendom AS"
]

INNTEKT_KATEGORIER = [
    "Rørleggerarbeid", "Vareoppdrag", "Serviceavtale", "Akuttoppdrag"
]

UTGIFT_KATEGORIER = [
    "Materialer", "Drivstoff", "Forsikring", "Telefon/internett",
    "Verktøy", "Lønn assistent"
]

BESKRIVELSER_INNTEKT = {
    "Rørleggerarbeid": [
        "Utskifting av vannrør kjøkken", "Installasjon ny dusj",
        "Rørlegging baderom", "Reparasjon vannlekkasje", "Ny varmtvannsbereder",
        "Utskifting toalett", "Legging nye avløpsrør", "Tilkobling oppvaskmaskin"
    ],
    "Vareoppdrag": [
        "Levering og montering kraner", "Installasjon radiatorer",
        "Montering baderomsinnredning", "Levering varmesystem"
    ],
    "Serviceavtale": [
        "Årsservice varmeanlegg", "Kvartalsvis vedlikehold", "Service borettslag jan",
        "Årsavtale vedlikehold", "Service varmepumpe"
    ],
    "Akuttoppdrag": [
        "Akutt rørbrudd natt", "Akutt vannlekkasje helg",
        "Hasteoppdrag oversvømt kjeller", "Akutt reparasjon helligdag"
    ]
}

BESKRIVELSER_UTGIFT = {
    "Materialer": [
        "Kobbrerør og koplinger", "Isolasjonsmateriale", "Pakninger og tetninger",
        "Rørfittings diverse", "Ventiler og kran-deler", "Avløpsdeler"
    ],
    "Drivstoff": [
        "Diesel firmabil januar", "Bensin og diesel", "Drivstoff arbeidsuke",
        "Tank firmabil"
    ],
    "Forsikring": [
        "Yrkesskadeforsikring", "Ansvarsforsikring bedrift", "Bilforsikring firmabil"
    ],
    "Telefon/internett": [
        "Mobilabonnement", "Internett kontor", "Telefon og data"
    ],
    "Verktøy": [
        "Nytt boreverktøy", "Reservedeler maskinpark", "Sikkerhetsutstyr",
        "Diverse håndverktøy"
    ],
    "Lønn assistent": [
        "Lønn lærlingassistent uke 1-2", "Lønn lærling uke 3-4",
        "Overtidsbetaling assistent", "Lønn vikar"
    ]
}


def tilfeldig_dato(ar, maned):
    """Generer tilfeldig dato i en gitt maned."""
    if maned == 2:
        max_dag = 28
    elif maned in [4, 6, 9, 11]:
        max_dag = 30
    else:
        max_dag = 31
    dag = random.randint(1, max_dag)
    return date(ar, maned, dag)


def generer_rad(ar, maned, fakturanr_teller):
    """Generer en rad med transaksjonsdata."""
    # Vekting: 60% inntekt, 40% utgift
    er_inntekt = random.random() < 0.60

    if er_inntekt:
        kategori = random.choice(INNTEKT_KATEGORIER)
        beskrivelse = random.choice(BESKRIVELSER_INNTEKT[kategori])
        kunde = random.choice(KUNDER_PRIVAT + KUNDER_BEDRIFT)

        # Variasjon etter maned — mars litt bedre
        if maned == 1:
            belop = random.randint(2000, 18000)
        elif maned == 2:
            belop = random.randint(2500, 20000)
        else:  # Mars
            belop = random.randint(3000, 25000)

        fakturanr = f"F{ar}-{fakturanr_teller:04d}"
    else:
        kategori = random.choice(UTGIFT_KATEGORIER)
        beskrivelse = random.choice(BESKRIVELSER_UTGIFT[kategori])
        kunde = "Intern utgift"

        if kategori == "Lønn assistent":
            belop = -random.randint(8000, 15000)
        elif kategori == "Forsikring":
            belop = -random.randint(2000, 5000)
        elif kategori == "Materialer":
            belop = -random.randint(500, 8000)
        else:
            belop = -random.randint(500, 3000)

        fakturanr = f"U{ar}-{fakturanr_teller:04d}"

    dato = tilfeldig_dato(ar, maned)
    return {
        "Dato": dato.strftime("%d.%m.%Y"),
        "Fakturanr": fakturanr,
        "Kunde": kunde,
        "Beskrivelse": beskrivelse,
        "Kategori": kategori,
        "Belop": belop
    }


def main():
    rader = []
    fakturanr = 100

    # 20 rader per maned (jan, feb, mars 2025) = 60 rader
    for maned in [1, 2, 3]:
        for _ in range(20):
            rad = generer_rad(2025, maned, fakturanr)
            rader.append(rad)
            fakturanr += 1

    df = pd.DataFrame(rader)

    # Sorter pa dato
    df["_dato_sort"] = pd.to_datetime(df["Dato"], format="%d.%m.%Y")
    df = df.sort_values("_dato_sort").drop(columns=["_dato_sort"])
    df = df.reset_index(drop=True)

    utfil = "testdata_ror_as.xlsx"
    with pd.ExcelWriter(utfil, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transaksjoner")

        # Kolonnebredder
        ws = writer.sheets["Transaksjoner"]
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 28
        ws.column_dimensions["D"].width = 40
        ws.column_dimensions["E"].width = 22
        ws.column_dimensions["F"].width = 14

    print(f"Demo-data generert: {utfil}")
    print(f"Antall rader: {len(df)}")
    print(f"Perioden: januar - mars 2025")

    # Kort statistikk
    inntekter = df[df["Belop"] > 0]["Belop"].sum()
    utgifter = df[df["Belop"] < 0]["Belop"].sum()
    print(f"\nRask oversikt:")
    print(f"  Totale inntekter: {inntekter:,.0f} NOK")
    print(f"  Totale utgifter:  {utgifter:,.0f} NOK")
    print(f"  Resultat:         {inntekter + utgifter:,.0f} NOK")


if __name__ == "__main__":
    main()
