"""
Lager en testfil (testdata.xlsx) du kan bruke med finansrapport.py
"""

import pandas as pd

data = [
    # Januar
    ("2024-01-05", "Faktura – Kunde Andersen AS",      "Inntekt",        12500),
    ("2024-01-10", "Faktura – Kunde Berg Consulting",  "Inntekt",         8000),
    ("2024-01-15", "Husleie kontor januar",            "Husleie",        -3500),
    ("2024-01-18", "Adobe Creative Cloud",             "Programvare",     -600),
    ("2024-01-20", "Internett og telefon",             "Programvare",     -499),
    ("2024-01-25", "Drivstoff og parkering",           "Transport",       -850),
    ("2024-01-28", "Lunsj med kunde",                  "Mat",             -480),
    # Februar
    ("2024-02-03", "Faktura – Kunde Dahl Regnskap",    "Inntekt",        15000),
    ("2024-02-07", "Faktura – Nettbutikk oppsett",     "Inntekt",         6500),
    ("2024-02-14", "Husleie kontor februar",           "Husleie",        -3500),
    ("2024-02-16", "Facebook-annonser",                "Markedsføring",  -1200),
    ("2024-02-19", "Regnskapsprogram (Tripletex)",     "Programvare",     -399),
    ("2024-02-22", "Drivstoff",                        "Transport",       -620),
    ("2024-02-27", "Forsikring næringsvirksomhet",     "Forsikring",      -890),
    # Mars
    ("2024-03-04", "Faktura – Kunde Eriksen Bygg",     "Inntekt",        18000),
    ("2024-03-06", "Faktura – Månedlig support",       "Inntekt",         4500),
    ("2024-03-12", "Husleie kontor mars",              "Husleie",        -3500),
    ("2024-03-15", "Google Ads",                       "Markedsføring",  -1500),
    ("2024-03-18", "Kurs og faglig utvikling",         "Annet",          -1800),
    ("2024-03-22", "Drivstoff og bompenger",           "Transport",       -730),
    ("2024-03-28", "Kontorrekvisita",                  "Annet",           -350),
]

df = pd.DataFrame(data, columns=["Dato", "Beskrivelse", "Kategori", "Beløp"])
df.to_excel("testdata.xlsx", index=False)
print("testdata.xlsx er opprettet!")
print("   Kjor na: python finansrapport.py testdata.xlsx")
