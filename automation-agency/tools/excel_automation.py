"""
excel_automation.py — Automatisk Excel-rensing og rapportgenerering
Bruk: python excel_automation.py minfil.xlsx
Krever: pandas, openpyxl
"""

import sys
import os
import pandas as pd
from datetime import datetime


def normaliser_kolonnenavn(df):
    """Gjor alle kolonnenavn til lowercase og fjern ekstra mellomrom."""
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df


def finn_pengekolonne(kolonner):
    """Finn kolonne som inneholder 'bel', 'amount' eller 'sum'."""
    for col in kolonner:
        if any(term in col for term in ["bel", "amount", "sum"]):
            return col
    return None


def finn_datokolonne(kolonner):
    """Finn kolonne som inneholder 'dato' eller 'date'."""
    for col in kolonner:
        if any(term in col for term in ["dato", "date"]):
            return col
    return None


def rens_data(df):
    """Fjern tomme rader og strip whitespace fra tekst-kolonner."""
    print("  Fjerner tomme rader...")
    df = df.dropna(how="all")

    print("  Renser whitespace fra tekstfelt...")
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].str.strip()

    df = df.reset_index(drop=True)
    return df


def konverter_numeriske_kolonner(df):
    """Forsok automatisk konvertering av kolonner til numerisk."""
    numeriske = []
    for col in df.columns:
        if df[col].dtype == object:
            forsok = pd.to_numeric(df[col].str.replace(",", ".").str.replace(" ", ""), errors="coerce")
            if forsok.notna().sum() > len(df) * 0.5:
                df[col] = forsok
                numeriske.append(col)
        elif pd.api.types.is_numeric_dtype(df[col]):
            numeriske.append(col)
    return df, numeriske


def legg_til_maned(df, datokolonne):
    """Legg til maned-kolonne basert pa datokolonne."""
    try:
        df[datokolonne] = pd.to_datetime(df[datokolonne], dayfirst=True, errors="coerce")
        df["maned"] = df[datokolonne].dt.strftime("%Y-%m")
        print(f"  Lagt til 'maned'-kolonne basert pa '{datokolonne}'")
    except Exception as e:
        print(f"  Kunne ikke konvertere datoer: {e}")
    return df


def analyser_pengekolonne(df, pengekolonne):
    """Separer inntekter og utgifter."""
    df[pengekolonne] = pd.to_numeric(df[pengekolonne], errors="coerce")
    inntekter = df[df[pengekolonne] > 0].copy()
    utgifter = df[df[pengekolonne] < 0].copy()
    return inntekter, utgifter


def lag_rapport(df, pengekolonne, datokolonne, kildefil, numeriske_kolonner):
    """Lag en norsk tekstrapport."""
    linjer = []
    linjer.append("=" * 60)
    linjer.append("AUTOMATISK RAPPORT — KABI AUTOMATION")
    linjer.append(f"Generert: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    linjer.append(f"Kildefil: {kildefil}")
    linjer.append("=" * 60)
    linjer.append("")
    linjer.append(f"DATAOVERSIKT")
    linjer.append(f"  Antall rader (etter rensing): {len(df)}")
    linjer.append(f"  Antall kolonner: {len(df.columns)}")
    linjer.append(f"  Kolonner: {', '.join(df.columns.tolist())}")
    linjer.append(f"  Numeriske kolonner: {', '.join(numeriske_kolonner) if numeriske_kolonner else 'Ingen funnet'}")
    linjer.append("")

    if pengekolonne:
        inntekter, utgifter = analyser_pengekolonne(df, pengekolonne)
        total_inntekt = inntekter[pengekolonne].sum()
        total_utgift = utgifter[pengekolonne].sum()
        resultat = total_inntekt + total_utgift

        linjer.append(f"OKONOMI (kolonne: '{pengekolonne}')")
        linjer.append(f"  Totale inntekter:  {total_inntekt:>12,.2f} NOK")
        linjer.append(f"  Totale utgifter:   {total_utgift:>12,.2f} NOK")
        linjer.append(f"  Resultat:          {resultat:>12,.2f} NOK")
        linjer.append(f"  Antall inntekter:  {len(inntekter)} transaksjoner")
        linjer.append(f"  Antall utgifter:   {len(utgifter)} transaksjoner")
        linjer.append("")

        if "maned" in df.columns:
            linjer.append("MANED FOR MANED:")
            maned_sum = df.groupby("maned")[pengekolonne].sum()
            for maned, total in maned_sum.items():
                prefix = "+" if total >= 0 else ""
                linjer.append(f"  {maned}: {prefix}{total:,.2f} NOK")
            linjer.append("")

    if datokolonne and "maned" in df.columns:
        linjer.append(f"DATO-INFORMASJON")
        try:
            linjer.append(f"  Tidligste dato: {df[datokolonne].min().strftime('%d.%m.%Y')}")
            linjer.append(f"  Seneste dato:   {df[datokolonne].max().strftime('%d.%m.%Y')}")
            linjer.append(f"  Antall maneder: {df['maned'].nunique()}")
        except Exception:
            pass
        linjer.append("")

    linjer.append("STATISTIKK PER NUMERISK KOLONNE:")
    for col in numeriske_kolonner:
        if col == pengekolonne:
            continue
        try:
            linjer.append(f"  {col}:")
            linjer.append(f"    Gjennomsnitt: {df[col].mean():,.2f}")
            linjer.append(f"    Min:          {df[col].min():,.2f}")
            linjer.append(f"    Max:          {df[col].max():,.2f}")
        except Exception:
            pass
    linjer.append("")
    linjer.append("=" * 60)
    linjer.append("Rapport generert av KABI AUTOMATION")
    linjer.append("=" * 60)

    return "\n".join(linjer)


def main():
    if len(sys.argv) < 2:
        print("FEIL: Ingen fil oppgitt.")
        print("Bruk: python excel_automation.py minfil.xlsx")
        sys.exit(1)

    innfil = sys.argv[1]

    if not os.path.exists(innfil):
        print(f"FEIL: Filen '{innfil}' ble ikke funnet.")
        sys.exit(1)

    if not innfil.lower().endswith((".xlsx", ".xls")):
        print("FEIL: Filen ma vaere en Excel-fil (.xlsx eller .xls)")
        sys.exit(1)

    print("\n===== KABI AUTOMATION — Excel-rensing =====\n")
    print(f"Leser fil: {innfil}")

    try:
        df = pd.read_excel(innfil)
    except Exception as e:
        print(f"FEIL: Kunne ikke lese Excel-filen: {e}")
        sys.exit(1)

    print(f"  Lest {len(df)} rader og {len(df.columns)} kolonner")

    print("\nNormaliserer kolonnenavn...")
    df = normaliser_kolonnenavn(df)
    print(f"  Kolonner funnet: {', '.join(df.columns.tolist())}")

    print("\nRenser data...")
    df = rens_data(df)
    print(f"  Rene rader: {len(df)}")

    print("\nKonverterer numeriske kolonner...")
    df, numeriske_kolonner = konverter_numeriske_kolonner(df)
    if numeriske_kolonner:
        print(f"  Numeriske kolonner: {', '.join(numeriske_kolonner)}")
    else:
        print("  Ingen numeriske kolonner funnet automatisk")

    pengekolonne = finn_pengekolonne(df.columns.tolist())
    datokolonne = finn_datokolonne(df.columns.tolist())

    if pengekolonne:
        print(f"\nPengekolonne funnet: '{pengekolonne}'")
        inntekter, utgifter = analyser_pengekolonne(df, pengekolonne)
        print(f"  Inntekter: {len(inntekter)} rader, total: {inntekter[pengekolonne].sum():,.2f} NOK")
        print(f"  Utgifter:  {len(utgifter)} rader, total: {utgifter[pengekolonne].sum():,.2f} NOK")
    else:
        print("\nIngen pengekolonne funnet (leter etter 'bel', 'amount', 'sum')")

    if datokolonne:
        print(f"\nDatokolonne funnet: '{datokolonne}'")
        df = legg_til_maned(df, datokolonne)
    else:
        print("\nIngen datokolonne funnet (leter etter 'dato', 'date')")

    # Generer filnavn
    base = os.path.splitext(innfil)[0]
    utfil_excel = base + "_renset.xlsx"
    utfil_rapport = base + "_rapport.txt"

    # Lagre renset Excel
    print(f"\nLagrer renset Excel: {utfil_excel}")
    try:
        with pd.ExcelWriter(utfil_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Renset data")
            if pengekolonne:
                inntekter_df, utgifter_df = analyser_pengekolonne(df, pengekolonne)
                if not inntekter_df.empty:
                    inntekter_df.to_excel(writer, index=False, sheet_name="Inntekter")
                if not utgifter_df.empty:
                    utgifter_df.to_excel(writer, index=False, sheet_name="Utgifter")
        print("  Excel-fil lagret!")
    except Exception as e:
        print(f"  FEIL ved lagring av Excel: {e}")

    # Generer og lagre rapport
    print(f"\nGenererer rapport: {utfil_rapport}")
    rapport_tekst = lag_rapport(df, pengekolonne, datokolonne, innfil, numeriske_kolonner)

    try:
        with open(utfil_rapport, "w", encoding="utf-8") as f:
            f.write(rapport_tekst)
        print("  Rapport lagret!")
    except Exception as e:
        print(f"  FEIL ved lagring av rapport: {e}")

    print("\n" + rapport_tekst)
    print("\n===== FERDIG =====")
    print(f"Renset fil:   {utfil_excel}")
    print(f"Rapport:      {utfil_rapport}")


if __name__ == "__main__":
    main()
