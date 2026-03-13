"""
pdf_extractor.py — Ekstraher data fra norske PDF-dokumenter med AI
Bruk: python pdf_extractor.py dokument.pdf
Krever: pdfplumber, anthropic
"""

import sys
import os
import json
import re

try:
    import pdfplumber
except ImportError:
    print("FEIL: 'pdfplumber' er ikke installert.")
    print("Installer med: pip install pdfplumber")
    sys.exit(1)

import anthropic
import pandas as pd


SYSTEM_PROMPT = (
    "Du er en ekspert pa a lese norske forretningsdokumenter. "
    "Ekstraher alle viktige felt fra denne teksten og returner dem som JSON. "
    "Bruk norske feltnavn. Inkluder alle datoer, belop, navn og referansenumre du finner. "
    "For fakturaer, bruk feltene: fakturanummer, dato, forfallsdato, leverandor, kunde, belop, mva, total. "
    "For andre dokumenter, bruk: tittel, dato, nokkelord, nokkelbelop, nokkelnavn, referanser. "
    "Returner KUN gyldig JSON, ingen forklarende tekst."
)


def les_pdf(filsti):
    """Ekstraher tekst fra alle sider i PDF."""
    print(f"  Leser PDF: {filsti}")
    all_tekst = []

    with pdfplumber.open(filsti) as pdf:
        antall_sider = len(pdf.pages)
        print(f"  Antall sider: {antall_sider}")

        for i, side in enumerate(pdf.pages, 1):
            tekst = side.extract_text()
            if tekst:
                all_tekst.append(f"--- Side {i} ---\n{tekst}")
                print(f"  Side {i}: {len(tekst)} tegn ekstrahert")
            else:
                print(f"  Side {i}: Ingen tekst funnet (muligens bilde)")

    return "\n\n".join(all_tekst)


def ekstraher_json_fra_tekst(tekst):
    """Forsok a parse JSON fra Claude sitt svar."""
    # Prover direkte parse
    try:
        return json.loads(tekst.strip())
    except json.JSONDecodeError:
        pass

    # Leter etter JSON-blokk i teksten
    json_match = re.search(r"\{.*\}", tekst, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            pass

    # Leter etter JSON i kodeblokk
    kode_match = re.search(r"```(?:json)?\s*(.*?)\s*```", tekst, re.DOTALL)
    if kode_match:
        try:
            return json.loads(kode_match.group(1))
        except json.JSONDecodeError:
            pass

    return None


def ekstraher_med_claude(klient, pdf_tekst):
    """Send PDF-tekst til Claude og fa strukturerte data tilbake."""
    print("\nSender til Claude AI for analyse...")

    # Begrens tekst til 15000 tegn for a unnga token-grenser
    if len(pdf_tekst) > 15000:
        pdf_tekst = pdf_tekst[:15000] + "\n\n[... resten av dokumentet er avkortet ...]"

    melding = klient.messages.create(
        model="claude-opus-4-5",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"Analyser dette dokumentet og ekstraher viktige felt som JSON:\n\n{pdf_tekst}"
            }
        ]
    )

    raa_svar = melding.content[0].text
    return raa_svar


def lagre_json(data, base_filnavn):
    """Lagre ekstraherte data til JSON-fil."""
    utfil = base_filnavn + "_ekstrahert.json"
    with open(utfil, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nJSON lagret: {utfil}")
    return utfil


def lagre_excel(data, base_filnavn):
    """Lagre ekstraherte data til Excel-fil."""
    utfil = base_filnavn + "_ekstrahert.xlsx"

    # Flatten nested dict til flat struktur
    flat_data = {}
    for nokkel, verdi in data.items():
        if isinstance(verdi, dict):
            for under_nokkel, under_verdi in verdi.items():
                flat_data[f"{nokkel}_{under_nokkel}"] = str(under_verdi)
        elif isinstance(verdi, list):
            flat_data[nokkel] = ", ".join(str(v) for v in verdi)
        else:
            flat_data[nokkel] = str(verdi) if verdi is not None else ""

    df = pd.DataFrame([flat_data])
    df.to_excel(utfil, index=False)
    print(f"Excel lagret: {utfil}")
    return utfil


def skriv_ut_felt(data):
    """Skriv ut ekstraherte felt pa norsk."""
    print("\n" + "=" * 50)
    print("EKSTRAHERTE FELT:")
    print("=" * 50)

    def skriv_dict(d, innrykk=0):
        for nokkel, verdi in d.items():
            prefix = "  " * innrykk
            if isinstance(verdi, dict):
                print(f"{prefix}{nokkel}:")
                skriv_dict(verdi, innrykk + 1)
            elif isinstance(verdi, list):
                print(f"{prefix}{nokkel}: {', '.join(str(v) for v in verdi)}")
            else:
                print(f"{prefix}{nokkel}: {verdi}")

    skriv_dict(data)
    print("=" * 50)


def main():
    if len(sys.argv) < 2:
        print("FEIL: Ingen fil oppgitt.")
        print("Bruk: python pdf_extractor.py dokument.pdf")
        sys.exit(1)

    innfil = sys.argv[1]

    if not os.path.exists(innfil):
        print(f"FEIL: Filen '{innfil}' ble ikke funnet.")
        sys.exit(1)

    if not innfil.lower().endswith(".pdf"):
        print("FEIL: Filen ma vaere en PDF (.pdf)")
        sys.exit(1)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("FEIL: ANTHROPIC_API_KEY er ikke satt.")
        print("Sett miljovariabelen med: set ANTHROPIC_API_KEY=din-nokkel-her")
        sys.exit(1)

    print("\n===== KABI AUTOMATION — PDF-ekstraktor =====\n")

    # Les PDF
    try:
        pdf_tekst = les_pdf(innfil)
    except Exception as e:
        print(f"FEIL: Kunne ikke lese PDF: {e}")
        sys.exit(1)

    if not pdf_tekst.strip():
        print("FEIL: Ingen tekst funnet i PDF-en.")
        print("Merk: Scannede PDF-er (bilder) krever OCR-programvare.")
        sys.exit(1)

    # Koble til Claude
    try:
        klient = anthropic.Anthropic(api_key=api_key)
    except Exception as e:
        print(f"FEIL: Kunne ikke opprette Anthropic-klient: {e}")
        sys.exit(1)

    # Ekstraher med AI
    try:
        raa_svar = ekstraher_med_claude(klient, pdf_tekst)
    except anthropic.AuthenticationError:
        print("\nFEIL: Ugyldig API-nokkel.")
        print("Kontroller ANTHROPIC_API_KEY pa: https://console.anthropic.com/")
        sys.exit(1)
    except Exception as e:
        print(f"\nFEIL ved Claude-analyse: {e}")
        sys.exit(1)

    # Parse JSON
    data = ekstraher_json_fra_tekst(raa_svar)

    if not data:
        print("\nAdvarsel: Kunne ikke parse JSON fra AI-svaret.")
        print("Raa AI-svar:")
        print(raa_svar)
        # Lagre raa svar som fallback
        base = os.path.splitext(innfil)[0]
        with open(base + "_raa_ekstrakt.txt", "w", encoding="utf-8") as f:
            f.write(raa_svar)
        print(f"\nRaat svar lagret til: {base}_raa_ekstrakt.txt")
        sys.exit(1)

    # Vis resultater
    skriv_ut_felt(data)

    # Lagre filer
    base = os.path.splitext(innfil)[0]
    json_fil = lagre_json(data, base)
    excel_fil = lagre_excel(data, base)

    print("\n===== FERDIG =====")
    print(f"PDF analysert:  {innfil}")
    print(f"JSON-fil:       {json_fil}")
    print(f"Excel-fil:      {excel_fil}")


if __name__ == "__main__":
    main()
