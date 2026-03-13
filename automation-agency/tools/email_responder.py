"""
email_responder.py — AI-genererte e-postsvar pa norsk
Bruk: python email_responder.py sporsmal.txt
Krever: anthropic
"""

import sys
import os
import anthropic


SYSTEM_PROMPT = (
    "Du er en profesjonell norsk kundeservicemedarbeider. "
    "Skriv et kort, hoflig og profesjonelt svar pa dette sporsmaalet fra en kunde. "
    "Svar direkte uten innledning. Maks 3 setninger."
)


def les_sporsmal(filsti):
    """Les sporsmal fra fil, ett per linje."""
    if not os.path.exists(filsti):
        print(f"FEIL: Filen '{filsti}' ble ikke funnet.")
        sys.exit(1)

    with open(filsti, "r", encoding="utf-8") as f:
        linjer = f.readlines()

    sporsmal = [linje.strip() for linje in linjer if linje.strip()]
    return sporsmal


def generer_svar(klient, sporsmal):
    """Generer svar for ett sporsmal med streaming."""
    print(f"\n{'=' * 60}")
    print(f"SPORSMAL: {sporsmal}")
    print("SVAR:")

    svar_tekst = ""

    with klient.messages.stream(
        model="claude-opus-4-5",
        max_tokens=300,
        system=SYSTEM_PROMPT,
        messages=[
            {"role": "user", "content": sporsmal}
        ]
    ) as stream:
        for tekst in stream.text_stream:
            print(tekst, end="", flush=True)
            svar_tekst += tekst

    print()  # Ny linje etter streaming
    return svar_tekst


def lagre_svar(alle_qa, utfil="email_svar.txt"):
    """Lagre alle Q&A-par til tekstfil."""
    with open(utfil, "w", encoding="utf-8") as f:
        f.write("E-POSTSVAR GENERERT AV KABI AUTOMATION\n")
        f.write(f"Antall sporsmal besvart: {len(alle_qa)}\n")
        f.write("=" * 60 + "\n\n")

        for i, (sporsmal, svar) in enumerate(alle_qa, 1):
            f.write(f"SPORSMAL {i}:\n{sporsmal}\n\n")
            f.write(f"SVAR {i}:\n{svar}\n\n")
            f.write("-" * 60 + "\n\n")

    print(f"\nAlle svar lagret til: {utfil}")


def main():
    if len(sys.argv) < 2:
        print("FEIL: Ingen fil oppgitt.")
        print("Bruk: python email_responder.py sporsmal.txt")
        sys.exit(1)

    innfil = sys.argv[1]

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("FEIL: ANTHROPIC_API_KEY er ikke satt.")
        print("Loesning: Sett miljovariabelen med:")
        print("  Windows: set ANTHROPIC_API_KEY=din-nokkel-her")
        print("  Mac/Linux: export ANTHROPIC_API_KEY=din-nokkel-her")
        sys.exit(1)

    print("\n===== KABI AUTOMATION — E-postsvar-generator =====\n")
    print(f"Leser sporsmal fra: {innfil}")

    sporsmal_liste = les_sporsmal(innfil)
    print(f"Fant {len(sporsmal_liste)} sporsmal\n")

    try:
        klient = anthropic.Anthropic(api_key=api_key)
    except Exception as e:
        print(f"FEIL: Kunne ikke opprette Anthropic-klient: {e}")
        sys.exit(1)

    alle_qa = []

    for i, sporsmal in enumerate(sporsmal_liste, 1):
        print(f"\nBehandler sporsmal {i} av {len(sporsmal_liste)}...")
        try:
            svar = generer_svar(klient, sporsmal)
            alle_qa.append((sporsmal, svar))
        except anthropic.AuthenticationError:
            print("\nFEIL: Ugyldig API-nokkel.")
            print("Kontroller at ANTHROPIC_API_KEY er riktig.")
            print("Du kan hente en ny nokkel pa: https://console.anthropic.com/")
            sys.exit(1)
        except anthropic.RateLimitError:
            print("\nFEIL: For mange foresprorsler. Vent litt og prov igjen.")
            break
        except Exception as e:
            print(f"\nFEIL ved sporsmal {i}: {e}")
            alle_qa.append((sporsmal, f"[Feil: Kunne ikke generere svar — {e}]"))

    if alle_qa:
        lagre_svar(alle_qa)

    print("\n===== FERDIG =====")
    print(f"Behandlet {len(alle_qa)} sporsmal")


if __name__ == "__main__":
    main()
