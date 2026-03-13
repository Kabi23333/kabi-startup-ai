"""
chatbot.py — FAQ-basert kundeservice-chatbot pa norsk
Bruk: python chatbot.py faq.txt
Krever: anthropic
"""

import sys
import os
import anthropic


MAX_HISTORIKK = 10  # Maks antall meldinger i historikken


def les_faq(filsti):
    """Les FAQ-innhold fra fil."""
    if not os.path.exists(filsti):
        print(f"FEIL: FAQ-filen '{filsti}' ble ikke funnet.")
        sys.exit(1)

    with open(filsti, "r", encoding="utf-8") as f:
        innhold = f.read()

    if not innhold.strip():
        print("FEIL: FAQ-filen er tom.")
        sys.exit(1)

    return innhold


def bygg_system_prompt(faq_innhold):
    """Bygg system prompt med FAQ-innhold."""
    return (
        "Du er en hjelpsom kundeservice-assistent for en norsk bedrift. "
        "Bruk kun informasjonen i FAQ-dokumentet nedenfor for a svare pa sporsmal. "
        "Svar alltid pa norsk. "
        "Hvis sporsmaalet ikke er besvart i FAQ-en, si at du ikke har informasjon om det "
        "og foresla at de kontakter bedriften direkte.\n\n"
        f"FAQ:\n{faq_innhold}"
    )


def send_melding(klient, system_prompt, historikk, bruker_melding):
    """Send melding til Claude og fa svar."""
    historikk.append({"role": "user", "content": bruker_melding})

    # Behold kun de siste MAX_HISTORIKK meldingene
    begrenset_historikk = historikk[-MAX_HISTORIKK:]

    svar = klient.messages.create(
        model="claude-opus-4-5",
        max_tokens=300,
        system=system_prompt,
        messages=begrenset_historikk
    )

    assistent_svar = svar.content[0].text

    # Legg til svar i full historikk
    historikk.append({"role": "assistant", "content": assistent_svar})

    return assistent_svar


def main():
    if len(sys.argv) < 2:
        print("FEIL: Ingen FAQ-fil oppgitt.")
        print("Bruk: python chatbot.py faq.txt")
        sys.exit(1)

    faq_fil = sys.argv[1]

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("FEIL: ANTHROPIC_API_KEY er ikke satt.")
        print("Sett miljovariabelen med: set ANTHROPIC_API_KEY=din-nokkel-her")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  KABI AUTOMATION — Kundeservice-assistent")
    print("=" * 60)

    faq_innhold = les_faq(faq_fil)
    system_prompt = bygg_system_prompt(faq_innhold)

    try:
        klient = anthropic.Anthropic(api_key=api_key)
    except Exception as e:
        print(f"FEIL: Kunne ikke opprette AI-tilkobling: {e}")
        sys.exit(1)

    print("\nHei! Jeg er din automatiske kundeservice-assistent.")
    print("Jeg kan hjelpe deg med sporsmal basert pa vart FAQ-dokument.")
    print("Skriv sporsmaalet ditt nedenfor, eller skriv 'avslutt' for a avslutte.\n")
    print("-" * 60)

    historikk = []

    while True:
        try:
            bruker_input = input("\nDu: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\nTakk for at du brukte kundeservice-assistenten. Ha en fin dag!")
            break

        if not bruker_input:
            print("Vennligst skriv et sporsmal.")
            continue

        if bruker_input.lower() in ["avslutt", "avslutt.", "quit", "exit", "bye"]:
            print("\nTakk for at du brukte kundeservice-assistenten. Ha en fin dag!")
            break

        try:
            print("\nAssistent: ", end="", flush=True)
            svar = send_melding(klient, system_prompt, historikk, bruker_input)
            print(svar)

        except anthropic.AuthenticationError:
            print("\nFEIL: Ugyldig API-nokkel. Kontroller ANTHROPIC_API_KEY.")
            sys.exit(1)
        except anthropic.RateLimitError:
            print("\nFor mange foresprorsler. Vent litt og prov igjen.")
        except Exception as e:
            print(f"\nFEIL: {e}")
            print("Prov igjen eller skriv 'avslutt' for a avslutte.")


if __name__ == "__main__":
    main()
