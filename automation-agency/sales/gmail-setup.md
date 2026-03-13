# Gmail-oppsett for KABI AUTOMATION

En enkel steg-for-steg guide til å sette opp Gmail profesjonelt.

---

## Steg 1 — Opprett en profesjonell Gmail-konto

Du har to alternativer:

### Alternativ A: Gratis Gmail-konto (enklest å starte med)
Bruk en adresse som ser profesjonell ut — ikke `karlobikic123@gmail.com`.

**Anbefalte adresser:**
- `kontakt.kabiautomation@gmail.com`
- `karlo.kabiautomation@gmail.com`
- `hei.kabiautomation@gmail.com`

**Slik gjør du det:**
1. Gå til [gmail.com](https://gmail.com)
2. Klikk **Opprett konto**
3. Fyll inn fornavn: `Karlo`, etternavn: `KABI Automation`
4. Velg e-postadresse (se forslag over)
5. Velg et sterkt passord (minst 12 tegn, tall og symboler)
6. Fullfør registreringen

### Alternativ B: Eget domene via Google Workspace (mer profesjonelt)
Bruk `karlo@kabiautomation.com` eller `kontakt@kabiautomation.com`.

**Krav:** Du må eie domenet `kabiautomation.com` (kjøp på f.eks. [domains.google.com](https://domains.google.com) for ca. 120-150 kr/år).

**Koster:** Google Workspace Business Starter = **ca. 68 kr/måned**

**Slik gjør du det:**
1. Kjøp domenet på [domains.google.com](https://domains.google.com)
2. Gå til [workspace.google.com](https://workspace.google.com)
3. Klikk **Start gratis prøveperiode** (14 dager gratis)
4. Følg veiviseren — du kobler domenet til Google Workspace
5. Din e-post blir `karlo@kabiautomation.com`

> **Anbefaling:** Start med Alternativ A og oppgrader til B når du har første betalende kunde.

---

## Steg 2 — Legg inn HTML-signaturen i Gmail

1. Åpne filen `signatur.html` i en nettleser (dobbeltklikk på filen)
2. Trykk **Ctrl+A** (Windows) eller **Cmd+A** (Mac) for å markere alt
3. Trykk **Ctrl+C** / **Cmd+C** for å kopiere
4. Åpne Gmail i nettleseren
5. Klikk på **tannhjulikonet ⚙️** øverst til høyre → **Se alle innstillinger**
6. Gå til fanen **Generelt**
7. Scroll ned til **Signatur**
8. Klikk **Opprett ny signatur**
9. Gi den et navn, f.eks. `KABI Automation`
10. Klikk i tekstfeltet og trykk **Ctrl+V** / **Cmd+V** for å lime inn
11. Under **Standardsignatur**, velg signaturen din for:
    - **Ny e-post:** KABI Automation
    - **Svar / videresending:** KABI Automation
12. Scroll ned og klikk **Lagre endringer**

> **Tips:** Åpne en ny e-post og se at signaturen vises korrekt før du sender noe.

---

## Steg 3 — Sett opp "Send som" med eget domene (valgfritt)

Hvis du kjøpte eget domene men bruker en vanlig Gmail-konto, kan du sende e-post som om det kom fra `karlo@kabiautomation.com`.

**Forutsetning:** Du må konfigurere e-posthosting hos domeneregistraren din (f.eks. Domeneshop, One.com, eller Cloudflare Email Routing).

**Slik gjør du det:**
1. Gå til Gmail → **Innstillinger ⚙️** → **Se alle innstillinger**
2. Klikk på fanen **Kontoer og import**
3. Under **Send e-post som**, klikk **Legg til en annen e-postadresse**
4. Fyll inn:
   - **Navn:** Karlo Bikic — KABI AUTOMATION
   - **E-postadresse:** `karlo@kabiautomation.com`
5. Klikk **Neste steg**
6. Velg **Send via Gmail** (enklest) eller konfigurer SMTP fra domenet ditt
7. Gmail sender en bekreftelseskode til den nye adressen — bekreft den
8. Nå kan du velge hvilken adresse du sender fra i nye e-poster

> **Gratis alternativ:** Bruk [Cloudflare Email Routing](https://cloudflare.com) (gratis) til å videresende e-post fra eget domene til Gmail.

---

## Steg 4 — Aktiver 2-faktor autentisering (2FA)

Dette beskytter kontoen din mot hackere. **Gjør dette nå — ikke vent.**

1. Gå til [myaccount.google.com/security](https://myaccount.google.com/security)
2. Under **Slik logger du på Google**, klikk **2-trinns bekreftelse**
3. Klikk **Kom i gang**
4. Følg veiviseren:
   - **Anbefalt:** Google Authenticator-appen (last ned på mobilen)
   - **Alternativ:** SMS-kode til telefonnummer ditt
5. Bekreft at du kan logge inn med den nye metoden

> **OBS:** Skriv ned backup-kodene og lagre dem på et trygt sted — du trenger dem hvis du mister telefonen.

---

## Steg 5 — Sett opp Gmail-filter mot spam

Når kunder begynner å svare, vil du sikre at e-postene ikke havner i spam-mappen.

### Filter 5A: Sørg for at svar på dine egne e-poster aldri blir spam

1. Gå til Gmail → **Innstillinger ⚙️** → **Se alle innstillinger**
2. Klikk på fanen **Filtre og blokkerte adresser**
3. Klikk **Opprett et nytt filter**
4. I feltet **Fra**, skriv `@gmail.com OR @hotmail.com OR @yahoo.com OR @outlook.com`
5. Klikk **Opprett filter med dette søket**
6. Huk av:
   - ☑ Merk aldri som spam
   - ☑ Bruk aldri automatisk svar fra (fravær)
7. Klikk **Opprett filter**

### Filter 5B: Etikett for KABI Automation-svar

Slik holder du oversikt over alle svar fra potensielle kunder:

1. Gå til **Innstillinger** → **Filtre og blokkerte adresser** → **Opprett nytt filter**
2. I **Emnelinjen**, skriv: `KABI Automation`
3. Klikk **Opprett filter**
4. Huk av:
   - ☑ Bruk etiketten → Velg **Ny etikett** → skriv `KABI Leads`
   - ☑ Merk aldri som spam
5. Klikk **Opprett filter**

> Nå samles alle svar automatisk i mappen **KABI Leads** — lett å holde oversikt!

---

## Rask oppsummering — gjør disse 5 tingene i dag

| # | Oppgave | Tid | Status |
|---|---------|-----|--------|
| 1 | Opprett profesjonell Gmail-konto | 5 min | ☐ |
| 2 | Legg inn HTML-signaturen | 5 min | ☐ |
| 3 | Sett opp "Send som" eget domene | 15 min | ☐ (valgfritt) |
| 4 | Aktiver 2-faktor autentisering | 3 min | ☐ |
| 5 | Sett opp filter mot spam | 5 min | ☐ |

**Total tid: ca. 18 minutter** — og du har et profesjonelt e-postsystem klart.

---

## Neste steg etter Gmail-oppsett

1. Kjør `python send_outreach.py` for å starte outreach til kunder
2. Sjekk innboksen daglig for svar
3. Svar innen 24 timer på alle henvendelser
4. Book møter — 20 minutter er nok til å demonstrere verdien

---

*Laget for KABI AUTOMATION — Karlo Bikic*
