"""
KABI STARTUP AI - Finansrapportgenerator
Leser Excel-fil med inntekter og utgifter og genererer en norsk finansrapport
med AI-drevet innsikt via Claude API.

Forventet Excel-format (kolonner):
  Dato        | Beskrivelse       | Kategori     | Beløp
  2024-01-15  | Faktura kunde A   | Inntekt      | 15000
  2024-01-20  | Husleie kontor    | Utgift       | -3500

Krev miljøvariabel: ANTHROPIC_API_KEY
"""

import sys
from datetime import datetime
import pandas as pd
import anthropic

# Sikrer UTF-8 output i Windows-terminal
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


KATEGORI_EMOJIS = {
    "inntekt": "💰",
    "salg": "💰",
    "faktura": "💰",
    "lønn": "💸",
    "husleie": "🏠",
    "markedsføring": "📣",
    "programvare": "💻",
    "transport": "🚗",
    "mat": "🍽️",
    "forsikring": "🛡️",
    "skatt": "🏛️",
    "annet": "📦",
}


def les_excel(filsti: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(filsti)
    except FileNotFoundError:
        print(f"❌ Fant ikke filen: {filsti}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Kunne ikke lese Excel-filen: {e}")
        sys.exit(1)

    # Normaliser kolonnenavn
    df.columns = [col.strip().lower() for col in df.columns]

    påkrevde = {"dato", "beskrivelse", "kategori", "beløp"}
    mangler = påkrevde - set(df.columns)
    if mangler:
        print(f"❌ Mangler kolonner i Excel-filen: {', '.join(mangler)}")
        print(f"   Forventede kolonner: Dato, Beskrivelse, Kategori, Beløp")
        sys.exit(1)

    df["beløp"] = pd.to_numeric(df["beløp"], errors="coerce").fillna(0)
    df["dato"] = pd.to_datetime(df["dato"], errors="coerce")
    df["kategori"] = df["kategori"].astype(str).str.strip()
    return df


def emoji_for_kategori(kategori: str) -> str:
    k = kategori.lower()
    for nøkkel, emoji in KATEGORI_EMOJIS.items():
        if nøkkel in k:
            return emoji
    return "📦"


def hent_ai_innsikt(
    df: pd.DataFrame,
    total_inntekt: float,
    total_utgift: float,
    resultat: float,
    margin: float,
) -> str:
    """Send finansdata til Claude og få tilbake norsk AI-innsikt via streaming."""

    inntekter = df[df["beløp"] > 0]
    utgifter = df[df["beløp"] < 0]

    # Bygg opp en kompakt dataoversikt til Claude
    inntekt_oversikt = (
        inntekter.groupby("kategori")["beløp"]
        .sum()
        .sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr")
        .to_string()
    )
    utgift_oversikt = (
        utgifter.groupby("kategori")["beløp"]
        .sum()
        .abs()
        .sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr")
        .to_string()
    )

    periode_fra = df["dato"].min()
    periode_til = df["dato"].max()
    periode = ""
    if pd.notna(periode_fra) and pd.notna(periode_til):
        periode = f"{periode_fra.strftime('%d.%m.%Y')} – {periode_til.strftime('%d.%m.%Y')}"

    prompt = f"""Du er en erfaren norsk regnskapsfører og forretningsrådgiver.
Analyser følgende finansdata for en liten norsk bedrift og gi konkret, praktisk innsikt på norsk.

PERIODE: {periode or 'Ukjent'}
TOTALE INNTEKTER: {total_inntekt:,.0f} kr
TOTALE UTGIFTER: {total_utgift:,.0f} kr
RESULTAT: {'+' if resultat >= 0 else ''}{resultat:,.0f} kr
FORTJENESTEMARGIN: {margin:.1f}%

INNTEKTER PER KATEGORI:
{inntekt_oversikt}

UTGIFTER PER KATEGORI:
{utgift_oversikt}

Gi en analyse på 4–6 setninger som inkluderer:
1. Vurdering av lønnsomhet og finansiell helse
2. Den viktigste styrken og den viktigste risikoen
3. Ett konkret, handlingsrettet råd for å forbedre resultatet

Svar direkte og enkelt — eieren er ikke regnskapsfaglig. Unngå fagsjargong."""

    try:
        claude = anthropic.Anthropic()
        tekst_deler = []

        print("  ⏳ Henter AI-innsikt fra Claude...", end="\r")

        with claude.messages.stream(
            model="claude-opus-4-6",
            max_tokens=512,
            thinking={"type": "adaptive"},
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            for delta in stream.text_stream:
                tekst_deler.append(delta)

        print("                                      ", end="\r")  # fjern spinner
        return "".join(tekst_deler).strip()

    except anthropic.AuthenticationError:
        return "⚠️  ANTHROPIC_API_KEY mangler eller er ugyldig. Sett miljøvariabelen for AI-innsikt."
    except Exception as e:
        return f"⚠️  Kunne ikke hente AI-innsikt: {e}"


def generer_rapport(df: pd.DataFrame, ai_tekst: str) -> str:
    inntekter = df[df["beløp"] > 0]
    utgifter = df[df["beløp"] < 0]

    total_inntekt = inntekter["beløp"].sum()
    total_utgift = abs(utgifter["beløp"].sum())
    resultat = total_inntekt - total_utgift
    margin = (resultat / total_inntekt * 100) if total_inntekt > 0 else 0

    periode_fra = df["dato"].min()
    periode_til = df["dato"].max()
    periode_str = ""
    if pd.notna(periode_fra) and pd.notna(periode_til):
        periode_str = f"{periode_fra.strftime('%d.%m.%Y')} – {periode_til.strftime('%d.%m.%Y')}"

    linjer = []
    skillelinje = "═" * 52

    linjer.append(skillelinje)
    linjer.append("         KABI STARTUP AI — FINANSRAPPORT")
    if periode_str:
        linjer.append(f"         Periode: {periode_str}")
    linjer.append(f"         Generert: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    linjer.append(skillelinje)

    # Sammendrag
    linjer.append("")
    linjer.append("📊 SAMMENDRAG")
    linjer.append("─" * 52)
    linjer.append(f"  Totale inntekter:     {total_inntekt:>12,.0f} kr")
    linjer.append(f"  Totale utgifter:      {total_utgift:>12,.0f} kr")
    linjer.append("  " + "·" * 48)
    if resultat >= 0:
        linjer.append(f"  ✅ Overskudd:          {resultat:>12,.0f} kr")
    else:
        linjer.append(f"  ❌ Underskudd:         {abs(resultat):>12,.0f} kr")
    linjer.append(f"  Fortjenestemargin:    {margin:>11.1f} %")

    # Inntekter per kategori
    if not inntekter.empty:
        linjer.append("")
        linjer.append("💰 INNTEKTER PER KATEGORI")
        linjer.append("─" * 52)
        for kat, gruppe in inntekter.groupby("kategori"):
            total = gruppe["beløp"].sum()
            antall = len(gruppe)
            emoji = emoji_for_kategori(kat)
            linjer.append(f"  {emoji} {kat:<28} {total:>10,.0f} kr  ({antall} poster)")

    # Utgifter per kategori
    if not utgifter.empty:
        linjer.append("")
        linjer.append("💸 UTGIFTER PER KATEGORI")
        linjer.append("─" * 52)
        utgift_per_kat = utgifter.groupby("kategori")["beløp"].agg(["sum", "count"])
        utgift_per_kat = utgift_per_kat.sort_values("sum")
        for kat, rad in utgift_per_kat.iterrows():
            total = abs(rad["sum"])
            antall = int(rad["count"])
            andel = (total / total_utgift * 100) if total_utgift > 0 else 0
            emoji = emoji_for_kategori(kat)
            linjer.append(f"  {emoji} {kat:<22} {total:>10,.0f} kr  ({andel:.0f}%)")

    # Månedlig oversikt (hvis data strekker seg over flere måneder)
    df_med_dato = df[df["dato"].notna()].copy()
    if not df_med_dato.empty:
        df_med_dato["måned"] = df_med_dato["dato"].dt.to_period("M")
        måneder = df_med_dato["måned"].nunique()
        if måneder > 1:
            linjer.append("")
            linjer.append("📅 MÅNEDLIG OVERSIKT")
            linjer.append("─" * 52)
            månedlig = df_med_dato.groupby("måned")["beløp"].agg(
                inntekt=lambda x: x[x > 0].sum(),
                utgift=lambda x: abs(x[x < 0].sum()),
            )
            månedlig["resultat"] = månedlig["inntekt"] - månedlig["utgift"]
            for måned, rad in månedlig.iterrows():
                tegn = "✅" if rad["resultat"] >= 0 else "❌"
                linjer.append(
                    f"  {str(måned):<10}  Innt: {rad['inntekt']:>8,.0f}  "
                    f"Utg: {rad['utgift']:>8,.0f}  {tegn} {rad['resultat']:>8,.0f} kr"
                )

    # AI-innsikt
    linjer.append("")
    linjer.append("🤖 AI-INNSIKT (Claude)")
    linjer.append("─" * 52)
    for linje in ai_tekst.splitlines():
        linjer.append(f"  {linje}")

    linjer.append("")
    linjer.append(skillelinje)
    linjer.append("  Rapporten er generert av KABI STARTUP AI")
    linjer.append(skillelinje)

    return "\n".join(linjer)


def eksporter_pdf(df: pd.DataFrame, ai_tekst: str, ut_fil: str) -> None:
    """Eksporter finansrapporten som en formatert PDF."""
    from fpdf import FPDF

    inntekter = df[df["beløp"] > 0]
    utgifter = df[df["beløp"] < 0]
    total_inntekt = inntekter["beløp"].sum()
    total_utgift = abs(utgifter["beløp"].sum())
    resultat = total_inntekt - total_utgift
    margin = (resultat / total_inntekt * 100) if total_inntekt > 0 else 0

    periode_fra = df["dato"].min()
    periode_til = df["dato"].max()
    periode_str = ""
    if pd.notna(periode_fra) and pd.notna(periode_til):
        periode_str = f"{periode_fra.strftime('%d.%m.%Y')} - {periode_til.strftime('%d.%m.%Y')}"

    # Farger (R, G, B)
    BLA = (30, 64, 175)
    GRA = (107, 114, 128)
    LYS_GRA = (243, 244, 246)
    GRONN = (22, 163, 74)
    ROED = (220, 38, 38)
    HVIT = (255, 255, 255)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # ── Header ──────────────────────────────────────────
    pdf.set_fill_color(*BLA)
    pdf.rect(0, 0, 210, 38, "F")
    pdf.set_text_color(*HVIT)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(10, 7)
    pdf.cell(190, 12, "KABI STARTUP AI", align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_xy(10, 21)
    pdf.cell(190, 8, "Finansrapport", align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.set_y(44)

    # Periode / dato
    if periode_str:
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(*GRA)
        pdf.cell(
            0, 5,
            f"Periode: {periode_str}     Generert: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            align="R",
        )
        pdf.ln(8)
    pdf.set_text_color(0, 0, 0)

    # ── Nøkkeltall-bokser ────────────────────────────────
    def boks(x, y, w, tittel, verdi, farge):
        pdf.set_fill_color(*LYS_GRA)
        pdf.rect(x, y, w, 22, "F")
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(*GRA)
        pdf.set_xy(x + 3, y + 3)
        pdf.cell(w - 6, 5, tittel)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(*farge)
        pdf.set_xy(x + 3, y + 10)
        pdf.cell(w - 6, 8, verdi)
        pdf.set_text_color(0, 0, 0)

    y0 = pdf.get_y()
    boks(10, y0, 44, "INNTEKTER", f"{total_inntekt:,.0f} kr", GRONN)
    boks(58, y0, 44, "UTGIFTER", f"{total_utgift:,.0f} kr", ROED)
    res_farge = GRONN if resultat >= 0 else ROED
    boks(106, y0, 44, "RESULTAT", f"{'+' if resultat >= 0 else ''}{resultat:,.0f} kr", res_farge)
    mar_farge = GRONN if margin >= 10 else ROED
    boks(154, y0, 44, "MARGIN", f"{margin:.1f}%", mar_farge)
    pdf.set_y(y0 + 28)

    # ── Hjelpefunksjoner ─────────────────────────────────
    def seksjon(tittel):
        pdf.ln(3)
        pdf.set_draw_color(*BLA)
        pdf.set_line_width(0.5)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*BLA)
        pdf.cell(0, 7, tittel)
        pdf.ln(8)
        pdf.set_text_color(0, 0, 0)
        pdf.set_line_width(0.2)

    def tabell_rad(kol1, kol2, kol3="", fet=False, skygge=False):
        if skygge:
            pdf.set_fill_color(*LYS_GRA)
            pdf.rect(10, pdf.get_y(), 190, 7, "F")
        pdf.set_font("Helvetica", "B" if fet else "", 9)
        pdf.set_x(12)
        pdf.cell(100, 7, str(kol1))
        pdf.cell(50, 7, str(kol2), align="R")
        pdf.cell(38, 7, str(kol3), align="R")
        pdf.ln()

    # ── Inntekter ────────────────────────────────────────
    if not inntekter.empty:
        seksjon("Inntekter per kategori")
        for i, (kat, grp) in enumerate(inntekter.groupby("kategori")):
            tabell_rad(kat, f"{grp['beloep'].sum():,.0f} kr" if "beloep" in grp else f"{grp['beløp'].sum():,.0f} kr",
                       f"{len(grp)} poster", skygge=(i % 2 == 0))
        tabell_rad("Totalt", f"{total_inntekt:,.0f} kr", fet=True)
        pdf.ln(2)

    # ── Utgifter ─────────────────────────────────────────
    if not utgifter.empty:
        seksjon("Utgifter per kategori")
        utgift_kat = utgifter.groupby("kategori")["beløp"].agg(["sum", "count"]).sort_values("sum")
        for i, (kat, r) in enumerate(utgift_kat.iterrows()):
            andel = abs(r["sum"]) / total_utgift * 100 if total_utgift > 0 else 0
            tabell_rad(kat, f"{abs(r['sum']):,.0f} kr", f"{andel:.0f}%", skygge=(i % 2 == 0))
        tabell_rad("Totalt", f"{total_utgift:,.0f} kr", fet=True)
        pdf.ln(2)

    # ── Maanedlig oversikt ───────────────────────────────
    df_dato = df[df["dato"].notna()].copy()
    if not df_dato.empty:
        df_dato["maned"] = df_dato["dato"].dt.to_period("M")
        if df_dato["maned"].nunique() > 1:
            seksjon("Manedlig oversikt")
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_x(12)
            pdf.cell(50, 7, "Maned")
            pdf.cell(50, 7, "Inntekter", align="R")
            pdf.cell(50, 7, "Utgifter", align="R")
            pdf.cell(38, 7, "Resultat", align="R")
            pdf.ln()
            manedlig = df_dato.groupby("maned")["beløp"].agg(
                inntekt=lambda x: x[x > 0].sum(),
                utgift=lambda x: abs(x[x < 0].sum()),
            )
            manedlig["res"] = manedlig["inntekt"] - manedlig["utgift"]
            for i, (maned, r) in enumerate(manedlig.iterrows()):
                if i % 2 == 0:
                    pdf.set_fill_color(*LYS_GRA)
                    pdf.rect(10, pdf.get_y(), 190, 7, "F")
                pdf.set_font("Helvetica", "", 9)
                pdf.set_x(12)
                pdf.cell(50, 7, str(maned))
                pdf.cell(50, 7, f"{r['inntekt']:,.0f} kr", align="R")
                pdf.cell(50, 7, f"{r['utgift']:,.0f} kr", align="R")
                pdf.set_text_color(*(GRONN if r["res"] >= 0 else ROED))
                pdf.cell(38, 7, f"{'+' if r['res'] >= 0 else ''}{r['res']:,.0f} kr", align="R")
                pdf.set_text_color(0, 0, 0)
                pdf.ln()
            pdf.ln(2)

    # ── AI-innsikt ───────────────────────────────────────
    seksjon("AI-innsikt (Claude Opus)")
    pdf.set_fill_color(239, 246, 255)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(30, 58, 138)
    # multi_cell med fill tegner bakgrunn per linje
    pdf.set_x(12)
    pdf.multi_cell(186, 6, ai_tekst or "Ingen AI-innsikt tilgjengelig.", fill=True)
    pdf.set_text_color(0, 0, 0)

    # ── Footer ───────────────────────────────────────────
    pdf.set_y(-14)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(*GRA)
    pdf.cell(0, 5, "Generert av KABI STARTUP AI", align="C")

    pdf.output(ut_fil)


def main():
    if len(sys.argv) < 2:
        print("Bruk: python finansrapport.py <excel-fil.xlsx>")
        print("Eksempel: python finansrapport.py data.xlsx")
        print("\nTips: Kjør 'python lag_testdata.py' for å lage en testfil.")
        sys.exit(1)

    filsti = sys.argv[1]
    df = les_excel(filsti)

    # Beregn nøkkeltall for AI-kallet
    inntekter = df[df["beløp"] > 0]
    utgifter = df[df["beløp"] < 0]
    total_inntekt = inntekter["beløp"].sum()
    total_utgift = abs(utgifter["beløp"].sum())
    resultat = total_inntekt - total_utgift
    margin = (resultat / total_inntekt * 100) if total_inntekt > 0 else 0

    ai_tekst = hent_ai_innsikt(df, total_inntekt, total_utgift, resultat, margin)

    rapport = generer_rapport(df, ai_tekst)
    print(rapport)

    base = filsti.replace(".xlsx", "").replace(".xls", "")

    # Tekstfil
    txt_fil = base + "_rapport.txt"
    with open(txt_fil, "w", encoding="utf-8") as f:
        f.write(rapport)
    print(f"\n📄 Tekstrapport lagret til: {txt_fil}")

    # PDF
    try:
        pdf_fil = base + "_rapport.pdf"
        eksporter_pdf(df, ai_tekst, pdf_fil)
        print(f"📑 PDF-rapport lagret til:  {pdf_fil}")
    except ImportError:
        print("⚠️  PDF-eksport krever fpdf2: pip install fpdf2")


if __name__ == "__main__":
    main()
