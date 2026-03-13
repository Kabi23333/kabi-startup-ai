"""
KABI STARTUP AI - Streamlit-app
Kjoer med: streamlit run app.py
"""

import os
import tempfile

import anthropic
import pandas as pd
import streamlit as st

from finansrapport import eksporter_pdf, generer_rapport

# ── Sidekonfigurasjon ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KABI STARTUP AI",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .kabi-header {
        background: linear-gradient(135deg, #1e40af, #3b82f6);
        padding: 2rem; border-radius: 12px; color: white;
        text-align: center; margin-bottom: 1.5rem;
    }
    .kabi-header h1 { margin: 0; font-size: 2.2rem; }
    .kabi-header p  { margin: 0.3rem 0 0; opacity: 0.85; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── API-nokkel: Streamlit Secrets → miljovariabel → manuell input ─────────────
try:
    if "ANTHROPIC_API_KEY" in st.secrets:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    pass

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Slik bruker du appen")
    st.markdown(
        "1. Last opp en Excel-fil (.xlsx)\n"
        "2. Rapporten genereres automatisk\n"
        "3. Last ned som **PDF** eller **TXT**"
    )
    st.divider()
    st.markdown("**Forventet kolonneformat:**")
    st.code("Dato | Beskrivelse | Kategori | Belop", language=None)
    st.caption("Positive belop = inntekter\nNegative belop = utgifter")
    st.divider()

    har_nokkel = bool(os.environ.get("ANTHROPIC_API_KEY"))
    api_input = st.text_input(
        "Anthropic API-nokkel",
        type="password",
        placeholder="Allerede satt" if har_nokkel else "Lim inn nokkel her",
        help="Trengs for AI-innsikt fra Claude.",
    )
    if api_input:
        os.environ["ANTHROPIC_API_KEY"] = api_input
        har_nokkel = True

    if not har_nokkel:
        st.warning("Uten API-nokkel vises ikke AI-innsikt.")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="kabi-header">'
    "<h1>KABI STARTUP AI</h1>"
    "<p>Finansrapportgenerator for norske smabedrifter</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ── Filopplasting ─────────────────────────────────────────────────────────────
fil = st.file_uploader(
    "Last opp Excel-filen din",
    type=["xlsx", "xls"],
    help="Excel-fil med kolonnene: Dato, Beskrivelse, Kategori, Belop",
)

if fil is None:
    st.info("Last opp en Excel-fil for a komme i gang.")
    st.stop()

# ── Les og valider ────────────────────────────────────────────────────────────
try:
    df = pd.read_excel(fil)
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

# Normaliser kolonnenavn til lowercase
df.columns = [c.strip().lower() for c in df.columns]

# Aksepter varianter av "belop" (beloep, belop, beløp)
belop_kandidater = [
    c for c in df.columns
    if "bel" in c and any(x in c for x in ("p", "o", "\u00f8"))
]
if belop_kandidater:
    beste = belop_kandidater[0]
    if beste != "bel\u00f8p":
        df = df.rename(columns={beste: "bel\u00f8p"})

mangler = {"dato", "beskrivelse", "kategori", "bel\u00f8p"} - set(df.columns)
if mangler:
    st.error(f"Mangler kolonner: **{', '.join(mangler)}**")
    st.caption("Forventede kolonner: Dato, Beskrivelse, Kategori, Belop")
    st.stop()

df["bel\u00f8p"] = pd.to_numeric(df["bel\u00f8p"], errors="coerce").fillna(0)
df["dato"] = pd.to_datetime(df["dato"], errors="coerce")
df["kategori"] = df["kategori"].astype(str).str.strip()

inntekter = df[df["bel\u00f8p"] > 0]
utgifter = df[df["bel\u00f8p"] < 0]
total_inntekt = inntekter["bel\u00f8p"].sum()
total_utgift = abs(utgifter["bel\u00f8p"].sum())
resultat = total_inntekt - total_utgift
margin = (resultat / total_inntekt * 100) if total_inntekt > 0 else 0

# ── Nokkeltall ────────────────────────────────────────────────────────────────
st.markdown("### Sammendrag")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Inntekter", f"{total_inntekt:,.0f} kr")
k2.metric("Utgifter", f"{total_utgift:,.0f} kr")
k3.metric(
    "Resultat",
    f"{'+' if resultat >= 0 else ''}{resultat:,.0f} kr",
    delta="Overskudd" if resultat >= 0 else "Underskudd",
    delta_color="normal" if resultat >= 0 else "inverse",
)
k4.metric("Margin", f"{margin:.1f}%")

# ── Diagrammer ────────────────────────────────────────────────────────────────
st.markdown("### Fordeling")
col_v, col_h = st.columns(2)

with col_v:
    st.markdown("**Inntekter per kategori**")
    if not inntekter.empty:
        data = inntekter.groupby("kategori")["bel\u00f8p"].sum().sort_values(ascending=False)
        st.bar_chart(data)

with col_h:
    st.markdown("**Utgifter per kategori**")
    if not utgifter.empty:
        data = utgifter.groupby("kategori")["bel\u00f8p"].abs().sum().sort_values(ascending=False)
        st.bar_chart(data)

# ── Manedlig oversikt ─────────────────────────────────────────────────────────
df_dato = df[df["dato"].notna()].copy()
if not df_dato.empty:
    df_dato["Maned"] = df_dato["dato"].dt.to_period("M").astype(str)
    manedlig = df_dato.groupby("Maned")["bel\u00f8p"].agg(
        Inntekter=lambda x: x[x > 0].sum(),
        Utgifter=lambda x: abs(x[x < 0].sum()),
    )
    manedlig["Resultat"] = manedlig["Inntekter"] - manedlig["Utgifter"]

    if len(manedlig) > 1:
        st.markdown("### Manedlig oversikt")
        styled = (
            manedlig.style
            .format("{:,.0f} kr")
            .map(
                lambda v: "color: green" if v >= 0 else "color: red",
                subset=["Resultat"],
            )
        )
        st.dataframe(styled, use_container_width=True)

# ── AI-innsikt ────────────────────────────────────────────────────────────────
st.markdown("### AI-innsikt fra Claude")

fil_nokkel = f"{fil.name}_{fil.size}"


def bygg_prompt() -> str:
    inntekt_oversikt = (
        inntekter.groupby("kategori")["bel\u00f8p"]
        .sum().sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr").to_string()
    )
    utgift_oversikt = (
        utgifter.groupby("kategori")["bel\u00f8p"]
        .sum().abs().sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr").to_string()
    )
    periode_fra = df["dato"].min()
    periode_til = df["dato"].max()
    periode = ""
    if pd.notna(periode_fra) and pd.notna(periode_til):
        periode = (
            f"{periode_fra.strftime('%d.%m.%Y')} - {periode_til.strftime('%d.%m.%Y')}"
        )
    return (
        "Du er en erfaren norsk regnskapsforer og forretningsradgiver. "
        "Analyser folgende finansdata og gi konkret innsikt pa norsk.\n\n"
        f"PERIODE: {periode or 'Ukjent'}\n"
        f"TOTALE INNTEKTER: {total_inntekt:,.0f} kr\n"
        f"TOTALE UTGIFTER: {total_utgift:,.0f} kr\n"
        f"RESULTAT: {'+' if resultat >= 0 else ''}{resultat:,.0f} kr\n"
        f"FORTJENESTEMARGIN: {margin:.1f}%\n\n"
        f"INNTEKTER:\n{inntekt_oversikt}\n\n"
        f"UTGIFTER:\n{utgift_oversikt}\n\n"
        "Gi 4-6 setninger: vurder lonnsomhet, nevn styrke og risiko, "
        "gi ett konkret rad. Svar enkelt - eieren er ikke regnskapsfaglig."
    )


def ai_stream():
    claude = anthropic.Anthropic()
    with claude.messages.stream(
        model="claude-opus-4-6",
        max_tokens=512,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": bygg_prompt()}],
    ) as stream:
        for delta in stream.text_stream:
            yield delta


har_nokkel_na = bool(os.environ.get("ANTHROPIC_API_KEY"))
ny_fil = st.session_state.get("fil_nokkel") != fil_nokkel

if not har_nokkel_na:
    st.info("Legg inn Anthropic API-nokkel i sidepanelet for a fa AI-innsikt.")
    ai_tekst = "AI-innsikt ikke tilgjengelig - API-nokkel mangler."
elif ny_fil:
    try:
        with st.container(border=True):
            ai_tekst = st.write_stream(ai_stream())
        st.session_state.ai_tekst = ai_tekst
        st.session_state.fil_nokkel = fil_nokkel
    except anthropic.AuthenticationError:
        st.error("Ugyldig API-nokkel. Sjekk nokkel i sidepanelet.")
        ai_tekst = "AI-innsikt ikke tilgjengelig - ugyldig API-nokkel."
    except Exception as e:
        st.warning(f"AI-analyse feilet: {e}")
        ai_tekst = "AI-innsikt ikke tilgjengelig."
else:
    with st.container(border=True):
        st.markdown(st.session_state.get("ai_tekst", ""))
    ai_tekst = st.session_state.get("ai_tekst", "")

# ── Last ned ──────────────────────────────────────────────────────────────────
st.markdown("### Last ned rapport")

dl1, dl2 = st.columns(2)

with dl1:
    try:
        txt_innhold = generer_rapport(df, ai_tekst)
        st.download_button(
            label="Last ned TXT",
            data=txt_innhold.encode("utf-8"),
            file_name="finansrapport.txt",
            mime="text/plain",
            use_container_width=True,
        )
    except Exception as e:
        st.warning(f"Teksteksport feilet: {e}")

with dl2:
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_sti = tmp.name
        eksporter_pdf(df, ai_tekst, tmp_sti)
        with open(tmp_sti, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(tmp_sti)
        st.download_button(
            label="Last ned PDF",
            data=pdf_bytes,
            file_name="finansrapport.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as e:
        st.warning(f"PDF-eksport feilet: {e}")

# ── Radata ────────────────────────────────────────────────────────────────────
with st.expander("Vis opplastet data"):
    st.dataframe(df, use_container_width=True)

st.caption("KABI STARTUP AI - bygget av Karlo Bikic")
