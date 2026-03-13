"""
KABI STARTUP AI — Streamlit-app
Kjør lokalt: streamlit run app.py
"""

import os
import sys
import tempfile

import anthropic
import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
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
    .kabi-header { background: linear-gradient(135deg, #1e40af, #3b82f6);
        padding: 2rem; border-radius: 12px; color: white; text-align: center;
        margin-bottom: 1.5rem; }
    .kabi-header h1 { margin: 0; font-size: 2.2rem; }
    .kabi-header p  { margin: 0.3rem 0 0; opacity: 0.85; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Hent API-nøkkel (Secrets → env → manuell) ────────────────────────────────
try:
    if "ANTHROPIC_API_KEY" in st.secrets:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    pass  # Ingen secrets konfigurert — går videre

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📋 Slik bruker du appen")
    st.markdown(
        "1. Last opp en Excel-fil (.xlsx)\n"
        "2. Rapporten genereres automatisk\n"
        "3. Last ned som **PDF** eller **TXT**"
    )
    st.divider()
    st.markdown("**Forventet kolonneformat:**")
    st.code("Dato | Beskrivelse | Kategori | Beløp", language=None)
    st.caption("Positive beløp = inntekter\nNegative beløp = utgifter")
    st.divider()

    har_nøkkel = bool(os.environ.get("ANTHROPIC_API_KEY"))
    api_input = st.text_input(
        "🔑 Anthropic API-nøkkel",
        type="password",
        placeholder="Allerede satt" if har_nøkkel else "Lim inn nøkkel her",
        help="Trengs for AI-innsikt fra Claude. Gratis å opprette på console.anthropic.com",
    )
    if api_input:
        os.environ["ANTHROPIC_API_KEY"] = api_input
        har_nøkkel = True

    if not har_nøkkel:
        st.warning("Uten API-nøkkel hoppes AI-innsikt over.")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="kabi-header">'
    "<h1>💰 KABI STARTUP AI</h1>"
    "<p>Finansrapportgenerator for norske småbedrifter</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ── Filopplasting ─────────────────────────────────────────────────────────────
fil = st.file_uploader(
    "Last opp Excel-filen din",
    type=["xlsx", "xls"],
    help="Excel-fil med kolonnene: Dato, Beskrivelse, Kategori, Beløp",
)

if fil is None:
    st.info("⬆️  Last opp en Excel-fil for å komme i gang.")
    st.stop()

# ── Les og valider ────────────────────────────────────────────────────────────
try:
    df = pd.read_excel(fil)
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

df.columns = [c.strip().lower() for c in df.columns]
mangler = {"dato", "beskrivelse", "kategori", "beløp"} - set(df.columns)
if mangler:
    st.error(f"❌ Mangler kolonner: **{', '.join(mangler)}**")
    st.caption("Forventede kolonner: Dato, Beskrivelse, Kategori, Beløp")
    st.stop()

df["beløp"] = pd.to_numeric(df["beløp"], errors="coerce").fillna(0)
df["dato"] = pd.to_datetime(df["dato"], errors="coerce")
df["kategori"] = df["kategori"].astype(str).str.strip()

inntekter = df[df["beløp"] > 0]
utgifter = df[df["beløp"] < 0]
total_inntekt = inntekter["beløp"].sum()
total_utgift = abs(utgifter["beløp"].sum())
resultat = total_inntekt - total_utgift
margin = (resultat / total_inntekt * 100) if total_inntekt > 0 else 0

# ── Nøkkeltall ────────────────────────────────────────────────────────────────
st.markdown("### 📊 Sammendrag")
k1, k2, k3, k4 = st.columns(4)
k1.metric("💰 Inntekter", f"{total_inntekt:,.0f} kr")
k2.metric("💸 Utgifter", f"{total_utgift:,.0f} kr")
k3.metric(
    "📈 Resultat",
    f"{'+' if resultat >= 0 else ''}{resultat:,.0f} kr",
    delta="Overskudd" if resultat >= 0 else "Underskudd",
    delta_color="normal" if resultat >= 0 else "inverse",
)
k4.metric("📉 Margin", f"{margin:.1f}%")

# ── Diagrammer ────────────────────────────────────────────────────────────────
st.markdown("### 📊 Fordeling")
col_v, col_h = st.columns(2)

with col_v:
    st.markdown("**Inntekter per kategori**")
    if not inntekter.empty:
        data = inntekter.groupby("kategori")["beløp"].sum().sort_values(ascending=False)
        st.bar_chart(data)

with col_h:
    st.markdown("**Utgifter per kategori**")
    if not utgifter.empty:
        data = utgifter.groupby("kategori")["beløp"].abs().sum().sort_values(ascending=False)
        st.bar_chart(data)

# ── Månedlig oversikt ─────────────────────────────────────────────────────────
df_dato = df[df["dato"].notna()].copy()
if not df_dato.empty:
    df_dato["Måned"] = df_dato["dato"].dt.to_period("M").astype(str)
    månedlig = df_dato.groupby("Måned")["beløp"].agg(
        Inntekter=lambda x: x[x > 0].sum(),
        Utgifter=lambda x: abs(x[x < 0].sum()),
    )
    månedlig["Resultat"] = månedlig["Inntekter"] - månedlig["Utgifter"]

    if len(månedlig) > 1:
        st.markdown("### 📅 Månedlig oversikt")
        # Bruk map() i stedet for applymap() (deprecated i pandas 2.x)
        styled = (
            månedlig.style
            .format("{:,.0f} kr")
            .map(
                lambda v: "color: green" if v >= 0 else "color: red",
                subset=["Resultat"],
            )
        )
        st.dataframe(styled, use_container_width=True)

# ── AI-innsikt ────────────────────────────────────────────────────────────────
st.markdown("### 🤖 AI-innsikt fra Claude")

fil_nøkkel = f"{fil.name}_{fil.size}"


def bygg_prompt() -> str:
    inntekt_oversikt = (
        inntekter.groupby("kategori")["beløp"]
        .sum().sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr").to_string()
    )
    utgift_oversikt = (
        utgifter.groupby("kategori")["beløp"]
        .sum().abs().sort_values(ascending=False)
        .apply(lambda x: f"{x:,.0f} kr").to_string()
    )
    periode_fra = df["dato"].min()
    periode_til = df["dato"].max()
    periode = ""
    if pd.notna(periode_fra) and pd.notna(periode_til):
        periode = f"{periode_fra.strftime('%d.%m.%Y')} - {periode_til.strftime('%d.%m.%Y')}"

    return f"""Du er en erfaren norsk regnskapsforer og forretningsradgiver.
Analyser folgende finansdata for en liten norsk bedrift og gi konkret, praktisk innsikt pa norsk.

PERIODE: {periode or 'Ukjent'}
TOTALE INNTEKTER: {total_inntekt:,.0f} kr
TOTALE UTGIFTER: {total_utgift:,.0f} kr
RESULTAT: {'+' if resultat >= 0 else ''}{resultat:,.0f} kr
FORTJENESTEMARGIN: {margin:.1f}%

INNTEKTER PER KATEGORI:
{inntekt_oversikt}

UTGIFTER PER KATEGORI:
{utgift_oversikt}

Gi en analyse pa 4-6 setninger som inkluderer:
1. Vurdering av lonnsomhet og finansiell helse
2. Den viktigste styrken og den viktigste risikoen
3. Ett konkret, handlingsrettet rad for a forbedre resultatet

Svar direkte og enkelt - eieren er ikke regnskapsfaglig. Unnga fagsjargong."""


def ai_stream_generator():
    claude = anthropic.Anthropic()
    with claude.messages.stream(
        model="claude-opus-4-6",
        max_tokens=512,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": bygg_prompt()}],
    ) as stream:
        for delta in stream.text_stream:
            yield delta


har_nøkkel_nå = bool(os.environ.get("ANTHROPIC_API_KEY"))
ny_fil = st.session_state.get("fil_nøkkel") != fil_nøkkel

if not har_nøkkel_nå:
    st.info("🔑 Legg inn Anthropic API-nøkkel i sidepanelet for å få AI-innsikt.")
    ai_tekst = "AI-innsikt ikke tilgjengelig — API-nøkkel mangler."
elif ny_fil:
    try:
        with st.container(border=True):
            ai_tekst = st.write_stream(ai_stream_generator())
        st.session_state.ai_tekst = ai_tekst
        st.session_state.fil_nøkkel = fil_nøkkel
    except anthropic.AuthenticationError:
        st.error("❌ Ugyldig API-nøkkel. Sjekk nøkkelen i sidepanelet.")
        ai_tekst = "AI-innsikt ikke tilgjengelig — ugyldig API-nøkkel."
    except Exception as e:
        st.warning(f"AI-analyse feilet: {e}")
        ai_tekst = "AI-innsikt ikke tilgjengelig."
else:
    with st.container(border=True):
        st.markdown(st.session_state.ai_tekst)
    ai_tekst = st.session_state.ai_tekst

# ── Last ned ──────────────────────────────────────────────────────────────────
st.markdown("### 📥 Last ned rapport")
dl1, dl2 = st.columns(2)

with dl1:
    txt_innhold = generer_rapport(df, ai_tekst)
    st.download_button(
        label="📄 Last ned TXT",
        data=txt_innhold.encode("utf-8"),
        file_name="finansrapport.txt",
        mime="text/plain",
        use_container_width=True,
    )

with dl2:
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_sti = tmp.name
        eksporter_pdf(df, ai_tekst, tmp_sti)
        with open(tmp_sti, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(tmp_sti)
        st.download_button(
            label="📑 Last ned PDF",
            data=pdf_bytes,
            file_name="finansrapport.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as e:
        st.warning(f"PDF-eksport feilet: {e}")

# ── Rådata ────────────────────────────────────────────────────────────────────
with st.expander("🔍 Vis opplastet data"):
    st.dataframe(df, use_container_width=True)

st.caption("KABI STARTUP AI — bygget av Karlo Bikic")
