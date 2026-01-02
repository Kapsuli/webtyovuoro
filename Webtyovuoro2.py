import pandas as pd
from datetime import datetime, time
from io import BytesIO
from openpyxl.utils import get_column_letter
import streamlit as st
import unicodedata
import re

# ── Apu‐funktiot ─────────────────────────────────

def korjaa_sahkoposti_merkit(s: str) -> str:
    if s is None:
        return ""

    s = str(s).strip()

    # Poista aksentit (ä ö ü Ü ü jne.)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))

    # pieniksi ja siivous
    s = s.lower()
    s = re.sub(r"[,\-\s]+", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)

    return s

def korjaa_merkit(n):
    return str(n).replace(",", "") if pd.notna(n) else ""

def save_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='vuorot', index=False)
        ws = writer.sheets['vuorot']
        for i, col in enumerate(ws.columns, 1):
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[get_column_letter(i)].width = max_len + 2
    buffer.seek(0)
    return buffer

# ── CSV MUUNNOS ─────────────────────────────────

def muunna_csv(df, listan_alku):

    df.columns = [
        "Päivämäärä","Viikko","Viikonpäivä","Sukunimi, Etunimi","Työntekijänumero",
        "Palkkausmuoto","Ammattinimike","Työryhmä","Työpiste","Työvuoro",
        "Työvuoron alku","Työvuoron loppu","Työvuoron kesto","Työajanlaatu",
        "Merkintä","Tietoja","Ruokatauon alku","Ruokatauon kesto"
    ]

    data = []

    for _, row in df.iterrows():
        pvm = pd.to_datetime(row['Päivämäärä'], dayfirst=True, errors='coerce')
        nimi = row['Sukunimi, Etunimi']

        if pd.isna(pvm) or pd.isna(nimi) or pvm < listan_alku:
            continue

        # Nimi varmasti oikein
        sukunimi, etunimi = "", ""
        if "," in nimi:
            sukunimi, etunimi = [x.strip() for x in nimi.split(",", 1)]
        else:
            parts = str(nimi).split()
            if len(parts) >= 2:
                sukunimi, etunimi = parts[0], parts[1]

        r = {
            "jäsen": korjaa_merkit(nimi),
            "työsähköposti": f"{korjaa_sahkoposti_merkit(etunimi)}.{korjaa_sahkoposti_merkit(sukunimi)}@verisure.fi",
            "ryhmä": "ARC"
        }

        vuoro = str(row['Työvuoro'])
        alku = row['Työvuoron alku']
        loppu = row['Työvuoron loppu']
        selite_txt = row['Työajanlaatu'] if pd.notna(row['Työajanlaatu']) else ""

        def parse(t):
            try:
                return pd.to_datetime(t, format="%H:%M", errors="coerce").time()
            except:
                return None

        if vuoro in ("0:00-0:00", "00:00-00:00"):
            alk_a, paa_a = "08:00", "16:00"
            selite, vari = "", "1. Valkoinen"
        else:
            alk_a, paa_a = parse(alku), parse(loppu)
            selite, vari = "", "1. Valkoinen"

        pvm_end = pvm + pd.Timedelta(days=1) if isinstance(alk_a, time) and alk_a.hour >= 19 else pvm

        r.update({
            "Aloituspäivä": pvm,
            "Alkamisaika": alk_a,
            "Päättymispäivä": pvm_end,
            "Päättymisaika": paa_a,
            "Teeman väri": vari,
            "Mukautettu selite": selite,
            "Palkaton tauko": "",
            "Huomautuksia": "",
            "Jaettu": "2. Ei jaettu"
        })

        data.append(r)

    return pd.DataFrame(data)

# ── Streamlit UI ─────────────────────────────────

st.title("Teams Shifts -vuoromuunnin")

uploaded = st.file_uploader("1. Valitse työvuorolista (.csv)", type="csv")
alkupv_input = st.date_input("2. Anna listan alkupäivä")

if uploaded and alkupv_input:
    if st.button("Muunna ja lataa Excel"):
        try:
            df_orig = pd.read_csv(uploaded, sep=";", header=None, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df_orig = pd.read_csv(uploaded, sep=";", header=None, encoding="latin1")

        df_m = muunna_csv(df_orig, pd.to_datetime(alkupv_input))
        out = save_excel_bytes(df_m)

        st.download_button(
            "📥 Lataa Excel",
            data=out,
            file_name=f"teams_shifts_{alkupv_input}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
