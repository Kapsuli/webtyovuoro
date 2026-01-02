import pandas as pd
from datetime import datetime, time
from io import BytesIO
from openpyxl.utils import get_column_letter
import streamlit as st

# ── Apu‐funktiot ─────────────────────────────────

def korjaa_sahkoposti_merkit(n):
    return n.replace("ä","a").replace("ö","o").replace(",","").replace("-","").replace(" ","").replace("ü","u")

def korjaa_merkit(n):
    return n.replace(",","")

def save_excel_bytes(df):
    """Palauta Excel-bytet objektina, sarakkeet autosize."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='vuorot', index=False)
        ws = writer.sheets['vuorot']
        for i, col in enumerate(ws.columns, 1):
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[get_column_letter(i)].width = max_len + 2
    buffer.seek(0)
    return buffer

def muunna_csv(df, listan_alku):
    # Aseta sarakenimet
    df.columns = [
        "Päivämäärä","Viikko","Viikonpäivä","Sukunimi, Etunimi","Työntekijänumero",
        "Palkkausmuoto","Ammattinimike","Työryhmä","Työpiste","Työvuoro",
        "Työvuoron alku","Työvuoron loppu","Työvuoron kesto","Työajanlaatu",
        "Merkintä","Tietoja","Ruokatauon alku","Ruokatauon kesto"
    ]

    data = []
    for _, row in df.iterrows():
        pvm = pd.to_datetime(row['Päivämäärä'], dayfirst=True, errors='coerce')
        if pd.isna(pvm) or pd.isna(row['Sukunimi, Etunimi']) or pvm < listan_alku:
            continue

        nimi = row['Sukunimi, Etunimi']
        r = {
            'jäsen': korjaa_merkit(nimi),
            'työsähköposti': "",
            'ryhmä': "ARC"
        }
        parts = nimi.split(" ")
        if len(parts)>=2:
            r['työsähköposti'] = f"{korjaa_sahkoposti_merkit(parts[1].lower())}.{korjaa_sahkoposti_merkit(parts[0].lower())}@verisure.fi"

        vuoro = str(row['Työvuoro'])
        alku = row['Työvuoron alku']; loppu = row['Työvuoron loppu']
        selite_txt = row['Työajanlaatu'] if pd.notna(row['Työajanlaatu']) else ""

        # Sama logiikka kuin aiemmin...
        if vuoro in ["0:00-0:00","00:00-00:00"]:
            alk_a, paa_a = "08:00","16:00"
            mapping = {
                "Toive Vapaa":("Vt","7. Harmaa"),
                "Vuosiloma":   ("VL","7. Harmaa"),
                "Vuosivapaa":  ("VV","7. Harmaa"),
                "Vapaa":       ("V","7. Harmaa"),
                "Muu palkallinen poissaolo":("poissa","7. Harmaa")
            }
            selite, vari = mapping.get(next((k for k in mapping if k in selite_txt), None), ("","1. Valkoinen"))
        else:
            # parse time
            def parse(t):
                try:
                    return pd.to_datetime(t, format='%H:%M', errors='coerce').time()
                except:
                    return t
            alk_a = parse(alku); paa_a = parse(loppu)

            mapping = {
                "Muu palkallinen poissaolo":("poissa","7. Harmaa"),
                "Vuosivapaa":               ("VV","7. Harmaa"),
                "Vuosiloma":                ("VL","7. Harmaa"),
                "Toive Vapaa":              ("Vt","7. Harmaa"),
                "Vapaa":                    ("V","7. Harmaa")
            }
            found = next(((s,v) for k,(s,v) in mapping.items() if k in selite_txt), (None,None))
            if found[0]:
                selite, vari = found
            else:
                if str(alk_a)[:5] in ("07:00","07:15") and str(paa_a)[:5] in ("19:00","19:15"):
                    selite, vari = "","5. Pinkki"
                elif isinstance(alk_a,time) and alk_a.hour==7:
                    selite, vari = "","6. Keltainen"
                elif isinstance(alk_a,time) and 7<alk_a.hour<19:
                    selite, vari = "","3. Vihreä"
                elif isinstance(alk_a,time) and alk_a.hour>=19:
                    selite, vari = "","2. Sininen"
                else:
                    selite, vari = "","1. Valkoinen"

        pvm_end = pvm + pd.Timedelta(days=1) if isinstance(alk_a,time) and alk_a.hour>=19 else pvm
        if selite in ("poissa","VV"):
            muk, huom = "","" + selite
        else:
            muk, huom = selite,""

        r.update({
            'Aloituspäivä':   pvm,
            'Alkamisaika':    alk_a,
            'Päättymispäivä': pvm_end,
            'Päättymisaika':  paa_a,
            'Teeman väri':    vari,
            'Mukautettu selite': muk,
            'Palkaton tauko': "",
            'Huomautuksia':    huom,
            'Jaettu':         "2. Ei jaettu"
        })
        data.append(r)

    return pd.DataFrame(data, columns=[
        'jäsen','työsähköposti','ryhmä','Aloituspäivä','Alkamisaika',
        'Päättymispäivä','Päättymisaika','Teeman väri',
        'Mukautettu selite','Palkaton tauko','Huomautuksia','Jaettu'
    ])

# ── Streamlit-sovellus ─────────────────────────────────

st.title("Teams Shifts -vuoromuunnin")
st.markdown("Lataa CSV, anna listan alkupäivä ja lataa valmis Excel-tiedosto.")

uploaded = st.file_uploader("1. Valitse tyovuorolista (.csv)", type="csv")
alkupv_input = st.date_input("2. Anna listan alkupäivä")
if uploaded and alkupv_input:
    if st.button("Muunna ja lataa Excel"):
        df_orig = pd.read_csv(uploaded, sep=';', header=None, encoding='latin1')
        df_m = muunna_csv(df_orig, pd.to_datetime(alkupv_input))
        out = save_excel_bytes(df_m)
        st.download_button(
            label="📥 Lataa Excel",
            data=out,
            file_name=f"teams_shifts_{alkupv_input.strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


