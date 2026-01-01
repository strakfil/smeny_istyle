import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io
import tempfile
import os

# --- NASTAVENÍ ---
st.set_page_config(page_title="Směny do kalendáře", page_icon="📅")
st.title("📅 Převodník směn")

if 'employee_map' not in st.session_state:
    st.session_state.employee_map = {
        "MAREK STRAKA FT": "MST",
        "ONDŘEJ TVRDÍK FT": "OTV",
        "ARPÁD NORCINI FT": "ANO",
        "ELIŠKA DESÁKOVÁ FT": "EDE",
        "JAN BIŠKO FT": "JB2",
        "FILIP STRAKA FT": "FIS",
        "LUKÁŠ SUCHOMEL FT": "LSU"
    }

def normalize_time(val):
    if pd.isna(val) or val == "" or val is None: return None
    if isinstance(val, time): return val
    if isinstance(val, datetime): return val.time()
    val_str = str(val).strip().replace('.', ':')
    if ":" not in val_str: return None
    for fmt in ["%H:%M", "%H:%M:%S"]:
        try: return datetime.strptime(val_str, fmt).time()
        except ValueError: continue
    return None

uploaded_file = st.file_uploader("Vyberte soubor rozpisu", type=["xlsx", "numbers"])

if uploaded_file:
    df = pd.DataFrame()
    try:
        if uploaded_file.name.endswith('.numbers'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".numbers") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            doc = Document(tmp_path)
            table = doc.sheets[0].tables[0]
            data = table.rows(values_only=True)
            df = pd.DataFrame(data)
            os.unlink(tmp_path)
            if not df.empty:
                # Vyčištění názvů sloupců a odstranění prázdných hodnot
                df.columns = [str(c).strip() if c is not None else f"Empty_{i}" for i, c in enumerate(df.iloc[0])]
                df = df[1:].reset_index(drop=True)
        else:
            df = pd.read_excel(uploaded_file)
        st.success("Soubor načten.")
    except Exception as e:
        st.error(f"Chyba při čtení: {e}")

    if not df.empty:
        relevant_columns = []
        mesice = ["LEDEN", "ÚNOR", "BŘEZEN", "DUBEN", "KVĚTEN", "ČERVEN", "ČERVENEC", "SRPEN", "ZÁŘÍ", "ŘÍJEN", "LISTOPAD", "PROSINEC"]
        
        for i, col_name in enumerate(df.columns):
            name_str = str(col_name).strip()
            name_upper = name_str.upper()
            
            # FILTRACE: Přeskoč datum, prázdné sloupce, titulky a dlouhé texty
            if i == 0: continue
            if any(m in name_upper for m in mesice): continue
            if "SMĚNY" in name_upper or "TABULKA" in name_upper: continue
            if "EMPTY_" in name_upper or "UNNAMED" in name_upper or name_upper == "NAN" or name_upper == "NONE": continue
            if len(name_str) > 25: continue # Jméno s FT by nemělo být delší než 25 znaků
            
            relevant_columns.append((i, name_str))

        with st.expander("👤 Správa zkratek"):
            for col_idx, full_name in relevant_columns:
                name_key = full_name.upper()
                if name_key not in st.session_state.employee_map:
                    safe_key = f"input_{col_idx}"
                    new_abbr = st.text_input(f"Zkratka pro: {full_name}", key=safe_key).strip().upper()
                    if new_abbr:
                        st.session_state.employee_map[name_key] = new_abbr
                else:
                    st.text(f"✅ {full_name} -> {st.session_state.employee_map[name_key]}")

        if st.button("🚀 Vygenerovat .ics"):
            ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//iStyle Rozpis//CZ", "METHOD:PUBLISH"]
            count = 0
            for index, row in df.iterrows():
                dt_val = pd.to_datetime(row.iloc[0], errors='coerce')
                if pd.isna(dt_val): continue
                curr_date = dt_val.date()
                for col_idx, full_name in relevant_columns:
                    abbr = st.session_state.employee_map.get(full_name.upper())
                    if abbr:
                        t_s = normalize_time(row.iloc[col_idx])
                        t_e = normalize_time(row.iloc[col_idx+1]) if (col_idx+1) < len(row) else None
                        if t_s and t_e:
                            start = datetime.combine(curr_date, t_s).strftime("%Y%m%dT%H%M00")
                            end = datetime.combine(curr_date, t_e).strftime("%Y%m%dT%H%M00")
                            ics_lines.extend(["BEGIN:VEVENT", f"DTSTART:{start}", f"DTEND:{end}", f"SUMMARY:{abbr}", f"UID:{start}-{abbr}@istyle", "END:VEVENT"])
                            count += 1
            ics_lines.append("END:VCALENDAR")
            if count > 0:
                st.download_button("📥 Stáhnout kalendář", "\n".join(ics_lines), "smeny.ics", "text/calendar")
