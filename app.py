import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io
import tempfile
import os

st.set_page_config(page_title="Směny iStyle", page_icon="📅")
st.title("📅 Převodník směn iStyle")

# Výchozí mapa zkratek
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

uploaded_file = st.file_uploader("Nahrajte soubor .xlsx nebo .numbers", type=["xlsx", "numbers"])

if uploaded_file:
    df_raw = pd.DataFrame()
    try:
        if uploaded_file.name.endswith('.numbers'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".numbers") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            doc = Document(tmp_path)
            table = doc.sheets[0].tables[0]
            data = table.rows(values_only=True)
            df_raw = pd.DataFrame(data)
            os.unlink(tmp_path)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        st.write("### Náhled nahraného souboru")
        st.dataframe(df_raw.head(10)) # Ukážeme prvních 10 řádků pro orientaci

        # Uživatel si vybere, na kterém řádku jsou jména (v Numbers je to často řádek 0 nebo 1)
        row_index = st.number_input("Na kterém řádku jsou jména zaměstnanců? (0 = první řádek)", min_value=0, max_value=len(df_raw)-1, value=0)
        
        if st.button("Potvrdit výběr řádku"):
            # Nastavíme vybraný řádek jako hlavičku
            new_df = df_raw.copy()
            new_df.columns = [str(c).strip() if c is not None else f"Empty_{i}" for i, c in enumerate(new_df.iloc[row_index])]
            new_df = new_df.iloc[row_index + 1:].reset_index(drop=True)
            st.session_state.df = new_df
            st.success("Hlavička nastavena.")

    except Exception as e:
        st.error(f"Chyba při čtení: {e}")

    if 'df' in st.session_state:
        df = st.session_state.df
        relevant_columns = []
        
        # Automatická filtrace jmen z vybrané hlavičky
        for i, col_name in enumerate(df.columns):
            name_str = col_name
            if i == 0 or any(x in name_str.upper() for x in ["EMPTY_", "NAN", "NONE", "UNNAMED", "SMĚNY", "2026"]):
                continue
            relevant_columns.append((i, name_str))

        with st.expander("👤 Kontrola zkratek"):
            for col_idx, full_name in relevant_columns:
                name_key = full_name.upper()
                if name_key not in st.session_state.employee_map:
                    abbr = st.text_input(f"Zkratka pro: {full_name}", key=f"key_{col_idx}").strip().upper()
                    if abbr: st.session_state.employee_map[name_key] = abbr
                else:
                    st.text(f"✅ {full_name} -> {st.session_state.employee_map[name_key]}")

        if st.button("🚀 Generovat .ics soubor"):
            ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//iStyle//CZ", "METHOD:PUBLISH"]
            count = 0
            for idx, row in df.iterrows():
                dt_val = pd.to_datetime(row.iloc[0], errors='coerce')
                if pd.isna(dt_val): continue
                
                for col_idx, full_name in relevant_columns:
                    abbr = st.session_state.employee_map.get(full_name.upper())
                    if abbr:
                        t_s = normalize_time(row.iloc[col_idx])
                        t_e = normalize_time(row.iloc[col_idx+1]) if (col_idx+1) < len(row) else None
                        if t_s and t_e:
                            start = datetime.combine(dt_val.date(), t_s).strftime("%Y%m%dT%H%M00")
                            end = datetime.combine(dt_val.date(), t_e).strftime("%Y%m%dT%H%M00")
                            ics_lines.extend(["BEGIN:VEVENT", f"DTSTART:{start}", f"DTEND:{end}", f"SUMMARY:{abbr}", f"UID:{start}-{abbr}-{col_idx}@istyle", "END:VEVENT"])
                            count += 1
            
            ics_lines.append("END:VCALENDAR")
            if count > 0:
                st.balloons()
                st.download_button("📥 Stáhnout kalendář", "\n".join(ics_lines), "smeny.ics", "text/calendar")
            else:
                st.warning("Nebyly nalezeny žádné směny. Zkontrolujte, zda jste vybrali správný řádek se jmény.")
