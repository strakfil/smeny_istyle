#STABILNÍ VERZE

import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io
import tempfile
import os

# --- KONFIGURACE ---
st.set_page_config(page_title="iStyle Kalendář", page_icon="📅")
st.title("📅 iStyle: Převodník směn")

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

uploaded_file = st.file_uploader("Nahrajte rozpis (.xlsx nebo .numbers)", type=["xlsx", "numbers"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.numbers'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".numbers") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            
            doc = Document(tmp_path)
            # Výběr listu (Sheet)
            sheet_names = [s.name for s in doc.sheets]
            selected_sheet_name = st.selectbox("Vyberte měsíc (list):", sheet_names)
            
            sheet = doc.sheets[selected_sheet_name]
            table = sheet.tables[0] # Bere první tabulku na listu
            data = table.rows(values_only=True)
            df_raw = pd.DataFrame(data)
            os.unlink(tmp_path)
        else:
            # Excel varianta
            xl = pd.ExcelFile(uploaded_file)
            selected_sheet_name = st.selectbox("Vyberte měsíc (list):", xl.sheet_names)
            df_raw = xl.parse(selected_sheet_name, header=None)

        # FIXNÍ NASTAVENÍ: Jména jsou na řádku 1 (index 1)
        row_names_index = 1 
        
        if len(df_raw) > row_names_index:
            df = df_raw.copy()
            # Nastavení hlavičky z řádku 1
            df.columns = [str(c).strip() if c is not None else f"Empty_{i}" for i, c in enumerate(df.iloc[row_names_index])]
            # Data začínají pod jmény
            df = df.iloc[row_names_index + 1:].reset_index(drop=True)

            relevant_columns = []
            for i, col_name in enumerate(df.columns):
                name_str = col_name
                if i == 0 or any(x in name_str.upper() for x in ["EMPTY_", "NAN", "NONE", "UNNAMED", "SMĚNY"]):
                    continue
                relevant_columns.append((i, name_str))

            with st.expander("👤 Kontrola zkratek"):
                for col_idx, full_name in relevant_columns:
                    name_key = full_name.upper()
                    if name_key not in st.session_state.employee_map:
                        abbr = st.text_input(f"Zkratka pro: {full_name}", key=f"k_{col_idx}").strip().upper()
                        if abbr: st.session_state.employee_map[name_key] = abbr
                    else:
                        st.text(f"✅ {full_name} -> {st.session_state.employee_map[name_key]}")

            if st.button("🚀 Vytvořit .ics pro tento měsíc"):
                ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//iStyle//CZ", "METHOD:PUBLISH"]
                count = 0
                for _, row in df.iterrows():
                    dt_val = pd.to_datetime(row.iloc[0], errors='coerce')
                    if pd.isna(dt_val): continue
                    
                    for col_idx, full_name in relevant_columns:
                        abbr = st.session_state.employee_map.get(full_name.upper())
                        if abbr:
                            # Čas začátku a konce
                            t_s = normalize_time(row.iloc[col_idx])
                            t_e = normalize_time(row.iloc[col_idx+1]) if (col_idx+1) < len(row) else None
                            
                            if t_s and t_e:
                                start = datetime.combine(dt_val.date(), t_s).strftime("%Y%m%dT%H%M00")
                                end = datetime.combine(dt_val.date(), t_e).strftime("%Y%m%dT%H%M00")
                                ics_lines.extend([
                                    "BEGIN:VEVENT",
                                    f"DTSTART:{start}",
                                    f"DTEND:{end}",
                                    f"SUMMARY:{abbr}",
                                    f"UID:{start}-{abbr}-{col_idx}@istyle",
                                    "END:VEVENT"
                                ])
                                count += 1
                
                ics_lines.append("END:VCALENDAR")
                if count > 0:
                    st.success(f"Zpracováno {count} směn z listu {selected_sheet_name}")
                    st.download_button("📥 Stáhnout .ics soubor", "\n".join(ics_lines), f"smeny_{selected_sheet_name}.ics", "text/calendar")
                else:
                    st.warning("Na tomto listu nebyly nalezeny žádné směny. Jsou jména opravdu na řádku 2 (index 1)?")
    
    except Exception as e:
        st.error(f"Chyba při zpracování: {e}")
