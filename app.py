import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

# --- KONFIGURACE ---
st.set_page_config(page_title="Převodník směn na .ICS", page_icon="📅")

st.title("📅 Převodník směn do kalendáře")
st.write("Nahrajte Excel s rozpisem a stáhněte si .ics soubor pro iPhone/Mac/Google.")

# Inicializace databáze zkratek v paměti aplikace
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
    if pd.isna(val): return None
    if isinstance(val, time): return val
    if isinstance(val, datetime): return val.time()
    if isinstance(val, str):
        val = val.strip()
        if ":" not in val: return None
        try:
            return datetime.strptime(val, "%H:%M").time()
        except ValueError: return None
    return None

# --- NAHRÁNÍ SOUBORU ---
uploaded_file = st.file_uploader("Vyberte soubor .xlsx", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Identifikace sloupců
    relevant_columns = []
    for i, col_name in enumerate(df.columns):
        name_str = str(col_name).strip()
        if i == 0 or "Unnamed" in name_str or name_str == "" or "datum" in name_str.lower():
            continue
        relevant_columns.append((i, name_str))

    # Kontrola nových zaměstnanců
    new_names = [n for _, n in relevant_columns if n.upper() not in st.session_state.employee_map]
    
    if new_names:
        st.warning("Byli nalezeni noví zaměstnanci. Zadejte prosím jejich zkratky:")
        for name in new_names:
            abbr = st.text_input(f"Zkratka pro {name}", key=name).strip().upper()
            if abbr:
                st.session_state.employee_map[name.upper()] = abbr

    # Tlačítko pro generování
    if st.button("Vygenerovat .ics soubor"):
        ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Rozpis Smen//CZ", "METHOD:PUBLISH"]
        count_events = 0

        for index, row in df.iterrows():
            date_val = pd.to_datetime(row.iloc[0], errors='coerce')
            if pd.isna(date_val): continue
            current_date = date_val.date()

            for col_idx, full_name in relevant_columns:
                name_key = full_name.upper()
                if name_key in st.session_state.employee_map:
                    abbr = st.session_state.employee_map[name_key]
                    
                    t_start = normalize_time(row.iloc[col_idx])
                    t_end = normalize_time(row.iloc[col_idx + 1]) if (col_idx + 1) < len(row) else None

                    if t_start and t_end:
                        dt_start = datetime.combine(current_date, t_start)
                        dt_end = datetime.combine(current_date, t_end)
                        fmt = "%Y%m%dT%H%M00"
                        
                        event = [
                            "BEGIN:VEVENT",
                            f"DTSTART:{dt_start.strftime(fmt)}",
                            f"DTEND:{dt_end.strftime(fmt)}",
                            f"SUMMARY:{abbr}",
                            "END:VEVENT"
                        ]
                        ics_lines.extend(event)
                        count_events += 1

        ics_lines.append("END:VCALENDAR")
        ics_string = "\n".join(ics_lines)
        
        st.success(f"Hotovo! Vygenerováno {count_events} směn.")
        
        # Nabídka ke stažení
        st.download_button(
            label="📥 Stáhnout .ics soubor",
            data=ics_string,
            file_name=f"smeny_{uploaded_file.name.replace('.xlsx', '')}.ics",
            mime="text/calendar"
        )