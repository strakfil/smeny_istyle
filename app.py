import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io

# --- NASTAVENÍ STRÁNKY ---
st.set_page_config(page_title="Směny do kalendáře", page_icon="📅")

st.title("📅 Převodník směn")

# --- DATABÁZE ZKRATEK ---
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
    """Převede buňku na časový objekt, ignoruje texty."""
    if pd.isna(val) or val == "" or val is None: 
        return None
    if isinstance(val, time): 
        return val
    if isinstance(val, datetime): 
        return val.time()
    
    # Pokud je to string (např. z Numbers nebo Excelu)
    val_str = str(val).strip().replace('.', ':')
    if ":" not in val_str: 
        return None
    
    formats = ["%H:%M", "%H:%M:%S", "%G:%M"]
    for fmt in formats:
        try:
            return datetime.strptime(val_str, fmt).time()
        except ValueError:
            continue
    return None

# --- NAHRÁNÍ SOUBORU ---
uploaded_file = st.file_uploader("Vyberte soubor rozpisu (.xlsx nebo .numbers)", type=["xlsx", "numbers"])

if uploaded_file:
    df = pd.DataFrame()
    
    try:
        if uploaded_file.name.endswith('.numbers'):
            doc = Document(uploaded_file)
            # Načtení první tabulky z prvního listu
            table = doc.sheets()[0].tables()[0]
            data = table.rows(values_only=True)
            df = pd.DataFrame(data)
            # První řádek jsou jména
            df.columns = [str(c) if c is not None else f"Empty_{i}" for i, c in enumerate(df.iloc[0])]
            df = df[1:].reset_index(drop=True)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Chyba při čtení souboru: {e}")

    if not df.empty:
        # Hledání sloupců se jmény (přeskakujeme datum v indexu 0)
        relevant_columns = []
        for i, col_name in enumerate(df.columns):
            name_str = str(col_name).strip()
            # Ignorujeme prázdné sloupce a sloupec s datem
            if i == 0 or "Empty_" in name_str or name_str.lower() == "none" or name_str == "":
                continue
            # Jméno se bere jen pokud není "Unnamed" (typické pro prázdné sloupce v Excelu)
            if "Unnamed" not in name_str:
                relevant_columns.append((i, name_str))

        # Správa zkratek (zobrazí se jen pokud je potřeba)
        with st.expander("👤 Správa zkratek"):
            for _, full_name in relevant_columns:
                name_key = full_name.upper()
                if name_key not in st.session_state.employee_map:
                    new_abbr = st.text_input(f"Zadejte zkratku pro: {full_name}", key=name_key).strip().upper()
                    if new_abbr:
                        st.session_state.employee_map[name_key] = new_abbr
                else:
                    st.text(f"✅ {full_name} -> {st.session_state.employee_map[name_key]}")

        if st.button("🚀 Vygenerovat .ics kalendář"):
            ics_lines = [
                "BEGIN:VCALENDAR",
                "VERSION:2.0",
                "PRODID:-//Rozpis Smen//CZ",
                "METHOD:PUBLISH"
            ]
            
            count_events = 0
            for index, row in df.iterrows():
                # První sloupec je datum
                raw_date = row.iloc[0]
                date_val = pd.to_datetime(raw_date, errors='coerce')
                if pd.isna(date_val): 
                    continue
                current_date = date_val.date()

                for col_idx, full_name in relevant_columns:
                    name_key = full_name.upper()
                    if name_key in st.session_state.employee_map:
                        abbr = st.session_state.employee_map[name_key]
                        
                        # Časy jsou v aktuálním a následujícím sloupci
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
                                f"UID:{dt_start.strftime(fmt)}-{abbr}@smeny",
                                "END:VEVENT"
                            ]
                            ics_lines.extend(event)
                            count_events += 1

            ics_lines.append("END:VCALENDAR")
            ics_string = "\n".join(ics_lines)

            if count_events > 0:
                st.success(f"Úspěšně zpracováno {count_events} směn.")
                st.download_button(
                    label="📥 Stáhnout kalendář",
                    data=ics_string,
                    file_name=f"smeny_export.ics",
                    mime="text/calendar"
                )
            else:
                st.warning("V nahraném souboru nebyly nalezeny žádné časy směn.")
