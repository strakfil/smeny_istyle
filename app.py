import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io

# --- NASTAVENÍ STRÁNKY ---
st.set_page_config(page_title="Směny do kalendáře", page_icon="📅")

st.title("📅 Převodník směn (Excel/Numbers -> .ics)")
st.info("Nahrajte rozpis.")

# --- DATABÁZE ZKRATEK (v paměti prohlížeče) ---
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
    """Převede různé formáty času na objekt datetime.time."""
    if pd.isna(val) or val == "": return None
    if isinstance(val, time): return val
    if isinstance(val, datetime): return val.time()
    if isinstance(val, str):
        val = val.strip().replace('.', ':') # Oprava teček na dvojtečky
        if ":" not in val: return None
        try:
            return datetime.strptime(val, "%H:%M").time()
        except ValueError:
            try: return datetime.strptime(val, "%H:%M:%S").time()
            except: return None
    return None

# --- NAHRÁNÍ SOUBORU ---
uploaded_file = st.file_uploader("Nahrajte soubor (Excel nebo Numbers)", type=["xlsx", "numbers"])

if uploaded_file:
    df = pd.DataFrame()
    
    try:
        if uploaded_file.name.endswith('.numbers'):
            # Načtení Apple Numbers
            doc = Document(uploaded_file)
            data = doc.sheets()[0].tables()[0].rows(values_only=True)
            df = pd.DataFrame(data)
            # Nastavení prvního řádku jako záhlaví
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        else:
            # Načtení Excelu
            df = pd.read_excel(uploaded_file)
        
        st.success(f"Soubor '{uploaded_file.name}' byl úspěšně načten.")
    except Exception as e:
        st.error(f"Chyba při čtení souboru: {e}")

    if not df.empty:
        # Identifikace sloupců se jmény (přeskakujeme datum a prázdné sloupce)
        relevant_columns = []
        for i, col_name in enumerate(df.columns):
            name_str = str(col_name).strip()
            if i == 0 or "Unnamed" in name_str or name_str == "None" or name_str == "":
                continue
            relevant_columns.append((i, name_str))

        # Kontrola nových zaměstnanců
        with st.expander("👤 Správa zkratek zaměstnanců"):
            for _, full_name in relevant_columns:
                name_key = full_name.upper()
                if name_key not in st.session_state.employee_map:
                    new_abbr = st.text_input(f"Neznámý zaměstnanec: {full_name}. Zadejte zkratku:", key=name_key).strip().upper()
                    if new_abbr:
                        st.session_state.employee_map[name_key] = new_abbr
                else:
                    st.text(f"✅ {full_name} -> {st.session_state.employee_map[name_key]}")

        # Tlačítko pro generování ICS
        if st.button("🚀 Vygenerovat kalendář (.ics)"):
            ics_lines = [
                "BEGIN:VCALENDAR",
                "VERSION:2.0",
                "PRODID:-//Rozpis Smen Streamlit//CZ",
                "CALSCALE:GREGORIAN",
                "METHOD:PUBLISH"
            ]
            
            count_events = 0
            for index, row in df.iterrows():
                # První sloupec musí být datum
                date_val = pd.to_datetime(row.iloc[0], errors='coerce')
                if pd.isna(date_val): continue
                current_date = date_val.date()

                for col_idx, full_name in relevant_columns:
                    name_key = full_name.upper()
                    if name_key in st.session_state.employee_map:
                        abbr = st.session_state.employee_map[name_key]
                        
                        # Načtení časů (aktuální sloupec a následující)
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
                st.balloons()
                st.success(f"Vytvořeno {count_events} událostí!")
                
                st.download_button(
                    label="📥 Stáhnout hotový kalendář",
                    data=ics_string,
                    file_name=f"smeny_{datetime.now().strftime('%Y_%m')}.ics",
                    mime="text/calendar"
                )
            else:
                st.warning("V souboru nebyly nalezeny žádné směny (buňky s časem).")
