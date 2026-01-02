import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io
import tempfile
import os

# --- KONFIGURACE ---
st.set_page_config(page_title="iStyle Kalendář", page_icon="📅", layout="centered")

# CSS pro úpravu vzhledu (volitelné - zjemní okraje)
st.markdown("""
    <style>
    .stSecondaryButton { border-radius: 20px; }
    .stPrimaryButton { border-radius: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📅 iStyle Kalendář")

if 'employee_map' not in st.session_state:
    st.session_state.employee_map = {
        "MAREK STRAKA FT": "MST", "ONDŘEJ TVRDÍK FT": "OTV", "ARPÁD NORCINI FT": "ANO",
        "ELIŠKA DESÁKOVÁ FT": "EDE", "JAN BIŠKO FT": "JB2", "FILIP STRAKA FT": "FIS", "LUKÁŠ SUCHOMEL FT": "LSU"
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
        # --- NAČTENÍ SOUBORU ---
        if uploaded_file.name.endswith('.numbers'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".numbers") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            doc = Document(tmp_path)
            sheet_names = [s.name for s in doc.sheets]
            selected_sheet_name = st.selectbox("📅 Vyberte měsíc:", sheet_names)
            sheet = doc.sheets[selected_sheet_name]
            table = sheet.tables[0]
            df_raw = pd.DataFrame(table.rows(values_only=True))
            os.unlink(tmp_path)
        else:
            xl = pd.ExcelFile(uploaded_file)
            selected_sheet_name = st.selectbox("📅 Vyberte měsíc:", xl.sheet_names)
            df_raw = xl.parse(selected_sheet_name, header=None)

        # --- NASTAVENÍ HLAVIČKY (ŘÁDEK 1) ---
        row_names_index = 1 
        if len(df_raw) > row_names_index:
            df = df_raw.copy()
            df.columns = [str(c).strip() if c is not None else f"Empty_{i}" for i, c in enumerate(df.iloc[row_names_index])]
            df = df.iloc[row_names_index + 1:].reset_index(drop=True)

            all_relevant_columns = []
            for i, col_name in enumerate(df.columns):
                if i == 0 or any(x in col_name.upper() for x in ["EMPTY_", "NAN", "NONE", "UNNAMED", "SMĚNY"]): continue
                all_relevant_columns.append((i, col_name))

            # --- ESTETICKÝ PŘEPÍNAČ (Segmented Control) ---
            st.write("---")
            mode = st.segmented_control(
                "Režim exportu",
                options=["Standardní", "Individuální"],
                default="Standardní",
                selection_mode="single"
            )

            target_columns = []
            custom_name_map = {}

            if mode == "Individuální":
                col1, col2 = st.columns(2)
                with col1:
                    person_names = [name for _, name in all_relevant_columns]
                    selected_person = st.selectbox("Kdo jste?", person_names)
                with col2:
                    custom_summary = st.text_input("Název v kalendáři:", value="Práce iStyle")
                
                for col_idx, full_name in all_relevant_columns:
                    if full_name == selected_person:
                        target_columns.append((col_idx, full_name))
                        custom_name_map[full_name.upper()] = custom_summary
            else:
                target_columns = all_relevant_columns
                with st.expander("👤 Kontrola zkratek týmu"):
                    for col_idx, full_name in target_columns:
                        name_key = full_name.upper()
                        if name_key not in st.session_state.employee_map:
                            abbr = st.text_input(f"Zkratka pro: {full_name}", key=f"k_{col_idx}").strip().upper()
                            if abbr: st.session_state.employee_map[name_key] = abbr
                        else:
                            st.text(f"✅ {full_name} → {st.session_state.employee_map[name_key]}")

            # --- TLAČÍTKO GENEROVÁNÍ ---
            st.write("")
            if st.button("🚀 Vygenerovat kalendář", use_container_width=True, type="primary"):
                ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//iStyle//CZ", "METHOD:PUBLISH"]
                count = 0
                for _, row in df.iterrows():
                    dt_val = pd.to_datetime(row.iloc[0], errors='coerce')
                    if pd.isna(dt_val): continue
                    for col_idx, full_name in target_columns:
                        summary = custom_name_map.get(full_name.upper()) if mode == "Individuální" else st.session_state.employee_map.get(full_name.upper())
                        if summary:
                            t_s, t_e = normalize_time(row.iloc[col_idx]), normalize_time(row.iloc[col_idx+1]) if (col_idx+1) < len(row) else None
                            if t_s and t_e:
                                start, end = datetime.combine(dt_val.date(), t_s).strftime("%Y%m%dT%H%M00"), datetime.combine(dt_val.date(), t_e).strftime("%Y%m%dT%H%M00")
                                ics_lines.extend(["BEGIN:VEVENT", f"DTSTART:{start}", f"DTEND:{end}", f"SUMMARY:{summary}", f"UID:{start}-{full_name.replace(' ','')}-{col_idx}@istyle", "END:VEVENT"])
                                count += 1
                
                ics_lines.append("END:VCALENDAR")
                if count > 0:
                    st.success(f"Hotovo! Vytvořeno {count} směn.")
                    st.download_button("📥 Stáhnout .ics soubor", "\n".join(ics_lines), "smeny.ics", "text/calendar", use_container_width=True)
                else:
                    st.warning("Žádné směny k exportu.")
    
    except Exception as e:
        st.error(f"Chyba: {e}")
