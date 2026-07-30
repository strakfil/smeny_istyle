import streamlit as st
import pandas as pd
from datetime import datetime, time
from numbers_parser import Document
import io
import tempfile
import os
import re
import json
import unicodedata
import difflib

# Cesta k souboru se seznamem zaměstnanců (leží vedle app.py v repu na GitHubu)
EMPLOYEES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "employees.json")

# --- KONFIGURACE ---
st.set_page_config(page_title="iStyle Kalendář", page_icon="📅", layout="centered")

# CSS pro úpravu vzhledu
st.markdown("""
    <style>
    .stSecondaryButton { border-radius: 20px; }
    .stPrimaryButton { border-radius: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("📅 iStyle Kalendář")

# --- FUNKCE PRO SJEDNOCENÍ TVARU JMEN ---
def normalize_name(name):
    name = str(name).upper()
    name = name.replace('\n', ' ').replace('\r', ' ') # Odstraní odřádkování (Alt+Enter v buňce)
    name = name.replace('-', ' ') # Nahradí pomlčky mezerou
    name = re.sub(r'\s+', ' ', name).strip() # Odstraní vícenásobné mezery a okraje
    return name

def load_employee_map():
    """Načte seznam zaměstnanců a jejich zkratek z externího souboru employees.json.
    Díky tomu se při měsíční změně týmu upravuje pouze tento soubor,
    aniž by bylo nutné zasahovat do kódu aplikace.
    Pokud soubor chybí nebo je poškozený, aplikace nespadne - jen upozorní
    a chybějící zkratky lze doplnit ručně přímo v UI (viz "Kontrola zkratek týmu")."""
    mapping = {}
    try:
        with open(EMPLOYEES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for entry in data:
            abbr = entry.get("abbr")
            for name in entry.get("names", []):
                mapping[normalize_name(name)] = abbr
    except FileNotFoundError:
        st.warning(f"⚠️ Soubor `employees.json` nebyl nalezen vedle app.py. Zkratky bude nutné doplnit ručně v sekci 'Kontrola zkratek týmu'.")
    except json.JSONDecodeError as e:
        st.error(f"⚠️ Soubor employees.json obsahuje chybu ve formátu JSON: {e}")
    return mapping

if 'employee_map' not in st.session_state:
    st.session_state.employee_map = load_employee_map()

def strip_diacritics(text):
    """Převede text bez diakritiky (á->a, š->s...) - používá se jen pro POROVNÁVÁNÍ,
    ne pro zobrazení či export (tam se pořád používá plné jméno se zkratkou z employees.json)."""
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if not unicodedata.combining(c))

def match_employee(full_name, employee_map):
    """Najde zkratku pro jméno ze sloupce tabulky. numbers-parser občas u buněk
    s více textovými styly vrátí jen část jména (např. 'ARPÁD NORCINI FT' -> 'D NORCINI FT'),
    takže se nespoléhá jen na přesnou shodu, ale postupně zkouší:
      1) přesná shoda (po normalizaci velikosti/mezer/pomlček)
      2) shoda bez diakritiky (kdyby o diakritiku přišlo úplně)
      3) shoda jako podřetězec (kdyby numbers-parser část jména 'ukously')
      4) přibližná (fuzzy) shoda podle podobnosti textu
    Vrací (zkratka, typ_shody) - typ_shody je 'exact', 'fuzzy' nebo None (nenalezeno)."""
    key = normalize_name(full_name)
    if not key:
        return None, None

    # 1) přesná shoda
    if key in employee_map:
        return employee_map[key], "exact"

    # 2) shoda bez diakritiky
    key_stripped = strip_diacritics(key)
    for k, abbr in employee_map.items():
        if strip_diacritics(k) == key_stripped:
            return abbr, "exact"

    # 3) podřetězec (jméno v tabulce je jen "ukousnutá" část správného jména, nebo naopak)
    for k, abbr in employee_map.items():
        if key in k or k in key:
            return abbr, "fuzzy"

    # 4) přibližná shoda podle podobnosti
    close = difflib.get_close_matches(key, list(employee_map.keys()), n=1, cutoff=0.6)
    if close:
        return employee_map[close[0]], "fuzzy"

    return None, None

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
            
            # OPRAVA 1: Automaticky vybere největší tabulku na listu (ignoruje legendy)
            largest_table = sheet.tables[0]
            for t in sheet.tables:
                if len(t.rows(values_only=True)) > len(largest_table.rows(values_only=True)):
                    largest_table = t
            
            df_raw = pd.DataFrame(largest_table.rows(values_only=True))
            os.unlink(tmp_path)
        else:
            xl = pd.ExcelFile(uploaded_file)
            selected_sheet_name = st.selectbox("📅 Vyberte měsíc:", xl.sheet_names)
            df_raw = xl.parse(selected_sheet_name, header=None)

        # --- DYNAMICKÉ NASTAVENÍ HLAVIČKY (ŘÁDEK SE JMÉNY) ---
        row_names_index = 1 
        for idx, row in df_raw.head(10).iterrows():
            # Kód vyhledá řádek, který obsahuje FT nebo PT (bezpečnější než fixní řádek 2)
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)]).upper()
            if "FT" in row_str or "PT" in row_str:
                row_names_index = idx
                break

        if len(df_raw) > row_names_index:
            df = df_raw.copy()
            df.columns = [str(c).strip() if pd.notna(c) else f"Empty_{i}" for i, c in enumerate(df.iloc[row_names_index])]
            df = df.iloc[row_names_index + 1:].reset_index(drop=True)
            
            # OPRAVA 2: Odstranění prázdných řádků podle sloupce s datumy (řeší 30/31 dní)
            df = df.dropna(subset=[df.columns[0]])

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
                    if st.button("🔄 Načíst seznam zaměstnanců znovu ze souboru"):
                        st.session_state.employee_map = load_employee_map()
                        st.rerun()
                    for col_idx, full_name in target_columns:

                        abbr, match_type = match_employee(full_name, st.session_state.employee_map)

                        if abbr is None:
                            typed = st.text_input(f"Zkratka pro: {full_name}", key=f"k_{col_idx}").strip().upper()
                            if typed:
                                st.session_state.employee_map[normalize_name(full_name)] = typed
                        elif match_type == "exact":
                            st.text(f"✅ {full_name} → {abbr}")
                        else:
                            st.warning(f"⚠️ {full_name} → {abbr}  (přibližná shoda - numbers-parser jméno zřejmě zkomolil, zkontrolujte)")

            # --- TLAČÍTKO GENEROVÁNÍ ---
            st.write("")
            if st.button("🚀 Vygenerovat kalendář", use_container_width=True, type="primary"):
                ics_lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//iStyle//CZ", "METHOD:PUBLISH"]
                count = 0
                for _, row in df.iterrows():
                    dt_val = pd.to_datetime(row.iloc[0], errors='coerce')
                    if pd.isna(dt_val): continue
                    for col_idx, full_name in target_columns:
                        
                        # POUŽITÍ CHYTRÉHO DOHLEDÁNÍ (match_employee) PŘI EXPORTU
                        if mode == "Individuální":
                            summary = custom_name_map.get(full_name.upper())
                        else:
                            summary, _ = match_employee(full_name, st.session_state.employee_map)
                        
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
