import os
from openai import OpenAI
from dotenv import load_dotenv
import json

load_dotenv()

# Use Llama 3.1 70b (128k context) if available, otherwise fallback to 3.0 70b
llm_client = OpenAI(
    api_key=os.getenv("LLM_API_KEY"),
    base_url=os.getenv("LLM_BASE_URL")
)
LLM_MODEL = os.getenv("LLM_MODEL")

def load_md_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return f"[ERROR: {filepath} not found.]"

DATA_DIR = "./data"
STORAGE_BUCKET = {
    "1_law/dvr_template.md": load_md_file(f"{DATA_DIR}/dvr_template.md"),
    "2_matrix/risk_matrix.md": load_md_file(f"{DATA_DIR}/risk_matrix.md"),
    "3_hazards/hazard_workplace.md": load_md_file(f"{DATA_DIR}/hazard_workplace.md"),
    "3_hazards/hazard_vdt.md": load_md_file(f"{DATA_DIR}/hazard_vdt.md"),
    "3_hazards/hazard_physical.md": load_md_file(f"{DATA_DIR}/hazard_physical.md"),
    "3_hazards/hazard_chemical.md": load_md_file(f"{DATA_DIR}/hazard_chemical.md"),
    "4_sectors/sector_supermercati.md": load_md_file(f"{DATA_DIR}/sector_supermercati.md")
}

# Kimi Fix #1: Aggressive Trimming
ESSENTIAL_LAW = """
Art. 17 - Obblighi del datore di lavoro: valutare rischi, adottare misure, nominare RSPP...
Art. 28 - Valutazione dei rischi (comma 1-2): Il documento deve contenere: 1) analisi rischi, 2) misure prevenzione, 3) programma miglioramento, 4) procedure attuazione e ruoli, 5) nominativi RSPP/RLS/medico, 6) mansioni a rischio specifico.
Art. 29 - Principi metodologici: valutazione per unità omogenee, consultazione RLS, aggiornamento...
Art. 55 - Formazione: obbligo informazione/formazione per tutti i lavoratori...
Title XII - Sanzioni: art. 55 e 55-bis, reato colposo, arresto/amminediativa...
"""

# Kimi Fix #1: Smart Hazard Routing (Only load 1-2 files per sector)
HAZARD_ROUTER = {
    "supermercati": ["hazard_vdt.md", "hazard_physical.md"],
    "ristorazione": ["hazard_physical.md", "hazard_chemical.md"],
    "edilizia": ["hazard_physical.md", "hazard_workplace.md"],
    "uffici": ["hazard_vdt.md", "hazard_workplace.md"],
    "industria": ["hazard_physical.md", "hazard_chemical.md"]
}