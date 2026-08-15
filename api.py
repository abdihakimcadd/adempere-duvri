from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from supabase import create_client, Client
import os
import io
import json
from datetime import datetime

# DOCX Imports ONLY
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Import your pipeline functions
from knowledge import get_context_block
from agent_drafter import draft_duvri
from agent_reviewer import review_duvri

app = FastAPI()

# UPDATED CORS TO FIX THE NETWORK ERROR
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://adempere.com", 
        "https://www.adempere.com",
        "http://localhost:3000",
        "http://localhost:5173"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize Supabase Client
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def generate_duvri_docx(data):
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # === HEADER ===
    title = doc.add_heading('DOCUMENTO UNICO DI VALUTAZIONE DEI RISCHI DA INTERFERENZA (DUVRI)', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
    
    subtitle = doc.add_paragraph('Redatto ai sensi dell\'Art. 26, comma 3, D.Lgs. 9 aprile 2008 n. 81')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.italic = True
    subtitle.runs[0].font.size = Pt(10)
    
    doc.add_paragraph()  # Spacer
    
    # === PARTE 1: COMMITTENTE ===
    doc.add_heading('Parte 1 — Azienda Committente', level=1)
    table1 = doc.add_table(rows=4, cols=2)
    table1.style = 'Table Grid'
    cells1 = [
        ('Ragione Sociale:', data['host_company']),
        ('P.IVA / C.F.:', data['host_vat']),
        ('Sede dei Lavori:', data['work_site']),
        ('Ruolo Redattore:', 'RSPP / Datore di Lavoro Committente')
    ]
    for i, (label, value) in enumerate(cells1):
        table1.rows[i].cells[0].text = label
        table1.rows[i].cells[1].text = value
        table1.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_paragraph()  # Spacer
    
    # === PARTE 5: APPALTATRICE ===
    doc.add_heading('Parte 5 — Azienda Appaltatrice', level=1)
    table2 = doc.add_table(rows=4, cols=2)
    table2.style = 'Table Grid'
    cells2 = [
        ('Ragione Sociale:', data['contractor_name']),
        ('P.IVA / C.F.:', data['contractor_vat']),
        ('Oggetto Appalto:', data['work_type']),
        ('Periodo Lavori:', f"Da {data['start_date']} a {data['end_date']}")
    ]
    for i, (label, value) in enumerate(cells2):
        table2.rows[i].cells[0].text = label
        table2.rows[i].cells[1].text = value
        table2.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_page_break()
    
    # === PARTE 4.A: RISK MATRIX ===
    doc.add_heading('Parte 4.A — Matrice dei Rischi Interferenti', level=1)
    risk_table = doc.add_table(rows=1 + len(data['risks']), cols=4)
    risk_table.style = 'Table Grid'
    hdr_cells = risk_table.rows[0].cells
    headers = ['Rischio Identificato', 'Probabilità (P)', 'Magnitudo (M)', 'Indice (P × M)']
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
    
    for i, risk in enumerate(data['risks']):
        row = risk_table.rows[i+1].cells
        row[0].text = risk['name']
        row[1].text = str(risk['probability'])
        row[2].text = str(risk['magnitude'])
        row[3].text = f"{risk['score']} ({risk['level']})"
    
    doc.add_paragraph('Classificazione: 1-4 Basso, 5-9 Medio, 10-16 Alto.').runs[0].font.italic = True
    
    doc.add_page_break()
    
    # === PARTE 2: VALUTAZIONE RICOGNITIVA ===
    doc.add_heading('Parte 2 — Valutazione Ricognitiva (Art. 26, comma 3-ter)', level=1)
    doc.add_paragraph(data['valutazione_text'])
    
    doc.add_page_break()
    
    # === PARTE 4.B: MISURE ===
    doc.add_heading('Parte 4.B — Misure di Prevenzione e Protezione', level=1)
    doc.add_paragraph(data['misure_text'])
    
    doc.add_page_break()
    
    # === COSTI ===
    doc.add_heading('Individuazione dei Costi della Sicurezza (Art. 26, comma 5)', level=1)
    doc.add_paragraph(data['costi_text'])
    
    doc.add_page_break()
    
    # === PARTE 3: REGOLE ===
    doc.add_heading('Parte 3 — Regole di Prevenzione ed Emergenza in vigore nel Sito', level=1)
    rules = [
        'Divieto di fumo in tutti i locali aziendali.',
        'Obbligo di utilizzo dei DPI ove previsti dalla segnaletica di sicurezza.',
        'Mantenimento delle vie di fuga e dei presidi antincendio sgombri da ostacoli.',
        'Rispetto delle procedure di emergenza aziendali (vie di fuga, punti di raccolta).'
    ]
    for rule in rules:
        p = doc.add_paragraph(rule, style='List Bullet')
    
    doc.add_page_break()
    
    # === PARTE 4.C: COORDINAMENTO ===
    doc.add_heading('Parte 4.C — Procedure di Coordinamento', level=1)
    doc.add_paragraph(data['coordination_text'])
    
    # Signatures
    doc.add_paragraph()
    doc.add_paragraph()
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = 'Table Grid'
    sig_table.rows[0].cells[0].text = '___________________________'
    sig_table.rows[0].cells[1].text = '___________________________'
    sig_table.rows[1].cells[0].text = 'Il Committente'
    sig_table.rows[1].cells[1].text = "L'Appaltatore"
    sig_table.rows[1].cells[0].paragraphs[
