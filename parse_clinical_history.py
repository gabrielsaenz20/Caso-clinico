#!/usr/bin/env python3
"""
parse_clinical_history.py — Parser para historia clínica en formato MSP/HOSVITAL (Ecuador)

Extrae automáticamente de la historia clínica (ej. "PLAZA QUIMIS GLADIS MARIA.pdf"):
  • Datos de filiación / demographics
  • Motivo de consulta
  • Enfermedad actual
  • Antecedentes patológicos personales (clínicos, quirúrgicos)
  • Alergias, antecedentes familiares, hábitos
  • Antecedentes ginecoobstétricos
  • Examen físico (signos vitales + examen por sistemas)
  • Análisis clínico (primera nota de ingreso)
  • Diagnósticos con código CIE-10
  • Notas de evolución (fecha, tipo, servicio, autor, subjetivo, plan, análisis)

Uso como script independiente:
    python3 parse_clinical_history.py "PLAZA QUIMIS GLADIS MARIA.pdf"
    python3 parse_clinical_history.py "HISTORIA_CLINICA.pdf" --output hc_data.yaml

Uso como módulo:
    from parse_clinical_history import parse_clinical_history
    data = parse_clinical_history("PLAZA QUIMIS GLADIS MARIA.pdf")

Dependencias:
    pip install pymupdf pyyaml
"""

import re
import sys
import argparse
from pathlib import Path
from typing import Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF no instalado. Ejecute: pip install pymupdf", file=sys.stderr)
    sys.exit(1)

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

# ─── Constantes de parseo ────────────────────────────────────────────────────

# Líneas del bloque de encabezado HOSVITAL que se repiten en cada página
# Terminan con "Atención Especial:" — todo lo antes es el header estandarizado
HEADER_END_MARKERS = [
    'Atención Especial:',
    'Atencion Especial:',
]

# Marcadores de sección dentro de las notas de hospitalización
SECTION_MARKERS = {
    'filiacion':          re.compile(r'DATOS\s+DE\s+FILIACI[ÓO]N', re.I),
    'referencia_familiar':re.compile(r'REFERENCIA\s+FAMILIAR\s*:', re.I),
    'antecedentes_pp':    re.compile(r'ANTECENTES?\s+PATOL[ÓO]GICOS\s+PERSONALES', re.I),
    'clinicos':           re.compile(r'CL[ÍI]NICOS\s*:', re.I),
    'quirurgicos':        re.compile(r'QUIR[ÚU]RGICOS\s*:', re.I),
    'alergias':           re.compile(r'ALERGIAS\s*:', re.I),
    'antecedentes_fam':   re.compile(r'ANTECEDENTES\s+FAMILIARES\s*:', re.I),
    'gineco':             re.compile(r'ANTECEDENTES\s+GINECO', re.I),
    'habitos':            re.compile(r'H[ÁA]BITOS\s*:', re.I),
    'motivo':             re.compile(r'MOTIVO\s+DE\s+CONSULTA\s*:', re.I),
    'enfermedad_actual':  re.compile(r'ENFERMEDAD\s+ACTUAL\s*:', re.I),
    'examen_fisico':      re.compile(r'EXAMEN\s+F[ÍI]SICO\s*:', re.I),
    'signos_vitales':     re.compile(r'SIGNOS\s+VITALES\s*:', re.I),
    'analisis':           re.compile(r'AN[ÁA]LISIS\s*:', re.I),
    'plan':               re.compile(r'\bPLAN\s*:', re.I),
    'diagnostico':        re.compile(r'DIAGN[ÓO]STICO', re.I),
    'subjetivo':          re.compile(r'SUBJETIVO\s*:', re.I),
    'objetivo':           re.compile(r'OBJETIVO\s*:', re.I),
    'evolucion':          re.compile(r'EVOLUCI[ÓO]N\s+M[EÉ]DICO|NOTA\s+DE\s+(?:INGRESO|EVOLUCI)', re.I),
}

# Regex para signos vitales
# Nota: los PDFs pueden usar caracteres acentuados (Ó, Í) en mayúsculas,
# se usan alternativas explícitas para mayor robustez.
VITAL_PATTERNS = {
    # Solo captura el número(s), sin requerir la unidad (que puede tener tildes)
    'pa':      re.compile(r'PRESI[ÓO]N\s+ARTERIAL\s*:?\s*([\d]+/[\d]+)', re.I),
    'fc':      re.compile(r'FRECUENCIA\s+CARD[IÍ]ACA\s*:?\s*([\d]+)', re.I),
    'fr':      re.compile(r'FRECUENCIA\s+RESPIRATORIA\s*:?\s*([\d]+)', re.I),
    'spo2':    re.compile(r'SATURACI[ÓO]N\s+DE\s+OX[IÍ]GENO\s*:?\s*([\d]+)', re.I),
    'temp':    re.compile(r'TEMPERATURA\s*:?\s*([\d]+[.,][\d]+)', re.I),
    'glasgow': re.compile(r'GLASGOW\s*([\d]+)\s*(?:SOBRE|/)\s*([\d]+)', re.I),
    'peso':    re.compile(r'PESO\s*:?\s*([\d]+[.,]?[\d]*)\s*(?:KG|KILOGRAMOS)', re.I),
    'talla':   re.compile(r'TALLA\s*:?\s*([\d]+[.,]?[\d]*)\s*(?:CM|M\b|METROS)', re.I),
}

# Regex para código CIE-10
CIE10_RE = re.compile(r'\b([A-Z][0-9]{1,2}\.?[0-9A-Z]{0,3})\b')

# Regex fecha (dd/mm/yyyy o dd/mm/yy)
DATE_RE = re.compile(r'(\d{2}/\d{2}/\d{4}|\d{2}/\d{2}/\d{2})')

# Líneas de pie de página que se deben ignorar
FOOTER_LINES = {
    'HISTORIA CLÍNICA No.',
    'Edad actual :',
    'Sexo:',
    'Grupo Sanguíneo:',
    'Empresa:',
    'Afiliado:',
    'Fecha Nacimiento:',
    'Dirección:',
    'Teléfono:',
    'Barrio:',
    'Municipio:',
    'Departamento:',
    'Estado Civil:',
    'Ocupacion:',
    'Etnia:',
    'Grupo Etnico:',
    'Nivel Educativo:',
    'Discapacidad:',
    'Grupo Poblacional:',
    'Atención Especial:',
    'HOSP EUGENIO ESPEJO',
    'HOSPITAL DE ESPECIALIDADES EUGENIO ESPEJO',
    'Pag:',
    'G.etareo:',
    'RHsClxFo',
    'Usuario:',
}

FOOTER_RE = re.compile(
    r'^(\*?\d{6,12}|Pag:|Fecha: \d{2}/\d{2}/\d{2}|\d+\nPag:'
    r'|HCU\s+\d+\s+--'
    r'|\d+\s+AÑOS|MSP-GENERAL|SIN DISCAPACIDAD|NO APLICA|Mestizo/a|Soltero|'
    r'7J\.0 \*HOSVITAL\*|de \d+|amas de casa)',
    re.I
)


# ─── Funciones de extracción ─────────────────────────────────────────────────

def _strip_hosvital_header(text: str) -> str:
    """
    Elimina el bloque de encabezado HOSVITAL que se repite en cada página.
    El encabezado termina con la línea "Atención Especial:".
    """
    for marker in HEADER_END_MARKERS:
        idx = text.find(marker)
        if idx >= 0:
            return text[idx + len(marker):]
    return text


def _clean_lines(text: str) -> list[str]:
    """Limpia el texto eliminando líneas vacías y de pie de página."""
    result = []
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        # Skip known footer / boilerplate lines
        if ln in FOOTER_LINES:
            continue
        if FOOTER_RE.match(ln):
            continue
        # Skip HOSVITAL stamp lines
        if re.match(r'^\*?\d{10}$', ln):
            continue
        if re.match(r'^\d{10}$', ln):
            continue
        if re.match(r'^\d{1,3}$', ln) and len(ln) <= 3:
            continue
        result.append(ln)
    return result


def _extract_section(text: str, start_pattern: re.Pattern,
                     end_patterns: list[re.Pattern],
                     max_chars: int = 3000) -> str:
    """
    Extrae texto desde start_pattern hasta el primer end_pattern encontrado
    (o hasta max_chars caracteres).
    """
    m = start_pattern.search(text)
    if not m:
        return ''
    start = m.end()
    end = len(text)
    for ep in end_patterns:
        em = ep.search(text, start)
        if em and em.start() < end:
            end = em.start()
    chunk = text[start:min(start + max_chars, end)]
    return chunk.strip()


def _parse_header_demographics(full_text_page1: str) -> dict:
    """
    Extrae datos demográficos del encabezado HOSVITAL de la primera página.
    El encabezado tiene un layout de columna único con valores antes y labels después.
    """
    demo = {}

    # HCU / medical record number
    m = re.search(r'\*(\d{8,12})\b', full_text_page1)
    if not m:
        m = re.search(r'HCU\s+(\d{8,12})', full_text_page1)
    if m:
        demo['hcu'] = m.group(1)

    # Patient name — in HOSVITAL header it appears right after "*HCU\nN\nPag:"
    # Format: "*0802060905\n1\nPag:\nFecha: 30/03/26\n11\nG.etareo:\nRHsClxFo\nde 736\nGLADIS MARIA..."
    m = re.search(r'de\s+\d+\n([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s,\.]+?)\n(Femenino|Masculino)',
                  full_text_page1, re.I)
    if m:
        demo['name'] = m.group(1).strip()
        demo['sex'] = m.group(2).strip()

    # Date of birth
    m = re.search(r'(\d{2}/\d{2}/\d{4})\nSoltero|(\d{2}/\d{2}/\d{4})\nCasado', full_text_page1)
    if not m:
        # Try alternative: find DOB in header block
        # It appears after province line: ESMERALDAS\n05/07/1976
        m = re.search(r'(?:ESMERALDAS|PICHINCHA|GUAYAS|AZUAY|MANABÍ|LOJA|[A-ZÁÉÍÓÚÑ]{4,})\n(\d{2}/\d{2}/\d{4})', full_text_page1)
    if m:
        demo['dob'] = m.group(1) if m.lastindex == 1 else (m.group(1) or m.group(2))

    # Phone
    m = re.search(r'\n(0\d{9})\n', full_text_page1)
    if m:
        demo['phone'] = m.group(1)

    # Age (in years) from header
    m = re.search(r'\n(\d+)\s+A[ÑN]OS\n', full_text_page1)
    if m:
        demo['age'] = m.group(1) + ' años'

    # Ethnicity
    m = re.search(r'\n(Mestizo/a|Afrodescendiente|Ind[íi]gena|Montubio/a|Blanco/a|Otro)\n', full_text_page1)
    if m:
        demo['ethnicity'] = m.group(1)

    # Hospital
    m = re.search(r'HOSPITAL DE ESPECIALIDADES ([^\n]+)', full_text_page1, re.I)
    if m:
        demo['hospital'] = m.group(0).strip()

    return demo


def _parse_vital_signs(text: str) -> dict:
    """Extrae signos vitales del texto (primera ocurrencia de cada parámetro)."""
    vs = {}
    for key, pat in VITAL_PATTERNS.items():
        m = pat.search(text)
        if m:
            if key == 'glasgow':
                vs[key] = f'{m.group(1)}/{m.group(2)}'
            else:
                vs[key] = m.group(1).replace(',', '.')
    return vs


def _parse_diagnoses(text: str) -> list[dict]:
    """Extrae diagnósticos con código CIE-10 del texto."""
    diagnoses = []
    # Pattern: "CELULITIS Y ABSCESO DE BOCA\nPRINCIPAL\nK122\nDIAGNÓSTICO"
    # or: "CELULITIS Y ABSCESO DE LA BOCA (CIE 10: K122)"
    # Pattern 1: inline CIE
    for m in re.finditer(r'([A-ZÁÉÍÓÚÑ][^(\n]+?)\s*\(CIE\s*10\s*:?\s*([A-Z]\d{1,2}\.?\d*)\)', text, re.I):
        diagnoses.append({'name': m.group(1).strip(), 'code': m.group(2)})

    # Pattern 2: multi-line block (HOSVITAL format)
    # name_line \n TYPE \n CODE \n DIAGNÓSTICO
    for m in re.finditer(
        r'^([A-ZÁÉÍÓÚÑ][^\n]{5,80})\n(PRINCIPAL|RELACIONADO|COMPLICACI[ÓO]N)\n([A-Z]\d{1,2}\.?\d*)\nDIAGN[ÓO]STICO',
        text, re.MULTILINE | re.I
    ):
        diagnoses.append({
            'name': m.group(1).strip(),
            'type': m.group(2).strip(),
            'code': m.group(3),
        })

    # Deduplicate by code
    seen = set()
    unique = []
    for d in diagnoses:
        if d['code'] not in seen:
            seen.add(d['code'])
            unique.append(d)
    return unique


def _parse_filiation(text: str) -> dict:
    """
    Extrae datos de filiación del párrafo narrativo de filiación.
    Ej: "PACIENTE FEMENINA DE 49 AÑOS, NACE Y RESIDENTE EN SANTO DOMINGO, 
          INSTRUCCIÓN: BÁSICA INCOMPLETA, OCUPACIÓN: NO TRABAJA, ESTADO CIVIL: UNIÓN LIBRE,
          RELIGIÓN: CATÓLICA, LATERALIDAD: DIESTRA, GRUPO SANGUÍNEO: NO CONOCE"
    """
    fil = {}
    patterns = {
        'origin':      r'(?:NACE Y )?RESIDENTE EN ([^,\.]+)',
        'education':   r'INSTRUCCI[ÓO]N\s*:?\s*([^,\.]+)',
        'occupation':  r'OCUPACI[ÓO]N\s*:?\s*([^,\.]+)',
        'civil_status':r'ESTADO CIVIL\s*:?\s*([^,\.]+)',
        'religion':    r'RELIGI[ÓO]N\s*:?\s*([^,\.]+)',
        'laterality':  r'LATERALIDAD\s*:?\s*([^,\.]+)',
        'blood_type':  r'GRUPO SANGU[ÍI]NEO\s*:?\s*([^,\.\n]+)',
        'transfusion': r'TRANSFUSIONES SANGU[ÍI]NEAS\s*:?\s*([^,\.\n]+)',
    }
    for key, pat in patterns.items():
        m = re.search(pat, text, re.I)
        if m:
            fil[key] = m.group(1).strip().title()
    return fil


def _parse_habits(text: str) -> dict:
    """Extrae hábitos de la sección de hábitos."""
    habits = {}
    fields = {
        'alcohol':  r'ALCOHOL\s*:?\s*([^\n]+)',
        'tobacco':  r'TABACO\s*:?\s*([^\n]+)',
        'drugs':    r'DROGAS?\s*:?\s*([^\n]+)',
        'biomass':  r'EXPOSICI[ÓO]N A BIOMASA\s*:?\s*([^\n]+)',
        'covid_vax':r'VACUNACI[ÓO]N (?:PARA )?COVID[^\n]*?(\d+\s*DOSIS)',
    }
    for key, pat in fields.items():
        m = re.search(pat, text, re.I)
        if m:
            habits[key] = m.group(1).strip().title()
    return habits


def _parse_gyneco(text: str) -> dict:
    """Extrae antecedentes ginecoobstétricos."""
    g = {}
    pairs = [
        ('gestas',   r'GESTAS?\s*:?\s*(\d+)'),
        ('partos',   r'PARTOS?\s*:?\s*(\d+)'),
        ('abortos',  r'ABORTOS?\s*:?\s*(\d+)'),
        ('cesareas', r'CES[AÁ]REAS?\s*:?\s*(\d+)'),
        ('hijos_vivos', r'HIJOS?\s+VIVOS?\s*:?\s*(\d+)'),
    ]
    for key, pat in pairs:
        m = re.search(pat, text, re.I)
        if m:
            g[key] = int(m.group(1))
    return g


def _parse_evolution_notes(full_content: str) -> list[dict]:
    """
    Extrae el resumen de notas de evolución (folio, fecha, tipo, servicio, autor).
    Busca el patrón FOLIO / FECHA / TIPO DE ATENCIÓN repetido en el documento.
    """
    notes = []
    # Pattern: "N\nDD/MM/YYYY HH:MM:SS\nTIPO\nSUBTIPO"
    pattern = re.compile(
        r'\bFOLIO\b[^\n]*\n[^\n]*\n[^\n]*\n[^\n]*\nEdad\s*:\s*\n'
        r'(\d+)\n'                            # folio number
        r'(\d{2}/\d{2}/\d{4}\s+[\d:]+)\n'    # date+time
        r'(URGENCIAS|HOSPITALIZACION|EMERGENCIA|[^\n]+)\n'  # type
        r'([^\n]+)',                           # subtype
        re.I
    )
    for m in pattern.finditer(full_content):
        notes.append({
            'folio': m.group(1),
            'date': m.group(2).strip(),
            'type': m.group(3).strip(),
            'subtype': m.group(4).strip(),
        })

    # Also extract author attribution lines
    author_re = re.compile(
        r'Evoluci[óo]n\s+realizada\s+por\s*:\s*([^-\n]+?)-?Fecha\s*:\s*(\d{2}/\d{2}/\d{2}\s+[\d:]+)',
        re.I
    )
    for m in author_re.finditer(full_content):
        notes.append({
            'author': m.group(1).strip(),
            'date': m.group(2).strip(),
            'type': 'NOTA',
            'subtype': '',
        })

    # Deduplicate by date+author
    seen = set()
    unique = []
    for n in notes:
        key = (n.get('date', ''), n.get('author', n.get('folio', '')))
        if key not in seen:
            seen.add(key)
            unique.append(n)
    return unique


def _find_admission_note(content: str) -> str:
    """
    Localiza el texto de la nota de ingreso (la más completa).
    Busca la nota con DATOS DE FILIACIÓN + ENFERMEDAD ACTUAL.
    """
    # Prefer the note that contains both DATOS DE FILIACIÓN and ENFERMEDAD ACTUAL
    best = ''
    idx = 0
    while True:
        m = SECTION_MARKERS['filiacion'].search(content, idx)
        if not m:
            break
        # Take up to 12000 chars from this position
        chunk = content[m.start():m.start() + 12000]
        if SECTION_MARKERS['enfermedad_actual'].search(chunk):
            if len(chunk) > len(best):
                best = chunk
        idx = m.end()
    return best


def _find_physical_exam_note(content: str) -> str:
    """
    Localiza la primera nota de ingreso que contenga examen físico completo
    con signos vitales, para extracción de datos.
    Acepta tanto 'EXAMEN FÍSICO:' como 'EXAMEN FÍSICO' sin colon.
    """
    # Find block with EXAMEN FÍSICO + SIGNOS VITALES + CABEZA/CUELLO
    pattern = re.compile(r'EXAMEN\s+F[ÍI]SICO', re.I)
    idx = 0
    while True:
        m = pattern.search(content, idx)
        if not m:
            break
        chunk = content[m.start():m.start() + 5000]
        if VITAL_PATTERNS['pa'].search(chunk) and re.search(r'CUELLO|CABEZA', chunk, re.I):
            return chunk
        idx = m.end()
    # Fallback: find PA anywhere and return surrounding block
    m_pa = VITAL_PATTERNS['pa'].search(content)
    if m_pa:
        return content[max(0, m_pa.start() - 300):m_pa.start() + 3000]
    return ''


# ─── Parser principal ─────────────────────────────────────────────────────────

def parse_clinical_history(pdf_path: str) -> dict:
    """
    Parsea la historia clínica PDF y devuelve un diccionario con todos los datos
    extraídos automáticamente.

    Args:
        pdf_path: ruta al PDF de la historia clínica

    Returns:
        dict con las claves:
            demographics, filiation, chief_complaint, present_illness,
            medical_history, surgical_history, allergies, family_history,
            habits, gyneco, vital_signs, physical_exam, clinical_analysis,
            diagnoses, evolution_notes, admission_date, discharge_date,
            raw_admission_note
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f'PDF no encontrado: {pdf_path}')

    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count

    # ── 1. Extraer datos demográficos del encabezado (primera página) ─────────
    first_page_text = doc[0].get_text()
    demographics = _parse_header_demographics(first_page_text)
    demographics['total_pages'] = total_pages

    # Buscar fecha de ingreso en primeras páginas (primera nota con fecha)
    for pg_idx in range(min(5, total_pages)):
        page_text = doc[pg_idx].get_text()
        m = DATE_RE.search(page_text, 100)  # skip first chars (HCU number)
        if m:
            # Look specifically for FECHA DE INGRESO A EMERGENCIA
            mi = re.search(r'FECHA Y HORA DE INGRESO[^\n]*?(\d{2}/\d{2}/\d{4})', page_text, re.I)
            if mi:
                demographics['admission_date'] = mi.group(1)
                break

    # ── 2. Construir texto unificado (sin headers repetidos) ──────────────────
    # Solo las primeras ~50 páginas suelen contener la nota de ingreso y datos relevantes
    # Las siguientes son órdenes, farmacia, etc. que son menos útiles para el reporte.
    # Procesamos por bloques para eficiencia con PDFs muy largos.
    MAX_PAGES_FULL = min(60, total_pages)
    MAX_PAGES_SUMMARY = min(total_pages, 200)   # para evolution notes

    full_content_parts = []
    for pg_idx in range(MAX_PAGES_FULL):
        page_text = doc[pg_idx].get_text()
        stripped = _strip_hosvital_header(page_text)
        full_content_parts.append(stripped)

    full_content = '\n'.join(full_content_parts)

    # Texto para notas de evolución (más páginas, menos detalle necesario)
    evol_parts = list(full_content_parts)
    for pg_idx in range(MAX_PAGES_FULL, MAX_PAGES_SUMMARY):
        page_text = doc[pg_idx].get_text()
        stripped = _strip_hosvital_header(page_text)
        evol_parts.append(stripped)
    evol_content = '\n'.join(evol_parts)

    doc.close()

    # ── 3. Nota de ingreso (la más completa con todos los datos) ──────────────
    admission_note = _find_admission_note(full_content)
    if not admission_note:
        # Fallback: take everything from first occurrence of MOTIVO DE CONSULTA
        m = SECTION_MARKERS['motivo'].search(full_content)
        if m:
            admission_note = full_content[m.start():m.start() + 8000]

    # ── 4. Datos de filiación ─────────────────────────────────────────────────
    filiation = {}
    sec_fil = _extract_section(
        admission_note,
        SECTION_MARKERS['filiacion'],
        [SECTION_MARKERS['antecedentes_pp'], SECTION_MARKERS['motivo']]
    )
    if sec_fil:
        filiation = _parse_filiation(sec_fil)

    # Referencia familiar
    m = re.search(r'REFERENCIA FAMILIAR\s*:\s*([^\n\r]+)', full_content, re.I)
    if m:
        filiation['family_contact'] = m.group(1).strip().title()

    # ── 5. Antecedentes patológicos ───────────────────────────────────────────
    sec_app = _extract_section(
        admission_note,
        SECTION_MARKERS['antecedentes_pp'],
        [SECTION_MARKERS['alergias'], SECTION_MARKERS['antecedentes_fam'],
         SECTION_MARKERS['habitos'], SECTION_MARKERS['motivo']],
        2000
    )

    medical_history_lines = []
    surgical_history_lines = []
    if sec_app:
        # Extract clinical history (between "CLÍNICOS:" and "QUIRÚRGICOS:" or "ALERGIAS:")
        m_cli = SECTION_MARKERS['clinicos'].search(sec_app)
        m_qx = SECTION_MARKERS['quirurgicos'].search(sec_app)
        if m_cli:
            cli_end = m_qx.start() if m_qx else len(sec_app)
            cli_text = sec_app[m_cli.end():cli_end].strip()
            medical_history_lines = [ln.strip() for ln in cli_text.splitlines() if ln.strip() and ln.strip() != '.']
        if m_qx:
            qx_end = len(sec_app)
            qx_text = sec_app[m_qx.end():qx_end].strip()
            surgical_history_lines = [ln.strip() for ln in qx_text.splitlines() if ln.strip() and ln.strip() != '.']

    # Alergias
    m = re.search(r'ALERGIAS\s*:\s*([^\n\r]{1,200})', admission_note, re.I)
    allergies = m.group(1).strip().title() if m else ''

    # Antecedentes familiares
    m = re.search(r'ANTECEDENTES\s+FAMILIARES\s*:\s*([^\n\r]{1,200})', admission_note, re.I)
    family_history = m.group(1).strip().title() if m else ''

    # ── 6. Hábitos ────────────────────────────────────────────────────────────
    sec_hab = _extract_section(
        admission_note,
        SECTION_MARKERS['habitos'],
        [SECTION_MARKERS['motivo'], SECTION_MARKERS['enfermedad_actual']],
        500
    )
    habits = _parse_habits(sec_hab) if sec_hab else {}

    # ── 7. Antecedentes ginecoobstétricos ─────────────────────────────────────
    sec_gineco = _extract_section(
        admission_note,
        SECTION_MARKERS['gineco'],
        [SECTION_MARKERS['habitos'], SECTION_MARKERS['motivo']],
        300
    )
    gyneco = _parse_gyneco(sec_gineco) if sec_gineco else {}

    # ── 8. Motivo de consulta ─────────────────────────────────────────────────
    m = re.search(r'MOTIVO\s+DE\s+CONSULTA\s*:\s*([^\n\r]{5,300})', admission_note, re.I)
    chief_complaint = m.group(1).strip().capitalize() if m else ''

    # ── 9. Enfermedad actual ──────────────────────────────────────────────────
    sec_ea = _extract_section(
        admission_note,
        SECTION_MARKERS['enfermedad_actual'],
        [SECTION_MARKERS['examen_fisico'], SECTION_MARKERS['analisis'],
         SECTION_MARKERS['plan']],
        4000
    )
    present_illness = re.sub(r'\n+', ' ', sec_ea).strip() if sec_ea else ''

    # ── 10. Examen físico y signos vitales ────────────────────────────────────
    physical_exam_block = _find_physical_exam_note(admission_note)
    if not physical_exam_block:
        physical_exam_block = _find_physical_exam_note(full_content)

    vital_signs = _parse_vital_signs(physical_exam_block) if physical_exam_block else {}

    # Systems exam: everything after SIGNOS VITALES until next section
    physical_exam_text = ''
    if physical_exam_block:
        vs_end = VITAL_PATTERNS['temp'].search(physical_exam_block)
        if vs_end:
            start = vs_end.end()
            # Find end of physical exam
            end_markers = ['ANÁLISIS', 'PLAN', 'DIAGNÓSTICO', 'EXÁMENES DE LABORATORIO']
            end = len(physical_exam_block)
            for em in end_markers:
                ei = physical_exam_block.find(em, start)
                if ei > 0 and ei < end:
                    end = ei
            physical_exam_text = re.sub(r'\n+', ' ', physical_exam_block[start:end]).strip()

    # ── 11. Análisis clínico ──────────────────────────────────────────────────
    sec_anal = _extract_section(
        full_content,
        SECTION_MARKERS['analisis'],
        [SECTION_MARKERS['plan'], SECTION_MARKERS['diagnostico']],
        2000
    )
    clinical_analysis = re.sub(r'\n+', ' ', sec_anal).strip() if sec_anal else ''

    # ── 12. Diagnósticos con CIE-10 ──────────────────────────────────────────
    diagnoses = _parse_diagnoses(full_content)

    # ── 13. Notas de evolución ────────────────────────────────────────────────
    evolution_notes = _find_admission_note_dates(evol_content)

    # ── 14. Resultado final ───────────────────────────────────────────────────
    return {
        'source_pdf': str(pdf_path),
        'demographics': demographics,
        'filiation': filiation,
        'chief_complaint': chief_complaint,
        'present_illness': present_illness,
        'medical_history': medical_history_lines,
        'surgical_history': surgical_history_lines,
        'allergies': allergies,
        'family_history': family_history,
        'habits': habits,
        'gyneco': gyneco,
        'vital_signs': vital_signs,
        'physical_exam': physical_exam_text,
        'clinical_analysis': clinical_analysis,
        'diagnoses': diagnoses,
        'evolution_notes': evolution_notes,
    }


def _find_admission_note_dates(content: str) -> list[dict]:
    """
    Extrae fechas y tipos de todas las notas de evolución para construir
    un resumen cronológico de la hospitalización.
    """
    notes = []
    seen_dates = set()

    # Pattern from HOSVITAL: folio number\ndate\ntype\nsubtype after FOLIO/FECHA/TIPO headers
    # The text after stripping headers looks like: "N\nDD/MM/YYYY HH:MM:SS\nTIPO\nSUBTIPO"
    folio_block_re = re.compile(
        r'(?:^|\n)(\d{1,4})\n'
        r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})\n'
        r'(URGENCIAS|HOSPITALIZACION|EMERGENCIA)\n'
        r'([^\n]{3,80})',
        re.MULTILINE
    )
    for m in folio_block_re.finditer(content):
        date_str = m.group(2)[:10]  # just DD/MM/YYYY
        if date_str not in seen_dates:
            seen_dates.add(date_str)
        notes.append({
            'folio': m.group(1),
            'date': m.group(2).strip(),
            'context': m.group(3).strip(),
            'note_type': m.group(4).strip(),
        })

    # Also extract signed evolution notes
    evol_re = re.compile(
        r'Evoluci[óo]n\s+realizada\s+por\s*:\s*([^\n\-]+?)(?:\s*-?\s*)?'
        r'Fecha\s*:\s*(\d{2}/\d{2}/\d{2,4}\s+[\d:]+)',
        re.I
    )
    for m in evol_re.finditer(content):
        notes.append({
            'author': m.group(1).strip(),
            'date': m.group(2).strip(),
            'context': 'NOTA DE EVOLUCIÓN',
            'note_type': '',
        })

    # Deduplicate keeping order
    seen = set()
    unique = []
    for n in notes:
        key = (n.get('folio', ''), n.get('date', ''))
        if key not in seen:
            seen.add(key)
            unique.append(n)

    return unique


def format_vital_signs(vs: dict) -> dict:
    """Formatea los signos vitales para mostrar en el documento."""
    labels = {
        'pa': ('PA (mmHg)', 'BP (mmHg)'),
        'fc': ('FC (lpm)', 'HR (bpm)'),
        'fr': ('FR (rpm)', 'RR (rpm)'),
        'spo2': ('SpO2 (%)', 'SpO2 (%)'),
        'temp': ('Temperatura (°C)', 'Temperature (°C)'),
        'glasgow': ('Glasgow', 'Glasgow'),
        'peso': ('Peso (kg)', 'Weight (kg)'),
        'talla': ('Talla (cm)', 'Height (cm)'),
    }
    formatted = {}
    for key, val in vs.items():
        if key in labels:
            formatted[key] = {
                'value': val,
                'label_es': labels[key][0],
                'label_en': labels[key][1],
            }
    return formatted


def print_summary(data: dict) -> None:
    """Imprime un resumen de los datos extraídos."""
    sep = '─' * 60
    print(sep)
    print('  HISTORIA CLÍNICA — RESUMEN DE EXTRACCIÓN')
    print(sep)
    demo = data.get('demographics', {})
    print(f"  Paciente:        {demo.get('name', '?')}")
    print(f"  HC:              {demo.get('hcu', '?')}")
    print(f"  Sexo:            {demo.get('sex', '?')}")
    print(f"  FN:              {demo.get('dob', '?')}")
    print(f"  Edad:            {demo.get('age', '?')}")
    print(f"  F. ingreso:      {demo.get('admission_date', '?')}")
    print(f"  Total páginas:   {demo.get('total_pages', '?')}")
    print()
    fil = data.get('filiation', {})
    print(f"  Procedencia:     {fil.get('origin', '?')}")
    print(f"  Estado civil:    {fil.get('civil_status', '?')}")
    print(f"  Ocupación:       {fil.get('occupation', '?')}")
    print(f"  Escolaridad:     {fil.get('education', '?')}")
    print(f"  Grupo sanguíneo: {fil.get('blood_type', '?')}")
    print(f"  Contacto fam.:   {fil.get('family_contact', '?')}")
    print()
    print(f"  Motivo consulta: {data.get('chief_complaint', '?')[:100]}")
    print()
    mh = data.get('medical_history', [])
    print(f"  Antec. clínicos ({len(mh)}):")
    for ln in mh[:5]:
        print(f"    • {ln[:90]}")
    print()
    vs = data.get('vital_signs', {})
    print(f"  Signos vitales ({len(vs)}):")
    for k, v in vs.items():
        print(f"    {k:10s}: {v}")
    print()
    dx = data.get('diagnoses', [])
    print(f"  Diagnósticos ({len(dx)}):")
    for d in dx[:8]:
        print(f"    [{d.get('code', '?')}] {d.get('name', '?')[:70]}")
    print()
    evol = data.get('evolution_notes', [])
    print(f"  Notas de evolución extraídas: {len(evol)}")
    print(sep)


# ─── CLI ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Extrae datos automáticamente de la historia clínica (formato MSP/HOSVITAL Ecuador)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('pdf', help='Ruta al PDF de la historia clínica')
    parser.add_argument('--output', '-o', help='Guardar resultado en archivo YAML')
    parser.add_argument('--summary', action='store_true', default=True,
                        help='Mostrar resumen en consola (por defecto: sí)')
    parser.add_argument('--no-summary', dest='summary', action='store_false')
    args = parser.parse_args()

    data = parse_clinical_history(args.pdf)

    if args.summary:
        print_summary(data)

    if args.output:
        if not HAS_YAML:
            print('ERROR: pyyaml no instalado. Ejecute: pip install pyyaml', file=sys.stderr)
            sys.exit(1)
        out_path = Path(args.output)
        with open(out_path, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False,
                      sort_keys=False, indent=2)
        print(f'\n✓ Datos guardados en: {out_path}')
    elif not args.summary:
        if HAS_YAML:
            print(yaml.dump(data, allow_unicode=True, default_flow_style=False,
                            sort_keys=False, indent=2))


if __name__ == '__main__':
    main()
