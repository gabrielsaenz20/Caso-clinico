#!/usr/bin/env python3
"""
build_case_report.py — Constructor genérico de reportes de caso clínico bilingüe (ES/EN)

Uso:
    python3 build_case_report.py --config case_config_template.yaml [--figs-only] [--skip-frames]

Dependencias:
    pip install python-docx pyyaml opencv-python-headless pymupdf

Secciones automáticas (extraídas/calculadas):
    - Portada (autores, afiliaciones, ORCIDs)
    - Tablas de datos (demographics, labs, cultivos, antibióticos, cirugías)
    - Figuras (fotogramas de video + imágenes clínicas)
    - Referencias (Vancouver numerado)
    - Declaración CRediT de autores

Secciones que requieren IA o revisión manual (marcadas en naranja en el documento):
    - Resumen / Abstract (párrafos de presentación del caso, conclusiones)
    - Introducción (texto de subsecciones)
    - Discusión (análisis comparativo)
    - Conclusiones con texto de redacción libre
"""

import argparse
import os
import sys
import re
import shutil
import zipfile
from pathlib import Path

import yaml
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import cv2
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False
    print("AVISO: opencv no disponible. La extracción de fotogramas de video no estará disponible.")

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# Importar parser de historia clínica si está disponible
try:
    from parse_clinical_history import parse_clinical_history as _parse_hc
    HAS_HC_PARSER = True
except ImportError:
    HAS_HC_PARSER = False

# ─── Colores ──────────────────────────────────────────────────────────────────
DARK_BLUE = RGBColor(0x1F, 0x38, 0x64)
MED_BLUE  = RGBColor(0x2E, 0x75, 0xB6)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
RED       = RGBColor(0xCC, 0x00, 0x00)
GRAY      = RGBColor(0x77, 0x77, 0x77)
ORANGE    = RGBColor(0xFF, 0x66, 0x00)   # marcador AI_PENDIENTE

AI_MARKER_PREFIX = '[AI_PENDIENTE'
AI_MARKER_PREFIX_EN = '[AI_PENDING'


def is_ai_placeholder(text):
    """Devuelve True si el texto es un marcador de contenido pendiente de IA."""
    if not isinstance(text, str):
        return False
    t = text.strip()
    return t.startswith(AI_MARKER_PREFIX) or t.startswith(AI_MARKER_PREFIX_EN)


def parse_timestamp(ts):
    """Convierte 'HH:MM:SS', 'MM:SS' o número a segundos (float)."""
    if isinstance(ts, (int, float)):
        return float(ts)
    ts = str(ts).strip()
    parts = ts.split(':')
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
    elif len(parts) == 2:
        return int(parts[0]) * 60 + float(parts[1])
    return float(ts)


# ─── XML helpers ─────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_margins(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top', 80), ('bottom', 80), ('left', 120), ('right', 120)]:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)


def set_cell_borders(cell, color='CCCCCC', size=4):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_spacing(para, before=4, after=4, line=None):
    pPr = para._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before * 20))
    sp.set(qn('w:after'), str(after * 20))
    if line:
        sp.set(qn('w:line'), str(line))
        sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)


# ─── Clinical history enrichment ─────────────────────────────────────────────

def _enrich_cfg_from_clinical_history(cfg: dict, hc_data: dict) -> None:
    """
    Enriquece el diccionario de configuración con datos extraídos
    automáticamente de la historia clínica.

    Solo rellena campos que estén vacíos o ausentes en el config — los
    valores definidos manualmente en el YAML siempre tienen prioridad.
    """
    demo = hc_data.get('demographics', {})
    fil  = hc_data.get('filiation', {})
    vs   = hc_data.get('vital_signs', {})
    dx   = hc_data.get('diagnoses', [])

    # ── Datos del paciente ────────────────────────────────────────────────────
    cp = cfg.setdefault('case_presentation', {})
    consent_es = (
        "Presentado con consentimiento informado de la paciente, en cumplimiento "
        "de la Ley Orgánica de Salud del Ecuador. La identidad se protege mediante iniciales."
    )
    consent_en = (
        "Reported with informed patient consent per Ecuador's Organic Health Law. "
        "Identity protected with initials."
    )
    cp.setdefault('consent_statement_es', consent_es)
    cp.setdefault('consent_statement_en', consent_en)

    # ── Demographics table (auto-generada desde HC) ───────────────────────────
    if 'demographics_table' not in cp:
        # Construir iniciales del paciente
        name = demo.get('name', '')
        initials = '.'.join(w[0] for w in name.split() if w) + '.' if name else '?'

        sex_es = demo.get('sex', '')
        sex_en = 'Female' if 'fem' in sex_es.lower() else ('Male' if 'masc' in sex_es.lower() else sex_es)
        age    = demo.get('age', '')
        dob    = demo.get('dob', '')
        hcu    = demo.get('hcu', '')
        adm_date = demo.get('admission_date', '')
        origin = fil.get('origin', '')
        hosp   = demo.get('hospital', 'Hospital de Especialidades Eugenio Espejo')
        contact = fil.get('family_contact', '')

        rows_es = [
            ['Paciente (iniciales)',    initials],
            ['Sexo / Edad',            f'{sex_es}, {age} (FN: {dob})'],
            ['Procedencia',            origin],
            ['Fecha de ingreso HEEE',  adm_date],
            ['Historia clínica HEEE',  hcu],
            ['Hospital',               hosp],
        ]
        rows_en = [
            ['Patient (initials)',       initials],
            ['Sex / Age',               f'{sex_en}, {age} (DOB: {dob})'],
            ['Origin',                  origin],
            ['Date of admission (HEEE)', adm_date],
            ['HEEE medical record',      hcu],
            ['Hospital',                hosp],
        ]
        if contact:
            rows_es.append(['Contacto familiar', contact])
            rows_en.append(['Family contact',    contact])

        cp['demographics_table'] = {
            'title_es': 'Datos generales de la paciente',
            'title_en': 'Patient demographic data',
            'headers_es': ['Parámetro', 'Dato'],
            'headers_en': ['Parameter', 'Data'],
            'widths': [6, 10],
            'rows_es': rows_es,
            'rows_en': rows_en,
        }

    # ── Chief complaint & present illness ─────────────────────────────────────
    cc = cp.setdefault('chief_complaint', {})

    # Motivo de consulta
    motivo = hc_data.get('chief_complaint', '')
    if motivo and 'narrative_es' not in cc:
        cc['narrative_es'] = motivo.capitalize()
        cc['narrative_en'] = '[AI_PENDING: Translate chief complaint to English]'

    # Enfermedad actual
    ea = hc_data.get('present_illness', '')
    if ea and 'present_illness_es' not in cc:
        cc['present_illness_es'] = ea.capitalize()
        cc['present_illness_en'] = '[AI_PENDING: Translate present illness to English]'

    # Antecedentes patológicos
    med_hx = hc_data.get('medical_history', [])
    surg_hx = hc_data.get('surgical_history', [])
    allergies = hc_data.get('allergies', '')
    family_hx = hc_data.get('family_history', '')
    habits = hc_data.get('habits', {})
    gyneco = hc_data.get('gyneco', {})

    if med_hx and 'medical_history_es' not in cc:
        cc['medical_history_es'] = '; '.join(ln.capitalize() for ln in med_hx)
    if surg_hx and 'surgical_history_es' not in cc:
        cc['surgical_history_es'] = '; '.join(ln.capitalize() for ln in surg_hx)
    if allergies and 'allergies_es' not in cc:
        cc['allergies_es'] = allergies
    if family_hx and 'family_history_es' not in cc:
        cc['family_history_es'] = family_hx
    if habits and 'habits_es' not in cc:
        parts = []
        labels = {
            'alcohol': 'Alcohol', 'tobacco': 'Tabaco', 'drugs': 'Drogas',
            'biomass': 'Biomasa', 'covid_vax': 'Vacuna COVID'
        }
        for k, v in habits.items():
            parts.append(f'{labels.get(k, k)}: {v}')
        cc['habits_es'] = '; '.join(parts)
    if gyneco and 'gyneco_es' not in cc:
        g = gyneco
        cc['gyneco_es'] = (
            f"G{g.get('gestas','?')} P{g.get('partos','?')} "
            f"A{g.get('abortos','?')} C{g.get('cesareas','?')} "
            f"HV{g.get('hijos_vivos','?')}"
        )

    # ── Physical exam / Vital signs ───────────────────────────────────────────
    pe = cp.setdefault('physical_exam', {})
    if vs and 'vital_signs_table' not in pe:
        vs_map = {
            'pa':      ('PA (mmHg)',         'BP (mmHg)'),
            'fc':      ('FC (lpm)',           'HR (bpm)'),
            'fr':      ('FR (rpm)',           'RR (rpm)'),
            'spo2':    ('SpO₂ (%)',           'SpO₂ (%)'),
            'temp':    ('Temperatura (°C)',   'Temperature (°C)'),
            'glasgow': ('Glasgow',            'Glasgow'),
            'peso':    ('Peso (kg)',          'Weight (kg)'),
            'talla':   ('Talla (cm)',         'Height (cm)'),
        }
        rows_es = []
        rows_en = []
        for k, val in vs.items():
            if k in vs_map:
                rows_es.append([vs_map[k][0], str(val)])
                rows_en.append([vs_map[k][1], str(val)])
        if rows_es:
            pe['vital_signs_table'] = {
                'title_es': 'Signos vitales al ingreso',
                'title_en': 'Vital signs on admission',
                'headers_es': ['Parámetro', 'Valor'],
                'headers_en': ['Parameter', 'Value'],
                'widths': [8, 8],
                'rows_es': rows_es,
                'rows_en': rows_en,
            }

    # Physical exam narrative
    exam_text = hc_data.get('physical_exam', '')
    if exam_text and 'findings_es' not in pe:
        pe['findings_es'] = exam_text.capitalize()
        pe['findings_en'] = '[AI_PENDING: Translate physical exam findings to English]'

    # ── Diagnoses list ────────────────────────────────────────────────────────
    if dx and 'diagnoses_list' not in cp:
        cp['diagnoses_list'] = [
            {'name': d.get('name', ''), 'code': d.get('code', ''),
             'type': d.get('type', '')}
            for d in dx
        ]

    # ── Clinical analysis ─────────────────────────────────────────────────────
    analysis = hc_data.get('clinical_analysis', '')
    if analysis and 'clinical_analysis_es' not in cp:
        cp['clinical_analysis_es'] = analysis.capitalize()
        cp['clinical_analysis_en'] = '[AI_PENDING: Translate clinical analysis to English]'
class CaseReportBuilder:
    def __init__(self, config_path):
        self.config_path = Path(config_path).resolve()
        self.base_dir = self.config_path.parent

        with open(config_path, 'r', encoding='utf-8') as f:
            self.cfg = yaml.safe_load(f)

        self.figs_dir = self.base_dir / self.cfg.get('figs_dir', 'figs')
        self.figs_dir.mkdir(exist_ok=True)

        # ── Auto-extract data from clinical history PDF if specified ──────────
        self._hc_data = {}
        sources = self.cfg.get('sources', {})
        hc_pdf = sources.get('clinical_history_pdf', '')
        if hc_pdf:
            hc_path = self.base_dir / hc_pdf
            if not hc_path.exists():
                print(f"AVISO: Historia clínica no encontrada: {hc_path}")
            elif not HAS_HC_PARSER:
                print("AVISO: parse_clinical_history.py no disponible. Datos del paciente no serán auto-extraídos.")
            else:
                print(f"\n── Extrayendo datos de historia clínica: {hc_pdf} ──")
                try:
                    hc_data = _parse_hc(str(hc_path))
                    _enrich_cfg_from_clinical_history(self.cfg, hc_data)
                    self._hc_data = hc_data
                    print(f"  ✓ Paciente: {hc_data['demographics'].get('name','?')}")
                    print(f"  ✓ Signos vitales extraídos: {len(hc_data.get('vital_signs',{}))}")
                    print(f"  ✓ Diagnósticos: {len(hc_data.get('diagnoses',[]))}")
                    print(f"  ✓ Notas de evolución: {len(hc_data.get('evolution_notes',[]))}")
                except Exception as e:
                    print(f"  AVISO: Error al parsear historia clínica: {e}")

        self.doc = Document()
        self._setup_document()
        self._tbl = [0]
        self._fig = [0]

    # ── Document setup ────────────────────────────────────────────────────────
    def _setup_document(self):
        for sec in self.doc.sections:
            sec.top_margin = sec.bottom_margin = Cm(2.5)
            sec.left_margin = sec.right_margin = Cm(2.5)
        self.doc.styles['Normal'].font.name = 'Arial'
        self.doc.styles['Normal'].font.size = Pt(11)
        for i, (sz, col) in enumerate([(14, DARK_BLUE), (12, MED_BLUE), (11, DARK_BLUE)], 1):
            hs = self.doc.styles[f'Heading {i}']
            hs.font.name = 'Arial'
            hs.font.size = Pt(sz)
            hs.font.color.rgb = col
            hs.font.bold = True

    # ── Content helpers ───────────────────────────────────────────────────────
    def h1(self, text):
        p = self.doc.add_heading(text, level=1)
        set_spacing(p, before=12, after=6)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single')
        bot.set(qn('w:sz'), '8')
        bot.set(qn('w:space'), '4')
        bot.set(qn('w:color'), '2E75B6')
        pBdr.append(bot)
        pPr.append(pBdr)

    def h2(self, text):
        p = self.doc.add_heading(text, level=2)
        set_spacing(p, before=10, after=4)

    def body(self, text, bold=False, italic=False, color=None,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=4, after=4):
        p = self.doc.add_paragraph()
        p.alignment = align
        set_spacing(p, before=before, after=after, line=276)
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        if color:
            r.font.color.rgb = color
        return p

    def ai_placeholder(self, text, before=4, after=4):
        """Renderiza un marcador AI_PENDIENTE en color naranja."""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_spacing(p, before=before, after=after, line=276)
        r = p.add_run(f'⚠ {text}')
        r.bold = True
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        r.font.color.rgb = ORANGE
        return p

    def body_or_placeholder(self, text, **kwargs):
        """Muestra texto normal o marcador si es AI_PENDIENTE."""
        if is_ai_placeholder(text):
            self.ai_placeholder(text, **kwargs)
        else:
            self.body(text, **kwargs)

    def mixed(self, parts, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=4, after=4):
        """Párrafo con partes de texto con formato mixto.
        parts: lista de dicts con claves: text, bold, italic, sup, color_red, color
        """
        p = self.doc.add_paragraph()
        p.alignment = align
        set_spacing(p, before=before, after=after, line=276)
        self._add_parts_to_para(p, parts)

    def _add_parts_to_para(self, p, parts):
        """Agrega partes con formato a un párrafo."""
        for pt in parts:
            text = pt.get('text', '')
            r = p.add_run(text)
            r.bold = pt.get('bold', False)
            r.italic = pt.get('italic', False)
            r.font.name = 'Arial'
            r.font.size = Pt(pt.get('size', 11))
            r.font.superscript = pt.get('sup', False)
            if pt.get('color_red'):
                r.font.color.rgb = RED
            elif pt.get('color'):
                c = pt['color']
                if isinstance(c, RGBColor):
                    r.font.color.rgb = c
                else:
                    r.font.color.rgb = GRAY

    def mixed_or_placeholder(self, parts_or_text, **kwargs):
        """Muestra mixed() o placeholder según el tipo/valor."""
        if isinstance(parts_or_text, str):
            self.body_or_placeholder(parts_or_text, **kwargs)
        elif isinstance(parts_or_text, list):
            # Si la lista tiene un solo elemento de texto que es placeholder
            if len(parts_or_text) == 1 and is_ai_placeholder(parts_or_text[0].get('text', '')):
                self.ai_placeholder(parts_or_text[0]['text'], **kwargs)
            else:
                self.mixed(parts_or_text, **kwargs)

    def bullet(self, text):
        if is_ai_placeholder(text):
            p = self.doc.add_paragraph(style='List Bullet')
            r = p.add_run(f'⚠ {text}')
            r.font.name = 'Arial'
            r.font.size = Pt(11)
            r.bold = True
            r.font.color.rgb = ORANGE
        else:
            p = self.doc.add_paragraph(style='List Bullet')
            r = p.add_run(text)
            r.font.name = 'Arial'
            r.font.size = Pt(11)
        set_spacing(p, before=3, after=3)

    def mixed_bullet(self, parts):
        p = self.doc.add_paragraph(style='List Bullet')
        set_spacing(p, before=3, after=3)
        self._add_parts_to_para(p, parts)

    def numbered(self, text):
        p = self.doc.add_paragraph(style='List Number')
        r = p.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        set_spacing(p, before=3, after=3)

    def space(self):
        p = self.doc.add_paragraph()
        set_spacing(p, before=2, after=2)

    def cline(self, text, bold=False, italic=False, size=11, color=None,
              before=2, after=2):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=before, after=after)
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.name = 'Arial'
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color

    def table(self, title, headers, rows, widths):
        """Genera una tabla con encabezado azul y filas alternadas."""
        self._tbl[0] += 1
        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_spacing(cap, before=10, after=4)
        r1 = cap.add_run(f'Cuadro {self._tbl[0]}. ')
        r1.bold = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(11)
        r1.font.color.rgb = DARK_BLUE
        r2 = cap.add_run(title)
        r2.font.name = 'Arial'
        r2.font.size = Pt(11)

        t = self.doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, w in enumerate(widths):
            for c in t.columns[i].cells:
                c.width = Cm(w)

        hdr = t.rows[0].cells
        for i, h in enumerate(headers):
            c = hdr[i]
            set_cell_bg(c, '2E75B6')
            set_cell_margins(c)
            set_cell_borders(c, '2E75B6', 6)
            pp = c.paragraphs[0]
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_spacing(pp, before=2, after=2)
            rr = pp.add_run(h)
            rr.bold = True
            rr.font.color.rgb = WHITE
            rr.font.name = 'Arial'
            rr.font.size = Pt(10)

        for ri, rd in enumerate(rows):
            row = t.add_row()
            bg = 'EBF3FB' if ri % 2 == 0 else 'FFFFFF'
            for ci, cd in enumerate(rd):
                c = row.cells[ci]
                set_cell_bg(c, bg)
                set_cell_margins(c)
                set_cell_borders(c, 'CCCCCC', 4)
                pp = c.paragraphs[0]
                pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                set_spacing(pp, before=2, after=2)
                if isinstance(cd, str):
                    rr = pp.add_run(cd)
                    rr.font.name = 'Arial'
                    rr.font.size = Pt(10)
                elif isinstance(cd, list):
                    for pt in cd:
                        rr = pp.add_run(pt.get('text', ''))
                        rr.font.name = 'Arial'
                        rr.font.size = Pt(10)
                        rr.bold = pt.get('bold', False)
                        rr.italic = pt.get('italic', False)
                        rr.font.superscript = pt.get('sup', False)
                        if pt.get('color_red'):
                            rr.font.color.rgb = RED
        self.space()

    def fig(self, path, w_cm, caption):
        """Inserta una figura centrada con su leyenda."""
        img_path = self._resolve_image(path)
        if img_path is None:
            self._missing_image_placeholder(path, caption)
            return
        self._fig[0] += 1
        n = self._fig[0]
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=8, after=2)
        p.add_run().add_picture(str(img_path), width=Cm(w_cm))
        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(cap, before=2, after=10)
        r1 = cap.add_run(f'Figura {n}. ')
        r1.bold = True
        r1.italic = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(10)
        r1.font.color.rgb = GRAY
        r2 = cap.add_run(caption)
        r2.italic = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(10)
        r2.font.color.rgb = GRAY

    def fig_row(self, items):
        """Muestra una fila de imágenes side-by-side.
        items: lista de (path, width_cm, caption)
        """
        n0 = self._fig[0] + 1
        available = [i for i in items if self._resolve_image(i[0]) is not None]
        missing = [i for i in items if self._resolve_image(i[0]) is None]

        for item in missing:
            self._fig[0] += 1
            self._missing_image_placeholder(item[0], item[2])

        if not available:
            return

        for _ in available:
            self._fig[0] += 1

        t = self.doc.add_table(rows=2, cols=len(available))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        cw = 16.0 / len(available)
        for i, (path, w_cm, caption) in enumerate(available):
            img_path = self._resolve_image(path)
            c = t.cell(0, i)
            c.width = Cm(cw)
            pp = c.paragraphs[0]
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_spacing(pp, before=4, after=2)
            pp.add_run().add_picture(str(img_path), width=Cm(min(w_cm, cw - 0.3)))
            cc = t.cell(1, i)
            capp = cc.paragraphs[0]
            capp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_spacing(capp, before=2, after=8)
            fn = n0 + i
            r1 = capp.add_run(f'Figura {fn}. ')
            r1.bold = True
            r1.italic = True
            r1.font.name = 'Arial'
            r1.font.size = Pt(9)
            r1.font.color.rgb = GRAY
            r2 = capp.add_run(caption)
            r2.italic = True
            r2.font.name = 'Arial'
            r2.font.size = Pt(9)
            r2.font.color.rgb = GRAY
            set_cell_borders(c, 'FFFFFF', 0)
            set_cell_borders(cc, 'FFFFFF', 0)
        self.space()

    def _resolve_image(self, filename):
        """Busca la imagen en figs_dir o base_dir. Devuelve Path o None."""
        if filename is None:
            return None
        candidates = [
            self.figs_dir / filename,
            self.base_dir / filename,
            Path(filename),
        ]
        for c in candidates:
            if c.exists():
                return c
        return None

    def _missing_image_placeholder(self, path, caption):
        """Inserta un marcador de imagen faltante."""
        self._fig[0] += 1
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=8, after=2)
        r = p.add_run(f'[⚠ IMAGEN FALTANTE: {path}]')
        r.bold = True
        r.font.name = 'Arial'
        r.font.size = Pt(10)
        r.font.color.rgb = ORANGE
        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(cap, before=2, after=10)
        r1 = cap.add_run(f'Figura {self._fig[0]}. ')
        r1.bold = True
        r1.italic = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(10)
        r1.font.color.rgb = GRAY
        r2 = cap.add_run(caption)
        r2.italic = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(10)
        r2.font.color.rgb = GRAY

    # ── Video frame extraction ────────────────────────────────────────────────
    def extract_video_frames(self):
        """Extrae fotogramas de los videos especificados en el config."""
        if not HAS_CV2:
            print("AVISO: opencv no disponible. Omitiendo extracción de fotogramas.")
            return

        frames_cfg = self.cfg.get('video_frames', [])
        if not frames_cfg:
            return

        print(f"\n── Extrayendo fotogramas de video ({len(frames_cfg)} frames) ──")
        for fc in frames_cfg:
            out_name = fc.get('output', '')
            out_path = self.figs_dir / out_name
            if out_path.exists():
                print(f"  ✓ Ya existe: {out_name}")
                continue

            video_file = fc.get('video', '')
            video_path = self.base_dir / video_file
            if not video_path.exists():
                print(f"  ✗ Video no encontrado: {video_file}")
                continue

            ts = parse_timestamp(fc.get('timestamp_sec', 0))
            cap = cv2.VideoCapture(str(video_path))
            fps = cap.get(cv2.CAP_PROP_FPS)
            total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
            duration = total_frames / fps if fps > 0 else 0

            if ts > duration:
                ts = duration * 0.5
                print(f"  ⚠ Timestamp {fc['timestamp_sec']}s excede duración {duration:.1f}s. "
                      f"Usando t={ts:.1f}s para {out_name}")

            frame_num = int(ts * fps)
            cap.set(cv2.CAP_PROP_POS_FRAMES, frame_num)
            ret, frame = cap.read()
            cap.release()

            if ret:
                cv2.imwrite(str(out_path), frame)
                desc = fc.get('description', '')
                print(f"  ✓ Extraído: {out_name} (t={ts:.1f}s) — {desc}")
            else:
                print(f"  ✗ No se pudo extraer frame de {video_file} en t={ts:.1f}s")

    # ── Image copies ──────────────────────────────────────────────────────────
    def copy_images(self):
        """Copia/renombra imágenes hacia figs_dir según image_copies del config."""
        copies_cfg = self.cfg.get('image_copies', [])
        if not copies_cfg:
            return
        print(f"\n── Copiando imágenes ({len(copies_cfg)} archivos) ──")
        for ic in copies_cfg:
            src_name = ic.get('source', '')
            dst_name = ic.get('dest', '')
            dst_path = self.figs_dir / dst_name
            if dst_path.exists():
                print(f"  ✓ Ya existe: {dst_name}")
                continue
            src_path = self.base_dir / src_name
            if not src_path.exists():
                print(f"  ✗ Fuente no encontrada: {src_name}")
                continue
            shutil.copy2(str(src_path), str(dst_path))
            desc = ic.get('description', '')
            print(f"  ✓ Copiado: {src_name} → {dst_name} — {desc}")

    # ── Text helpers with lang support ───────────────────────────────────────
    def _t(self, obj, key, lang, default=''):
        """Obtiene texto bilingüe de un dict de config."""
        if obj is None:
            return default
        # Primero intentar key_lang, luego key
        val = obj.get(f'{key}_{lang}', obj.get(key, default))
        return val if val is not None else default

    def _parts(self, obj, key, lang):
        """Obtiene lista de partes (para mixed()) o None."""
        return obj.get(f'parts_{lang}', obj.get(f'{key}_{lang}', None))

    # ── Separator between ES and EN versions ─────────────────────────────────
    def build_separator(self):
        self.doc.add_page_break()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=60, after=10)
        r = p.add_run('━' * 55)
        r.font.name = 'Arial'
        r.font.size = Pt(14)
        r.font.color.rgb = MED_BLUE

        p2 = self.doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p2, before=10, after=10)
        r2 = p2.add_run('ENGLISH VERSION')
        r2.bold = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(20)
        r2.font.color.rgb = DARK_BLUE

        p3 = self.doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p3, before=4, after=10)
        r3 = p3.add_run('Complete text in English  /  Texto completo en inglés')
        r3.italic = True
        r3.font.name = 'Arial'
        r3.font.size = Pt(12)
        r3.font.color.rgb = GRAY

        p4 = self.doc.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p4, before=4, after=60)
        r4 = p4.add_run('━' * 55)
        r4.font.name = 'Arial'
        r4.font.size = Pt(14)
        r4.font.color.rgb = MED_BLUE

    def reset_counters(self):
        self._tbl[0] = 0
        self._fig[0] = 0

    # ══════════════════════════════════════════════════════════════════════════
    # DOCUMENT SECTIONS
    # ══════════════════════════════════════════════════════════════════════════

    def build_cover(self, lang):
        """Construye la portada / cover page."""
        inst = self.cfg.get('institution', {})
        ministry = self._t(inst, 'ministry', lang, 'INSTITUCIÓN')
        name = self._t(inst, 'name', lang, '')
        dept = self._t(inst, 'department', lang, '')
        city = inst.get('city', '')

        self.cline(ministry, bold=True, size=11, color=MED_BLUE, before=0, after=3)
        self.cline(f'{name} | {dept} | {city}', size=10, color=GRAY, before=0, after=6)
        self.space()

        # Título
        title_cfg = self.cfg.get('title', {})
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=8, after=6)

        title_text = self._t(title_cfg, '', lang, '[TÍTULO FALTANTE]')
        # Manejar '' key para ES/EN
        title_text = title_cfg.get(lang, '[TÍTULO FALTANTE]')
        r = p.add_run(title_text)
        r.bold = True
        r.font.name = 'Arial'
        r.font.size = Pt(14)
        r.font.color.rgb = DARK_BLUE

        suffix_italic = title_cfg.get(f'suffix_italic_{lang}', '')
        if suffix_italic:
            r2 = p.add_run(suffix_italic)
            r2.bold = True
            r2.italic = True
            r2.font.name = 'Arial'
            r2.font.size = Pt(14)
            r2.font.color.rgb = DARK_BLUE

        suffix_end = title_cfg.get(f'suffix_end_{lang}', '')
        if suffix_end:
            r3 = p.add_run(suffix_end)
            r3.bold = True
            r3.font.name = 'Arial'
            r3.font.size = Pt(14)
            r3.font.color.rgb = DARK_BLUE

        self.space()

        # Autores
        authors = self.cfg.get('authors', [])
        for a in authors:
            name_str = a.get('name', '')
            suffix = a.get(f'suffix_{lang}', a.get('suffix_es', ''))
            self.cline(f'{name_str}{suffix}', size=11, before=1, after=1)
        self.space()

        # Afiliaciones
        affiliations = self.cfg.get('affiliations', {})
        for num, aff in affiliations.items():
            aff_text = aff.get(lang, aff.get('es', ''))
            self.cline(f'{num} {aff_text}', italic=True, size=9, color=GRAY, before=0, after=1)

        # Nota del supervisor
        for a in authors:
            if a.get('supervisor'):
                note = a.get(f'supervisor_note_{lang}', '')
                if note:
                    sup_name = a.get('name', '')
                    self.cline(f'* {sup_name} — {note}', italic=True, size=9, color=GRAY, before=0, after=2)

        # ORCIDs
        orcid_parts = []
        for a in authors:
            orcid = a.get('orcid', '')
            if orcid:
                # Usar iniciales del nombre
                parts = a.get('name', '').split()
                initials = '.'.join(p[0] for p in parts[:2] if p) + '.'
                orcid_parts.append(f'{initials}: {orcid}')
        if orcid_parts:
            self.cline('ORCIDs: ' + ' | '.join(orcid_parts),
                       italic=True, size=9, color=GRAY, before=0, after=6)

        self.space()

        # Correspondencia
        corr_label = 'Correspondencia:' if lang == 'es' else 'Corresponding author:'
        self.cline(corr_label, bold=True, size=10, before=4, after=2)
        for a in authors:
            if a.get('corresponding'):
                self.cline(a.get('name', ''), size=10, before=1, after=1)
                email = a.get('email', '')
                orcid = a.get('orcid', '')
                if email or orcid:
                    self.cline(f'{email}  |  ORCID: {orcid}', size=10, before=1, after=1)
                inst_city = a.get('institution_city', '')
                if inst_city:
                    self.cline(inst_city, size=10, before=1, after=6)

        conflict_text = self.cfg.get(f'conflict_funding_consent_{lang}', '')
        if conflict_text:
            self.cline(conflict_text, italic=True, size=9, color=GRAY, before=4, after=2)

    def build_abstract(self, lang):
        """Construye el resumen / abstract."""
        self.doc.add_page_break()
        title = 'RESUMEN' if lang == 'es' else 'ABSTRACT'
        self.h1(title)

        ab = self.cfg.get('abstract', {})

        intro_label = 'Introducción: ' if lang == 'es' else 'Introduction: '
        intro_text = self._t(ab, 'introduction', lang, '')
        self.mixed([{'text': intro_label, 'bold': True}, {'text': intro_text}])

        case_label = 'Presentación del caso: ' if lang == 'es' else 'Case Presentation: '
        case_text = self._t(ab, 'case', lang, '')
        if is_ai_placeholder(case_text):
            self.mixed([{'text': case_label, 'bold': True}])
            self.ai_placeholder(case_text)
        else:
            self.mixed([{'text': case_label, 'bold': True}, {'text': case_text}])

        conc_label = self._t(ab, 'conclusions_label', lang, 'Conclusiones') + ': '
        conc_text = self._t(ab, 'conclusions', lang, '')
        if is_ai_placeholder(conc_text):
            self.mixed([{'text': conc_label, 'bold': True}])
            self.ai_placeholder(conc_text)
        else:
            self.mixed([{'text': conc_label, 'bold': True}, {'text': conc_text}])

        kw_label = 'Palabras clave: ' if lang == 'es' else 'Keywords: '
        kw_text = self._t(ab, 'keywords', lang, '')
        self.mixed([{'text': kw_label, 'bold': True, 'italic': True}, {'text': kw_text, 'italic': True}])

    def build_introduction(self, lang):
        """Construye la introducción."""
        self.doc.add_page_break()
        num = '1'
        title = f'{num}. INTRODUCCIÓN' if lang == 'es' else f'{num}. INTRODUCTION'
        self.h1(title)

        intro_cfg = self.cfg.get('introduction', {})
        for sub in intro_cfg.get('subsections', []):
            sub_title = self._t(sub, 'title', lang, '')
            self.h2(sub_title)
            text = self._t(sub, 'text', lang, '')
            parts = sub.get(f'parts_{lang}', None)
            if parts:
                self.mixed(parts)
            else:
                self.body_or_placeholder(text)

        obj_text = self._t(intro_cfg, 'objective', lang, '')
        if obj_text:
            parts = intro_cfg.get(f'objective_parts_{lang}', None)
            if parts:
                self.mixed(parts)
            else:
                self.body_or_placeholder(obj_text)

    def build_case_presentation(self, lang):
        """Construye la presentación del caso."""
        self.doc.add_page_break()
        num = '2'
        title = f'{num}. PRESENTACIÓN DEL CASO' if lang == 'es' else f'{num}. CASE PRESENTATION'
        self.h1(title)

        cp = self.cfg.get('case_presentation', {})
        consent = self._t(cp, 'consent_statement', lang, '')
        if consent:
            self.body(consent)

        # 2.1 Datos generales / Demographics
        demo_title = '2.1 Datos generales' if lang == 'es' else '2.1 Patient demographics'
        self.h2(demo_title)
        dt = cp.get('demographics_table', {})
        if dt:
            self.table(
                self._t(dt, 'title', lang, ''),
                dt.get(f'headers_{lang}', []),
                dt.get(f'rows_{lang}', []),
                dt.get('widths', [6, 10])
            )

        # 2.2 Motivo de consulta / Chief complaint
        cc_title = '2.2 Motivo de consulta y antecedentes' if lang == 'es' else '2.2 Chief complaint and medical history'
        self.h2(cc_title)
        cc = cp.get('chief_complaint', {})

        # Motivo de consulta
        narrative = cc.get(f'narrative_{lang}', '')
        if narrative:
            if isinstance(narrative, list):
                self.mixed(narrative)
            else:
                self.body_or_placeholder(str(narrative))

        # Antecedentes médicos / medical history
        mh_label_default = 'Antecedentes patológicos: ' if lang == 'es' else 'Medical history: '
        mh_label = self._t(cc, 'medical_history_label', lang, mh_label_default)
        mh_text = self._t(cc, 'medical_history', lang, '')
        if mh_text:
            self.mixed([{'text': mh_label, 'bold': True}, {'text': mh_text}])

        # Antecedentes quirúrgicos / surgical history
        sx_label_default = 'Antecedentes quirúrgicos: ' if lang == 'es' else 'Surgical history: '
        sx_text = cc.get(f'surgical_history_{lang}', cc.get('surgical_history_es', ''))
        if sx_text:
            self.mixed([{'text': sx_label_default, 'bold': True}, {'text': sx_text}])

        # Alergias / Allergies
        al_label = 'Alergias: ' if lang == 'es' else 'Allergies: '
        al_text = cc.get('allergies_es', '')
        if al_text:
            self.mixed([{'text': al_label, 'bold': True}, {'text': al_text}])

        # Antecedentes gineco-obstétricos (solo ES)
        if lang == 'es':
            gy_text = cc.get('gyneco_es', '')
            if gy_text:
                self.mixed([{'text': 'Gineco-obstétricos: ', 'bold': True}, {'text': gy_text}])

        # Hábitos
        hab_label = 'Hábitos: ' if lang == 'es' else 'Habits: '
        hab_text = cc.get('habits_es', '')
        if hab_text:
            self.mixed([{'text': hab_label, 'bold': True}, {'text': hab_text}])

        # Enfermedad actual / Present illness
        ea_label = 'Enfermedad actual: ' if lang == 'es' else 'Present illness: '
        ea_text = cc.get(f'present_illness_{lang}', cc.get('present_illness_es', ''))
        if ea_text:
            self.mixed([{'text': ea_label, 'bold': True}])
            self.body_or_placeholder(ea_text)

        # Nota pie de laboratorio
        lab_fn = self._t(cc, 'lab_footnote', lang, '')
        if lab_fn:
            self.body(lab_fn, italic=True, color=GRAY, before=0, after=4)

        # 2.3 Examen físico / Physical exam
        pe_title = '2.3 Examen físico al ingreso' if lang == 'es' else '2.3 Physical exam on admission'
        self.h2(pe_title)
        pe = cp.get('physical_exam', {})
        vs_table = pe.get('vital_signs_table', {})
        if vs_table:
            self.table(
                self._t(vs_table, 'title', lang, ''),
                vs_table.get(f'headers_{lang}', []),
                vs_table.get(f'rows_{lang}', []),
                vs_table.get('widths', [8, 8])
            )

        # Hallazgos del examen físico
        findings_text = self._t(pe, 'findings', lang, '')
        if not findings_text:
            # Legacy format: findings_label + findings_bullets
            findings_label = self._t(pe, 'findings_label', lang, '')
            if findings_label:
                self.body(findings_label)
            for item in pe.get(f'findings_bullets_{lang}', []):
                self.bullet(item)
        else:
            self.body_or_placeholder(findings_text)

        # 2.4 Análisis clínico / Clinical analysis
        analysis = cp.get(f'clinical_analysis_{lang}', cp.get('clinical_analysis_es', ''))
        if analysis:
            an_title = '2.4 Análisis clínico' if lang == 'es' else '2.4 Clinical analysis'
            self.h2(an_title)
            self.body_or_placeholder(analysis)

        # 2.5 Diagnósticos / Diagnoses
        dx_list = cp.get('diagnoses_list', [])
        if dx_list:
            dx_title = '2.5 Diagnósticos (CIE-10)' if lang == 'es' else '2.5 Diagnoses (ICD-10)'
            self.h2(dx_title)
            for d in dx_list:
                name = d.get('name', '')
                code = d.get('code', '')
                dtype = d.get('type', '')
                type_label = f' [{dtype}]' if dtype else ''
                self.bullet(f'[{code}]{type_label} {name}')

        # 2.6 Laboratorios preingreso
        pil = cp.get('preingreso_labs', {})
        lab_table = pil.get('table', {})
        if lab_table:
            lab_title = '2.6 Exámenes de laboratorio preingreso' if lang == 'es' else '2.6 Pre-admission laboratory results'
            self.h2(lab_title)
            self.table(
                self._t(lab_table, 'title', lang, ''),
                lab_table.get(f'headers_{lang}', []),
                lab_table.get(f'rows_{lang}', []),
                lab_table.get('widths', [])
            )

    def build_imaging(self, lang):
        """Construye la sección de estudios de imagen."""
        self.doc.add_page_break()
        img_cfg = self.cfg.get('imaging', {})
        sec_title = self._t(img_cfg, 'section_title', lang, '3. ESTUDIOS DE IMAGEN')
        self.h1(sec_title)

        for group in img_cfg.get('image_groups', []):
            self.h2(self._t(group, 'subsection_title', lang, ''))
            narrative = group.get(f'narrative_{lang}', '')
            if isinstance(narrative, list):
                self.mixed(narrative)
            else:
                self.body_or_placeholder(str(narrative))

            images = group.get('images', [])
            if images:
                items = [(img['file'], img.get('width_cm', 5.5),
                          self._t(img, 'caption', lang, '')) for img in images]
                self.fig_row(items)

        # Hallazgos clave
        kf = img_cfg.get('key_findings', {})
        if kf:
            self.h2(self._t(kf, 'title', lang, ''))
            quote = self._t(kf, 'ct_report_quote', lang, '')
            if quote:
                self.body(quote, italic=True)
            intro = self._t(kf, 'findings_intro', lang, '')
            if intro:
                self.body(intro)
            for item in kf.get(f'findings_bullets_{lang}', []):
                self.bullet(item)

        # Foto clínica
        cp_img = img_cfg.get('clinical_photo', {})
        if cp_img:
            self.h2(self._t(cp_img, 'title', lang, ''))
            intro_text = self._t(cp_img, 'intro', lang, '')
            if intro_text:
                self.body(intro_text)
            self.fig(
                cp_img.get('file', ''),
                cp_img.get('width_cm', 10),
                self._t(cp_img, 'caption', lang, '')
            )

    def build_multidisciplinary(self, lang):
        """Construye la evaluación multidisciplinaria."""
        self.doc.add_page_break()
        num = '4'
        title = f'{num}. EVALUACIÓN MULTIDISCIPLINARIA' if lang == 'es' else f'{num}. MULTIDISCIPLINARY ASSESSMENT'
        self.h1(title)

        md = self.cfg.get('multidisciplinary', {})
        for sub in md.get('subsections', []):
            self.h2(self._t(sub, 'title', lang, ''))
            parts = sub.get(f'parts_{lang}', None)
            text = self._t(sub, 'text', lang, '')
            if parts:
                self.mixed_or_placeholder(parts)
            elif text:
                self.body_or_placeholder(text)

    def build_surgical(self, lang):
        """Construye la sección de manejo quirúrgico."""
        self.doc.add_page_break()
        num = '5'
        title = f'{num}. MANEJO QUIRÚRGICO' if lang == 'es' else f'{num}. SURGICAL MANAGEMENT'
        self.h1(title)

        s = self.cfg.get('surgical', {})

        prep_title = '5.1 Preparación preoperatoria' if lang == 'es' else '5.1 Preoperative preparation'
        self.h2(prep_title)
        for item in s.get(f'prep_bullets_{lang}', []):
            self.bullet(item)

        tech = s.get('technique', {})
        if tech:
            self.h2(self._t(tech, 'title', lang, '5.2 Técnica quirúrgica'))
            date_label = self._t(tech, 'date_label', lang, 'Fecha: ')
            date_val = self._t(tech, 'date', lang, '')
            self.mixed([{'text': date_label, 'bold': True}, {'text': date_val}])

            surg_label = self._t(tech, 'surgeon_label', lang, 'Cirujano: ')
            surg_val = self._t(tech, 'surgeon', lang, '')
            self.mixed([{'text': surg_label, 'bold': True}, {'text': surg_val}])

            approach = tech.get(f'approach_parts_{lang}', None)
            if approach:
                self.mixed(approach)

            proc_label = self._t(tech, 'procedure_label', lang, '')
            if proc_label:
                self.body(proc_label)
            for step in tech.get(f'procedure_steps_{lang}', []):
                self.numbered(step)

            intraop = tech.get(f'intraop_parts_{lang}', None)
            if intraop:
                self.mixed(intraop)

            intraop_intro = self._t(tech, 'intraop_figs_intro', lang, '')
            if intraop_intro:
                self.body(intraop_intro)

            intraop_imgs = tech.get('intraop_images', [])
            if intraop_imgs:
                items = [(img['file'], img.get('width_cm', 8),
                          self._t(img, 'caption', lang, '')) for img in intraop_imgs]
                self.fig_row(items)

        # 5.3 Procedimientos subsecuentes
        sp = s.get('subsequent_procedures', {})
        if sp:
            self.doc.add_page_break()
            self.h2(self._t(sp, 'title', lang, '5.3 Procedimientos subsecuentes'))
            intro = self._t(sp, 'intro', lang, '')
            if intro:
                self.body(intro)
            sp_table = sp.get('table', {})
            if sp_table:
                self.table(
                    self._t(sp_table, 'title', lang, ''),
                    sp_table.get(f'headers_{lang}', []),
                    sp_table.get(f'rows_{lang}', []),
                    sp_table.get('widths', [])
                )

    def build_postoperative(self, lang):
        """Construye la evolución posoperatoria."""
        num = '6'
        title = f'{num}. EVOLUCIÓN POSOPERATORIA' if lang == 'es' else f'{num}. POSTOPERATIVE COURSE'
        self.h1(title)

        po = self.cfg.get('postoperative', {})

        # 6.1 Microbiología
        micro_title = '6.1 Resultados microbiológicos definitivos' if lang == 'es' else '6.1 Definitive microbiological results'
        self.h2(micro_title)
        mt = po.get('microbiology_table', {})
        if mt:
            self.table(
                self._t(mt, 'title', lang, ''),
                mt.get(f'headers_{lang}', []),
                mt.get(f'rows_{lang}', []),
                mt.get('widths', [])
            )

        # 6.2 Antibióticos
        ab_title = '6.2 Antibioticoterapia definitiva' if lang == 'es' else '6.2 Definitive antibiotic therapy'
        self.h2(ab_title)
        at = po.get('antibiotic_table', {})
        if at:
            self.table(
                self._t(at, 'title', lang, ''),
                at.get(f'headers_{lang}', []),
                at.get(f'rows_{lang}', []),
                at.get('widths', [])
            )

        # 6.3 Laboratorios
        lab_title = '6.3 Evolución de laboratorios' if lang == 'es' else '6.3 Laboratory evolution'
        self.h2(lab_title)
        lt = po.get('lab_evolution_table', {})
        if lt:
            self.table(
                self._t(lt, 'title', lang, ''),
                lt.get(f'headers_{lang}', []),
                lt.get(f'rows_{lang}', []),
                lt.get('widths', [])
            )
            fn = self._t(lt, 'footnote', lang, '')
            if fn:
                self.body(fn, italic=True, color=GRAY, before=0, after=4)
            narr = lt.get(f'narrative_parts_{lang}', None)
            if narr:
                self.mixed(narr)

        # 6.4 Fotos de herida
        wp = po.get('wound_photos', {})
        if wp:
            self.h2(self._t(wp, 'title', lang, '6.4 Fotos de herida'))
            intro = self._t(wp, 'intro', lang, '')
            if intro:
                self.body(intro)
            for img_row in wp.get('image_rows', []):
                images = img_row.get('images', [])
                items = [(img['file'], img.get('width_cm', 8),
                          self._t(img, 'caption', lang, '')) for img in images]
                self.fig_row(items)

        # 6.5 Complicaciones
        comp = po.get('complications', {})
        if comp:
            self.h2(self._t(comp, 'title', lang, '6.5 Complicaciones'))
            for item in comp.get(f'items_{lang}', []):
                label = item.get('label', '')
                text = item.get('text', '')
                self.mixed_bullet([
                    {'text': label, 'bold': True},
                    {'text': text}
                ])

    def build_discussion(self, lang):
        """Construye la discusión."""
        self.doc.add_page_break()
        num = '7'
        title = f'{num}. DISCUSIÓN' if lang == 'es' else f'{num}. DISCUSSION'
        self.h1(title)

        disc = self.cfg.get('discussion', {})

        # 7.1 Diagnóstico diferencial
        ddx_title = '7.1 Diagnóstico diferencial' if lang == 'es' else '7.1 Differential diagnosis'
        self.h2(ddx_title)
        ddx = disc.get('differential_dx_table', {})
        if ddx:
            self.table(
                self._t(ddx, 'title', lang, ''),
                ddx.get(f'headers_{lang}', []),
                ddx.get(f'rows_{lang}', []),
                ddx.get('widths', [])
            )

        # Subsecciones de discusión
        for sub in disc.get('subsections', []):
            sub_id = sub.get('id', '')
            sub_title = self._t(sub, 'title', lang, '')
            self.h2(sub_title)

            # Subsecciones especiales con tabla
            if sub_id == 'surgical_approaches':
                sa_table = disc.get('surgical_approaches_table', {})
                if sa_table:
                    self.table(
                        self._t(sa_table, 'title', lang, ''),
                        sa_table.get(f'headers_{lang}', []),
                        sa_table.get(f'rows_{lang}', []),
                        sa_table.get('widths', [])
                    )
                continue

            if sub_id == 'complications_discussion':
                for item in disc.get(f'complications_discussion_bullets_{lang}', []):
                    label = item.get('label', '')
                    text = item.get('text', '')
                    sup = item.get('sup', '')
                    parts = [{'text': label, 'bold': True}, {'text': text}]
                    if sup:
                        parts.append({'text': sup, 'sup': True})
                    self.mixed_bullet(parts)
                continue

            if sub_id == 'limitations':
                for item in disc.get(f'limitations_bullets_{lang}', []):
                    label = item.get('label', '')
                    text = item.get('text', '')
                    self.mixed_bullet([
                        {'text': label, 'bold': True},
                        {'text': text}
                    ])
                continue

            # Texto o partes genérico
            parts = sub.get(f'parts_{lang}', None)
            text = self._t(sub, 'text', lang, '')
            if parts:
                self.mixed_or_placeholder(parts)
            else:
                self.body_or_placeholder(text)

    def build_conclusions(self, lang):
        """Construye las conclusiones."""
        self.doc.add_page_break()
        num = '8'
        title = f'{num}. CONCLUSIONES' if lang == 'es' else f'{num}. CONCLUSIONS'
        self.h1(title)

        conc = self.cfg.get('conclusions', {})
        for item in conc.get(f'items_{lang}', []):
            self.bullet(item)

    def build_references(self, lang):
        """Construye las referencias en estilo Vancouver."""
        self.doc.add_page_break()
        num = '9'
        title = f'{num}. REFERENCIAS' if lang == 'es' else f'{num}. REFERENCES'
        self.h1(title)

        style_note = 'Estilo Vancouver. Numeración por orden de primera citación en el texto.' if lang == 'es' \
            else 'Vancouver style. Numbered in order of first citation in text.'
        self.body(style_note, italic=True, color=GRAY)
        self.space()

        refs = self.cfg.get('references', [])
        for i, ref in enumerate(refs, 1):
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            set_spacing(p, before=3, after=3, line=276)
            rn = p.add_run(f'{i}. ')
            rn.bold = True
            rn.font.name = 'Arial'
            rn.font.size = Pt(11)
            rt = p.add_run(ref)
            rt.font.name = 'Arial'
            rt.font.size = Pt(11)

    def build_author_contributions(self, lang):
        """Construye la declaración CRediT de contribución de autores."""
        self.doc.add_page_break()
        title_es = 'DECLARACIÓN DE CONTRIBUCIÓN DE AUTORES'
        title_en = 'AUTHORSHIP CONTRIBUTION STATEMENT (CRediT)'
        self.h1(title_es if lang == 'es' else title_en)

        intro_es = 'De acuerdo con la taxonomía CRediT (Contributor Roles Taxonomy), las contribuciones de cada autor al presente reporte de caso son las siguientes:'
        intro_en = 'In accordance with the CRediT (Contributor Roles Taxonomy), author contributions to this case report are as follows:'
        self.body(intro_es if lang == 'es' else intro_en, before=4, after=6)

        ac = self.cfg.get('author_contributions', {})
        roles = ac.get(f'roles_{lang}', [])
        contributions = ac.get('contributions', [])
        authors = self.cfg.get('authors', [])

        if not roles or not authors:
            return

        # Construir encabezados: Rol + un col por autor (iniciales)
        author_names = []
        for a in authors:
            parts = a.get('name', '').split(',')[0].strip().split()
            # Iniciales: primeras dos palabras
            initials = '.'.join(p[0] for p in parts[:2] if p) + '.' if len(parts) >= 2 else parts[0] if parts else ''
            author_names.append(initials)

        headers = ['Rol / Role'] + author_names
        widths = [5.5] + [3.0] * len(authors)

        # Construir filas
        rows = []
        for contrib in contributions:
            idx = contrib.get('role_index', 0)
            if idx < len(roles):
                role_name = roles[idx]
                marks = contrib.get('marks', [''] * len(authors))
                rows.append([role_name] + marks)

        title_table = 'Contribuciones de autores según taxonomía CRediT' if lang == 'es' \
            else 'Author contributions according to CRediT taxonomy'
        self.table(title_table, headers, rows, widths)

        closing = ac.get(f'closing_statement_{lang}', '')
        if closing:
            self.body(closing, italic=True, before=6, after=4)

    # ══════════════════════════════════════════════════════════════════════════
    # MAIN BUILD
    # ══════════════════════════════════════════════════════════════════════════

    def build(self, skip_frames=False, figs_only=False):
        """Orquesta la construcción completa del documento."""
        print(f"\n{'='*60}")
        print(f"  Construyendo reporte: {self.cfg.get('output_file', 'case_report.docx')}")
        print(f"  Config: {self.config_path}")
        print(f"  Figuras: {self.figs_dir}")
        print(f"{'='*60}\n")

        # Paso 1: extraer fotogramas de video
        if not skip_frames:
            self.extract_video_frames()
        else:
            print("  (Extracción de fotogramas omitida con --skip-frames)")

        # Paso 2: copiar/renombrar imágenes
        self.copy_images()

        if figs_only:
            print("\n  --figs-only: omitiendo construcción del documento.")
            return

        # ── VERSIÓN ESPAÑOL ──────────────────────────────────────────────────
        print("\n── Construyendo versión en español ──")
        self.build_cover('es')
        self.build_abstract('es')
        self.build_introduction('es')
        self.build_case_presentation('es')
        self.build_imaging('es')
        self.build_multidisciplinary('es')
        self.build_surgical('es')
        self.doc.add_page_break()
        self.build_postoperative('es')
        self.build_discussion('es')
        self.build_conclusions('es')
        self.build_references('es')
        self.build_author_contributions('es')

        # ── SEPARADOR ────────────────────────────────────────────────────────
        self.build_separator()
        self.reset_counters()

        # ── VERSIÓN INGLÉS ───────────────────────────────────────────────────
        print("── Construyendo versión en inglés ──")
        self.doc.add_page_break()
        self.build_cover('en')
        self.build_abstract('en')
        self.build_introduction('en')
        self.build_case_presentation('en')
        self.build_imaging('en')
        self.build_multidisciplinary('en')
        self.build_surgical('en')
        self.doc.add_page_break()
        self.build_postoperative('en')
        self.build_discussion('en')
        self.build_conclusions('en')
        self.build_references('en')
        self.build_author_contributions('en')

        # ── GUARDAR ──────────────────────────────────────────────────────────
        output_path = self.base_dir / self.cfg.get('output_file', 'case_report.docx')
        self.doc.save(str(output_path))

        # Verificar integridad del ZIP (docx es un ZIP)
        try:
            with zipfile.ZipFile(str(output_path)) as zf:
                imgs = [n for n in zf.namelist() if 'media' in n]
                print(f"\n✓ Guardado: {output_path}")
                print(f"  ZIP OK — {len(zf.namelist())} archivos — {len(imgs)} imágenes embebidas")
        except Exception as e:
            print(f"\n✓ Guardado: {output_path} (verificación ZIP: {e})")

        # Resumen de marcadores AI_PENDIENTE
        self._print_ai_summary()

    def _print_ai_summary(self):
        """Imprime resumen de secciones que requieren procesamiento de IA."""
        print("\n" + "─" * 60)
        print("  SECCIONES PENDIENTES DE REVISIÓN / IA:")
        print("─" * 60)
        placeholders = self._find_placeholders(self.cfg)
        if placeholders:
            for p in placeholders:
                print(f"  ⚠ {p}")
        else:
            print("  ✓ No hay marcadores AI_PENDIENTE en el config.")
        print("─" * 60)

    def _find_placeholders(self, obj, path=''):
        """Busca recursivamente marcadores AI_PENDIENTE en el config."""
        results = []
        if isinstance(obj, str):
            if is_ai_placeholder(obj):
                results.append(f"{path}: {obj[:80]}...")
        elif isinstance(obj, dict):
            for k, v in obj.items():
                results.extend(self._find_placeholders(v, f"{path}.{k}" if path else k))
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                results.extend(self._find_placeholders(item, f"{path}[{i}]"))
        return results


# ─── main ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='Constructor genérico de reportes de caso clínico bilingüe (ES/EN)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python3 build_case_report.py --config case_config_template.yaml
  python3 build_case_report.py --config mi_caso.yaml --skip-frames
  python3 build_case_report.py --config mi_caso.yaml --figs-only

Flujo de trabajo:
  1. Copie case_config_template.yaml a mi_caso.yaml
  2. Edite mi_caso.yaml con los datos del nuevo caso
  3. (Opcional) Ejecute parse_lab_pdfs.py para extraer datos de laboratorio
  4. Ejecute este script para generar el documento final
  5. Complete las secciones marcadas con ⚠ en naranja (requieren IA o redacción)
        """
    )
    parser.add_argument('--config', required=True,
                        help='Ruta al archivo YAML de configuración del caso')
    parser.add_argument('--skip-frames', action='store_true',
                        help='Omitir extracción de fotogramas de video')
    parser.add_argument('--figs-only', action='store_true',
                        help='Solo extraer fotogramas y copiar imágenes (no construir documento)')
    args = parser.parse_args()

    if not Path(args.config).exists():
        print(f"Error: archivo de configuración no encontrado: {args.config}", file=sys.stderr)
        sys.exit(1)

    builder = CaseReportBuilder(args.config)
    builder.build(skip_frames=args.skip_frames, figs_only=args.figs_only)


if __name__ == '__main__':
    main()
