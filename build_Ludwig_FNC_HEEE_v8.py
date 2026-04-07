from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import zipfile

DARK_BLUE = RGBColor(0x1F, 0x38, 0x64)
MED_BLUE  = RGBColor(0x2E, 0x75, 0xB6)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
RED       = RGBColor(0xCC, 0x00, 0x00)
GRAY      = RGBColor(0x77, 0x77, 0x77)

F = '/home/claude/figs/'

# ─── XML helpers ─────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color); tcPr.append(shd)

def set_cell_margins(cell):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top',80),('bottom',80),('left',120),('right',120)]:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), str(val)); el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

def set_cell_borders(cell, color='CCCCCC', size=4):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def set_spacing(para, before=4, after=4, line=None):
    pPr = para._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before*20)); sp.set(qn('w:after'), str(after*20))
    if line: sp.set(qn('w:line'), str(line)); sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

# ─── Document setup ───────────────────────────────────────────────────────────
doc = Document()
for sec in doc.sections:
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.5)
doc.styles['Normal'].font.name = 'Arial'
doc.styles['Normal'].font.size = Pt(11)
for i, (sz, col) in enumerate([(14,DARK_BLUE),(12,MED_BLUE),(11,DARK_BLUE)], 1):
    hs = doc.styles[f'Heading {i}']
    hs.font.name = 'Arial'; hs.font.size = Pt(sz)
    hs.font.color.rgb = col; hs.font.bold = True

# ─── Content helpers ──────────────────────────────────────────────────────────
def h1(text):
    p = doc.add_heading(text, level=1); set_spacing(p, before=12, after=6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr'); bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'8')
    bot.set(qn('w:space'),'4'); bot.set(qn('w:color'),'2E75B6')
    pBdr.append(bot); pPr.append(pBdr)

def h2(text):
    p = doc.add_heading(text, level=2); set_spacing(p, before=10, after=4)

def body(text, bold=False, italic=False, color=None,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=4, after=4):
    p = doc.add_paragraph(); p.alignment = align
    set_spacing(p, before=before, after=after, line=276)
    r = p.add_run(text); r.bold = bold; r.italic = italic
    r.font.name = 'Arial'; r.font.size = Pt(11)
    if color: r.font.color.rgb = color
    return p

def mixed(parts, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=4, after=4):
    p = doc.add_paragraph(); p.alignment = align
    set_spacing(p, before=before, after=after, line=276)
    for pt in parts:
        r = p.add_run(pt.get('text',''))
        r.bold = pt.get('bold',False); r.italic = pt.get('italic',False)
        r.font.name = 'Arial'; r.font.size = Pt(pt.get('size',11))
        r.font.superscript = pt.get('sup',False)
        if pt.get('color'): r.font.color.rgb = pt['color']

def bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(text); r.font.name = 'Arial'; r.font.size = Pt(11)
    set_spacing(p, before=3, after=3)

def mixed_bullet(parts):
    p = doc.add_paragraph(style='List Bullet'); set_spacing(p, before=3, after=3)
    for pt in parts:
        r = p.add_run(pt.get('text',''))
        r.bold = pt.get('bold',False); r.italic = pt.get('italic',False)
        r.font.name = 'Arial'; r.font.size = Pt(11)
        r.font.superscript = pt.get('sup',False)
        if pt.get('color'): r.font.color.rgb = pt['color']

def numbered(text):
    p = doc.add_paragraph(style='List Number')
    r = p.add_run(text); r.font.name = 'Arial'; r.font.size = Pt(11)
    set_spacing(p, before=3, after=3)

def space():
    p = doc.add_paragraph(); set_spacing(p, before=2, after=2)

def cline(text, bold=False, italic=False, size=11, color=None, before=2, after=2):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p, before=before, after=after)
    r = p.add_run(text); r.bold = bold; r.italic = italic
    r.font.name = 'Arial'; r.font.size = Pt(size)
    if color: r.font.color.rgb = color

_tbl = [0]
def table(title, headers, rows, widths):
    _tbl[0] += 1
    cap = doc.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_spacing(cap, before=10, after=4)
    r1 = cap.add_run(f'Cuadro {_tbl[0]}. ')
    r1.bold = True; r1.font.name = 'Arial'; r1.font.size = Pt(11); r1.font.color.rgb = DARK_BLUE
    r2 = cap.add_run(title); r2.font.name = 'Arial'; r2.font.size = Pt(11)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, w in enumerate(widths):
        for c in t.columns[i].cells: c.width = Cm(w)
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        c = hdr[i]; set_cell_bg(c,'2E75B6'); set_cell_margins(c); set_cell_borders(c,'2E75B6',6)
        pp = c.paragraphs[0]; pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(pp, before=2, after=2)
        rr = pp.add_run(h); rr.bold = True; rr.font.color.rgb = WHITE
        rr.font.name = 'Arial'; rr.font.size = Pt(10)
    for ri, rd in enumerate(rows):
        row = t.add_row(); bg = 'EBF3FB' if ri%2==0 else 'FFFFFF'
        for ci, cd in enumerate(rd):
            c = row.cells[ci]; set_cell_bg(c,bg); set_cell_margins(c); set_cell_borders(c,'CCCCCC',4)
            pp = c.paragraphs[0]; pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_spacing(pp, before=2, after=2)
            if isinstance(cd, str):
                rr = pp.add_run(cd); rr.font.name = 'Arial'; rr.font.size = Pt(10)
            elif isinstance(cd, list):
                for pt in cd:
                    rr = pp.add_run(pt.get('text',''))
                    rr.font.name = 'Arial'; rr.font.size = Pt(10)
                    rr.bold = pt.get('bold',False); rr.italic = pt.get('italic',False)
                    rr.font.superscript = pt.get('sup',False)
                    if pt.get('color'): rr.font.color.rgb = pt['color']
    space()

_fig = [0]
def fig(path, w_cm, caption):
    _fig[0] += 1; n = _fig[0]
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p, before=8, after=2)
    p.add_run().add_picture(path, width=Cm(w_cm))
    cap = doc.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(cap, before=2, after=10)
    r1 = cap.add_run(f'Figura {n}. ')
    r1.bold = True; r1.italic = True; r1.font.name = 'Arial'
    r1.font.size = Pt(10); r1.font.color.rgb = GRAY
    r2 = cap.add_run(caption)
    r2.italic = True; r2.font.name = 'Arial'
    r2.font.size = Pt(10); r2.font.color.rgb = GRAY

def fig_row(items):
    """items = list of (path, w_cm, caption)"""
    n0 = _fig[0] + 1
    for _ in items: _fig[0] += 1
    t = doc.add_table(rows=2, cols=len(items))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cw = 16.0 / len(items)
    for i, (path, w_cm, caption) in enumerate(items):
        c = t.cell(0,i); c.width = Cm(cw)
        pp = c.paragraphs[0]; pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(pp, before=4, after=2)
        pp.add_run().add_picture(path, width=Cm(min(w_cm, cw-0.3)))
        cc = t.cell(1,i); capp = cc.paragraphs[0]
        capp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(capp, before=2, after=8)
        fn = n0 + i
        r1 = capp.add_run(f'Figura {fn}. ')
        r1.bold = True; r1.italic = True; r1.font.name = 'Arial'
        r1.font.size = Pt(9); r1.font.color.rgb = GRAY
        r2 = capp.add_run(caption)
        r2.italic = True; r2.font.name = 'Arial'
        r2.font.size = Pt(9); r2.font.color.rgb = GRAY
        set_cell_borders(c,'FFFFFF',0); set_cell_borders(cc,'FFFFFF',0)
    space()

# Shared references list
refs = [
    "Lamont G, Peterson DC. Ludwig Angina. In: StatPearls [Internet]. Treasure Island (FL): StatPearls Publishing; 2024 [citado/cited 2026 Jan 20]. Disponible en/Available from: https://www.ncbi.nlm.nih.gov/books/NBK470191/",
    "Sichel JY, Attal P, Hocwald E, Eliashar R. Redefining parapharyngeal space infections. Ann Otol Rhinol Laryngol. 2006;115(2):117-23. doi: 10.1177/000348940611500207.",
    "Huang TT, Liu TC, Chen PR, Tseng FY, Yeh TH, Chen YS. Deep neck infection: analysis of 185 cases. Head Neck. 2004;26(10):854-60. doi: 10.1002/hed.20014.",
    "Prado-Calleros HM, Jimenez-Fuentes E, Jimenez-Islas O. Descending necrotizing mediastinitis: systematic review on its treatment in the last 6 years, 2010-2015. Head Neck. 2016;38 Suppl 1:E2275-83. doi: 10.1002/hed.24408.",
    "Har-El G, Aroesty JH, Shaha A, Lucente FE. Changing trends in deep neck abscess. A retrospective study of 110 patients. Oral Surg Oral Med Oral Pathol. 1994;77(5):446-50. doi: 10.1016/0030-4220(94)90175-9.",
    "Vieira F, Allen SM, Stocks RM, Thompson JW. Deep neck infection. Otolaryngol Clin North Am. 2008;41(3):459-83. doi: 10.1016/j.otc.2008.01.002.",
    "Sartelli M, Guirao X, Hardcastle TC, Kluger Y, Catena F, Corsi D, et al. 2018 WSES/SIS-E consensus conference: recommendations for antibiotic therapy in abdominal infections in adult patients. World J Emerg Surg. 2019;14:40. doi: 10.1186/s13017-019-0259-0.",
    "Huang TT, Tseng FY, Liu TC, Hsu CJ, Chen YS. Deep neck infection in diabetic patients: comparison of clinical picture and outcomes with nondiabetic patients. Otolaryngol Head Neck Surg. 2005;132(6):943-7. doi: 10.1016/j.otohns.2005.01.028.",
    "Stevens DL, Bisno AL, Chambers HF, Dellinger EP, Goldstein EJ, Gorbach SL, et al. Practice guidelines for the diagnosis and management of skin and soft tissue infections: 2014 update by the IDSA. Clin Infect Dis. 2014;59(2):147-59. doi: 10.1093/cid/ciu296.",
    "Becker M, Zbaren P, Hermans R, Becker CD, Marchal F, Kurt AM, et al. Necrotizing fasciitis of the head and neck: role of CT in diagnosis and management. Radiology. 1997;202(2):471-6. doi: 10.1148/radiology.202.2.9015076.",
    "Lazor JB, Cunningham MJ, Eavey RD, Weber AL. Comparison of computed tomography and surgical findings in deep neck infections. Otolaryngol Head Neck Surg. 1994;111(6):746-50. doi: 10.1177/019459989411100607.",
    "Boscolo-Rizzo P, Stellin M, Muzzi E, Mantovani M, Fuson R, Lupato V, et al. Deep neck infections: a study of 365 cases highlighting recommendations for management and treatment. Eur Arch Otorhinolaryngol. 2012;269(4):1241-9. doi: 10.1007/s00405-011-1761-1.",
    "Marioni G, Staffieri A, Parisi S, Marchese-Ragona R, Zuccon A, Staffieri C, et al. Rational diagnostic and therapeutic management of deep neck infections: analysis of 233 consecutive cases. Ann Otol Rhinol Laryngol. 2010;119(3):181-7. doi: 10.1177/000348941011900307.",
    "Thottam PJ, Vimalachandran S, Shivakumar AM. Airway management in deep neck space infections: assessment of factors associated with difficult airway. J Laryngol Otol. 2012;126(8):835-9. doi: 10.1017/S0022215112001211.",
    "Rhodes A, Evans LE, Alhazzani W, Levy MM, Antonelli M, Ferrer R, et al. Surviving Sepsis Campaign: International Guidelines for Management of Sepsis and Septic Shock: 2016. Crit Care Med. 2017;45(3):486-552. doi: 10.1097/CCM.0000000000002255.",
    "Grisaru-Soen G, Komisar O, Aizenstein O, Soudack M, Schwartz D, Paret G. Retropharyngeal and parapharyngeal abscess in children: epidemiology, clinical features and treatment. Int J Pediatr Otorhinolaryngol. 2010;74(9):1016-20. doi: 10.1016/j.ijporl.2010.05.030.",
    "Parhiscar A, Har-El G. Deep neck abscess: a retrospective review of 210 cases. Ann Otol Rhinol Laryngol. 2001;110(11):1051-4. doi: 10.1177/000348940111001111.",
    "Endo S, Murayama F, Hasegawa T, Yamamoto S, Yamaguchi T, Sohara Y, et al. Guideline of surgical management based on diffusion of descending necrotizing mediastinitis. Jpn J Thorac Cardiovasc Surg. 1999;47(1):14-9. doi: 10.1007/BF03217709.",
    "Riordan T. Human infection with Fusobacterium necrophorum (Necrobacillosis), with a focus on Lemierre's syndrome. Clin Microbiol Rev. 2007;20(4):622-59. doi: 10.1128/CMR.00011-07.",
    "Kullberg BJ, Arendrup MC. Invasive fungal disease: management challenges and new treatment options for candidemia and aspergillosis. Br J Haematol. 2015;171(5):617-29. doi: 10.1111/bjh.13614.",
    "Mariano RC, de Moraes M, Santos FA, Ribeiro Junior PD, Campos AC. Ludwig's Angina: severity, clinical course and outcomes. Oral Maxillofac Surg. 2021;25(3):307-14. doi: 10.1007/s10006-020-00905-5.",
    "Martins JR, Chagas OL Jr, Weschenfelder F, Matos CM, Roehe AV, Rivero LF, et al. The use of antibiotics in odontogenic infections: what is the best choice? A systematic review. J Oral Maxillofac Surg. 2017;75(12):2606.e1-2606.e11. doi: 10.1016/j.joms.2017.08.003.",
    "Pinto A, Scaglione M, Scuderi MG, Tortora G, Daniele S, Romano L. Infections of the neck leading to descending necrotizing mediastinitis: role of multi-detector row computed tomography. Eur J Radiol. 2008;65(3):389-94. doi: 10.1016/j.ejrad.2007.04.015.",
    "Simon L, Gauvin F, Amre DK, Saint-Louis P, Lacroix J. Serum procalcitonin and C-reactive protein levels as markers of bacterial infection: a systematic review and meta-analysis. Clin Infect Dis. 2004;39(2):206-17. doi: 10.1086/421997.",
    "Marioni G, Rinaldi R, Staffieri C, Marchese-Ragona R, Ottaviano G, Staffieri A, et al. Deep neck infection with dental origin: analysis of 85 consecutive cases (2000-2006). Acta Otolaryngol. 2008;128(2):201-6. doi: 10.1080/00016480701387077.",
]

# ╔══════════════════════════════════════════════════════════════════════════════
# ██████████████████   VERSIÓN EN ESPAÑOL   ████████████████████████████████████
# ╔══════════════════════════════════════════════════════════════════════════════

# ── PORTADA ES ────────────────────────────────────────────────────────────────
cline('MINISTERIO DE SALUD PÚBLICA DEL ECUADOR', bold=True, size=11, color=MED_BLUE, before=0, after=3)
cline('Hospital de Especialidades Eugenio Espejo | Servicio de Otorrinolaringología | Quito, Ecuador',
      size=10, color=GRAY, before=0, after=6)
space()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_spacing(p, before=8, after=6)
r = p.add_run('ANGINA DE LUDWIG COMPLICADA CON FASCITIS NECROTIZANTE CERVICAL EN PACIENTE DIABÉTICA: REPORTE DE UN CASO CON AISLAMIENTO DE ')
r.bold = True; r.font.name = 'Arial'; r.font.size = Pt(14); r.font.color.rgb = DARK_BLUE
r2 = p.add_run('Klebsiella pneumoniae'); r2.bold = True; r2.italic = True
r2.font.name = 'Arial'; r2.font.size = Pt(14); r2.font.color.rgb = DARK_BLUE
r3 = p.add_run(' PRODUCTORA DE BLEE'); r3.bold = True
r3.font.name = 'Arial'; r3.font.size = Pt(14); r3.font.color.rgb = DARK_BLUE
space()

for a in ['Jean Pierre Gavilanez Heras, MD, MSc¹',
          'Christian Alfonso Freire Zamora, MD, Esp. ORL¹*',
          'Hernán Javier Campoverde Sánchez, MD, MSc¹',
          'Gabriel Saenz Ortega, MD²']:
    cline(a, size=11, before=1, after=1)
space()
cline('¹ Servicio de Otorrinolaringología, Hospital de Especialidades Eugenio Espejo (HEEE), Ministerio de Salud Pública del Ecuador. Quito, Ecuador.',
      italic=True, size=9, color=GRAY, before=0, after=1)
cline('² Médico independiente / Independent Physician. Quito, Ecuador.',
      italic=True, size=9, color=GRAY, before=0, after=1)
cline('* Dr. Christian Alfonso Freire Zamora — Especialista en ORL, Jefe del Servicio de ORL, HEEE. Tutor del caso.',
      italic=True, size=9, color=GRAY, before=0, after=2)
cline('ORCIDs: J.P. Gavilanez Heras: 0009-0008-6518-6041 | H.J. Campoverde Sánchez: 0009-0000-4802-0105 | G. Saenz Ortega: 0009-0009-5669-1027 | C.A. Freire Zamora: 0000-0003-2237-1462',
      italic=True, size=9, color=GRAY, before=0, after=6)
space()
cline('Correspondencia:', bold=True, size=10, before=4, after=2)
cline('Jean Pierre Gavilanez Heras, MD, MSc', size=10, before=1, after=1)
cline('jean.gavilanez@hee.gob.ec  |  ORCID: 0009-0008-6518-6041', size=10, before=1, after=1)
cline('Hospital de Especialidades Eugenio Espejo, Quito, Ecuador', size=10, before=1, after=6)
cline('Conflicto de intereses: Ninguno.  |  Financiamiento: Ninguno.  |  Consentimiento informado: Obtenido.',
      italic=True, size=9, color=GRAY, before=4, after=2)

# ── RESUMEN ES ────────────────────────────────────────────────────────────────
doc.add_page_break()
h1('RESUMEN')
mixed([{'text':'Introducción: ','bold':True},
       {'text':'La Angina de Ludwig es una celulitis indurada bilateral del suelo de la boca que, sin tratamiento oportuno, puede progresar a fascitis necrotizante cervical (FNC) con mortalidad del 15–40%. La diabetes mellitus (DM) es el principal factor de riesgo de progresión.'}])
mixed([{'text':'Presentación del caso: ','bold':True},
       {'text':'Paciente femenina, 49 años, con diabetes mellitus tipo 2 (DM2) de 11 años y extracción molar (14/12/2025). Desarrolló Angina de Ludwig con progresión a FNC en aproximadamente un mes de evolución sin atención oportuna, con necrosis cervical extensa, fístulas cutáneas y sepsis. Transferida el 16/01/2026 al HEEE (Quito) con leucocitosis (19.550 x10³/µL), anemia (Hb 8,9 g/dL), hiponatremia grave (Na 126,9 mEq/L) e hiperglucemia (glucosa 299,7 mg/dL). La TAC mostró proceso necrotizante cervical con gas en tejidos blandos sin extensión mediastínica. Desbridamiento quirúrgico el 18/01/2026 bajo anestesia general. Cultivos: '},
       {'text':'Klebsiella pneumoniae','italic':True},
       {'text':' productora de beta-lactamasas de espectro extendido (BLEE) + '},{'text':'Escherichia coli','italic':True},
       {'text':' BLEE (sensibles sólo a carbapenems) + '},
       {'text':'Candida tropicalis','italic':True},{'text':' y '},
       {'text':'Candida parapsilosis','italic':True},
       {'text':'. Hemocultivos negativos (4 sets). Meropenem como antibiótico definitivo; proteína C reactiva (PCR): 202 → 7,35 mg/L en 22 días. Estancia 33 días (16/01–18/02/2026) con reconstrucción por Cirugía Plástica.'}])
mixed([{'text':'Conclusiones: ','bold':True},
       {'text':'La Angina de Ludwig complicada con FNC con aislamiento de '},
       {'text':'K. pneumoniae','italic':True},
       {'text':' BLEE exige desbridamiento radical urgente y meropenem. El abordaje multidisciplinario ORL–UCI–Infectología–Cirugía Plástica fue determinante en el resultado favorable.'}])
mixed([{'text':'Palabras clave: ','bold':True,'italic':True},
       {'text':'Angina de Ludwig; fascitis necrotizante cervical; infección odontogénica; Klebsiella pneumoniae BLEE; diabetes mellitus; desbridamiento quirúrgico; meropenem; Hospital Eugenio Espejo.','italic':True}])

# ── INTRODUCCIÓN ES ───────────────────────────────────────────────────────────
doc.add_page_break()
h1('1. INTRODUCCIÓN')
h2('1.1 Angina de Ludwig: definición y anatomía')
mixed([{'text':'La Angina de Ludwig, descrita por Wilhelm Friedrich von Ludwig en 1836, se define como una celulitis indurada bilateral del suelo de la boca que compromete simultáneamente los espacios submandibular, submentoniano y sublingual, sin formación de absceso inicial.'},
       {'text':'1','sup':True},
       {'text':" La comunicación con el espacio retrofaríngeo y el espacio peligroso (danger space), que desciende hasta el mediastino posterior, explica la rápida diseminación descendente y el riesgo de mediastinitis descendente necrotizante (MDN)."},
       {'text':'2,3,4','sup':True}])
h2('1.2 Epidemiología y factores de riesgo')
mixed([{'text':'La etiología odontogénica representa el 40–60% de los casos en adultos, con predominio de afectación del segundo y tercer molar inferior.'},
       {'text':'5,6,25','sup':True},
       {'text':" La DM es el principal factor de riesgo para progresión a FNC, presente en el 50–80% de los casos, mediada por disfunción de neutrófilos, microangiopatía tisular y neuropatía periférica que retrasan el diagnóstico."},
       {'text':'8,21','sup':True}])
h2('1.3 Progresión a fascitis necrotizante cervical')
mixed([{'text':'La FNC se define por necrosis progresiva a lo largo de los planos fasciales cervicales, con mortalidad del 15–40% incluso con tratamiento adecuado.'},
       {'text':'4,9','sup':True},
       {'text':" La tríada diagnóstica incluye: (1) necrosis cutánea con escara oscura, (2) drenaje purulento fétido a través de fístulas y (3) disección fascial confirmada quirúrgicamente."},
       {'text':'10','sup':True}])

mixed([{'text':'El objetivo del presente reporte es describir el caso de una paciente con Angina de Ludwig complicada con fascitis necrotizante cervical (FNC), con aislamiento inusual de '},
       {'text':'Klebsiella pneumoniae','italic':True},
       {'text':' productora de beta-lactamasas de espectro extendido (BLEE) en una infección de inicio comunitario en el contexto de diabetes mellitus tipo 2 (DM2) sin control metabólico adecuado, y analizar los aspectos diagnósticos imagenológicos, microbiológicos, quirúrgicos y multidisciplinarios que determinaron el resultado favorable.'},
       {'text':'1,21','sup':True}])

# ── PRESENTACIÓN DEL CASO ES ──────────────────────────────────────────────────
doc.add_page_break()
h1('2. PRESENTACIÓN DEL CASO')
body('Presentado con consentimiento informado de la paciente, en cumplimiento de la Ley Orgánica de Salud del Ecuador. La identidad se protege mediante iniciales.')

h2('2.1 Datos generales')
table('Datos generales de la paciente',
    ['Parámetro','Dato'],
    [['Paciente (iniciales)','G.M.P.Q.'],
     ['Sexo / Edad','Femenino, 49 años (05/07/1976)'],
     ['Procedencia','Santo Domingo de los Tsáchilas, Ecuador'],
     ['Centro remitente','Hospital Dr. Gustavo Domínguez Zambrano, Santo Domingo'],
     ['Fecha de ingreso HEEE','16 de enero de 2026'],
     ['Servicio de ingreso','Otorrinolaringología — Piso 6, HEEE, Quito'],
     ['Historia clínica HEEE','0802060905'],
     ['Médico tutor / Jefe de servicio','Dr. Christian Alfonso Freire Zamora, MD, Esp. ORL — Jefe del Servicio ORL, HEEE. Tutor del caso.']],
    [6,10])

h2('2.2 Motivo de consulta y antecedentes')
mixed([{'text':'Cuadro de aproximadamente un mes de evolución (inicio: diciembre 2025) de edema mandibular progresivo que se extiende a región cervical y tórax, alzas térmicas, disfagia y aparición de fístulas cutáneas cervicales con drenaje purulento fétido. '},
       {'text':'Evento desencadenante: extracción molar el 14/12/2025','bold':True},
       {'text':' en contexto ambulatorio sin cobertura antibiótica documentada.'},
       {'text':'22','sup':True}])
mixed([{'text':'Antecedentes: ','bold':True},
       {'text':'DM2 × 11 años (metformina 500 mg VO). Cardiopatía hipertensiva sin IC congestiva (CIE-10: I119). Sin alergias medicamentosas conocidas.'}])

body('¹ NLR: razón neutrófilo-linfocito (neutrophil-to-lymphocyte ratio). ² PLR: razón plaqueta-linfocito (platelet-to-lymphocyte ratio).', italic=True, color=GRAY, before=0, after=4)
h2('2.3 Examen físico al ingreso HEEE')
table('Signos vitales durante el traslado al HEEE (16/01/2026)',
    ['Parámetro','Valor inicial','Valor final','Referencia'],
    [['FC (lpm)','~85','~104','60–100'],
     ['Temperatura (°C)','39,2','38,8','36,5–37,5'],
     ['TA (mmHg)','~152/89','~105/75','< 130/80'],
     ['FR (rpm)','20','20','12–20'],
     ['SpO2 (%)','99','94','≥ 95'],
     ['Glasgow','15/15','15/15','15/15']],
    [5,3.5,3.5,4])
body('Hallazgos clínicos relevantes al ingreso:')
for item in [
    'Necrosis tisular extensa en región cervical anterior y lateral con escara negra y áreas de tejido fibrinoide amarillento.',
    'Fístulas cutáneas activas con drenaje de líquido purulento fétido. Edema indurado submandibular-supraclavicular bilateral.',
    'Trismus funcional. Voz empastada. Sin estridor audible al ingreso. SpO2 94% durante el traslado.',
    'Taquicardia (FC ~104 lpm). Temperatura 38,8–39,2°C. Edema grado 2+ en miembros inferiores.',
]: bullet(item)

h2('2.4 Exámenes de laboratorio preingreso (Hospital Domínguez Zambrano)')
table('Evolución de laboratorios preingreso — Hospital Dr. Gustavo Domínguez Zambrano (15–16/01/2026)',
    ['Parámetro','15/01 15:00 h','16/01 01:41 h','16/01 10:50 h','Referencia'],
    [['Leucocitos x10³/µL','13,17 ↑','10,42','19,55 ↑↑','4,0–10,0'],
     ['Neutrófilos (%)','83,3','75,4','79,4','50–70'],
     ['NLR¹','9,3 ↑','5,6','7,5 ↑','< 3,0'],
     ['PLR²','628 ↑↑','479 ↑↑','648 ↑↑','< 150'],
     ['Hemoglobina (g/dL)','9,7 ↓','8,2 ↓','8,9 ↓','12,3–15,3'],
     ['Plaquetas x10³/µL','741 ↑↑','675 ↑↑','727 ↑↑','149–409'],
     ['Glucosa (mg/dL)','299,7 ↑↑','—','164,9 ↑','60–100'],
     ['Sodio (mEq/L)','126,9 ↓↓','—','130,6 ↓','136–145'],
     ['Potasio (mEq/L)','3,3 ↓','—','3,1 ↓','3,5–5,1'],
     ['pH (gasometría)','7,581 ↑','—','—','7,35–7,45'],
     ['Lactato (mmol/L)','1,2','—','—','0,6–1,7'],
     ['HBsAg / Anti-VHC','NEG / NEG','—','—','Negativos']],
    [4.5,2.8,2.8,2.8,3.1])

# ── ESTUDIOS DE IMAGEN ES — CON FIGURAS ───────────────────────────────────────
doc.add_page_break()
h1('3. ESTUDIOS DE IMAGEN')

h2('3.1 Plano sagital — TAC simple de cuello y tórax')
mixed([{'text':'La TAC simple de cuello y tórax (sin contraste) del 17/01/2026 reportó gas en región cervical bilateral y anterior, con colección en región pectoral izquierda. Los cortes sagitales (Figuras 1–3) demuestran gas en tejidos blandos cervicales, colección purulenta retrofaríngea y extensión infero-superior del proceso necrotizante desde C3 hasta la región supraclavicular.'},
       {'text':'11,23','sup':True}])
fig_row([
    (F+'tac1_sagital_gas.jpg', 5.5,
     'TAC sagital (16/01/2026): gas retrofaríngeo (burbujas hipodensas) y colección purulenta cervical. '
     'Proceso necrotizante con gas difuso en tejidos blandos, signo patognomónico de FNC.'),
    (F+'tac2_sagital_coleccion.jpg', 5.5,
     'TAC sagital (16/01/2026): colección purulenta con múltiples burbujas de gas en espacios cervicales '
     'profundos. Extensión desde el espacio retrofaríngeo hacia el mediastino superior.'),
    (F+'tac3_sagital_extension.jpg', 5.5,
     'TAC sagital (16/01/2026): extensión inferosuperior del proceso necrotizante cervical bilateral '
     'desde C3 hasta la región supraclavicular con compromiso del espacio peligroso.'),
])

h2('3.2 Plano coronal')
mixed([{'text':'Los cortes coronales (Figuras 4–6) muestran el absceso necrotizante cervical bilateral, gas difuso en los espacios fasciales profundos del cuello y la extensión hacia el mediastino superior.'},
       {'text':'11,12','sup':True}])
fig_row([
    (F+'tac4_coronal_absceso.jpg', 5.5,
     'TAC coronal (16/01/2026): absceso necrotizante cervical bilateral. Ocupación de los espacios '
     'fasciales profundos con gas (hipodensidades puntiformes) en ambos lados del cuello.'),
    (F+'tac5_coronal_gas.jpg', 5.5,
     'TAC coronal (16/01/2026): gas retrofaríngeo bilateral y engrosamiento difuso de tejidos blandos '
     'cervicales. Sin desplazamiento laríngeo ni extensión mediastínica al momento del estudio.'),
    (F+'tac6_coronal_espacios.jpg', 5.5,
     'TAC coronal (16/01/2026): compromiso bilateral de espacios fasciales profundos cervicales con gas '
     'y proceso inflamatorio difuso. Nótese la asimetría del proceso con mayor afectación derecha.'),
])


h2('3.3 Plano axial')
mixed([{'text':'Los cortes axiales (Figuras 7 y 8) muestran la afectación bilateral de los espacios parafaríngeos '
        'y submandibulares con hipodensidades compatible con gas extrafascial — hallazgo patognomónico de FNC.'},
       {'text':'10,11,23','sup':True}])
fig_row([
    (F+'tac7_axial_parafaringeo.jpg', 8,
     'TAC axial (16/01/2026): corte a nivel suprahioideo. Hipodensidades bilaterales en espacios parafaríngeos '
     'y masticadores compatibles con gas extrafascial, confirmando la naturaleza bilateral de la FNC.'),
    (F+'tac8_axial_submandibular.jpg', 8,
     'TAC axial (16/01/2026): segundo corte a nivel suprahioideo. Afectación bilateral simétrica de los espacios '
     'fasciales profundos. La distribución bilateral confirma el diagnóstico de Angina de Ludwig complicada con FNC.'),
])

h2('3.4 Hallazgos imagenológicos clave')
body('Reporte TAC simple de cabeza, cuello y tórax (17/01/2026): "Gas en región cervical derecha, izquierda y anterior. No aparenta mediastinitis. Gas y colección en región pectoral izquierda. Sin invasión ni desplazamiento de vía aérea."', italic=True)
body('Los hallazgos tomográficos que fundamentaron la indicación de desbridamiento urgente:')
for item in [
    'Gas en tejidos blandos cervicales en múltiples planos (sagital y coronal) — hallazgo patognomónico de FNC.',
    'Colección purulenta retrofaríngea con extensión de C3 a región supraclavicular.',
    'Compromiso bilateral de espacios fasciales profundos sin extensión mediastínica al momento del estudio.',
    'Ausencia de desplazamiento laríngeo — vía aérea no comprometida directamente por la colección.',
]: bullet(item)

h2('3.5 Imagen clínica al ingreso HEEE (17/01/2026)')
body('La Figura 9 documenta la extensión de la necrosis cervical, la escara oscura característica de la '
     'FNC y las fístulas activas con drenaje purulento al momento de la evaluación en el HEEE.')
fig(F+'fig2_preop_necrosis_jan17.jpg', 10,
    'Fotografía clínica al ingreso HEEE (17/01/2026). Necrosis tisular extensa de región cervical anterior '
    'y lateral con escara oscura, tejido fibrinoide y fístulas cutáneas activas con drenaje purulento fétido. '
    'Paciente con cánula nasal de O₂ y drenaje cervical colocado. Edema submandibular bilateral marcado.')

# ── EVALUACIÓN MULTIDISCIPLINARIA ES ─────────────────────────────────────────
doc.add_page_break()
h1('4. EVALUACIÓN MULTIDISCIPLINARIA')
h2('4.1 Proceso de referencia y recepción')
body('La paciente fue referida por la Dra. Patricia Verónica Morales Cabezas (Especialista en Medicina de '
     'Emergencias, SENESCYT: 1027-2017-1891207) desde el Hospital Dr. Gustavo Domínguez Zambrano al HEEE. '
     'Traslado en ambulancia SVA terrestre el 16/01/2026 a las 07:00 h. Recibida por la Dra. Mariuxi Báez Moreira.')
h2('4.2 Equipo ORL — HEEE')
body('El Servicio de Otorrinolaringología del HEEE evaluó a la paciente al ingreso e identificó Angina de '
     'Ludwig complicada con FNC bilateral como emergencia quirúrgica de primer orden, con indicación '
     'de desbridamiento radical urgente.')
h2('4.3 Anestesiología')
mixed([{'text':'El edema cervical masivo y el trismus funcional requirieron protocolo de vía aérea difícil prevista. '
        'La intubación orotraqueal se realizó exitosamente bajo anestesia general (confirmada por cultivo de '
        'secreción traqueal del 18/01/2026). Clasificación ASA: III.'},
       {'text':'14','sup':True}])
h2('4.4 UCI')
body('Ingreso a UCI (Piso 1, HEEE) el 17/01/2026 a las 23:48 h en el posquirúrgico inmediato (Dra. Soraya Acaro) para monitorización '
     'continua, soporte metabólico e insulinoterapia IV (meta glucémica perioperatoria: 140–180 mg/dL).')
h2('4.5 Infectología')
mixed([{'text':'Ajuste a las 48–72 h: meropenem instaurado como antibiótico definitivo al confirmar '},
       {'text':'K. pneumoniae','italic':True},{'text':' BLEE y '},{'text':'E. coli','italic':True},
       {'text':' BLEE con CIM meropenem ≤0,25 µg/mL — única opción activa frente a ambos aislados.'},
       {'text':'8,9','sup':True}])
h2('4.6 Cirugía Plástica')
body('El 16/02/2026 (~POD 29), con parámetros inflamatorios normalizados (WBC 5,97 x10³/µL, Na 138 mEq/L), '
     'traslado al Servicio de Cirugía Plástica (Dr. Juan Gutiérrez, Piso 3, HEEE) para manejo '
     'reconstructivo del defecto cervical residual.')

# ── MANEJO QUIRÚRGICO ES ──────────────────────────────────────────────────────
doc.add_page_break()
h1('5. MANEJO QUIRÚRGICO')
h2('5.1 Preparación preoperatoria')
for item in [
    'Resucitación con cristaloides IV. Corrección de hiponatremia (Na 126,9 mEq/L) e hipopotasemia (K 2,9 mEq/L).',
    'Antibioticoterapia empírica IV (beta-lactámico amplio espectro + metronidazol) hasta disponibilidad de cultivos.',
    'Insulinoterapia IV (meta 140–180 mg/dL). Solicitud de 2 unidades de paquete globular (Hb 8,2–8,9 g/dL).',
    'Consentimiento informado quirúrgico y anestésico. Clasificación ASA: III.',
]: bullet(item)

h2('5.2 Técnica quirúrgica — Desbridamiento cervical radical')
mixed([{'text':'Fecha: ','bold':True},{'text':'17 de enero de 2026.'}])
mixed([{'text':'Cirujano principal: ','bold':True},
       {'text':'Dr. Christian Alfonso Freire Zamora, MD, Especialista en ORL — Jefe del Servicio ORL, HEEE.'}])
mixed([{'text':'Abordaje: ','bold':True},
       {'text':'Desbridamiento cervical radical bilateral bajo anestesia general con intubación orotraqueal. '
        'La naturaleza necrotizante del proceso exigió exéresis radical de todos los tejidos desvitalizados.'},
       {'text':'9,10','sup':True}])
body('Descripción del procedimiento:')
for item in [
    'Anestesia general con intubación orotraqueal (IOT) bajo protocolo de vía aérea difícil prevista.',
    'Delimitación de márgenes de necrosis. Extensión bilateral en región cervical anterior, lateral y submandibular.',
    'Incisiones cervicales con exéresis de escara necrótica y tejido fibrinoide desvitalizado.',
    'Desbridamiento roma y cortante hasta tejido sano sangrante.',
    'Evacuación de material purulento fétido. Envío de muestras para cultivos aerobio, anaerobio, hongos y antibiograma.',
    'Lavado exhaustivo con solución salina normal (SSN) 0,9% tibia hasta irrigación limpia.',
    'Colocación de drenajes Penrose por contraincisiones. Herida abierta — sin cierre primario.',
]: numbered(item)

mixed([{'text':'Datos intraoperatorios: ','bold':True},
       {'text':'Duración del procedimiento: 3 horas 10 minutos. Sangrado estimado: 200 mL aproximadamente. '
        'Volumen de líquido purulento evacuado: 350 mL aproximadamente. Técnica de intubación: '
        'no documentada en el protocolo anestésico.'}])
body('Las Figuras 10 y 11 documentan el campo quirúrgico intraoperatorio.')
fig_row([
    (F+'fig3a_intraop_desbrid_jan18.jpg', 8,
     'Campo quirúrgico intraoperatorio (18/01/2026): desbridamiento cervical bilateral con retractores. '
     'Se evidencia el tejido necrótico y fibrinoide en exéresis, con sangrado activo de tejido vital. '
     'Cirujano: Dr. Christian Alfonso Freire Zamora. Servicio de ORL, HEEE.'),
    (F+'fig3b_intraop_diseccion_jan18.jpg', 8,
     'Disección intraoperatoria (18/01/2026): exposición de los planos fasciales profundos cervicales. '
     'Se aprecian las estructuras del triángulo anterior del cuello tras el desbridamiento radical.'),
])

# ── EVOLUCIÓN POSOPERATORIA ES ────────────────────────────────────────────────
doc.add_page_break()
h2('5.3 Procedimientos quirúrgicos subsecuentes / Subsequent surgical procedures')
body('Durante la hospitalización se realizaron tres procedimientos quirúrgicos adicionales bajo anestesia general, '
     'documentados en la historia clínica del HEEE (HC 0802060905):')
table('Procedimientos quirúrgicos durante la hospitalización — HEEE, enero 2026',
    ['Fecha / POD','Procedimiento','Hallazgos quirúrgicos','Servicio / Cirujano'],
    [['17/01/2026 (POD 0)\n20:20–22:20 h',
      'Cervicotomía amplia + drenaje absceso profundo de cuello + limpieza profunda + debridación tejido necrótico región cervical anterior + traqueotomía',
      '• Abundantes gleras en región cervical (~80 g)\n• FNC en región cervical anterior\n• Líquido purulento verdoso fétido ~150 mL\n• Colección pectoral izquierdo ~40 mL\n• Tráquea lateralizada a la izquierda\n• Tiroides desplazada hacia inferior\n• Sangrado: ~200 mL. Líquido purulento + gleras: ~350 mL',
      'ORL / Dr. Christian Alfonso Freire Zamora\nAyudantes: MD Cárdenas, MD León\nAnestesiología: Dr. Vásquez'],
     ['21/01/2026 (POD 4)',
      'Limpieza profunda de cuello + debridación de gleras en región cervical anterior + toma de cultivo de tejido',
      '• Escasas gleras a nivel cervical anterior\n• Sin tejido necrótico evidente\n• Secreción seropurulenta en escasa cantidad\n• Traqueotomo in situ permeable\n• Absceso pectoral izquierdo con edema y eritema',
      'Cirugía Plástica'],
     ['27/01/2026 (POD 10)',
      'Limpieza quirúrgica + debridamiento limitado + plastia parcial de heridas en cuello y tórax anterior',
      '• Herida cervical anterior sin tejido necrótico, escasas gleras, sin secreción purulenta\n• Traqueotomo in situ funcional\n• Ostomía sin signos de infección\n• Tórax anterior: área cruenta con tejido de granulación moderado, bolsillos de 10 cm hacia posterior\n• Tórax izquierdo: herida con exposición de planos musculares (pectoral mayor)',
      'Cirugía Plástica']],
    [3.5, 5, 6, 3])
space()

h1('6. EVOLUCIÓN POSOPERATORIA')

h2('6.1 Resultados microbiológicos definitivos')
table('Resultados microbiológicos definitivos — Laboratorio de Microbiología HEEE (enero–febrero 2026)',
    ['Muestra (fecha)','Microorganismo aislado','Sensible a','Resistente a'],
    [['Hisopado herida cuello der. (17/01/2026)',
      [{'text':'Klebsiella pneumoniae ss. pneumoniae BLEE','italic':True,'bold':True,'color':RED},
       {'text':' + '},{'text':'Candida tropicalis','italic':True},
       {'text':' (aislado concomitante, cultivo de hongos 17/01/2026)'}],
      'Amikacina (CIM ≤1), Gentamicina, Imipenem, Meropenem (CIM ≤0,25)',
      'Amp/Sulbactam (CIM ≥32), Cefepima, Ceftriaxona (CIM ≥64), Ciprofloxacino, Pip/Tazo, TMP/SMX'],
     ['Hisopado herida cuello der. — cultivo de hongos (19/01/2026)',
      [{'text':'Candida tropicalis','italic':True},
       {'text':' + '},{'text':'Candida parapsilosis','italic':True},
       {'text':' — ambas especies confirmadas. Sin antifungigrama disponible.'}],
      'Sin antifungigrama realizado','—'],
     ['Tejido necrótico cuello (18/01/2026)',
      [{'text':'Klebsiella pneumoniae BLEE','italic':True},{'text':' (2ª confirmación)'}],
      'Meropenem, Imipenem, Amikacina, Gentamicina','Mismo perfil BLEE'],
     ['Secreción cuello — 2ª muestra (18/01/2026)',
      [{'text':'E. coli BLEE + K. pneumoniae BLEE','italic':True},{'text':' (polimicrobiana)'}],
      'Amikacina (S), Gentamicina (S), Imipenem (CIM ≤1), Meropenem (CIM ≤1)',
      'Amp/Sulbactam, Cefepima, Ceftriaxona, Pip/Tazo, TMP/SMX'],
     ['Secreción traqueal (18/01/2026)','Sin desarrollo bacteriano','—','—'],
     ['Hemocultivos ×4 sets (18/01/2026)','SIN DESARROLLO BACTERIANO — bacteriemia descartada','—','—'],
     ['Hisopado rectal vigilancia (17/01/2026)','NEGATIVO para resistentes a carbapenems (sin KPC/MBL)','—','—'],
     ['Urocultivo por sonda (17/01/2026)','Sin desarrollo bacteriano','—','—'],
     ['Orina — candiduria (04/02/2026)',
      [{'text':'Candida spp.','italic':True},{'text':' — Levaduras +++ / Hifas +++ (candiduria nosocomial, POD 17). Especie no identificada.'}],
      'Antifungigrama no realizado','—']],
    [4.5,5,4.5,4])

h2('6.2 Antibioticoterapia definitiva')
table('Esquema antibiótico durante la hospitalización',
    ['Período','Fármaco','Justificación','Duración'],
    [['Empírico (desde 17/01/2026)',
      'Meropenem 1 g c/8 h IV + Vancomicina 1 g c/12 h IV (desde 17/01/2026)',
      'Cobertura empírica: Meropenem vs gramnegativos BLEE; Vancomicina vs cocos grampositivos. Instaurado desde ORL urgencias 17/01/2026.',
      'Meropenem: continuó indefinido (definitivo). Vancomicina: 14 dosis (~hasta 24/01/2026)'],
     ['Definitivo (desde 19/01)',
      [{'text':'Meropenem 1 g c/8 h IV','bold':True},
       {'text':' — único activo vs K. pneumoniae BLEE y E. coli BLEE'}],
      'CIM ≤0,25 µg/mL para ambos aislados. Resistencia completa a cefalosporinas y beta-lactámicos.',
      '14–21 días IV'],
     ['Antifúngico (desde 20/01/2026 — POD 3)',
      'Fluconazol 150 mg IV c/24 h',
      'Candiduria nosocomial (POD 17) + Candida en herida/herida cervical. Fluconazol como antifúngico azólico de primera línea.','14 días']],
    [3,5.5,5,4])

h2('6.3 Evolución de laboratorios durante hospitalización HEEE')
table('Evolución de parámetros de laboratorio — HEEE (enero–febrero 2026)',
    ['Fecha / Servicio','WBC ×10³','Hb g/dL','PCR¹ mg/L','PCT ng/mL','Na mEq/L','K mEq/L','Glucosa mg/dL'],
    [['17/01 — ORL Urgencias','13,11 ↑','8,3 ↓','202 ↑↑↑','0,31','133 ↓','2,9 ↓↓','120 ↑'],
     ['18/01 — UCI preop','13,69 ↑','8,9 ↓','189 ↑↑↑','0,26','130 ↓↓','3,3 ↓','164 ↑↑'],
     ['19/01 — UCI posop','7,06','8,2 ↓','—','—','132 ↓','3,8','93'],
     ['20/01 — ORL Piso 6','8,31','12,0 ↓','114 ↑↑','0,12','129 ↓↓','4,1','—'],
     ['22/01 — ORL Piso 6','5,35','10,4 ↓','—','0,08','134 ↓','4,1','129 ↑'],
     ['26/01 — ORL Piso 6','7,41','12,0 ↓','16,7 ↑','—','135','3,6','114 ↑'],
     ['05/02 — ORL Piso 6','7,67','11,1 ↓','8,85 ↑','—','135','4,9','135 ↑'],
     ['09/02 — ORL Piso 6','9,27','12,0 ↓','7,35 ↑','—','134 ↓','5,4 ↑','139 ↑'],
     ['16/02 — Cir. Plástica','5,97','10,9 ↓','—','—','138','4,2','131 ↑']],
    [4,2,2,2,2,2,2,2.5])
body('¹ PCR: proteína C reactiva. PCT: procalcitonina.', italic=True, color=GRAY, before=0, after=4)
mixed([{'text':'La proteína C reactiva (PCR) descendió de 202 mg/L al ingreso hasta 7,35 mg/L al día 22 posoperatorio, '
        'reflejando la respuesta favorable al desbridamiento y a meropenem.'},
       {'text':'12','sup':True},
       {'text':" La procalcitonina se mantuvo < 0,5 ng/mL en todos los puntos, concordante con los hemocultivos negativos."},
       {'text':'24,15','sup':True}])

h2('6.4 Evolución fotográfica de la herida')
body('Las Figuras 12 y 13 muestran la evolución posquirúrgica temprana (día posoperatorio [POD] 6, 24/01/2026). '
     'Las Figuras 15–17 documentan el control ambulatorio (~POD 71, 30–31/03/2026) con '
     'granulación activa en los defectos cervical e infraclavicular.')
fig_row([
    (F+'fig4b_postop_herida_jan24.jpg', 8,
     'Evolución posquirúrgica (POD 6, 24/01/2026): vista general de la herida cervical con drenaje activo (azul) '
     'y cura abierta bilateral. Reducción evidente del edema perilesional respecto al ingreso (17/01/2026).'),
    (F+'fig4a_postop_traqueo_jan24.jpg', 8,
     'Evolución posquirúrgica (POD 6, 24/01/2026): acercamiento de la traqueostomía cervical con cánula in situ, '
     'suturas de afrontamiento y tejido de granulación incipiente en los bordes de la herida.'),
])
fig_row([
    (F+'fig5a_control_cervical_mar30.jpg', 5.5,
     'Control ambulatorio (~POD 71, 30/03/2026): defecto cervical submandibular con tejido de '
     'granulación activo y defecto infraclavicular derecho en proceso de cierre por segunda intención.'),
    (F+'fig5b_control_infraclavicular_mar30.jpg', 5.5,
     'Control ambulatorio (30/03/2026): acercamiento del defecto infraclavicular con tejido de '
     'granulación activo, bordes bien definidos sin signos de infección activa.'),
    (F+'fig5c_control_panoramico_mar31.jpg', 5.5,
     'Control ambulatorio (31/03/2026): vista panorámica cervical con evolución favorable del '
     'defecto. Granulación activa en resolución. Cicatrización por segunda intención en curso.'),
])

h2('6.5 Complicaciones posoperatorias')
mixed_bullet([{'text':'Hipoxemia aguda (POD 5, 23/01): ','bold':True},
              {'text':'SatO2 85,7%, PaO2 47,9 mmHg, Dímero D 1,27 µg/mL. Compatible con atelectasia o neumonía aspirativa.'}])
mixed_bullet([{'text':'Hiponatremia e hipopotasemia recurrentes: ','bold':True},
              {'text':'Na mín. 129 mEq/L, K mín. 2,9 mEq/L. Corrección progresiva con reposición IV.'}])
mixed_bullet([{'text':'Candiduria nosocomial (POD 17, 04/02): ','bold':True},
              {'text':'Candida spp. levaduras +++ / hifas +++. Tratada con fluconazol 150 mg IV c/24 h por 14 días con resolución favorable.'}])
mixed_bullet([{'text':'Eosinofilia progresiva: ','bold':True},
              {'text':'Hasta 18,7% (09/02). Compatible con reacción a antibioticoterapia prolongada.'}])
mixed_bullet([{'text':'Anemia de enfermedad crónica: ','bold':True},
              {'text':'Hb 8,2–12,0 g/dL durante toda la hospitalización.'}])

# ── DISCUSIÓN ES ──────────────────────────────────────────────────────────────
doc.add_page_break()
h1('7. DISCUSIÓN')

h2('7.1 Diagnóstico diferencial')
table('Diagnóstico diferencial de las infecciones cervicofaciales profundas',
    ['Entidad','Características clínicas','Hallazgos imagenológicos','Manejo'],
    [['Flemón cervical',
      'Induración difusa, sin fluctuación ni necrosis.',
      'Engrosamiento difuso. Sin gas ni colección.',
      [{'text':'ATB IV + vigilancia estrecha.'},{'text':'12,13','sup':True}]],
     ['Absceso profundo de cuello',
      'Fluctuación, trismus, fiebre moderada.',
      'Colección hipodensa con rim enhancement. Gas posible.',
      [{'text':'Drenaje quirúrgico + ATB IV.'},{'text':'12,13','sup':True}]],
     ['Angina de Ludwig (inicio)',
      [{'text':'CASO (inicio): ','bold':True},
       {'text':'Celulitis bilateral suelo boca, trismus funcional, sin fluctuación.'}],
      'Engrosamiento bilateral submandibular/sublingual. Sin colección definida.',
      [{'text':'Manejo vía aérea + ATB + desbridamiento si progresión.'},{'text':'1,2','sup':True}]],
     ['FNC cervical',
      [{'text':'CASO (progresión): ','bold':True},
       {'text':'Necrosis extensa, escara negra, fístulas con pus fétido, crepitación, sepsis sistémica.'}],
      'Gas tejidos blandos, necrosis fascial, fístulas confirmadas (ver Figuras 1–9).',
      [{'text':'Desbridamiento radical + carbapenems + UCI.'},{'text':'9,10','sup':True}]]],
    [3.5,4.5,4.5,4])
space()

h2('7.2 Progresión de Angina de Ludwig a FNC en paciente diabética')
mixed([{'text':'La DM genera disfunción de neutrófilos, microangiopatía y neuropatía periférica que favorecen '
        'la progresión silente de la Angina de Ludwig a FNC.'},
       {'text':'1,2,8,21','sup':True},
       {'text':" La hiperglucemia grave al ingreso (glucosa 299–331 mg/dL) y la evolución de un mes sin atención "
        "determinaron la progresión irreversible a necrosis fascial, evidenciada en los tres planos tomográficos "
        "(Figuras 1–8) y en las fotografías clínicas de ingreso (Figura 9)."},
       {'text':'8','sup':True}])

mixed([{'text':'En consonancia con Huang et al., los pacientes diabéticos con infecciones '
        'profundas de cuello presentan estancias más prolongadas y mayor tasa de complicaciones que los no diabéticos.'},
       {'text':'8','sup':True},
       {'text':' Mariano et al. destacan que la progresión a FNC en el paciente diabético constituye un factor '
        'de mal pronóstico y que el desbridamiento precoz es el único elemento modificable que impacta en la supervivencia.'},
       {'text':'21','sup':True},
       {'text':' En el presente caso, el retraso de un mes en la atención transformó una Angina de Ludwig '
        'potencialmente manejable en una FNC establecida que requirió UCI y reconstrucción quirúrgica.'}])
h2('7.3 Klebsiella pneumoniae BLEE en infección de inicio comunitario')
mixed([{'text':'El aislamiento de '},{'text':'K. pneumoniae','italic':True},
       {'text':' BLEE como agente principal en FNC de inicio aparentemente comunitario (extracción dental '
        'ambulatoria) es infrecuente y de implicaciones críticas: exige carbapenems como único tratamiento '
        'activo, descartando todos los beta-lactámicos y cefalosporinas.'},
       {'text':'7,8','sup':True},
       {'text':" La elección de meropenem (CIM ≤0,25 µg/mL) fue óptima. La negatividad de los 4 sets de "
        "hemocultivos y el descenso de la PCR (proteína C reactiva) confirman que el proceso fue controlado sin bacteriemia."},
       {'text':'15','sup':True}])

mixed([{'text':'Marioni et al. reportaron que el 74% de las infecciones profundas de cuello de '
        'origen dental fueron causadas por flora mixta polimicrobiana (estreptococos y anaerobios).'},
       {'text':'25','sup':True},
       {'text':' El aislamiento de '},{'text':'K. pneumoniae','italic':True},
       {'text':' BLEE como microorganismo dominante es por tanto un hallazgo inusual en infecciones de inicio '
        'comunitario, que sugiere la posibilidad de colonización previa o exposición nosocomial no identificada.'}])
h2('7.4 Abordajes quirúrgicos según espacio fascial comprometido')
table('Abordajes quirúrgicos según espacio fascial cervical comprometido',
    ['Espacio fascial','Abordaje de elección','Abordaje alternativo','Ref.'],
    [['Submandibular (Angina de Ludwig)',
      'Bilateral submandibular + submentoniano',
      'Traqueostomía profiláctica',
      [{'text':'1,2','sup':True}]],
     ['FNC cervical extensa (CASO ACTUAL)',
      'Desbridamiento radical bilateral (Figuras 10–11)',
      'Múltiples incisiones de relajación ± VAC therapy',
      [{'text':'9,10','sup':True}]],
     ['Parafaríngeo','Transcervical submandibular','Transoral seleccionado',[{'text':'13','sup':True}]],
     ['Retrofaríngeo','Transcervical + transoral','Transoral (pediatría)',[{'text':'16','sup':True}]],
     ['Carotídeo','Transcervical anterior al ECM','—',[{'text':'17','sup':True}]],
     ['Espacio peligroso / Mediastino','Transcervical extendido + VATS','Toracotomía derecha',[{'text':'18','sup':True}]]],
    [4,5,5,2])
space()

h2('7.5 Complicaciones')
mixed_bullet([{'text':'Mediastinitis descendente necrotizante (MDN): ','bold':True},
              {'text':'Mortalidad 11,8–40%. La TAC descartó extensión mediastínica al momento del estudio.'},
              {'text':'4','sup':True}])
mixed_bullet([{'text':'Bacteriemia / Sepsis: ','bold':True},
              {'text':'4 sets de hemocultivos negativos. Respuesta favorable al control del foco y meropenem.'},
              {'text':'15','sup':True}])
mixed_bullet([{'text':'Síndrome de Lemierre: ','bold':True},
              {'text':'No documentado (hemocultivos negativos). Dímero D elevado (1,27 µg/mL) ameritó vigilancia.'},
              {'text':'19','sup':True}])
mixed_bullet([{'text':'Candidosis nosocomial: ','bold':True},
              {'text':'Candiduria POD 17 en DM2 + antibioticoterapia prolongada + catéter vesical. Requirió antifúngico.'},
              {'text':'20','sup':True}])
mixed_bullet([{'text':'Reconstrucción cervical: ','bold':True},
              {'text':'Defecto cervical manejado por Cirugía Plástica (~POD 29), documentado en Figuras 12–17.'},
              {'text':'9','sup':True}])


h2('7.6 Limitaciones del caso')
mixed_bullet([{'text':'Diseño retrospectivo: ','bold':True},
              {'text':'Datos obtenidos mediante revisión de historia clínica tras el alta. Riesgo de datos incompletos.'}])
mixed_bullet([{'text':'Técnica de intubación no documentada: ','bold':True},
              {'text':'El protocolo anestésico no especifica el método utilizado (videolaringoscopía, fibroscopía u otro) '
               'para el manejo de la vía aérea difícil prevista.'}])
mixed_bullet([{'text':'Evaluación pulmonar incompleta: ','bold':True},
              {'text':'El episodio de hipoxemia grave del POD 5 (SatO₂ 85,7%, Dímero D 1,27 µg/mL) no cuenta con '
               'registro de angioTAC pulmonar, limitando la exclusión formal de tromboembolia pulmonar.'}])
mixed_bullet([{'text':'Caso único (n = 1): ','bold':True},
              {'text':'Los hallazgos no son generalizables. Se requieren series multicentricas para validar las conclusiones.'}])
mixed_bullet([{'text':'Seguimiento no protocolizado: ','bold':True},
              {'text':'Aunque la paciente continúa en control ambulatorio activo, no existe protocolo formal de seguimiento '
               'a largo plazo, limitando la evaluación del resultado funcional y estético definitivo.'}])
mixed_bullet([{'text':'Ausencia de cultivos anaeróbicos confirmatorios: ','bold':True},
              {'text':'Los cultivos anaeróbicos no reportaron crecimiento, posiblemente por sensibilidad técnica o '
               'uso de antibioticoterapia previa, sin descartar su participación en la infección polimicrobiana.'}])

# ── CONCLUSIONES ES ───────────────────────────────────────────────────────────
doc.add_page_break()
h1('8. CONCLUSIONES')
for item in [
    'La Angina de Ludwig complicada con fascitis necrotizante cervical es una emergencia quirúrgica con mortalidad potencial del 15–40%. El diagnóstico precoz con TAC multiplanar (sagital, coronal y axial), el traslado urgente a tercer nivel y el desbridamiento radical inmediato son los pilares del tratamiento.',
    'La extracción dental sin antibioticoprofilaxis en paciente diabética con hiperglucemia grave (glucosa 299–331 mg/dL) fue el evento desencadenante. La prevención requiere control glucémico estricto previo a procedimientos odontológicos en grupos de riesgo.',
    'El aislamiento de Klebsiella pneumoniae BLEE como agente principal — inusual en infecciones de inicio comunitario — modifica sustancialmente la estrategia antibiótica: meropenem (CIM ≤0,25 µg/mL) fue la única opción definitiva activa frente a ambos aislados BLEE.',
    'Los hemocultivos negativos (4 sets), el descenso de PCR de 202 a 7,35 mg/L en 22 días y la ausencia de extensión mediastínica confirman la eficacia del manejo instaurado.',
    'El abordaje multidisciplinario ORL–UCI–Infectología–Cirugía Plástica, supervisado por el Servicio de ORL del HEEE, fue determinante en el resultado favorable con 33 días de hospitalización (16/01–18/02/2026) y evolución a cicatrización por segunda intención (documentada en Figuras 12–17).',
    'La evolución de un mes sin atención oportuna desde Santo Domingo refleja una brecha crítica en el sistema de salud ecuatoriano para la detección temprana de complicaciones odontogénicas en el primer nivel de atención en pacientes diabéticos.',
]: bullet(item)

# ── REFERENCIAS ES ────────────────────────────────────────────────────────────
doc.add_page_break()
h1('9. REFERENCIAS')
body('Estilo Vancouver. Numeración por orden de primera citación en el texto.', italic=True, color=GRAY)
space()
for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_spacing(p, before=3, after=3, line=276)
    rn = p.add_run(f'{i}. ')
    rn.bold = True; rn.font.name = 'Arial'; rn.font.size = Pt(11)
    rt = p.add_run(ref)
    rt.font.name = 'Arial'; rt.font.size = Pt(11)



# ── DECLARACIÓN DE CONTRIBUCIÓN DE AUTORES (CRediT) ES ───────────────────────
doc.add_page_break()
h1('DECLARACIÓN DE CONTRIBUCIÓN DE AUTORES')
body('De acuerdo con la taxonomía CRediT (Contributor Roles Taxonomy), las contribuciones de cada autor al presente reporte de caso son las siguientes:', before=4, after=6)
table('Contribuciones de autores según taxonomía CRediT',
    ['Rol / Role','J.P. Gavilanez Heras','C.A. Freire Zamora','H.J. Campoverde Sánchez','G. Saenz Ortega'],
    [['Conceptualización','✓','','',''],
     ['Curación de datos','✓','','✓',''],
     ['Análisis formal','✓','','',''],
     ['Investigación','✓','','',''],
     ['Metodología','✓','','',''],
     ['Supervisión','','✓','',''],
     ['Validación','','✓','','✓'],
     ['Visualización','✓','','',''],
     ['Software / Análisis digital','','','','✓'],
     ['Escritura — borrador original','✓','','',''],
     ['Escritura — revisión y edición','✓','✓','✓','']],
    [5.5, 3, 3, 3, 3])
body('Todos los autores han leído y aprobado la versión final del manuscrito para su publicación.', italic=True, before=6, after=4)

# ╔══════════════════════════════════════════════════════════════════════════════
# SEPARADOR ENTRE VERSIONES
# ╔══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_spacing(p, before=60, after=10)
r = p.add_run('━' * 55)
r.font.name = 'Arial'; r.font.size = Pt(14); r.font.color.rgb = MED_BLUE

p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_spacing(p2, before=10, after=10)
r2 = p2.add_run('ENGLISH VERSION')
r2.bold = True; r2.font.name = 'Arial'; r2.font.size = Pt(20); r2.font.color.rgb = DARK_BLUE

p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_spacing(p3, before=4, after=10)
r3 = p3.add_run('Complete text in English  /  Texto completo en inglés')
r3.italic = True; r3.font.name = 'Arial'; r3.font.size = Pt(12); r3.font.color.rgb = GRAY

p4 = doc.add_paragraph(); p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_spacing(p4, before=4, after=60)
r4 = p4.add_run('━' * 55)
r4.font.name = 'Arial'; r4.font.size = Pt(14); r4.font.color.rgb = MED_BLUE

# Reset figure and table counters for English version
_fig[0] = 0
_tbl[0] = 0


# ╔══════════════════════════════════════════════════════════════════════════════
# ██████████████████   ENGLISH VERSION   ███████████████████████████████████████
# ╔══════════════════════════════════════════════════════════════════════════════

doc.add_page_break()
# ── COVER EN ──────────────────────────────────────────────────────────────────
cline('MINISTRY OF PUBLIC HEALTH OF ECUADOR', bold=True, size=11, color=MED_BLUE, before=0, after=3)
cline('Eugenio Espejo Specialty Hospital | Otorhinolaryngology Department | Quito, Ecuador',
      size=10, color=GRAY, before=0, after=6)
space()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_spacing(p, before=8, after=6)
r = p.add_run("LUDWIG'S ANGINA COMPLICATED WITH CERVICAL NECROTIZING FASCIITIS IN A DIABETIC PATIENT: "
              "A CASE REPORT WITH ISOLATION OF ESBL-PRODUCING ")
r.bold = True; r.font.name = 'Arial'; r.font.size = Pt(14); r.font.color.rgb = DARK_BLUE
r2 = p.add_run('Klebsiella pneumoniae'); r2.bold = True; r2.italic = True
r2.font.name = 'Arial'; r2.font.size = Pt(14); r2.font.color.rgb = DARK_BLUE
space()

for a in ['Jean Pierre Gavilanez Heras, MD, MSc¹',
          'Christian Alfonso Freire Zamora, MD, ORL Specialist¹*',
          'Hernán Javier Campoverde Sánchez, MD, MSc¹',
          'Gabriel Saenz Ortega, MD²']:
    cline(a, size=11, before=1, after=1)
space()
cline('¹ Department of Otorhinolaryngology, Eugenio Espejo Specialty Hospital (HEEE), Ministry of Public Health of Ecuador. Quito, Ecuador.',
      italic=True, size=9, color=GRAY, before=0, after=1)
cline('² Independent Physician / Médico independiente. Quito, Ecuador.',
      italic=True, size=9, color=GRAY, before=0, after=1)
cline('* Dr. Christian Alfonso Freire Zamora — ORL Specialist, Head of ORL Department, HEEE. Case Supervisor.',
      italic=True, size=9, color=GRAY, before=0, after=2)
cline('ORCIDs: J.P. Gavilanez Heras: 0009-0008-6518-6041 | H.J. Campoverde Sánchez: 0009-0000-4802-0105 | G. Saenz Ortega: 0009-0009-5669-1027 | C.A. Freire Zamora: 0000-0003-2237-1462',
      italic=True, size=9, color=GRAY, before=0, after=6)
space()
cline('Corresponding author:', bold=True, size=10, before=4, after=2)
cline('Jean Pierre Gavilanez Heras, MD, MSc', size=10, before=1, after=1)
cline('jean.gavilanez@hee.gob.ec  |  ORCID: 0009-0008-6518-6041', size=10, before=1, after=1)
cline('Eugenio Espejo Specialty Hospital, Quito, Ecuador', size=10, before=1, after=6)
cline('Conflicts of interest: None.  |  Funding: None.  |  Informed consent: Obtained.',
      italic=True, size=9, color=GRAY, before=4, after=2)

# ── ABSTRACT EN ───────────────────────────────────────────────────────────────
doc.add_page_break()
h1('ABSTRACT')
mixed([{'text':'Introduction: ','bold':True},
       {'text':"Ludwig's Angina is a bilateral indurated cellulitis of the floor of the mouth that, "
        "without timely treatment, can progress to cervical necrotizing fasciitis (CNF) with 15–40% mortality. "
        "Diabetes mellitus (DM) is the main risk factor for progression."},
       ])
mixed([{'text':'Case Presentation: ','bold':True},
       {'text':"A 49-year-old female with T2DM (11 years) and molar extraction (14/12/2025) developed Ludwig's "
        "Angina progressing to CNF over one month without treatment, with extensive cervical necrosis, cutaneous "
        "fistulas, and sepsis. Transferred January 16, 2026 to HEEE (Quito) with leukocytosis (19,550 x10³/µL), "
        "anemia (Hb 8.9 g/dL), severe hyponatremia (Na 126.9 mEq/L), and hyperglycemia (glucose 299.7 mg/dL). "
        "CT scan showed necrotizing cervical process with soft tissue gas without mediastinal involvement. "
        "Surgical debridement on January 17, 2026 under general anesthesia. Cultures isolated ESBL-producing "},
       {'text':'Klebsiella pneumoniae','italic':True},
       {'text':' and ESBL '},{'text':'Escherichia coli','italic':True},
       {'text':" (sensitive only to carbapenems) plus "},
       {'text':'Candida tropicalis','italic':True},{'text':' and '},
       {'text':'Candida parapsilosis','italic':True},
       {'text':'. Four blood culture sets were negative. Meropenem initiated as definitive therapy; '
        'C-reactive protein (CRP): 202 → 7.35 mg/L over 22 days. Total hospitalization ~30 days with Plastic Surgery reconstruction.'}])
mixed([{'text':'Conclusions: ','bold':True},
       {'text':"Ludwig's Angina complicated with CNF with ESBL-producing "},
       {'text':'K. pneumoniae','italic':True},
       {'text':' requires urgent radical debridement and meropenem. A multidisciplinary ORL–ICU–Infectious '
        'Disease–Plastic Surgery approach was key to the favorable outcome.'}])
mixed([{'text':'Keywords: ','bold':True,'italic':True},
       {'text':"Ludwig's Angina; cervical necrotizing fasciitis; odontogenic infection; ESBL Klebsiella pneumoniae; "
        "diabetes mellitus; surgical debridement; meropenem; Eugenio Espejo Hospital.",'italic':True}])

# ── INTRODUCTION EN ───────────────────────────────────────────────────────────
doc.add_page_break()
h1('1. INTRODUCTION')
h2("1.1 Ludwig's Angina: definition and anatomy")
mixed([{'text':"Ludwig's Angina, described by Wilhelm Friedrich von Ludwig in 1836, is defined as bilateral "
        "indurated cellulitis of the floor of the mouth simultaneously involving the submandibular, submental, "
        "and sublingual spaces, without initial abscess formation."},
       {'text':'1','sup':True},
       {'text':" Communication with the retropharyngeal space and the danger space, which descends to the "
        "posterior mediastinum, explains the rapid descending spread and the risk of descending necrotizing "
        "mediastinitis (DNM)."},
       {'text':'2,3,4','sup':True}])
h2('1.2 Epidemiology and risk factors')
mixed([{'text':'Odontogenic etiology accounts for 40–60% of adult cases, predominantly involving the lower second and third molars.'},
       {'text':'5,6,25','sup':True},
       {'text':" DM is the main risk factor for progression to CNF, present in 50–80% of cases, mediated by "
        "neutrophil dysfunction, tissue microangiopathy, and peripheral neuropathy that delay diagnosis."},
       {'text':'8,21','sup':True}])
h2('1.3 Progression to cervical necrotizing fasciitis')
mixed([{'text':'CNF is defined by progressive necrosis along cervical fascial planes, with mortality of '
        '15–40% even with adequate treatment.'},
       {'text':'4,9','sup':True},
       {'text':" The diagnostic triad includes: (1) cutaneous necrosis with dark eschar, "
        "(2) fetid purulent drainage through fistulas, and (3) fascial dissection confirmed surgically."},
       {'text':'10','sup':True}])

mixed([{'text':"The objective of this report is to describe the case of a patient with Ludwig's Angina complicated by cervical necrotizing fasciitis (CNF), with an unusual isolation of "},
       {'text':'Klebsiella pneumoniae','italic':True},
       {'text':' producing extended-spectrum beta-lactamases (ESBL) in a community-onset infection in the context of uncontrolled type 2 diabetes mellitus (T2DM), and to analyze the imaging, microbiological, surgical, and multidisciplinary aspects that determined the favorable outcome.'},
       {'text':'1,21','sup':True}])

# ── CASE PRESENTATION EN ──────────────────────────────────────────────────────
doc.add_page_break()
h1('2. CASE PRESENTATION')
body("Reported with informed patient consent per Ecuador's Organic Health Law. Identity protected with initials.")

h2('2.1 Patient demographics')
table('Patient demographic data',
    ['Parameter','Data'],
    [['Patient (initials)','G.M.P.Q.'],
     ['Sex / Age','Female, 49 years (07/05/1976)'],
     ['Origin','Santo Domingo de los Tsáchilas, Ecuador'],
     ['Referring hospital','Hospital Dr. Gustavo Domínguez Zambrano, Santo Domingo'],
     ['Date of admission (HEEE)','January 16, 2026'],
     ['Admitting service','Otorhinolaryngology — Floor 6, HEEE, Quito'],
     ['HEEE medical record','0802060905'],
     ['Case supervisor / Dept. head','Dr. Christian Alfonso Freire Zamora, MD, ORL Specialist — Head of ORL Dept., HEEE. Case Supervisor.']],
    [6,10])

h2('2.2 Chief complaint and medical history')
mixed([{'text':'~One month history (onset December 2025) of progressive mandibular edema extending to '
        'neck and chest, fever spikes, dysphagia, and cervical fistulas draining fetid purulent material. '},
       {'text':'Triggering event: molar extraction on December 14, 2025','bold':True},
       {'text':", outpatient setting, without documented antibiotic coverage."},
       {'text':'22','sup':True}])
mixed([{'text':'Past medical history: ','bold':True},
       {'text':'T2DM × 11 years (metformin 500 mg PO). Hypertensive heart disease without CHF (ICD-10: I119). No known drug allergies.'}])

body('¹ NLR: neutrophil-to-lymphocyte ratio. ² PLR: platelet-to-lymphocyte ratio.', italic=True, color=GRAY, before=0, after=4)
h2('2.3 Physical exam on HEEE admission')
table('Vital signs during transfer to HEEE (January 16, 2026)',
    ['Parameter','Initial value','Final value','Reference range'],
    [['HR (bpm)','~85','~104','60–100'],
     ['Temperature (°C)','39.2','38.8','36.5–37.5'],
     ['BP (mmHg)','~152/89','~105/75','< 130/80'],
     ['RR (rpm)','20','20','12–20'],
     ['SpO2 (%)','99','94','≥ 95'],
     ['Glasgow','15/15','15/15','15/15']],
    [5,3.5,3.5,4])
body('Relevant clinical findings on admission:')
for item in [
    'Extensive tissue necrosis of anterior and lateral cervical region with black eschar and yellowish fibrinoid areas.',
    'Active cutaneous fistulas with fetid purulent drainage. Bilateral submandibular-supraclavicular indurated edema.',
    'Functional trismus. Muffled voice. No audible stridor on admission. SpO2 94% during transfer.',
    'Tachycardia (HR ~104 bpm). Temperature 38.8–39.2°C. Grade 2+ lower limb edema.',
]: bullet(item)

h2('2.4 Pre-admission laboratory results')
table('Pre-admission laboratory evolution — Hospital Dr. Gustavo Domínguez Zambrano (January 15–16, 2026)',
    ['Parameter','15/01 15:00 h','16/01 01:41 h','16/01 10:50 h','Reference range'],
    [['WBC (×10³/µL)','13.17 ↑','10.42','19.55 ↑↑','4.0–10.0'],
     ['Neutrophils (%)','83.3','75.4','79.4','50–70'],
     ['NLR¹','9.3 ↑','5.6','7.5 ↑','< 3.0'],
     ['PLR²','628 ↑↑','479 ↑↑','648 ↑↑','< 150'],
     ['Hemoglobin (g/dL)','9.7 ↓','8.2 ↓','8.9 ↓','12.3–15.3'],
     ['Platelets (×10³/µL)','741 ↑↑','675 ↑↑','727 ↑↑','149–409'],
     ['Glucose (mg/dL)','299.7 ↑↑','—','164.9 ↑','60–100'],
     ['Sodium (mEq/L)','126.9 ↓↓','—','130.6 ↓','136–145'],
     ['Potassium (mEq/L)','3.3 ↓','—','3.1 ↓','3.5–5.1'],
     ['pH (blood gas)','7.581 ↑','—','—','7.35–7.45'],
     ['Lactate (mmol/L)','1.2','—','—','0.6–1.7'],
     ['HBsAg / Anti-HCV','NEG / NEG','—','—','Negative']],
    [4.5,2.8,2.8,2.8,3.1])

# ── IMAGING EN — WITH FIGURES ─────────────────────────────────────────────────
doc.add_page_break()
h1('3. IMAGING STUDIES')

h2('3.1 Sagittal plane — non-contrast CT of neck and chest')
mixed([{'text':'Non-contrast CT of the neck and chest (January 17, 2026) reported bilateral and anterior cervical gas with left pectoral collection. Sagittal CT sections (Figures 1–3) demonstrate soft tissue gas, retropharyngeal purulent '
        'collection, and inferosuperior extension of the necrotizing process from C3 to the '
        'supraclavicular region.'},
       {'text':'11,23','sup':True}])
fig_row([
    (F+'tac1_sagital_gas.jpg', 5.5,
     'Sagittal CT (16/01/2026): retropharyngeal gas (hypodense bubbles) and cervical purulent collection. '
     'Necrotizing process with diffuse soft tissue gas — pathognomonic sign of CNF.'),
    (F+'tac2_sagital_coleccion.jpg', 5.5,
     'Sagittal CT (16/01/2026): purulent collection with multiple gas bubbles in deep cervical spaces. '
     'Extension from the retropharyngeal space toward the superior mediastinum.'),
    (F+'tac3_sagital_extension.jpg', 5.5,
     'Sagittal CT (16/01/2026): inferosuperior extension of the bilateral cervical necrotizing process '
     'from C3 to the supraclavicular region with involvement of the danger space.'),
])

h2('3.2 Coronal plane')
mixed([{'text':'Coronal CT sections (Figures 4–6) show bilateral necrotizing cervical abscess, diffuse '
        'gas in the deep fascial spaces, and extension toward the superior mediastinum.'},
       {'text':'11,12','sup':True}])
fig_row([
    (F+'tac4_coronal_absceso.jpg', 5.5,
     'Coronal CT (16/01/2026): bilateral necrotizing cervical abscess with extension to deep fascial spaces. '
     'Note gas (punctiform hypodensities) on both sides of the neck.'),
    (F+'tac5_coronal_gas.jpg', 5.5,
     'Coronal CT (16/01/2026): bilateral retropharyngeal gas and diffuse cervical soft tissue thickening. '
     'No laryngeal displacement. No mediastinal extension at time of study.'),
    (F+'tac6_coronal_espacios.jpg', 5.5,
     'Coronal CT (16/01/2026): bilateral deep cervical fascial space involvement with gas and diffuse '
     'inflammatory process. Greater involvement noted on the right side.'),
])


h2('3.3 Axial plane')
mixed([{'text':'Axial CT sections (Figures 7 and 8) show bilateral involvement of the parapharyngeal and submandibular '
        'spaces with hypodensities consistent with extrafascial gas — a pathognomonic finding of CNF.'},
       {'text':'10,11,23','sup':True}])
fig_row([
    (F+'tac7_axial_parafaringeo.jpg', 8,
     'Axial CT (16/01/2026): suprahyoid level section. Bilateral hypodensities in parapharyngeal and masticator '
     'spaces consistent with extrafascial gas, confirming the bilateral nature of CNF.'),
    (F+'tac8_axial_submandibular.jpg', 8,
     'Axial CT (16/01/2026): second suprahyoid section. Symmetric bilateral involvement of deep cervical fascial '
     "spaces. Bilateral distribution confirms the diagnosis of Ludwig's Angina complicated with CNF."),
])

h2('3.4 Key imaging findings')
body('Official CT report (MSP/HCU form.053/2021, January 16, 2026): "Necrotizing abscess of the cervical '
     'region with cutaneous fistulas producing fetid purulent drainage. No laryngeal invasion or '
     'displacement observed."', italic=True)
body('Key CT findings supporting the indication for urgent debridement:')
for item in [
    'Soft tissue gas in multiple planes (sagittal and coronal) — pathognomonic sign of CNF.',
    'Retropharyngeal purulent collection extending from C3 to the supraclavicular region.',
    'Bilateral deep cervical fascial space involvement without mediastinal extension at time of study.',
    'No laryngeal displacement — airway not directly compromised by the collection.',
]: bullet(item)

h2('3.5 Clinical photograph on HEEE admission (January 17, 2026)')
body('Figure 9 documents the extent of cervical necrosis, the characteristic dark eschar of CNF, '
     'and the active fistulas with purulent drainage on evaluation at HEEE.')
fig(F+'fig2_preop_necrosis_jan17.jpg', 10,
    'Clinical photograph on HEEE admission (January 17, 2026). Extensive tissue necrosis of anterior and '
    'lateral cervical region with dark eschar, fibrinoid tissue areas, and active cutaneous fistulas with '
    'fetid purulent drainage. Patient with nasal O₂ cannula and cervical drain. Bilateral submandibular edema.')

# ── MULTIDISCIPLINARY EN ──────────────────────────────────────────────────────
doc.add_page_break()
h1('4. MULTIDISCIPLINARY ASSESSMENT')
h2('4.1 Referral and reception')
body('Referred by Dr. Patricia Verónica Morales Cabezas (Emergency Medicine Specialist, '
     'SENESCYT: 1027-2017-1891207) from Hospital Dr. Gustavo Domínguez Zambrano to HEEE. '
     'ALS ground transfer January 16, 2026 at 07:00 h. Received by Dr. Mariuxi Báez Moreira.')
h2('4.2 ORL team — HEEE')
body("The Otorhinolaryngology Department team at HEEE evaluated the patient on admission and identified "
     "Ludwig's Angina complicated with bilateral CNF as a first-order surgical emergency requiring "
     "urgent radical debridement.")
h2('4.3 Anesthesiology')
mixed([{'text':'Massive cervical edema and functional trismus required a predicted difficult airway protocol. '
        'Orotracheal intubation performed successfully under general anesthesia '
        '(confirmed by tracheal secretion culture, January 17, 2026). ASA classification: III.'},
       {'text':'14','sup':True}])
h2('4.4 ICU')
body('Postoperative ICU admission (Floor 1, HEEE) on January 17, 2026 at 23:48 h (Dr. Soraya Acaro) for continuous '
     'monitoring, metabolic support, and IV insulin therapy (perioperative target: 140–180 mg/dL).')
h2('4.5 Infectious Disease')
mixed([{'text':'Antibiotic adjustment at 48–72 h: meropenem established as definitive therapy upon '
        'confirmation of ESBL-producing '},
       {'text':'K. pneumoniae','italic':True},{'text':' and ESBL '},
       {'text':'E. coli','italic':True},
       {'text':' with MIC ≤0.25 µg/mL — the only active option against both ESBL isolates.'},
       {'text':'8,9','sup':True}])
h2('4.6 Plastic Surgery')
body('February 16, 2026 (~POD 29), with normalized inflammatory parameters (WBC 5.97 ×10³/µL, '
     'Na 138 mEq/L), transfer to Plastic Surgery (Dr. Juan Gutiérrez, Floor 3, HEEE) for '
     'reconstructive management of the residual cervical defect.')

# ── SURGICAL MANAGEMENT EN ────────────────────────────────────────────────────
doc.add_page_break()
h1('5. SURGICAL MANAGEMENT')
h2('5.1 Preoperative preparation')
for item in [
    'IV fluid resuscitation. Correction of hyponatremia (Na 126.9 mEq/L) and hypokalemia (K 2.9 mEq/L).',
    'Empirical broad-spectrum IV antibiotherapy (beta-lactam + metronidazole) pending cultures.',
    'IV insulin therapy (target 140–180 mg/dL). Request for 2 units packed RBCs (Hb 8.2–8.9 g/dL).',
    'Signed surgical and anesthetic informed consent. ASA classification: III.',
]: bullet(item)

h2('5.2 Surgical technique — Radical cervical debridement')
mixed([{'text':'Date: ','bold':True},{'text':'January 17, 2026 (20:20–22:20 h).'}])
mixed([{'text':'Lead surgeon: ','bold':True},
       {'text':'Dr. Christian Alfonso Freire Zamora, MD, ORL Specialist — Head of ORL Department, HEEE.'}])
mixed([{'text':'Approach: ','bold':True},
       {'text':'Bilateral radical cervical debridement under general anesthesia with orotracheal intubation. '
        'The necrotizing nature required radical excision of all devitalized tissue.'},
       {'text':'9,10','sup':True}])
body('Procedure description:')
for item in [
    'General anesthesia with orotracheal intubation (OTI) under predicted difficult airway protocol.',
    'Necrosis margin demarcation. Bilateral extent involving anterior, lateral, and submandibular cervical region.',
    'Cervical incisions with excision of necrotic eschar and devitalized fibrinoid tissue.',
    'Blunt and sharp debridement to healthy bleeding tissue.',
    'Evacuation of fetid purulent material. Samples sent for aerobic, anaerobic, fungal cultures and antibiogram.',
    'Thorough irrigation with warm 0.9% normal saline until clean return.',
    'Penrose drains placed through counter-incisions. Open wound — no primary closure.',
]: numbered(item)

mixed([{'text':'Intraoperative data: ','bold':True},
       {'text':'Procedure duration: 3 hours 10 minutes. Estimated blood loss: ~200 mL. '
        'Purulent fluid evacuated: ~350 mL. Intubation technique: '
        'not documented in the anesthetic protocol.'}])
body('Figures 10 and 11 document the intraoperative surgical field.')
fig_row([
    (F+'fig3a_intraop_desbrid_jan18.jpg', 8,
     'Intraoperative debridement (January 17, 2026): bilateral cervical surgical field with retractors. '
     'Necrotic and fibrinoid tissue in excision with active bleeding from viable tissue. '
     'Surgeon: Dr. Christian Alfonso Freire Zamora. ORL Department, HEEE.'),
    (F+'fig3b_intraop_diseccion_jan18.jpg', 8,
     'Intraoperative dissection (January 17, 2026): exposure of deep cervical fascial planes. '
     'Structures of the anterior cervical triangle visible after radical debridement.'),
])

# ── POSTOPERATIVE COURSE EN ───────────────────────────────────────────────────
doc.add_page_break()
h2('5.3 Subsequent surgical procedures')
body('Three additional surgical procedures were performed under general anesthesia during hospitalization, '
     'documented in the HEEE medical record (HC 0802060905):')
table('Surgical procedures during hospitalization — HEEE, January 2026',
    ['Date / POD','Procedure','Surgical findings','Department / Surgeon'],
    [['17/01/2026 (POD 0)\n20:20–22:20 h',
      'Wide cervicotomy + deep neck abscess drainage + deep neck debridement + necrotic tissue debridement anterior cervical region + tracheotomy',
      '• Abundant slough in cervical region (~80 g)\n• CNF anterior cervical region\n• Fetid greenish purulent fluid ~150 mL\n• Left pectoral collection ~40 mL\n• Trachea lateralized to the left\n• Thyroid gland inferiorly displaced\n• Blood loss: ~200 mL. Purulent fluid + slough: ~350 mL',
      'ORL / Dr. Christian Alfonso Freire Zamora\nAssistants: MD Cárdenas, MD León\nAnesthesiology: Dr. Vásquez'],
     ['21/01/2026 (POD 4)',
      'Deep neck debridement + slough debridement anterior cervical region + tissue culture',
      '• Scant slough in anterior cervical region\n• No necrotic tissue\n• Scant seropurulent secretion\n• Tracheotomy in situ, patent\n• Left pectoral abscess with edema and erythema',
      'Plastic Surgery'],
     ['27/01/2026 (POD 10)',
      'Surgical debridement + limited debridement + partial wound plasty of neck and anterior chest',
      '• Anterior cervical wound: no necrotic tissue, scant slough, no purulent discharge\n• Tracheotomy in situ, functional\n• Stoma without signs of infection\n• Anterior chest: raw area with moderate granulation tissue, 10 cm posterior pockets\n• Left chest: wound with pectoral major muscle exposure',
      'Plastic Surgery']],
    [3.5, 5, 6, 3])
space()

h1('6. POSTOPERATIVE COURSE')

h2('6.1 Definitive microbiological results')
table('Definitive microbiological results — HEEE Microbiology Laboratory (January–February 2026)',
    ['Sample (date)','Isolated organism','Sensitive to','Resistant to'],
    [['Right neck wound swab (17/01/2026)',
      [{'text':'Klebsiella pneumoniae ss. pneumoniae ESBL-producing','italic':True,'bold':True,'color':RED},
       {'text':' + '},{'text':'Candida tropicalis','italic':True},
       {'text':' (concomitant isolate, fungal culture 17/01/2026)'}],
      'Amikacin (MIC ≤1), Gentamicin, Imipenem, Meropenem (MIC ≤0.25)',
      'Amp/Sulbactam (MIC ≥32), Cefepime, Ceftriaxone (MIC ≥64), Ciprofloxacin, Pip/Tazo, TMP/SMX'],
     ['Neck wound swab — fungal culture (19/01/2026)',
      [{'text':'Candida tropicalis','italic':True},
       {'text':' + '},{'text':'Candida parapsilosis','italic':True},
       {'text':' — both species confirmed. No antifungal susceptibility testing performed.'}],
      'No antifungal susceptibility performed','—'],
     ['Necrotic neck tissue (18/01/2026)',
      [{'text':'Klebsiella pneumoniae ESBL','italic':True},{'text':' (2nd confirmation)'}],
      'Meropenem, Imipenem, Amikacin, Gentamicin','Same ESBL profile'],
     ['Neck secretion — 2nd sample (18/01/2026)',
      [{'text':'E. coli ESBL + K. pneumoniae ESBL','italic':True},{'text':' (polymicrobial)'}],
      'Amikacin (S), Gentamicin (S), Imipenem (MIC ≤1), Meropenem (MIC ≤1)',
      'Amp/Sulbactam, Cefepime, Ceftriaxone, Pip/Tazo, TMP/SMX'],
     ['Tracheal secretion (18/01/2026)','No bacterial growth','—','—'],
     ['Blood cultures ×4 sets (18/01/2026)','NO BACTERIAL GROWTH — bacteremia ruled out','—','—'],
     ['Rectal swab surveillance (17/01/2026)','NEGATIVE for carbapenem-resistant organisms (no KPC/MBL)','—','—'],
     ['Urine culture — catheter (17/01/2026)','No bacterial growth','—','—'],
     ['Urine — candiduria (04/02/2026)',
      [{'text':'Candida spp.','italic':True},{'text':' — Yeast +++ / Hyphae +++ (nosocomial candiduria, POD 17). Species not identified.'}],
      'Antifungal susceptibility not performed','—']],
    [4.5,5,4.5,4])

h2('6.2 Definitive antibiotic therapy')
table('Antibiotic regimen during hospitalization',
    ['Period','Drug','Rationale','Duration'],
    [['Empirical (from 17/01/2026)',
      'Meropenem 1 g q8h IV + Vancomycin 1 g q12h IV (from 17/01/2026)',
      'Empirical: Meropenem for ESBL gram-negatives; Vancomycin for gram-positive cocci. Initiated from ORL emergency on 17/01/2026.',
      'Meropenem: continued indefinitely (definitive). Vancomycin: 14 doses (~until 24/01/2026)'],
     ['Definitive (from 19/01)',
      [{'text':'Meropenem 1 g q8h IV','bold':True},
       {'text':' — only active agent vs ESBL K. pneumoniae and E. coli (MIC ≤0.25)'}],
      'First-line carbapenem for ESBL Enterobacteriaceae. Complete resistance to cephalosporins and beta-lactams.',
      '14–21 days IV'],
     ['Antifungal (from 20/01/2026 — POD 3)',
      'Fluconazole 150 mg IV q24h',
      'Nosocomial candiduria (POD 17) + Candida in cervical wound. Fluconazole as first-line azole antifungal.','14 days']],
    [3,5.5,5,4])

h2('6.3 Laboratory evolution during HEEE hospitalization')
table('Laboratory parameter evolution during hospitalization — HEEE (January–February 2026)',
    ['Date / Service','WBC ×10³','Hb g/dL','CRP¹ mg/L','PCT ng/mL','Na mEq/L','K mEq/L','Glucose mg/dL'],
    [['17/01 — ORL Emergency','13.11 ↑','8.3 ↓','202 ↑↑↑','0.31','133 ↓','2.9 ↓↓','120 ↑'],
     ['18/01 — ICU (preop)','13.69 ↑','8.9 ↓','189 ↑↑↑','0.26','130 ↓↓','3.3 ↓','164 ↑↑'],
     ['19/01 — ICU (postop)','7.06','8.2 ↓','—','—','132 ↓','3.8','93'],
     ['20/01 — ORL Floor 6','8.31','12.0 ↓','114 ↑↑','0.12','129 ↓↓','4.1','—'],
     ['22/01 — ORL Floor 6','5.35','10.4 ↓','—','0.08','134 ↓','4.1','129 ↑'],
     ['26/01 — ORL Floor 6','7.41','12.0 ↓','16.7 ↑','—','135','3.6','114 ↑'],
     ['05/02 — ORL Floor 6','7.67','11.1 ↓','8.85 ↑','—','135','4.9','135 ↑'],
     ['09/02 — ORL Floor 6','9.27','12.0 ↓','7.35 ↑','—','134 ↓','5.4 ↑','139 ↑'],
     ['16/02 — Plastic Surgery','5.97','10.9 ↓','—','—','138','4.2','131 ↑']],
    [4,2,2,2,2,2,2,2.5])
body('¹ CRP: C-reactive protein. PCT: procalcitonin.', italic=True, color=GRAY, before=0, after=4)
mixed([{'text':'C-reactive protein (CRP) declined from 202 mg/L to 7.35 mg/L by postoperative day 22, confirming response '
        'to debridement and meropenem.'},
       {'text':'12','sup':True},
       {'text':" Procalcitonin remained < 0.5 ng/mL throughout follow-up, consistent with negative blood cultures."},
       {'text':'24,15','sup':True}])

h2('6.4 Wound evolution — Photographic documentation')
body('Figures 12–13 show early postoperative evolution (postoperative day [POD] 6, January 24, 2026). '
     'Figures 15–17 document outpatient follow-up (~POD 71, March 30–31, 2026) with '
     'active granulation tissue in the cervical and infraclavicular defects.')
fig_row([
    (F+'fig4b_postop_herida_jan24.jpg', 8,
     'Postoperative evolution (POD 6, January 24, 2026): general view of cervical wound with active blue drain '
     'and bilateral open wound management. Evident perilesional edema reduction compared to admission (January 17, 2026).'),
    (F+'fig4a_postop_traqueo_jan24.jpg', 8,
     'Postoperative evolution (POD 6, January 24, 2026): close-up of cervical tracheostomy with cannula in place, '
     'approximation sutures and incipient granulation tissue at wound margins.'),
])
fig_row([
    (F+'fig5a_control_cervical_mar30.jpg', 5.5,
     'Outpatient follow-up (~POD 71, March 30, 2026): submandibular cervical defect with active '
     'granulation tissue and right infraclavicular defect healing by secondary intention.'),
    (F+'fig5b_control_infraclavicular_mar30.jpg', 5.5,
     'Outpatient follow-up (March 30, 2026): close-up of infraclavicular defect with active '
     'granulation tissue, well-defined edges, and no signs of active infection.'),
    (F+'fig5c_control_panoramico_mar31.jpg', 5.5,
     'Outpatient follow-up (March 31, 2026): panoramic cervical view showing favorable defect '
     'evolution. Active granulation in resolution. Healing by secondary intention ongoing.'),
])

h2('6.5 Postoperative complications')
mixed_bullet([{'text':'Acute hypoxemia (POD 5, January 23): ','bold':True},
              {'text':'SpO2 85.7%, PaO2 47.9 mmHg, D-dimer 1.27 µg/mL. Compatible with atelectasis or aspiration pneumonia.'}])
mixed_bullet([{'text':'Recurrent hyponatremia and hypokalemia: ','bold':True},
              {'text':'Minimum Na 129 mEq/L, K 2.9 mEq/L. Progressive IV correction.'}])
mixed_bullet([{'text':'Nosocomial candiduria (POD 17, February 4): ','bold':True},
              {'text':'Candida spp. yeast +++ / hyphae +++. Treated with fluconazole 150 mg IV q24h for 14 days with favorable resolution.'}])
mixed_bullet([{'text':'Progressive eosinophilia: ','bold':True},
              {'text':'Up to 18.7% (February 9). Compatible with drug reaction to prolonged antibiotherapy.'}])
mixed_bullet([{'text':'Persistent anemia of chronic disease: ','bold':True},
              {'text':'Hb 8.2–12.0 g/dL throughout hospitalization.'}])

# ── DISCUSSION EN ─────────────────────────────────────────────────────────────
doc.add_page_break()
h1('7. DISCUSSION')

h2('7.1 Differential diagnosis')
table('Differential diagnosis of deep cervicofacial infections',
    ['Entity','Clinical features','Imaging findings','Management'],
    [['Cervical phlegmon',
      'Diffuse induration, no fluctuation, no necrosis.',
      'Diffuse soft tissue thickening. No gas or collection.',
      [{'text':'IV antibiotics + close monitoring.'},{'text':'12,13','sup':True}]],
     ['Deep neck abscess',
      'Fluctuation, trismus, moderate fever.',
      'Hypodense collection with rim enhancement. Gas possible.',
      [{'text':'Transcervical drainage + IV antibiotics.'},{'text':'12,13','sup':True}]],
     ["Ludwig's Angina (onset)",
      [{'text':'CASE (onset): ','bold':True},
       {'text':'Bilateral cellulitis floor of mouth, functional trismus, no fluctuation.'}],
      'Bilateral submandibular/sublingual thickening. No defined collection.',
      [{'text':'Airway management + IV ATB + debridement if progression.'},{'text':'1,2','sup':True}]],
     ['Cervical CNF',
      [{'text':'CASE (progression): ','bold':True},
       {'text':'Extensive necrosis, black eschar, fetid fistulas, crepitation, systemic sepsis.'}],
      'Soft tissue gas, fascial necrosis, confirmed cutaneous fistulas (see Figures 1–9).',
      [{'text':'Urgent radical debridement + carbapenems + ICU.'},{'text':'9,10','sup':True}]]],
    [3.5,4.5,4.5,4])
space()

h2("7.2 Progression from Ludwig's Angina to CNF in a diabetic patient")
mixed([{'text':"DM causes neutrophil dysfunction, microangiopathy, and peripheral neuropathy that favor "
        "silent progression of Ludwig's Angina to CNF."},
       {'text':'1,2,8,21','sup':True},
       {'text':" Severe hyperglycemia on admission (glucose 299–331 mg/dL) and the one-month delay in care "
        "determined the irreversible progression to fascial necrosis, evidenced in all three CT planes "
        "(Figures 1–8) and in the clinical photographs on admission (Figure 9)."},
       {'text':'8','sup':True}])

mixed([{'text':"In line with Huang et al., diabetic patients with deep neck infections have longer "
        "hospital stays and higher complication rates than non-diabetic patients."},
       {'text':'8','sup':True},
       {'text':" Mariano et al. highlight that progression to CNF in diabetic patients is a poor prognostic factor, "
        "with early debridement being the only modifiable element impacting survival."},
       {'text':'21','sup':True},
       {'text':" In the present case, the one-month delay in care transformed a potentially manageable "
        "Ludwig's Angina into established CNF requiring ICU management and subsequent surgical reconstruction."}])
h2('7.3 ESBL-producing Klebsiella pneumoniae in a community-onset infection')
mixed([{'text':'Isolation of ESBL-producing '},{'text':'K. pneumoniae','italic':True},
       {'text':' as the primary pathogen in community-onset CNF (outpatient dental extraction) is unusual '
        'and critically important: it mandates carbapenems as the only active treatment, ruling out all '
        'beta-lactams and cephalosporins.'},
       {'text':'7,8','sup':True},
       {'text':" Meropenem selection (MIC ≤0.25 µg/mL — optimal MIC in the antibiogram) was ideal. "
        "Negative blood cultures (4 sets) and CRP decline confirm the infection was controlled without bacteremia."},
       {'text':'15','sup':True}])

mixed([{'text':'Marioni et al. reported that 74% of deep neck infections of dental origin were caused '
        'by mixed polymicrobial flora (streptococci and anaerobes).'},
       {'text':'25','sup':True},
       {'text':' The isolation of '},{'text':'K. pneumoniae','italic':True},
       {'text':' ESBL as the dominant microorganism is therefore unusual in community-onset infections, '
        'suggesting prior colonization or unidentified nosocomial exposure.'}])
h2('7.4 Surgical approaches by fascial space')
table('Surgical approaches based on compromised cervical fascial space',
    ['Fascial space','Approach of choice','Alternative approach','Ref.'],
    [["Submandibular (Ludwig's Angina)",
      'Bilateral submandibular + submental',
      'Prophylactic tracheostomy',
      [{'text':'1,2','sup':True}]],
     ['Extensive bilateral CNF (PRESENT CASE)',
      'Wide radical debridement (Figures 10–11)',
      'Multiple relaxation incisions ± VAC therapy',
      [{'text':'9,10','sup':True}]],
     ['Parapharyngeal','Transcervical submandibular','Transoral selected',[{'text':'13','sup':True}]],
     ['Retropharyngeal','Transcervical + transoral','Transoral in pediatrics',[{'text':'16','sup':True}]],
     ['Carotid','Transcervical anterior to SCM','—',[{'text':'17','sup':True}]],
     ['Danger space / Mediastinum','Extended transcervical + VATS','Right thoracotomy',[{'text':'18','sup':True}]]],
    [4,5,5,2])
space()

h2('7.5 Complications')
mixed_bullet([{'text':'DNM: ','bold':True},
              {'text':'Mortality 11.8–40%. CT ruled out mediastinal extension.'},
              {'text':'4','sup':True}])
mixed_bullet([{'text':'Bacteremia / Sepsis: ','bold':True},
              {'text':'Four blood culture sets negative. Favorable response to source control and meropenem.'},
              {'text':'15','sup':True}])
mixed_bullet([{'text':"Lemierre's syndrome: ",'bold':True},
              {'text':'Not documented (negative cultures). Elevated D-dimer (1.27 µg/mL) warranted surveillance.'},
              {'text':'19','sup':True}])
mixed_bullet([{'text':'Nosocomial candidosis: ','bold':True},
              {'text':'Candiduria POD 17 in T2DM + prolonged antibiotics + urinary catheter. Required antifungal therapy.'},
              {'text':'20','sup':True}])
mixed_bullet([{'text':'Cervical reconstruction: ','bold':True},
              {'text':'Residual defect managed by Plastic Surgery (~POD 29), documented in Figures 10–14.'},
              {'text':'9','sup':True}])


h2('7.6 Limitations of the case')
mixed_bullet([{'text':'Retrospective design: ','bold':True},
              {'text':'Data obtained through medical record review after discharge, with inherent risk of incomplete information.'}])
mixed_bullet([{'text':'Intubation technique not documented: ','bold':True},
              {'text':'The anesthetic protocol does not specify the method used for difficult airway management '
               '(videolaryngoscopy, fiberoptic bronchoscopy, or other).'}])
mixed_bullet([{'text':'Incomplete pulmonary evaluation: ','bold':True},
              {'text':'The POD 5 acute hypoxemia episode (SpO₂ 85.7%, D-dimer 1.27 µg/mL) has no CT pulmonary '
               'angiography record, limiting formal exclusion of pulmonary thromboembolism.'}])
mixed_bullet([{'text':'Single case (n = 1): ','bold':True},
              {'text':'Findings are not generalizable. Multicenter prospective series are required to validate conclusions.'}])
mixed_bullet([{'text':'Non-protocolized follow-up: ','bold':True},
              {'text':'Although active outpatient follow-up continues, no formal long-term protocol exists, '
               'limiting assessment of the definitive functional and aesthetic outcome.'}])
mixed_bullet([{'text':'Absence of confirmatory anaerobic cultures: ','bold':True},
              {'text':'Anaerobic cultures showed no growth, possibly due to technical sensitivity or prior antibiotics, '
               'not ruling out their role in the polymicrobial infection.'}])

# ── CONCLUSIONS EN ────────────────────────────────────────────────────────────
doc.add_page_break()
h1('8. CONCLUSIONS')
for item in [
    "Ludwig's Angina complicated with cervical necrotizing fasciitis is a surgical emergency with potential mortality of 15–40%. Early multiplanar CT diagnosis (sagittal, coronal, and axial planes), urgent tertiary center transfer, and immediate radical debridement are the cornerstones of treatment.",
    "Molar extraction without antibiotic prophylaxis in a diabetic patient with severe hyperglycemia (glucose 299–331 mg/dL) was the triggering event. Prevention requires strict glycemic control prior to dental procedures in high-risk patients.",
    "ESBL-producing Klebsiella pneumoniae as the primary pathogen — unusual in community-onset infections — substantially modifies the antibiotic strategy: meropenem (MIC ≤0.25 µg/mL) was the only definitively active option against both ESBL isolates.",
    "Negative blood cultures (4 sets), CRP decline from 202 to 7.35 mg/L over 22 days, and the absence of mediastinal extension confirm the effectiveness of the established management.",
    "A multidisciplinary ORL–ICU–Infectious Disease–Plastic Surgery approach supervised by the HEEE ORL Department was key in achieving a favorable outcome with 33 days of hospitalization (January 16 – February 18, 2026) and healing by secondary intention (documented in Figures 12–17).",
    "The one-month delay in seeking care from Santo Domingo reflects a critical gap in Ecuador's health system for early detection of odontogenic complications at the primary care level in diabetic patients.",
]: bullet(item)

# ── REFERENCES EN ─────────────────────────────────────────────────────────────
doc.add_page_break()
h1('9. REFERENCES')
body('Vancouver style. Numbered in order of first citation in text.', italic=True, color=GRAY)
space()
for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_spacing(p, before=3, after=3, line=276)
    rn = p.add_run(f'{i}. ')
    rn.bold = True; rn.font.name = 'Arial'; rn.font.size = Pt(11)
    rt = p.add_run(ref)
    rt.font.name = 'Arial'; rt.font.size = Pt(11)


# ── CRediT AUTHORSHIP CONTRIBUTION STATEMENT EN ───────────────────────────────
doc.add_page_break()
h1('AUTHORSHIP CONTRIBUTION STATEMENT (CRediT)')
body('In accordance with the CRediT (Contributor Roles Taxonomy), author contributions to this case report are as follows:', before=4, after=6)
table('Author contributions according to CRediT taxonomy',
    ['Role','J.P. Gavilanez Heras','C.A. Freire Zamora','H.J. Campoverde Sánchez','G. Saenz Ortega'],
    [['Conceptualization','✓','','',''],
     ['Data curation','✓','','✓',''],
     ['Formal analysis','✓','','',''],
     ['Investigation','✓','','',''],
     ['Methodology','✓','','',''],
     ['Supervision','','✓','',''],
     ['Validation','','✓','','✓'],
     ['Visualization','✓','','',''],
     ['Software / Digital analysis','','','','✓'],
     ['Writing — original draft','✓','','',''],
     ['Writing — review & editing','✓','✓','✓','']],
    [5.5, 3, 3, 3, 3])
body('All authors have read and approved the final version of the manuscript for publication.', italic=True, before=6, after=4)

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
out = '/home/claude/BILINGUE_v8.docx'
doc.save(out)
print(f'Saved: {out}')
with zipfile.ZipFile(out) as zf:
    imgs = [n for n in zf.namelist() if 'media' in n]
    print(f'ZIP OK — {len(zf.namelist())} files — {len(imgs)} images embedded')

# ════════════════════════════════════════════════════════════════════════════
# POST-PROCESSING: inject axial CT section, timeline, limitaciones, discussion
# ════════════════════════════════════════════════════════════════════════════
# NOTE: This block is appended and will be inserted via build injection below
