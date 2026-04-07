# Caso Clínico — Constructor de Reporte Bilingüe (ES/EN)

Sistema de scripts Python para generar automáticamente un reporte de caso clínico académico en formato Word (`.docx`), bilingüe español/inglés, a partir de una historia clínica en formato MSP/HOSVITAL y archivos de resultados de laboratorio.

---

## Tabla de contenidos

1. [Descripción general](#descripción-general)
2. [Requisitos](#requisitos)
3. [Instalación](#instalación)
4. [Inicio rápido](#inicio-rápido)
5. [Flujo de trabajo](#flujo-de-trabajo)
6. [Scripts](#scripts)
   - [build_case_report.py](#build_case_reportpy)
   - [parse_clinical_history.py](#parse_clinical_historypy)
   - [parse_lab_pdfs.py](#parse_lab_pdfspy)
7. [Configuración YAML](#configuración-yaml)
   - [Campos auto-extraídos vs. manuales](#campos-auto-extraídos-vs-manuales)
   - [Estructura de secciones](#estructura-de-secciones)
8. [Estructura de archivos del proyecto](#estructura-de-archivos-del-proyecto)
9. [Marcadores de contenido pendiente](#marcadores-de-contenido-pendiente)
10. [Adaptar a otro caso clínico](#adaptar-a-otro-caso-clínico)

---

## Descripción general

El sistema genera un documento Word con las siguientes secciones:

| # | Sección | Fuente principal |
|---|---------|-----------------|
| Portada | Autores, afiliaciones, ORCIDs, título | YAML (manual) |
| Resumen / Abstract | Introducción + presentación + conclusiones | YAML / IA |
| 1. Introducción | Subsecciones + objetivo | YAML / IA |
| 2. Presentación del caso | Datos paciente, motivo, antecedentes, signos vitales, examen físico, diagnósticos | **Auto: HC PDF** |
| 3. Estudios de imagen | Grupos de imágenes + TAC | YAML + archivos imagen |
| 4. Evaluación multidisciplinaria | Subsecciones | YAML |
| 5. Manejo quirúrgico | Cronograma de cirugías | YAML |
| 6. Evolución posoperatoria | Labs, antibióticos, fotos herida | YAML + imágenes |
| 7. Discusión | Diagnóstico diferencial, subsecciones | YAML / IA |
| 8. Conclusiones | Lista de conclusiones | YAML / IA |
| 9. Referencias | Estilo Vancouver | YAML |
| CRediT | Tabla de contribución de autores | YAML |

---

## Requisitos

- **Python 3.10+**
- Paquetes:

```bash
pip install python-docx pyyaml pymupdf
# Opcional — para extracción de fotogramas de video:
pip install opencv-python-headless
```

| Paquete | Uso |
|---------|-----|
| `python-docx` | Generación del documento Word |
| `pyyaml` | Lectura del archivo de configuración |
| `pymupdf` (fitz) | Extracción de texto de PDFs |
| `opencv-python-headless` | Extracción de fotogramas de videos (opcional) |

---

## Instalación

```bash
git clone https://github.com/gabrielsaenz20/Caso-clinico.git
cd Caso-clinico
pip install python-docx pyyaml pymupdf
```

---

## Inicio rápido

```bash
# 1. Generar el documento Word completo
python3 build_case_report.py --config case_config_template.yaml

# 2. Solo extraer fotogramas e imágenes (sin generar el .docx)
python3 build_case_report.py --config case_config_template.yaml --figs-only

# 3. Ver qué datos se extraen de la historia clínica
python3 parse_clinical_history.py "PLAZA QUIMIS GLADIS MARIA.pdf"

# 4. Exportar los datos extraídos a YAML
python3 parse_clinical_history.py "PLAZA QUIMIS GLADIS MARIA.pdf" --output datos_paciente.yaml

# 5. Extraer resultados de laboratorio de los PDFs
python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/"
python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/" --output lab_results.yaml
```

El resultado es `case_report_FINAL.docx` (o el nombre definido en `output_file`).

---

## Flujo de trabajo

```
┌─────────────────────────────────┐
│  PLAZA QUIMIS GLADIS MARIA.pdf  │  ← Historia clínica (MSP/HOSVITAL)
│  (formato estándar MSP Ecuador) │
└────────────────┬────────────────┘
                 │  parse_clinical_history.py (automático)
                 ▼
    ┌────────────────────────┐
    │  Datos auto-extraídos  │
    │  • Demographics        │
    │  • Motivo consulta     │
    │  • Enfermedad actual   │
    │  • Antecedentes        │
    │  • Signos vitales      │
    │  • Examen físico       │
    │  • Diagnósticos CIE-10 │
    │  • Notas evolución     │
    └────────────┬───────────┘
                 │
┌────────────────┼────────────────────────────────────────┐
│  case_config_template.yaml                              │
│  (solo campos que NO están en la HC):                   │
│  • institution, title, authors, affiliations            │
│  • abstract, introduction, discussion, conclusions      │
│  • imaging (grupos de imágenes)                         │
│  • surgical management, postoperative evolution         │
│  • references, CRediT contributions                     │
│  • preingreso_labs (tabla multi-fecha del hospital ref) │
└────────────────┬────────────────────────────────────────┘
                 │
                 ▼
         build_case_report.py
                 │
                 ▼
    ┌────────────────────────┐
    │  case_report_FINAL.docx│  ← Documento Word bilingüe ES/EN
    └────────────────────────┘
```

---

## Scripts

### `build_case_report.py`

Constructor principal del reporte.

```
Uso: python3 build_case_report.py --config ARCHIVO.yaml [opciones]

Opciones:
  --config FILE      Archivo de configuración YAML (requerido)
  --figs-only        Solo extraer fotogramas/imágenes, no generar el .docx
  --skip-frames      Omitir extracción de fotogramas de video
  -h, --help         Mostrar ayuda
```

**Comportamiento al inicio:**
1. Lee el archivo YAML.
2. Si `sources.clinical_history_pdf` está definido, llama automáticamente a `parse_clinical_history.py` y enriquece la configuración con los datos extraídos.
3. Si `sources.lab_pdf_dir` está definido, puede usarse con `parse_lab_pdfs.py`.
4. Genera el documento Word con versiones en español e inglés separadas por una página divisora.

**Secciones que se generan automáticamente** (sin necesidad de datos en el YAML):

| Campo YAML generado automáticamente | Fuente |
|--------------------------------------|--------|
| `case_presentation.demographics_table` | HC PDF |
| `case_presentation.chief_complaint.narrative_es` | HC PDF |
| `case_presentation.chief_complaint.present_illness_es` | HC PDF |
| `case_presentation.chief_complaint.medical_history_es` | HC PDF |
| `case_presentation.chief_complaint.surgical_history_es` | HC PDF |
| `case_presentation.chief_complaint.allergies_es` | HC PDF |
| `case_presentation.chief_complaint.gyneco_es` | HC PDF |
| `case_presentation.chief_complaint.habits_es` | HC PDF |
| `case_presentation.physical_exam.vital_signs_table` | HC PDF |
| `case_presentation.physical_exam.findings_es` | HC PDF |
| `case_presentation.clinical_analysis_es` | HC PDF |
| `case_presentation.diagnoses_list` | HC PDF |

---

### `parse_clinical_history.py`

Parser para historias clínicas en formato MSP/HOSVITAL (Ecuador).

```
Uso: python3 parse_clinical_history.py PDF [opciones]

Argumentos:
  PDF                Ruta al archivo PDF de la historia clínica

Opciones:
  --output FILE      Guardar datos extraídos en YAML
  --summary          Mostrar resumen en consola (predeterminado: activado)
  --no-summary       Suprimir el resumen en consola
  -h, --help         Mostrar ayuda
```

**Datos extraídos:**

| Campo | Descripción |
|-------|-------------|
| `demographics.name` | Nombre completo del paciente |
| `demographics.hcu` | Número de historia clínica |
| `demographics.sex` | Sexo |
| `demographics.dob` | Fecha de nacimiento (DD/MM/YYYY) |
| `demographics.age` | Edad en años |
| `demographics.phone` | Teléfono |
| `demographics.admission_date` | Fecha de ingreso |
| `demographics.total_pages` | Total de páginas del PDF |
| `filiation.origin` | Procedencia |
| `filiation.education` | Nivel de instrucción |
| `filiation.occupation` | Ocupación |
| `filiation.civil_status` | Estado civil |
| `filiation.blood_type` | Grupo sanguíneo |
| `filiation.family_contact` | Contacto familiar (nombre y teléfono) |
| `chief_complaint` | Motivo de consulta |
| `present_illness` | Enfermedad actual (texto completo) |
| `medical_history` | Antecedentes patológicos clínicos (lista) |
| `surgical_history` | Antecedentes quirúrgicos (lista) |
| `allergies` | Alergias |
| `family_history` | Antecedentes familiares |
| `habits` | Hábitos (alcohol, tabaco, drogas, biomasa, vacuna COVID) |
| `gyneco` | Antecedentes gineco-obstétricos (G/P/A/C/HV) |
| `vital_signs` | PA, FC, FR, SpO₂, temperatura, Glasgow |
| `physical_exam` | Examen físico por sistemas (texto) |
| `clinical_analysis` | Análisis clínico del médico tratante |
| `diagnoses` | Lista de diagnósticos con código CIE-10 |
| `evolution_notes` | Resumen de notas de evolución (folio, fecha, tipo) |

**Uso como módulo Python:**

```python
from parse_clinical_history import parse_clinical_history

data = parse_clinical_history("PLAZA QUIMIS GLADIS MARIA.pdf")

print(data['demographics']['name'])      # GLADIS MARIA PLAZA QUIMIZ
print(data['vital_signs']['pa'])         # 100/47
print(data['diagnoses'][0]['code'])      # K122
print(len(data['evolution_notes']))      # 146
```

**Formato de historia clínica compatible:**

El parser está diseñado para el formato estándar exportado por el sistema HOSVITAL del Ministerio de Salud Pública del Ecuador (MSP). Características del formato:
- Encabezado repetido en cada página con datos del paciente y labels (`Sexo:`, `Edad actual:`, etc.)
- Bloque de encabezado termina con la línea `Atención Especial:`
- Secciones en mayúsculas: `DATOS DE FILIACIÓN`, `ANTECEDENTES PATOLÓGICOS PERSONALES`, `MOTIVO DE CONSULTA:`, `ENFERMEDAD ACTUAL:`, `EXAMEN FÍSICO:`, `ANÁLISIS:`, `PLAN:`
- Signos vitales con formato: `PRESIÓN ARTERIAL: NNN/NNN MILÍMETROS DE MERCURIO`
- Diagnósticos en formato: `NOMBRE DEL DIAGNÓSTICO (CIE 10: XXXN)` o bloque tabulado

---

### `parse_lab_pdfs.py`

Extractor de resultados de laboratorio a partir de PDFs del laboratorio hospitalario.

```
Uso: python3 parse_lab_pdfs.py DIRECTORIO [opciones]

Argumentos:
  DIRECTORIO         Carpeta que contiene los PDFs de resultados de laboratorio

Opciones:
  --output FILE      Guardar resultados en YAML
  --summary          Mostrar resumen en consola
  -h, --help         Mostrar ayuda
```

**Categorías de laboratorio detectadas:**
`HEMATOLOGÍA`, `BIOQUÍMICOS`, `ELECTROLITOS`, `GASOMETRÍA`, `HEMOSTASIA`, `INMUNOQUÍMICA`, `UROANÁLISIS`, `MICROBIOLOGÍA`, `CULTIVO`, `HEMOCULTIVO`, entre otras.

**Resultado por cada parámetro:**
```yaml
- category: HEMATOLOGÍA
  name: Leucocitos
  result: "13.11"
  unit: x10^3/UL
  reference: 4.32 - 10.42
  pathological: true
```

---

## Configuración YAML

El archivo `case_config_template.yaml` es el punto de entrada principal para personalizar el reporte. Todas las rutas son **relativas** a la ubicación del propio archivo YAML.

### Campos auto-extraídos vs. manuales

#### ✅ Auto-extraídos (no hay que escribirlos)

Estos campos se llenan automáticamente desde `sources.clinical_history_pdf`:

- Nombre del paciente, HC, sexo, edad, fecha de nacimiento, teléfono
- Procedencia, instrucción, ocupación, estado civil, grupo sanguíneo
- Contacto familiar
- Motivo de consulta, enfermedad actual
- Antecedentes patológicos, quirúrgicos, familiares, alergias
- Hábitos (alcohol, tabaco, biomasa, vacuna COVID)
- Antecedentes gineco-obstétricos
- Signos vitales (PA, FC, FR, SpO₂, temperatura, Glasgow)
- Examen físico completo por sistemas
- Análisis clínico del médico
- Lista de diagnósticos con código CIE-10

#### ✏️ Manuales (deben completarse en el YAML)

Estos campos no se pueden extraer de la historia clínica y deben ingresarse a mano:

| Campo | Descripción |
|-------|-------------|
| `sources.clinical_history_pdf` | Ruta al PDF de la historia clínica |
| `institution` | Datos del hospital (nombre, servicio, ciudad) |
| `title` | Título del artículo (ES y EN), con partes en itálica |
| `authors` | Lista de autores con nombre, ORCID, afiliación, email |
| `affiliations` | Texto de afiliaciones institucionales |
| `abstract.introduction_*` | Párrafo introductorio del resumen |
| `abstract.case_*` | Párrafo de presentación del caso en el resumen *(IA)* |
| `abstract.conclusions_*` | Conclusiones del resumen *(IA)* |
| `abstract.keywords_*` | Palabras clave |
| `introduction.subsections[*].text_*` | Texto de cada subsección de la introducción *(IA)* |
| `introduction.objective_*` | Objetivo del reporte *(IA)* |
| `case_presentation.preingreso_labs` | Tabla de laboratorios multi-fecha del hospital remitente |
| `imaging` | Grupos de imágenes, archivos, pies de figura |
| `multidisciplinary` | Subsecciones de evaluación multidisciplinaria |
| `surgical_management` | Cronograma de procedimientos quirúrgicos |
| `postoperative_evolution` | Labs posoperatorios, antibióticos, fotos de herida |
| `discussion.subsections[*].text_*` | Texto de cada subsección de discusión *(IA)* |
| `conclusions.items_*` | Lista de conclusiones *(parcialmente IA)* |
| `references` | Lista de referencias en estilo Vancouver |
| `author_contributions` | Tabla CRediT de contribución de autores |
| `video_frames` | Timestamps para extracción de fotogramas |
| `image_copies` | Copias/renombres de imágenes hacia `figs_dir` |

> Los campos marcados con *(IA)* aparecerán en el documento con un marcador naranja `⚠ [AI_PENDIENTE: ...]` mientras no tengan contenido final.

### Estructura de secciones

```yaml
# ── Fuentes (auto-extracción) ─────────────────────────────────────────────────
sources:
  clinical_history_pdf: "APELLIDO NOMBRE.pdf"
  lab_pdf_dir: "EXAMENES DE LABORATORIO"

# ── Salida ────────────────────────────────────────────────────────────────────
output_file: "case_report_FINAL.docx"
figs_dir: "figs"

# ── Institución ───────────────────────────────────────────────────────────────
institution:
  ministry_es: "MINISTERIO DE SALUD PÚBLICA DEL ECUADOR"
  ministry_en: "MINISTRY OF PUBLIC HEALTH OF ECUADOR"
  name_es: "Hospital de Especialidades Eugenio Espejo"
  name_en: "Eugenio Espejo Specialty Hospital"
  department_es: "Servicio de Otorrinolaringología"
  department_en: "Otorhinolaryngology Department"
  city: "Quito, Ecuador"

# ── Título ────────────────────────────────────────────────────────────────────
title:
  es: "TÍTULO EN ESPAÑOL CON PARTE FINAL EN "
  en: "TITLE IN ENGLISH WITH FINAL PART IN "
  suffix_italic_es: "Klebsiella pneumoniae"   # texto en itálica
  suffix_end_es: " PRODUCTORA DE BLEE"        # texto final tras la itálica
  suffix_italic_en: "Klebsiella pneumoniae"
  suffix_end_en: ""

# ── Autores ───────────────────────────────────────────────────────────────────
authors:
  - name: "Apellido Nombre, MD"
    affiliation: 1
    orcid: "0000-0000-0000-0000"
    corresponding: true
    email: "correo@hospital.gob.ec"
    institution_city: "Hospital, Ciudad, País"
    suffix_es: "¹"
    suffix_en: "¹"
  - name: "Apellido Nombre, MD, Esp."
    affiliation: 1
    orcid: "0000-0000-0000-0000"
    corresponding: false
    supervisor: true
    supervisor_note_es: "Jefe del Servicio, Tutor del caso."
    supervisor_note_en: "Department Head, Case Supervisor."
    suffix_es: "¹*"
    suffix_en: "¹*"

affiliations:
  1:
    es: "Servicio X, Hospital Y, MSP del Ecuador. Ciudad, País."
    en: "Department X, Hospital Y, Ministry of Public Health of Ecuador. City, Country."

# ── Presentación del caso ─────────────────────────────────────────────────────
# (la mayor parte se auto-extrae de clinical_history_pdf)
case_presentation:
  preingreso_labs:        # tabla de laboratorios del hospital remitente
    table:
      title_es: "Laboratorios preingreso"
      headers_es: ["Parámetro", "Fecha 1", "Fecha 2", "Referencia"]
      widths: [4.5, 3, 3, 3]
      rows_es:
        - ["Leucocitos x10³/µL", "13,17 ↑", "19,55 ↑↑", "4,0–10,0"]

# ── Imágenes ──────────────────────────────────────────────────────────────────
imaging:
  image_groups:
    - id: grupo1
      subsection_title_es: "3.1 Descripción"
      subsection_title_en: "3.1 Description"
      narrative_es: "Texto descriptivo de las imágenes."
      narrative_en: "Descriptive text of the images."
      images:
        - file: "figs/imagen.jpg"
          width_cm: 8
          caption_es: "Pie de figura en español."
          caption_en: "Figure caption in English."

# ── Fotogramas de video (extracción automática) ───────────────────────────────
video_frames:
  - video: "VIDEO.mp4"
    timestamp_sec: 5.0
    output: "figs/frame_video.jpg"
    description: "Descripción del fotograma"

# ── Copias de imágenes hacia figs_dir ─────────────────────────────────────────
image_copies:
  - source: "IMG-20260117-WA0012.jpg"
    dest: "figs/fig2_descripcion.jpg"
    description: "Descripción"
```

---

## Estructura de archivos del proyecto

```
Caso-clinico/
├── build_case_report.py         # Constructor principal del reporte .docx
├── parse_clinical_history.py    # Parser de historia clínica (MSP/HOSVITAL)
├── parse_lab_pdfs.py            # Extractor de resultados de laboratorio
├── case_config_template.yaml    # Configuración del reporte (plantilla)
│
├── PLAZA QUIMIS GLADIS MARIA.pdf   # Historia clínica del paciente (HOSVITAL)
├── referencia plaza quimiz.pdf     # Epicrisis del hospital remitente
├── referencia plaza quimiz 2.pdf   # Epicrisis del hospital remitente (2)
│
├── EXAMENES DE LABORATORIO/     # PDFs de resultados de laboratorio (HEEE)
│
├── figs/                        # Imágenes procesadas para el reporte
│   ├── fig2_preop_necrosis_jan17.jpg
│   ├── fig3a_intraop_desbrid_jan18.jpg
│   ├── fig3b_intraop_diseccion_jan18.jpg
│   ├── fig4a_postop_traqueo_jan24.jpg
│   ├── fig4b_postop_herida_jan24.jpg
│   ├── fig5a_control_cervical_mar30.jpg
│   ├── fig5b_control_infraclavicular_mar30.jpg
│   └── fig5c_control_panoramico_mar31.jpg
│
├── IMG-20260117-*.jpg           # Imágenes clínicas originales (WhatsApp/cámara)
├── WhatsApp Image 2026-*.jpeg   # Imágenes clínicas originales
├── MOVIE-*.mp4 / VID-*.mp4      # Videos clínicos
│
└── case_report_FINAL.docx       # Documento Word generado (salida)
```

---

## Marcadores de contenido pendiente

Los campos que aún no tienen contenido final (secciones que requieren redacción académica o revisión por IA) aparecen en el documento Word con un marcador naranja:

```
⚠ [AI_PENDIENTE: Redactar párrafo de presentación del caso...]
```

En inglés:
```
⚠ [AI_PENDING: Write case presentation paragraph...]
```

Al ejecutar `build_case_report.py` se muestra al final una lista de todos los campos pendientes:

```
────────────────────────────────────────────────────────────
  SECCIONES PENDIENTES DE REVISIÓN / IA:
────────────────────────────────────────────────────────────
  ⚠ abstract.case_es: [AI_PENDIENTE: ...]
  ⚠ introduction.subsections[0].text_es: [AI_PENDIENTE: ...]
  ...
```

Para marcar un campo como pendiente en el YAML, use la sintaxis:
```yaml
text_es: "[AI_PENDIENTE: Describir el mecanismo de progresión en 100 palabras]"
text_en: "[AI_PENDING: Describe the progression mechanism in 100 words]"
```

---

## Adaptar a otro caso clínico

Para reutilizar el sistema con un nuevo caso:

1. **Copiar la carpeta del proyecto** o crear una nueva con los mismos scripts.

2. **Agregar el PDF de la historia clínica** — el nombre puede ser cualquiera, por ejemplo:
   ```
   APELLIDO NOMBRES.pdf
   ```

3. **Crear un nuevo YAML** basándose en `case_config_template.yaml`:
   ```bash
   cp case_config_template.yaml mi_nuevo_caso.yaml
   ```

4. **Editar el YAML** — solo los campos manuales (institución, título, autores, imágenes, texto de discusión, referencias). Cambiar `sources.clinical_history_pdf` al nombre del nuevo PDF.

5. **Colocar las imágenes y videos** en la carpeta del proyecto y actualizar las secciones `imaging`, `video_frames` e `image_copies` del YAML.

6. **Ejecutar el build:**
   ```bash
   python3 build_case_report.py --config mi_nuevo_caso.yaml
   ```

7. **Revisar los marcadores naranja** en el documento generado y completar las secciones pendientes de redacción académica.

> **Nota:** El parser de historia clínica (`parse_clinical_history.py`) está optimizado para el formato HOSVITAL del MSP Ecuador. Si el PDF proviene de otro sistema hospitalario, puede ser necesario ajustar los patrones de extracción en `SECTION_MARKERS` y `VITAL_PATTERNS`.
