#!/usr/bin/env python3
"""
parse_lab_pdfs.py — Extractor automático de resultados de laboratorio desde PDFs

Uso:
    python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/"
    python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/" --output lab_results.yaml
    python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/" --summary

Dependencias:
    pip install pymupdf pyyaml

Descripción:
    Escanea todos los PDFs en el directorio especificado, extrae:
    - Datos del paciente (nombre, HC, fecha)
    - Resultados de laboratorio (parámetro, resultado, unidad, referencia, patológico)
    - Agrupados por fecha y categoría

    La salida en YAML puede pegarse en la sección lab_tables del config principal.
"""

import argparse
import sys
import re
from pathlib import Path
from collections import defaultdict

import yaml

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF no disponible. Instalar con: pip install pymupdf", file=sys.stderr)
    sys.exit(1)


# ─── Patrones de extracción ───────────────────────────────────────────────────
# Encabezado del reporte
PATIENT_RE   = re.compile(r'Paciente:\s*(.+)', re.IGNORECASE)
HC_RE        = re.compile(r'N[°º]\s*Historia[:\s]+(\d+)', re.IGNORECASE)
DATE_RE      = re.compile(r'Fecha de Ingreso[:\s]+(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
SERVICE_RE   = re.compile(r'Servicio[:\s]+(.+)', re.IGNORECASE)
REQUEST_RE   = re.compile(r'N[°º]\.\s*de petici[oó]n[:\s]+(\d+)', re.IGNORECASE)

# Categorías de laboratorio (separadores de sección)
CATEGORY_RE  = re.compile(
    r'^(HEMATOLOG[IÍ]A|BIOQU[IÍ]MICOS?|ELECTROLITOS|GASOMETR[IÍ]A|'
    r'HEMOSTASIA|INMUNOQUÍMICA|UROAN[AÁ]LISIS|MICROBIOLOG[IÍ]A|'
    r'ORINA|HISOPADO|SECRECION|TEJIDO|HEMOCULTIVO|CULTIVO|'
    r'COAGULACIÓN|INMUNOLOG[IÍ]A)\s*$',
    re.IGNORECASE
)

# Línea de resultado: Nombre  Valor  Unidad  Rango
# Acepta líneas con resultado numérico o texto corto
RESULT_RE = re.compile(
    r'^(.{5,50?}?)\s+'                  # nombre del parámetro
    r'([<>≤≥]?\s*[\d,\.]+\s*[↑↓\*]*|'  # resultado numérico
    r'(?:POSITIVO|NEGATIVO|PRESENTE|AUSENTE|SIN DESARROLLO|NO ES APARENTE|NORMAL|TRANSPARENTE|AMARILLO|'
    r'NEGATIVO PARA|SIN CRECIMIENTO|MENOR A|MAYOR A|NEG|POS).*?)'  # resultado textual
    r'\s+([\w/\^µ\(\)\.\s%³]{1,20})?'  # unidad (opcional)
    r'\s+([\d\.,\-\s––]+)?'             # rango de referencia (opcional)
    r'\s*(\*{1,2})?$',                  # marcador patológico (opcional)
    re.IGNORECASE
)

# Resultado de microbiología (freetext)
MICRO_RE = re.compile(
    r'(POSITIVO|NEGATIVO|SIN DESARROLLO|NEGATIVO PARA|SIN CRECIMIENTO BACTERIANO|'
    r'Microorganismo identificado[:\s]+.+)',
    re.IGNORECASE
)


def extract_pdf_text(pdf_path: Path) -> str:
    """Extrae el texto completo de un PDF."""
    try:
        doc = fitz.open(str(pdf_path))
        pages_text = []
        for page in doc:
            pages_text.append(page.get_text())
        doc.close()
        return '\n'.join(pages_text)
    except Exception as e:
        print(f"  Error leyendo {pdf_path.name}: {e}", file=sys.stderr)
        return ''


def parse_report_header(text: str) -> dict:
    """Extrae metadatos del encabezado del reporte."""
    meta = {}
    for line in text.splitlines():
        line = line.strip()
        m = PATIENT_RE.search(line)
        if m and 'patient' not in meta:
            meta['patient'] = m.group(1).strip()
        m = HC_RE.search(line)
        if m and 'medical_record' not in meta:
            meta['medical_record'] = m.group(1).strip()
        m = DATE_RE.search(line)
        if m and 'date' not in meta:
            meta['date'] = m.group(1).strip()
        m = REQUEST_RE.search(line)
        if m and 'request_id' not in meta:
            meta['request_id'] = m.group(1).strip()
        m = SERVICE_RE.search(line)
        if m and 'service' not in meta:
            meta['service'] = m.group(1).strip()
    return meta


def is_numeric_value(s: str) -> bool:
    """Verifica si una cadena es un valor numérico de laboratorio."""
    s = s.strip().rstrip('*').strip()
    # Remover prefijos comparativos
    s = re.sub(r'^[<>≤≥]\s*', '', s)
    # Reemplazar comas decimales
    s = s.replace(',', '.')
    try:
        float(s)
        return True
    except ValueError:
        return False


def is_reference_range(s: str) -> bool:
    """Verifica si una cadena parece un rango de referencia (ej. '4.32 - 10.42')."""
    s = s.strip()
    # Patrón: número - número
    return bool(re.match(r'^[<>≤≥]?\s*[\d,\.]+\s*[-–]\s*[\d,\.]+$', s) or
                re.match(r'^[<>≤≥]\s*[\d,\.]+$', s) or
                re.match(r'^[\d,\.]+\s*[-–]\s*[\d,\.]+\s*$', s))


def is_category_header(s: str) -> bool:
    """Verifica si la línea es un encabezado de categoría de laboratorio."""
    return bool(CATEGORY_RE.match(s.strip()))


def is_method_line(s: str) -> bool:
    """Verifica si la línea es un nombre de método analítico (ignorar)."""
    methods = ('citometría', 'fotometría', 'impedancia', 'electrodo', 'coagulometría',
               'inmunoturbidimetria', 'electroquimioluminiscencia', 'potenciometría',
               'kirby bauer', 'espectrofotometría', 'nefelometría')
    return s.strip().lower().startswith(methods)


SKIP_LINES = {
    'exámenes', 'resultado', 'unidad', 'intervalo de referencia', 'método',
    'fecha validación', 'fecha validadación', 'responsable', 'preinforme',
    'reimpresión de', 'resultados', 'pág:', 'de', '* marca de resultados',
    '** marca de resultados', 'dra.', 'lic.', 'informe de resultados',
    'n°. de petición:', 'paciente:', 'origen:', 'nº historia:', 'edad:',
    'servicio:', 'género:', 'médico:', 'fecha de ingreso:', 'años',
    'fecha nacimiento:', 'años'
}


def parse_lab_results(text: str) -> list:
    """
    Extrae resultados de laboratorio del texto del PDF.
    
    Los PDFs del laboratorio HEEE tienen layout de columnas donde cada campo
    aparece en una línea separada: nombre / resultado / unidad / referencia / *
    
    Devuelve lista de dicts: {category, name, result, unit, reference, pathological}
    """
    results = []
    current_category = 'GENERAL'
    lines = [l.strip() for l in text.splitlines()]

    # Filtrar líneas vacías y de encabezado
    clean_lines = []
    for ln in lines:
        if not ln:
            continue
        if any(ln.lower().startswith(skip) for skip in SKIP_LINES):
            continue
        if re.match(r'^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}$', ln):
            continue
        clean_lines.append(ln)

    i = 0
    while i < len(clean_lines):
        line = clean_lines[i]

        # Detectar cambio de categoría
        if is_category_header(line):
            current_category = re.sub(r'\s+', ' ', line.upper().strip())
            i += 1
            continue

        # Ignorar líneas de método analítico
        if is_method_line(line):
            i += 1
            continue

        # Detectar resultados microbiológicos
        micro_keywords = ('positivo', 'negativo', 'sin desarrollo', 'negativo para',
                          'sin crecimiento', 'microorganismo identificado',
                          'investigacion de cepa', 'no es aparente', 'no aparente')
        if any(line.lower().startswith(kw) or kw in line.lower() for kw in micro_keywords):
            results.append({
                'category': current_category,
                'name': 'Resultado microbiológico',
                'result': line[:300],
                'unit': '',
                'reference': '',
                'pathological': False,
                'type': 'microbiology'
            })
            i += 1
            continue

        # Intentar detectar nombre de parámetro seguido de valor numérico
        # Patrón multi-línea: nombre (línea i) → valor (línea i+1 o i+2)
        if not is_numeric_value(line) and not is_reference_range(line) and len(line) >= 3:
            param_name = line

            # Buscar el valor en las próximas líneas
            j = i + 1
            result_val = ''
            result_unit = ''
            result_ref = ''
            result_path = False

            while j < min(i + 6, len(clean_lines)):
                next_line = clean_lines[j]
                if is_category_header(next_line):
                    break
                if is_method_line(next_line):
                    j += 1
                    continue
                # Línea con valor + posible * al final
                stripped = next_line.rstrip('*').strip()
                if is_numeric_value(stripped) and not result_val:
                    result_val = stripped.replace(',', '.')
                    result_path = next_line.endswith('*') or next_line.endswith('**')
                    j += 1
                    # Siguiente línea puede ser unidad
                    if j < len(clean_lines):
                        unit_line = clean_lines[j]
                        if not is_numeric_value(unit_line) and \
                           not is_reference_range(unit_line) and \
                           not is_category_header(unit_line) and \
                           not is_method_line(unit_line) and \
                           len(unit_line) < 30 and \
                           not any(unit_line.lower().startswith(kw) for kw in micro_keywords):
                            result_unit = unit_line
                            j += 1
                    # Siguiente línea puede ser rango de referencia
                    if j < len(clean_lines):
                        ref_line = clean_lines[j]
                        if is_reference_range(ref_line):
                            result_ref = ref_line
                            j += 1
                    # Posible marcador de patológico suelto
                    if j < len(clean_lines) and clean_lines[j] in ('*', '**'):
                        result_path = True
                        j += 1
                    break
                j += 1

            if result_val and param_name and len(param_name) >= 3:
                results.append({
                    'category': current_category,
                    'name': param_name,
                    'result': result_val,
                    'unit': result_unit,
                    'reference': result_ref,
                    'pathological': result_path,
                    'type': 'numeric'
                })
                i = j
                continue

        i += 1

    return results


def group_by_date(all_reports: list) -> dict:
    """Agrupa reportes por fecha."""
    by_date = defaultdict(list)
    for report in all_reports:
        date = report.get('meta', {}).get('date', 'Fecha desconocida')
        by_date[date].append(report)
    return dict(sorted(by_date.items()))


def format_yaml_output(grouped: dict, include_microbiology: bool = True) -> dict:
    """
    Genera estructura YAML para insertar en el config del reporte de caso.
    Produce tablas por fecha con los parámetros extraídos.
    """
    output = {
        'lab_reports_summary': [],
        'suggested_lab_table': {
            'title_es': 'Evolución de laboratorios',
            'title_en': 'Laboratory evolution',
            'note': 'Tabla sugerida — revisar y ajustar valores antes de usar en el config principal',
            'dates': [],
            'parameters': []
        }
    }

    all_dates = sorted(grouped.keys())
    output['suggested_lab_table']['dates'] = all_dates

    # Recopilar todos los parámetros únicos
    all_params = {}
    for date, reports in grouped.items():
        for report in reports:
            for result in report.get('results', []):
                if result.get('type') == 'numeric':
                    name = result['name']
                    if name not in all_params:
                        all_params[name] = {
                            'unit': result['unit'],
                            'reference': result['reference'],
                            'category': result['category']
                        }

    # Construir tabla con una fila por parámetro y columna por fecha
    param_rows = []
    for param_name, param_info in all_params.items():
        row = {
            'parameter': param_name,
            'unit': param_info['unit'],
            'reference': param_info['reference'],
            'category': param_info['category'],
            'values': {}
        }
        for date, reports in grouped.items():
            for report in reports:
                for result in report.get('results', []):
                    if result.get('type') == 'numeric' and result['name'] == param_name:
                        flag = ' ↑' if result['pathological'] else ''
                        row['values'][date] = result['result'] + flag
        param_rows.append(row)

    output['suggested_lab_table']['parameters'] = param_rows

    # Resumen de microbiología
    micro_results = []
    for date, reports in grouped.items():
        for report in reports:
            request_id = report.get('meta', {}).get('request_id', '')
            for result in report.get('results', []):
                if result.get('type') == 'microbiology' and result['result'].strip():
                    micro_results.append({
                        'date': date,
                        'request_id': request_id,
                        'result': result['result'][:200]
                    })

    output['microbiology_findings'] = micro_results

    # Resumen por reporte
    for date, reports in grouped.items():
        for report in reports:
            meta = report.get('meta', {})
            n_results = len([r for r in report.get('results', []) if r.get('type') == 'numeric'])
            output['lab_reports_summary'].append({
                'date': date,
                'request_id': meta.get('request_id', ''),
                'patient': meta.get('patient', ''),
                'medical_record': meta.get('medical_record', ''),
                'service': meta.get('service', ''),
                'n_numeric_results': n_results,
                'file': report.get('file', '')
            })

    return output


def print_summary(grouped: dict):
    """Imprime resumen en consola."""
    total_pdfs = sum(len(r) for r in grouped.values())
    total_results = sum(len(r.get('results', [])) for reports in grouped.values() for r in reports)

    print(f"\n{'─'*60}")
    print(f"  RESUMEN DE EXTRACCIÓN DE LABORATORIOS")
    print(f"{'─'*60}")
    print(f"  PDFs procesados: {total_pdfs}")
    print(f"  Fechas únicas:   {len(grouped)}")
    print(f"  Resultados:      {total_results}")
    print()

    for date in sorted(grouped.keys()):
        reports = grouped[date]
        print(f"  📅 {date}:")
        for r in reports:
            meta = r.get('meta', {})
            n = len([x for x in r.get('results', []) if x.get('type') == 'numeric'])
            n_micro = len([x for x in r.get('results', []) if x.get('type') == 'microbiology'])
            print(f"     {r.get('file', '')} — {n} numéricos, {n_micro} microbiología")
        print()

    # Mostrar primeros parámetros
    print("  Parámetros numéricos encontrados (muestra):")
    seen = set()
    for reports in grouped.values():
        for report in reports:
            for result in report.get('results', []):
                if result.get('type') == 'numeric' and result['name'] not in seen:
                    seen.add(result['name'])
                    print(f"     • {result['name']:40s} [{result['unit']:15s}] ref: {result['reference']}")
                    if len(seen) >= 20:
                        break
            if len(seen) >= 20:
                break
        if len(seen) >= 20:
            break
    if len(seen) == 20:
        print("     ... (use --output para ver todos los resultados)")


def main():
    parser = argparse.ArgumentParser(
        description='Extrae resultados de laboratorio de PDFs para reportes de caso clínico',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  # Ver resumen en consola
  python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/"

  # Guardar en YAML para revisar e incorporar al config
  python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/" --output lab_data.yaml

  # Solo resumen sin output
  python3 parse_lab_pdfs.py "EXAMENES DE LABORATORIO/" --summary

Uso del output:
  El archivo YAML generado contiene una tabla sugerida con parámetros y valores
  por fecha. Copie y adapte la sección suggested_lab_table al config principal
  (case_config_template.yaml) en la sección postoperative.lab_evolution_table.
        """
    )
    parser.add_argument('pdf_dir', help='Directorio con los archivos PDF de laboratorio')
    parser.add_argument('--output', '-o', help='Archivo YAML de salida (default: stdout)')
    parser.add_argument('--summary', '-s', action='store_true',
                        help='Solo mostrar resumen en consola sin YAML completo')
    parser.add_argument('--pattern', default='Report*.pdf',
                        help='Patrón glob de archivos (default: Report*.pdf)')
    args = parser.parse_args()

    pdf_dir = Path(args.pdf_dir)
    if not pdf_dir.exists():
        print(f"Error: directorio no encontrado: {pdf_dir}", file=sys.stderr)
        sys.exit(1)

    # Buscar PDFs
    pdfs = sorted(pdf_dir.glob(args.pattern))
    if not pdfs:
        # Intentar con todos los PDFs
        pdfs = sorted(pdf_dir.glob('*.pdf'))

    if not pdfs:
        print(f"Error: no se encontraron archivos PDF en {pdf_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"Procesando {len(pdfs)} archivos PDF en {pdf_dir}...")

    # Procesar cada PDF
    all_reports = []
    for pdf_path in pdfs:
        text = extract_pdf_text(pdf_path)
        if not text.strip():
            print(f"  ⚠ Sin texto: {pdf_path.name}")
            continue

        meta = parse_report_header(text)
        results = parse_lab_results(text)

        all_reports.append({
            'file': pdf_path.name,
            'meta': meta,
            'results': results
        })
        print(f"  ✓ {pdf_path.name} — fecha: {meta.get('date', '?')} — {len(results)} resultados")

    if not all_reports:
        print("No se pudieron extraer datos de ningún PDF.", file=sys.stderr)
        sys.exit(1)

    # Agrupar por fecha
    grouped = group_by_date(all_reports)

    if args.summary:
        print_summary(grouped)
        return

    # Generar output YAML
    output_data = format_yaml_output(grouped)

    # Añadir instrucciones
    output_data['_instructions'] = (
        'Este archivo fue generado automáticamente por parse_lab_pdfs.py. '
        'Revise y edite los valores antes de incorporar al config principal. '
        'La sección suggested_lab_table puede adaptarse directamente a '
        'postoperative.lab_evolution_table en case_config_template.yaml.'
    )

    yaml_str = yaml.dump(output_data, allow_unicode=True, sort_keys=False,
                         default_flow_style=False, width=120)

    if args.output:
        output_path = Path(args.output)
        output_path.write_text(yaml_str, encoding='utf-8')
        print(f"\n✓ Datos guardados en: {output_path}")
        print_summary(grouped)
    else:
        print(f"\n{'─'*60}")
        print("  DATOS EXTRAÍDOS (YAML):")
        print("─'─'*60")
        print(yaml_str)
        print_summary(grouped)


if __name__ == '__main__':
    main()
