#!/usr/bin/env python3
"""
sync_crocoveen.py
─────────────────────────────────────────────────────────────────
Lee todos los archivos *Crocoveen*.xlsx de la carpeta donde vive
este script y actualiza en crocoveen.html los campos:
  • towers[].sold   (Vendidas)
  • S.rhythm        (ritmo mensual de ventas)
  • S.deliveries    (cronograma de entregas)

Ejecución:
    python3 sync_crocoveen.py              # usa la carpeta del script
    python3 sync_crocoveen.py /ruta/otra  # ruta explícita
"""

import openpyxl, glob, re, os, sys, json
from datetime import datetime

# ── 1. Configuración ────────────────────────────────────────────
FOLDER = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
HTML   = os.path.join(FOLDER, "crocoveen.html")

MO = {'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,
      'jul':7,'ago':8,'sep':9,'oct':10,'nov':11,'dic':12}

def norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

# ── 2. Tabla nombre-torre → ID de app ───────────────────────────
# Clave: (keyword_del_archivo, nombre_normalizado_de_torre)
TOWER_MAP = {
    # Country Life y Verano
    ('country','tgct1a'): 'C-T1A', ('country','tgct1b'): 'C-T1B',
    ('country','tgct2a'): 'C-T2A', ('country','tgct2b'): 'C-T2B',
    ('country','veranot1a'): 'V-T1A', ('country','veranot1b'): 'V-T1B',
    ('country','veranot2a'): 'V-T2A', ('country','veranot2b'): 'V-T2B',
    # Zitizen
    ('zitizen','t1'): 'Z-T1', ('zitizen','t2'): 'Z-T2',
    # Primavera
    ('primavera','t1'): 'P-T1', ('primavera','t2'): 'P-T2', ('primavera','t3'): 'P-T3',
    # Mágica (usa keyword 'gica' para cubrir la ó)
    ('gica','t2a'): 'M-T2A', ('gica','t2b'): 'M-T2B',
    ('gica','t3a'): 'M-T3A', ('gica','t3b'): 'M-T3B',
    ('gica','t4a'): 'M-T4A', ('gica','t4b'): 'M-T4B', ('gica','t4c'): 'M-T4C',
    # Bosketo
    ('bosketo','t1'): 'B-T1', ('bosketo','t2a'): 'B-T2A',
    ('bosketo','t2b'): 'B-T2B', ('bosketo','t3'): 'B-T3',
    # Camino Verde
    ('camino','t3'): 'CV-T3', ('camino','t4'): 'CV-T4',
    # Ambarte
    ('ambarte','t1a'): 'A-T1A', ('ambarte','t1b'): 'A-T1B', ('ambarte','t1c'): 'A-T1C',
    # La Vida es Bella (nombre normalizado empieza con 'lvb2t')
    ('bella','lvb2t1'): 'L-T1', ('bella','lvb2t2'): 'L-T2',
    ('bella','lvb2t3'): 'L-T3', ('bella','lvb2t4'): 'L-T4',
    ('bella','lvb2t5'): 'L-T5', ('bella','lvb2t6'): 'L-T6',
    # Summit – la hoja se llama UAU5
    ('uau5','t1'): 'S-T1', ('uau5','t2'): 'S-T2',
    ('uau5','t1s2'): 'S-T1S2', ('uau5','t1-s2'): 'S-T1S2',
    ('summit','t1'): 'S-T1', ('summit','t2'): 'S-T2',
    # Summit Grand
    ('summitgrand','t1'): 'SG-T1', ('summitgrand','t2'): 'SG-T2', ('summitgrand','t3'): 'SG-T3',
    ('grand','t1'): 'SG-T1', ('grand','t2'): 'SG-T2', ('grand','t3'): 'SG-T3',
}

NOT_TOWER = {'total','ventas','rph','torre','locales','condiciones',
             'cronentregas','planadepagos','actualizacion'}

def resolve_tower_id(file_keys, tower_name_raw):
    tn = norm(tower_name_raw)
    for fk in file_keys:
        hit = TOWER_MAP.get((fk, tn))
        if hit:
            return hit
    # Partial key match (exact normalized name)
    for (fk, tk), tid in TOWER_MAP.items():
        if tk == tn and any(fk in k for k in file_keys):
            return tid
    # Prefix match: handles names like 'LVB 2 T1\nARMONIA' → norm = 'lvb2t1armonia'
    # where the key is 'lvb2t1'. Only match if the char after the prefix is NOT a digit
    # (to avoid 'lvb2t10' matching key 'lvb2t1')
    for (fk, tk), tid in TOWER_MAP.items():
        if tn.startswith(tk) and any(fk in k for k in file_keys):
            suffix = tn[len(tk):]
            if not suffix or not suffix[0].isdigit():
                return tid
    return None

# ── 3. Mapeo columna → mes (YYYY-MM) ───────────────────────────
def build_col_month(hdr_vals, data_start_idx):
    """hdr_vals: tuple de 1 fila; data_start_idx: índice 0-based donde empiezan los meses."""
    col_month = {}
    i = data_start_idx
    while i < len(hdr_vals):
        v = hdr_vals[i]
        col = i + 1  # 1-based
        if v is None:
            i += 1
        elif isinstance(v, datetime):
            col_month[col] = v.strftime('%Y-%m')
            i += 1
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            i += 1  # número de año, saltar
        elif isinstance(v, str):
            clean = v.strip().replace(' ', '')
            # Formato 'Abr-Jun-26', 'Ene-Mar27', etc.
            m = re.match(r'([A-Za-z]{3})[-]?[A-Za-z]*[-]?(\d{2,4})', clean)
            if m:
                mo_str = m.group(1).lower()
                yr_str = m.group(2)
                year = int('20' + yr_str) if len(yr_str) == 2 else int(yr_str)
                if mo_str in MO:
                    mo = MO[mo_str]
                    for j in range(3):
                        cm = mo + j
                        cy = year
                        if cm > 12: cm -= 12; cy += 1
                        col_month[col + j] = f'{cy:04d}-{cm:02d}'
                    i += 3
                else:
                    i += 1
            else:
                i += 1
        else:
            i += 1
    return col_month

# ── 4. Parsear un archivo Excel ──────────────────────────────────
def parse_file(fpath):
    fname  = os.path.basename(fpath)
    fnorm  = norm(fname)
    results = []

    try:
        wb = openpyxl.load_workbook(fpath, data_only=True)
    except Exception as e:
        print(f"  [ERROR] No se pudo abrir {fname}: {e}")
        return []

    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    sheet_key = norm(sheet_name)

    # Keywords para lookup: del nombre de archivo y del nombre de hoja
    file_keys = []
    for kw in ['country','zitizen','primavera','gica','bosketo',
               'camino','ambarte','bella','uau5','summitgrand','grand','summit']:
        if kw in fnorm or kw in sheet_key:
            file_keys.append(kw)
    # Summit Grand antes de Summit para evitar falsos positivos
    if 'summitgrand' in file_keys and 'summit' in file_keys:
        file_keys = [k for k in file_keys if k != 'summit']

    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    # Encontrar fila de encabezado
    hdr_idx = tot_col = sold_col = disp_col = None
    for i, row in enumerate(rows[:10]):
        if 'Unidades' in row:
            hdr_idx = i
            for j, v in enumerate(row):
                if v == 'Unidades': tot_col  = j + 1
                if v == 'Vendidas': sold_col = j + 1
                if v == 'Disponibles': disp_col = j + 1
            break

    if hdr_idx is None or tot_col is None:
        print(f"  [SKIP] Sin encabezado Unidades/Vendidas: {fname}")
        return []

    # Mapa columna → mes
    col_month = build_col_month(rows[hdr_idx], disp_col)  # disp_col es 1-based, pasamos como 0-based index
    # Corregir: disp_col es 1-based, necesitamos 0-based para build_col_month
    col_month = build_col_month(rows[hdr_idx], disp_col)  # disp_col - 1 + 1 = disp_col en 1-based → pasar disp_col como idx 0-based
    # Fix the call:
    col_month = {}
    _hdr = rows[hdr_idx]
    _i = disp_col  # disp_col is 1-based, so index disp_col means column disp_col+1... let's redo cleanly
    # disp_col = column number (1-based) of 'Disponibles'
    # data starts at column disp_col+1, which is index disp_col in 0-based
    col_month = build_col_month(_hdr, disp_col)  # pass disp_col as 0-based start index → that's col disp_col+1 (1-based)

    def row_val(row, col_1based):
        idx = col_1based - 1
        return row[idx] if idx < len(row) else None

    def is_tower_row(row):
        tv = row_val(row, tot_col)
        sv = row_val(row, sold_col)
        if not (isinstance(tv, (int,float)) and not isinstance(tv, bool)): return False
        if not (isinstance(sv, (int,float)) and not isinstance(sv, bool)): return False
        if tv < 5 or tv > 1000: return False        # sanity: unit counts
        if sv < 0 or sv > tv + 5: return False
        return True

    def get_tower_name(row):
        """Busca el nombre de torre en columnas a la izquierda de tot_col."""
        candidates = []
        for j in range(min(tot_col - 1, 5)):
            v = row[j]
            if isinstance(v, str) and v.strip():
                n = norm(v)
                if n not in NOT_TOWER and len(n) >= 2:
                    candidates.append(v.strip())
        return candidates[-1] if candidates else None  # tomar el más cercano a los números

    def extract_rhythm(row):
        rh = {}
        for j, v in enumerate(row):
            col = j + 1
            if col in col_month and isinstance(v, (int,float)) and not isinstance(v,bool) and v > 0:
                rh[col_month[col]] = int(v)
        return rh

    for i, row in enumerate(rows):
        if i <= hdr_idx:
            continue
        if not is_tower_row(row):
            continue

        tname = get_tower_name(row)
        if not tname:
            continue

        tot  = int(row_val(row, tot_col))
        sold = int(row_val(row, sold_col))

        # Saltar filas TOTAL
        if norm(tname) in NOT_TOWER:
            continue

        # App ID
        app_id = resolve_tower_id(file_keys, tname)
        if not app_id:
            print(f"  ? Sin mapeo: archivo='{fname}' torre='{tname}' keys={file_keys}")
            continue

        # Ritmo de ventas (datos mensuales en la fila de torre)
        rhythm = extract_rhythm(row)

        # Buscar filas C/E (construcción/entregas) y Cron. entregas en las siguientes filas
        delivery = {}
        constr_cols = []
        e_cols = []
        # Scan up to 12 rows forward, stop at next tower row
        for k in range(1, 13):
            if i + k >= len(rows): break
            nrow = rows[i + k]
            if is_tower_row(nrow):
                break  # llegamos a la siguiente torre

            # ¿Es fila de marcadores C/E? (todas las celdas no-nulas son 'C', 'E' o 'V')
            ce_vals = [v for v in nrow if v is not None]
            if ce_vals and all(v in ('C', 'E', 'V') for v in ce_vals):
                for j, v in enumerate(nrow):
                    if v == 'C':
                        constr_cols.append(j + 1)  # columna 1-based
                    elif v == 'E':
                        e_cols.append(j + 1)       # columna E = lista para entrega

            # ¿Es fila de entregas?
            n2 = str(nrow[1]).strip() if len(nrow) > 1 and nrow[1] is not None else ''
            n3 = str(nrow[2]).strip() if len(nrow) > 2 and nrow[2] is not None else ''
            if 'Cron' in n2 or 'Cron' in n3 or 'entregas' in n2.lower() or 'entregas' in n3.lower():
                for j, v in enumerate(nrow):
                    col = j + 1
                    if col in col_month and isinstance(v,(int,float)) and not isinstance(v,bool) and v > 0:
                        delivery[col_month[col]] = int(v)

        # Fechas de construcción: meses C (obra) y E (lista para entrega)
        constr = None
        if constr_cols:
            c_months = sorted(col_month[c] for c in constr_cols if c in col_month)
            e_months = sorted(col_month[c] for c in e_cols if c in col_month)
            if c_months:
                constr = {'start': c_months[0], 'end': c_months[-1]}
                if e_months:
                    constr['eEnd'] = e_months[-1]

        parts = []
        if rhythm:   parts.append(f"ventas={sum(rhythm.values())}")
        if delivery: parts.append(f"entregas={sum(delivery.values())}")
        if constr:   parts.append(f"constr={constr['start']}→{constr['end']}")
        status_str = ' | '.join(parts) if parts else 'sin datos mensuales'

        print(f"  ✓ '{tname}' → {app_id}: sold={sold}/{tot} | {status_str}")
        results.append({'id': app_id, 'sold': sold, 'rhythm': rhythm,
                        'delivery': delivery, 'constr': constr})

    return results

# ── 5. Actualizar el HTML ───────────────────────────────────────
def update_html(all_results):
    if not os.path.exists(HTML):
        print(f"\n[ERROR] No se encontró {HTML}")
        print("Asegúrate de que crocoveen.html esté en la misma carpeta que este script.")
        return False

    with open(HTML, encoding='utf-8') as f:
        content = f.read()

    # Consolidar por tower ID (último archivo gana en caso de duplicados)
    data = {}
    for r in all_results:
        tid = r['id']
        if tid not in data:
            data[tid] = r
        else:
            # Fusionar
            if r['sold'] > 0 or r['rhythm'] or r['delivery']:
                data[tid] = r

    changed_sold = []
    changed_rhythm = []
    changed_del = []

    # 5a. Actualizar sold en towers array
    def replace_sold(m):
        tid = None
        # Buscar el id en la línea completa
        id_m = re.search(r"id:'([^']+)'", m.group(0))
        if id_m:
            tid = id_m.group(1)
        if tid and tid in data:
            new_sold = data[tid]['sold']
            changed_sold.append(f"{tid}:sold={new_sold}")
            return re.sub(r"sold:\d+", f"sold:{new_sold}", m.group(0))
        return m.group(0)

    content = re.sub(
        r"\{id:'[^']+',proj:'[^']*',name:'[^']*',tot:\d+,sold:\d+\}",
        replace_sold,
        content
    )

    def replace_js_block(txt, block_name, new_content_lines):
        """Reemplaza el bloque 'block_name:{...},' en el JS usando conteo de llaves."""
        marker = f'  {block_name}:{{'
        start = txt.find(marker)
        if start == -1:
            print(f"  [WARN] Bloque '{block_name}' no encontrado")
            return txt
        brace_pos = start + len(marker) - 1  # posición de la '{'
        depth = 0
        end = None
        for i in range(brace_pos, len(txt)):
            if txt[i] == '{':  depth += 1
            elif txt[i] == '}':
                depth -= 1
                if depth == 0:
                    end = i
                    break
        if end is None:
            print(f"  [WARN] No se encontró el cierre de '{block_name}'")
            return txt
        # Avanzar hasta después de la coma que sigue al bloque
        tail = end + 1
        while tail < len(txt) and txt[tail] in (' ', '\n', '\r', '\t'):
            tail += 1
        if tail < len(txt) and txt[tail] == ',':
            tail += 1
        new_block = f"  {block_name}:{{\n" + "\n".join(new_content_lines) + "\n  },"
        return txt[:start] + new_block + txt[tail:]

    # 5b. Reconstruir bloque rhythm:{...}
    new_rhythm_lines = []
    for tid, d in sorted(data.items()):
        if d['rhythm']:
            pairs = ','.join(f"'{k}':{v}" for k, v in sorted(d['rhythm'].items()))
            new_rhythm_lines.append(f"    '{tid}':{{{pairs}}},")
            changed_rhythm.append(tid)

    content = replace_js_block(content, 'rhythm', new_rhythm_lines)

    # 5c. Reconstruir bloque deliveries:{...}
    new_del_lines = []
    for tid, d in sorted(data.items()):
        if d['delivery']:
            pairs = ','.join(f"'{k}':{v}" for k, v in sorted(d['delivery'].items()))
            new_del_lines.append(f"    '{tid}':{{{pairs}}},")
            changed_del.append(tid)

    content = replace_js_block(content, 'deliveries', new_del_lines)

    # 5d. Reconstruir bloque constr:{...}
    # Estrategia: leer el constr actual del HTML, actualizar solo las torres
    # donde el Excel tiene marcadores C; el resto se mantiene igual.
    changed_constr = []

    # Extraer constr actual del HTML (antes de reemplazarlo)
    constr_start = content.find('  constr:{')
    brace_pos = constr_start + len('  constr:{') - 1
    depth = 0
    constr_end = None
    for i in range(brace_pos, len(content)):
        if content[i] == '{':  depth += 1
        elif content[i] == '}':
            depth -= 1
            if depth == 0:
                constr_end = i
                break

    existing_constr = {}
    if constr_end:
        block_text = content[constr_start:constr_end+1]
        # Parsear líneas como: 'C-T1A':{start:'2025-12',end:'2026-06'} o con eEnd opcional
        for m in re.finditer(r"'([^']+)':\{start:'([^']+)',end:'([^']+)'(?:,eEnd:'([^']+)')?\}", block_text):
            existing_constr[m.group(1)] = {'start': m.group(2), 'end': m.group(3)}
            if m.group(4):
                existing_constr[m.group(1)]['eEnd'] = m.group(4)

    # Mezclar: usar datos del Excel donde hay marcadores C; mantener el resto
    merged_constr = dict(existing_constr)
    for tid, d in data.items():
        if d.get('constr'):
            if existing_constr.get(tid) != d['constr']:
                changed_constr.append(
                    f"{tid}: {existing_constr.get(tid,{}).get('start','?')}→"
                    f"{existing_constr.get(tid,{}).get('end','?')}  →  "
                    f"{d['constr']['start']}→{d['constr']['end']}"
                )
            merged_constr[tid] = d['constr']

    new_constr_lines = []
    for tid in sorted(merged_constr):
        c = merged_constr[tid]
        epart = f",eEnd:'{c['eEnd']}'" if c.get('eEnd') else ''
        new_constr_lines.append(f"    '{tid}':" + "{" + f"start:'{c['start']}',end:'{c['end']}'{epart}" + "},")

    content = replace_js_block(content, 'constr', new_constr_lines)

    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"\n✅ HTML actualizado: {HTML}")
    print(f"   sold actualizado:    {len(changed_sold)} torres")
    print(f"   rhythm actualizado:  {len(changed_rhythm)} torres")
    print(f"   deliveries actualizado: {len(changed_del)} torres")
    print(f"   constr actualizado:  {len(changed_constr)} torres")
    if changed_constr:
        for c in changed_constr:
            print(f"     · {c}")
    return True

# ── 6. Main ─────────────────────────────────────────────────────
def main():
    print(f"📂 Carpeta: {FOLDER}")
    print(f"📄 HTML: {HTML}\n")

    xl_files = [
        f for f in glob.glob(os.path.join(FOLDER, "*.xlsx"))
        if 'Crocoveen' in os.path.basename(f) or 'crocoveen' in os.path.basename(f).lower()
    ]

    if not xl_files:
        print("[ERROR] No se encontraron archivos *Crocoveen*.xlsx")
        return

    all_results = []
    for fpath in sorted(xl_files):
        fname = os.path.basename(fpath)
        print(f"\n📊 {fname}")
        results = parse_file(fpath)
        all_results.extend(results)

    print(f"\n── Total: {len(all_results)} torres parseadas ──")
    update_html(all_results)
    push_to_github(FOLDER)

# ── 7. GitHub push ──────────────────────────────────────────────
def push_to_github(folder):
    """Sube crocoveen.html a GitHub Pages automáticamente."""
    import subprocess
    print("\n🚀 Subiendo a GitHub...")
    # Verificar que git esté instalado
    try:
        subprocess.run(['git', '--version'], capture_output=True, check=True)
    except FileNotFoundError:
        print("❌  Git no está instalado.")
        print("    Descárgalo en: https://git-scm.com/download/win")
        return
    # Verificar que la carpeta sea un repositorio git
    result = subprocess.run(['git', 'rev-parse', '--git-dir'],
                            capture_output=True, cwd=folder)
    if result.returncode != 0:
        print("⚠️   Esta carpeta no está conectada a GitHub todavía.")
        print("    Sigue la Guía de Configuración Inicial (GitHub_Setup.docx).")
        return
    # Agregar y commitear
    from datetime import datetime
    fecha = datetime.now().strftime('%Y-%m-%d')
    subprocess.run(['git', 'add', 'crocoveen.html'], cwd=folder, check=True)
    # ¿Hay cambios para subir?
    sin_cambios = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=folder)
    if sin_cambios.returncode == 0:
        print("✓  Sin cambios nuevos — GitHub ya está al día.")
        return
    try:
        subprocess.run(['git', 'commit', '-m', f'Sync {fecha}'], cwd=folder, check=True)
        subprocess.run(['git', 'push'], cwd=folder, check=True)
        print(f"✓  Publicado en GitHub ({fecha}) — el sitio se actualiza en ~1 minuto.")
    except subprocess.CalledProcessError:
        print("❌  Error al subir. Verifica tu conexión y que hayas hecho la configuración inicial.")

# ── 8. Main ─────────────────────────────────────────────────────
if __name__ == '__main__':
    main()
