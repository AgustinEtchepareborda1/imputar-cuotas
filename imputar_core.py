"""
Motor de imputación. Importable desde app.py (Streamlit) o scripts CLI.
No tiene side effects al importar.
"""

import io
import re
import datetime
import openpyxl
from openpyxl.styles import PatternFill

YELLOW_FILL = PatternFill(patternType='solid', fgColor='FFFF00')

MESES_ES = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo',  6: 'junio',   7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre',
}

SHEETS_BASE_PESOS = {
    'INDICE CAC': {
        'header_row': 5, 'data_start': 6,
        'cuit_col': 12, 'nombre_col': 9, 'lote_col': 8,
    },
    'INDICE CAC M. OBRA': {
        'header_row': 3, 'data_start': 4,
        'cuit_col': 11, 'nombre_col': 9, 'lote_col': 8,
    },
    'BOLSA CEMENTO': {
        'header_row': 3, 'data_start': 4,
        'cuit_col': 14, 'nombre_col': 10, 'lote_col': 9,
    },
}

SHEETS_BASE_USD = {
    '$  USD fijo': {
        'header_row': 3, 'data_start': 4,
        'cuit_col': 13, 'nombre_col': 10, 'lote_col': 9,
    },
}


def normalize_cuits(raw):
    if not raw:
        return []
    results = []
    for part in str(raw).split('/'):
        digits = re.sub(r'[^0-9]', '', part.strip())
        if len(digits) >= 10:
            results.append(digits)
    return results


def is_row_yellow(ws, row_num):
    for cell in ws[row_num]:
        if cell.value is None:
            continue
        fill = cell.fill
        if fill and fill.patternType and fill.patternType != 'none':
            rgb = str(fill.fgColor.rgb) if fill.fgColor else ''
            if 'FFFF00' in rgb or rgb == 'FFFFFF00':
                return True
        break
    return False


def extract_cuit_from_concepto(concepto):
    if not concepto:
        return None
    matches = re.findall(r'\b(\d{11,12})\b', str(concepto))
    return matches[0] if matches else None


def parse_date(val):
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, str):
        for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y'):
            try:
                return datetime.datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None


def _norm(s):
    return s.lower().replace('ó', 'o').replace('é', 'e').replace('á', 'a').replace('í', 'i').replace('ú', 'u')


def detectar_mes_transferencias(ws_imp, max_row=500):
    """Lee col A de imputaciones y retorna (year, month) más frecuente, o (None, None)."""
    from collections import Counter
    conteo = Counter()
    for row in ws_imp.iter_rows(min_row=4, max_row=max_row, max_col=1):
        dt = parse_date(row[0].value)
        if dt:
            conteo[(dt.year, dt.month)] += 1
    if not conteo:
        return None, None
    (year, month), _ = conteo.most_common(1)[0]
    return year, month


def detectar_columnas_mes(ws, header_row, year=None, month=None):
    """
    Retorna la columna 'teorico' para el mes/año indicado.
    Si year/month son None o no hay match, retorna la más a la derecha (fallback).
    """
    mes_nombre = _norm(MESES_ES.get(month, '')) if month else None
    yr_str = str(year % 100) if year else None       # "26" para 2026
    yr_full = str(year) if year else None             # "2026"

    teo_col_match = None
    teo_col_fallback = None
    for c in range(1, ws.max_column + 1):
        h = ws.cell(header_row, c).value
        if not h:
            continue
        h_norm = _norm(str(h))
        if 'teorico' not in h_norm:
            continue
        teo_col_fallback = c
        if mes_nombre and yr_str and mes_nombre in h_norm:
            if yr_str in h_norm or (yr_full and yr_full in h_norm):
                teo_col_match = c
    return teo_col_match if teo_col_match is not None else teo_col_fallback


def build_sheets_cfg(wb_deu_data, sheets_base, year=None, month=None):
    """Auto-detecta columnas del mes y construye sheets_cfg completo."""
    sheets_cfg = {}
    mes_info = {}
    for sheet_name, base in sheets_base.items():
        try:
            ws = wb_deu_data[sheet_name]
        except KeyError:
            continue
        teo_col = detectar_columnas_mes(ws, base['header_row'], year=year, month=month)
        if teo_col is None:
            continue
        cfg = dict(base)
        cfg['teo_col'] = teo_col
        cfg['real_col'] = teo_col + 1
        cfg['cuota_col'] = teo_col + 2
        cfg['fecha_col'] = teo_col + 3
        cfg['max_cuota_col'] = 1
        sheets_cfg[sheet_name] = cfg
        header_val = ws.cell(base['header_row'], teo_col).value
        mes_info[sheet_name] = str(header_val).strip() if header_val else f'col {teo_col}'
    return sheets_cfg, mes_info


def build_indices(wb_deu_data, sheets_cfg):
    cuit_index = {}
    nombre_index = {}
    cuota_history_cols = {}

    for sheet_name, cfg in sheets_cfg.items():
        ws_data = wb_deu_data[sheet_name]
        max_row = ws_data.max_row

        for r in range(cfg['data_start'], max_row + 1):
            nombre = ws_data.cell(r, cfg['nombre_col']).value
            cuit_raw = ws_data.cell(r, cfg['cuit_col']).value
            if not nombre and not cuit_raw:
                continue
            for c in normalize_cuits(cuit_raw):
                cuit_index.setdefault(c, []).append((sheet_name, r, nombre))

        for r in range(cfg['data_start'], max_row + 1):
            nombre = ws_data.cell(r, cfg['nombre_col']).value
            if not nombre:
                continue
            for palabra in str(nombre).upper().split():
                if len(palabra) >= 4:
                    nombre_index.setdefault(palabra, []).append((sheet_name, r, nombre))

        cols = []
        for c in range(1, ws_data.max_column + 1):
            h = ws_data.cell(cfg['header_row'], c).value
            if h and ('numero' in str(h).lower() or 'n°' in str(h).lower() or 'nro' in str(h).lower()) and 'cuota' in str(h).lower():
                if c != cfg['cuota_col']:
                    cols.append(c)
        cuota_history_cols[sheet_name] = cols

    return cuit_index, nombre_index, cuota_history_cols


def buscar_en_deudores_por_nombre(nombre_str, nombre_index):
    from collections import Counter
    palabras = [p.upper() for p in nombre_str.split() if len(p) >= 4]
    if not palabras:
        return []
    sets = [set((s, r) for s, r, _ in nombre_index.get(p, [])) for p in palabras]
    if not sets:
        return []
    comunes = sets[0]
    for s in sets[1:]:
        comunes &= s
    if not comunes:
        cnt = Counter()
        for p in palabras:
            for s, r, _ in nombre_index.get(p, []):
                cnt[(s, r)] += 1
        max_hits = max(cnt.values()) if cnt else 0
        comunes = {k for k, v in cnt.items() if v == max_hits and v >= max(2, len(palabras) - 1)}
    resultado = []
    seen = set()
    for p in palabras:
        for s, r, n in nombre_index.get(p, []):
            if (s, r) in comunes and (s, r) not in seen:
                resultado.append((s, r, n))
                seen.add((s, r))
    return resultado


def build_previo(wb_imp, imp_sheet, es_usd=False):
    all_sheets = wb_imp.sheetnames
    try:
        idx_actual = all_sheets.index(imp_sheet)
        previas = list(reversed(all_sheets[:idx_actual]))
        if es_usd:
            previas = [h for h in previas if 'usd' in h.lower()][:6]
        else:
            previas = previas[:8]
    except ValueError:
        previas = []

    cuit_to_nombre_previo = {}
    cuit_to_cuota_previo = {}
    for hoja in previas:
        try:
            ws_prev = wb_imp[hoja]
        except Exception:
            continue
        for row in ws_prev.iter_rows(min_row=4, max_row=ws_prev.max_row):
            concepto = row[2].value if len(row) > 2 else None
            col_h = row[7].value if len(row) > 7 else None
            if not concepto or not col_h:
                continue
            col_h_str = str(col_h).strip()
            if not col_h_str or 'PAGO MENOS' in col_h_str or 'Saldo' in col_h_str:
                continue
            cuit = extract_cuit_from_concepto(concepto)
            if cuit and cuit not in cuit_to_nombre_previo:
                m_cuota = re.search(r'\bc(\d+)\s*$', col_h_str, re.IGNORECASE)
                if m_cuota:
                    cuit_to_cuota_previo[cuit] = int(m_cuota.group(1))
                nombre_prev = re.sub(r'\s+c\d+\s*$', '', col_h_str, flags=re.IGNORECASE).strip()
                if nombre_prev:
                    cuit_to_nombre_previo[cuit] = nombre_prev

    return cuit_to_nombre_previo, cuit_to_cuota_previo


def procesar(
    imp_bytes, deu_bytes, imp_sheet,
    es_usd=False,
    tolerance=None,
    max_row=500,
    cuota_override=None,
    comprobantes_cache=None,
    log_fn=None,
):
    """
    Corre la imputación en modo simulación.
    Retorna (results, pago_menos, ambiguous, mes_info, sheets_cfg).
    """
    if tolerance is None:
        tolerance = 5 if es_usd else 3000
    if cuota_override is None:
        cuota_override = {}
    if comprobantes_cache is None:
        comprobantes_cache = {}
    if log_fn is None:
        log_fn = lambda msg: None

    sheets_base = SHEETS_BASE_USD if es_usd else SHEETS_BASE_PESOS

    # Solo data_only para leer — no cargamos el workbook editable acá
    wb_deu_data = openpyxl.load_workbook(io.BytesIO(deu_bytes), data_only=True)
    wb_imp = openpyxl.load_workbook(io.BytesIO(imp_bytes))

    tx_year, tx_month = detectar_mes_transferencias(wb_imp[imp_sheet], max_row)
    if tx_year:
        log_fn(f'Mes detectado en transferencias: {MESES_ES.get(tx_month, "?")} {tx_year}')

    sheets_cfg, mes_info = build_sheets_cfg(wb_deu_data, sheets_base, year=tx_year, month=tx_month)
    cuit_index, nombre_index, cuota_history_cols = build_indices(wb_deu_data, sheets_cfg)
    log_fn(f'{len(cuit_index)} CUITs indexados en deudores')

    cuit_to_nombre_previo, cuit_to_cuota_previo = build_previo(wb_imp, imp_sheet, es_usd)
    log_fn(f'{len(cuit_to_nombre_previo)} CUITs con nombre desde hojas anteriores')
    log_fn(f'{len(comprobantes_cache)} CUITs en cache de comprobantes')

    def _buscar_nombre(nombre_str):
        return buscar_en_deudores_por_nombre(nombre_str, nombre_index)

    ws_imp = wb_imp[imp_sheet]
    results = []
    ambiguous = []
    pago_menos = []
    written_deu_rows = set()

    for row in ws_imp.iter_rows(min_row=4, max_row=max_row):
        fecha_val = row[0].value
        monto_val = row[5].value
        concepto = row[2].value
        col_h_val = row[7].value

        if fecha_val is None and monto_val is None:
            continue

        row_num = row[0].row

        if is_row_yellow(ws_imp, row_num):
            continue

        if col_h_val and ('PAGO MENOS' in str(col_h_val) or 'Saldo Disponible' in str(col_h_val)):
            continue

        cuit_raw = extract_cuit_from_concepto(concepto)
        if not cuit_raw:
            ambiguous.append({'row': row_num, 'motivo': 'Sin CUIT extraíble', 'concepto': str(concepto)[:60] if concepto else '', 'monto': monto_val, 'fecha': fecha_val})
            continue

        matches = cuit_index.get(cuit_raw, [])
        if not matches:
            nombre_previo = cuit_to_nombre_previo.get(cuit_raw)
            if nombre_previo:
                m2 = _buscar_nombre(nombre_previo)
                if len(m2) == 1:
                    matches = m2
                    log_fn(f'Fila {row_num}: fallback nombre "{nombre_previo}" → {m2[0][2]}')
                elif len(m2) > 1:
                    ambiguous.append({'row': row_num, 'motivo': f'CUIT {cuit_raw} no en deudores; nombre "{nombre_previo}" da {len(m2)} candidatos', 'concepto': str(concepto)[:60], 'monto': monto_val, 'fecha': fecha_val, 'matches': [(n, None) for _, _, n in m2]})
                    continue
                else:
                    ambiguous.append({'row': row_num, 'motivo': f'CUIT {cuit_raw} no en deudores; "{nombre_previo}" no encontrado en deudores', 'concepto': str(concepto)[:60], 'monto': monto_val, 'fecha': fecha_val})
                    continue
            else:
                nombre_comp = comprobantes_cache.get(cuit_raw)
                if nombre_comp:
                    m2 = _buscar_nombre(nombre_comp)
                    if len(m2) == 1:
                        matches = m2
                        log_fn(f'Fila {row_num}: fallback comprobantes "{nombre_comp}" → {m2[0][2]}')
                    elif len(m2) > 1:
                        ambiguous.append({'row': row_num, 'motivo': f'CUIT {cuit_raw} en comprobantes como "{nombre_comp}"; da {len(m2)} candidatos', 'concepto': str(concepto)[:60], 'monto': monto_val, 'fecha': fecha_val, 'matches': [(n, None) for _, _, n in m2]})
                        continue
                    else:
                        ambiguous.append({'row': row_num, 'motivo': f'CUIT {cuit_raw} en comprobantes como "{nombre_comp}"; no encontrado en deudores', 'concepto': str(concepto)[:60], 'monto': monto_val, 'fecha': fecha_val})
                        continue
                else:
                    ambiguous.append({'row': row_num, 'motivo': f'CUIT {cuit_raw} no encontrado en deudores, semanas anteriores ni comprobantes', 'concepto': str(concepto)[:60], 'monto': monto_val, 'fecha': fecha_val})
                    continue

        if len(matches) > 1:
            candidatos = []
            for (sname, srow, snombre) in matches:
                cfg = sheets_cfg[sname]
                ws_d = wb_deu_data[sname]
                teo = ws_d.cell(srow, cfg['teo_col']).value
                real = ws_d.cell(srow, cfg['real_col']).value
                candidatos.append((sname, srow, snombre, teo, real))

            sin_imputar = [(s, r, n, t, rv) for s, r, n, t, rv in candidatos if rv is None]
            if not sin_imputar:
                ambiguous.append({'row': row_num, 'motivo': 'Todos los matches ya tienen el mes imputado', 'cuit': cuit_raw, 'monto': monto_val, 'matches': [(n, t) for _, _, n, t, _ in candidatos]})
                continue
            if len(sin_imputar) > 1:
                mejor = min(sin_imputar, key=lambda x: abs((x[3] or 0) - (monto_val or 0)))
                dif_mejor = abs((mejor[3] or 0) - (monto_val or 0))
                empate = [x for x in sin_imputar if x != mejor and abs(abs((x[3] or 0) - (monto_val or 0)) - dif_mejor) < (0.01 if es_usd else 1)]
                if empate:
                    disponibles = [x for x in sin_imputar if (x[0], x[1]) not in written_deu_rows]
                    if not disponibles:
                        ambiguous.append({'row': row_num, 'motivo': 'Todos los lotes ya asignados en este run', 'cuit': cuit_raw, 'monto': monto_val, 'matches': [(n, t) for _, _, n, t, _ in sin_imputar]})
                        continue
                    selected = sorted(disponibles, key=lambda x: (x[0], x[1]))[0]
                else:
                    selected = mejor
            else:
                selected = sin_imputar[0]
        else:
            sname, srow, snombre = matches[0]
            cfg = sheets_cfg[sname]
            ws_d = wb_deu_data[sname]
            teo = ws_d.cell(srow, cfg['teo_col']).value
            real = ws_d.cell(srow, cfg['real_col']).value
            selected = (sname, srow, snombre, teo, real)

        sname, srow, snombre, teo_val, real_existente = selected
        cfg = sheets_cfg[sname]

        if real_existente is not None or (sname, srow) in written_deu_rows:
            ambiguous.append({'row': row_num, 'motivo': f'Mes ya imputado ({real_existente}) o destino duplicado en {sname} fila {srow}', 'cliente': snombre, 'cuit': cuit_raw, 'monto': monto_val})
            continue

        monto_num = monto_val if isinstance(monto_val, (int, float)) else 0
        teo_num = teo_val if isinstance(teo_val, (int, float)) else 0

        if monto_num < teo_num and (teo_num - monto_num) > tolerance:
            pago_menos.append({'row': row_num, 'cliente': snombre, 'cuit': cuit_raw, 'transferido': monto_num, 'teorico': round(teo_num, 2 if es_usd else 0), 'diferencia': round(teo_num - monto_num, 2 if es_usd else 0)})
            continue

        ws_d = wb_deu_data[sname]
        cuota_col_val = ws_d.cell(srow, cfg['cuota_col']).value

        if isinstance(cuota_col_val, str) and 'parte' in cuota_col_val.lower():
            m = re.search(r'\d+', cuota_col_val)
            if m:
                next_cuota = int(m.group()) + 1
            else:
                ambiguous.append({'row': row_num, 'motivo': f'Cuota dice "parte de..." pero no se pudo extraer número ({cuota_col_val!r})', 'cliente': snombre, 'cuit': cuit_raw, 'hoja': sname, 'hoja_fila': srow})
                continue
        else:
            hist_cols = cuota_history_cols.get(sname, [])
            max_hist = None
            for hc in hist_cols:
                val = ws_d.cell(srow, hc).value
                if isinstance(val, (int, float)) and 0 < val <= 200:
                    if max_hist is None or val > max_hist:
                        max_hist = int(val)

            if max_hist is not None:
                next_cuota = max_hist + 1
            elif cuit_raw in cuit_to_cuota_previo:
                next_cuota = cuit_to_cuota_previo[cuit_raw] + 1
                log_fn(f'Fila {row_num}: cuota previa {cuit_to_cuota_previo[cuit_raw]}+1={next_cuota}')
            elif cuit_raw in cuota_override:
                next_cuota = cuota_override[cuit_raw]
                log_fn(f'Fila {row_num}: cuota override {next_cuota} para {snombre}')
            else:
                ambiguous.append({'row': row_num, 'motivo': 'No se pudo determinar número de cuota (sin historial)', 'cliente': snombre, 'cuit': cuit_raw, 'hoja': sname, 'hoja_fila': srow})
                continue

        fecha_dt = parse_date(fecha_val)

        todos_matches = cuit_index.get(cuit_raw, []) or _buscar_nombre(cuit_to_nombre_previo.get(cuit_raw, ''))
        nombres_distintos = {str(n).strip().upper() for _, _, n in todos_matches}
        if len(todos_matches) > 1 and len(nombres_distintos) == 1:
            lote_val = wb_deu_data[sname].cell(srow, cfg['lote_col']).value
            lote_str = f' l{lote_val}' if lote_val is not None else ''
        else:
            lote_str = ''

        written_deu_rows.add((sname, srow))
        results.append({
            'imp_row': row_num,
            'cuit': cuit_raw,
            'cliente': snombre,
            'lote_str': lote_str,
            'hoja': sname,
            'hoja_fila': srow,
            'monto_real': monto_num,
            'monto_teo': round(teo_num, 2 if es_usd else 0),
            'diferencia': round(monto_num - teo_num, 2 if es_usd else 0),
            'cuota': next_cuota,
            'fecha': fecha_dt,
        })

    wb_deu_data.close()
    wb_imp.close()

    return results, pago_menos, ambiguous, mes_info, sheets_cfg


def aplicar(results, pago_menos, imp_bytes, deu_bytes, imp_sheet, sheets_cfg):
    """Carga los workbooks desde bytes, escribe y retorna (imp_bytes, deu_bytes)."""
    wb_imp = openpyxl.load_workbook(io.BytesIO(imp_bytes))
    wb_deu_edit = openpyxl.load_workbook(io.BytesIO(deu_bytes))

    ws_edit = wb_imp[imp_sheet]

    for p in pago_menos:
        ws_edit.cell(p['row'], 8).value = 'PAGO MENOS'

    for r in results:
        row_num = r['imp_row']
        sname = r['hoja']
        srow = r['hoja_fila']
        cfg = sheets_cfg[sname]
        ws_deu = wb_deu_edit[sname]

        ws_deu.cell(srow, cfg['real_col']).value = r['monto_real']
        ws_deu.cell(srow, cfg['cuota_col']).value = r['cuota']
        ws_deu.cell(srow, cfg['fecha_col']).value = r['fecha']

        nombre_corto = str(r['cliente'])[:38]
        ws_edit.cell(row_num, 8).value = f"{nombre_corto}{r['lote_str']} c{r['cuota']}"
        for cell in ws_edit[row_num]:
            cell.fill = YELLOW_FILL

    imp_out = io.BytesIO()
    deu_out = io.BytesIO()
    wb_imp.save(imp_out)
    wb_deu_edit.save(deu_out)
    wb_imp.close()
    wb_deu_edit.close()
    return imp_out.getvalue(), deu_out.getvalue()
