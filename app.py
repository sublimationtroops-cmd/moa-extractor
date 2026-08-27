import streamlit as st
import pdfplumber, openpyxl, io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─── Constants ────────────────────────────────────────────────────────────────
SIZE_TOKENS = (
    'K4','K6','K8','K10','K12','K14','K16',
    'XXS','XS','S','M','L','XL',
    '2XL','3XL','4XL','5XL','6XL','7XL','8XL',
    'L6','L8','L10','L12','L14','L16','L18','L20','L22','L24','L26',
    'TO 0','TO 1','TO 2','TO 3','TO 4','TO 5','TO 6',
    'One'
)
SIZE_TOKENS_SET = set(SIZE_TOKENS)

LINE_COLORS = [
    'FF6D01','FFFF00','FF6D01','F6B3AE',
    'C2EF51','FFE599','ADD8E6','D8BFD8','90EE90','FFDAB9'
]

# ─── Core parsing helpers (unchanged logic) ───────────────────────────────────
def get_order_no(words):
    for i, w in enumerate(words):
        if w['text'] == 'Order' and i+1 < len(words) and words[i+1]['text'] in ('No', 'Number'):
            for j in range(i+2, min(i+8, len(words))):
                if words[j]['text'] in (':', ' ', 'No', 'Number'):
                    continue
                cleaned = words[j]['text'].replace(',', '')
                if len(cleaned) >= 4 and cleaned.isdigit():
                    return cleaned
    return ''

def get_line_no(words):
    for i, w in enumerate(words):
        if w['text'] == 'Line' and i+1 < len(words) and words[i+1]['text'] in ('No', 'Number'):
            for j in range(i+2, min(i+8, len(words))):
                val = words[j]['text'].replace(':', '').strip()
                if val.isdigit() and len(val) < 4:
                    return val
    return None

def merge_size_tokens(words):
    result = []
    i = 0
    while i < len(words):
        w = words[i]
        if w['text'].upper() == 'TO' and i + 1 < len(words):
            combined = f"TO {words[i+1]['text']}"
            if combined in SIZE_TOKENS_SET:
                merged = dict(w); merged['text'] = combined
                result.append(merged); i += 2; continue
        result.append(w); i += 1
    return result

def is_name_number_page(wbt, words):
    if not all(k in wbt for k in ('Name','Number','Qty','Size')):
        return False
    rows = defaultdict(list)
    for key in ('Name','Number','Qty','Size'):
        for w in wbt[key]:
            rows[round(w['top']/4)*4].append(w['text'])
    return any(all(t in rt for t in ('Name','Number','Qty','Size')) for rt in rows.values())

def is_names_grid_page(wbt):
    return all(k in wbt for k in ('Names','Sizes','Qunatities'))

def is_name_requirements_page(wbt):
    return all(k in wbt for k in ('Name','Requirements','Number','Size'))

def parse_names_grid_page(words, wbt):
    size_header_rows = {}
    for w in words:
        if w['text'] in SIZE_TOKENS_SET:
            size_header_rows.setdefault(round(w['top']/4)*4, []).append(w)
    if not size_header_rows:
        return []
    sorted_tops = sorted(size_header_rows)
    inner_top = next((ht for ht in sorted_tops if ht > 130), sorted_tops[-1])
    header_sizes = sorted(size_header_rows[inner_top], key=lambda w: w['x0'])
    total_words = [w for w in wbt.get('Total',[]) if abs(round(w['top']/4)*4 - inner_top) < 8]
    total_x = total_words[0]['x0'] if total_words else float('inf')
    excluded = {'Names','Sizes','Qunatities','Total','Printed','Page','of',
                ':','Type','Order','Details','Customer','ID','Ship','By',
                'Date','No','Line','AJAX','JFC'}
    name_words = [w for w in words
                  if w['x0'] < 80 and w['top'] > inner_top + 10
                  and w['text'] not in excluded
                  and not w['text'].replace(':','').strip().isdigit()
                  and w['text'] not in SIZE_TOKENS_SET
                  and len(w['text']) > 1
                  and not w['text'].startswith('2')
                  and ':' not in w['text']]
    name_rows = defaultdict(list)
    for w in name_words:
        name_rows[round(w['top']/4)*4].append(w)
    digit_by_row = defaultdict(list)
    for w in words:
        if w['text'].isdigit() and w['x0'] < total_x - 5:
            digit_by_row[round(w['top']/4)*4].append(w)
    entries = []
    for rk, rnw in name_rows.items():
        rnw_s = sorted(rnw, key=lambda w: w['x0'])
        name = ' '.join(w['text'] for w in rnw_s)
        name_top = rnw_s[0]['top']
        for dw in digit_by_row[round(name_top/4)*4]:
            nearest = min(header_sizes, key=lambda s: abs(s['x0']-dw['x0']))
            if abs(nearest['x0']-dw['x0']) < 40:
                try:
                    entries.append({'size':nearest['text'],'name':name,'number':None,'qty':int(dw['text'])})
                except ValueError:
                    pass
    return entries

def parse_name_number_page(words, wbt):
    header_words = [w for k in ('Size','Name','Number','Qty') for w in wbt.get(k,[])]
    if not header_words:
        return []
    header_top = min(w['top'] for w in header_words)
    left_h  = sorted([w for w in header_words if w['x0'] <  350 and abs(w['top']-header_top)<8], key=lambda w: w['x0'])
    right_h = sorted([w for w in header_words if w['x0'] >= 350 and abs(w['top']-header_top)<8], key=lambda w: w['x0'])
    skip = {'Printed',':','Page','of','Total'}

    def parse_group(headers, all_words, min_x, max_x):
        if len(headers) < 4: return []
        col_x = {h['text']: h['x0'] for h in headers}
        col_coords = sorted([(k, col_x[k]) for k in ('Size','Name','Number','Qty') if k in col_x], key=lambda x: x[1])
        data = [w for w in all_words if w['top']>header_top+5 and min_x<=w['x0']<=max_x and w['text'] not in skip]
        prox_rows = []
        for w in sorted(data, key=lambda w: w['top']):
            placed = False
            for pr in prox_rows:
                if abs(w['top'] - pr[0]['top']) <= 10:
                    pr.append(w); placed = True; break
            if not placed:
                prox_rows.append([w])
        entries = []
        for prox_row in prox_rows:
            rw = sorted(prox_row, key=lambda w: w['x0'])
            crd = defaultdict(list)
            for w in rw:
                col = min(col_coords, key=lambda c: abs(c[1]-w['x0']))
                if abs(col[1]-w['x0']) < 60:
                    crd[col[0]].append(w['text'])
            rd = {k: ' '.join(crd[k]).strip() if crd[k] else None for k in ('Size','Name','Number','Qty')}
            if rd['Size'] and rd['Size'] in SIZE_TOKENS_SET:
                try:
                    qty = int(rd['Qty']) if rd['Qty'] and rd['Qty'].isdigit() else 1
                    num = int(rd['Number']) if rd['Number'] and rd['Number'].isdigit() else None
                    entries.append({'size':rd['Size'],'name':rd['Name'],'number':num,'qty':qty})
                except ValueError:
                    pass
        return entries

    return parse_group(left_h, words, 0, 370) + parse_group(right_h, words, 370, 900)

def parse_name_requirements_page(words, wbt):
    header_words = [w for k in ('Number','Size','Name') for w in wbt.get(k,[])]
    if not header_words: return []
    rbt = defaultdict(list)
    for w in header_words: rbt[round(w['top']/4)*4].append(w)
    header_top = next((tk for tk, ws in rbt.items()
                       if all(t in [w['text'] for w in ws] for t in ('Number','Size','Name'))), None)
    if header_top is None: return []
    col_ws = [w for w in header_words if abs(round(w['top']/4)*4-header_top)<8]
    col_x  = {w['text']:w['x0'] for w in col_ws}
    num_x, size_x, name_x = col_x.get('Number',36), col_x.get('Size',90), col_x.get('Name',147)
    skip = {'Number','Size','Name','Requirements','Printed','Page','of','Total',':',
            'Customer','JAMBEROO','SUPEROOS','RLFC'}
    data = [w for w in words if w['top'] > header_top/1 + 8 and w['text'] not in skip]
    data_rows = defaultdict(list)
    for w in data: data_rows[round(w['top']/4)*4].append(w)
    entries = []
    for tk in sorted(data_rows):
        rw = sorted(data_rows[tk], key=lambda w: w['x0'])
        if any(t in [w['text'] for w in rw] for t in ('Printed','Apr','Page')): continue
        merged = []
        skip_next = False
        for idx, w in enumerate(rw):
            if skip_next: skip_next=False; continue
            if w['text']=='TO' and idx+1<len(rw):
                ct = f"TO {rw[idx+1]['text']}"
                if ct in SIZE_TOKENS_SET:
                    mw=dict(w); mw['text']=ct; merged.append(mw); skip_next=True; continue
            merged.append(w)
        num_val=size_val=None; name_parts=[]
        for w in merged:
            dists = [abs(w['x0']-num_x), abs(w['x0']-size_x), abs(w['x0']-name_x)]
            md = min(dists)
            if md > 80: continue
            if dists[0]==md: num_val=w['text']
            elif dists[1]==md: size_val=w['text']
            else: name_parts.append(w['text'])
        if not size_val or not num_val or size_val not in SIZE_TOKENS_SET: continue
        number = None if num_val in ('N/A','n/a','') else (int(num_val) if num_val.isdigit() else None)
        entries.append({'size':size_val,'name':' '.join(name_parts).strip() or None,'number':number,'qty':1})
    return entries

def process_page(page_num, page, total_pages):
    words_raw = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False, use_text_flow=False)
    if len(words_raw) < 5:
        return page_num, None
    words = merge_size_tokens(words_raw)
    wbt = defaultdict(list)
    for w in words: wbt[w['text']].append(w)
    return page_num, {'order_no': get_order_no(words), 'line_no': get_line_no(words), 'words': words, 'wbt': wbt}

def parse_po(pdf_bytes, progress_cb=None):
    order_no = ''
    pages_data = {}
    last_known_line_no = None

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)
        page_results = {}
        with ThreadPoolExecutor(max_workers=4) as ex:
            fmap = {ex.submit(process_page, n, p, total): n for n, p in enumerate(pdf.pages, 1)}
            done = 0
            for fut in as_completed(fmap):
                done += 1
                if progress_cb:
                    progress_cb(done, total)
                pn, res = fut.result()
                if res: page_results[pn] = res

        for pn in sorted(page_results):
            r = page_results[pn]
            words, wbt = r['words'], r['wbt']
            ol = r['order_no']
            if ol and len(ol) > len(order_no): order_no = ol
            line_no = r['line_no']
            if not line_no:
                if last_known_line_no: line_no = last_known_line_no
                else: continue
            else: last_known_line_no = line_no
            if line_no not in pages_data:
                pages_data[line_no] = {'has_roster':False,'sizes':[],'roster':[],'size_order':[],'type':'grid'}

            if is_name_number_page(wbt, words):
                pages_data[line_no]['type'] = 'name_table'
                pages_data[line_no]['roster'].extend(parse_name_number_page(words, wbt))
                pages_data[line_no]['has_roster'] = True
            elif is_name_requirements_page(wbt):
                pages_data[line_no]['type'] = 'name_requirements'
                pages_data[line_no]['roster'].extend(parse_name_requirements_page(words, wbt))
                pages_data[line_no]['has_roster'] = True
                if 'Sizes' in wbt:
                    _sw = [w for w in words if w['text'] in SIZE_TOKENS_SET]
                    if _sw:
                        _ht = min(w['top'] for w in _sw)
                        _hs = sorted([w for w in _sw if abs(w['top']-_ht)<6], key=lambda w: w['x0'])
                        _tc = [w for w in wbt.get('Total',[]) if abs(w['top']-_ht)<6]
                        _tx = _tc[0]['x0'] if _tc else float('inf')
                        _nw = [w for w in words if w['text'].isdigit() and int(w['text'])>0 and _ht+5<w['top']<_ht+60]
                        _rt = defaultdict(list)
                        for w in _nw: _rt[round(w['top']/4)*4].append(w)
                        _st = sorted(_rt)
                        if _st:
                            _sq = [min(_hs,key=lambda s:abs(s['x0']-nw['x0']))['text']
                                   for nw in _rt[_st[0]] if abs(nw['x0']-_tx)>=20
                                   and abs(min(_hs,key=lambda s:abs(s['x0']-nw['x0']))['x0']-nw['x0'])<40]
                            if _sq:
                                pages_data[line_no]['sizes'] = _sq
                                pages_data[line_no]['size_order'] = [s['text'] for s in _hs]
            elif is_names_grid_page(wbt):
                pages_data[line_no]['type'] = 'names_grid'
                pages_data[line_no]['roster'].extend(parse_names_grid_page(words, wbt))
                pages_data[line_no]['has_roster'] = True
            else:
                size_words = [w for w in words if w['text'] in SIZE_TOKENS_SET]
                if not size_words: continue
                has_sizes = 'Sizes' in wbt
                has_number = 'Number' in wbt
                if has_sizes:
                    ht = min(w['top'] for w in size_words)
                    hs = sorted([w for w in size_words if abs(w['top']-ht)<6], key=lambda w: w['x0'])
                    so = [s['text'] for s in hs]
                    nw = [w for w in words if w['text'].isdigit() and int(w['text'])>0 and ht+5<w['top']<ht+60]
                    rbt2 = defaultdict(list)
                    for w in nw: rbt2[round(w['top']/4)*4].append(w)
                    st = sorted(rbt2)
                    if not st: continue
                    tcw = [w for w in wbt.get('Total',[]) if abs(w['top']-ht)<6]
                    tx = tcw[0]['x0'] if tcw else float('inf')
                    sq = []
                    for nw2 in rbt2[st[0]]:
                        if abs(nw2['x0']-tx)<20: continue
                        near = min(hs, key=lambda s: abs(s['x0']-nw2['x0']))
                        if abs(near['x0']-nw2['x0'])<40: sq.append(near['text'])
                    pages_data[line_no]['sizes'] = sq
                    pages_data[line_no]['has_roster'] = has_number
                    pages_data[line_no]['size_order'] = so
                    if 'Numbering' in wbt:
                        _nl = wbt['Numbering']; _nt = min(w['top'] for w in _nl)
                        _nbs = [w for w in words if w['text'] in SIZE_TOKENS_SET and _nt-5<w['top']<_nt+20]
                        if _nbs:
                            _nbht = min(w['top'] for w in _nbs)
                            _nbhs = sorted([w for w in _nbs if abs(w['top']-_nbht)<6], key=lambda w: w['x0'])
                            _nbtc = [w for w in wbt.get('Total',[]) if abs(w['top']-_nbht)<6]
                            _nbtx = _nbtc[0]['x0'] if _nbtc else float('inf')
                            _nbfx = min(s['x0'] for s in _nbhs)
                            _nbtops = [w['top'] for w in wbt.get('Total',[]) if w['top']>_nbht+5]
                            _nbad = [w for w in words if w['text'].isdigit() and w['top']>_nbht+5]
                            def _on_tot(top): return any(abs(top-tt)<8 for tt in _nbtops)
                            _nbj  = [w for w in _nbad if w['x0']<_nbfx-5 and not _on_tot(w['top'])]
                            _nbq  = [w for w in _nbad if w['x0']>=_nbfx-5 and not _on_tot(w['top']) and abs(w['x0']-_nbtx)>15]
                            for _jw in _nbj:
                                _num = int(_jw['text'])
                                for _qw in _nbq:
                                    if abs(_qw['top'] - _jw['top']) <= 10:
                                        _nr = min(_nbhs, key=lambda s: abs(s['x0']-_qw['x0']))
                                        if abs(_nr['x0']-_qw['x0'])<40 and abs(_qw['x0']-_nbtx)>15:
                                            pages_data[line_no]['roster'].append({'size':_nr['text'],'name':None,'number':_num,'qty':int(_qw['text'])})
                            if not pages_data[line_no]['sizes'] and _nbhs:
                                pages_data[line_no]['sizes'] = [s['text'] for s in _nbhs]
                else:
                    ht  = min(w['top'] for w in size_words)
                    hs  = sorted([w for w in size_words if abs(w['top']-ht)<6], key=lambda w: w['x0'])
                    fsx = min(s['x0'] for s in hs)
                    total_tops = [w['top'] for w in wbt.get('Total',[]) if w['top']>ht]
                    tcw2= [w for w in wbt.get('Total',[]) if abs(w['top']-ht)<6]
                    tcx = tcw2[0]['x0'] if tcw2 else float('inf')
                    nw2 = [w for w in words if w['text'].isdigit() and w['top']>ht+5]
                    # Exclude words on Total rows (use proximity, not bucket)
                    def on_total_row(top):
                        return any(abs(top - tt) < 8 for tt in total_tops)
                    jn  = [w for w in nw2 if w['x0']<fsx-5 and not on_total_row(w['top'])]
                    qc  = [w for w in nw2 if w['x0']>=fsx-5 and not on_total_row(w['top']) and abs(w['x0']-tcx)>15]
                    # Use proximity row matching (tolerance=10px) instead of fixed bucket
                    for jw in jn:
                        num = int(jw['text'])
                        for qw in qc:
                            if abs(qw['top'] - jw['top']) <= 10:
                                nr = min(hs, key=lambda s: abs(s['x0']-qw['x0']))
                                if abs(nr['x0']-qw['x0'])<40 and abs(qw['x0']-tcx)>15:
                                    try: pages_data[line_no]['roster'].append({'size':nr['text'],'name':None,'number':num,'qty':int(qw['text'])})
                                    except ValueError: pass
                    if not pages_data[line_no].get('size_order'):
                        pages_data[line_no]['size_order'] = [s['text'] for s in hs]

    rows = []
    for ln in sorted(pages_data, key=int):
        d = pages_data[ln]
        label = f'{order_no} line {ln}'
        so = d.get('size_order',[])
        all_s = set(d['sizes']) | {r['size'] for r in d['roster']}
        ord_s = [s for s in so if s in all_s]
        for s in all_s:
            if s not in ord_s: ord_s.append(s)
        by_sz = defaultdict(list)
        for r in d['roster']: by_sz[r['size']].append(r)
        for sz in ord_s:
            rows.append({'line':ln,'label':label,'size':sz,'number':None,'name':None,'scope':'BASE DESIGN'})
            for r in by_sz.get(sz,[]):
                rows.append({'line':ln,'label':label,'size':sz,'number':r['number'],'name':r.get('name'),'scope':'ROSTER','qty':r.get('qty',1)})
    return order_no, rows

def build_excel(order_no_master, rows):
    unique_ids = []
    seen = set()
    for item in rows:
        uid = item.get('subf_number') or item['label'].split(' ')[0]
        if uid not in seen: unique_ids.append(uid); seen.add(uid)
    color_map = {uid: LINE_COLORS[i % len(LINE_COLORS)] for i, uid in enumerate(unique_ids)}
    rql = defaultdict(int)
    for item in rows:
        if item['scope']=='ROSTER': rql[item['label']] += item.get('qty',1)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hf = PatternFill('solid', start_color='002060')
    hfont = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    for col, h in enumerate(['SUBF NUMBER','LINE #','SIZE','NUMBERS','NAMES','SCOPE','VIS NAME'], 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font=hfont; c.fill=hf; c.alignment=Alignment(horizontal='center',vertical='center'); c.border=border
    for ri, item in enumerate(rows, 2):
        uid = item.get('subf_number') or item['label'].split(' ')[0]
        rf = PatternFill('solid', start_color=color_map[uid])
        font = Font(name='Arial', size=11)
        ncf = PatternFill('solid', start_color='FFFF00') if item['scope']=='ROSTER' and rql[item['label']]>33 else None
        for col, val in enumerate([item.get('subf_number',''),item['label'],item['size'],item['number'],item['name'],item['scope'],item.get('vis_name','')], 1):
            c = ws.cell(row=ri, column=col, value=val)
            c.font=font
            if col<=2: c.fill=rf
            if col in (4,5) and ncf: c.fill=ncf
            c.alignment=Alignment(horizontal='center',vertical='center'); c.border=border
    for col, w in zip('ABCDEFG', [15,22,10,12,18,16,25]):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

def get_subf_numerical_value(s):
    if s.upper().startswith('SUBF '):
        try: return int(s[5:])
        except: return float('inf')
    return float('inf')

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="MOA Count Sheet Extractor", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
@import url('https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css');

html, body, [class*="css"], .stApp {
    font-family: 'Inter', sans-serif !important;
    background: #13152e !important;
}
.block-container { padding: 0 !important; max-width: 100% !important; }
header[data-testid="stHeader"] { background: transparent !important; }
.stDeployButton, footer { display: none !important; }

/* Topbar */
.topbar {
    background: #1a1d3d;
    padding: 12px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid rgba(255,255,255,0.06);
}
.logo { font-size: 18px; font-weight: 800; color: #fff; letter-spacing: 0.05em; }
.logo-accent {
    background: linear-gradient(135deg,#6c63ff,#00e5a0);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
}
.nav-tabs { display: flex; gap: 6px; }
.ntab {
    background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.09);
    border-radius: 8px; padding: 6px 14px; font-size: 11px; font-weight: 600;
    color: rgba(255,255,255,0.4); cursor: pointer;
}
.ntab.active {
    color: #fff; border-color: rgba(108,99,255,0.5);
    background: rgba(108,99,255,0.15);
}

/* Welcome bar */
.welcomebar {
    background: #1a1d3d;
    padding: 12px 24px;
    display: flex; align-items: center; justify-content: space-between;
    border-bottom: 1px solid rgba(255,255,255,0.05);
    flex-wrap: wrap; gap: 10px;
}
.wleft { display: flex; align-items: center; gap: 12px; }
.avatar {
    width: 40px; height: 40px; border-radius: 12px;
    background: linear-gradient(135deg,#6c63ff,#00e5a0);
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.avatar i { color: #fff; font-size: 20px; }
.wtitle { font-size: 14px; font-weight: 700; color: #fff; }
.wsub { font-size: 11px; color: rgba(255,255,255,0.35); margin-top: 2px; }
.clocks { display: flex; gap: 8px; flex-wrap: wrap; }
.clock-card {
    background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.08);
    border-radius: 10px; padding: 7px 14px; text-align: center; min-width: 95px;
}
.clock-zone { font-size: 9px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 3px; }
.clock-time { font-size: 15px; font-weight: 700; color: #fff; line-height: 1; font-variant-numeric: tabular-nums; }
.clock-date { font-size: 9px; color: rgba(255,255,255,0.35); margin-top: 2px; }

/* Body */
.body-wrap { padding: 16px 20px; background: #13152e; }

/* Cards */
.card {
    background: #1e2146; border: 1px solid rgba(255,255,255,0.07);
    border-radius: 16px; padding: 18px;
}
.card-hd {
    font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.45);
    text-transform: uppercase; letter-spacing: 0.08em;
    margin-bottom: 14px; display: flex; align-items: center; justify-content: space-between;
}
.card-hd span { color: #fff; text-transform: none; letter-spacing: 0; font-size: 13px; font-weight: 700; }

/* Upload */
.upload-drop {
    border: 1.5px dashed rgba(108,99,255,0.4); border-radius: 12px;
    padding: 26px 16px; text-align: center; background: rgba(108,99,255,0.05);
    cursor: pointer; margin-bottom: 12px; transition: all 0.2s;
}
.up-ring {
    width: 50px; height: 50px; border-radius: 14px; margin: 0 auto 12px;
    display: flex; align-items: center; justify-content: center; font-size: 24px;
    background: linear-gradient(135deg,rgba(108,99,255,0.3),rgba(0,229,160,0.2));
    border: 1px solid rgba(108,99,255,0.3);
}
.up-ring i { color: #a78bfa; font-size: 24px; }
.up-title { font-size: 14px; font-weight: 600; color: #fff; margin-bottom: 4px; }
.up-sub { font-size: 11px; color: rgba(255,255,255,0.3); }
.subf-acc {
    background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.07);
    border-radius: 9px; padding: 10px 13px; display: flex; align-items: center;
    justify-content: space-between; font-size: 12px; color: rgba(255,255,255,0.4);
    cursor: pointer; margin-bottom: 12px;
}

/* Extract button */
.stButton > button {
    width: 100% !important;
    padding: 14px !important;
    border: none !important;
    border-radius: 12px !important;
    font-size: 14px !important;
    font-weight: 700 !important;
    color: #13152e !important;
    background: linear-gradient(135deg,#6c63ff,#00e5a0) !important;
    letter-spacing: 0.02em !important;
}
.stButton > button:hover { filter: brightness(1.08) !important; }

/* Stat mini cards */
.stat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 14px; }
.scard { background: #252854; border-radius: 10px; padding: 11px 13px; }
.scard-l { font-size: 10px; color: rgba(255,255,255,0.35); text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 5px; }
.scard-v { font-size: 18px; font-weight: 700; }
.g1 { background: linear-gradient(135deg,#6c63ff,#00b4d8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.g2 { background: linear-gradient(135deg,#f43f5e,#ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.g3 { background: linear-gradient(135deg,#fbbf24,#f97316); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.g4 { background: linear-gradient(135deg,#00e5a0,#00b4d8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.pill { display: inline-flex; padding: 2px 8px; border-radius: 100px; font-size: 10px; font-weight: 600; margin-top: 5px; }
.pill-g { background: rgba(0,229,160,0.12); color: #00e5a0; }
.pill-p { background: rgba(108,99,255,0.15); color: #a78bfa; }
.pill-o { background: rgba(251,191,36,0.12); color: #fbbf24; }

/* Bars */
.bars-wrap { display: flex; align-items: flex-end; gap: 5px; height: 70px; }
.bar-col { flex: 1; display: flex; flex-direction: column; align-items: center; gap: 4px; }
.bar-body { width: 100%; border-radius: 4px 4px 0 0; }
.bar-lbl { font-size: 9px; color: rgba(255,255,255,0.3); }

/* Donut */
.donut-row { display: flex; align-items: center; gap: 14px; margin-bottom: 14px; }
.donut { position: relative; width: 84px; height: 84px; flex-shrink: 0; }
.donut svg { transform: rotate(-90deg); }
.dc { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; font-size: 16px; font-weight: 700; color: #fff; line-height: 1.2; }
.dc small { font-size: 9px; color: rgba(255,255,255,0.4); font-weight: 400; }
.leg-item { display: flex; align-items: center; gap: 7px; font-size: 11px; color: rgba(255,255,255,0.55); margin-bottom: 6px; }
.leg-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

/* Progress bars */
.prog-item { margin-bottom: 10px; }
.prog-top { display: flex; justify-content: space-between; font-size: 11px; color: rgba(255,255,255,0.5); margin-bottom: 5px; }
.prog-top strong { color: #fff; font-weight: 600; }
.track { height: 6px; background: rgba(255,255,255,0.07); border-radius: 100px; overflow: hidden; }
.prog-fill { height: 100%; border-radius: 100px; }

/* Success result */
.result-box {
    background: rgba(0,229,160,0.08); border: 1px solid rgba(0,229,160,0.3);
    border-radius: 14px; padding: 18px; margin-top: 12px;
}
.result-title { font-size: 15px; font-weight: 700; color: #00e5a0; margin-bottom: 12px; text-align: center; }
.result-stats { display: grid; grid-template-columns: repeat(4,1fr); gap: 8px; }
.rs { background: rgba(0,0,0,0.2); border-radius: 8px; padding: 10px; text-align: center; }
.rs-val { font-size: 18px; font-weight: 700; color: #fff; display: block; }
.rs-lbl { font-size: 9px; color: rgba(255,255,255,0.35); text-transform: uppercase; letter-spacing: 0.06em; }

/* Footer */
.footer-bar {
    background: #1a1d3d; border-top: 1px solid rgba(255,255,255,0.05);
    padding: 10px 24px; display: flex; align-items: center; justify-content: space-between;
    font-size: 11px; color: rgba(255,255,255,0.25); margin-top: 6px;
}
.fl { display: flex; align-items: center; gap: 8px; color: rgba(255,255,255,0.4); font-weight: 600; }
.live-dot { width: 7px; height: 7px; background: #00e5a0; border-radius: 50%;
    display: inline-block; animation: gp 2s infinite; }
@keyframes gp { 0%,100%{box-shadow:0 0 4px #00e5a0} 50%{box-shadow:0 0 12px #00e5a0,0 0 24px rgba(0,229,160,0.3)} }

/* Streamlit overrides */
section[data-testid="stFileUploadDropzone"] {
    background: transparent !important; border: none !important;
}
.stTextArea textarea {
    background: rgba(255,255,255,0.05) !important; border: 1px solid rgba(255,255,255,0.1) !important;
    color: #fff !important; border-radius: 9px !important; font-size: 12px !important;
}
.stProgress > div > div { background: linear-gradient(90deg,#6c63ff,#00e5a0) !important; }
label, .stTextArea label { color: rgba(255,255,255,0.5) !important; font-size: 12px !important; }
.stDownloadButton > button {
    width: 100% !important; padding: 13px !important; border-radius: 12px !important;
    font-size: 14px !important; font-weight: 700 !important;
    background: linear-gradient(135deg,#00e5a0,#00b4d8) !important;
    color: #13152e !important; border: none !important;
}
</style>

<!-- Live clocks JS injected via component -->
""", unsafe_allow_html=True)

CLOCK_JS = """
<script>
function fmtT(d){return d.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true})}
function fmtD(d){return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})}
function tick(){
    var now=new Date();
    var ist=new Date(now.toLocaleString('en-US',{timeZone:'Asia/Kolkata'}));
    var fjt=new Date(now.toLocaleString('en-US',{timeZone:'Pacific/Fiji'}));
    var est=new Date(now.toLocaleString('en-US',{timeZone:'America/New_York'}));
    var e=document.getElementById.bind(document);
    if(e('ist-t')){e('ist-t').textContent=fmtT(ist);e('ist-d').textContent=fmtD(ist);}
    if(e('fjt-t')){e('fjt-t').textContent=fmtT(fjt);e('fjt-d').textContent=fmtD(fjt);}
    if(e('est-t')){e('est-t').textContent=fmtT(est);e('est-d').textContent=fmtD(est);}
}
tick();setInterval(tick,1000);
</script>
"""

# ── TOPBAR ──
st.markdown("""
<div class="topbar">
  <div class="logo">MOA <span class="logo-accent">Extractor</span></div>
  <div class="nav-tabs">
    <div class="ntab">Upload</div>
    <div class="ntab active">Dashboard</div>
    <div class="ntab">History</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── WELCOME + CLOCKS ──
st.markdown(f"""
<div class="welcomebar">
  <div class="wleft">
    <div class="avatar"><i class="ti ti-file-spreadsheet"></i></div>
    <div>
      <div class="wtitle">Hsenid — PO Count Sheet Extractor</div>
      <div class="wsub">Upload your PO PDF and get a formatted Excel instantly · No Google account needed</div>
    </div>
  </div>
  <div class="clocks">
    <div class="clock-card">
      <div class="clock-zone" style="color:#6c63ff">🇮🇳 IST</div>
      <div class="clock-time" id="ist-t">--:--:--</div>
      <div class="clock-date" id="ist-d">--</div>
    </div>
    <div class="clock-card">
      <div class="clock-zone" style="color:#00e5a0">🇫🇯 FJT</div>
      <div class="clock-time" id="fjt-t">--:--:--</div>
      <div class="clock-date" id="fjt-d">--</div>
    </div>
    <div class="clock-card">
      <div class="clock-zone" style="color:#ec4899">🇺🇸 EST</div>
      <div class="clock-time" id="est-t">--:--:--</div>
      <div class="clock-date" id="est-d">--</div>
    </div>
  </div>
</div>
{CLOCK_JS}
""", unsafe_allow_html=True)

# ── BODY ──
st.markdown('<div class="body-wrap">', unsafe_allow_html=True)

col1, col2, col3 = st.columns([1.1, 1, 1], gap="small")

# ── COL 1: UPLOAD ──
with col1:
    st.markdown("""
    <div class="card">
      <div class="card-hd"><span>Upload PO files</span><i class="ti ti-upload" style="color:#6c63ff;font-size:15px"></i></div>
      <div class="upload-drop">
        <div class="up-ring"><i class="ti ti-cloud-upload"></i></div>
        <div class="up-title">Drop PDF files here</div>
        <div class="up-sub">or use the button below · PDF only · up to 200MB</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Select PDF files", type=["pdf"], accept_multiple_files=True,
        label_visibility="collapsed"
    )

    st.markdown("""
    <div class="card" style="margin-top:10px">
      <div class="subf-acc">
        <span><i class="ti ti-hash" style="margin-right:6px;font-size:13px;vertical-align:-2px"></i>Optional: SUBF number mapping</span>
        <i class="ti ti-chevron-down" style="font-size:13px"></i>
      </div>
    </div>
    """, unsafe_allow_html=True)

    subf_text = st.text_area(
        "SUBF mapping", placeholder="SUBF 103566\t88251-24\nSUBF 103567\t88251-25",
        height=80, label_visibility="collapsed"
    )

    extract_btn = st.button("⚡  Extract to Excel", disabled=not uploaded_files, use_container_width=True)

# ── COL 2: STATS + BARS ──
with col2:
    st.markdown("""
    <div class="card">
      <div class="card-hd"><span>App capabilities</span></div>
      <div class="stat-grid">
        <div class="scard">
          <div class="scard-l">File limit</div>
          <div class="scard-v g1">200 MB</div>
          <div><span class="pill pill-p">PDF</span></div>
        </div>
        <div class="scard">
          <div class="scard-l">Upload mode</div>
          <div class="scard-v g2">Multi</div>
          <div><span class="pill pill-g">Batch</span></div>
        </div>
        <div class="scard">
          <div class="scard-l">Output</div>
          <div class="scard-v g3">Excel</div>
          <div><span class="pill pill-o">.xlsx</span></div>
        </div>
        <div class="scard">
          <div class="scard-l">Engine</div>
          <div class="scard-v g4">Fast</div>
          <div><span class="pill pill-g">Parallel</span></div>
        </div>
      </div>
      <div class="card-hd"><span>Sample — lines per order</span></div>
      <div class="bars-wrap">
        <div class="bar-col"><div class="bar-body" style="height:40px;background:linear-gradient(180deg,#6c63ff,#00b4d8)"></div><div class="bar-lbl">L2</div></div>
        <div class="bar-col"><div class="bar-body" style="height:55px;background:linear-gradient(180deg,#f43f5e,#ec4899)"></div><div class="bar-lbl">L3</div></div>
        <div class="bar-col"><div class="bar-body" style="height:70px;background:linear-gradient(180deg,#6c63ff,#00b4d8)"></div><div class="bar-lbl">L4</div></div>
        <div class="bar-col"><div class="bar-body" style="height:45px;background:linear-gradient(180deg,#fbbf24,#f97316)"></div><div class="bar-lbl">L5</div></div>
        <div class="bar-col"><div class="bar-body" style="height:60px;background:linear-gradient(180deg,#00e5a0,#00b4d8)"></div><div class="bar-lbl">L6</div></div>
        <div class="bar-col"><div class="bar-body" style="height:70px;background:linear-gradient(180deg,#6c63ff,#00b4d8)"></div><div class="bar-lbl">L9</div></div>
        <div class="bar-col"><div class="bar-body" style="height:50px;background:linear-gradient(180deg,#f43f5e,#ec4899)"></div><div class="bar-lbl">L11</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

# ── COL 3: DONUT + PROGRESS ──
with col3:
    st.markdown("""
    <div class="card">
      <div class="card-hd"><span>Roster breakdown</span></div>
      <div class="donut-row">
        <div class="donut">
          <svg width="84" height="84" viewBox="0 0 84 84">
            <circle cx="42" cy="42" r="32" fill="none" stroke="rgba(255,255,255,0.05)" stroke-width="12"/>
            <circle cx="42" cy="42" r="32" fill="none" stroke="url(#dg1)" stroke-width="12" stroke-dasharray="120 81" stroke-linecap="round"/>
            <circle cx="42" cy="42" r="32" fill="none" stroke="url(#dg2)" stroke-width="12" stroke-dasharray="81 120" stroke-dashoffset="-120" stroke-linecap="round"/>
            <defs>
              <linearGradient id="dg1" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" stop-color="#6c63ff"/><stop offset="100%" stop-color="#00e5a0"/></linearGradient>
              <linearGradient id="dg2" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" stop-color="#f43f5e"/><stop offset="100%" stop-color="#ec4899"/></linearGradient>
            </defs>
          </svg>
          <div class="dc">60%<small>roster</small></div>
        </div>
        <div>
          <div class="leg-item"><div class="leg-dot" style="background:linear-gradient(135deg,#6c63ff,#00e5a0)"></div>Roster entries (60%)</div>
          <div class="leg-item"><div class="leg-dot" style="background:linear-gradient(135deg,#f43f5e,#ec4899)"></div>Base design (40%)</div>
        </div>
      </div>
      <div class="card-hd"><span>Size group coverage</span></div>
      <div class="prog-item">
        <div class="prog-top"><strong>M / L / XL</strong><span>Adult standard</span></div>
        <div class="track"><div class="prog-fill" style="width:85%;background:linear-gradient(90deg,#6c63ff,#00e5a0)"></div></div>
      </div>
      <div class="prog-item">
        <div class="prog-top"><strong>2XL / 3XL+</strong><span>Extended</span></div>
        <div class="track"><div class="prog-fill" style="width:52%;background:linear-gradient(90deg,#f43f5e,#ec4899)"></div></div>
      </div>
      <div class="prog-item">
        <div class="prog-top"><strong>Kids K4–K16</strong><span>Junior</span></div>
        <div class="track"><div class="prog-fill" style="width:38%;background:linear-gradient(90deg,#fbbf24,#f97316)"></div></div>
      </div>
      <div class="prog-item">
        <div class="prog-top"><strong>Ladies L6–L26</strong><span>Ladies</span></div>
        <div class="track"><div class="prog-fill" style="width:28%;background:linear-gradient(90deg,#00e5a0,#00b4d8)"></div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ── EXTRACTION LOGIC ──
if extract_btn and uploaded_files:
    progress_bar = st.progress(0, text="Starting extraction...")
    status = st.empty()
    all_rows_by_key = defaultdict(list)
    all_order_numbers = []
    total_files = len(uploaded_files)

    for fi, uf in enumerate(uploaded_files):
        pdf_bytes = uf.read()
        def progress_cb(done, total, fi=fi, fname=uf.name):
            overall = int(((fi + done/max(total,1)) / total_files) * 90)
            progress_bar.progress(overall, text=f"Processing {fname} — page {done}/{total}")
        order_no, rows = parse_po(pdf_bytes, progress_cb=progress_cb)
        if order_no: all_order_numbers.append(order_no)
        for row in rows:
            key = (order_no, row['line'])
            all_rows_by_key[key].append(row)

    progress_bar.progress(92, text="Building Excel...")

    subf_num_map = {}
    subf_filter_keys = set()
    if subf_text.strip():
        for entry in subf_text.strip().split('\n'):
            entry = entry.strip()
            if '\t' in entry:
                subf_label, order_line = entry.split('\t', 1)
                subf_number = subf_label.strip()
                if not subf_number.upper().startswith('SUBF'):
                    subf_number = 'SUBF ' + subf_number
                if '-' in order_line:
                    order, line = order_line.split('-', 1)
                    key = (order.strip(), line.strip())
                    subf_num_map[key] = subf_number
                    subf_filter_keys.add(key)

    if subf_filter_keys:
        all_rows_by_key = {k: v for k, v in all_rows_by_key.items() if k in subf_filter_keys}

    if not all_rows_by_key:
        st.error("No data extracted. Check your PDF format.")
    else:
        if subf_filter_keys:
            keys_ordered = sorted(all_rows_by_key, key=lambda k: get_subf_numerical_value(subf_num_map.get(k,'')))
        else:
            keys_ordered = sorted(all_rows_by_key, key=lambda k: (k[0], int(k[1])))

        final_rows = []
        for key in keys_ordered:
            for row in all_rows_by_key[key]:
                row['subf_number'] = subf_num_map.get(key, '')
                final_rows.append(row)

        roster_counter = defaultdict(int)
        for item in final_rows:
            parts = [item.get('subf_number') or 'N/A_SUBF', item['size']]
            if item['scope'] == 'BASE DESIGN':
                parts.append('B')
            elif item['scope'] == 'ROSTER':
                k = (item.get('subf_number','N/A_SUBF'), item['size'])
                roster_counter[k] += 1
                parts.append(f'R{roster_counter[k]}')
            item['vis_name'] = '_'.join(parts)

        unique_orders = list(dict.fromkeys(all_order_numbers))
        if len(unique_orders) == 1: out_name = unique_orders[0]
        elif len(unique_orders) > 1: out_name = f'{unique_orders[0]}_and_{len(unique_orders)-1}_others'
        else: out_name = 'Combined'

        excel_bytes = build_excel(out_name, final_rows)
        output_filename = f'{out_name.replace(" ","_")}_count_sheet.xlsx'
        line_count = len(set(r['label'] for r in final_rows))
        roster_count = sum(r.get('qty',1) for r in final_rows if r['scope']=='ROSTER')

        progress_bar.progress(100, text="Done!")

        st.markdown(f"""
        <div class="result-box">
          <div class="result-title">✅ Excel Ready — {output_filename}</div>
          <div class="result-stats">
            <div class="rs"><span class="rs-val">{out_name}</span><span class="rs-lbl">Order</span></div>
            <div class="rs"><span class="rs-val">{line_count}</span><span class="rs-lbl">Lines</span></div>
            <div class="rs"><span class="rs-val">{len(final_rows)}</span><span class="rs-lbl">Rows</span></div>
            <div class="rs"><span class="rs-val">{roster_count}</span><span class="rs-lbl">Roster</span></div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="⬇️  Download Excel",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ── FOOTER ──
st.markdown("""
<div class="footer-bar">
  <div class="fl">
    <span class="live-dot"></span>
    Hsenid — MOA Count Sheet Extractor · moacountsheet.streamlit.app
  </div>
  <span>Powered by Streamlit · Files processed in memory · Never stored</span>
</div>
""", unsafe_allow_html=True)
