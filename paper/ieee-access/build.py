import copy, re, os, shutil, zipfile, sys
KN_MAX = int(os.environ.get('KN_MAX', '18'))
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}
W = '{%s}' % NS['w']
R = '{%s}' % NS['r']
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

TPL = '/tmp/tpl'
SRC = '/tmp/sd'          # unzipped styled0.docx
OUT = '/tmp/build_access/pkg'
if os.path.exists(OUT):
    shutil.rmtree(OUT)
shutil.copytree(TPL, OUT)

# ---------- helpers ----------
def el(tag, **attrs):
    e = etree.Element(W + tag)
    for k, v in attrs.items():
        e.set(W + k, str(v))
    return e

def sub(parent, tag, **attrs):
    e = el(tag, **attrs)
    parent.append(e)
    return e

def para(style=None, jc=None, keep_next=False):
    p = el('p')
    ppr = sub(p, 'pPr')
    if style:
        sub(ppr, 'pStyle', val=style)
    if keep_next:
        sub(ppr, 'keepNext')
    if jc:
        sub(ppr, 'jc', val=jc)
    return p

def run(text, rpr=None):
    r = el('r')
    if rpr is not None:
        r.append(rpr)
    t = sub(r, 't')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r

def rpr(**flags):
    e = el('rPr')
    if flags.get('style'):
        sub(e, 'rStyle', val=flags['style'])
    if flags.get('b'):
        sub(e, 'b')
    if flags.get('i'):
        sub(e, 'i')
    if flags.get('sup'):
        sub(e, 'vertAlign', val='superscript')
    return e

def text_of(e):
    return ''.join(e.itertext())

def style_of(p):
    s = p.find('w:pPr/w:pStyle', NS)
    return s.get(W + 'val') if s is not None else None

# ---------- load sources ----------
src_doc = etree.parse(SRC + '/word/document.xml').getroot()
src_body = src_doc.find('w:body', NS)
src_rels = etree.parse(SRC + '/word/_rels/document.xml.rels').getroot()
src_rel = {r.get('Id'): r for r in src_rels}
src_num = etree.parse(SRC + '/word/numbering.xml').getroot()

# numbering format of the pandoc lists
def list_kind(numid):
    for n in src_num.findall('w:num', NS):
        if n.get(W + 'numId') == str(numid):
            a = n.find('w:abstractNumId', NS).get(W + 'val')
            for an in src_num.findall('w:abstractNum', NS):
                if an.get(W + 'abstractNumId') == a:
                    lvl = an.find('w:lvl', NS)
                    return lvl.find('w:numFmt', NS).get(W + 'val')
    return 'bullet'

tpl_doc_path = OUT + '/word/document.xml'
tpl = etree.parse(tpl_doc_path)
root = tpl.getroot()
body = root.find('w:body', NS)
tpl_children = list(body)
tpl_front = []   # template paragraphs 0..11 (index by position among w:p)
for ch in tpl_children:
    tpl_front.append(ch)

def find_tpl(style, k=0):
    ms = [c for c in tpl_children if c.tag == W + 'p' and style_of(c) == style]
    return ms[k]

# ---------- relationships / media ----------
tpl_rels_path = OUT + '/word/_rels/document.xml.rels'
tpl_rels = etree.parse(tpl_rels_path)
rr = tpl_rels.getroot()
rid_map = {}
counter = [0]

def map_rid(old):
    if old in rid_map:
        return rid_map[old]
    rel = src_rel[old]
    counter[0] += 1
    new = 'rIdAura%d' % counter[0]
    tgt = rel.get('Target')
    typ = rel.get('Type')
    nr = etree.SubElement(rr, '{%s}Relationship' % REL)
    nr.set('Id', new)
    nr.set('Type', typ)
    if rel.get('TargetMode') == 'External':
        nr.set('Target', tgt)
        nr.set('TargetMode', 'External')
    else:
        base = os.path.basename(tgt)
        newname = 'aura_' + base
        shutil.copy(os.path.join(SRC, 'word', tgt), os.path.join(OUT, 'word', 'media', newname))
        nr.set('Target', 'media/' + newname)
    rid_map[old] = new
    return new

def remap_rids(e):
    for n in e.iter():
        for att in list(n.attrib):
            if att.startswith(R):
                n.set(att, map_rid(n.get(att)))

# ---------- numbering additions (per-section A., B. lettering) ----------
num_path = OUT + '/word/numbering.xml'
numt = etree.parse(num_path)
nroot = numt.getroot()
ab = el('abstractNum', abstractNumId=900)
sub(ab, 'multiLevelType', val='singleLevel')
lvl = sub(ab, 'lvl', ilvl=0)
sub(lvl, 'start', val=1)
sub(lvl, 'numFmt', val='upperLetter')
sub(lvl, 'suff', val='space')
sub(lvl, 'lvlText', val='%1.')
sub(lvl, 'lvlJc', val='left')
ppr_ = sub(lvl, 'pPr')
sub(ppr_, 'ind', left=0, firstLine=0)
rp = sub(lvl, 'rPr')
sub(rp, 'color', val='58595B')
first_num = nroot.find('w:num', NS)
first_num.addprevious(ab)
next_num_id = [900]
def new_h2_list():
    next_num_id[0] += 1
    n = el('num', numId=next_num_id[0])
    sub(n, 'abstractNumId', val=900)
    lo = sub(n, 'lvlOverride', ilvl=0)
    sub(lo, 'startOverride', val=1)
    mac = nroot.find('w:numIdMacAtCleanup', NS)
    if mac is not None:
        mac.addprevious(n)
    else:
        nroot.append(n)
    return next_num_id[0]

# ---------- section properties ----------
def sectpr(ncols, first=False):
    s = el('sectPr')
    if first:
        sub(s, 'headerReference', type='default').set(R + 'id', 'rId12')
        sub(s, 'headerReference', type='even').set(R + 'id', 'rId12')
        sub(s, 'headerReference', type='first').set(R + 'id', 'rId9')
        sub(s, 'footerReference', type='default').set(R + 'id', 'rId10')
        sub(s, 'footerReference', type='even').set(R + 'id', 'rId13')
        sub(s, 'footerReference', type='first').set(R + 'id', 'rId10')
    else:
        sub(s, 'headerReference', type='default').set(R + 'id', 'rId12')
        sub(s, 'headerReference', type='even').set(R + 'id', 'rId12')
        sub(s, 'footerReference', type='default').set(R + 'id', 'rId10')
        sub(s, 'footerReference', type='even').set(R + 'id', 'rId13')
        sub(s, 'type', val='continuous')
    sub(s, 'pgSz', w=11520, h=15660, code=1)
    sub(s, 'pgMar', top=1300 if not first else 1280, right=740, bottom=1040, left=740, header=360, footer=640, gutter=0)
    if ncols == 2:
        sub(s, 'cols', num=2, space=400)
    else:
        sub(s, 'cols', space=720)
    if first:
        sub(s, 'titlePg')
    sub(s, 'docGrid', linePitch=360)
    return s

# ---------- front matter ----------
new_children = []
# DOP, DOI, title, AU, affiliations, abstract, index terms
new_children.append(copy.deepcopy(find_tpl('DOP')))
new_children.append(copy.deepcopy(find_tpl('DOI')))

title_text = text_of(src_body[0]).strip()
p = para('PaperTitle')
p.append(run(title_text))
new_children.append(p)

authors = [
    ('Radha Mohan Pattanayak', 'Associate Professor, SCOPE, VIT-AP University, Amaravati, India', 'radhamohan.pattanayak@vitap.ac.in'),
    ('Soubhagya Swastik Dhal', 'B.Tech (Computer Science & Engineering), VIT-AP University, Amaravati, India', 'soubhagya3696@gmail.com'),
    ('Basant Kumar', 'B.Tech (Computer Science & Engineering), VIT-AP University, Amaravati, India', 'gbasant201036@gmail.com'),
    ('Mehal Mathur', 'B.Tech (Computer Science & Engineering), VIT-AP University, Amaravati, India', 'mehalmathur24@gmail.com'),
    ('Aditya Kumar Jena', 'B.Tech (Computer Science & Engineering), VIT-AP University, Amaravati, India', 'adi.jena1904@gmail.com'),
    ('Sachi Nandan Mohanty', 'Professor Grade 2, SCOPE, VIT-AP University, Amaravati, India', 'sachinandan.m@vitap.ac.in'),
]
au = para('AU')
au.find('w:pPr', NS).append(el('spacing', after=0))
for k, (n, aff, mail) in enumerate(authors):
    pre = ''
    if k == len(authors) - 1:
        pre = 'and '
    elif k > 0:
        pre = ''
    au.append(run(('' if k == 0 else ' ') + pre + n))
    au.append(run(str(k + 1), rpr(sup=True)))
    if k < len(authors) - 1:
        au.append(run(','))
new_children.append(au)

def small(sz):
    e = el('rPr')
    sub(e, 'sz', val=sz)
    sub(e, 'szCs', val=sz)
    return e

for k, (n, aff, mail) in enumerate(authors):
    p = para('PINoSpace')
    ppr = p.find('w:pPr', NS)
    sub(ppr, 'ind', firstLine=0)
    e = el('rPr'); sub(e, 'sz', val=14); sub(e, 'szCs', val=14); sub(e, 'vertAlign', val='superscript')
    p.append(run(str(k + 1), e))
    p.append(run('%s (e-mail: %s)' % (aff, mail), small(14)))
    new_children.append(p)

p = para('PI')
p.append(run('Corresponding author: Radha Mohan Pattanayak (e-mail: radhamohan.pattanayak@vitap.ac.in).'))
new_children.append(p)

# abstract
abs_src = src_body[3]
abs_p = copy.deepcopy(find_tpl('Abstract'))
# keep first two runs (label + space), drop the rest
runs = abs_p.findall('w:r', NS)
for r_ in runs[2:]:
    abs_p.remove(r_)
for c in list(abs_p):
    if c.tag not in (W + 'pPr', W + 'r'):
        abs_p.remove(c)
txt_runs = []
for r_ in abs_src.findall('w:r', NS):
    rp_ = r_.find('w:rPr', NS)
    t = ''.join(x.text or '' for x in r_.findall('w:t', NS))
    if not t:
        continue
    it = rp_ is not None and rp_.find('w:i', NS) is not None
    txt_runs.append((t, it))
full = ''.join(t for t, _ in txt_runs)
assert full.startswith('Abstract—')
cut = len('Abstract—')
for t, it in txt_runs:
    if cut:
        if len(t) <= cut:
            cut -= len(t)
            continue
        t = t[cut:]
        cut = 0
    abs_p.append(run(t, rpr(i=True) if it else None))
new_children.append(abs_p)

# index terms
kw_src = text_of(src_body[4])
assert kw_src.startswith('Index Terms—')
kws = [x.strip().rstrip('.') for x in kw_src[len('Index Terms—'):].split(',')]
kws = sorted(kws, key=lambda s: s.lower())
it_p = copy.deepcopy(find_tpl('IT'))
runs = it_p.findall('w:r', NS)
for r_ in runs[2:]:
    it_p.remove(r_)
for c in list(it_p):
    if c.tag not in (W + 'pPr', W + 'r'):
        it_p.remove(c)
it_p.append(run(', '.join(kws)))
new_children.append(it_p)

# section break paragraph (single-column front matter)
brk = para('IT')
brk.find('w:pPr', NS).append(sectpr(1, first=True))
new_children.append(brk)

# ---------- body ----------
ROMAN = re.compile(r'^[IVXL]+\.\s*')
LETTER = re.compile(r'^[A-Z]\.\s*')

def clean_runs(p, drop=('sz', 'szCs', 'smallCaps', 'b', 'bCs')):
    for rp_ in p.iter(W + 'rPr'):
        for c in list(rp_):
            if c.tag.replace(W, '') in drop:
                rp_.remove(c)

def fix_verbatim(p):
    for r_ in p.iter(W + 'r'):
        rp_ = r_.find('w:rPr', NS)
        if rp_ is None:
            continue
        rs = rp_.find('w:rStyle', NS)
        if rs is not None and rs.get(W + 'val') == 'VerbatimChar':
            rp_.remove(rs)
            rf = el('rFonts', ascii='Courier New', hAnsi='Courier New', cs='Courier New')
            rp_.insert(0, rf)
            sz = el('sz', val=17); szc = el('szCs', val=17)
            rp_.append(sz); rp_.append(szc)
    # normalise child order of rPr: rStyle, rFonts, b, i, ..., sz, szCs, vertAlign
    order = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'color', 'sz', 'szCs', 'vertAlign']
    for rp_ in p.iter(W + 'rPr'):
        kids = list(rp_)
        kids.sort(key=lambda c: order.index(c.tag.replace(W, '')) if c.tag.replace(W, '') in order else 99)
        for c in kids:
            rp_.remove(c)
        seen = set()
        for c in kids:
            t_ = c.tag
            if t_ in seen and t_ != W + 'rFonts':
                continue
            seen.add(t_)
            rp_.append(c)

def convert_sect(p):
    sp = p.find('w:pPr/w:sectPr', NS)
    cols = sp.find('w:cols', NS)
    n = int(cols.get(W + 'num', '1'))
    ppr = p.find('w:pPr', NS)
    ppr.remove(sp)
    # spacer paragraph carrying the section break: no style (tiny)
    for c in list(ppr):
        ppr.remove(c)
    sub(ppr, 'spacing', before=0, after=0, line=20, lineRule='exact')
    rp_ = el('rPr'); sub(rp_, 'sz', val=2); sub(rp_, 'szCs', val=2)
    ppr.append(rp_)
    ppr.append(sectpr(n))
    for c in list(p):
        if c.tag != W + 'pPr':
            p.remove(c)
    return p

def marker(n):
    p = el('p'); ppr = sub(p, 'pPr')
    sub(ppr, 'spacing', before=0, after=0, line=20, lineRule='exact')
    rp_ = el('rPr'); sub(rp_, 'sz', val=2); sub(rp_, 'szCs', val=2)
    ppr.append(rp_)
    ppr.append(sectpr(n))
    return p

single_blocks = []
state = {'app': False, 'auto_wide': False, 'prev': 'heading', 'h2': None, 'h1_has_h2': False, 'h1_count': 0, 'first_h1': True,
         'in_refs': False, 'refs_started': False}
list_counts = {}
IMG_W_WIDE = int(6.9 * 914400)
IMG_W_COL = int(3.3 * 914400)

body_items = list(src_body)[5:]
next_cols = {}
for i_, it_ in enumerate(body_items):
    if it_.tag == W + 'p' and text_of(it_).startswith('TABLE '):
        for j_ in range(i_ + 1, len(body_items)):
            if body_items[j_].tag == W + 'tbl':
                next_cols[i_] = len(body_items[j_].findall('w:tblGrid/w:gridCol', NS))
                break

# which section markers to drop (tables are placed by my own width rule; Fig. 1 goes in one column)
def _is_marker(it):
    return it.tag == W + 'p' and it.find('w:pPr/w:sectPr', NS) is not None
def _ncols(it):
    return int(it.find('w:pPr/w:sectPr/w:cols', NS).get(W + 'num', '1'))
skip_markers = set()
first_img_idx = None
for i_, it_ in enumerate(body_items):
    if it_.tag == W + 'p' and it_.find('.//w:drawing', NS) is not None:
        first_img_idx = i_
        break
i_ = 0
while i_ < len(body_items):
    it_ = body_items[i_]
    if _is_marker(it_) and _ncols(it_) == 2:
        j_ = i_ + 1
        while j_ < len(body_items) and not (_is_marker(body_items[j_]) and _ncols(body_items[j_]) == 1):
            j_ += 1
        has_fig = any(b.tag == W + 'p' and b.find('.//w:drawing', NS) is not None for b in body_items[i_:j_])
        is_fig1 = first_img_idx is not None and i_ < first_img_idx < j_
        if is_fig1 or not has_fig:
            skip_markers.add(i_); skip_markers.add(j_)
        i_ = j_ + 1
    else:
        i_ += 1

def table_is_wide(tb):
    rows_ = tb.findall('w:tr', NS)
    n = len(tb.findall('w:tblGrid/w:gridCol', NS))
    if n <= 1:
        return False
    lens = [0] * n
    for tr in rows_:
        for k_, tc in enumerate(tr.findall('w:tc', NS)):
            if k_ < n:
                lens[k_] = max(lens[k_], max((len(x) for x in text_of(tc).split('\n')), default=0))
    cap = 22
    est = sum(min(l, cap) for l in lens) * 0.5 * 7 / 72 + 0.09 * n
    return est > 3.3

wide_cap = {}
for i_, it_ in enumerate(body_items):
    if it_.tag == W + 'p' and text_of(it_).startswith('TABLE '):
        for j_ in range(i_ + 1, len(body_items)):
            if body_items[j_].tag == W + 'tbl':
                wide_cap[i_] = table_is_wide(body_items[j_])
                for pre_ in os.environ.get('FORCE_SINGLE', 'TABLE XXPer,TABLE XXVIIIAl,TABLE XXXIReac').split(','):
                    if pre_ and text_of(it_).startswith(pre_):
                        wide_cap[i_] = False
                for pre_ in os.environ.get('FORCE_WIDE', '').split(','):
                    if pre_ and text_of(it_).startswith(pre_):
                        wide_cap[i_] = True
                print('table', text_of(it_)[:14], 'wide' if wide_cap[i_] else 'single')
                break

wide = False

def h1(text, first=False, unnumbered=False, appendix=False):
    p = para('H1' if (unnumbered or appendix) else ('H1ListNoSpace' if first else 'H1ListSpace'), keep_next=True)
    p.append(run(text.upper()))
    return p

for idx_, ch in enumerate(body_items):
    ch_idx = idx_
    tag = ch.tag.split('}')[1]
    if tag in ('bookmarkStart', 'bookmarkEnd'):
        new_children.append(copy.deepcopy(ch))
        continue
    if tag == 'tbl':
        t = copy.deepcopy(ch)
        remap_rids(t)
        tp = t.find('w:tblPr', NS)
        for b in tp.findall('w:tblBorders', NS):
            tp.remove(b)
        ncols = len(t.findall('w:tblGrid/w:gridCol', NS))
        rows = t.findall('w:tr', NS)
        onecol = ncols == 1
        # borders in template style
        b = el('tblBorders')
        if onecol:
            for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                sub(b, e, val='single', sz=6, space=0, color='808080')
        else:
            for e in ('top', 'bottom'):
                sub(b, e, val='single', sz=12, space=0, color='808080')
        # tblBorders must come before tblLayout/tblLook in schema: tblStyle,tblpPr,tblOverlap,bidiVisual,tblStyleRowBandSize,tblStyleColBandSize,tblW,jc,tblCellSpacing,tblInd,tblBorders,shd,tblLayout,tblCellMar,tblLook
        pos = None
        for i, c in enumerate(tp):
            if c.tag in (W + 'shd', W + 'tblLayout', W + 'tblCellMar', W + 'tblLook'):
                pos = i
                break
        if pos is None:
            tp.append(b)
        else:
            tp.insert(pos, b)
        order_ = ['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription']
        kids_ = list(tp)
        for c_ in kids_:
            tp.remove(c_)
        seen_ = set()
        for c_ in sorted(kids_, key=lambda c: order_.index(c.tag.replace(W,'')) if c.tag.replace(W,'') in order_ else 99):
            if c_.tag in seen_:
                continue
            seen_.add(c_.tag)
            tp.append(c_)
        # header row borders (template: double top, single bottom)
        if rows and not onecol and len(rows) > 1:
            for tc in rows[0].findall('w:tc', NS):
                tcpr = tc.find('w:tcPr', NS)
                if tcpr is None:
                    tcpr = el('tcPr')
                    tc.insert(0, tcpr)
                for old in tcpr.findall('w:tcBorders', NS):
                    tcpr.remove(old)
                tb = el('tcBorders')
                sub(tb, 'top', val='double', sz=6, space=0, color='auto')
                sub(tb, 'bottom', val='single', sz=6, space=0, color='auto')
                # schema order in tcPr: tcW, gridSpan, vMerge, tcBorders, shd, noWrap, tcMar...
                idx = 0
                for i, c in enumerate(tcpr):
                    if c.tag in (W + 'tcW', W + 'gridSpan', W + 'vMerge', W + 'hMerge'):
                        idx = i + 1
                tcpr.insert(idx, tb)
        tcorder = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark']
        for tc_ in t.iter(W + 'tcPr'):
            kids_ = list(tc_)
            for c_ in kids_:
                tc_.remove(c_)
            seen_ = set()
            for c_ in sorted(kids_, key=lambda c: tcorder.index(c.tag.replace(W,'')) if c.tag.replace(W,'') in tcorder else 99):
                if c_.tag in seen_:
                    continue
                seen_.add(c_.tag)
                tc_.append(c_)
        # cell paragraphs: drop pandoc styles
        for cp in t.iter(W + 'p'):
            ppr = cp.find('w:pPr', NS)
            if ppr is None:
                ppr = el('pPr')
                cp.insert(0, ppr)
            for s in ppr.findall('w:pStyle', NS):
                ppr.remove(s)
            for s in ppr.findall('w:keepNext', NS):
                ppr.remove(s)
            fix_verbatim(cp)
        tbl_wide = state['auto_wide']
        for t_el in t.iter(W + 't'):
            if t_el.text and re.search(r'\d-\d\(\w+\)\d-\d', t_el.text):
                t_el.text = re.sub(r'\d-\d\(\w+\)\d-\d\s*', '', t_el.text)
        avail = 10000 if tbl_wide else 4800
        gcols = t.findall('w:tblGrid/w:gridCol', NS)
        if len(gcols) > 1:
            lens_ = [0] * len(gcols)
            for tr in rows:
                k_ = 0
                for tc in tr.findall('w:tc', NS):
                    gs = tc.find('w:tcPr/w:gridSpan', NS)
                    span = int(gs.get(W + 'val')) if gs is not None else 1
                    if span == 1 and k_ < len(lens_):
                        lens_[k_] = max(lens_[k_], max((len(x) for x in text_of(tc).split('\n')), default=0))
                    k_ += span
            wts = [min(max(l, 3), 22) + 2 for l in lens_]
            tot = sum(wts)
            widths = [int(avail * w_ / tot) for w_ in wts]
            for g_, w_ in zip(gcols, widths):
                g_.set(W + 'w', str(w_))
            for tr in rows:
                k_ = 0
                for tc in tr.findall('w:tc', NS):
                    gs = tc.find('w:tcPr/w:gridSpan', NS)
                    span = int(gs.get(W + 'val')) if gs is not None else 1
                    tcw = tc.find('w:tcPr/w:tcW', NS)
                    if tcw is None:
                        tcpr_ = tc.find('w:tcPr', NS)
                        if tcpr_ is None:
                            tcpr_ = el('tcPr'); tc.insert(0, tcpr_)
                        tcw = el('tcW'); tcpr_.insert(0, tcw)
                    tcw.set(W + 'type', 'dxa')
                    tcw.set(W + 'w', str(sum(widths[k_:k_ + span])))
                    k_ += span
            tw = tp.find('w:tblW', NS)
            tw.set(W + 'type', 'dxa'); tw.set(W + 'w', str(sum(widths)))
            tl = tp.find('w:tblLayout', NS)
            if tl is None:
                tl = el('tblLayout'); tp.append(tl)
            tl.set(W + 'type', 'fixed')
            # re-sort tblPr
            kids_ = list(tp)
            for c_ in kids_:
                tp.remove(c_)
            for c_ in sorted(kids_, key=lambda c: order_.index(c.tag.replace(W,'')) if c.tag.replace(W,'') in order_ else 99):
                tp.append(c_)
        fsz = 16 if tbl_wide else (int(os.environ.get('APPSZ', '13')) if state['app'] else 14)
        for rp_ in t.iter(W + 'rPr'):
            for c_ in list(rp_):
                if c_.tag in (W + 'sz', W + 'szCs'):
                    rp_.remove(c_)
            rp_.append(el('sz', val=fsz)); rp_.append(el('szCs', val=fsz))
        for cp in t.iter(W + 'p'):
            hasrun = cp.find('w:r', NS) is not None
            ppr = cp.find('w:pPr', NS)
            if ppr.find('w:rPr', NS) is None:
                ppr.append(el('rPr'))
            for c_ in list(ppr.find('w:rPr', NS)):
                if c_.tag in (W + 'sz', W + 'szCs'):
                    ppr.find('w:rPr', NS).remove(c_)
            ppr.find('w:rPr', NS).append(el('sz', val=fsz)); ppr.find('w:rPr', NS).append(el('szCs', val=fsz))
        if state['app'] and not tbl_wide:
            for cp in t.iter(W + 'p'):
                ppr = cp.find('w:pPr', NS)
                for sp_ in ppr.findall('w:spacing', NS):
                    ppr.remove(sp_)
                spc = el('spacing', before=0, after=0, line=int(os.environ.get('APPLINE', '140')), lineRule='exact')
                # spacing goes after keepNext/numPr/tabs, before ind/jc/rPr
                pos_ = 0
                for k_, c_ in enumerate(ppr):
                    if c_.tag in (W + 'pStyle', W + 'keepNext', W + 'numPr', W + 'tabs'):
                        pos_ = k_ + 1
                ppr.insert(pos_, spc)
        for ri, tr in enumerate(rows):
            trpr = tr.find('w:trPr', NS)
            if trpr is None:
                trpr = el('trPr')
                tr.insert(0, trpr)
                # trPr must follow tblPrEx if present
            for c_ in list(trpr):
                if c_.tag in (W + 'cantSplit', W + 'tblHeader'):
                    trpr.remove(c_)
            trpr.append(el('cantSplit'))
        if len(gcols) == 1:
            gcols[0].set(W + 'w', str(avail))
            for tcw_ in t.iter(W + 'tcW'):
                tcw_.set(W + 'type', 'dxa'); tcw_.set(W + 'w', str(avail))
            tw_ = tp.find('w:tblW', NS)
            tw_.set(W + 'type', 'dxa'); tw_.set(W + 'w', str(avail))
        # wrap caption + table in a one-row, non-splitting outer table so the pair never breaks across a page/column
        outer = el('tbl')
        otp = sub(outer, 'tblPr')
        sub(otp, 'tblW', w=avail, type='dxa')
        sub(otp, 'jc', val='center')
        sub(otp, 'tblLayout', type='fixed')
        omar = sub(otp, 'tblCellMar')
        for side in ('top', 'left', 'bottom', 'right'):
            sub(omar, side, w=0, type='dxa')
        sub(otp, 'tblLook', val='0000')
        og = sub(outer, 'tblGrid')
        sub(og, 'gridCol', w=avail)
        otr = sub(outer, 'tr')
        otrp = sub(otr, 'trPr')
        sub(otrp, 'cantSplit')
        otc = sub(otr, 'tc')
        otcp = sub(otc, 'tcPr')
        sub(otcp, 'tcW', w=avail, type='dxa')
        for cp_ in state.get('caps', []):
            otc.append(cp_)
        state['caps'] = []
        otc.append(t)
        tail = para(None)
        tp_ = tail.find('w:pPr', NS)
        sub(tp_, 'spacing', before=0, after=0, line=20, lineRule='exact')
        trp = sub(tp_, 'rPr'); sub(trp, 'sz', val=2); sub(trp, 'szCs', val=2)
        otc.append(tail)
        new_children.append(outer)
        if not state['auto_wide'] and not state['app']:
            single_blocks.append(outer)
        if state['auto_wide']:
            new_children.append(marker(1))
            state['auto_wide'] = False
            wide = False
        state['prev'] = 'table'
        continue
    if tag != 'p':
        continue
    p = ch
    st = style_of(p)
    txt = text_of(p)
    has_sect = p.find('w:pPr/w:sectPr', NS) is not None
    has_img = p.find('.//w:drawing', NS) is not None
    if has_sect and ch_idx in skip_markers:
        continue
    if has_sect:
        cols = p.find('w:pPr/w:sectPr/w:cols', NS)
        n = int(cols.get(W + 'num', '1'))
        # the break paragraph with 2 cols precedes wide content; 1 col ends wide content
        wide = (n == 2)  # after a 2-col break marker the next content is wide
        new_children.append(convert_sect(copy.deepcopy(p)))
        continue
    if has_img:
        np_ = para('FigureCaption', jc='center', keep_next=True)
        for r_ in p.findall('w:r', NS):
            r2 = copy.deepcopy(r_)
            remap_rids(r2)
            if ch_idx == first_img_idx:
                shutil.copy('/tmp/figv/fig1_policies_v.png', OUT + '/word/media/aura_fig1_v.png')
                counter[0] += 1
                nid = 'rIdAura%d' % counter[0]
                nr = etree.SubElement(rr, '{%s}Relationship' % REL)
                nr.set('Id', nid)
                nr.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
                nr.set('Target', 'media/aura_fig1_v.png')
                for bl in r2.iter('{%s}blip' % NS['a']):
                    bl.set(R + 'embed', nid)
                for ext in r2.iter('{%s}extent' % NS['wp']):
                    ext.set('cx', '1007'); ext.set('cy', '1678')
                for e in r2.iter('{%s}ext' % NS['a']):
                    if e.get('cx'):
                        e.set('cx', '1007'); e.set('cy', '1678')
            for ext in r2.iter('{%s}extent' % NS['wp']):
                cx = int(ext.get('cx')); cy = int(ext.get('cy'))
                nw = IMG_W_WIDE if wide else IMG_W_COL
                ext.set('cx', str(nw)); ext.set('cy', str(int(cy * nw / cx)))
            for e in r2.iter('{%s}ext' % NS['a']):
                if e.get('cx'):
                    cx = int(e.get('cx')); cy = int(e.get('cy'))
                    nw = IMG_W_WIDE if wide else IMG_W_COL
                    e.set('cx', str(nw)); e.set('cy', str(int(cy * nw / cx)))
            np_.append(r2)
        new_children.append(np_)
        state['prev'] = 'img'
        continue
    # ----- headings -----
    if st == 'Heading1':
        t = txt.strip()
        # appendix headings contain a line break
        if t.startswith('Appendix') or t.startswith('Data and Code'):
            state['app'] = True
        if t.startswith('Appendix'):
            m = re.match(r'Appendix ([A-Z])\s*(.*)', t, re.S)
            letter, ttl = m.group(1), m.group(2).strip()
            new_children.append(h1('Appendix ' + letter, appendix=True))
            p2 = h1(ttl, appendix=True)
            p2.find('w:pPr', NS).append(el('spacing', before=0))
            new_children.append(p2)
            state['h2'] = None
        elif t.startswith('Data and Code'):
            new_children.append(h1(t, unnumbered=True))
            state['h2'] = None
        else:
            t = ROMAN.sub('', t)
            new_children.append(h1(t, first=state['first_h1']))
            state['first_h1'] = False
            state['h2'] = None
        state['prev'] = 'heading'
        continue
    if st == 'Heading2':
        t = LETTER.sub('', txt.strip())
        if state['h2'] is None:
            state['h2'] = new_h2_list()
            sty = 'H2First'
        else:
            sty = 'H2Cont'
        p2 = para(sty, keep_next=True)
        npr = el('numPr'); sub(npr, 'ilvl', val=0); sub(npr, 'numId', val=state['h2'])
        p2.find('w:pPr', NS).append(npr)
        p2.append(run(t.upper()))
        new_children.append(p2)
        state['prev'] = 'heading'
        continue
    # ----- bibliography -----
    if st == 'Bibliography':
        if not state['refs_started']:
            state['refs_started'] = True
            new_children.append(h1('References', unnumbered=True))
        np_ = para('References')
        rs = p.findall('w:r', NS)
        for r_ in rs[3:]:
            r2 = copy.deepcopy(r_)
            np_.append(r2)
        for c in p:
            if c.tag in (W + 'hyperlink',):
                np_.append(copy.deepcopy(c))
        remap_rids(np_)
        fix_verbatim(np_)
        new_children.append(np_)
        continue
    # ----- table title -----
    if txt.startswith('TABLE ') and st == 'BodyText':
        nc_ = next_cols.get(ch_idx, 0)
        if (not wide) and wide_cap.get(ch_idx, False):
            new_children.append(marker(2))
            state['auto_wide'] = True
            wide = True
        runs_ = p.findall('w:r', NS)
        before, after = [], []
        seen_br = False
        for r_ in runs_:
            if r_.find('w:br', NS) is not None:
                seen_br = True
                continue
            (after if seen_br else before).append(r_)
        p1 = para('TableTitle', keep_next=True)
        for r_ in before:
            r2 = copy.deepcopy(r_)
            clean_runs(r2)
            p1.append(r2)
        p2 = para('TableTitle', keep_next=True)
        p2.find('w:pPr', NS).append(el('spacing', after=60))
        for r_ in after:
            r2 = copy.deepcopy(r_)
            clean_runs(r2)
            p2.append(r2)
        fix_verbatim(p1); fix_verbatim(p2)
        state['caps'] = [p1, p2]
        state['prev'] = 'caption'
        continue
    # ----- figure caption -----
    if txt.startswith('Fig. ') and len(txt) > 20 and st == 'BodyText' and re.match(r'Fig\. \d+\.', txt):
        runs_ = p.findall('w:r', NS)
        np_ = para('FigCaption', jc='both')
        # drop label run ("Fig. n.") and one following space
        skipped = False
        for r_ in runs_:
            t = ''.join(x.text or '' for x in r_.findall('w:t', NS))
            if not skipped:
                skipped = True
                continue
            if t == ' ' and len(np_.findall('w:r', NS)) == 0:
                continue
            r2 = copy.deepcopy(r_)
            clean_runs(r2)
            np_.append(r2)
        fix_verbatim(np_)
        remap_rids(np_)
        new_children.append(np_)
        state['prev'] = 'caption'
        continue
    # ----- equations -----
    m = re.match(r'^(CR = .*|P\(S\) = .*?)\s*\((\d)\)\s*$', txt, re.S)
    if m and st == 'BodyText':
        np_ = para('Equation')
        ppr = np_.find('w:pPr', NS)
        tabs = el('tabs')
        sub(tabs, 'tab', val='clear', pos=4584)
        sub(tabs, 'tab', val='right', pos=4820)
        ppr.append(tabs)
        ppr.append(el('ind', left=0))
        np_.append(run(m.group(1)))
        rr_ = el('r'); rp_ = el('rPr'); sub(rp_, 'i', val=0); rr_.append(rp_)
        sub(rr_, 'tab')
        t = sub(rr_, 't'); t.text = '(%s)' % m.group(2)
        np_.append(rr_)
        new_children.append(np_)
        state['prev'] = 'eq'
        continue
    # ----- lists -----
    numpr = p.find('w:pPr/w:numPr', NS)
    if numpr is not None:
        nid = numpr.find('w:numId', NS).get(W + 'val')
        kind = list_kind(nid)
        list_counts[nid] = list_counts.get(nid, 0) + 1
        label = '•' if kind == 'bullet' else '%d.' % list_counts[nid]
        np_ = para('PARA')
        ppr = np_.find('w:pPr', NS)
        ppr.append(el('ind', left=300, hanging=300))
        ppr.append(el('jc', val='left'))
        np_.append(run(label))
        rr_ = el('r'); sub(rr_, 'tab'); np_.append(rr_)
        for c in p:
            if c.tag == W + 'r':
                np_.append(copy.deepcopy(c))
            elif c.tag == W + 'hyperlink':
                np_.append(copy.deepcopy(c))
        fix_verbatim(np_)
        remap_rids(np_)
        new_children.append(np_)
        state['prev'] = 'list'
        continue
    # ----- ordinary paragraphs -----
    if not txt.strip():
        continue
    style = 'PARA' if state['prev'] in ('heading',) else 'PARAIndent'
    if st == 'FirstParagraph' and state['prev'] == 'heading':
        style = 'PARA'
    np_ = para(style)
    for c in p:
        if c.tag in (W + 'r', W + 'hyperlink', W + 'bookmarkStart', W + 'bookmarkEnd'):
            np_.append(copy.deepcopy(c))
    fix_verbatim(np_)
    remap_rids(np_)
    new_children.append(np_)
    state['prev'] = 'para'


# pair adjacent one-column table blocks side by side in a full-width row (removes column gaps)
def _is_skip(e):
    return e.tag in (W + 'bookmarkStart', W + 'bookmarkEnd')
sb_ids = set(id(b) for b in single_blocks)
res = []
i_ = 0
npaired = 0
while i_ < len(new_children):
    e = new_children[i_]
    if id(e) in sb_ids:
        j_ = i_ + 1
        while j_ < len(new_children) and _is_skip(new_children[j_]):
            j_ += 1
        if j_ < len(new_children) and id(new_children[j_]) in sb_ids:
            a_, b_ = e, new_children[j_]
            grid = el('tbl')
            gp = sub(grid, 'tblPr')
            sub(gp, 'tblW', w=10000, type='dxa')
            sub(gp, 'jc', val='center')
            sub(gp, 'tblLayout', type='fixed')
            gm = sub(gp, 'tblCellMar')
            for side in ('top', 'left', 'bottom', 'right'):
                sub(gm, side, w=0, type='dxa')
            sub(gp, 'tblLook', val='0000')
            gg = sub(grid, 'tblGrid')
            for w_ in (4800, 400, 4800):
                sub(gg, 'gridCol', w=w_)
            gr = sub(grid, 'tr')
            grp = sub(gr, 'trPr'); sub(grp, 'cantSplit')
            for w_, blk in ((4800, a_), (400, None), (4800, b_)):
                c_ = sub(gr, 'tc')
                cp_ = sub(c_, 'tcPr'); sub(cp_, 'tcW', w=w_, type='dxa')
                if blk is not None:
                    c_.append(blk)
                tl_ = para(None)
                tpp = tl_.find('w:pPr', NS)
                sub(tpp, 'spacing', before=0, after=0, line=20, lineRule='exact')
                trp_ = sub(tpp, 'rPr'); sub(trp_, 'sz', val=2); sub(trp_, 'szCs', val=2)
                c_.append(tl_)
            res.append(marker(2))
            res.append(grid)
            res.append(marker(1))
            npaired += 1
            i_ = j_ + 1
            continue
    res.append(e)
    i_ += 1
new_children = res
print('paired table blocks:', npaired)

# merge adjacent wide regions: marker(1) immediately followed by marker(2) -> drop both
def _mk(e):
    if e.tag != W + 'p':
        return None
    sp = e.find('w:pPr/w:sectPr', NS)
    if sp is None:
        return None
    return int(sp.find('w:cols', NS).get(W + 'num', '1'))
changed = True
while changed:
    changed = False
    for i_ in range(len(new_children) - 1):
        if _mk(new_children[i_]) == 1 and _mk(new_children[i_ + 1]) == 2:
            # only merge the later-added markers (not the front matter one)
            if new_children[i_].find('w:pPr/w:sectPr/w:titlePg', NS) is None:
                del new_children[i_:i_ + 2]
                changed = True
                break

# final section
final = sectpr(2)
# ---------- write document.xml ----------
for ch in list(body):
    body.remove(ch)
for c in new_children:
    body.append(c)
body.append(final)
tpl.write(tpl_doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
tpl_rels.write(tpl_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
numt.write(num_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------- settings: even/odd headers ----------
sp = OUT + '/word/settings.xml'
s = open(sp, encoding='utf8').read()
if 'evenAndOddHeaders' not in s:
    if '<w:characterSpacingControl' in s:
        s = s.replace('<w:characterSpacingControl', '<w:evenAndOddHeaders/><w:characterSpacingControl', 1)
    else:
        s = s.replace('</w:settings>', '<w:evenAndOddHeaders/></w:settings>')
open(sp, 'w', encoding='utf8').write(s)

# ---------- header text and footer page numbers ----------
def fld(rpr_xml):
    return ('<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r><w:r>%s<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r><w:r>%s<w:t>1</w:t></w:r><w:r>%s<w:fldChar w:fldCharType="end"/></w:r>') % ((rpr_xml,) * 5)

for fn, lit in (('footer1.xml', '1'), ('footer2.xml', '2')):
    path = OUT + '/word/' + fn
    x = open(path, encoding='utf8').read()
    x = x.replace('VOLUME XX, 2017', 'VOLUME XX, 2026')
    # replace the run holding the literal page number
    rp_m = re.search(r'<w:rPr><w:rFonts w:ascii="Helvetica"[^>]*?/><w:sz w:val="12"/><w:szCs w:val="1[24]"/></w:rPr>', x)
    rpx = rp_m.group(0) if rp_m else '<w:rPr><w:sz w:val="12"/></w:rPr>'
    if fn == 'footer1.xml':
        x = re.sub(r'<w:tab/><w:t>1</w:t></w:r>', '<w:tab/></w:r>' + fld(rpx), x, count=1)
    else:
        x = re.sub(r'(<w:r [^>]*>)(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)<w:t>2</w:t></w:r>',
                   lambda m: fld(m.group(2)), x, count=1)
    open(path, 'w', encoding='utf8').write(x)

hp = OUT + '/word/header2.xml'
x = open(hp, encoding='utf8').read()
x = x.replace('>Author Name<', '>Pattanayak et al.<').replace('>Preparation of Papers for IEEE Access (February 2017)<', '>Review Gates in User-Controlled LLM Memory<')
assert 'Pattanayak et al.' in x
open(hp, 'w', encoding='utf8').write(x)

# ---------- core properties ----------
cp = OUT + '/docProps/core.xml'
x = open(cp, encoding='utf8').read()
x = re.sub(r'<dc:title>.*?</dc:title>', '<dc:title>%s</dc:title>' % title_text, x)
x = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator>Radha Mohan Pattanayak; Soubhagya Swastik Dhal; Basant Kumar; Mehal Mathur; Aditya Kumar Jena; Sachi Nandan Mohanty</dc:creator>', x)
open(cp, 'w', encoding='utf8').write(x)

# ---------- content types: png ----------
ctp = OUT + '/[Content_Types].xml'
x = open(ctp, encoding='utf8').read()
if 'Extension="png"' not in x:
    x = x.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1)
open(ctp, 'w', encoding='utf8').write(x)

# ---------- zip ----------
outdocx = '/tmp/build_access/aura-paper-ieee-access.docx'
if os.path.exists(outdocx):
    os.remove(outdocx)
with zipfile.ZipFile(outdocx, 'w', zipfile.ZIP_DEFLATED) as z:
    # content types first
    z.write(OUT + '/[Content_Types].xml', '[Content_Types].xml')
    for dp, dn, fns in os.walk(OUT):
        for f in fns:
            full = os.path.join(dp, f)
            arc = os.path.relpath(full, OUT)
            if arc == '[Content_Types].xml':
                continue
            z.write(full, arc)
print('built', outdocx, len(new_children), 'blocks')
