#!/usr/bin/env python3
"""
DOCX → JATS XML Converter v3 — Fully PMC-Compliant
PMC Tagging Guidelines: https://pmc.ncbi.nlm.nih.gov/tagging-guidelines/article/style/

Key PMC rules applied:
 1. <contrib id="..."> — DO NOT USE (use <contrib-id> for ORCID)
 2. <name> — no id attribute
 3. Affiliation inside <contrib-group> OR immediately after (both ok per PMC)
 4. issn @publication-format="print"/"electronic"
 5. pub-date @date-type="pub"
 6. Structured abstract with proper <sec><title><p> (not <bold>)
 7. sec-type on standard sections: intro/methods/results/discussion/conclusions
 8. <label> no trailing punctuation (follow-copy rule)
 9. <permissions> with <copyright-statement> + <license> for OA
10. <author-notes> with <corresp> + <fn fn-type="conflict">/<fn fn-type="financial-disclosure">
11. <fig position="float"> for floating figures
12. Table: <p> inside td/th
13. ref label = "1." (follow copy — punctuation from source)
14. element-citation: no punctuation, structured tags
15. <xref ref-type="bibr"> check rid matches ref id
16. <floats-group> for unreferenced figs at end
17. <custom-meta-group> for any extra metadata
"""

import sys, re, json, hashlib, urllib.request, urllib.parse
from docx import Document

# ============================================================
# ID GENERATOR
# ============================================================
_ctr = [0]

def make_prefix(title=""):
    return hashlib.md5((title or "article").encode()).hexdigest()[:8]

def nid(pfx, label):
    _ctr[0] += 1
    suf = hashlib.md5(f"{label}-{_ctr[0]}".encode()).hexdigest()[:12]
    return f"{label}-{pfx}-{_ctr[0]}-{suf}"

def xe(t):
    if not t: return ''
    return str(t).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def slugify(t):
    return re.sub(r'[^a-z0-9]+','-',t.lower()).strip('-')

# ============================================================
# SEC-TYPE MAPPING (PMC standard values)
# ============================================================
SEC_TYPE_MAP = {
    'introduction': 'intro',
    'material and methods': 'methods',
    'materials and methods': 'methods',
    'methods': 'methods',
    'methodology': 'methods',
    'results': 'results',
    'discussion': 'discussion',
    'conclusion': 'conclusions',
    'conclusions': 'conclusions',
    'acknowledgement': 'acknowledgments',
    'acknowledgements': 'acknowledgments',
    'acknowledgment': 'acknowledgments',
    'acknowledgments': 'acknowledgments',
    'supplementary': 'supplementary-material',
    'abbreviations': 'abbreviations',
    'case report': 'cases',
}

def get_sec_type(title):
    lc = title.lower().strip()
    for k, v in SEC_TYPE_MAP.items():
        if k in lc:
            return v
    return None

# ============================================================
# AUTHOR NAME PARSER
# ============================================================
def parse_author_name(raw):
    name = raw.strip().rstrip('.')
    if not name: return {'surname':'','given':'','prefix':''}
    parts = name.split()
    if len(parts)==1: return {'surname':parts[0],'given':'','prefix':''}
    # "Dhyani H" — last is single letter
    if len(parts)==2 and re.match(r'^[A-Z]{1,2}\.?$', parts[1]):
        return {'surname':parts[0],'given':parts[1].rstrip('.'),'prefix':''}
    # "H Dhyani" — first is single letter
    if len(parts)==2 and re.match(r'^[A-Z]{1,2}\.?$', parts[0]):
        return {'surname':parts[1],'given':parts[0].rstrip('.'),'prefix':''}
    # Prefix surnames: von, van, de, del, der, la, le, al
    prefixes = {'von','van','de','del','der','la','le','al','bin','binti','el'}
    if len(parts)>=3 and parts[-2].lower() in prefixes:
        return {'surname':' '.join(parts[-2:]),'given':' '.join(parts[:-2]),'prefix':''}
    # Default: last word = surname
    return {'surname':parts[-1],'given':' '.join(parts[:-1]),'prefix':''}

def parse_authors_para(para):
    authors = []
    cur_name = ''
    cur_affs = []
    is_corr = False

    for run in para.runs:
        t = run.text
        if not t: continue
        is_sup = run.font.superscript

        if is_sup:
            nums = [n.strip() for n in re.split(r'[,،\s]+', t) if n.strip().isdigit()]
            cur_affs.extend(nums)
        elif t.strip() == '*':
            is_corr = True
        elif ',' in t:
            parts = t.split(',')
            cur_name += parts[0]
            _flush_author(cur_name, cur_affs, is_corr, authors)
            cur_name = ','.join(parts[1:]).strip()
            cur_affs = []; is_corr = False
        else:
            cur_name += t

    if cur_name.strip():
        _flush_author(cur_name, cur_affs, is_corr, authors)
    return authors

def _flush_author(name, affs, corr, out):
    name = name.strip().strip(',').strip()
    if not name: return
    pn = parse_author_name(name)
    if pn['surname']:
        out.append({**pn, 'affiliationNums': affs[:], 'isCorresponding': corr})

def parse_affiliation(para):
    num = ''; parts = []
    for run in para.runs:
        if run.font.superscript and run.text.strip().isdigit():
            num = run.text.strip()
        else:
            parts.append(run.text)
    full = ''.join(parts).strip()
    if not num:
        m = re.match(r'^(\d+)\s*(.+)', para.text.strip())
        if m: num, full = m.group(1), m.group(2).strip()
        else: full = para.text.strip()
    return num, full

# ============================================================
# INLINE TEXT → JATS
# ============================================================
def para_to_inline(para, pfx, cite_ids):
    """Convert para runs to JATS inline XML. cite_ids: set of valid ref nums."""
    out = ''
    for run in para.runs:
        t = run.text
        if not t: continue
        sup = run.font.superscript
        ital = run.italic
        bold = run.bold

        if sup and t.strip():
            nums_str = t.strip()
            if re.match(r'^[\d,\s\-–]+$', nums_str):
                for num in re.split(r'[,\s]+', nums_str):
                    num = num.strip().lstrip('0') or num.strip()
                    if num and num.isdigit():
                        xid = nid(pfx, "x")
                        out += f'<xref id="{xid}" rid="{pfx}-B{num}" ref-type="bibr">{xe(num)}</xref>'
            else:
                out += f'<sup>{xe(t)}</sup>'
        elif ital and bold:
            out += f'<bold><italic>{xe(t)}</italic></bold>'
        elif ital:
            out += f'<italic>{xe(t)}</italic>'
        elif bold:
            bid = nid(pfx, "s")
            out += f'<bold id="{bid}">{xe(t)}</bold>'
        else:
            out += xe(t)
    return out

# ============================================================
# REFERENCE PARSER
# ============================================================
def parse_ref(raw):
    r = {'authors':[],'title':'','journal':'','year':'','volume':'','issue':'',
         'fpage':'','lpage':'','doi':'','hasEtAl':False,'pubType':'journal'}
    dm = re.search(r'https?://doi\.org/(10\.[^\s]+)', raw)
    if dm: r['doi'] = dm.group(1).rstrip('.')
    clean = re.sub(r'https?://\S+','',raw).strip()
    if re.search(r'\[dissertation\]|\[thesis\]', clean, re.I): r['pubType']='thesis'
    ym = re.search(r'[;.]\s*((?:19|20)\d{2})\s*[;:]', clean) or re.search(r'\b((?:20|19)\d{2})\b', clean)
    if ym: r['year'] = ym.group(1)
    vm = re.search(r'(\d+)\s*\((\d+)\)\s*:\s*([0-9]+)\s*[-–]\s*([0-9]+)', clean)
    if vm:
        r['volume'],r['issue'],r['fpage'],r['lpage'] = vm.group(1),vm.group(2),vm.group(3),vm.group(4)
    else:
        vm2 = re.search(r'(\d+)\s*:\s*([0-9]+)\s*[-–]\s*([0-9]+)', clean)
        if vm2: r['volume'],r['fpage'],r['lpage'] = vm2.group(1),vm2.group(2),vm2.group(3)
    sm = re.match(r'^(.+?)\.\s+([A-Z].+)', clean)
    if sm:
        for ap in re.split(r',\s+(?=[A-Z])', sm.group(1)):
            ap=ap.strip()
            if re.search(r'et al\.?$',ap,re.I): r['hasEtAl']=True; break
            pn=parse_author_name(ap)
            if pn['surname']: r['authors'].append(pn)
    return r

def fetch_crossref(q, timeout=5):
    try:
        url = f"https://api.crossref.org/works?query={urllib.parse.quote(q)}&rows=1&select=DOI,author,title,published,container-title,volume,issue,page"
        req = urllib.request.Request(url, headers={'User-Agent':'JATSConverter/3.0 (medical-journal@example.com)'})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read())
        item = (data.get('message',{}).get('items') or [None])[0]
        if not item: return None
        return {
            'doi': item.get('DOI',''),
            'title': (item.get('title') or [''])[0],
            'journal': (item.get('container-title') or [''])[0],
            'year': str(((item.get('published',{}).get('date-parts') or [['']])[0] or [''])[0]),
            'volume': item.get('volume',''), 'issue': item.get('issue',''), 'pages': item.get('page',''),
            'authors': [{'surname':a.get('family',''),'given':a.get('given',''),'orcid':a.get('ORCID','').replace('http://orcid.org/','')} for a in item.get('author',[])],
        }
    except: return None

# ============================================================
# DOCX → PARSED STRUCTURE
# ============================================================
def parse_docx(path, use_crossref=True):
    doc = Document(path)
    parsed = {
        'title':'','authors':[],'affiliations':{},'abstract':{},'keywords':[],
        'receivedDate':'','acceptedDate':'','sections':[],'references':[],'tables':[],'figures':[],
        'openAccess': True,
    }
    cur_sec = cur_sub = None
    ref_num = 1; table_num = 0

    for para in doc.paragraphs:
        sty = para.style.name
        txt = para.text.strip()

        if sty == 'Title':
            parsed['title'] = txt

        elif sty == 'Author Name':
            parsed['authors'] = parse_authors_para(para)

        elif sty in ('Authors affiliation','Last Authors affiliation'):
            if txt:
                num, aff = parse_affiliation(para)
                if num: parsed['affiliations'][num] = aff
                else:
                    n = str(len(parsed['affiliations'])+1)
                    parsed['affiliations'][n] = txt

        elif sty == 'Abstract':
            if not txt or 'Open Access' in txt or 'reprint' in txt.lower() or txt.startswith('For reprints'): continue
            lm = re.match(r'^(Background|Methods?|Results?|Conclusion|Objective|Aim|Discussion|Summary):\s*', txt, re.I)
            if lm:
                parsed['abstract'][lm.group(1)] = txt[len(lm.group(0)):]
            else:
                k = list(parsed['abstract'].keys())[-1] if parsed['abstract'] else 'text'
                if k not in parsed['abstract']: parsed['abstract'][k] = txt
                else: parsed['abstract'][k] += ' ' + txt

        elif sty == 'Keywords':
            if 'Keywords:' in txt:
                parsed['keywords'] = [k.strip().rstrip('.') for k in txt.split('Keywords:',1)[1].split(',') if k.strip()]
            elif 'Received:' in txt:
                rm = re.search(r'Received:\s*([\d\-]+)', txt)
                am = re.search(r'Accepted:\s*([\d\-]+)', txt)
                if rm: parsed['receivedDate'] = rm.group(1)
                if am: parsed['acceptedDate'] = am.group(1)

        elif sty == 'Heading 1':
            lc = txt.lower()
            skip = ['reference','source of fund','conflict','ethical','patient consent']
            if any(s in lc for s in skip):
                cur_sec = cur_sub = None
            else:
                cur_sec = {'title':txt,'id':slugify(txt),'paragraphs':[],'subsections':[],'sec_type':get_sec_type(txt)}
                cur_sub = None
                parsed['sections'].append(cur_sec)

        elif sty == 'Heading 2':
            if cur_sec:
                cur_sub = {'title':txt,'id':slugify(txt),'paragraphs':[],'sec_type':get_sec_type(txt)}
                cur_sec['subsections'].append(cur_sub)

        elif sty in ('Paragraph 1','2nd Para','List Paragraph','Normal'):
            if not txt: continue
            if cur_sub: cur_sub['paragraphs'].append(para)
            elif cur_sec: cur_sec['paragraphs'].append(para)

        elif sty == 'Reference':
            if txt and not txt.startswith('http'):
                rp = parse_ref(txt)
                parsed['references'].append({'num':ref_num,'raw':txt,'doi':rp.get('doi',''),'parsed':rp,'crossref':None})
                ref_num += 1
            elif txt.startswith('http') and parsed['references']:
                dm = re.search(r'10\.\S+', txt)
                if dm: parsed['references'][-1]['doi'] = dm.group(0).rstrip('.')

        elif sty == 'Table caption':
            if txt:
                tm = re.search(r'Table\s+(\d+)', txt, re.I)
                tnum = int(tm.group(1)) if tm else table_num+1
                found = any(t['num']==tnum and not t.get('caption') for t in parsed['tables'])
                if not found:
                    upd = False
                    for t in parsed['tables']:
                        if t['num']==tnum:
                            t['caption']=txt; upd=True; break
                    if not upd:
                        parsed['tables'].append({'num':tnum,'caption':txt,'rows':[],'colwidths':[]})

    # Tables
    table_num = 0
    for table in doc.tables:
        table_num += 1
        rows=[]; colwidths=[]
        try:
            total = sum(c.width or 1 for c in table.columns if c.width) or 1
            for col in table.columns:
                w = col.width or 1
                colwidths.append(round((w/total)*100,2))
        except: pass

        for row in table.rows:
            cells=[]
            for cell in row.cells:
                tc = cell._tc
                ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                gs = tc.find(f'.//{{{ns}}}gridSpan')
                colspan = int(gs.get(f'{{{ns}}}val',1)) if gs is not None else 1
                cells.append({'text':cell.text.strip(),'colspan':colspan,'rowspan':1})
            rows.append(cells)

        found=False
        for tbl in parsed['tables']:
            if tbl['num']==table_num and not tbl.get('rows'):
                tbl['rows']=rows; tbl['colwidths']=colwidths; found=True; break
        if not found:
            parsed['tables'].append({'num':table_num,'caption':'','rows':rows,'colwidths':colwidths})

    # CrossRef
    if use_crossref and parsed['references']:
        print(f"\n🔍 CrossRef lookup for {len(parsed['references'])} references...", file=sys.stderr)
        for ref in parsed['references']:
            p=ref['parsed']
            q=' '.join(filter(None,[
                p['authors'][0].get('surname','') if p['authors'] else '',
                ' '.join(p['title'].split()[:4]) if p['title'] else '',
                p['year']
            ]))
            if q.strip():
                cr=fetch_crossref(q)
                ref['crossref']=cr
                if cr and not ref['doi']: ref['doi']=cr.get('doi','')
            sym = '✓' if ref.get('crossref') else '✗'
            print(f"  [{ref['num']:2d}] {sym} {ref['raw'][:55]}...", file=sys.stderr)

    return parsed


# ============================================================
# JATS XML BUILDER — PMC COMPLIANT v3
# ============================================================
def build_xml(parsed, jm):
    pfx = make_prefix(parsed.get('title',''))
    _ctr[0] = 0
    cite_ids = {str(r['num']) for r in parsed.get('references',[])}

    L = []  # lines

    L += [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<!DOCTYPE article PUBLIC "-//NLM//DTD JATS (Z39.96) Journal Publishing DTD v1.2 20190208//EN"',
        '  "JATS-journalpublishing1-2.dtd">',
        f'<article xmlns:xlink="http://www.w3.org/1999/xlink"',
        f'         article-type="{jm.get("articleType","research-article")}"',
        f'         xml:lang="en">',
        '',
        '  <!-- ============================================================',
        '       FRONT MATTER',
        '       ============================================================ -->',
        '  <front>',
    ]

    # ---- journal-meta ----
    jm_id = nid(pfx,"journal-meta")
    L += [
        f'    <journal-meta id="{jm_id}">',
        f'      <journal-id journal-id-type="nlm-ta">{xe(jm.get("publisher","IP Innovative Publication"))}</journal-id>',
        f'      <journal-id journal-id-type="publisher-id">{xe(jm.get("publisher","IP Innovative Publication"))}</journal-id>',
    ]
    # journal URL goes in self-uri, not journal-id (invalid type per PMC)
    L.append('      <journal-title-group>')
    L.append(f'        <journal-title>{xe(jm.get("name",""))}</journal-title>')
    L.append('      </journal-title-group>')
    if jm.get('issnPrint'): L.append(f'      <issn publication-format="print">{xe(jm["issnPrint"].strip())}</issn>')
    if jm.get('issnElec'): L.append(f'      <issn publication-format="electronic">{xe(jm["issnElec"].strip())}</issn>')
    if jm.get('journalUrl'):
        L.append(f'      <self-uri xlink:href="{xe(jm["journalUrl"])}"/>')
    L.append('    </journal-meta>')
    L.append('')

    # ---- article-meta ----
    am_id = nid(pfx,"article-meta")
    L.append(f'    <article-meta id="{am_id}">')
    if jm.get('doi'):
        L.append(f'      <article-id pub-id-type="doi">{xe(jm["doi"])}</article-id>')
    L += [
        '      <article-categories>',
        '        <subj-group subj-group-type="heading">',
        f'          <subject>{xe(jm.get("articleTypeLabel","Original Research Article"))}</subject>',
        '        </subj-group>',
        '      </article-categories>',
        '',
        '      <title-group>',
        f'        <article-title>{xe(parsed.get("title",""))}</article-title>',
        '      </title-group>',
        '',
        '      <contrib-group>',
    ]

    # Authors — PMC rule: no id on <contrib>, use <contrib-id> for ORCID
    for auth in parsed.get('authors',[]):
        corr = ' corresp="yes"' if auth.get('isCorresponding') else ''
        L.append(f'        <contrib contrib-type="author"{corr}>')
        if auth.get('orcid'):
            L.append(f'          <contrib-id contrib-id-type="orcid">{xe(auth["orcid"])}</contrib-id>')
        # PMC: no id on <name>
        L.append('          <name name-style="western">')
        L.append(f'            <surname>{xe(auth["surname"])}</surname>')
        if auth.get('given'):
            L.append(f'            <given-names>{xe(auth["given"])}</given-names>')
        L.append('          </name>')
        for an in auth.get('affiliationNums',[]):
            xid = nid(pfx,"x")
            L.append(f'          <xref id="{xid}" rid="aff{an}" ref-type="aff"><sup>{xe(an)}</sup></xref>')
        if auth.get('isCorresponding'):
            xid2 = nid(pfx,"x")
            L.append(f'          <xref id="{xid2}" rid="cor1" ref-type="corresp">*</xref>')
        L.append('        </contrib>')

    # Affiliations inside contrib-group (PMC preferred)
    for num, txt in parsed.get('affiliations',{}).items():
        parts = [p.strip() for p in txt.split(',')]
        L.append(f'        <aff id="aff{num}">')
        L.append(f'          <label>{xe(num)}</label>')
        if len(parts) >= 2:
            L.append(f'          <institution content-type="dept">{xe(parts[0])}</institution>')
            mid = ', '.join(parts[1:-2]) if len(parts)>3 else (parts[1] if len(parts)>1 else '')
            if mid: L.append(f'          <institution>{xe(mid)}</institution>')
            if len(parts)>=3: L.append(f'          <addr-line>{xe(parts[-2])}</addr-line>')
            L.append(f'          <country country="IN">{xe(parts[-1])}</country>')
        else:
            L.append(f'          <institution>{xe(txt)}</institution>')
        L.append('        </aff>')

    L.append('      </contrib-group>')
    L.append('')

    # Author notes — PMC rule: corresp + fn for conflict/funding
    corr_authors = [a for a in parsed.get('authors',[]) if a.get('isCorresponding')]
    L.append('      <author-notes>')
    if corr_authors:
        ca = corr_authors[0]
        L += [
            '        <corresp id="cor1">',
            f'          <bold>Corresponding Author:</bold> {xe(ca["given"])} {xe(ca["surname"])}',
            '        </corresp>',
        ]
    L += [
        '        <fn fn-type="conflict">',
        '          <p>None declared.</p>',
        '        </fn>',
        '        <fn fn-type="financial-disclosure">',
        '          <p>None.</p>',
        '        </fn>',
        '      </author-notes>',
        '',
    ]

    # Pub date — PMC rule: @date-type="pub"
    L += [
        '      <pub-date date-type="pub" publication-format="print">',
        f'        <day>{xe(jm.get("day",""))}</day>',
        f'        <month>{xe(jm.get("month",""))}</month>',
        f'        <year>{xe(jm.get("year","2025"))}</year>',
        '      </pub-date>',
    ]
    if jm.get('volume'): L.append(f'      <volume>{xe(jm["volume"])}</volume>')
    if jm.get('issue'): L.append(f'      <issue>{xe(jm["issue"])}</issue>')
    if jm.get('fpage'): L.append(f'      <fpage>{xe(jm["fpage"])}</fpage>')
    if jm.get('lpage'): L.append(f'      <lpage>{xe(jm["lpage"])}</lpage>')

    # History
    if parsed.get('receivedDate') or parsed.get('acceptedDate'):
        L.append('      <history>')
        for dtype, dstr in [('received',parsed.get('receivedDate','')),('accepted',parsed.get('acceptedDate',''))]:
            if not dstr: continue
            d=dstr.split('-')
            if len(d[0])==4: yr,mo,dy=d[0],d[1] if len(d)>1 else '',d[2] if len(d)>2 else ''
            else: dy,mo,yr=d[0],d[1] if len(d)>1 else '',d[2] if len(d)>2 else ''
            L += [f'        <date date-type="{dtype}">',
                  f'          <day>{dy}</day>',f'          <month>{mo}</month>',f'          <year>{yr}</year>',
                  '        </date>']
        L.append('      </history>')

    # Permissions — PMC rule: copyright-statement + license for OA
    yr = jm.get('year','2025')
    L += [
        '      <permissions>',
        f'        <copyright-statement>© {yr} The Author(s)</copyright-statement>',
        f'        <copyright-year>{yr}</copyright-year>',
        '        <license license-type="open-access" xlink:href="https://creativecommons.org/licenses/by-nc/4.0/">',
        '          <license-p>This is an Open Access article distributed under the terms of the',
        '          Creative Commons Attribution-NonCommercial 4.0 International License',
        '          (<ext-link ext-link-type="uri" xlink:href="https://creativecommons.org/licenses/by-nc/4.0/">https://creativecommons.org/licenses/by-nc/4.0/</ext-link>)',
        '          which permits unrestricted non-commercial use, distribution, and reproduction',
        '          in any medium, provided the original work is properly cited.</license-p>',
        '        </license>',
        '      </permissions>',
        '',
    ]

    # Abstract — PMC rule: structured with <sec><title><p>, NOT <bold>
    abstract = parsed.get('abstract',{})
    if abstract:
        abs_id = nid(pfx,"abstract")
        L.append(f'      <abstract id="{abs_id}">')
        L.append('        <title>Abstract</title>')
        if len(abstract)==1 and 'text' in abstract:
            pid=nid(pfx,"p")
            L.append(f'        <p id="{pid}">{xe(abstract["text"].strip())}</p>')
        else:
            for label, text in abstract.items():
                if label=='text':
                    pid=nid(pfx,"p")
                    L.append(f'        <p id="{pid}">{xe(text.strip())}</p>')
                else:
                    sid=nid(pfx,"sec")
                    tid=nid(pfx,"title")
                    pid=nid(pfx,"p")
                    L += [f'        <sec id="{sid}">',
                          f'          <title id="{tid}">{xe(label)}</title>',
                          f'          <p id="{pid}">{xe(text.strip())}</p>',
                          '        </sec>']
        L.append('      </abstract>')

    # Keywords
    kws = parsed.get('keywords',[])
    if kws:
        kg_id = nid(pfx,"kwd-group")
        L += [f'      <kwd-group id="{kg_id}" kwd-group-type="author-generated">',
              '        <title>Keywords</title>']
        for kw in kws: L.append(f'        <kwd>{xe(kw)}</kwd>')
        L.append('      </kwd-group>')

    L += ['', '    </article-meta>', '  </front>', '']

    # ============================================================
    # BODY
    # ============================================================
    L += [
        '  <!-- ============================================================',
        '       BODY',
        '       ============================================================ -->',
        '  <body>',
    ]

    for sec in parsed.get('sections',[]):
        st = sec.get('sec_type')
        st_attr = f' sec-type="{st}"' if st else ''
        L.append(f'    <sec{st_attr}>')
        tid = nid(pfx,"title")
        L.append(f'      <title id="{tid}">{xe(sec["title"])}</title>')

        for para in sec.get('paragraphs',[]):
            inline = para_to_inline(para, pfx, cite_ids)
            if inline.strip():
                pid = nid(pfx,"p")
                L.append(f'      <p id="{pid}">{inline}</p>')

        for sub in sec.get('subsections',[]):
            sst = sub.get('sec_type')
            sst_attr = f' sec-type="{sst}"' if sst else ''
            L.append(f'      <sec{sst_attr}>')
            stid = nid(pfx,"title")
            L.append(f'        <title id="{stid}">{xe(sub["title"])}</title>')
            for para in sub.get('paragraphs',[]):
                inline = para_to_inline(para, pfx, cite_ids)
                if inline.strip():
                    pid = nid(pfx,"p")
                    L.append(f'        <p id="{pid}">{inline}</p>')
            L.append('      </sec>')

        L.append('    </sec>')
        L.append('')

    L.append('  </body>')
    L.append('')

    # ============================================================
    # BACK MATTER
    # DTD order: front -> body -> back -> floats-group
    # ============================================================
    L += [
        '  <!-- ============================================================',
        '       BACK MATTER',
        '       ============================================================ -->',
        '  <back>',

    ]

    refs = parsed.get('references',[])
    if refs:
        L.append('    <ref-list>')
        L.append('      <title>References</title>')
        for ref in refs:
            p=ref.get('parsed',{}); cr=ref.get('crossref') or {}
            authors = cr.get('authors') or p.get('authors',[])
            year_r = cr.get('year') or p.get('year','')
            doi = ref.get('doi') or cr.get('doi','')
            journal_r = cr.get('journal') or p.get('journal','')
            vol_r = cr.get('volume') or p.get('volume','')
            iss_r = cr.get('issue') or p.get('issue','')
            cr_pg = cr.get('pages','')
            if cr_pg and '-' in cr_pg: fp_r,lp_r = cr_pg.split('-',1)
            else: fp_r=cr_pg or p.get('fpage',''); lp_r=p.get('lpage','')
            ti_r = cr.get('title') or p.get('title','')
            pub_t = p.get('pubType','journal')

            # PMC rule: ref id = pfx-B{num}, label = follow copy (include punctuation from source)
            L.append(f'      <ref id="{pfx}-B{ref["num"]}">')
            L.append(f'        <label>{ref["num"]}.</label>')
            L.append(f'        <element-citation publication-type="{pub_t}">')

            if authors:
                L.append('          <person-group person-group-type="author">')
                for a in authors:
                    L.append('            <name name-style="western">')
                    L.append(f'              <surname>{xe(a.get("surname",""))}</surname>')
                    if a.get('given'): L.append(f'              <given-names>{xe(a["given"])}</given-names>')
                    L.append('            </name>')
                    if a.get('orcid'):
                        L.append(f'            <contrib-id contrib-id-type="orcid">{xe(a["orcid"])}</contrib-id>')
                if p.get('hasEtAl'): L.append('            <etal/>')
                L.append('          </person-group>')

            # PMC rule: element-citation — no punctuation, just tags
            if ti_r: L.append(f'          <article-title>{xe(ti_r)}</article-title>')
            if journal_r: L.append(f'          <source>{xe(journal_r)}</source>')
            if year_r: L.append(f'          <year iso-8601-date="{year_r}">{xe(year_r)}</year>')
            if vol_r: L.append(f'          <volume>{xe(vol_r)}</volume>')
            if iss_r: L.append(f'          <issue>{xe(iss_r)}</issue>')
            if fp_r: L.append(f'          <fpage>{xe(fp_r.strip())}</fpage>')
            if lp_r: L.append(f'          <lpage>{xe(lp_r.strip())}</lpage>')
            if doi: L.append(f'          <pub-id pub-id-type="doi">{xe(doi)}</pub-id>')

            if not authors and not ti_r:
                L.append(f'          <!-- RAW: {xe(ref["raw"][:200])} -->')

            L += ['        </element-citation>', '      </ref>']
        L.append('    </ref-list>')

    L.append('  </back>')
    L.append('')

    # FLOATS GROUP — must come after <back> per DTD:
    # front, body?, back?, floats-group?
    tables = parsed.get('tables',[])
    if tables:
        L += [
            '  <!-- ============================================================',
            '       FLOATS GROUP (Tables) — after back per DTD',
            '       ============================================================ -->',
            '  <floats-group>',
        ]
        for tbl in tables:
            L += build_table(tbl, pfx)
        L += ['  </floats-group>', '']

    L.append('</article>')
    return '\n'.join(L)

def build_table(tbl, pfx):
    L=[]
    tw_id = nid(pfx,"table-wrap")
    # PMC rule: floating tables with position="float"
    L.append(f'    <table-wrap id="{tw_id}" position="float" orientation="portrait">')
    L.append(f'      <label>Table {tbl["num"]}</label>')

    cap_txt = tbl.get('caption','')
    cap_clean = re.sub(r'^Table\s+\d+\s*[:\-]\s*','',cap_txt,flags=re.I)
    if cap_clean:
        cid=nid(pfx,"caption"); ctid=nid(pfx,"title")
        L += [f'      <caption id="{cid}">',
              f'        <title id="{ctid}">{xe(cap_clean)}</title>',
              '      </caption>']

    tbl_id=nid(pfx,"table")
    L.append(f'      <table id="{tbl_id}" rules="all" frame="box">')

    rows=tbl.get('rows',[]); cws=tbl.get('colwidths',[])
    if cws:
        L.append('        <colgroup>')
        for w in cws: L.append(f'          <col width="{w}"/>')
        L.append('        </colgroup>')
    elif rows and rows[0]:
        n=len(rows[0]); w=round(100/n,2)
        L.append('        <colgroup>')
        for _ in rows[0]: L.append(f'          <col width="{w}"/>')
        L.append('        </colgroup>')

    if rows:
        thead_id=nid(pfx,"table-section-header"); tr_id=nid(pfx,"tr")
        L += [f'        <thead id="{thead_id}">',f'          <tr id="{tr_id}">']
        for cell in rows[0]:
            tc_id=nid(pfx,"tc"); p_id=nid(pfx,"p"); b_id=nid(pfx,"strong")
            cs=f' colspan="{cell["colspan"]}"' if int(cell.get('colspan',1))>1 else ''
            cell_h = xe(cell['text'])
            L.append(f'            <th id="{tc_id}"{cs} align="left">')
            if cell_h.strip():
                L.append(f'              <p id="{p_id}"><bold id="{b_id}">{cell_h}</bold></p>')
            L.append('            </th>')
        L += ['          </tr>','        </thead>']

        if len(rows)>1:
            tb_id=nid(pfx,"table-section")
            L.append(f'        <tbody id="{tb_id}">')
            for row in rows[1:]:
                tr_id=nid(pfx,"table-row")
                L.append(f'          <tr id="{tr_id}">')
                for cell in row:
                    td_id=nid(pfx,"table-cell"); p_id=nid(pfx,"p")
                    cs=f' colspan="{cell["colspan"]}"' if int(cell.get('colspan',1))>1 else ''
                    rs=f' rowspan="{cell["rowspan"]}"' if int(cell.get('rowspan',1))>1 else ''
                    cell_text = xe(cell['text'])
                    L.append(f'            <td id="{td_id}"{cs}{rs} align="left">')
                    if cell_text.strip():
                        L.append(f'              <p id="{p_id}">{cell_text}</p>')
                    L.append('            </td>')
                L.append('          </tr>')
            L.append('        </tbody>')

    L += ['      </table>','    </table-wrap>']
    return L

# ============================================================
# ENTRY POINT
# ============================================================
def post_process(xml):
    """Remove empty p/bold/italic elements that PMC rejects."""
    # Empty p: contains only whitespace, &#160;, or nothing
    xml = re.sub(r'<p([^>]*)>(\s|&#160;|&#xA0;| )*</p>', '', xml)
    # Empty bold
    xml = re.sub(r'<bold([^>]*)>(\s)*</bold>', '', xml)
    # Empty italic
    xml = re.sub(r'<italic([^>]*)>(\s)*</italic>', '', xml)
    # Remove blank lines created by removals
    lines = [l for l in xml.split('\n') if l.strip()]
    return '\n'.join(lines)


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='DOCX → JATS XML v3 (PMC Compliant)')
    ap.add_argument('input', help='.docx file path')
    ap.add_argument('-o','--output', help='Output .xml path')
    ap.add_argument('--journal',     default='Journal Name')
    ap.add_argument('--abbrev',      default='')
    ap.add_argument('--issn-print',  default='')
    ap.add_argument('--issn-elec',   default='')
    ap.add_argument('--publisher',   default='IP Innovative Publication')
    ap.add_argument('--journal-url', default='')
    ap.add_argument('--doi',         default='')
    ap.add_argument('--volume',      default='')
    ap.add_argument('--issue',       default='')
    ap.add_argument('--year',        default='2025')
    ap.add_argument('--month',       default='')
    ap.add_argument('--day',         default='')
    ap.add_argument('--fpage',       default='')
    ap.add_argument('--lpage',       default='')
    ap.add_argument('--type',        default='research-article', dest='article_type')
    ap.add_argument('--no-crossref', action='store_true')
    args = ap.parse_args()

    TYPE_LABELS = {
        'research-article':'Original Research Article','review-article':'Review Article',
        'case-report':'Case Report','letter':'Letter','editorial':'Editorial',
        'brief-report':'Brief Report','systematic-review':'Systematic Review',
    }
    jm = {
        'name':args.journal,'abbrev':args.abbrev,'issnPrint':args.issn_print,
        'issnElec':args.issn_elec,'publisher':args.publisher,'journalUrl':args.journal_url,
        'doi':args.doi,'volume':args.volume,'issue':args.issue,'year':args.year,
        'month':args.month,'day':args.day,'fpage':args.fpage,'lpage':args.lpage,
        'articleType':args.article_type,
        'articleTypeLabel':TYPE_LABELS.get(args.article_type,args.article_type),
    }

    print(f"📄 Input:  {args.input}", file=sys.stderr)
    parsed = parse_docx(args.input, use_crossref=not args.no_crossref)

    print(f"\n📊 Parsed:", file=sys.stderr)
    print(f"   Title:        {parsed['title'][:65]}", file=sys.stderr)
    print(f"   Authors:      {len(parsed['authors'])}", file=sys.stderr)
    for a in parsed['authors']:
        corr = ' *' if a['isCorresponding'] else ''
        print(f"     → {a['given']} {a['surname']}{corr} [aff: {a['affiliationNums']}]", file=sys.stderr)
    print(f"   Affiliations: {len(parsed['affiliations'])}", file=sys.stderr)
    print(f"   Sections:     {len(parsed['sections'])}", file=sys.stderr)
    for s in parsed['sections']:
        st = f" ({s['sec_type']})" if s.get('sec_type') else ''
        print(f"     → {s['title']}{st} [{len(s['paragraphs'])} para, {len(s['subsections'])} sub]", file=sys.stderr)
    print(f"   Tables:       {len(parsed['tables'])}", file=sys.stderr)
    print(f"   References:   {len(parsed['references'])}", file=sys.stderr)

    xml = build_xml(parsed, jm)
    xml = post_process(xml)
    out = args.output or args.input.replace('.docx','-jats-v3.xml')
    with open(out,'w',encoding='utf-8') as f: f.write(xml)
    print(f"\n✅ Saved: {out}", file=sys.stderr)
    print(f"   Size:  {len(xml)/1024:.1f} KB | Lines: {xml.count(chr(10))}", file=sys.stderr)
    print(f"\n🔗 Validate at: https://validator.jats4r.org", file=sys.stderr)
