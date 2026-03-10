"""
工事写真管理 Flask アプリ
- 元xlsxテンプレートをzipレベルで操作し写真・テキストを差し替えてExcel出力
- db.xlsx から管理番号→物件名称・担当者を検索する /api/lookup エンドポイント
"""

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from PIL import Image
from datetime import date
import io, base64, os, re, zipfile
from lxml import etree
import openpyxl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

# GitHub Pages など別オリジンからのアクセスを許可
CORS(app, resources={r"/api/*": {"origins": "*"}})

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')
DB_PATH       = os.path.join(os.path.dirname(__file__), 'db.xlsx')

SS_NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
R_NS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS  = 'http://schemas.openxmlformats.org/package/2006/content-types'
A_NS   = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# ══════════════════════════════════════════════════════
# DB検索  db.xlsx: A=管理番号 B=物件名称 C=担当者
# ══════════════════════════════════════════════════════
def load_db():
    if not os.path.exists(DB_PATH):
        return {}
    wb = openpyxl.load_workbook(DB_PATH, read_only=True, data_only=True)
    ws = wb.active
    db = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        code    = str(row[0]).strip()
        name    = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        manager = str(row[2]).strip() if len(row) > 2 and row[2] else ''
        if code:
            db[code] = {'name': name, 'manager': manager}
    wb.close()
    return db

@app.route('/api/lookup')
def lookup():
    code = request.args.get('code', '').strip()
    if not code:
        return jsonify({'error': 'コードが指定されていません'}), 400
    db = load_db()
    if code in db:
        return jsonify({'found': True, 'code': code, **db[code]})
    return jsonify({'found': False, 'code': code}), 404

# ══════════════════════════════════════════════════════
# Excel生成ユーティリティ
# ══════════════════════════════════════════════════════
def find_or_add_ss(ss_root, text):
    sis = ss_root.findall(f'{{{SS_NS}}}si')
    for i, si in enumerate(sis):
        t = si.find(f'{{{SS_NS}}}t')
        if t is not None and t.text == text:
            return i
    new_si = etree.SubElement(ss_root, f'{{{SS_NS}}}si')
    new_t  = etree.SubElement(new_si, f'{{{SS_NS}}}t')
    new_t.text = text
    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return len(sis)

def set_cell_ss(sheet_root, ref, idx):
    for c in sheet_root.iter(f'{{{SS_NS}}}c'):
        if c.get('r') == ref:
            c.set('t', 's')
            for f in c.findall(f'{{{SS_NS}}}f'):
                c.remove(f)
            v = c.find(f'{{{SS_NS}}}v')
            if v is None:
                v = etree.SubElement(c, f'{{{SS_NS}}}v')
            v.text = str(idx)
            return

def set_cell_formula_val(sheet_root, ref, val):
    for c in sheet_root.iter(f'{{{SS_NS}}}c'):
        if c.get('r') == ref:
            v = c.find(f'{{{SS_NS}}}v')
            if v is None:
                v = etree.SubElement(c, f'{{{SS_NS}}}v')
            v.text = val
            return

def replace_blip_rids(drawing_xml_bytes, old_to_new):
    root = etree.fromstring(drawing_xml_bytes)
    for blip in root.iter(f'{{{A_NS}}}blip'):
        old = blip.get(f'{{{R_NS}}}embed')
        if old in old_to_new:
            blip.set(f'{{{R_NS}}}embed', old_to_new[old])
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def make_rels_xml(entries):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>',
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rid, rtype, target in entries:
        lines.append(f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>')
    lines.append('</Relationships>')
    return '\n'.join(lines).encode('utf-8')

# ══════════════════════════════════════════════════════
# Excel生成メイン
# ══════════════════════════════════════════════════════
def generate_excel(project, parts):
    with open(TEMPLATE_PATH, 'rb') as f:
        tmpl = f.read()
    with zipfile.ZipFile(io.BytesIO(tmpl), 'r') as tz:
        files = {n: tz.read(n) for n in tz.namelist()}

    tmpl_sheet_xml   = files['xl/worksheets/sheet1.xml']
    tmpl_drawing_xml = files['xl/drawings/drawing1.xml']
    ss_root          = etree.fromstring(files['xl/sharedStrings.xml'])
    wb_root          = etree.fromstring(files['xl/workbook.xml'])

    sheets_el = wb_root.find(f'{{{SS_NS}}}sheets') or wb_root.find('sheets')
    for s in list(sheets_el): sheets_el.remove(s)
    for tag in ['definedNames', f'{{{SS_NS}}}definedNames']:
        for dn in wb_root.findall(tag): wb_root.remove(dn)

    wb_rels_root = etree.fromstring(files['xl/_rels/workbook.xml.rels'])
    for r in list(wb_rels_root):
        if 'worksheet' in r.get('Type', ''): wb_rels_root.remove(r)

    ct_root = etree.fromstring(files['[Content_Types].xml'])
    for ov in list(ct_root):
        pn = ov.get('PartName', '')
        if '/xl/worksheets/sheet' in pn or '/xl/drawings/drawing' in pn:
            ct_root.remove(ov)

    for k in list(files.keys()):
        if (re.match(r'xl/worksheets/sheet\d+\.xml$', k) or
            re.match(r'xl/worksheets/_rels/sheet\d+.*$', k) or
            re.match(r'xl/drawings/drawing\d+\.xml$', k) or
            re.match(r'xl/drawings/_rels/drawing\d+.*$', k) or
            k.startswith('xl/media/')):
            del files[k]

    j1     = '\u3000'.join(filter(None, [project.get('code',''), project.get('name','')])) or '（物件名未設定）'
    today  = date.today()
    wd     = f'{today.year}.{today.month}.{today.day}\u3000作業'
    img_no = 1

    for idx, part in enumerate(parts):
        sn    = idx + 1
        sname = re.sub(r'[\[\]:*?/\\]', '_', part['name'])[:31]

        sr    = etree.fromstring(tmpl_sheet_xml)
        j1_i  = find_or_add_ss(ss_root, j1)
        j2_i  = find_or_add_ss(ss_root, part['name'])
        j20_i = find_or_add_ss(ss_root, wd)
        old_i = find_or_add_ss(ss_root, ' ' + part.get('oldDesc', '■内：旧部品'))
        new_i = find_or_add_ss(ss_root, ' ' + part.get('newDesc', '■内：新部品'))

        set_cell_ss(sr, 'J1',  j1_i);  set_cell_ss(sr, 'J2',  j2_i)
        set_cell_ss(sr, 'J20', j20_i); set_cell_ss(sr, 'J10', old_i)
        set_cell_ss(sr, 'J35', old_i); set_cell_ss(sr, 'J32', new_i)
        set_cell_ss(sr, 'J56', new_i)
        set_cell_formula_val(sr, 'J25', part['name']); set_cell_formula_val(sr, 'J48', part['name'])
        set_cell_formula_val(sr, 'J43', wd);           set_cell_formula_val(sr, 'J66', wd)

        files[f'xl/worksheets/sheet{sn}.xml'] = etree.tostring(
            sr, xml_declaration=True, encoding='UTF-8', standalone=True)

        rid_map, rels = {}, []
        for slot, old_rid in [('before','rId1'),('compare','rId2'),('after','rId3')]:
            b64 = part.get('photos',{}).get(slot)
            if not b64: continue
            if ',' in b64: b64 = b64.split(',',1)[1]
            pil = Image.open(io.BytesIO(base64.b64decode(b64))).convert('RGB')
            pil.thumbnail((3000,3000), Image.LANCZOS)
            buf = io.BytesIO(); pil.save(buf, 'JPEG', quality=90)
            mname   = f'image{img_no}.jpeg'
            new_rid = f'rId{sn*10+len(rels)+1}'
            files[f'xl/media/{mname}'] = buf.getvalue()
            rid_map[old_rid] = new_rid
            rels.append((new_rid,
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                f'../media/{mname}'))
            img_no += 1

        files[f'xl/drawings/drawing{sn}.xml']           = replace_blip_rids(tmpl_drawing_xml, rid_map)
        files[f'xl/drawings/_rels/drawing{sn}.xml.rels']= make_rels_xml(rels)
        files[f'xl/worksheets/_rels/sheet{sn}.xml.rels']= make_rels_xml([('rId1',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
            f'../drawings/drawing{sn}.xml')])

        wb_rid = f'rId{100+sn}'
        etree.SubElement(sheets_el, f'{{{SS_NS}}}sheet',
            {'name':sname,'sheetId':str(sn),'state':'visible',f'{{{R_NS}}}id':wb_rid})
        etree.SubElement(wb_rels_root, f'{{{REL_NS}}}Relationship',
            {'Id':wb_rid,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
             'Target':f'worksheets/sheet{sn}.xml'})
        etree.SubElement(ct_root, f'{{{CT_NS}}}Override',
            {'PartName':f'/xl/worksheets/sheet{sn}.xml',
             'ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
        if rels:
            etree.SubElement(ct_root, f'{{{CT_NS}}}Override',
                {'PartName':f'/xl/drawings/drawing{sn}.xml',
                 'ContentType':'application/vnd.openxmlformats-officedocument.drawing+xml'})

    si_list = ss_root.findall(f'{{{SS_NS}}}si')
    ss_root.set('count', str(len(si_list))); ss_root.set('uniqueCount', str(len(si_list)))
    files['xl/sharedStrings.xml']       = etree.tostring(ss_root,      xml_declaration=True, encoding='UTF-8', standalone=True)
    files['xl/workbook.xml']            = etree.tostring(wb_root,      xml_declaration=True, encoding='UTF-8', standalone=True)
    files['xl/_rels/workbook.xml.rels'] = etree.tostring(wb_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['[Content_Types].xml']        = etree.tostring(ct_root,      xml_declaration=True, encoding='UTF-8', standalone=True)

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as oz:
        for fname, fdata in files.items(): oz.writestr(fname, fdata)
    out.seek(0)
    return out

# ══════════════════════════════════════════════════════
# ルート
# ══════════════════════════════════════════════════════
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        data    = request.get_json(force=True)
        project = data.get('project', {})
        parts   = data.get('parts', [])
        if not parts:
            return jsonify({'error': '部品が登録されていません'}), 400
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({'error': 'template.xlsx が見つかりません'}), 500
        buf   = generate_excel(project, parts)
        fname = re.sub(r'[^\w\u3040-\u9fff._-]', '_',
                       '_'.join(filter(None,[project.get('code',''), project.get('name','工事写真')]))) + '.xlsx'
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
