"""从备份 + 损坏文件恢复 Whole Market.xlsx。

用法:
    python recover_market.py <备份文件> <损坏文件> [输出文件]

示例:
    python recover_market.py backup.xlsx corrupted.xlsx recovered.xlsx
"""
import zipfile, re, io, struct, zlib, sys
import pandas as pd

BAK = sys.argv[1] if len(sys.argv) > 1 else "Whole Market.xlsx.bak"
CORRUPTED = sys.argv[2] if len(sys.argv) > 2 else "Whole Market.xlsx"
DST = sys.argv[3] if len(sys.argv) > 3 else CORRUPTED


def read_zip_safe(path):
    """读取ZIP中所有有效条目，跳过CRC错误的。"""
    result = {}
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            try:
                result[info.filename] = z.read(info.filename)
            except Exception:
                pass
    return result


def scan_local_headers(path):
    """从损坏的ZIP中扫描所有本地文件头，提取条目。"""
    with open(path, 'rb') as f:
        data = f.read()
    entries = {}
    pos = 0
    while pos < len(data) - 30:
        if data[pos:pos+4] != b'PK\x03\x04':
            pos += 1
            continue
        compression = struct.unpack('<H', data[pos+8:pos+10])[0]
        comp_size = struct.unpack('<I', data[pos+18:pos+22])[0]
        name_len = struct.unpack('<H', data[pos+26:pos+28])[0]
        extra_len = struct.unpack('<H', data[pos+28:pos+30])[0]
        name = data[pos+30:pos+30+name_len].decode('latin-1')
        start = pos + 30 + name_len + extra_len
        raw = data[start:start+comp_size]
        if compression == 8:
            content = zlib.decompress(raw, -zlib.MAX_WBITS)
        else:
            content = raw
        entries[name] = content
        pos = start + comp_size
    return entries


def extract_sheet_codes(xml):
    """提取sheet XML中前10个股代码作为指纹。"""
    return tuple(re.findall(r'<t>(\d+\.\w+)</t>', xml.decode('latin-1'))[:10])


# ── 1. 读取备份 ──
bak_all = read_zip_safe(BAK)
bak_xml = bak_all['xl/workbook.xml'].decode('utf-8')

bak_date_map = {}
for m in re.finditer(r'name="(\d+)"[^>]*sheetId="(\d+)"', bak_xml):
    date_name = m.group(1)
    sid = m.group(2)
    fn = f'xl/worksheets/sheet{sid}.xml'
    if fn in bak_all:
        bak_date_map[date_name] = bak_all[fn]

print(f"备份装载: {len(bak_date_map)} sheets: {sorted(bak_date_map)}")

# ── 2. 读取损坏文件 ──
corrupted = scan_local_headers(CORRUPTED)
corr_sheets = {k: v for k, v in corrupted.items() if k.startswith('xl/worksheets/sheet')}
print(f"损坏文件: {len(corr_sheets)} sheet 条目")

# ── 3. 指纹匹配 ──
bak_fingerprints = {d: extract_sheet_codes(x) for d, x in bak_date_map.items()}
corr_fingerprints = {s: extract_sheet_codes(x) for s, x in corr_sheets.items()}

found = {}
for date_name, codes in bak_fingerprints.items():
    for sname, ccodes in corr_fingerprints.items():
        if codes == ccodes and len(codes) >= 5:
            found[date_name] = sname
            break

print(f"\n精确匹配: {len(found)}")
already = set(found.keys())

# ── 4. 未匹配的损坏sheet ──
matched_set = set(found.values())
unmatched = [(s, c) for s, c in corr_sheets.items() if s not in matched_set]

# 通过行数排序（减少的顺序 = 时间往后）
unmatched_rows = []
for sname, xml in unmatched:
    text = xml.decode('latin-1')
    rows = len(re.findall(r'<row r="(\d+)"', text)) - 1
    unmatched_rows.append((rows, sname, xml))

unmatched_rows.sort(key=lambda x: x[0], reverse=True)

new_dates = {}
for i, (rows, sname, xml) in enumerate(unmatched_rows):
    new_dates[f"sheet{i+1}"] = xml
    print(f"  新条目 {sname} ({rows} rows)")

# ── 5. 构建最终workbook ──
all_dates = sorted(bak_date_map.keys()) + sorted(new_dates.keys())
print(f"\n最终 {len(all_dates)} sheets")

overrides = []
for i in range(len(all_dates)):
    overrides.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')

ct_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{"".join(overrides)}
</Types>'''

workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
{"".join(f'<sheet name="{d}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, d in enumerate(all_dates))}
</sheets>
</workbook>'''

rels_parts = '\n'.join(
    f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
    for i in range(len(all_dates))
)
rels_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels_parts}
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
<dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">quant</dc:creator>
<dcterms:created xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">2026-05-27T00:00:00Z</dcterms:created>
</cp:coreProperties>'''

# ── 6. 写入 ──
output = io.BytesIO()
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml', ct_xml)
    zf.writestr('_rels/.rels', root_rels)
    zf.writestr('docProps/app.xml', bak_all['docProps/app.xml'])
    zf.writestr('docProps/core.xml', core_xml)
    zf.writestr('xl/workbook.xml', workbook_xml)
    zf.writestr('xl/_rels/workbook.xml.rels', rels_xml)
    zf.writestr('xl/styles.xml', bak_all['xl/styles.xml'])
    zf.writestr('xl/theme/theme1.xml', bak_all['xl/theme/theme1.xml'])
    for i, d in enumerate(all_dates):
        xml = bak_date_map.get(d) or new_dates.get(d)
        zf.writestr(f'xl/worksheets/sheet{i+1}.xml', xml)

with open(DST, 'wb') as f:
    f.write(output.getvalue())

print(f"\n已写入 {DST}")

# ── 7. 验证 ──
try:
    xls = pd.ExcelFile(DST)
    print(f"验证成功: {len(xls.sheet_names)} sheets: {xls.sheet_names}")
    for sn in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sn)
        print(f"  {sn}: {len(df)} rows")
except Exception as e:
    print(f"验证失败: {e}")
