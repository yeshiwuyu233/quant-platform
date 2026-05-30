"""从损坏的 ZIP/XLSX 中提取所有条目并重建为有效文件。

用法:
    python fix_excel.py <损坏文件> [输出文件]

示例:
    python fix_excel.py corrupted.xlsx rebuilt.xlsx
"""
import struct
import zlib
import os
import sys
import zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else "corrupted.xlsx"
DST = sys.argv[2] if len(sys.argv) > 2 else SRC.rsplit('.', 1)[0] + "_rebuilt.xlsx"

with open(SRC, "rb") as f:
    data = f.read()

# 扫描所有本地文件头
pos = 0
entries = []
while pos < len(data) - 30:
    if data[pos:pos+4] != b'PK\x03\x04':
        pos += 1
        continue
    compression = struct.unpack('<H', data[pos+8:pos+10])[0]
    comp_size = struct.unpack('<I', data[pos+18:pos+22])[0]
    uncomp_size = struct.unpack('<I', data[pos+22:pos+26])[0]
    name_len = struct.unpack('<H', data[pos+26:pos+28])[0]
    extra_len = struct.unpack('<H', data[pos+28:pos+30])[0]
    name = data[pos+30:pos+30+name_len].decode('latin-1')
    start = pos + 30 + name_len + extra_len
    raw = data[start:start+comp_size]

    entries.append((name, compression, raw, uncomp_size))
    pos = start + comp_size

print(f"找到 {len(entries)} 个条目")

# 解压所有条目
extracted = []
for name, comp, raw, uncomp_size in entries:
    if comp == 0:
        content = raw[:uncomp_size]
    elif comp == 8:
        content = zlib.decompress(raw, -zlib.MAX_WBITS)
    else:
        print(f"不支持压缩方式 {comp} for {name}")
        continue
    extracted.append((name, content))
    print(f"  提取: {name} ({len(content)} bytes)")

# 检查是否缺失必要文件
has_content_types = any(n == '[Content_Types].xml' for n, _ in extracted)
has_rels = any(n == '_rels/.rels' for n, _ in extracted)
has_workbook = any(n.startswith('xl/workbook') for n, _ in extracted)

if not has_content_types:
    # 从工作表推断
    sheet_names = [n for n, _ in extracted if n.startswith('xl/worksheets/sheet')]
    sheet_count = len(sheet_names)
    print(f"\n缺失 [Content_Types].xml，自动生成 ({sheet_count} sheets)")

    overrides = []
    overrides.append(f'<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    overrides.append(f'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    overrides.append(f'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
    overrides.append(f'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    for sn in sheet_names:
        overrides.append(f'<Override PartName="/{sn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')

    ct_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{"".join(overrides)}</Types>'
    extracted.insert(0, ('[Content_Types].xml', ct_xml.encode('utf-8')))

if not has_rels:
    rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    extracted.insert(1, ('_rels/.rels', rels_xml.encode('utf-8')))

if not has_workbook:
    sheet_names_sorted = sorted([n for n, _ in extracted if n.startswith('xl/worksheets/sheet')])
    print(f"缺失 workbook.xml，自动生成 ({len(sheet_names_sorted)} sheets)")

    sheets_xml = []
    for i, sn in enumerate(sheet_names_sorted):
        r_id = i + 1
        # Extract sheet name from sheet file if possible
        sheet_content = dict(extracted).get(sn, b'')
        if sheet_content:
            import re
            sheet_name = re.search(r'<sheetPr[^>]*/>\s*<dimension[^>]*/>', sheet_content.decode('latin-1'))
        sid = i + 1
        # Use MMDD format based on the date - but we don't know the actual date
        # Just name them sequentially
        sheets_xml.append(f'<sheet name="sheet{sid}" sheetId="{sid}" r:id="rId{r_id}"/>')

    # Create xl/workbook.xml
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{"".join(sheets_xml)}</sheets>
</workbook>'''
    extracted.append(('xl/workbook.xml', wb_xml.encode('utf-8')))

    # Create xl/_rels/workbook.xml.rels
    wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, sn in enumerate(sheet_names_sorted):
        r_id = i + 1
        wb_rels.append(f'<Relationship Id="rId{r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{os.path.basename(sn)}"/>')
    wb_rels.append('<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    wb_rels.append('</Relationships>')
    extracted.append(('xl/_rels/workbook.xml.rels', "".join(wb_rels).encode('utf-8')))

# 追加缺失的 sharedStrings.xml（如果没有的话）
has_ss = any('sharedStrings' in n for n, _ in extracted)
if not has_ss:
    ss_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>'
    extracted.append(('xl/sharedStrings.xml', ss_xml.encode('utf-8')))

# 重建 ZIP
with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zf:
    for name, content in extracted:
        zf.writestr(name, content)

print(f"\n重建完成: {DST}")
print(f"总条目: {len(extracted)}")

# 验证
try:
    xls = __import__('pandas').ExcelFile(DST)
    print(f"验证成功: {len(xls.sheet_names)} sheets: {xls.sheet_names}")
except Exception as e:
    print(f"验证失败: {e}")
