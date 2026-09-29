"""Generate the VBA field map from the unmodified CBI v1.0 sources. Read only XLSX."""
from pathlib import Path
import csv, json, re, shutil
import openpyxl
from lxml import etree as ET

BASE = Path(__file__).resolve().parents[1]
SRC = Path(r'C:/Users/wkowa/Downloads/funds-administrators-files-v1-0')
OUT = BASE / 'output' / 'CBI_XML_Macro'
OUT.mkdir(parents=True, exist_ok=True)
NS = {'x': 'http://www.w3.org/2001/XMLSchema'}
xsd = ET.parse(str(SRC / 'Fund Admin - XSD v1.0.xsd'))
wb = openpyxl.load_workbook(SRC / 'Funds Administrators - Risk Evaluation Questionnaire v1.0.xlsx')
instructions = {str(r[1]).strip(): r[3] for r in list(openpyxl.load_workbook(next(SRC.glob('*Individual*.xlsx'))).active.values)[3:] if r[1]}
locations = {}
for sheet in wb:
    if sheet.title == 'Drop-Downs':
        continue
    for row in sheet:
        for cell in row:
            if isinstance(cell.value, str) and re.fullmatch(r'L[A-Z]{2}', cell.value):
                locations.setdefault(cell.value, []).append((sheet.title, cell.coordinate))

def hash_label(text):
    h = 0
    for c in str(text or ''):
        h = (h * 31 + ord(c)) % 2147483647
    return str(h)

scalars = []
blank = ET.parse(str(SRC / 'Fund Admin - Example Blank XML v1.0.xml'))
for group in blank.getroot():
    for item in group:
        if 'Identifier' not in item.attrib:
            continue
        ref = item.get('Identifier')
        assert len(locations[ref]) == 1, ref
        sheet, addr = locations[ref][0]
        attr = next(a for a in item.attrib if a != 'Identifier')
        attrs = xsd.xpath(f"/x:schema/x:element[@name='{ET.QName(group).localname}']//x:attribute/@name", namespaces=NS)
        tail = instructions[ref].split('Variable name:')[-1].strip()
        declared = next((a for a in sorted(attrs, key=len, reverse=True) if tail.startswith(a)), None)
        assert declared == attr, (ref, declared, attr)
        row = wb[sheet][addr].row
        scalars.append(dict(ref=ref, sheet=sheet, cell=addr, group=ET.QName(group).localname,
                            element=ET.QName(item).localname, attribute=attr,
                            label_cell=f'B{row}', label_hash=hash_label(wb[sheet][f'B{row}'].value)))
assert len(scalars) == 456

tables = [
    ('A1201B','8.2 General',28,29,58,'B,D,E,F,G','LAR,LAS,LAT,LAU,LAV'),
    ('A1202A','8.3.1 Inherent Risk Funds',4,5,0,'A,B,C,D,E,F,G,H,I,J,K,L,M,N,O','LBB,LBC,LBD,LBE,LBF,LBG,LBH,LBI,LBJ,LBK,LBL,LBM,LBN,LBO,LBP'),
    ('A1202B','8.3.2 Inherent Risks Investors',7,8,34,'B,D','LBQ,LBR'),
    ('A1204','8.5 Physical Presence',3,4,0,'A,B,C,D,E','LSK,LSL,LSM,LSN,LSO'),
    ('A1205','8.6 Residence & Establishment',4,5,0,'A,B,C,D,E,F,G,H,I','LSP,LSQ,LSR,LSS,LST,LSU,LSV,LSW,LSX'),
    ('A1206','8.7 Beneficial Owner',4,5,0,'A,B','LSY,LSZ'),
    ('A1207','8.8 Politically Exposed Person',3,4,0,'A,B,C','LTA,LTB,LTC'),
    ('A1208','8.9 Geography of Funds Flow',4,5,0,'A,B,C,D,E,F,G,H,I,J,K,L,M','LTD,LTE,LTF,LTG,LTH,LTI,LTJ,LTK,LTL,LTM,LTN,LTO,LTP'),
    ('A1209','8.10 Transaction Monitoring',3,4,0,'A,B,C,D,E','LTQ,LTR,LTS,LTT,LTU'),
]
for group, sheet, header, first, last, cols, refs in tables:
    for col, ref in zip(cols.split(','), refs.split(',')):
        assert wb[sheet][f'{col}{header}'].value == ref

lines = ['Private Sub LoadScalarMap(ByRef records As Collection)']
for i in range(0, len(scalars), 60):
    lines.append(f'    MapPart{i//60} records')
lines.append('End Sub\n')
for i in range(0, len(scalars), 60):
    lines.append(f'Private Sub MapPart{i//60}(ByRef records As Collection)')
    for r in scalars[i:i+60]:
        data = '|'.join(r[k] for k in ['ref','sheet','cell','group','element','attribute','label_cell','label_hash'])
        lines.append(f'    records.Add "{data}"')
    lines.append('End Sub\n')
lines.append('Private Sub LoadTableMap(ByRef records As Collection)')
for t in tables:
    lines.append('    records.Add "' + '|'.join(map(str,t)) + '"')
lines.append('End Sub\n')
core = (BASE / 'tools' / 'CBI_REQ_Export_core.bas').read_text(encoding='utf-8')
(OUT / 'CBI_REQ_Export.bas').write_bytes((core + '\n' + '\n'.join(lines)).replace('\r\n','\n').replace('\n','\r\n').encode('ascii'))
shutil.copy2(SRC / 'Fund Admin - XSD v1.0.xsd', OUT)
with (OUT / 'mapowanie_pol.csv').open('w', newline='', encoding='utf-8-sig') as f:
    writer = csv.DictWriter(f, fieldnames=list(scalars[0]), delimiter=';')
    writer.writeheader()
    writer.writerows(scalars)
with (OUT / 'mapowanie_tabel.csv').open('w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f, delimiter=';')
    writer.writerow(['group','sheet','header_row','first_data_row','last_data_row_0_means_dynamic','columns','refs'])
    writer.writerows(tables)
(BASE/'tools'/'mapping.json').write_text(json.dumps({'scalars':scalars,'tables':tables}, indent=2), encoding='utf-8')
print(f'Generated {len(scalars)} scalar mappings and {len(tables)} tables in {OUT}')
