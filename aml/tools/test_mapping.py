"""Independent checks of the generated cell map against both CBI XML examples.
Does not execute VBA. Does not modify bank files. Outputs test XML only under tools/qa.
"""
from pathlib import Path
import json, datetime, decimal, re
import openpyxl
from lxml import etree as E

BASE=Path(__file__).resolve().parents[1]
SRC=Path(r'C:/Users/wkowa/Downloads/funds-administrators-files-v1-0')
QA=BASE/'tools'/'qa'
QA.mkdir(exist_ok=True)
mapping=json.loads((BASE/'tools'/'mapping.json').read_text())
schema_doc=E.parse(str(SRC/'Fund Admin - XSD v1.0.xsd'))
schema=E.XMLSchema(schema_doc)
ns={'x':'http://www.w3.org/2001/XMLSchema'}
uri=schema_doc.getroot().get('targetNamespace')
types={}
for g in schema_doc.xpath('/x:schema/x:element', namespaces=ns):
    types[g.get('name')]={a.get('name'):a.get('type').split(':')[-1] for a in g.xpath('.//x:attribute',namespaces=ns)}

def value(v, typ):
    if v is None: raise ValueError('Missing value')
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if typ in ['NoNegInt20','Percentage','Decimal206']:
        d=decimal.Decimal(str(v).replace(',','.'))
        s=format(d,'f')
        if '.' in s: s=s.rstrip('0').rstrip('.')
        return s
    return str(v).strip()

def build(wb):
    root=E.Element('{'+uri+'}FundAdmin',nsmap={None:uri})
    groups={}
    for ref in schema_doc.xpath("/x:schema/x:element[@name='FundAdmin']//x:element/@ref",namespaces=ns):
        group=ref.split(':')[1]
        groups[group]=E.SubElement(root,'{'+uri+'}'+group)
    for m in mapping['scalars']:
        node=E.SubElement(groups[m['group']],'{'+uri+'}'+m['element'])
        node.set('Identifier',m['ref'])
        node.set(m['attribute'],value(wb[m['sheet']][m['cell']].value,types[m['group']][m['attribute']]))
    for group,sheet,header,first,last,cols,refs in mapping['tables']:
        sh=wb[sheet]; cols=cols.split(','); refs=refs.split(',')
        tag=schema_doc.xpath(f"/x:schema/x:element[@name='{group}']//x:element/@name",namespaces=ns)[0]
        for r in range(first,(last or sh.max_row)+1):
            vals=[sh[f'{c}{r}'].value for c in cols]
            if not any(v is not None for v in vals) and not last: continue
            node=E.SubElement(groups[group],'{'+uri+'}'+tag)
            for ref,v in zip(refs,vals): node.set(ref,value(v,types[group][ref]))
    return root

summary=[]
for name,xml_name in [
 ('Funds Admin - Example Completed file ILLUSTRATIVE ONLYv1.0.xlsx','Fund Admin - Example Completed XML v1.0.xml'),
 ('Funds Admin - Example Completed Blank file ILLUSTRATIVE ONLY v1.0.xlsx','Fund Admin - Example Blank XML v1.0.xml')]:
    wb=openpyxl.load_workbook(SRC/name,data_only=True)
    root=build(wb)
    other=E.parse(str(SRC/xml_name)).getroot()
    diffs=[]
    for a,b in zip(root,other):
        assert a.tag==b.tag and len(a)==len(b), (a.tag,len(a),len(b))
        for i,(aa,bb) in enumerate(zip(a,b)):
            assert aa.tag==bb.tag
            assert set(aa.attrib)==set(bb.attrib)
            for k in aa.attrib:
                if aa.get(k)!=bb.get(k): diffs.append((E.QName(a).localname,aa.get('Identifier',str(i+1)),k,aa.get(k),bb.get(k)))
    summary.append({'workbook':name,'rows':sum(len(g) for g in root),'differences_vs_bank_xml':diffs,'valid_before_repairs':schema.validate(root),'errors':[str(e) for e in schema.error_log]})
    print(name, 'rows',sum(len(g) for g in root),'differences',len(diffs),'XSD errors',len(schema.error_log))
    for m in mapping['scalars']:
        h=0
        for c in str(wb[m['sheet']][m['label_cell']].value or ''): h=(h*31+ord(c))%2147483647
        assert str(h)==m['label_hash'], ('Layout label mismatch',name,m['ref'],m['label_cell'])
    # Corrections are exclusively in synthetic test memory, never in bank sources.
    for n in root.iter():
        if n.get('Identifier') in ['LAG','LAO']: n.set('LEI','00000000000000000000')
        if 'Blank' in name and n.get('Identifier') in ['LGG','LGH','LPL']: n.set('EnumerationYesNo','No')
        if n.get('Identifier')=='LIX': n.set('EnumerationYesNoNA','N/A')
        if n.get('Identifier')=='LIY' and n.get('EnumerationMonitoringApproach')=='N/A': n.set('EnumerationMonitoringApproach','No_transaction_monitoring')
    assert schema.validate(root), str(schema.error_log)
    # Force explicit closing tags, matching exporter contract.
    text=E.tostring(root,encoding='unicode')
    text=re.sub(r'<([A-Za-z0-9]+)([^<>]*?)/>',r'<\1\2></\1>',text)
    prefix='completed' if 'Blank' not in name else 'blank'
    (QA/(prefix+'_corrected_test.xml')).write_text('<?xml version="1.0" encoding="utf-8"?>\n'+text,encoding='utf-8')
    assert schema.validate(E.fromstring(text.encode()))

# Every scalar identifier and every table attribute must be represented exactly once.
for g in ['A1201A','A1202C','A1203A']:
    enum_type=types[g]['Identifier']
    identifiers=set(schema_doc.xpath(f"/x:schema/x:simpleType[@name='{enum_type}']//x:enumeration/@value",namespaces=ns))
    mapped=[m['ref'] for m in mapping['scalars'] if m['group']==g]
    assert len(mapped)==len(set(mapped)) and set(mapped)==identifiers
for g,_,_,_,_,_,refs in mapping['tables']:
    assert set(refs.split(','))==set(types[g])

(QA/'mapping_test_results.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print('PASS: complete scalar/table coverage; both example layouts produce XSD-valid XML after explicit synthetic corrections.')
