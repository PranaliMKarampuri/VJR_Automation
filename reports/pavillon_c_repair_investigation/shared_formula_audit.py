from __future__ import annotations
import json, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from openpyxl.utils.cell import range_boundaries, get_column_letter

ROOT=Path(__file__).parent.parent.parent
OUT=Path(__file__).parent
TEMPLATE=ROOT/'source_files/Ventilation Jean Roy excel.xlsx'
GENERATED=ROOT/'output/jobs/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_8bc5caddc1/TAKEOFF_Plans_Pav._C_-r_vis_add1_8_VJR.xlsx'
M={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def audit(z, part):
    root=ET.fromstring(z.read(part)); cells={c.attrib['r']:c for c in root.findall('.//m:sheetData/m:row/m:c',M)}
    masters=[]
    for ref,c in cells.items():
        f=c.find('m:f',M)
        if f is not None and f.attrib.get('t')=='shared' and f.attrib.get('ref'):
            si=f.attrib.get('si'); x1,y1,x2,y2=range_boundaries(f.attrib['ref']); missing=[]; nonformula=[]
            for row in range(y1,y2+1):
                for col in range(x1,x2+1):
                    cr=f'{get_column_letter(col)}{row}'; cc=cells.get(cr); ff=cc.find('m:f',M) if cc is not None else None
                    if ff is None or ff.attrib.get('t')!='shared' or ff.attrib.get('si')!=si:
                        nonformula.append(cr)
            masters.append({'master':ref,'si':si,'ref':f.attrib['ref'],'formula':f.text,'range_cell_count':(x2-x1+1)*(y2-y1+1),'non_shared_formula_cells_in_ref':nonformula})
    follower_cells={f.attrib.get('si'):[] for c in cells.values() for f in c.findall('m:f',M) if f.attrib.get('t')=='shared' and not f.attrib.get('ref')}
    for ref,c in cells.items():
        for f in c.findall('m:f',M):
            if f.attrib.get('t')=='shared' and not f.attrib.get('ref'):
                follower_cells.setdefault(f.attrib.get('si'),[]).append(ref)
    follower_ids=set(follower_cells)
    master_ids={x['si'] for x in masters}
    return {'shared_formula_master_count':len(masters),'shared_formula_follower_ids_without_master':sorted(follower_ids-master_ids,key=lambda x:int(x or 0)),
            'follower_cells_by_id':follower_cells, 'masters_with_non_shared_cells_in_ref':[x for x in masters if x['non_shared_formula_cells_in_ref']], 'all_masters':masters}

res={}
with zipfile.ZipFile(TEMPLATE) as t,zipfile.ZipFile(GENERATED) as g:
 for p in ['xl/worksheets/sheet2.xml','xl/worksheets/sheet3.xml']:
  res[p]={'template':audit(t,p),'generated':audit(g,p)}
out=OUT/'shared_formula_audit.json'; out.write_text(json.dumps(res,indent=2,ensure_ascii=False),encoding='utf-8')
for p,x in res.items():
 print(p)
 for typ in ('template','generated'):
  y=x[typ]; print(typ,'groups',y['shared_formula_master_count'],'orphan ids',y['shared_formula_follower_ids_without_master'],'groups with holes',len(y['masters_with_non_shared_cells_in_ref']))
  if typ=='generated':
   template_masters={m['si']:m for m in x['template']['all_masters']}
   for si in y['shared_formula_follower_ids_without_master']:
    print(' ORPHAN GROUP',si,'original master',template_masters.get(si),'generated followers',y['follower_cells_by_id'].get(si))
  for m in y['masters_with_non_shared_cells_in_ref'][:25]: print(' ',m['master'],m['ref'],'formula=',m['formula'],'gaps=',m['non_shared_formula_cells_in_ref'][:25])
print('saved',out)
