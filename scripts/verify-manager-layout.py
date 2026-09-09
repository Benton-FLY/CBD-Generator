import xml.etree.ElementTree as E,re,json,pathlib
results=[]
for kind in ['100pct','print']:
 path=pathlib.Path(f'reports/manager-render/{kind}-text.html');root=E.parse(path).getroot()
 pages=root.findall('.//{*}page');assert len(pages)==13,(kind,len(pages))
 for i,page in enumerate(pages,1):
  words=[{'text':w.text or '',**{k:float(w.get(k)) for k in ['xMin','xMax','yMin','yMax']}} for w in page.findall('.//{*}word')]
  text=' '.join(w['text'] for w in words)
  assert not re.search(r'#VALUE!|#REF!|#DIV/0!|#NAME\?|#N/A|NaN|#{3,}',text),(kind,i,'error text')
  assert text.count('FOB')>=2,(kind,i,'missing left/right FOB')
  assert text.count('Material')>=2,(kind,i,'missing left/right header')
  width,height=float(page.get('width')),float(page.get('height'))
  assert all(w['xMin']>=0 and w['xMax']<=width+.1 and w['yMin']>=0 and w['yMax']<=height+.1 for w in words),(kind,i,'clipped text')
  overlap=[]
  numeric=lambda t:bool(re.fullmatch(r'\(?-?[\d,.]+%?\)?',t))
  for a,w in enumerate(words):
   if not numeric(w['text']):continue
   for v in words[a+1:]:
    if min(w['xMax'],v['xMax'])-max(w['xMin'],v['xMin'])>0.5 and min(w['yMax'],v['yMax'])-max(w['yMin'],v['yMin'])>0.5:overlap.append([w['text'],v['text']])
  assert not overlap,(kind,i,overlap)
  results.append({'render':kind,'sheet':i,'pageWidth':width,'pageHeight':height,'words':len(words),'numericTextOverlaps':overlap,'leftRightFobPresent':True})
pathlib.Path('reports/manager-layout-verification.json').write_text(json.dumps(results,indent=2));print('PASS: 13 natural-size + 13 A4 pages; all text in bounds, left/right FOB present, no numeric overlaps or error tokens.')
