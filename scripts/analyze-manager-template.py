import zipfile,pathlib,xml.etree.ElementTree as E,json,re,collections
source=next(pathlib.Path('reference').glob('*.xlsx'));z=zipfile.ZipFile(source)
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
strings=[''.join(si.itertext()) for si in E.fromstring(z.read('xl/sharedStrings.xml'))]
book=E.fromstring(z.read('xl/workbook.xml'));styles=E.fromstring(z.read('xl/styles.xml'))
styledefs=list(styles.find('m:cellXfs',ns));fills=list(styles.find('m:fills',ns));fonts=list(styles.find('m:fonts',ns));borders=list(styles.find('m:borders',ns))
report=['# Manager Format 원본 13시트 분석','',f'원본: `{source.name}`','', '신규 A:I / 비교 K:O / 기존 Q:Y, J/P 간격 열. Model Name C2:G2 및 S2:V2 병합. B4/R4 REFERENCES, 5행 열 제목, 본문 Arial 8pt, 주요 제목 Arial 10pt Bold. 그룹 병합과 빈 행은 STYLE별로 다름.','', '| STYLE | 그룹 블록 (원본 행) | 자재행 신규/기존 | 상대편 공란 신규/기존 | Total/FOB 행 |','|---|---|---:|---:|---|']
records=[]
for i,sh in enumerate(book.find('m:sheets',ns),1):
 root=E.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'));cells={};used=set()
 for c in root.findall('m:sheetData/m:row/m:c',ns):
  v=c.find('m:v',ns);value=v.text if v is not None else ''
  if c.get('t')=='s':value=strings[int(value)]
  cells[c.get('r')]={'value':value,'formula':c.findtext('m:f',None,ns),'style':int(c.get('s','0'))};used.add(int(c.get('s','0')))
 def value(address):return cells.get(address,{}).get('value','')
 groups=[];pairs=[];counts=[0,0];orphans=[0,0];summary=[]
 for r in range(6,101):
  a=value(f'A{r}').replace('\n',' ')
  if a in ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)']:groups.append([a,r])
  if a.lower() in ['total material cost','fob price']:summary.append([a,r])
  b,rr=value(f'B{r}'),value(f'R{r}')
  if not groups or summary:continue
  if b:counts[0]+=1
  if rr:counts[1]+=1
  if b and not rr:orphans[0]+=1
  if rr and not b:orphans[1]+=1
  if b and rr and b!=rr:pairs.append([r,b,rr])
 merges=[m.get('ref') for m in root.findall('m:mergeCells/m:mergeCell',ns)]
 printarea=next((n.text for n in book.findall('m:definedNames/m:definedName',ns) if n.get('name')=='_xlnm.Print_Area' and n.get('localSheetId')==str(i-1)),None)
 record={'name':sh.get('name'),'groups':groups,'materialCounts':counts,'oneSidedCounts':orphans,'substitutions':pairs,'summary':summary,'merges':merges,'sourcePrintArea':printarea,'pageSetup':root.find('m:pageSetup',ns).attrib,'views':[x.attrib for x in root.findall('m:sheetViews/m:sheetView',ns)],'columnWidths':[c.attrib for c in root.findall('m:cols/m:col',ns) if int(c.get('min'))<=25],'styles':[{'id':j,'cell':styledefs[j].attrib,'font':E.tostring(fonts[int(styledefs[j].get('fontId','0'))]).decode(),'fill':E.tostring(fills[int(styledefs[j].get('fillId','0'))]).decode(),'border':E.tostring(borders[int(styledefs[j].get('borderId','0'))]).decode()} for j in sorted(used)]}
 records.append(record)
 report.append(f"| {sh.get('name')} | "+', '.join(f'{g}: {r}' for g,r in groups)+f" | {counts[0]}/{counts[1]} | {orphans[0]}/{orphans[1]} | "+'/'.join(str(r) for _,r in summary)+' |')
report+=['','## 공통 규칙 및 예외','', '- EVO: COOL MESH 4540 NYLON ↔ Q-SPAN, RHEON NUCKLE ↔ TPR KNUCKLE의 명시적인 같은 행 대체 사례. 전체 시트에 J RING PLASTIC SNAP HOOK ↔ J RING과 시즌/공급처 설명이 다른 LABEL/PACKAGING 사례가 있다. 이들은 템플릿 관계를 확인한 테스트 자료이며, 출력기는 별도 문자열 매칭을 하지 않는다.', '- KINETIC 2 Sonic/sub는 원본 자재 수·빈칸 배치·Total·FOB가 서로 다르다. LITE 계열 TRIMS에는 신규가 비어 있는 기존 전용 자재가 있다. PRO LITE DBK 2 원본 시트명 끝에 불필요한 닫는 괄호가 있다.', '- 노란 소계가 각 블록 뒤에 있고 그룹별 빈 행 수는 가변이다. SPECIAL PROCESS에는 비용 비교가 없는 목록과 원본 Remark가 있다. Labor/Overhead/Profit은 원본에서도 비어 있는 경우가 있으며 생성기는 값을 추정하지 않는다.', '- 원본 중앙에는 직접 뺄셈과 비율 수식이 있다. 하단 일부 IFERROR 수식은 복사하지 않고 분모 0/공란을 명시적으로 처리한다.', '- 색상: 헤더 99CCFF, 모델명 C0C0C0, 소계 FFFF00, 하단 CCFFFF, 변경 FFCCFF. 신규 모델명 빨강/기존 검정. 음수 금액은 빨간 괄호. 원본은 비율 서식이 일관되지 않아 신규 생성기는 음수 비율도 빨간 괄호로 통일한다.', '- 로고: 모든 시트 2개, A1/Q1 부근, 약 106×36 px의 drawing extent. xl/media/image1.emf 하나를 공유한다. EMF 기록은 HEADER/STRETCHDIBITS/EOF로, 실제 내용은 143×44 32bit 비트맵이다. 비트맵을 추출하여 흰 배경을 투명화한 PNG를 재사용한다. 원본에 없는 고해상도 디테일은 생성하지 않는다.', '- 인쇄: 원본 오른쪽만 지정된 범위는 복제하지 않는다. 신규 출력은 A1:Y마지막행, A4 가로, 가로 1페이지/세로 자동. 전체 시트 렌더링도 별도 검사한다.', '- 검증에서 원본 Total과 기존 비교 계산의 차이를 발견: SEWING THREAD의 자재 Extended Cost 없이 별도 소계로만 입력한 $0.25는 기존 공통 계산이 참고용으로 제외한다. 신규 시즌 원본 Total에는 그룹 반올림 차이도 있다. 기존 계산 보존 요구에 따라 Manager 비교 Total은 기존 상세와 동일하며, 원본 소계/Total 및 차이를 별도로 명시한다. 원본 인쇄 Total과 비교 Total이 같다는 불변 조건은 이 자료에서 성립하지 않는다. 사용자에게 계산 정책 선택을 질문했다.','', '시트별 열 너비, 모든 사용 스타일의 font/fill/border, 병합, 원본 인쇄 범위, 대체 자재 목록: `manager-template-structure.json`.']
pathlib.Path('reports/manager-template-structure.json').write_text(json.dumps(records,ensure_ascii=False,indent=2));pathlib.Path('reports/manager-template-analysis.md').write_text('\n'.join(report)+'\n')
print('\n'.join(report[:19]))
