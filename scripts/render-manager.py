"""Actual Calc recalculation and PDF rendering, including full sheets at natural size.
Requires libreoffice-calc, python3-uno, poppler-utils; launch LO on port 2002 first.
Run with /usr/bin/python3, not a virtual environment interpreter.
"""
import json,pathlib,sys,uno
from com.sun.star.beans import PropertyValue

def prop(name,value):
 p=PropertyValue();p.Name=name;p.Value=value;return p
ctx=uno.getComponentContext();resolver=ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',ctx)
remote=resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
desktop=remote.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',remote)
root=pathlib.Path.cwd();out=root/'reports/manager-render';out.mkdir(exist_ok=True)
results=[]
paths=sys.argv[1:] or ['reports/EVO_GLOVE_CBD_Comparison_Manager_Format.xlsx','reports/All_Styles_CBD_Comparison_Manager_Format.xlsx']
for path in paths:
 doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(root/path)),'_blank',0,(prop('Hidden',True),))
 if not doc:raise RuntimeError('Could not load '+path)
 doc.calculateAll();sheets=[]
 for sheet in doc.Sheets:
  cursor=sheet.createCursor();cursor.gotoEndOfUsedArea(False);last=cursor.RangeAddress.EndRow
  errors=[];formulas=0;cached=[]
  for r in range(last+1):
   for c in range(25):
    cell=sheet.getCellByPosition(c,r)
    if cell.Formula.startswith('='):
     formulas+=1;cached.append({'row':r+1,'column':c+1,'formula':cell.Formula,'value':cell.Value,'text':cell.String})
    if cell.Error:errors.append({'row':r+1,'column':c+1,'error':cell.Error,'formula':cell.Formula})
  assert not errors,(sheet.Name,errors)
  assert sheet.DrawPage.Count==2,(sheet.Name,sheet.DrawPage.Count)
  sheets.append({'name':sheet.Name,'formulaCount':formulas,'errors':errors,'images':sheet.DrawPage.Count,'cells':cached})
 name=pathlib.Path(path).stem
 doc.storeToURL(uno.systemPathToFileUrl(str(out/(name+'-print.pdf'))),(prop('FilterName','calc_pdf_Export'),))
 doc.storeToURL(uno.systemPathToFileUrl(str(out/(name+'-recalculated.xlsx'))),(prop('FilterName','Calc MS Excel 2007 XML'),))
 # Render a separate natural-size PDF at 100%; retain the workbook's A4 setup above.
 styles=doc.StyleFamilies.getByName('PageStyles')
 for sheet in doc.Sheets:
  cursor=sheet.createCursor();cursor.gotoEndOfUsedArea(False);last=cursor.RangeAddress.EndRow
  page=styles.getByName(sheet.PageStyle);size=sheet.getCellRangeByPosition(0,0,24,last).Size
  page.ScaleToPagesX=0;page.ScaleToPagesY=0;page.ScaleToPages=0;page.PageScale=100
  page.Width=size.Width+page.LeftMargin+page.RightMargin+200
  page.Height=size.Height+page.TopMargin+page.BottomMargin+200
 doc.storeToURL(uno.systemPathToFileUrl(str(out/(name+'-100pct.pdf'))),(prop('FilterName','calc_pdf_Export'),))
 results.append({'file':path,'sheets':sheets});doc.close(True)
 print(name, len(sheets),'sheets:',sum(s['formulaCount'] for s in sheets),'formulas, no calculation errors, 2 images each')
(out/'recalculation.json').write_text(json.dumps(results,ensure_ascii=False,indent=2))
