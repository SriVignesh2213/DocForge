import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn
import os

tp = r'c:\Users\SRI VIGNESH\Downloads\DocForge\templates'
for f in os.listdir(tp):
    if not f.endswith('.docx'):
        continue
    d = Document(os.path.join(tp, f))
    print('='*60)
    print('FILE: ' + f)
    print('  Total Sections: ' + str(len(d.sections)))
    for i, s in enumerate(d.sections):
        sectPr = s._sectPr
        cols = sectPr.findall(qn('w:cols'))
        num = cols[0].get(qn('w:num')) if cols else '1(default)'
        hrefs = sectPr.findall(qn('w:headerReference'))
        frefs = sectPr.findall(qn('w:footerReference'))
        print('  Section ' + str(i) + ': cols=' + str(num))
        RI = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        hlist = [(h.get(RI), h.get(qn('w:type'))) for h in hrefs]
        flist = [(h.get(RI), h.get(qn('w:type'))) for h in frefs]
        print('    headerRefs: ' + str(hlist))
        print('    footerRefs: ' + str(flist))
        pgMar = sectPr.find(qn('w:pgMar'))
        if pgMar is not None:
            top = pgMar.get(qn('w:top'))
            bot = pgMar.get(qn('w:bottom'))
            left = pgMar.get(qn('w:left'))
            right = pgMar.get(qn('w:right'))
            hdr = pgMar.get(qn('w:header'))
            ftr = pgMar.get(qn('w:footer'))
            print('    pgMar: top='+str(top)+' bot='+str(bot)+' left='+str(left)+' right='+str(right)+' hdr='+str(hdr)+' ftr='+str(ftr))
