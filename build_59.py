# -*- coding: utf-8 -*-
"""master md -> .docx (รูปแบบเอกสารราชการ ตาม BDI: TH Sarabun New 16pt, หัวข้อดำ, ตารางหัวเทา)"""
import io, re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

md = io.open('59_coe-creative-fullproposal.md', encoding='utf-8').read()
FONT='TH Sarabun New'; BLUE=(0x1F,0x4E,0x78); BLACK=(0,0,0); GRAY=(0x44,0x44,0x44)
doc=Document()
st=doc.styles['Normal']; st.font.name=FONT; st.font.size=Pt(16)
st._element.rPr.rFonts.set(qn('w:cs'),FONT); st._element.rPr.rFonts.set(qn('w:eastAsia'),FONT)
def cs(r):
    r.font.name=FONT; rf=r._element.get_or_add_rPr().get_or_add_rFonts()
    rf.set(qn('w:cs'),FONT); rf.set(qn('w:ascii'),FONT); rf.set(qn('w:hAnsi'),FONT)
def add_runs(p,text,bold=False,size=16,color=None):
    for i,part in enumerate(text.split('**')):
        if part=='':continue
        r=p.add_run(part); r.bold=bold or (i%2==1); r.font.size=Pt(size)
        if color:r.font.color.rgb=RGBColor(*color)
        cs(r)
def para(text='',bold=False,size=16,align=None,color=None,after=6):
    p=doc.add_paragraph()
    if align:p.alignment=align
    p.paragraph_format.space_after=Pt(after)
    if text:add_runs(p,text,bold,size,color)
    return p
def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd')
    sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),hexc); tcPr.append(sh)
def emit_table(rows):
    data=[r for r in rows if not re.match(r'^[\s:|-]+$','|'.join(r))]
    if not data:return
    n=max(len(r) for r in data)
    t=doc.add_table(rows=0,cols=n); t.style='Table Grid'
    for ri,row in enumerate(data):
        c=t.add_row().cells
        for ci in range(n):
            v=row[ci].strip() if ci<len(row) else ''
            c[ci].text=''
            add_runs(c[ci].paragraphs[0], v, bold=(ri==0), size=14, color=(BLACK if ri==0 else None))
            if ri==0: shade(c[ci],'D9D9D9')
    para('',after=4)

lines=md.split('\n'); i=0; in_cover=True
while i<len(lines):
    s=lines[i].strip()
    if s.startswith('## (1)'):in_cover=False
    if s.startswith('|') and '|' in s[1:]:
        tb=[]
        while i<len(lines) and lines[i].strip().startswith('|'):
            tb.append([c.strip() for c in lines[i].strip().strip('|').split('|')]); i+=1
        emit_table(tb); continue
    if s in ('','---'):i+=1;continue
    if s.startswith('<div') or s.startswith('</div'):i+=1;continue
    al=WD_ALIGN_PARAGRAPH.CENTER if in_cover else None
    if s.startswith('#### ') or s.startswith('### '):
        para(s.lstrip('# ').strip(),bold=True,size=16,color=BLACK,after=3,align=al)
    elif s.startswith('## '):
        para(s[3:].strip(),bold=True,size=16,color=BLACK,after=5,align=al)
    elif s.startswith('# '):
        if in_cover: para(s[2:].strip(),bold=True,size=18,color=BLUE,after=6,align=WD_ALIGN_PARAGRAPH.CENTER)
        else: para(s[2:].strip(),bold=True,size=17,color=BLACK,after=6)
    elif s.startswith('> '):
        para(s[2:].strip(),size=14,color=GRAY,after=4,align=al)
    elif re.match(r'^[-*•]\s',s):
        p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.left_indent=Cm(1.0); p.paragraph_format.space_after=Pt(2)
        add_runs(p,re.sub(r'^[-*•]\s','',s),size=15)
    else:
        para(s,after=5,align=al)
    i+=1
doc.save('59_COE_Creative_Economy_fullproposal.docx')
print('saved docx')
