# -*- coding: utf-8 -*-
"""master md -> styled HTML (+ SVG figures) -> for Edge PDF"""
import io, re, html as H
md = io.open('59_coe-creative-fullproposal.md', encoding='utf-8').read()

IG='#F26B21'; GR='#2E8B57'; GOV='#1F4E79'; EQ='#6A4C93'

FIG1 = '''<figure><svg viewBox="0 0 740 250" xmlns="http://www.w3.org/2000/svg">
<defs><marker id="ar" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 z" fill="#64748B"/></marker></defs>
<rect x="20" y="14" width="700" height="30" rx="6" fill="#1F4E79"/>
<text x="370" y="34" text-anchor="middle" fill="#fff" font-size="15" font-weight="700">ศูนย์ความเป็นเลิศฯ (COE) — กลไกขับเคลื่อน + leverage งบเดิม + จ่ายเมื่อเกิดผล</text>
'''
_nodes=[('1 ข้อมูล/สิทธิ์','Data & Rights','80 ลบ.',EQ),('2 กำลังคน','Talent','100 ลบ.',EQ),
('3 การผลิต','Production','140 ลบ.',IG),('4 มาตรฐาน','Standards','100 ลบ.',GOV),('5 ตลาด/ส่งออก','Market','60 ลบ.',GR)]
x=28
for idx,(th,en,bud,col) in enumerate(_nodes):
    FIG1+=f'<rect x="{x}" y="78" width="118" height="78" rx="8" fill="{col}"/>'
    FIG1+=f'<text x="{x+59}" y="104" text-anchor="middle" fill="#fff" font-size="14" font-weight="700">{th}</text>'
    FIG1+=f'<text x="{x+59}" y="122" text-anchor="middle" fill="#fff" font-size="11">{en}</text>'
    FIG1+=f'<text x="{x+59}" y="142" text-anchor="middle" fill="#fff" font-size="12" font-weight="700">{bud}</text>'
    if idx<4:
        FIG1+=f'<line x1="{x+118}" y1="117" x2="{x+138}" y2="117" stroke="#64748B" stroke-width="2.5" marker-end="url(#ar)"/>'
    x+=138
FIG1+='<path d="M 660 165 C 660 205, 87 205, 87 165" fill="none" stroke="#94A3B8" stroke-width="2" stroke-dasharray="5,4" marker-end="url(#ar)"/>'
FIG1+='<text x="370" y="200" text-anchor="middle" fill="#64748B" font-size="12">ป้อนกลับ: ข้อมูลตลาด / ความต้องการ / รายได้ (closed-loop)</text>'
FIG1+='<text x="370" y="232" text-anchor="middle" fill="#334155" font-size="12" font-weight="700">ภาพที่ 1 — วงจรระบบนิเวศปัญญาประดิษฐ์เพื่อเศรษฐกิจสร้างสรรค์ครบวงจร (Creative AI Content Lifecycle)</text>'
FIG1+='</svg></figure>'

# budget by year (62/82/92/90/80/74) + by program (80/100/140/100/60)
yrs=[('2570',62),('2571',82),('2572',92),('2573',90),('2574',80),('2575',74)]
FIGBUD='<figure><div class="figrow">'
sv='<svg viewBox="0 0 360 260" xmlns="http://www.w3.org/2000/svg">'
sv+='<line x1="40" y1="210" x2="345" y2="210" stroke="#334155"/>'
bx=52
for yr,v in yrs:
    h=v*1.9; y=210-h
    sv+=f'<rect x="{bx}" y="{y:.0f}" width="34" height="{h:.0f}" fill="#1F4E79"/>'
    sv+=f'<text x="{bx+17}" y="{y-4:.0f}" text-anchor="middle" font-size="12" font-weight="700" fill="#1F4E79">{v}</text>'
    sv+=f'<text x="{bx+17}" y="226" text-anchor="middle" font-size="11" fill="#334155">{yr}</text>'
    bx+=50
sv+='<text x="192" y="248" text-anchor="middle" font-size="11.5" font-weight="700" fill="#334155">ภาพที่ 4 — งบศูนย์ฯ รายปี (รวม 480 ลบ.)</text></svg>'
FIGBUD+=sv
progs=[('3 การผลิต',140,IG),('2 กำลังคน',100,EQ),('4 มาตรฐาน',100,GOV),('1 ข้อมูล/สิทธิ์',80,EQ),('5 ตลาด/ส่งออก',60,GR)]
sv2='<svg viewBox="0 0 360 260" xmlns="http://www.w3.org/2000/svg">'
yy=24
for name,v,col in progs:
    w=v*1.4
    sv2+=f'<text x="8" y="{yy+15}" font-size="11.5" fill="#334155">{name}</text>'
    sv2+=f'<rect x="120" y="{yy}" width="{w:.0f}" height="24" fill="{col}"/>'
    sv2+=f'<text x="{120+w+5:.0f}" y="{yy+17}" font-size="12" font-weight="700" fill="#334155">{v}</text>'
    yy+=42
sv2+='<text x="180" y="248" text-anchor="middle" font-size="11.5" font-weight="700" fill="#334155">ภาพที่ 5 — งบศูนย์ฯ ตามโปรแกรม (ลบ.)</text></svg>'
FIGBUD+=sv2+'</div></figure>'

FIGLEV='<figure><svg viewBox="0 0 740 300" xmlns="http://www.w3.org/2000/svg">'
FIGLEV+='<defs><marker id="ar2" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 z" fill="#64748B"/></marker></defs>'
FIGLEV+='<rect x="20" y="110" width="150" height="80" rx="10" fill="#1F4E79"/>'
FIGLEV+='<text x="95" y="142" text-anchor="middle" fill="#fff" font-size="15" font-weight="700">งบ COE</text>'
FIGLEV+='<text x="95" y="165" text-anchor="middle" fill="#fff" font-size="14" font-weight="700">480 ลบ.</text>'
mech=['co-pay','matching','rebate','first-loss','pay-for-outcome']
my=40
for mc in mech:
    FIGLEV+=f'<text x="250" y="{my+118}" text-anchor="middle" font-size="11" fill="#475569">{mc}</text>'
    my+=22
FIGLEV+='<line x1="172" y1="150" x2="330" y2="150" stroke="#64748B" stroke-width="2.5" marker-end="url(#ar2)"/>'
srcs=['สสว. BDS voucher','depa transformation voucher','BOI rebate / สิทธิประโยชน์','NIA grants','EXIM / SME D Bank สินเชื่อ','กองทุนพัฒนาดิจิทัลฯ','กองทุน ววน. (บพข.)','กองทุนพัฒนาสื่อฯ','งบส่งเสริมการค้า DITP']
sy=18
for sc in srcs:
    FIGLEV+=f'<rect x="360" y="{sy}" width="360" height="24" rx="4" fill="#ECFDF5" stroke="#A7F3D0"/>'
    FIGLEV+=f'<text x="372" y="{sy+16}" font-size="12" fill="#065F46">{sc}</text>'
    sy+=29
FIGLEV+='<text x="370" y="292" text-anchor="middle" font-size="11.5" font-weight="700" fill="#334155">ภาพที่ 6 — กลไกทวีคูณ: งบศูนย์ฯ ปลดล็อกงบเดิมของหน่วยงาน/กองทุน/เอกชนอีกหลายเท่า</text>'
FIGLEV+='</svg></figure>'

# inject placeholders
md = md.replace('## (1)', '[[FIG1]]\n\n## (1)', 1)
md = md.replace('### (ข) แหล่งงบ', '[[FIGBUD]]\n\n### (ข) แหล่งงบ', 1)
md = md.replace('### (ค) หลักการงบประมาณ', '[[FIGLEV]]\n\n### (ค) หลักการงบประมาณ', 1)

# ---- md -> html ----
def inline(t):
    t=H.escape(t, quote=False)
    t=re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    return t
lines=md.split('\n'); out=[]; i=0
while i<len(lines):
    s=lines[i].strip()
    if s.startswith('|'):
        rows=[]
        while i<len(lines) and lines[i].strip().startswith('|'):
            rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')]); i+=1
        rows=[r for r in rows if not re.match(r'^[\s:|-]+$','|'.join(r))]
        if rows:
            out.append('<table>')
            for ri,r in enumerate(rows):
                tag='th' if ri==0 else 'td'
                out.append('<tr>'+''.join(f'<{tag}>{inline(c)}</{tag}>' for c in r)+'</tr>')
            out.append('</table>')
        continue
    if s in ('','---'): out.append('<hr>' if s=='---' else ''); i+=1; continue
    if s.startswith('<div') or s.startswith('</div'): i+=1; continue
    if s.startswith('[[') and s.endswith(']]'): out.append({'[[FIG1]]':FIG1,'[[FIGBUD]]':FIGBUD,'[[FIGLEV]]':FIGLEV}.get(s,'')); i+=1; continue
    if s.startswith('#### ') or s.startswith('### '): out.append(f'<h3>{inline(s.lstrip("# "))}</h3>')
    elif s.startswith('## '): out.append(f'<h2>{inline(s[3:])}</h2>')
    elif s.startswith('# '): out.append(f'<h1>{inline(s[2:])}</h1>')
    elif s.startswith('> '): out.append(f'<blockquote>{inline(s[2:])}</blockquote>')
    elif re.match(r'^[-*•]\s', s):
        items=[]
        while i<len(lines) and re.match(r'^[-*•]\s', lines[i].strip()):
            items.append('<li>'+inline(re.sub(r'^[-*•]\s','',lines[i].strip()))+'</li>'); i+=1
        out.append('<ul>'+''.join(items)+'</ul>'); continue
    else: out.append(f'<p>{inline(s)}</p>')
    i+=1
htmlbody='\n'.join(out)

TPL='''<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
<style>
@page{size:A4;margin:1.4cm 1.5cm;}
*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
body{font-family:'Sarabun','TH Sarabun New','Leelawadee UI','Tahoma',sans-serif;color:#1F2937;font-size:11px;line-height:1.5;margin:0;}
h1{color:#1F4E78;font-size:1.5em;text-align:center;margin:14px 0 6px;}
h2{color:#1F4E78;font-size:1.18em;margin:14px 0 5px;border-bottom:2px solid #2E75B6;padding-bottom:2px;}
h3{color:#2E5A88;font-size:1.04em;margin:9px 0 3px;}
p{margin:4px 0;text-align:justify;}
ul{margin:4px 0 6px 0;padding-left:20px;} li{margin:2px 0;}
blockquote{margin:5px 0;padding:5px 10px;background:#F1F5F9;border-left:3px solid #94A3B8;color:#475569;font-size:0.95em;}
table{width:100%;border-collapse:collapse;font-size:0.92em;margin:6px 0;}
th,td{border:1px solid #CBD5E1;padding:4px 7px;text-align:left;vertical-align:top;}
th{background:#1F4E78;color:#fff;}
tr:nth-child(even) td{background:#F6F9FC;}
hr{border:none;border-top:1px solid #E2E8F0;margin:8px 0;}
figure{margin:12px 0;text-align:center;page-break-inside:avoid;}
figure svg{max-width:100%;height:auto;border:1px solid #E2E8F0;border-radius:6px;padding:6px;background:#fff;}
.figrow{display:flex;gap:10px;justify-content:center;flex-wrap:wrap;}
.figrow svg{flex:1;min-width:300px;}
strong{color:#0F172A;}
</style></head><body>
'''+htmlbody+'\n</body></html>'
io.open('59_coe-creative-fullproposal.html','w',encoding='utf-8').write(TPL)
print('html len', len(TPL))
