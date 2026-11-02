# -*- coding: utf-8 -*-
import json,os,re,copy
from pptx import Presentation
from pptx.util import Inches as In, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from PIL import Image

S=os.path.dirname(os.path.abspath(__file__))
D=json.load(open(S+'/content.json',encoding='utf-8'))
U=D['units']; ORDER=D['order']
MEMO=open(S+'/memo.md',encoding='utf-8').read()

BLUE=RGBColor(0x00,0x70,0xC0); NAVY=RGBColor(0x30,0x54,0x96); DARK=RGBColor(0x1F,0x30,0x48)
GREY=RGBColor(0x59,0x59,0x59); LGREY=RGBColor(0x80,0x80,0x80); BLACK=RGBColor(0x26,0x26,0x26)
BAND=RGBColor(0xDE,0xEB,0xF7); SOFT=RGBColor(0xF2,0xF5,0xF9); ORANGE=RGBColor(0xC5,0x5A,0x11)
WHITE=RGBColor(0xFF,0xFF,0xFF); LINE=RGBColor(0xBF,0xCD,0xDD)
EA='微软雅黑'; LAT='Arial'

def setf(r,sz=10,b=False,color=BLACK,italic=False,ea=EA,lat=LAT):
    f=r.font; f.size=Pt(sz); f.bold=b; f.italic=italic; f.name=lat
    f.color.rgb=color
    rPr=r._r.get_or_add_rPr()
    for tag in ('a:ea','a:cs'):
        e=rPr.find(qn(tag))
        if e is None:
            e=rPr.makeelement(qn(tag),{}); rPr.append(e)
        e.set('typeface',ea if tag=='a:ea' else lat)

def tb(sl,x,y,w,h,wrap=True,anchor=MSO_ANCHOR.TOP):
    s=sl.shapes.add_textbox(In(x),In(y),In(w),In(h)); tf=s.text_frame
    tf.word_wrap=wrap; tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    tf.vertical_anchor=anchor
    return s,tf

def para(tf,text,sz=10,b=False,color=BLACK,space_after=3,first=False,align=PP_ALIGN.LEFT,line=None,italic=False):
    p=tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment=align
    p.space_after=Pt(space_after)
    if line: p.line_spacing=line
    r=p.add_run(); r.text=text; setf(r,sz,b,color,italic)
    return p

def rect(sl,x,y,w,h,fill=None,lineclr=None,lw=0.75):
    from pptx.enum.shapes import MSO_SHAPE
    s=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,In(x),In(y),In(w),In(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb=fill
    else: s.fill.background()
    if lineclr: s.line.color.rgb=lineclr; s.line.width=Pt(lw)
    else: s.line.fill.background()
    s.shadow.inherit=False
    s.text_frame.word_wrap=True
    for m in ('margin_left','margin_right','margin_top','margin_bottom'): setattr(s.text_frame,m,0)
    s.text_frame.paragraphs[0].alignment=PP_ALIGN.LEFT
    return s

def clean(t):
    t=t.replace('\\','').replace('**','')
    t=re.sub(r'\s+',' ',t).strip()
    return t

def cut(t,n):
    t=clean(t)
    return t if len(t)<=n else t[:n-1]+'…'

# ---------- figure captions ----------
CAPS={}
for m in re.finditer(r'^(图 (M\d+u?)-\d+[^\n]*)',MEMO,re.M):
    CAPS.setdefault(m.group(2),clean(m.group(1)))

prs=Presentation(S+'/tpl.pptx')
# drop existing slides
sldIdLst=prs.slides._sldIdLst
for sid in list(sldIdLst):
    rId=sid.get(qn('r:id')); prs.part.drop_rel(rId); sldIdLst.remove(sid)

L_COVER=prs.slide_layouts[11]; L_MAIN=prs.slide_layouts[12]

W=13.333; ML=0.56; CW=12.09

def add_slidenum(sl,layout):
    for sp in layout.shapes:
        el=sp._element
        ph=el.find('.//'+qn('p:ph'))
        if ph is not None and ph.get('type')=='sldNum':
            new=copy.deepcopy(el)
            sl.shapes._spTree.append(new)
            return

def newslide(title=None,tag=None,layout=L_MAIN,title_sz=18):
    sl=prs.slides.add_slide(layout)
    add_slidenum(sl,layout)
    # remove inherited empty placeholders we don't fill, keep slide-number
    for ph in list(sl.placeholders):
        if ph.placeholder_format.idx==13 and title is None:
            ph._element.getparent().remove(ph._element)
    if title is not None:
        ph=None
        for p in sl.placeholders:
            if p.placeholder_format.idx==13: ph=p
        ph.left=In(ML); ph.top=In(0.60); ph.width=In(CW); ph.height=In(0.70)
        tf=ph.text_frame; tf.word_wrap=True
        tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
        p=tf.paragraphs[0]
        for r in list(p.runs): r._r.getparent().remove(r._r)
        r=p.add_run(); r.text=title; setf(r,title_sz,True,BLUE)
        p.line_spacing=1.0
    if tag:
        s,t=tb(sl,7.0,1.33,5.65,0.26)
        para(t,'/  '+tag,9,True,LGREY,0,True,PP_ALIGN.RIGHT)
    return sl

def claimband(sl,text,y=1.33,h=0.66):
    bx=rect(sl,ML,y,CW,h,SOFT,LINE,0.75)
    tfr=bx.text_frame; tfr.margin_left=In(0.12); tfr.margin_right=In(0.12); tfr.margin_top=In(0.05)
    tfr.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tfr.paragraphs[0]; p.line_spacing=1.08; p.alignment=PP_ALIGN.LEFT
    r=p.add_run(); r.text='本页主张  '; setf(r,10,True,ORANGE)
    r=p.add_run(); r.text=clean(text); setf(r,10,False,BLACK)
    return bx

def footnotes(sl,limits,src,pend=''):
    s,t=tb(sl,ML,6.44,CW,0.46)
    first=True
    for L in limits[:2]:
        para(t,'· '+cut(L,112),7.5,False,GREY,1,first,line=1.0); first=False
    s2,t2=tb(sl,2.10,6.93,10.2,0.30)
    txt='来源：'+cut(src,150)
    if pend: txt+='　｜　本页待定 '+clean(pend).split('；')[0]
    para(t2,txt,7,False,LGREY,0,True,line=1.0)

def points(sl,u,x,y,w,h,n=4,sz=9.5):
    s,t=tb(sl,x,y,w,h)
    para(t,'本页要点',10,True,NAVY,5,True)
    pts=(u['what']+u['how'])[:n]
    for i,p in enumerate(pts):
        pp=t.add_paragraph(); pp.space_after=Pt(6); pp.line_spacing=1.12
        r=pp.add_run(); r.text='%d　'%(i+1); setf(r,sz,True,ORANGE)
        r=pp.add_run(); r.text=cut(p,54); setf(r,sz,False,BLACK)
    return s

def tablelist(sl,u,x,y,w,h,skip=0):
    caps=[clean(c) for c in u['tables']][skip:skip+4]
    if not caps: return
    s,t=tb(sl,x,y,w,h)
    para(t,'本页表（全表见备忘录）',8.5,True,NAVY,4,True)
    for c in caps:
        para(t,'· '+cut(c,52),7.5,False,GREY,3,line=1.05)

def addtable(sl,grid,x,y,w,maxrows=5,maxcols=5,sz=8,hsz=8):
    grid=[row[:maxcols] for row in grid]
    ncol=max(len(r) for r in grid)
    grid=[r+['']*(ncol-len(r)) for r in grid]
    body=grid[1:1+maxrows]
    rows=1+len(body)
    h=0.30+0.36*len(body)
    gf=sl.shapes.add_table(rows,ncol,In(x),In(y),In(w),In(h))
    tb_=gf.table
    tb_.first_row=True; tb_.horz_banding=True
    # column widths: first col wider
    base=w/ncol
    widths=[base]*ncol
    if ncol>=3:
        widths[0]=base*1.25
        extra=(base*0.25)/(ncol-1)
        for i in range(1,ncol): widths[i]=base-extra
    for i,ww in enumerate(widths): tb_.columns[i].width=In(ww)
    tb_.rows[0].height=In(0.30)
    for r in range(1,rows): tb_.rows[r].height=In(0.36)
    for c in range(ncol):
        cell=tb_.cell(0,c); cell.fill.solid(); cell.fill.fore_color.rgb=NAVY
        cell.margin_left=In(0.06);cell.margin_right=In(0.06);cell.margin_top=In(0.03);cell.margin_bottom=In(0.03)
        cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=cell.text_frame.paragraphs[0]; r_=p.add_run(); r_.text=cut(grid[0][c],int(widths[c]*9)); setf(r_,hsz,True,WHITE)
        p.line_spacing=1.0
    for ri,row in enumerate(body):
        for c in range(ncol):
            cell=tb_.cell(ri+1,c)
            cell.fill.solid(); cell.fill.fore_color.rgb=WHITE if ri%2==0 else SOFT
            cell.margin_left=In(0.06);cell.margin_right=In(0.06);cell.margin_top=In(0.02);cell.margin_bottom=In(0.02)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            p=cell.text_frame.paragraphs[0]; p.line_spacing=1.0
            budget=max(10,int((widths[c]-0.14)*8.4*2))
            r_=p.add_run(); r_.text=cut(row[c],budget)
            setf(r_,sz,c==0,BLACK if c else NAVY)
    return gf

def pickgrid(u,maxcols=5):
    best=None
    for t in u['mtables']:
        g=t['grid']
        if len(g)<2: continue
        if best is None or len(g[0])<=maxcols and len(g[0])>len(best[0]['grid'][0]) and len(best[0]['grid'][0])>maxcols:
            pass
        if best is None: best=(t,0)
    # prefer a table with 3..maxcols columns and >=3 rows
    cands=[(i,t) for i,t in enumerate(u['mtables']) if len(t['grid'])>=3 and 3<=len(t['grid'][0])<=maxcols]
    if cands: i,t=cands[0]; return t,i
    if u['mtables']: return u['mtables'][0],0
    return None,0

def addimage(sl,fn,x,y,w,h,caption=None):
    path=S+'/img/'+fn.replace('.png','.jpg')
    iw,ih=Image.open(path).size
    sc=min(w/iw,h/ih)
    dw,dh=iw*sc,ih*sc
    sl.shapes.add_picture(path,In(x+(w-dw)/2),In(y+(h-dh)/2),In(dw),In(dh))
    return dw,dh

SECTAG={}
for sec,uid in D['sections']:
    SECTAG[sec]=sec.split('　')[1]

# ================= COVER =================
sl=prs.slides.add_slide(L_COVER)
s,t=tb(sl,1.22,2.30,11.0,0.9)
para(t,'AOC 新公司设立｜投资委员会讨论材料',30,True,DARK,0,True)
s,t=tb(sl,1.22,3.35,11.0,0.45)
para(t,'36 单元主文　＋　附录页组　｜　本轮请批：限额准备验证期',13,False,BLUE,0,True)
s,t=tb(sl,1.22,4.00,11.0,0.7)
para(t,'2026-11-02',12,True,GREY,4,True)
para(t,'Ruochen Zhang, M.D.　｜　Confidential material　｜　依据 IC Memo v2（36 单元）编制',10,False,LGREY,0)
logo=S+'/../tpl/ppt/media/image8.png'
if os.path.exists(logo): sl.shapes.add_picture(logo,In(1.22),In(1.10),height=In(0.62))

# ================= 执行摘要 =================
u1=U['M01']
sl=newslide('本轮请批：一段限额准备验证期——按六道关口分段拨付，三构型并行评估','执行摘要')
claimband(sl,u1['claim'])
cards=[('买三类读出','权利方向性意见（E0→G1）｜生产性递送判别（E1→G2／G3）｜偶联物可解释性（E4→G6，与 E1 并行）'),
       ('授权范围','止于 G2／G3，加全程并行的 G6；G4／G5 另行上会。拨付挂读出，不挂商务事件与日历'),
       ('三构型并行','A 肌肉核内修复（GAA IVS1＋DYSF）｜B 病理成纤维细胞 HSP47｜C 特定肾细胞入口（TAL／足细胞）'),
       ('金额口径','额度、估值、持股与审批结果一律「待定—主责方」；阈值在 TPP 定稿后反推，本轮不给通用数字')]
x=ML; cw=(CW-3*0.18)/4
for i,(h,b) in enumerate(cards):
    bx=rect(sl,x+i*(cw+0.18),2.10,cw,1.40,WHITE,LINE)
    tfr=bx.text_frame; tfr.margin_left=In(0.12);tfr.margin_right=In(0.12);tfr.margin_top=In(0.10)
    p=tfr.paragraphs[0]; r=p.add_run(); r.text=h; setf(r,11,True,BLUE); p.space_after=Pt(5)
    para(tfr,b,9,False,BLACK,0,line=1.15)
g,_=pickgrid(u1)
grid=u1['mtables'][2]['grid']
s,t=tb(sl,ML,3.64,CW,0.26)
para(t,'追加投资与六道关口的挂接（拨付挂读出；金额一律待定）',10,True,NAVY,0,True)
addtable(sl,grid,ML,3.94,CW,maxrows=6,maxcols=5,sz=8)
footnotes(sl,u1['limits'],u1['src'],u1['pend'])

# ================= 目录 =================
sl=newslide('本材料的十个部分：从投资命题到投资与治理，36 个单元逐页给结论','目录')
claimband(sl,'每页一个结论式标题；页内给图、给表、给限制与来源；阈值与金额一律「待定—主责方」，在 TPP 定稿后反推。')
secs=D['sections']
x=ML; cw=(CW-0.22)/2
for i,(sec,first) in enumerate(secs):
    col=i//5; row=i%5
    bx=rect(sl,ML+col*(cw+0.22),2.20+row*0.84,cw,0.74,WHITE,LINE)
    tfr=bx.text_frame; tfr.margin_left=In(0.14);tfr.margin_right=In(0.12);tfr.margin_top=In(0.07)
    num,name=sec.split('　')
    p=tfr.paragraphs[0]; p.space_after=Pt(3)
    r=p.add_run(); r.text=num+'　'; setf(r,12,True,ORANGE)
    r=p.add_run(); r.text=name; setf(r,12,True,NAVY)
    ids=[k for k in ORDER if U[k]['sec']==sec]
    para(tfr,'　'.join(ids)+'　（%d 页）'%len(ids),8.5,False,GREY,0)
footnotes(sl,['页码以本文件为准；附录页组列于主文之后（待定事项台账、术语表、原图档案 26 幅）。'],
          '架构依据 IC Memo v2 的 36 个单元；版式参照 BaseCure 原版 16:9（仅版式，不借事实）。')

# ================= 正文 =================
for sec,firstuid in D['sections']:
    num,name=sec.split('　')
    ids=[k for k in ORDER if U[k]['sec']==sec]
    sl=newslide(None,None)
    s,t=tb(sl,ML,2.10,2.0,1.6)
    para(t,num,72,True,RGBColor(0xDC,0xE5,0xEF),0,True)
    s,t=tb(sl,2.30,2.46,9.9,0.9)
    para(t,'第 %s 部分'%num,12,True,ORANGE,6,True)
    para(t,name,32,True,NAVY,0)
    s,t=tb(sl,2.30,4.10,9.9,2.2)
    for k in ids:
        p=t.add_paragraph() if (k!=ids[0]) else t.paragraphs[0]
        p.space_after=Pt(8); p.line_spacing=1.05
        r=p.add_run(); r.text=k+'　'; setf(r,11,True,ORANGE)
        r=p.add_run(); r.text=U[k]['title']; setf(r,11,False,BLACK)

    for uid in ids:
        u=U[uid]
        sl=newslide(uid+'　'+u['title'],name)
        claimband(sl,u['claim'])
        imgs=u['imgs']
        grid,gi=pickgrid(u)
        if imgs:
            fn=imgs[0]
            iw,ih=Image.open(S+'/img/'+fn.replace('.png','.jpg')).size
            ar=iw/ih
            if ar>=2.9:   # very wide: 图上 表下
                addimage(sl,fn,ML,2.08,CW,2.00)
                yy=4.22
                if grid:
                    s,t=tb(sl,ML,yy,CW,0.24)
                    para(t,clean(grid['cap']),9,True,NAVY,0,True)
                    addtable(sl,grid['grid'],ML,yy+0.30,CW,maxrows=4,maxcols=6,sz=8)
                else:
                    points(sl,u,ML,yy,CW,1.9,n=4)
                s,t=tb(sl,ML,6.26,CW,0.18)
                para(t,CAPS.get(uid,'')[:70],7,False,LGREY,0,True)
            else:          # tall: 图左 要点右
                ih_=4.24
                iwid=min(ar*ih_,8.60)
                addimage(sl,fn,ML,2.08,iwid,ih_)
                s,t=tb(sl,ML,6.34,iwid,0.18)
                para(t,CAPS.get(uid,'')[:int(iwid*9)],7,False,LGREY,0,True)
                px=ML+iwid+0.34; pw=12.65-px
                if pw>=5.4 and grid:
                    points(sl,u,px,2.08,pw,2.00,n=4)
                    s,t=tb(sl,px,4.18,pw,0.22)
                    para(t,clean(grid['cap']),8.5,True,NAVY,0,True)
                    addtable(sl,grid['grid'],px,4.46,pw,maxrows=4,maxcols=4,sz=7.5,hsz=7.5)
                else:
                    points(sl,u,px,2.08,pw,2.70,n=4)
                    tablelist(sl,u,px,4.92,pw,1.44)
        else:              # 无图：要点条 + 大表
            pts=(u['what']+u['how'])[:4]
            cw2=(CW-3*0.16)/4
            for i,p in enumerate(pts):
                bx=rect(sl,ML+i*(cw2+0.16),2.10,cw2,0.74,WHITE,LINE)
                tfr=bx.text_frame; tfr.margin_left=In(0.10);tfr.margin_right=In(0.10);tfr.margin_top=In(0.06)
                pp=tfr.paragraphs[0]; pp.space_after=Pt(2)
                r=pp.add_run(); r.text='要点 %d'%(i+1); setf(r,8,True,ORANGE)
                para(tfr,cut(p,40),9,True,NAVY,0,line=1.1)
            if grid:
                s,t=tb(sl,ML,3.04,CW,0.24)
                para(t,clean(grid['cap']),9.5,True,NAVY,0,True)
                nr=min(6,len(grid['grid'])-1)
                addtable(sl,grid['grid'],ML,3.36,CW,maxrows=6,maxcols=6,sz=8)
                bot=3.36+0.30+0.36*nr
                if bot<5.80:
                    tablelist(sl,u,ML,bot+0.12,CW,6.36-bot,skip=gi+1)
        footnotes(sl,u['limits'],u['src'],u['pend'])

# ================= 附录 =================
sl=newslide('附录页组：待定事项台账、术语表与原图档案','附录')
claimband(sl,'附录按主责方归集待定事项、按中英对照列术语、并以带日期的对照存档保留 26 幅原图；原图保留不等于事实已更新。')
rowsA=[['附-A','待定事项台账（按主责方归集）','本轮全部「待定—主责方」逐条登记，讲解时逐条点名主责方'],
       ['附-B','术语表（中英对照）','全文术语与缩写，按首次出现单元标注'],
       ['附-C','原图档案 26 幅','F01–F26 带日期的对照存档，正文使用改绘版'],
       ['附-D','实验方案 E0–E4 逐项','目的／对照／读出／关口／资源'],
       ['附-E','94 条疾病与基因型记录','多页分组表'],
       ['附-F','资产与注册状态全表','核查日 2026-10-11，ClinicalTrials.gov 逐项核对']]
addtable(sl,[['组','内容','说明']]+rowsA,ML,2.20,CW,maxrows=6,maxcols=3,sz=9)
footnotes(sl,['空白格统一读作「未收录／本轮未核到」，不等于「无」；不设加权总分页。'],
          '附录内容依据 IC Memo v2 附录一／附录二／附录三。')

# 附-A 待定事项台账（前两页）
led=MEMO[re.search(r'^\*\*附录一　待定事项台账\*\*$',MEMO,re.M).start():]
owners=re.findall(r'^\*\*(.+?)　（(\d+) 条）\*\*$',led,re.M)
own=[(o,int(n)) for o,n in owners][:24]
for page in range(2):
    chunk=own[page*12:(page+1)*12]
    if not chunk: break
    sl=newslide('附-A　待定事项台账：按主责方归集（%d/2）'%(page+1),'附录')
    if page==0:
        claimband(sl,'待定事项按主责方归集，讲解时逐条点名主责方；本轮不以估算值补上任何缺失输入。')
    else:
        claimband(sl,'下列主责方各自承担的条数；条目明细见备忘录附录一。')
    g=[['主责方','条数','主责方','条数']]
    half=(len(chunk)+1)//2
    for i in range(half):
        a=chunk[i]; b=chunk[i+half] if i+half<len(chunk) else ('','')
        g.append([a[0],str(a[1]),b[0],str(b[1]) if b[1] else ''])
    addtable(sl,g,ML,2.20,CW,maxrows=6,maxcols=4,sz=9)
    footnotes(sl,['本页只给归集口径与条数，条目原文见备忘录附录一。'],'IC Memo v2 附录一　待定事项台账。')

# 附-C 原图档案
figs=re.findall(r'^\*\*(F\d\d)\*\*\n\n<img src="([^"]+)"',MEMO,re.M)
sl=newslide('附-C　原图档案 26 幅：带日期的对照存档','附录')
claimband(sl,'原图保留不等于事实已更新；需修订的原图，正文使用改绘版，此处保留原图作带日期的对照存档。')
s,t=tb(sl,ML,2.20,CW,0.3)
para(t,'来源：AOC 技术平台与疾病立项全景研究 汇总核验版　｜　以下 %d 页一图一页'%len(figs),9.5,False,GREY,0,True)
cols=7
for i,(fid,p) in enumerate(figs):
    bx=rect(sl,ML+(i%cols)*((CW-0.1*(cols-1))/cols+0.1),2.70+(i//cols)*0.52,(CW-0.1*(cols-1))/cols,0.42,SOFT,LINE)
    tfr=bx.text_frame; tfr.vertical_anchor=MSO_ANCHOR.MIDDLE
    pp=tfr.paragraphs[0]; pp.alignment=PP_ALIGN.CENTER
    r=pp.add_run(); r.text=fid; setf(r,11,True,NAVY)
footnotes(sl,['原图未作改绘，日期与口径以原报告为准。'],'IC Memo v2 附录三　原图档案（26 幅）。')
for fid,p in figs:
    fn=os.path.basename(p)
    sl=newslide('附-C　%s（原图存档）'%fid,'附录')
    addimage(sl,fn,ML,1.40,CW,5.0)
    s,t=tb(sl,ML,6.50,CW,0.2)
    para(t,'原图保留不等于事实已更新；正文如有改绘版以正文为准。',7.5,False,LGREY,0,True)

prs.save(S+'/AOC_IC.pptx')
print('slides',len(prs.slides.__iter__.__self__._sldIdLst))
