import re,json,os
S='/tmp/claude-0/-home-user-TEST01/bedc6b37-c512-5e8e-b8d3-45462d502183/scratchpad/work'
MD=open('/root/.claude/uploads/bedc6b37-c512-5e8e-b8d3-45462d502183/c2943e88-22_AOC_ICMemo_PPT___2026-11-02.md',encoding='utf-8').read()
MEMO=open(S+'/memo.md',encoding='utf-8').read()

# ---- MD architecture ----
units={}
order=[]
sections=[]
cur_sec=None
parts=re.split(r'\n(?=## |### )',MD)
for p in parts:
    if p.startswith('## ') and not p.startswith('### '):
        t=p.split('\n')[0][3:].strip()
        cur_sec=t
    if p.startswith('### '):
        head=p.split('\n')[0][4:].strip()
        uid,title=head.split('　',1)
        body=p
        def grab(tag):
            m=re.search(r'\*\*'+tag+r'\*\*(.*?)(?=\n\*\*|\Z)',body,re.S)
            return m.group(1).strip() if m else ''
        claim=grab('本页主张')
        claim=re.sub(r'^[　\s]*','',claim)
        claim=re.sub(r'^对 NewCo 的含义：','',claim)
        layout=grab('版式')
        figs=grab('图')
        tables=[l.strip('- ').strip() for l in grab('表').split('\n') if l.strip().startswith('-')]
        lim_blk=grab(r'必须同页保留的限制[^*]*')
        limits=[l.strip('- ').strip() for l in lim_blk.split('\n') if l.strip().startswith('-')]
        src=grab('来源脚注')
        pend=grab('本页待定')
        units[uid]=dict(id=uid,sec=cur_sec,title=title,claim=claim,layout=layout,figs=figs,
                        tables=tables,limits=limits,src=src,pend=pend)
        order.append(uid)
        if not sections or sections[-1][0]!=cur_sec: sections.append((cur_sec,uid))

# ---- memo per unit ----
# split memo by unit headings
idx=[]
for m in re.finditer(r'^\*\*(M\d+u?)　(.+?)\*\*$',MEMO,re.M):
    idx.append((m.start(),m.group(1)))
idx.append((re.search(r'^\*\*附录一　待定事项台账\*\*$',MEMO,re.M).start(),'END'))
for i,(st,uid) in enumerate(idx[:-1]):
    body=MEMO[st:idx[i+1][0]]
    if uid not in units: continue
    u=units[uid]
    # intro paragraph (first non-heading para)
    lines=[l for l in body.split('\n')[1:] if l.strip()]
    intro=next((l for l in lines if not l.startswith('**')),'')
    u['intro']=re.sub(r'\\','',intro).strip()
    # subheads under 一、是什么 / 二、怎么做
    subs=re.findall(r'^\*\*(\d)　(.+?)\*\*$',body,re.M)
    sec1=body.split('**二、怎么做**')[0]
    sec2=body.split('**二、怎么做**')[1] if '**二、怎么做**' in body else ''
    sec2=sec2.split('**三、关键数据**')[0]
    def heads(t):
        r=[re.sub(r'\\','',h).rstrip('。') for n,h in re.findall(r'^\*\*(\d)　(.+?)\*\*',t,re.M)]
        if not r: r=[re.sub(r'\\','',h) for n,h in re.findall(r'^\*\*([一二三四五六七八九])、(.+?)\*\*$',t,re.M)]
        return r
    if '**一、是什么**' in body:
        s1=body.split('**一、是什么**')[1].split('**二、怎么做**')[0]
        s2=body.split('**二、怎么做**')[1].split('**三、关键数据**')[0] if '**二、怎么做**' in body else ''
    else:
        pre=body.split('**三、关键数据**')[0]
        hs=re.split(r'\n(?=\*\*二、)',pre,maxsplit=1)
        s1=hs[0]; s2=hs[1] if len(hs)>1 else ''
    u['what']=heads(s1); u['how']=heads(s2)
    # images
    u['imgs']=[os.path.basename(x) for x in re.findall(r'<img src="([^"]+)"',body)]
    # markdown tables in 关键数据
    kd=body.split('**三、关键数据**')[-1]
    tabs=[]
    for tm in re.finditer(r'(\*\*表 [^\n*]+\*\*)\n\n((?:\|.*\n)+)',kd):
        cap=tm.group(1).strip('*')
        rows=[r for r in tm.group(2).strip().split('\n')]
        grid=[]
        for r in rows:
            if re.match(r'^\|[\s\-:|]+\|$',r): continue
            cells=[c.strip().strip('*').replace('\\','') for c in r.strip().strip('|').split('|')]
            grid.append(cells)
        tabs.append(dict(cap=cap,grid=grid))
    u['mtables']=tabs
    # 含义
    m=re.search(r'\*\*对 NewCo 的含义　?\*\*(.+)',body)
    u['meaning']=re.sub(r'^对 NewCo 的含义：','',m.group(1).strip()) if m else u['claim']

json.dump(dict(order=order,sections=sections,units=units),open(S+'/content.json','w'),ensure_ascii=False,indent=1)
print(len(order),'units')
for uid in order[:3]+order[-2:]:
    u=units[uid];print(uid,u['sec'],'|figs',u.get('imgs'),'|tabs',[t['cap'] for t in u.get('mtables',[])],'|what',len(u.get('what',[])),'how',len(u.get('how',[])))
