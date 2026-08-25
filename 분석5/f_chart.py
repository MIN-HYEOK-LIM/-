# -*- coding: utf-8 -*-
"""분석10_그래프 — 14종 (반도체 2종, 나머지는 비반도체·전산업)"""
SPAN10=16
w10=newsheet('분석10_그래프',[26]+[12]*20)
title(w10,1,'⑩ 그래프',SPAN10)
note(w10,2,1,'보조표는 분석3·4·5·6·8을 참조하는 수식입니다. 5·6번 산점도가 이 시트의 핵심으로, 사분면 위치가 곧 섹터의 국면입니다.',SPAN10)
r=4
sec(w10,r,'보조표 A. 섹터 좌표 (3M YoY · 가속도 · 단가Δ · 물량Δ · 규모 · 점수)',SPAN10); r+=1
A10H=r
QUAD=[('가속 성장',8),('바닥 통과',10),('성장 둔화',12),('침체 심화',14)]
DRIV=[('가격 주도',16),('물량 주도',18),('혼합',20)]
for i,h in enumerate(['섹터','3M YoY','가속도','단가Δ','물량Δ','2026 수출','종합점수']): head(w10,r,i+1,h)
for lab,c in QUAD: head(w10,r,c,lab+' X'); head(w10,r,c+1,lab+' Y')
for lab,c in DRIV: head(w10,r,c,lab+' X'); head(w10,r,c+1,lab+' Y')
r+=1; A10=r
S='분석4_섹터스코어보드'
for i,(sc,nm,g,ds) in enumerate(SEC_LIVE):
    sr=S4+i
    put(w10,r,1,f"='{S}'!B{sr}",sz=9)
    put(w10,r,2,f"=IFERROR('{S}'!M{sr},\"\")",PCT,sz=9)
    put(w10,r,3,f"=IFERROR('{S}'!O{sr}/100,\"\")",PCT,sz=9)
    put(w10,r,4,f"=IFERROR('{S}'!Q{sr},\"\")",PCT,sz=9)
    put(w10,r,5,f"=IFERROR('{S}'!R{sr},\"\")",PCT,sz=9)
    put(w10,r,6,f"='{S}'!F{sr}",NUM,sz=9)
    put(w10,r,7,f"=IFERROR('{S}'!X{sr},\"\")",'0.0',sz=9)
    for lab,c in QUAD:
        put(w10,r,c,  f'=IF(\'{S}\'!P{sr}="{lab}",B{r},"")',PCT,sz=9)
        put(w10,r,c+1,f'=IF(\'{S}\'!P{sr}="{lab}",C{r},"")',PCT,sz=9)
    for lab,c in DRIV:
        put(w10,r,c,  f'=IF(\'{S}\'!S{sr}="{lab}",D{r},"")',PCT,sz=9)
        put(w10,r,c+1,f'=IF(\'{S}\'!S{sr}="{lab}",E{r},"")',PCT,sz=9)
    r+=1
E10A=r-1; r+=1
sec(w10,r,'보조표 B. 그룹별 2026년 1~7월 수출 (천달러)',SPAN10); r+=1
B10H=r; head(w10,r,1,'그룹'); head(w10,r,2,'수출')
r+=1; B10=r
GRPS=[]
for sc,nm,g,ds in SEC_LIVE:
    if g not in GRPS: GRPS.append(g)
for g in GRPS:
    idxs=[j for j,(s2,n2,g2,d2) in enumerate(SEC_LIVE) if g2==g]
    put(w10,r,1,g,sz=9)
    put(w10,r,2,'='+'+'.join(f"'{S}'!F{S4+j}" for j in idxs),NUM,sz=9)
    r+=1
E10B=r-1; r+=1
sec(w10,r,'보조표 C. 비반도체 성장 상위 12 섹터 (3M YoY)',SPAN10); r+=1
C10H=r; head(w10,r,1,'섹터'); head(w10,r,2,'3M YoY')
r+=1; C10=r
import data4 as _d4
def _sz(sc):
    return sum(H4[c][m][1] for c in SEC_CODES[sc] for m in M17['2026'] if c in H4 and m in H4[c])
def _g3(sc):
    a=sum(H4[c][MONTHS[k]][1] for c in SEC_CODES[sc] for k in KL3 if c in H4 and MONTHS[k] in H4[c])
    b=sum(H4[c][MONTHS[k]][1] for c in SEC_CODES[sc] for k in KL3P if c in H4 and MONTHS[k] in H4[c])
    return (a/b-1) if b else -9
_ns=[(sc,i) for i,(sc,nm,g,ds) in enumerate(SEC_LIVE) if sc not in SEMI and _sz(sc)>200000]
_ns.sort(key=lambda t:-_g3(t[0]))
for sc,i in _ns[:12]:
    sr=S4+i
    put(w10,r,1,f"='{S}'!B{sr}",sz=9); put(w10,r,2,f"=IFERROR('{S}'!M{sr},0)",PCT,sz=9)
    r+=1
E10C=r-1
CT=E10C+3
def add10(ch,anchor,t,w=27,h=12,ylab=None,xlab=None,legend=True,off=0):
    flat(ch,off); ch.title=t; ch.width=w; ch.height=h; ch.style=2
    if ylab: ch.y_axis.title=ylab
    if xlab: ch.x_axis.title=xlab
    if not legend: ch.legend=None
    ch.y_axis.majorGridlines=ChartLines(); ch.dispBlanksAs='gap'
    w10.add_chart(ch,anchor)
def scat(anchor,groups,xt,yt,ttl,colmap):
    s_=ScatterChart(); s_.x_axis.title=xt; s_.y_axis.title=yt; s_.style=13
    for lab,c in groups:
        se=Series(Reference(w10,min_col=c+1,min_row=A10,max_row=E10A),
                  Reference(w10,min_col=c,  min_row=A10,max_row=E10A),title=lab)
        se.marker.symbol='circle'; se.marker.size=9
        se.marker.graphicalProperties.solidFill=colmap[lab]
        se.marker.graphicalProperties.line.solidFill='1F3864'
        se.graphicalProperties.line.noFill=True
        s_.series.append(se)
    s_.title=ttl; s_.width=28; s_.height=15; s_.y_axis.majorGridlines=ChartLines()
    w10.add_chart(s_,anchor)
W2='분석2_월별거시'; W3='분석3_수출구조분해'; W5='분석5_섹터월별'; W6='분석6_산업체인'; W8='분석8_마진·단가'
w2s=wb[W2]; w3s=wb[W3]; w5s=wb[W5]; w6s=wb[W6]; w8s=wb[W8]
A=0
# 1. 국가 수출 월별 + 12M 이동합
c1=LineChart()
c1.add_data(Reference(w2s,min_col=2,min_row=A2H,max_row=A2H+NM),titles_from_data=True)
c1.set_categories(Reference(w2s,min_col=1,min_row=A2,max_row=A2+NM-1))
c2_=LineChart()
c2_.add_data(Reference(w2s,min_col=10,min_row=A2H,max_row=A2H+NM),titles_from_data=True)
c2_.y_axis.axId=200; c2_.y_axis.title='12M 이동합'; c2_.y_axis.crosses='max'
flat(c2_,1); c1+=c2_
add10(c1,f'A{CT+A}','1. 국가 월별 수출과 12개월 이동합 — 추세선으로 본 사이클',ylab='천달러'); A+=24
# 2. 반도체 vs 비반도체 월별 YoY  ★
c=LineChart()
rowdata(c,w3s,[Y_TOT,Y_SEMI,Y_NS,Y_NSD],3+12,2+NM)
c.set_categories(Reference(w3s,min_col=3+12,max_col=2+NM,min_row=A3H))
add10(c,f'A{CT+A}','2. 반도체 vs 비반도체 월별 전년동월비 — 반도체를 빼면 무엇이 남는가',ylab='전년동월비'); A+=24
# 3. 반도체 비중 추이
c=LineChart()
rowdata(c,w3s,[R_SH],3,2+NM)
c.set_categories(Reference(w3s,min_col=3,max_col=2+NM,min_row=A3H))
add10(c,f'A{CT+A}','3. 국가 수출에서 반도체가 차지하는 몫 — 21.8%에서 45.9%로',ylab='비중',legend=False,off=4); A+=24
# 4. 연도 오버레이 (총수출 / 비반도체)
c=LineChart()
c.add_data(Reference(w2s,min_col=1,max_col=13,min_row=OVR['총수출'],max_row=OVR['총수출']+2),from_rows=True,titles_from_data=True)
c.add_data(Reference(w2s,min_col=1,max_col=13,min_row=OVR['비반도체'],max_row=OVR['비반도체']+2),from_rows=True,titles_from_data=True)
c.set_categories(Reference(w2s,min_col=2,max_col=13,min_row=OVR['총수출']-1))
add10(c,f'A{CT+A}','4. 1~12월 축에 겹쳐 본 3개 연도 — 총수출과 비반도체',ylab='천달러'); A+=24
# 5. 국면 지도
scat(f'A{CT+A}',QUAD,'최근 3개월 전년동월비','가속도 (3M − 직전 3M)',
     '5. 섹터 국면 지도 — 오른쪽 위 = 가속 성장, 왼쪽 위 = 바닥 통과',
     {'가속 성장':'2E75B6','바닥 통과':'70AD47','성장 둔화':'ED7D31','침체 심화':'C00000'}); A+=32
# 6. 성장의 질
scat(f'A{CT+A}',DRIV,'단가 증감률','물량 증감률',
     '6. 성장의 질 — 오른쪽 아래 = 가격 주도(사이클), 왼쪽 위 = 물량 주도(실수요)',
     {'가격 주도':'C00000','물량 주도':'548235','혼합':'BF9000'}); A+=32
print('분석10 전반부',CT,A)

# 7. 비반도체 성장 상위 12 (bar)
b=BarChart(); b.type='bar'
b.add_data(Reference(w10,min_col=2,min_row=C10H,max_row=E10C),titles_from_data=True)
b.set_categories(Reference(w10,min_col=1,min_row=C10,max_row=E10C))
add10(b,f'A{CT+A}','7. 비반도체 성장 상위 12 섹터 — 최근 3개월 전년동월비',h=13,ylab='3M YoY',legend=False,off=2); A+=26
# 8. 화학 체인 단계별 월별 YoY
ch_f,ch_l,ch_g=CH_ROWS['화학 — 원유에서 플라스틱까지']
c=LineChart(); rowdata(c,w6s,range(ch_f+1,ch_l+1),9+NM,8+NM+len(KY))
c.set_categories(Reference(w6s,min_col=9+NM,max_col=8+NM+len(KY),min_row=ch_f-1))
add10(c,f'A{CT+A}','8. 화학 체인 — 정유 · 기초유분 · 합성수지 · 가공품의 월별 전년동월비',ylab='전년동월비',off=1); A+=24
# 9. 철강 체인
st_f,st_l,st_g=CH_ROWS['철강 — 스크랩에서 강판까지']
c=LineChart(); rowdata(c,w6s,range(st_f+1,st_l+1),9+NM,8+NM+len(KY))
c.set_categories(Reference(w6s,min_col=9+NM,max_col=8+NM+len(KY),min_row=st_f-1))
add10(c,f'A{CT+A}','9. 철강 체인 — 상류(반제품·봉형강)와 하류(판재)가 갈라진다',ylab='전년동월비',off=3); A+=24
# 10. 자동차 체인
au_f,au_l,au_g=CH_ROWS['자동차 — 완성차·부품·전동화']
c=LineChart(); rowdata(c,w6s,range(au_f,au_l+1),9+NM,8+NM+len(KY))
c.set_categories(Reference(w6s,min_col=9+NM,max_col=8+NM+len(KY),min_row=au_f-1))
add10(c,f'A{CT+A}','10. 자동차 체인 — 완성차 · 부품 · 엔진 · 축전지의 월별 전년동월비',ylab='전년동월비',off=5); A+=24
# 11. 전동화 지표 + 특수 비율
c=LineChart(); rowdata(c,w6s,[SP6,SP6+2,SP6+3],14,13+NM)
c.set_categories(Reference(w6s,min_col=14,max_col=13+NM,min_row=SP6-1))
add10(c,f'A{CT+A}','11. 구조 전환 지표 — 전동화(축전지/엔진) · 판재/봉형강 · 정유/석화',ylab='배수',off=6); A+=24
# 12. 조선 3개월 이동합
c=LineChart(); rowdata(c,w6s,[SH6],9,8+NM)
c.set_categories(Reference(w6s,min_col=9,max_col=8+NM,min_row=SH6-1))
add10(c,f'A{CT+A}','12. 조선 상선 3개월 이동합 — 월별 편차를 걷어낸 추세',ylab='천달러',legend=False,off=8); A+=24
# 13. 마진 프록시
c=LineChart(); rowdata(c,w8s,[A8,A8+1,A8+2,A8+4],14,13+NM)
c.set_categories(Reference(w8s,min_col=14,max_col=13+NM,min_row=A8H))
add10(c,f'A{CT+A}','13. 마진 프록시 — 정제 · 석화 · 롤마진 · 이차전지 (배수 스케일이 비슷한 4종)',ylab='배수',off=9); A+=24
# 14. 반도체 체인 (밸류체인 단계별)
se_f,se_l,se_g=CH_ROWS['반도체 — 소자에서 장비까지']
c=LineChart(); rowdata(c,w6s,range(se_f,se_l+1),9+NM,8+NM+len(KY))
c.set_categories(Reference(w6s,min_col=9+NM,max_col=8+NM+len(KY),min_row=se_f-1))
add10(c,f'A{CT+A}','14. 반도체 체인 — 소자는 폭발했는데 장비·검사는 아직',ylab='전년동월비'); A+=24
# 15. 그룹별 규모
b=BarChart(); b.type='bar'
b.add_data(Reference(w10,min_col=2,min_row=B10H,max_row=E10B),titles_from_data=True)
b.set_categories(Reference(w10,min_col=1,min_row=B10,max_row=E10B))
add10(b,f'A{CT+A}','15. 산업 그룹별 2026년 1~7월 수출 규모',h=11,ylab='천달러',legend=False,off=7); A+=24
print('분석10 완료 차트',len(w10._charts))
