# -*- coding: utf-8 -*-
"""분석2_월별거시 — 총괄 추이 · MoM/QoQ/YoY · 물량가격 분해 · 연도 오버레이 원본"""
w2=newsheet('분석2_월별거시',[10,13,13,13,10,12,11,11,11,13,13,12,12,12,11,11])
SPAN2=16
title(w2,1,'② 월별 거시 — 31개월 총괄 추이와 증감의 성격',SPAN2)
note(w2,2,1,'모든 값은 원본 「수출입 총괄」을 참조합니다. 조업일수로 나눈 일평균은 달력 효과를 걷어낸 실질 흐름이고, '
            '12개월 이동합은 계절성을 걷어낸 추세선입니다. 단가는 금액÷중량으로, 성장이 물량에서 왔는지 가격에서 왔는지 가르는 기준입니다.',SPAN2)
r=4
sec(w2,r,'[A] 월별 총괄 (2024.01~2026.07)',SPAN2); r+=1
A2H=r
for i,h in enumerate(['기간','수출 금액','수입 금액','무역수지','조업일수','일평균 수출','수출 MoM','수출 YoY',
                      '일평균 YoY','수출 12M 이동합','12M 이동합 YoY','수출 중량','수출 단가','수입 단가','교역조건','반도체 비중']):
    head(w2,r,i+1,h)
w2.row_dimensions[r].height=30
r+=1; A2=r
SEMI_LIVE2=[c for c in SEMI_CODES if c in CROW]
for k in range(NM):
    rr=A2+k
    put(w2,rr,1,MONTHS[k],align='center',sz=9)
    put(w2,rr,2,'='+TX(k),NUM); put(w2,rr,3,'='+TI(k),NUM); put(w2,rr,4,f'=B{rr}-C{rr}',NUM)
    put(w2,rr,5,'='+TD(k),NUM1); put(w2,rr,6,f'=IFERROR(B{rr}/E{rr},"-")',NUM)
    put(w2,rr,7,(f'=IFERROR(B{rr}/B{rr-1}-1,"-")' if k>0 else None),PCT)
    put(w2,rr,8,(f'=IFERROR(B{rr}/B{rr-12}-1,"-")' if k>=12 else None),PCT,bold=True)
    put(w2,rr,9,(f'=IFERROR(F{rr}/F{rr-12}-1,"-")' if k>=12 else None),PCT)
    put(w2,rr,10,(f'=SUM(B{rr-11}:B{rr})' if k>=11 else None),NUM)
    put(w2,rr,11,(f'=IFERROR(J{rr}/J{rr-12}-1,"-")' if k>=23 else None),PCT)
    put(w2,rr,12,'='+TWX(k),NUM1)
    put(w2,rr,13,f'=IFERROR(B{rr}/L{rr},"-")',UNIT)
    put(w2,rr,14,f'=IFERROR(C{rr}/{DT}!J{trow(k)},"-")',UNIT)
    put(w2,rr,15,f'=IFERROR(M{rr}/N{rr},"-")',UNIT)
    put(w2,rr,16,f'=IFERROR(({csum(PEX,SEMI_LIVE2,k)})/B{rr},"-")',PCT)
r=A2+NM; E2A=r-1; r+=1

sec(w2,r,'[B] 분기별 — QoQ · YoY',SPAN2); r+=1
for i,h in enumerate(['분기','수출 금액','수입 금액','무역수지','조업일수','일평균 수출','QoQ','YoY','일평균 YoY','반도체','비반도체','반도체 YoY','비반도체 YoY']):
    head(w2,r,i+1,h)
r+=1; B2=r
QS=[]
for y in (2024,2025,2026):
    for q in (1,2,3,4):
        ms=[f'{y}.{mm:02d}' for mm in range(3*q-2,3*q+1) if f'{y}.{mm:02d}' in MONTHS]
        if ms: QS.append((f'{y} {q}Q',[MONTHS.index(m) for m in ms]))
for i,(lab,ks) in enumerate(QS):
    rr=B2+i
    ex='+'.join(f'B{A2+k}' for k in ks); im='+'.join(f'C{A2+k}' for k in ks); dy='+'.join(f'E{A2+k}' for k in ks)
    sm='+'.join(f'({csum(PEX,SEMI_LIVE2,k)})' for k in ks)
    put(w2,rr,1,lab,align='center',bold=True,sz=9)
    put(w2,rr,2,'='+ex,NUM); put(w2,rr,3,'='+im,NUM); put(w2,rr,4,f'=B{rr}-C{rr}',NUM)
    put(w2,rr,5,'='+dy,NUM1); put(w2,rr,6,f'=IFERROR(B{rr}/E{rr},"-")',NUM)
    put(w2,rr,7,(f'=IFERROR(B{rr}/B{rr-1}-1,"-")' if i>0 else None),PCT)
    put(w2,rr,8,(f'=IFERROR(B{rr}/B{rr-4}-1,"-")' if i>=4 else None),PCT,bold=True)
    put(w2,rr,9,(f'=IFERROR(F{rr}/F{rr-4}-1,"-")' if i>=4 else None),PCT)
    put(w2,rr,10,'='+sm,NUM); put(w2,rr,11,f'=B{rr}-J{rr}',NUM)
    put(w2,rr,12,(f'=IFERROR(J{rr}/J{rr-4}-1,"-")' if i>=4 else None),PCT)
    put(w2,rr,13,(f'=IFERROR(K{rr}/K{rr-4}-1,"-")' if i>=4 else None),PCT,bold=True)
r=B2+len(QS); E2B=r-1; r+=1

sec(w2,r,'[C] 연도 오버레이 원본 — 1~12월 축에 2024·2025·2026을 겹쳐 보기',SPAN2); r+=1
note(w2,r,1,'그래프(분석10)의 오버레이 차트가 이 표를 참조합니다. 2026년은 7월까지만 있어 8~12월은 빈칸으로 두었습니다(선이 0으로 떨어지지 않도록).',SPAN2); r+=1
OV=[('총수출',lambda k: f'B{A2+k}'),('반도체',lambda k: f'({csum(PEX,SEMI_LIVE2,k)})'),
    ('비반도체',lambda k: f'B{A2+k}-({csum(PEX,SEMI_LIVE2,k)})'),('일평균 수출',lambda k: f'F{A2+k}'),
    ('무역수지',lambda k: f'D{A2+k}')]
OVR={}
for lab,fn in OV:
    for i,h in enumerate(['지표/연도']+[f'{mm}월' for mm in range(1,13)]): head(w2,r,i+1,h)
    r+=1; base=r
    for y in (2024,2025,2026):
        put(w2,r,1,f'{lab} {y}',bold=True,sz=9)
        for mm in range(1,13):
            key=f'{y}.{mm:02d}'
            if key in MONTHS: put(w2,r,1+mm,'='+fn(MONTHS.index(key)),NUM,sz=9)
            else: put(w2,r,1+mm,None,NUM,sz=9)
        r+=1
    OVR[lab]=base
    r+=1
E2C=r-1
print('분석2 [A][B][C] 완료',A2,B2,OVR)

r=E2C+2
sec(w2,r,'[D] 월별 증감의 성격 — 물량효과 · 가격효과 · 교차효과',SPAN2); r+=1
note(w2,r,1,'금액 증감 = 물량효과(전년 단가 × 물량 증감) + 가격효과(전년 물량 × 단가 증감) + 교차효과(증감 × 증감). '
            '가격효과가 압도적이면 사이클성 호황이고, 물량효과가 크면 실수요가 받쳐 주는 성장입니다.',SPAN2); r+=1
for i,h in enumerate(['기간','수출 금액','전년 금액','증감액','물량효과','가격효과','교차효과','물량 기여율','가격 기여율','중량 YoY','단가 YoY','판정']):
    head(w2,r,i+1,h)
r+=1; D2=r
for k in KY:
    rr=D2+(k-12); a=A2+k; p=A2+k-12
    put(w2,rr,1,MONTHS[k],align='center',sz=9)
    put(w2,rr,2,f'=B{a}',NUM); put(w2,rr,3,f'=B{p}',NUM); put(w2,rr,4,f'=B{rr}-C{rr}',NUM)
    put(w2,rr,5,f'=IFERROR((L{a}-L{p})*M{p},"-")',NUM)
    put(w2,rr,6,f'=IFERROR((M{a}-M{p})*L{p},"-")',NUM)
    put(w2,rr,7,f'=IFERROR((L{a}-L{p})*(M{a}-M{p}),"-")',NUM)
    put(w2,rr,8,f'=IFERROR(E{rr}/D{rr},"-")',PCT)
    put(w2,rr,9,f'=IFERROR(F{rr}/D{rr},"-")',PCT)
    put(w2,rr,10,f'=IFERROR(L{a}/L{p}-1,"-")',PCT)
    put(w2,rr,11,f'=IFERROR(M{a}/M{p}-1,"-")',PCT)
    put(w2,rr,12,f'=IF(D{rr}=0,"-",IF(ABS(F{rr})>ABS(E{rr})*2,"가격이 주도",IF(ABS(E{rr})>ABS(F{rr})*2,"물량이 주도","혼합")))',align='center',sz=9)
r=D2+len(KY); E2D=r-1; r+=1
put(w2,r,1,'2026 1~7 누계',bold=True,align='center',sz=9,fill=SUB)
_s=lambda c: '+'.join(f'{c}{D2+(k-12)}' for k in K26)
for col,ltr in [(2,'B'),(3,'C'),(4,'D'),(5,'E'),(6,'F'),(7,'G')]:
    put(w2,r,col,'='+_s(ltr),NUM,bold=True,fill=SUB)
put(w2,r,8,f'=IFERROR(E{r}/D{r},"-")',PCT,bold=True,fill=SUB)
put(w2,r,9,f'=IFERROR(F{r}/D{r},"-")',PCT,bold=True,fill=SUB)
put(w2,r,10,'',fill=SUB); put(w2,r,11,'',fill=SUB)
put(w2,r,12,f'=IF(ABS(F{r})>ABS(E{r})*2,"가격이 주도",IF(ABS(E{r})>ABS(F{r})*2,"물량이 주도","혼합"))',align='center',sz=9,bold=True,fill=SUB)
SUM2D=r
w2.freeze_panes='B6'
print('분석2 완료',D2,SUM2D)
