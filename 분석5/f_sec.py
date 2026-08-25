# -*- coding: utf-8 -*-
"""분석5_섹터월별  →  분석4_섹터스코어보드"""
# ══════════════════════════════════════════════════════════════════
# 분석5_섹터월별
# ══════════════════════════════════════════════════════════════════
w5=newsheet('분석5_섹터월별',[9,24]+[11]*NM)
SPAN5=2+NM
title(w5,1,'⑤ 섹터별 월별 수출·수입 — 38개 섹터 × 31개월',SPAN5)
note(w5,2,1,'스코어보드(분석4)와 그래프(분석10)가 모두 이 표를 참조합니다. [B]는 명목 전년동월비, [D]는 조업일수로 나눈 일평균 기준 전년동월비입니다. '
            '2026년 2월처럼 조업일수가 3일 적은 달은 명목으로 보면 부진해 보이지만 일평균으로는 증가인 경우가 있어 두 가지를 함께 둡니다.',SPAN5)
r=4
sec(w5,r,'[A] 섹터별 월별 수출금액 (천달러)',SPAN5); r+=1
A5H=r; head(w5,r,1,'코드'); head(w5,r,2,'섹터')
for k in range(NM): head(w5,r,3+k,MONTHS[k])
r+=1; A5=r
for sc,nm,g,ds in SEC_LIVE:
    put(w5,r,1,sc,align='center',sz=9); put(w5,r,2,nm,sz=9,bold=(sc in SEMI))
    for k in range(NM): put(w5,r,3+k,'='+ssum(PEX,sc,k),NUM,sz=8)
    r+=1
E5A=r-1; r+=1

sec(w5,r,'[B] 섹터별 월별 전년동월비 (명목)',SPAN5); r+=1
B5H=r; head(w5,r,1,'코드'); head(w5,r,2,'섹터')
for k in KY: head(w5,r,3+k-12,MONTHS[k])
r+=1; B5=r
for i,(sc,nm,g,ds) in enumerate(SEC_LIVE):
    a=A5+i
    put(w5,r,1,sc,align='center',sz=9); put(w5,r,2,nm,sz=9,bold=(sc in SEMI))
    for k in KY:
        put(w5,r,3+k-12,f'=IFERROR({L(3+k)}{a}/{L(3+k-12)}{a}-1,"-")',PCT,sz=8)
    r+=1
E5B=r-1; r+=1

sec(w5,r,'[C] 섹터별 월별 수입금액 (천달러)',SPAN5); r+=1
C5H=r; head(w5,r,1,'코드'); head(w5,r,2,'섹터')
for k in range(NM): head(w5,r,3+k,MONTHS[k])
r+=1; C5=r
for sc,nm,g,ds in SEC_LIVE:
    put(w5,r,1,sc,align='center',sz=9); put(w5,r,2,nm,sz=9)
    for k in range(NM): put(w5,r,3+k,'='+ssum(PIM,sc,k),NUM,sz=8)
    r+=1
E5C=r-1; r+=1

sec(w5,r,'[D] 섹터별 월별 전년동월비 (조업일수 보정 = 일평균 기준)',SPAN5); r+=1
D5H=r; head(w5,r,1,'코드'); head(w5,r,2,'섹터')
for k in KY: head(w5,r,3+k-12,MONTHS[k])
r+=1; D5=r
for i,(sc,nm,g,ds) in enumerate(SEC_LIVE):
    a=A5+i
    put(w5,r,1,sc,align='center',sz=9); put(w5,r,2,nm,sz=9,bold=(sc in SEMI))
    for k in KY:
        put(w5,r,3+k-12,f'=IFERROR(({L(3+k)}{a}/{TD(k)})/({L(3+k-12)}{a}/{TD(k-12)})-1,"-")',PCT,sz=8)
    r+=1
E5D=r-1; r+=1

sec(w5,r,'[E] 섹터별 월별 수출중량 (톤) — 물량 흐름',SPAN5); r+=1
E5H=r; head(w5,r,1,'코드'); head(w5,r,2,'섹터')
for k in range(NM): head(w5,r,3+k,MONTHS[k])
r+=1; F5=r
for sc,nm,g,ds in SEC_LIVE:
    put(w5,r,1,sc,align='center',sz=9); put(w5,r,2,nm,sz=9)
    for k in range(NM): put(w5,r,3+k,'='+ssum(PWT,sc,k),NUM1,sz=8)
    r+=1
E5E=r-1
w5.freeze_panes='C6'
print('분석5 완료',A5,B5,C5,D5,F5)

# ══════════════════════════════════════════════════════════════════
# 분석4_섹터스코어보드
# ══════════════════════════════════════════════════════════════════
w4=newsheet('분석4_섹터스코어보드',[8,22,13,9,26,12,8,11,11,11,11,10,10,10,10,9,13,13,13,9,9,11,10,9,46])
SPAN4=25
title(w4,1,'④ 섹터 스코어보드 — 38개 섹터를 같은 잣대로 세운 표',SPAN4)
note(w4,2,1,'최근 3개월(2026.05~07)을 전년 동월과 비교한 값이 「3M YoY」, 그 직전 3개월(2026.02~04)과의 차이가 「가속도」입니다. 국면은 이 둘의 부호 조합으로 자동 판정합니다.',SPAN4)
note(w4,3,1,'성장동력 : 단가 상승이 주도하면 사이클(되돌림 위험), 물량 증가가 주도하면 실수요(지속성 높음). 같은 성장률이라도 성격이 다르므로 종합점수에 물량 기여를 가중했습니다.',SPAN4)
note(w4,5,1,'Z~AC 열은 국면별 규모를 담는 계산용 보조 열입니다(아래 국면별 분포표가 참조합니다).',SPAN4)
note(w4,4,1,'※ 무역통계 기반 정량 스크리닝입니다. 개별 기업의 실적·밸류에이션·수급은 반영돼 있지 않으므로 종목 선택은 별도 검증이 필요합니다. 「구분」 열로 반도체/비반도체를 걸러 보십시오.',SPAN4,color=RED)
r=6
H4C=['코드','섹터','그룹','구분','포함 HS 4단위','2026 수출','비중','2024','2025','2026','25 YoY','26 YoY',
     '3M YoY','직전 3M','가속도','국면','단가Δ','물량Δ','성장동력','증가 월수','변동성','수입 26','수지 Δ','종합','등급']
for i,h in enumerate(H4C): head(w4,r,i+1,h)
for _j,_ph in enumerate(['가속 성장','성장 둔화','바닥 통과','침체 심화']):
    head(w4,r,26+_j,_ph+' 규모',fill=HDR2)
    w4.column_dimensions[L(26+_j)].width=13
w4.row_dimensions[r].height=30
r+=1; S4=r
_tot26='+'.join(TX(k) for k in K26)
for i,(sc,nm,g,ds) in enumerate(SEC_LIVE):
    a=A5+i; b=B5+i; c=C5+i; d=D5+i
    m=lambda ks,rw=None: '+'.join(f"'분석5_섹터월별'!{L(3+k)}{rw or a}" for k in ks)
    wt=lambda ks: '+'.join(f"({ssum(PWT,sc,k)})" for k in ks)
    put(w4,r,1,sc,align='center',bold=True); put(w4,r,2,nm,bold=True)
    put(w4,r,3,g,align='center',sz=9)
    put(w4,r,4,'반도체' if sc in SEMI else '비반도체',align='center',sz=9,
        fill=(WARN if sc in SEMI else None))
    put(w4,r,5,'·'.join(SEC_CODES[sc][:10])+('…' if len(SEC_CODES[sc])>10 else ''),sz=8,wrap=True)
    put(w4,r,6,'='+m(K26),NUM); put(w4,r,7,f'=IFERROR(F{r}/({_tot26}),"-")',PCT2)
    put(w4,r,8,'='+m(K24),NUM); put(w4,r,9,'='+m(K25),NUM); put(w4,r,10,f'=F{r}',NUM)
    put(w4,r,11,f'=IFERROR(I{r}/H{r}-1,"-")',PCT)
    put(w4,r,12,f'=IFERROR(J{r}/I{r}-1,"-")',PCT)
    put(w4,r,13,f'=IFERROR(({m(KL3)})/({m(KL3P)})-1,"-")',PCT,bold=True)
    put(w4,r,14,f'=IFERROR(({m(KL3B)})/({m(KL3BP)})-1,"-")',PCT)
    put(w4,r,15,f'=IFERROR((M{r}-N{r})*100,"-")',PP,bold=True)
    put(w4,r,16,f'=IF(M{r}="-","-",IF(AND(M{r}>0,O{r}>0),"가속 성장",IF(M{r}>0,"성장 둔화",IF(O{r}>0,"바닥 통과","침체 심화"))))',align='center',bold=True)
    put(w4,r,17,f'=IFERROR((J{r}/({wt(K26)}))/(I{r}/({wt(K25)}))-1,"-")',PCT)
    put(w4,r,18,f'=IFERROR(({wt(K26)})/({wt(K25)})-1,"-")',PCT)
    put(w4,r,19,f'=IF(OR(Q{r}="-",R{r}="-"),"-",IF(Q{r}>ABS(R{r})*2,"가격 주도",IF(R{r}>ABS(Q{r})*2,"물량 주도","혼합")))',align='center',sz=9)
    put(w4,r,20,f"=COUNTIF('분석5_섹터월별'!{L(3+K26[0]-12)}{b}:{L(3+K26[-1]-12)}{b},\">0\")",NUM)
    put(w4,r,21,f"=IFERROR(STDEV('분석5_섹터월별'!{L(3+K26[0]-12)}{b}:{L(3+K26[-1]-12)}{b}),\"-\")",PCT)
    imp26=m(K26,c); imp25=m(K25,c)
    put(w4,r,22,'='+imp26,NUM)
    put(w4,r,23,f'=(J{r}-({imp26}))-(I{r}-({imp25}))',NUM)
    put(w4,r,24,
        f'=IFERROR(ROUND(MEDIAN(0,M{r},1)*30'
        f'+MEDIAN(0,O{r}/100+0.3,0.6)/0.6*25'
        f'+IF(S{r}="물량 주도",20,IF(S{r}="혼합",12,5))*T{r}/7'
        f'+MEDIAN(0,G{r}*100/10,1)*15'
        f'+IF(U{r}="-",5,MEDIAN(0,1-U{r}/1.5,1)*10),1),"-")','0.0',bold=True)
    put(w4,r,25,f'=IF(X{r}="","-",IF(AND(X{r}>=55,P{r}="가속 성장",OR(U{r}="-",U{r}<1.2)),"비중확대",'
                f'IF(AND(X{r}>=42,OR(P{r}="가속 성장",P{r}="바닥 통과")),"중립+",'
                f'IF(AND(P{r}="성장 둔화",O{r}<-30),"비중축소 검토",'
                f'IF(P{r}="침체 심화","회피",IF(AND(P{r}="바닥 통과",O{r}>5,T{r}<=2),"매집 검토","중립"))))))',align='center',bold=True)
    for _j,_ph in enumerate(['가속 성장','성장 둔화','바닥 통과','침체 심화']):
        put(w4,r,26+_j,f'=IF($P{r}="{_ph}",$F{r},"")',NUM,sz=8)
    r+=1
E4=r-1
w4.auto_filter.ref=f'A{S4-1}:{L(SPAN4)}{E4}'
w4.freeze_panes=f'F{S4}'
r+=1
sec(w4,r,'국면별 분포',SPAN4); r+=1
for i,h in enumerate(['국면','섹터 수','반도체','비반도체','2026 수출 합계','전체 비중','대표 섹터(규모 상위)']): head(w4,r,i+1,h)
r+=1; Q4=r
for ph,dsc in [('가속 성장','최근 3개월도 늘고 속도까지 빨라진 구간'),('성장 둔화','아직 플러스지만 속도가 꺾인 구간'),
               ('바닥 통과','아직 마이너스지만 감소폭이 줄어드는 구간'),('침체 심화','마이너스가 더 깊어지는 구간')]:
    rng=f'$P${S4}:$P${E4}'
    put(w4,r,1,ph,bold=True,align='center')
    put(w4,r,2,f'=COUNTIF({rng},"{ph}")',NUM)
    put(w4,r,3,f'=COUNTIFS({rng},"{ph}",$D${S4}:$D${E4},"반도체")',NUM)
    put(w4,r,4,f'=COUNTIFS({rng},"{ph}",$D${S4}:$D${E4},"비반도체")',NUM)
    put(w4,r,5,f'=SUMIF({rng},"{ph}",$F${S4}:$F${E4})',NUM)
    put(w4,r,6,f'=IFERROR(E{r}/({_tot26}),"-")',PCT2)
    _hc=L(26+['가속 성장','성장 둔화','바닥 통과','침체 심화'].index(ph))
    put(w4,r,7,f'=IFERROR(INDEX($B${S4}:$B${E4},MATCH(MAX(${_hc}${S4}:${_hc}${E4}),${_hc}${S4}:${_hc}${E4},0))&IF(B{r}>1," 외 "&(B{r}-1)&"개",""),"-")',sz=9)
    put(w4,r,8,dsc,sz=9,color='595959')
    r+=1
print('분석4 완료',S4,E4)
