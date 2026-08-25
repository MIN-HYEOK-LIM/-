# -*- coding: utf-8 -*-
"""분석7_HS4상세 — 4단위 321개 전수"""
SPAN7=25
w7=newsheet('분석7_HS4상세',[8,44,8,9,22,12,12,12,10,10,10,10,10,10,10,12,12,9,9,9,9,9,9,9,52])
title(w7,1,'⑦ HS 4단위 전수 — 321개 코드 · 2026년 7개월을 월별로',SPAN7)
note(w7,2,1,'수집한 10개 류(27·29·39·71·72·84·85·87·89·90)의 4단위 코드 전부입니다. 누계 YoY는 어느 달에 무슨 일이 있었는지를 감춥니다. '
            '오른쪽 월별 YoY 7개 열에서 변화가 언제 시작됐는지 확인하십시오. 자동 필터로 섹터·구분을 걸러 쓰십시오.',SPAN7)
note(w7,3,1,'기여도 = 그 코드의 증감액 ÷ 국가 총수출 증감액. 비중 = 그 코드의 2026년 1~7월 수출 ÷ 국가 총수출.',SPAN7)
r=4
H7=['HS4','품목명','2단위','섹터','구분','2024','2025','2026','25 YoY','26 YoY','기여도','비중','3M YoY','가속도','단가Δ','물량Δ','수입 26','수지 26']\
   +[f'{m} YoY' for m in M17['2026']]+['해석']
for i,h in enumerate(H7): head(w7,r,i+1,h)
w7.row_dimensions[r].height=30
r+=1; S7=r
_t24='+'.join(TX(k) for k in K24); _t25='+'.join(TX(k) for k in K25); _t26='+'.join(TX(k) for k in K26)
for i,code in enumerate(CODES):
    rr=S7+i
    sc=SEC_OF.get(code,'-')
    put(w7,rr,1,code,align='center',sz=9); put(w7,rr,2,CNAME[code],sz=8,wrap=True)
    put(w7,rr,3,code[:2],align='center',sz=9)
    put(w7,rr,4,SEC_NAME.get(sc,'-'),sz=9)
    put(w7,rr,5,'반도체' if sc in SEMI else '비반도체',align='center',sz=9)
    for j,ks in enumerate([K24,K25,K26]):
        put(w7,rr,6+j,'='+'+'.join(pref(PEX,code,k) for k in ks),NUM,sz=9)
    put(w7,rr,9,f'=IFERROR(G{rr}/F{rr}-1,"-")',PCT,sz=9)
    put(w7,rr,10,f'=IFERROR(H{rr}/G{rr}-1,"-")',PCT,sz=9,bold=True)
    put(w7,rr,11,f'=IFERROR((H{rr}-G{rr})/(({_t26})-({_t25})),"-")',PCT2,sz=9)
    put(w7,rr,12,f'=IFERROR(H{rr}/({_t26}),"-")',PCT2,sz=9)
    m3='+'.join(pref(PEX,code,k) for k in KL3); m3p='+'.join(pref(PEX,code,k) for k in KL3P)
    m3b='+'.join(pref(PEX,code,k) for k in KL3B); m3bp='+'.join(pref(PEX,code,k) for k in KL3BP)
    put(w7,rr,13,f'=IFERROR(({m3})/({m3p})-1,"-")',PCT,sz=9)
    put(w7,rr,14,f'=IFERROR((M{rr}-(({m3b})/({m3bp})-1))*100,"-")',PP0,sz=9)
    w26='+'.join(pref(PWT,code,k) for k in K26); w25='+'.join(pref(PWT,code,k) for k in K25)
    put(w7,rr,15,f'=IFERROR((H{rr}/({w26}))/(G{rr}/({w25}))-1,"-")',PCT,sz=9)
    put(w7,rr,16,f'=IFERROR(({w26})/({w25})-1,"-")',PCT,sz=9)
    imp='+'.join(pref(PIM,code,k) for k in K26)
    put(w7,rr,17,'='+imp,NUM,sz=9); put(w7,rr,18,f'=H{rr}-Q{rr}',NUM,sz=9)
    for j,k in enumerate(K26):
        put(w7,rr,19+j,f'=IFERROR({pref(PEX,code,k)}/{pref(PEX,code,k-12)}-1,"-")',PCT,sz=8)
    mr=f'S{rr}:Y{rr}'
    put(w7,rr,26,
        f'=IF(J{rr}="-","비교할 전년 값이 없습니다.",'
        f'IF(J{rr}>1,"2배 이상 급증. ",IF(J{rr}>0.2,"큰 폭 증가. ",IF(J{rr}>0.05,"증가. ",IF(J{rr}>-0.05,"보합. ",IF(J{rr}>-0.2,"감소. ","큰 폭 감소. ")))))'
        f'&IF(NOT(AND(ISNUMBER(O{rr}),ISNUMBER(P{rr}))),"",IF(O{rr}>ABS(P{rr})*2,"단가 상승이 주도(단가 "&TEXT(O{rr},"+0.0%")&", 물량 "&TEXT(P{rr},"+0.0%")&"). ",'
        f'IF(P{rr}>ABS(O{rr})*2,"물량 증가가 주도(물량 "&TEXT(P{rr},"+0.0%")&", 단가 "&TEXT(O{rr},"+0.0%")&"). ","단가·물량이 함께 움직임. ")))'
        f'&"7개월 중 "&COUNTIF({mr},">0")&"개월 증가."'
        f'&IF(ISNUMBER(N{rr})," 최근 3개월 속도는 직전 3개월 대비 "&TEXT(N{rr}/100,"+0.0%")&"p.","")',sz=8,wrap=True)
E7=S7+len(CODES)-1
w7.auto_filter.ref=f'A{S7-1}:{L(SPAN7)}{E7}'
w7.freeze_panes=f'F{S7}'
r=E7+2
sec(w7,r,'합계 (수집한 10개 류)',SPAN7); r+=1
put(w7,r,1,'합계',bold=True,align='center'); put(w7,r,2,'4단위를 수집한 10개 류의 총계',sz=9)
for col,ltr in [(6,'F'),(7,'G'),(8,'H'),(17,'Q')]:
    put(w7,r,col,f'=SUM({ltr}{S7}:{ltr}{E7})',NUM,bold=True)
put(w7,r,9,f'=IFERROR(G{r}/F{r}-1,"-")',PCT,bold=True)
put(w7,r,10,f'=IFERROR(H{r}/G{r}-1,"-")',PCT,bold=True)
put(w7,r,12,f'=IFERROR(H{r}/({_t26}),"-")',PCT2,bold=True)
put(w7,r,26,f'="수집 10개 류가 2026년 1~7월 국가 총수출의 "&TEXT(H{r}/({_t26}),"0.0%")&"를 차지합니다. 나머지 영역은 분석3 [C] 부문별 표에서 확인하십시오."',sz=8,wrap=True)
print('분석7 완료',S7,E7)
