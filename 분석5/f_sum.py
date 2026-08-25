# -*- coding: utf-8 -*-
"""분석1_핵심요약 (원본 정합성 검증 포함)"""
SPAN1=12
w1=newsheet('분석1_핵심요약',[28,17,17,15,13,13,13,13,13,13,13,13])
title(w1,1,'① 핵심 요약 — 이 파일이 말하는 것',SPAN1)
note(w1,2,1,'2024.01~2026.07 관세청 수출입실적. 원본 9개 시트는 수정하지 않았고, 분석 셀은 전부 원본을 참조하는 수식입니다.',SPAN1)
def SRCTOT(col): return f'IFERROR(VALUE(SUBSTITUTE(TRIM(\'수출입 총괄\'!{col}6),",","")),"")'
r=4
sec(w1,r,'[A] 2026년 1~7월 한눈에',SPAN1); r+=1
for i,h in enumerate(['항목','2024 1~7','2025 1~7','2026 1~7','25 YoY','26 YoY','최근 3M YoY','비고']): head(w1,r,i+1,h)
r+=1; A1=r
_m=lambda f,ks: '+'.join(f(k) for k in ks)
ROWS1=[('수출 금액 (천$)',TX,NUM,'국가 총수출'),
       ('수입 금액 (천$)',TI,NUM,'국가 총수입'),
       ('조업일수',TD,NUM1,'2026년은 전년보다 짧은 달이 섞여 있습니다'),
       ('수출 중량 (톤)',TWX,NUM1,'물량 — 금액과 따로 봐야 합니다')]
for lab,fn,fmt,memo in ROWS1:
    put(w1,r,1,lab,bold=True)
    for j,ks in enumerate([K24,K25,K26]): put(w1,r,2+j,'='+_m(fn,ks),fmt)
    put(w1,r,5,f'=IFERROR(C{r}/B{r}-1,"-")',PCT); put(w1,r,6,f'=IFERROR(D{r}/C{r}-1,"-")',PCT,bold=True)
    put(w1,r,7,f'=IFERROR(({_m(fn,KL3)})/({_m(fn,KL3P)})-1,"-")',PCT)
    put(w1,r,8,memo,sz=9,color='595959'); r+=1
put(w1,r,1,'무역수지 (천$)',bold=True)
for j in range(3): put(w1,r,2+j,f'={L(2+j)}{A1}-{L(2+j)}{A1+1}',NUM)
put(w1,r,5,'',None); put(w1,r,6,'',None); put(w1,r,7,'',None)
put(w1,r,8,'수출−수입',sz=9,color='595959'); r+=1
put(w1,r,1,'일평균 수출 (천$)',bold=True)
for j in range(3): put(w1,r,2+j,f'=IFERROR({L(2+j)}{A1}/{L(2+j)}{A1+2},"-")',NUM)
put(w1,r,5,f'=IFERROR(C{r}/B{r}-1,"-")',PCT); put(w1,r,6,f'=IFERROR(D{r}/C{r}-1,"-")',PCT,bold=True)
put(w1,r,7,'',None); put(w1,r,8,'조업일수 효과를 걷어낸 실질 흐름',sz=9,color='595959'); r+=1
put(w1,r,1,'수출 단가 (천$/톤)',bold=True)
for j in range(3): put(w1,r,2+j,f'=IFERROR({L(2+j)}{A1}/{L(2+j)}{A1+3},"-")',UNIT)
put(w1,r,5,f'=IFERROR(C{r}/B{r}-1,"-")',PCT); put(w1,r,6,f'=IFERROR(D{r}/C{r}-1,"-")',PCT,bold=True)
put(w1,r,7,'',None); put(w1,r,8,'금액÷중량. 성장이 가격에서 왔는지 물량에서 왔는지의 기준',sz=9,color='595959'); r+=2

sec(w1,r,'[B] 반도체와 비반도체 — 2026년은 정말 반도체만의 이야기인가',SPAN1); r+=1
for i,h in enumerate(['구분','2025 1~7','2026 1~7','26 YoY','조업보정 26 YoY','최근 3M YoY','2026 비중','증가 기여','판정']): head(w1,r,i+1,h)
r+=1; B1=r
W3='분석3_수출구조분해'
for lab,src in [('국가 총수출',R_TOT),('반도체 (8542·8473·8523·8541)',R_SEMI),('반도체 장비·검사 (8486·9030·9031)',R_EQP),('비반도체 (나머지 전부)',R_NS)]:
    m=lambda ks: '+'.join(f"'{W3}'!{L(3+k)}{src}" for k in ks)
    dm=lambda ks: '+'.join(f"('{W3}'!{L(3+k)}{src}/{TD(k)})" for k in ks)
    hl=GOOD if '비반도체' in lab else None
    put(w1,r,1,lab,bold=True,fill=hl)
    put(w1,r,2,'='+m(K25),NUM,fill=hl); put(w1,r,3,'='+m(K26),NUM,fill=hl)
    put(w1,r,4,f'=IFERROR(C{r}/B{r}-1,"-")',PCT,bold=True,fill=hl)
    put(w1,r,5,f'=IFERROR(({dm(K26)})/({dm(K25)})-1,"-")',PCT,bold=True,fill=hl)
    put(w1,r,6,f'=IFERROR(({m(KL3)})/({m(KL3P)})-1,"-")',PCT,fill=hl)
    put(w1,r,7,f'=IFERROR(C{r}/C${B1},"-")',PCT,fill=hl)
    put(w1,r,8,f'=IFERROR((C{r}-B{r})/(C${B1}-B${B1}),"-")',PCT,fill=hl)
    put(w1,r,9,f'=IF(E{r}="-","-",IF(E{r}>0.1,"강한 증가",IF(E{r}>0.03,"증가",IF(E{r}>-0.03,"보합","감소"))))',align='center',bold=True,fill=hl)
    r+=1
E_B1=r-1
put(w1,r,1,'비반도체가 플러스였던 달',bold=True)
put(w1,r,2,f"=COUNTIF('{W3}'!{L(3+K26[0])}{Y_NSD}:{L(3+K26[-1])}{Y_NSD},\">0\")&\" / 7개월 (조업일수 보정 기준)\"",sz=10)
w1.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
put(w1,r,6,'명목 기준으로는',sz=9,color='595959')
put(w1,r,7,f"=COUNTIF('{W3}'!{L(3+K26[0])}{Y_NS}:{L(3+K26[-1])}{Y_NS},\">0\")&\" / 7개월\"",sz=9)
r+=2
S='분석4_섹터스코어보드'
sec(w1,r,'[C] 국가 수출 증가를 만든 것과 깎은 것 — 섹터 상·하위',SPAN1); r+=1
note(w1,r,1,'기여도 = 그 섹터의 2026년 증감액 ÷ 국가 총수출 증감액. 반도체 4개 섹터는 「구분」 열로 구별됩니다.',SPAN1); r+=1
for i,h in enumerate(['순위','섹터','구분','2025 1~7','2026 1~7','26 YoY','증감액','기여도','3M YoY','가속도','국면','등급']): head(w1,r,i+1,h)
r+=1; C1=r
_szmap={sc:sum(H4[c][m][1] for c in SEC_CODES[sc] for m in M17['2026'] if c in H4 and m in H4[c]) for sc,nm,g,ds in SEC_LIVE}
_dmap={}
for sc,nm,g,ds in SEC_LIVE:
    a=sum(H4[c][m][1] for c in SEC_CODES[sc] for m in M17['2025'] if c in H4 and m in H4[c])
    _dmap[sc]=_szmap[sc]-a
_rank=sorted(SEC_LIVE,key=lambda t:-_dmap[t[0]])
_tot_d=f"(({'+'.join(TX(k) for k in K26)})-({'+'.join(TX(k) for k in K25)}))"
def idea_rows(lst,startrank):
    global r
    for n,(sc,nm,g,ds) in enumerate(lst):
        sr=S4+IDX_OF[sc]
        put(w1,r,1,startrank+n if startrank>0 else '',align='center',sz=9)
        put(w1,r,2,f"='{S}'!B{sr}",sz=9,bold=True)
        put(w1,r,3,f"='{S}'!D{sr}",align='center',sz=9)
        put(w1,r,4,f"='{S}'!I{sr}",NUM,sz=9); put(w1,r,5,f"='{S}'!J{sr}",NUM,sz=9)
        put(w1,r,6,f"='{S}'!L{sr}",PCT,sz=9)
        put(w1,r,7,f"='{S}'!J{sr}-'{S}'!I{sr}",NUM,sz=9)
        put(w1,r,8,f"=IFERROR(G{r}/{_tot_d},\"-\")",PCT2,sz=9,bold=True)
        put(w1,r,9,f"='{S}'!M{sr}",PCT,sz=9); put(w1,r,10,f"='{S}'!O{sr}",PP0,sz=9)
        put(w1,r,11,f"='{S}'!P{sr}",align='center',sz=9); put(w1,r,12,f"='{S}'!Y{sr}",align='center',sz=9,bold=True)
        r+=1
idea_rows(_rank[:12],1)
put(w1,r,1,'▼ 하위',bold=True,align='center',sz=9,fill=BAD)
for c in range(2,13): put(w1,r,c,'',fill=BAD)
r+=1
idea_rows(_rank[-6:],0)
r+=1
sec(w1,r,'[D] 국면 분포와 성장의 질',SPAN1); r+=1
for i,h in enumerate(['구분','섹터 수','반도체','비반도체','2026 수출','비중','설명']): head(w1,r,i+1,h)
r+=1
_rngP=f"'{S}'!$P${S4}:$P${E4}"; _rngD=f"'{S}'!$D${S4}:$D${E4}"; _rngF=f"'{S}'!$F${S4}:$F${E4}"; _rngS=f"'{S}'!$S${S4}:$S${E4}"
_t26='+'.join(TX(k) for k in K26)
for lab,rng,key,dsc in [('가속 성장',_rngP,'가속 성장','최근 3개월도 늘고 속도까지 빨라진 구간'),
                        ('성장 둔화',_rngP,'성장 둔화','아직 플러스지만 속도가 꺾인 구간'),
                        ('바닥 통과',_rngP,'바닥 통과','아직 마이너스지만 감소폭이 줄어드는 구간'),
                        ('침체 심화',_rngP,'침체 심화','마이너스가 더 깊어지는 구간'),
                        ('― 물량 주도',_rngS,'물량 주도','실수요형 성장. 가격 되돌림 위험이 낮습니다'),
                        ('― 가격 주도',_rngS,'가격 주도','사이클성 성장. 단가가 꺾이면 곧바로 반전됩니다'),
                        ('― 혼합',_rngS,'혼합','단가와 물량이 함께 움직이는 경우')]:
    put(w1,r,1,lab,bold=True,align='center')
    put(w1,r,2,f'=COUNTIF({rng},"{key}")',NUM)
    put(w1,r,3,f'=COUNTIFS({rng},"{key}",{_rngD},"반도체")',NUM)
    put(w1,r,4,f'=COUNTIFS({rng},"{key}",{_rngD},"비반도체")',NUM)
    put(w1,r,5,f'=SUMIF({rng},"{key}",{_rngF})',NUM)
    put(w1,r,6,f'=IFERROR(E{r}/({_t26}),"-")',PCT2)
    put(w1,r,7,dsc,sz=9,color='595959'); r+=1
r+=1
sec(w1,r,'[E] 원본 정합성 검증 — 개별항목의 합과 총합이 맞는가',SPAN1); r+=1
note(w1,r,1,'원본은 숫자가 텍스트(쉼표 포함)로 저장돼 있어 정제 시트를 거쳐 검증합니다. 원본 시트는 수정하지 않았습니다.',SPAN1); r+=1
put(w1,r,1,'허용오차(중량, 톤)',bold=True); put(w1,r,2,1.0,NUM1,color=BLUE,fill=IN)
put(w1,r,3,'허용오차(금액, 천$)',bold=True); put(w1,r,4,50,NUM,color=BLUE,fill=IN)
TW,TA=f'$B${r}',f'$D${r}'; r+=1
for i,h in enumerate(['항목','총계행(원본)','31개월 합계','차이','판정','','','']): head(w1,r,i+1,h)
r+=1
for nm_,sc_,dc_,fmt_,tol in [('조업일수','B','E',NUM1,TW),('수출 건수','C','F',NUM,TA),('수출 중량(톤)','D','G',NUM1,TW),
        ('수출 금액(천$)','E','H',NUM,TA),('수입 건수','F','I',NUM,TA),('수입 중량(톤)','G','J',NUM1,TW),
        ('수입 금액(천$)','H','K',NUM,TA),('무역수지(천$)','I','L',NUM,TA)]:
    put(w1,r,1,nm_,bold=True); put(w1,r,2,'='+SRCTOT(sc_),fmt_)
    put(w1,r,3,f'=SUM({DT}!${dc_}${T0}:${dc_}${T1})',fmt_); put(w1,r,4,f'=B{r}-C{r}',NUM1)
    put(w1,r,5,f'=IF(ABS(D{r})<={tol},"일치(반올림오차 이내)","불일치")',align='center'); r+=1
put(w1,r,1,'품목별 합 vs 총괄',bold=True)
put(w1,r,2,f'=SUM({DH}!$H${H0}:$H${H1})',NUM)
put(w1,r,3,f'=SUM({DT}!$H${T0}:$H${T1})',NUM)
put(w1,r,4,f'=IFERROR(B{r}/C{r}-1,"-")',PCT2)
put(w1,r,5,f'=IF(ABS(D{r})<0.001,"일치(HS 98 특수분류 제외분)","차이 확인 필요")',align='center')
put(w1,r,6,'품목별(HS)에는 특수분류 HS 98이 빠져 총괄보다 0.03~0.05% 작습니다. 정상입니다.',sz=9,color='595959')
w1.merge_cells(start_row=r,start_column=6,end_row=r,end_column=SPAN1)
r+=1
put(w1,r,1,'HS4 합 vs 총괄',bold=True)
put(w1,r,2,f"=SUM('분석7_HS4상세'!$H${S7}:$H${E7})",NUM)
put(w1,r,3,'='+'+'.join(TX(k) for k in K26),NUM)
put(w1,r,4,f'=IFERROR(B{r}/C{r},"-")',PCT2)
put(w1,r,5,'참고',align='center')
put(w1,r,6,'4단위를 수집한 10개 류가 2026년 1~7월 국가 총수출에서 차지하는 몫입니다. 나머지는 분석3 [C] 부문별 표로 커버합니다.',sz=9,color='595959')
w1.merge_cells(start_row=r,start_column=6,end_row=r,end_column=SPAN1)
print('분석1 완료',A1,B1,C1)
