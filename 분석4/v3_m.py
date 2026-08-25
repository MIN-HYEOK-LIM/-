# -*- coding: utf-8 -*-
# 공통 헬퍼 (누계 참조용 — 월별 분해의 합계 확인에만 사용)
def hs_y(code,y,col='H',mmax=7):
    return f'SUMIFS({hR(col)},{hR("E")},"{code}",{hR("B")},{y},{hR("C")},"<={mmax}")'
def tot_sum(y,mmax=None,col='H'):
    cond=f',{tR("C")},"<={mmax}"' if mmax else ''
    return f'SUMIFS({tR(col)},{tR("B")},{y}{cond})'
def it_sum(keys,y,mmax=None,col='G'):
    cond=f',{nR("C")},"<={mmax}"' if mmax else ''
    return '+'.join(f'SUMIFS({nR(col)},{nR("E")},"{k}",{nR("B")},{y}{cond})' for k in keys)
_e26=tot_sum(2026,7); _e25=tot_sum(2025,7); _e24=tot_sum(2024,7)
# 월별 매트릭스 + 분석5~7 (월별 분해 중심으로 재구성)
MONTH_LABEL=MONTHS
M26L=[m for m in MONTHS if m.startswith('2026')]      # 2026.01~07
YOY_M=[m for m in MONTHS if not m.startswith('2024')] # 2025.01~2026.07 (YoY 계산 가능 구간)

# ══════════════════════════════════════════════════════════════════
# 분석4B_월별매트릭스 (모든 월별 분석의 원천)
# ══════════════════════════════════════════════════════════════════
wm=wb.create_sheet('분석4B_월별매트릭스'); wm.sheet_view.showGridLines=False
widths(wm,[8,30]+[12]*NM)
title(wm,1,'④-B 월별 매트릭스 — 모든 항목의 31개월 월별 값 (이후 시트의 월별 YoY 원천)',2+NM)
note(wm,2,1,'이 시트는 품목·부문·성질별 항목의 월별 값을 한 번만 계산해 두는 곳입니다. 분석6·7·8의 월별 증감은 모두 이 표를 참조하므로, 값이 바뀌면 월별 분석 전체가 함께 갱신됩니다.',2+NM)
r=4
def matrix(rows_def, label, key_fmt, valcol_expr):
    global r
    sec(wm,r,label,2+NM); r+=1
    hr=r; head(wm,r,1,'코드'); head(wm,r,2,'항목')
    for k in range(NM): head(wm,r,3+k,MONTHS[k])
    r+=1; r0=r
    for code,name in rows_def:
        put(wm,r,1,code,align='center',sz=9).number_format='@'
        put(wm,r,2,name,sz=9)
        for k in range(NM):
            put(wm,r,3+k,valcol_expr(code,MONTHS[k]),NUM,sz=8)
        r+=1
    r+=1
    return hr,r0,r-2
HS_ORDER_M=sorted(HS)
MX_HS_EX=matrix([(c,HSNAME[c]) for c in HS_ORDER_M],'[A] HS 류별 월별 수출금액 (천달러)',None,
    lambda c,m: f'=SUMIFS({hR("H")},{hR("E")},"{c}",{hR("A")},"{m}")')
MX_HS_IM=matrix([(c,HSNAME[c]) for c in HS_ORDER_M],'[B] HS 류별 월별 수입금액 (천달러)',None,
    lambda c,m: f'=SUMIFS({hR("J")},{hR("E")},"{c}",{hR("A")},"{m}")')
SEC_ROWS=[(s,f'{s}. {n}') for s,n,chs in SECTIONS if [c for c in chs if c in HS]]
SEC_CH={s:[c for c in chs if c in HS] for s,n,chs in SECTIONS}
MX_SEC=matrix(SEC_ROWS,'[C] HS 부문별 월별 수출금액 (천달러)',None,
    lambda s,m: '='+'+'.join(f'SUMIFS({hR("H")},{hR("E")},"{c}",{hR("A")},"{m}")' for c in SEC_CH[s]))
MX_X=matrix([(str(i+1),n) for i,n in enumerate(OLDX_ITEMS)],'[D] 성질별(수출) 월별 금액 (천달러)',None,
    lambda i,m: f'=SUMIFS({xR("G")},{xR("E")},$B{{r}},{xR("A")},"{m}")')
MX_M=matrix([(str(i+1),n) for i,n in enumerate(OLDM_ITEMS)],'[E] 성질별(수입) 월별 금액 (천달러)',None,
    lambda i,m: f'=SUMIFS({mR("G")},{mR("E")},$B{{r}},{mR("A")},"{m}")')
MX_N=matrix([(str(i+1),n) for i,n in enumerate(NEW_ITEMS)],'[F] 신성질별 월별 수출금액 (천달러)',None,
    lambda i,m: f'=SUMIFS({nR("G")},{nR("E")},$B{{r}},{nR("A")},"{m}")')
MX_NI=matrix([(str(i+1),n) for i,n in enumerate(NEW_ITEMS)],'[G] 신성질별 월별 수입금액 (천달러)',None,
    lambda i,m: f'=SUMIFS({nR("I")},{nR("E")},$B{{r}},{nR("A")},"{m}")')
# $B{r} 치환 (행 참조가 필요한 블록)
for hr,r0,r1 in [MX_X,MX_M,MX_N,MX_NI]:
    for rr in range(r0,r1+1):
        for cc in range(3,3+NM):
            v=wm.cell(rr,cc).value
            if isinstance(v,str) and '{r}' in v: wm.cell(rr,cc).value=v.replace('{r}',str(rr))
wm.freeze_panes='C5'
def mx_row(block,idx): return block[1]+idx
def mx_col(k): return L(3+k)
print('분석4B 월별매트릭스 완료 :', MX_HS_EX, MX_SEC, MX_N)

WM="'분석4B_월별매트릭스'"
def mref(block,idx,k): return f'{WM}!{mx_col(k)}{mx_row(block,idx)}'
IT_IDX=[NEW_ITEMS.index('라.IT부품'), NEW_ITEMS.index('다.IT제품')]
HS85_IDX=HS_ORDER_M.index('85')

# ══════════════════════════════════════════════════════════════════
# 분석5_반도체vs비반도체  (월별 분해 중심)
# ══════════════════════════════════════════════════════════════════
w5=wb.create_sheet('분석5_반도체vs비반도체'); w5.sheet_view.showGridLines=False
widths(w5,[24]+[12]*NM)
title(w5,1,'⑤ 반도체 vs 비반도체 — 월별로 쪼개서 본다',1+NM)
note(w5,2,1,'HS 2단위 자료라 반도체(HS 8541·8542)를 단독으로 뽑을 수 없어 세 가지 대용 정의를 병기했습니다. ① 협의 = 신성질별 IT부품 ② 광의 = IT부품+IT제품 ③ HS85 기준.',1+NM)
note(w5,3,1,'★ 누계(1~7월 합산)가 아니라 매달 따로 계산했습니다. 어느 달에 무엇이 얼마나 달라졌는지는 [B]·[C]에서 월 단위로 확인하십시오.',1+NM,color=RED)
r=5
sec(w5,r,'[A] 월별 금액 (31개월, 천달러)',1+NM); r+=1
A5H=r; head(w5,r,1,'구분')
for k in range(NM): head(w5,r,2+k,MONTHS[k])
r+=1; A5=r
put(w5,r,1,'① 반도체(협의) IT부품',bold=True)
for k in range(NM): put(w5,r,2+k,f'={mref(MX_N,IT_IDX[0],k)}',NUM,sz=9)
R_IT1=r; r+=1
put(w5,r,1,'② 반도체(광의) IT부품+IT제품',bold=True)
for k in range(NM): put(w5,r,2+k,f'={mref(MX_N,IT_IDX[0],k)}+{mref(MX_N,IT_IDX[1],k)}',NUM,sz=9)
R_IT2=r; r+=1
put(w5,r,1,'③ HS85 전기기기',bold=True)
for k in range(NM): put(w5,r,2+k,f'={mref(MX_HS_EX,HS85_IDX,k)}',NUM,sz=9)
R_HS85=r; r+=1
put(w5,r,1,'전체 수출',bold=True)
for k in range(NM): put(w5,r,2+k,f"='분석2_월별추이'!E{D2+k}",NUM,sz=9)
R_TOT=r; r+=1
put(w5,r,1,'비반도체(광의 기준)',bold=True)
for k in range(NM): put(w5,r,2+k,f'={L(2+k)}{R_TOT}-{L(2+k)}{R_IT2}',NUM,sz=9)
R_NON=r; r+=1
put(w5,r,1,'반도체(광의) 비중',bold=True)
for k in range(NM): put(w5,r,2+k,f'={L(2+k)}{R_IT2}/{L(2+k)}{R_TOT}',PCT2,sz=9)
R_SH=r; r+=2
sec(w5,r,'[B] 월별 전년동월비 — 어느 달에 어느 쪽이 좋았나 (2025.01~2026.07)',1+NM); r+=1
B5H=r; head(w5,r,1,'구분')
for k in range(12,NM): head(w5,r,2+k-12,MONTHS[k])
r+=1; B5=r
def yoy_row(label,src,fmt=PCT,bold=True):
    global r
    put(w5,r,1,label,bold=bold)
    for k in range(12,NM):
        put(w5,r,2+k-12,f'={L(2+k)}{src}/{L(2+k-12)}{src}-1',fmt,sz=9)
    rr=r; r+=1; return rr
Y_IT1=yoy_row('① 반도체(협의) YoY',R_IT1)
Y_IT2=yoy_row('② 반도체(광의) YoY',R_IT2)
Y_HS85=yoy_row('③ HS85 YoY',R_HS85)
Y_NON=yoy_row('비반도체(광의) YoY',R_NON)
Y_TOT=yoy_row('전체 수출 YoY',R_TOT)
put(w5,r,1,'격차 (반도체−비반도체)',bold=True)
for k in range(12,NM): put(w5,r,2+k-12,f'=({L(2+k-12)}{Y_IT2}-{L(2+k-12)}{Y_NON})*100',PP,sz=9)
r+=1
put(w5,r,1,'그 달의 우위',bold=True)
for k in range(12,NM):
    cl=L(2+k-12)
    put(w5,r,2+k-12,f'=IF({cl}{Y_IT2}>{cl}{Y_NON},"반도체","비반도체")',align='center',sz=9)
R_WIN=r; r+=1
put(w5,r,1,'반도체가 만든 몫(기여도)',bold=True)
for k in range(12,NM):
    c1=L(2+k); c0=L(2+k-12)
    put(w5,r,2+k-12,f'=IFERROR(({c1}{R_IT2}-{c0}{R_IT2})/({c1}{R_TOT}-{c0}{R_TOT}),"-")',PCT2,sz=9)
R_CONT=r; r+=1
put(w5,r,1,'비반도체가 만든 몫',bold=True)
for k in range(12,NM):
    c1=L(2+k); c0=L(2+k-12)
    put(w5,r,2+k-12,f'=IFERROR(({c1}{R_NON}-{c0}{R_NON})/({c1}{R_TOT}-{c0}{R_TOT}),"-")',PCT2,sz=9)
r+=1
note(w5,r,1,'기여도는 그 달의 전체 수출 증감액을 100%로 놓았을 때의 몫입니다. 전체가 감소한 달에는 부호가 뒤집혀 보일 수 있으니(예 : 전체는 줄었는데 반도체는 늘어난 달) 위의 YoY와 함께 읽으십시오.',1+NM)
r+=2
sec(w5,r,'[C] 월별 증감액 분해 (천달러) — 전년 동월 대비',1+NM); r+=1
C5H=r; head(w5,r,1,'구분')
for k in range(12,NM): head(w5,r,2+k-12,MONTHS[k])
r+=1
for lab,src in [('전체 수출 증감액',R_TOT),('반도체(광의) 증감액',R_IT2),('비반도체 증감액',R_NON),('HS85 증감액',R_HS85)]:
    put(w5,r,1,lab,bold=True)
    for k in range(12,NM):
        put(w5,r,2+k-12,f'={L(2+k)}{src}-{L(2+k-12)}{src}',NUM,sz=9)
    r+=1
r+=1
sec(w5,r,'[D] 참고 : 연도별 1~7월 누계 (월별 분해의 합계 확인용)',1+NM); r+=1
D5H=r
for i,h in enumerate(['구분','2024','2025','2026','25 YoY','26 YoY','26 비중']): head(w5,r,i+1,h)
r+=1
D5_0=r
_d5=[('반도체(협의)',R_IT1),('반도체(광의)',R_IT2),('HS85',R_HS85),('비반도체(광의)',R_NON),('전체 수출',R_TOT)]
for i5,(lab,src) in enumerate(_d5):
    put(w5,r,1,lab,bold=(i5==4))
    for i,y in enumerate(YRS):
        k0=MONTHS.index(f'{y}.01'); cols='+'.join(f'{L(2+k0+j)}{src}' for j in range(7))
        put(w5,r,2+i,'='+cols,NUM,bold=(i5==4))
    put(w5,r,5,f'=C{r}/B{r}-1',PCT); put(w5,r,6,f'=D{r}/C{r}-1',PCT)
    put(w5,r,7,f'=D{r}/D${D5_0+4}',PCT2)
    r+=1
w5.freeze_panes='B7'
print('분석5(월별) 완료')

# ══════════════════════════════════════════════════════════════════
# 분석6_HS부문별 (월별 분해 중심)
# ══════════════════════════════════════════════════════════════════
w6=wb.create_sheet('분석6_HS부문별'); w6.sheet_view.showGridLines=False
widths(w6,[7,24,26]+[12]*20)
title(w6,1,'⑥ HS 부문(Section)별 — 21개 부문 전체를 월별로 쪼개서 본다',23)
note(w6,2,1,'HS 97개 류를 21개 부문으로 묶었습니다. 세 번째 열에 각 부문에 어떤 류가 들어가는지 명시했습니다. 정렬은 2026년 1~7월 수출 증감액 순(작성 시점 고정).',23)
note(w6,3,1,'★ [A]는 부문별 월별 전년동월비, [B]는 월별 증감액입니다. 누계로 뭉뚱그리지 않고 매달 따로 계산했으므로, 어느 부문이 몇 월에 꺾였는지·튀었는지 바로 보입니다.',23,color=RED)
SEC_ORDER=[]
for s,n,chs in SECTIONS:
    ch=[c for c in chs if c in HS]
    if ch: SEC_ORDER.append((s,n,ch))
def _sv(chs,y): return sum(HS[c][m][1] for c in chs for m in M17[str(y)] if m in HS[c])
SEC_ORDER.sort(key=lambda t:-(_sv(t[2],2026)-_sv(t[2],2025)))
SEC_IDX={s:[x[0] for x in SEC_ROWS].index(s) for s,n,ch in SEC_ORDER}
r=5
sec(w6,r,'[A] 부문별 월별 전년동월비 (2025.01~2026.07)',23); r+=1
A6H=r; head(w6,r,1,'부'); head(w6,r,2,'부문 명칭')
for k in range(12,NM): head(w6,r,3+k-12,MONTHS[k])
r+=1; A6=r
for s,n,ch in SEC_ORDER:
    idx=SEC_IDX[s]
    put(w6,r,1,s,align='center',bold=True,sz=9); put(w6,r,2,n,sz=9)
    for k in range(12,NM):
        cur=mref(MX_SEC,idx,k); pre=mref(MX_SEC,idx,k-12)
        put(w6,r,3+k-12,f'=IFERROR({cur}/{pre}-1,"-")',PCT,sz=9)
    r+=1
E6A=r-1; r+=1
sec(w6,r,'[B] 부문별 월별 증감액 (전년 동월 대비, 천달러)',23); r+=1
B6H=r; head(w6,r,1,'부'); head(w6,r,2,'부문 명칭')
for k in range(12,NM): head(w6,r,3+k-12,MONTHS[k])
r+=1; B6=r
for s,n,ch in SEC_ORDER:
    idx=SEC_IDX[s]
    put(w6,r,1,s,align='center',bold=True,sz=9); put(w6,r,2,n,sz=9)
    for k in range(12,NM):
        put(w6,r,3+k-12,f'={mref(MX_SEC,idx,k)}-{mref(MX_SEC,idx,k-12)}',NUM,sz=9)
    r+=1
E6B=r-1; r+=1
sec(w6,r,'[C] 부문별 월별 수출금액 (2026.01~07, 천달러) 및 구성 류',23); r+=1
C6H=r
for i,h in enumerate(['부','부문 명칭','포함 HS류']): head(w6,r,i+1,h)
for j,m in enumerate(M26L): head(w6,r,4+j,m)
head(w6,r,4+len(M26L),'26년 1~7월'); head(w6,r,5+len(M26L),'26 비중')
r+=1; C6=r
k26=[MONTHS.index(m) for m in M26L]
for s,n,ch in SEC_ORDER:
    idx=SEC_IDX[s]
    put(w6,r,1,s,align='center',bold=True,sz=9); put(w6,r,2,n,sz=9)
    put(w6,r,3,', '.join(ch),sz=8,wrap=True)
    for j,k in enumerate(k26): put(w6,r,4+j,f'={mref(MX_SEC,idx,k)}',NUM,sz=9)
    put(w6,r,4+len(M26L),f'=SUM({L(4)}{r}:{L(3+len(M26L))}{r})',NUM,bold=True)
    put(w6,r,5+len(M26L),f'=K{r}/${L(4+len(M26L))}${r+len(SEC_ORDER)-SEC_ORDER.index((s,n,ch))}'.replace('K'+str(r),f'{L(4+len(M26L))}{r}'),PCT2)
    r+=1
E6C=r-1
put(w6,r,2,'합계',bold=True)
for j in range(len(M26L)+1): put(w6,r,4+j,f'=SUM({L(4+j)}{C6}:{L(4+j)}{E6C})',NUM,bold=True)
SUM6C=r
for rr in range(C6,E6C+1):
    put(w6,rr,5+len(M26L),f'={L(4+len(M26L))}{rr}/${L(4+len(M26L))}${SUM6C}',PCT2)
r+=2
sec(w6,r,'[D] 부문별 연도 비교 (참고 — 월별 분해의 합계)',23); r+=1
D6H=r
for i,h in enumerate(['부','부문 명칭','수출 24(1~7)','수출 25(1~7)','수출 26(1~7)','25 YoY','26 YoY','기여도','수입 26','무역수지 26','성격 판정']): head(w6,r,i+1,h)
r+=1; D6=r
for s,n,ch in SEC_ORDER:
    put(w6,r,1,s,align='center',bold=True); put(w6,r,2,n)
    for i,y in enumerate(YRS): put(w6,r,3+i,'='+'+'.join(hs_y(c,y) for c in ch),NUM)
    put(w6,r,6,f'=IFERROR(D{r}/C{r}-1,"-")',PCT); put(w6,r,7,f'=IFERROR(E{r}/D{r}-1,"-")',PCT)
    put(w6,r,8,f'=(E{r}-D{r})/({_e26}-{_e25})',PCT2)
    put(w6,r,9,'='+'+'.join(hs_y(c,2026,'J') for c in ch),NUM)
    put(w6,r,10,f'=E{r}-I{r}',NUM)
    put(w6,r,11,f'=IF(H{r}>0.3,"전체 증가를 주도",IF(H{r}>0.03,"증가에 뚜렷이 기여",IF(G{r}<0,"감소 — 부진 부문",IF(H{r}>0,"소폭 기여","영향 미미"))))',sz=9)
    r+=1
E6D=r-1
w6.freeze_panes='C7'
print('분석6(월별) 완료')

# ══════════════════════════════════════════════════════════════════
# 분석7_HS류별상세 (월별 분해 중심)
# ══════════════════════════════════════════════════════════════════
w7=wb.create_sheet('분석7_HS류별상세'); w7.sheet_view.showGridLines=False
widths(w7,[6,7,24,32]+[11]*7+[12]*7+[13,13,10,10,12,12,11,13,13,44])
title(w7,1,'⑦ HS 류(2단위)별 전수 분석 — 97개 류를 월별로 쪼개서 본다',33)
note(w7,2,1,'원본에 있는 HS류를 하나도 빼지 않고 담았습니다. 「주요 구성 품목」은 해당 류에 실제로 들어가는 대표 품목이며, 「해석」 열은 월별 패턴까지 반영해 수식으로 자동 판정한 문장입니다.',33)
note(w7,3,1,'★ C~I열 = 2026년 각 월의 전년동월비, J~P열 = 각 월의 증감액. 누계가 아니라 월별로 나눠 놓았으므로 "어느 달에 무엇이 달라졌는지"를 류 단위로 볼 수 있습니다. Q열부터는 참고용 누계입니다.',33,color=RED)
r=5
H7=['부','HS','품목명(원본)','주요 구성 품목']+[f'{m} YoY' for m in M26L]+[f'{m} Δ' for m in M26L]+ \
   ['수출 25(1~7)','수출 26(1~7)','26 YoY','기여도','중량 26','단가 증감','성장 유형','수입 26','무역수지 26','해석']
for i,h in enumerate(H7): head(w7,r,i+1,h)
w7.row_dimensions[r].height=34
r+=1; R7=r
order7=sorted(HS,key=lambda c:-(sum(HS[c][m][1] for m in M17['2026'] if m in HS[c])-sum(HS[c][m][1] for m in M17['2025'] if m in HS[c])))
HSI={c:HS_ORDER_M.index(c) for c in HS_ORDER_M}
for code in order7:
    s,n=CH2SEC.get(code,('-','-')); idx=HSI[code]
    put(w7,r,1,s,align='center',sz=9); put(w7,r,2,code,align='center',bold=True).number_format='@'
    put(w7,r,3,f'=IFERROR(INDEX({hR("F")},MATCH($B{r},{hR("E")},0)),"")',sz=8,wrap=True)
    put(w7,r,4,DESC.get(code,''),sz=8,wrap=True)
    for j,m in enumerate(M26L):
        k=MONTHS.index(m)
        put(w7,r,5+j,f'=IFERROR({mref(MX_HS_EX,idx,k)}/{mref(MX_HS_EX,idx,k-12)}-1,"-")',PCT,sz=9)
        put(w7,r,12+j,f'={mref(MX_HS_EX,idx,k)}-{mref(MX_HS_EX,idx,k-12)}',NUM,sz=9)
    put(w7,r,19,'='+hs_y(code,2025),NUM)
    put(w7,r,20,'='+hs_y(code,2026),NUM)
    put(w7,r,21,f'=IFERROR(T{r}/S{r}-1,"-")',PCT)
    put(w7,r,22,f'=(T{r}-S{r})/({_e26}-{_e25})',PCT2)
    put(w7,r,23,'='+hs_y(code,2026,'G'),NUM1)
    put(w7,r,24,f'=IFERROR((T{r}/W{r})/(S{r}/'+hs_y(code,2025,'G')+f')-1,"-")',PCT)
    put(w7,r,25,f'=IF(OR(W{r}<100,{hs_y(code,2025,"G")}<100),"단가 불안정",'
                f'IF(T{r}-S{r}<0,"감소",IF(X{r}>0.05,"가격 주도",IF(X{r}<0.02,"물량 주도","혼합"))))',align='center',sz=9)
    put(w7,r,26,'='+hs_y(code,2026,'J'),NUM)
    put(w7,r,27,f'=T{r}-Z{r}',NUM)
    put(w7,r,28,
        f'=IF(T{r}=0,"자료 없음",'
        f'IF(U{r}>1,"1~7월 2배 이상 급증. ",IF(U{r}>0.5,"1~7월 +50% 초과 급증. ",IF(U{r}>0.15,"1~7월 뚜렷한 증가. ",'
        f'IF(U{r}>0.02,"1~7월 소폭 증가. ",IF(U{r}>-0.02,"1~7월 보합. ",IF(U{r}>-0.15,"1~7월 소폭 감소. ","1~7월 뚜렷한 감소. "))))))'
        f'&"월별로는 "&COUNTIF(E{r}:K{r},">0")&"개월 증가·"&COUNTIF(E{r}:K{r},"<0")&"개월 감소"'
        f'&IF(COUNTIF(E{r}:K{r},">0")=7," (7개월 내내 증가)",IF(COUNTIF(E{r}:K{r},"<0")=7," (7개월 내내 감소)",""))'
        f'&IFERROR(", 최대 "&TEXT(MAX(E{r}:K{r}),"0.0%")&"("&SUBSTITUTE(INDEX($E$5:$K$5,MATCH(MAX(E{r}:K{r}),E{r}:K{r},0))," YoY","")&")","")'
        f'&IFERROR(", 최소 "&TEXT(MIN(E{r}:K{r}),"0.0%")&"("&SUBSTITUTE(INDEX($E$5:$K$5,MATCH(MIN(E{r}:K{r}),E{r}:K{r},0))," YoY","")&")","")&". "'
        f'&IF(Y{r}="가격 주도","단가 상승이 주도. ",IF(Y{r}="물량 주도","물량 확대가 주도. ",IF(Y{r}="혼합","물량·단가 동반. ",IF(Y{r}="감소","감소 품목. ","단가 해석 유보. "))))'
        f'&IF(V{r}>0.05,"전체 증가의 5% 이상을 설명하는 대형 품목.",IF(V{r}>0.005,"전체 증가에 의미 있게 기여.",IF(T{r}/{_e26}<0.001,"규모가 작아 총량 영향은 미미.","총량 영향은 제한적.")))',
        sz=8,wrap=True)
    w7.row_dimensions[r].height=34
    r+=1
E7=r-1
put(w7,r,3,'합계(HS 전체)',bold=True)
for j in range(7): put(w7,r,12+j,f'=SUM({L(12+j)}{R7}:{L(12+j)}{E7})',NUM,bold=True)
for col in [19,20,23,26,27]: put(w7,r,col,f'=SUM({L(col)}{R7}:{L(col)}{E7})',NUM1 if col==23 else NUM,bold=True)
put(w7,r,21,f'=T{r}/S{r}-1',PCT,bold=True); put(w7,r,22,f'=(T{r}-S{r})/({_e26}-{_e25})',PCT2,bold=True)
SUM7=r; r+=2
sec(w7,r,'[요약 1] 성장 유형별 집계',33); r+=1
for i,h in enumerate(['유형','류 수','2026 수출 합','Δ금액 합','전체 증가분 대비']): head(w7,r,i+1,h)
r+=1; T7=r
for t in ['가격 주도','물량 주도','혼합','감소','단가 불안정']:
    put(w7,r,1,t,bold=True)
    put(w7,r,2,f'=COUNTIF($Y${R7}:$Y${E7},A{r})',NUM)
    put(w7,r,3,f'=SUMIF($Y${R7}:$Y${E7},A{r},$T${R7}:$T${E7})',NUM)
    put(w7,r,4,f'=SUMIF($Y${R7}:$Y${E7},A{r},$T${R7}:$T${E7})-SUMIF($Y${R7}:$Y${E7},A{r},$S${R7}:$S${E7})',NUM)
    put(w7,r,5,f'=D{r}/({_e26}-{_e25})',PCT2)
    r+=1
r+=1
sec(w7,r,'[요약 2] 월별 증감 품목 수 — 그 달에 몇 개 류가 늘고 줄었나',33); r+=1
for i,h in enumerate(['구분']+M26L): head(w7,r,i+1,h)
r+=1; D7=r
for lab,op in [('증가한 류 수','">0"'),('감소한 류 수','"<0"')]:
    put(w7,r,1,lab,bold=True)
    for j in range(7): put(w7,r,2+j,f'=COUNTIF({L(5+j)}${R7}:{L(5+j)}${E7},{op})',NUM)
    r+=1
put(w7,r,1,'증가 류 비율',bold=True)
for j in range(7): put(w7,r,2+j,f'={L(2+j)}{D7}/({L(2+j)}{D7}+{L(2+j)}{D7+1})',PCT2)
r+=1
put(w7,r,1,'증가액이 총변동액에서 차지하는 비중',bold=True)
for j in range(7):
    pos=f'SUMIF({L(12+j)}${R7}:{L(12+j)}${E7},">0")'
    neg=f'SUMIF({L(12+j)}${R7}:{L(12+j)}${E7},"<0")'
    put(w7,r,2+j,f'=IFERROR({pos}/({pos}-{neg}),"-")',PCT2)
r+=1
note(w7,r,1,'증가 류 비율(개수 기준)과 증가액 비중(금액 기준)을 함께 보십시오. 개수는 낮은데 금액 비중이 높으면 "큰 품목만 늘었다"는 뜻입니다. 증가액 비중 = 늘어난 류들의 증가액 합 ÷ (증가액 합 + 감소액 절대합).',33)
w7.freeze_panes='E6'
print('분석7(월별) 완료', R7, E7)
