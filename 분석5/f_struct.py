# -*- coding: utf-8 -*-
"""분석3_수출구조분해 — 반도체/비반도체 · 조업일수 보정 · 21개 HS 부문"""
SPAN3=2+NM
w3=newsheet('분석3_수출구조분해',[26,30]+[11]*NM)
title(w3,1,'③ 수출 구조 분해 — 반도체를 빼면 무엇이 남는가',SPAN3)
note(w3,2,1,'「2026년은 반도체가 좋았다」는 사실이지만 절반의 이야기입니다. 반도체를 제외한 나머지가 같은 기간 어떤 흐름이었는지, '
            '그리고 조업일수 차이를 걷어내면 그림이 어떻게 달라지는지를 이 시트에서 확인하십시오.',SPAN3)
note(w3,3,1,'반도체 = HS 8542(집적회로)·8473(모듈·부품)·8523(저장장치)·8541(개별소자). 장비(8486)와 검사(9030·9031)는 별도로 두어 「소자 호황이 장비로 번졌는지」를 따로 볼 수 있게 했습니다.',SPAN3)
r=5
sec(w3,r,'[A] 반도체 vs 비반도체 — 월별 분해',SPAN3); r+=1
A3H=r; head(w3,r,1,'구분'); head(w3,r,2,'설명')
for k in range(NM): head(w3,r,3+k,MONTHS[k])
r+=1; A3=r
SEMI_LIVE=[c for c in SEMI_CODES if c in CROW]
EQP=[c for c in ['8486','9030','9031'] if c in CROW]
ROWS3=[('국가 총수출','원본 「수출입 총괄」의 월별 수출금액',lambda k: TX(k),NUM),
       ('반도체 (4개 코드)','8542+8473+8523+8541',lambda k: csum(PEX,SEMI_LIVE,k),NUM),
       ('반도체 장비·검사','8486+9030+9031 — 소자와 시차가 있어 따로 봅니다',lambda k: csum(PEX,EQP,k),NUM),
       ('비반도체 (총수출−반도체)','반도체 4개 코드를 뺀 나머지 전부',lambda k: f'{TX(k)}-({csum(PEX,SEMI_LIVE,k)})',NUM)]
for lab,dsc,fn,fmt in ROWS3:
    put(w3,r,1,lab,bold=True,sz=9); put(w3,r,2,dsc,sz=8,wrap=True)
    for k in range(NM): put(w3,r,3+k,'='+fn(k),fmt,sz=8)
    r+=1
R_TOT,R_SEMI,R_EQP,R_NS=A3,A3+1,A3+2,A3+3
put(w3,r,1,'반도체 비중',bold=True,sz=9,fill=GOOD); put(w3,r,2,'국가 총수출에서 반도체 4개 코드가 차지하는 몫',sz=8,wrap=True,fill=GOOD)
for k in range(NM): put(w3,r,3+k,f'=IFERROR({L(3+k)}{R_SEMI}/{L(3+k)}{R_TOT},"-")',PCT,sz=8,fill=GOOD)
R_SH=r; r+=1
for lab,src,dsc in [('총수출 YoY',R_TOT,'명목 전년동월비'),
                    ('반도체 YoY',R_SEMI,'명목 전년동월비'),
                    ('비반도체 YoY',R_NS,'명목 전년동월비 — 이 줄이 이 시트의 핵심입니다')]:
    put(w3,r,1,lab,bold=True,sz=9,fill=(GOOD if '비반도체' in lab else None))
    put(w3,r,2,dsc,sz=8,wrap=True,fill=(GOOD if '비반도체' in lab else None))
    for k in range(NM):
        put(w3,r,3+k,(f'=IFERROR({L(3+k)}{src}/{L(3+k-12)}{src}-1,"-")' if k>=12 else ''),PCT,sz=8,
            fill=(GOOD if '비반도체' in lab else None))
    r+=1
Y_TOT,Y_SEMI,Y_NS=r-3,r-2,r-1
put(w3,r,1,'비반도체 YoY (조업일수 보정)',bold=True,sz=9,fill=GOOD)
put(w3,r,2,'일평균 기준. 2026.02는 조업일수가 19일로 전년 22일보다 3일 적어 명목으로는 마이너스가 나옵니다.',sz=8,wrap=True,fill=GOOD)
for k in range(NM):
    put(w3,r,3+k,(f'=IFERROR(({L(3+k)}{R_NS}/{TD(k)})/({L(3+k-12)}{R_NS}/{TD(k-12)})-1,"-")' if k>=12 else ''),PCT,sz=8,fill=GOOD)
Y_NSD=r; r+=1
put(w3,r,1,'조업일수',bold=True,sz=9); put(w3,r,2,'원본 「수출입 총괄」 조업일수',sz=8)
for k in range(NM): put(w3,r,3+k,'='+TD(k),NUM1,sz=8)
R_DAY=r; r+=1
put(w3,r,1,'조업일수 YoY',bold=True,sz=9); put(w3,r,2,'전년 동월 대비 조업일수 증감률 — 명목 YoY를 읽을 때의 보정치',sz=8,wrap=True)
for k in range(NM):
    put(w3,r,3+k,(f'=IFERROR({L(3+k)}{R_DAY}/{L(3+k-12)}{R_DAY}-1,"-")' if k>=12 else ''),PCT,sz=8)
R_DAYY=r; r+=1
put(w3,r,1,'반도체 성장 기여도',bold=True,sz=9); put(w3,r,2,'그 달 총수출 증감액 중 반도체가 만든 몫',sz=8,wrap=True)
for k in range(NM):
    put(w3,r,3+k,(f'=IFERROR(({L(3+k)}{R_SEMI}-{L(3+k-12)}{R_SEMI})/({L(3+k)}{R_TOT}-{L(3+k-12)}{R_TOT}),"-")' if k>=12 else ''),PCT,sz=8)
R_CTB=r; r+=1
put(w3,r,1,'판정',bold=True,sz=9,fill=SUB); put(w3,r,2,'비반도체가 그 달에 실제로 늘었는지 (조업일수 보정 기준)',sz=8,wrap=True,fill=SUB)
for k in range(NM):
    put(w3,r,3+k,(f'=IF({L(3+k)}{Y_NSD}="","-",IF({L(3+k)}{Y_NSD}>0.05,"양호",IF({L(3+k)}{Y_NSD}>0,"소폭 증가",IF({L(3+k)}{Y_NSD}>-0.05,"소폭 감소","부진"))))' if k>=12 else ''),
        None,sz=8,align='center',fill=SUB)
r+=2
sec(w3,r,'[B] 요약 — 누계와 최근 3개월',SPAN3); r+=1
for i,h in enumerate(['구분','설명','2024 1~7','2025 1~7','2026 1~7','25 YoY','26 YoY','최근 3M YoY','직전 3M YoY','가속도','조업보정 26 YoY']): head(w3,r,i+1,h)
r+=1; B3=r
for lab,src in [('국가 총수출',R_TOT),('반도체',R_SEMI),('반도체 장비·검사',R_EQP),('비반도체',R_NS)]:
    m=lambda ks: '+'.join(f'{L(3+k)}{src}' for k in ks)
    dm=lambda ks: '+'.join(f'({L(3+k)}{src}/{TD(k)})' for k in ks)
    put(w3,r,1,lab,bold=True,sz=10,fill=(GOOD if lab=='비반도체' else None))
    put(w3,r,2,{'국가 총수출':'원본 총괄 기준','반도체':'8542·8473·8523·8541','반도체 장비·검사':'8486·9030·9031',
                '비반도체':'총수출에서 반도체 4개 코드를 뺀 값'}[lab],sz=8,wrap=True)
    put(w3,r,3,'='+m(K24),NUM); put(w3,r,4,'='+m(K25),NUM); put(w3,r,5,'='+m(K26),NUM)
    put(w3,r,6,f'=IFERROR(D{r}/C{r}-1,"-")',PCT); put(w3,r,7,f'=IFERROR(E{r}/D{r}-1,"-")',PCT,bold=True)
    put(w3,r,8,f'=IFERROR(({m(KL3)})/({m(KL3P)})-1,"-")',PCT,bold=True)
    put(w3,r,9,f'=IFERROR(({m(KL3B)})/({m(KL3BP)})-1,"-")',PCT)
    put(w3,r,10,f'=IFERROR((H{r}-I{r})*100,"-")',PP0)
    put(w3,r,11,f'=IFERROR(({dm(K26)})/({dm(K25)})-1,"-")',PCT,bold=True)
    r+=1
E_B3=r-1; r+=1
note(w3,r,1,'※ 「비반도체」 행의 조업보정 26 YoY가 플러스라면, 2026년 수출 호조는 반도체만의 이야기가 아닙니다. '
            '단월로는 2026년 2월만 마이너스인데 그 달은 조업일수가 3일 적었습니다.',SPAN3,color=NAVY); r+=2

sec(w3,r,'[C] HS 21개 부문 월별 전년동월비 — 97개 류 전체 (4단위 미수집 영역 포함)',SPAN3); r+=1
note(w3,r,1,'4단위 상세(분석7)는 10개 류만 담고 있습니다. 나머지 영역까지 놓치지 않도록 HS 21개 부문 전체를 여기에 둡니다. '
            '농수산·섬유·비금속 등 4단위를 수집하지 않은 영역의 흐름은 이 표로 확인하십시오.',SPAN3); r+=1
C3H=r; head(w3,r,1,'부문'); head(w3,r,2,'포함 류')
for k in range(NM): head(w3,r,3+k,MONTHS[k])
r+=1; C3=r
_HRNG=f"{DH}!$E${H0}:$E${H1}"; _PRNG=f"{DH}!$A${H0}:$A${H1}"; _VRNG=f"{DH}!$H${H0}:$H${H1}"
SEC_ROWS=[]
for rn,sname,chs in SECTIONS:
    sname=f'{rn}. {sname}'
    put(w3,r,1,sname,bold=True,sz=9); put(w3,r,2,('·'.join(chs) if len(chs)<=8 else chs[0]+'~'+chs[-1]+f' ({len(chs)}개)'),sz=8,wrap=True)
    SEC_ROWS.append((r,sname,chs))
    r+=1
E_C3=r-1
# 값행(월별 금액)을 뒤쪽에 숨겨 두고 YoY만 표시
r+=1
sec(w3,r,'[C-2] 부문별 월별 수출금액 (위 표의 계산 근거)',SPAN3); r+=1
D3H=r; head(w3,r,1,'부문'); head(w3,r,2,'포함 류')
for k in range(NM): head(w3,r,3+k,MONTHS[k])
r+=1; D3=r
for i,(rr,sname,chs) in enumerate(SEC_ROWS):
    put(w3,r,1,sname,bold=True,sz=9); put(w3,r,2,('·'.join(chs) if len(chs)<=8 else chs[0]+'~'+chs[-1]),sz=8,wrap=True)
    for k in range(NM):
        terms='+'.join(f'SUMIFS({_VRNG},{_PRNG},"{MONTHS[k]}",{_HRNG},"{c}")' for c in chs)
        put(w3,r,3+k,'='+terms,NUM,sz=8)
    r+=1
E_D3=r-1
for i,(rr,sname,chs) in enumerate(SEC_ROWS):
    src=D3+i
    for k in range(NM):
        put(w3,rr,3+k,(f'=IFERROR({L(3+k)}{src}/{L(3+k-12)}{src}-1,"-")' if k>=12 else ''),PCT,sz=8)
w3.freeze_panes='C7'
print('분석3 완료',A3,B3,C3,D3)
