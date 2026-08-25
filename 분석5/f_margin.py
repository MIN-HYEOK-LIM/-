# -*- coding: utf-8 -*-
"""분석8_마진·단가 — 수출단가와 투입 원가의 스프레드 (전 산업)"""
SPAN8=13+NM
w8=newsheet('분석8_마진·단가',[30,40,14,12,12,12,10,10,12,12,10,44,10]+[11]*NM)
title(w8,1,'⑧ 마진·단가 프록시 — 금액이 아니라 남는 것을 봅니다',SPAN8)
note(w8,2,1,'수출 톤당 단가 ÷ 수입 톤당 단가로 산업별 스프레드를 근사합니다. 절대 수준이 아니라 방향과 변화폭을 보십시오. '
            '수출금액이 늘어도 스프레드가 좁아지고 있으면 물량으로 버티는 것이고, 그 반대면 가격으로 버는 것입니다.',SPAN8)
note(w8,3,1,'※ 톤당 단가는 대당·개당 ASP가 아닙니다. 품목 믹스가 바뀌면 톤당 단가도 함께 움직이므로 방향 지표로만 쓰십시오.',SPAN8,color=RED)
# (라벨, 수출코드들, 수입코드들 or None, 설명, 형식)
MARGIN=[
 ('정유 : 석유제품 ÷ 원유',['2710'],['2709'],'정제마진 프록시. 수출하는 석유제품 단가와 사올 때 원유 단가의 비율입니다.','0.000'),
 ('석유화학 : 합성수지 ÷ 석유제품',['3901','3902','3903'],['2710'],'나프타 투입 대비 수지 판가. 석화 스프레드 프록시입니다.','0.000'),
 ('철강 : 도금강판 ÷ 철스크랩',['7210'],['7204'],'롤마진 프록시. 전기로 원가 대비 고부가 판재 판가.','0.000'),
 ('철강 : 스테인리스 ÷ 철스크랩',['7219'],['7204'],'STS 스프레드. 니켈 가격 영향이 섞여 있습니다.','0.000'),
 ('이차전지 : 축전지 수출 ÷ 수입',['8507'],['8507'],'같은 코드의 수출단가와 수입단가 비율. 국내 생산분의 상대적 부가가치를 봅니다.','0.000'),
 ('자동차 : 승용차 수출 ÷ 수입',['8703'],['8703'],'1을 넘으면 수입차보다 비싼 차를 팔고 있다는 뜻입니다(톤당 기준).','0.000'),
 ('기계 : 건설기계 ÷ 철강 판재',['8429'],['7208'],'투입 강재 대비 기계 판가. 기계업 마진 방향의 근사치.','0.000'),
 ('반도체 : 집적회로 단가',['8542'],None,'8542 톤당 단가 — 메모리 ASP 방향 프록시.',UNIT),
 ('반도체 : 모듈 단가',['8473'],None,'8473 톤당 단가 — 모듈 ASP 방향 프록시.',UNIT),
 ('디스플레이 : 8524 단가',['8524'],None,'패널 단가 방향.',UNIT),
 ('조선 : 상선 단가',['8901'],None,'톤당 선가 프록시. 고부가 선종 비중이 오르면 함께 오릅니다.',UNIT),
 ('의료기기 : 9018 단가',['9018'],None,'의료기기 톤당 단가 — 고부가화 여부.',UNIT),
 ('합성수지 : 3901 단가',['3901'],None,'PE 판가 방향.',UNIT),
 ('건설기계 : 8429 단가',['8429'],None,'장비 대당 가격 방향의 근사치.',UNIT),
]
r=5
sec(w8,r,'[A] 월별 추이',SPAN8); r+=1
A8H=r
for i,h in enumerate(['지표','설명','구성','','','','','','','','','','']): head(w8,r,i+1,h)
for k in range(NM): head(w8,r,14+k,MONTHS[k],fill=HDR3)
r+=1; A8=r
for lab,ex,im,dsc,fmt in MARGIN:
    ex=[c for c in ex if c in CROW]; im=[c for c in im if c in CROW] if im else None
    put(w8,r,1,lab,bold=True,sz=9); put(w8,r,2,dsc,sz=8,wrap=True)
    put(w8,r,3,('·'.join(ex)+' / '+'·'.join(im)) if im else '·'.join(ex),align='center',sz=8)
    for k in range(NM):
        exu=f'(({csum(PEX,ex,k)})/({csum(PWT,ex,k)}))'
        if im:
            imu=f'(({csum(PIM,im,k)})/({csum(PWI,im,k)}))'
            put(w8,r,14+k,f'=IFERROR({exu}/{imu},"-")',fmt,sz=8)
        else:
            put(w8,r,14+k,f'=IFERROR({exu},"-")',fmt,sz=8)
    r+=1
E8A=r-1; r+=1
sec(w8,r,'[B] 구간 비교와 방향',SPAN8); r+=1
for i,h in enumerate(['지표','설명','구성','2024 1~7','2025 1~7','2026 1~7','25 변화','26 변화','최근 3개월','직전 3개월','방향','해석','']):
    head(w8,r,i+1,h)
r+=1; B8=r
for i,(lab,ex,im,dsc,fmt) in enumerate(MARGIN):
    a=A8+i
    avg=lambda ks: 'AVERAGE('+','.join(f'{L(14+k)}{a}' for k in ks)+')'
    put(w8,r,1,lab,bold=True,sz=9); put(w8,r,2,dsc,sz=8,wrap=True)
    put(w8,r,3,f'{"스프레드" if im else "단가"}',align='center',sz=8)
    put(w8,r,4,f'=IFERROR({avg(K24)},"-")',fmt); put(w8,r,5,f'=IFERROR({avg(K25)},"-")',fmt)
    put(w8,r,6,f'=IFERROR({avg(K26)},"-")',fmt)
    put(w8,r,7,f'=IFERROR(E{r}/D{r}-1,"-")',PCT); put(w8,r,8,f'=IFERROR(F{r}/E{r}-1,"-")',PCT,bold=True)
    put(w8,r,9,f'=IFERROR({avg(KL3)},"-")',fmt,bold=True); put(w8,r,10,f'=IFERROR({avg(KL3B)},"-")',fmt)
    put(w8,r,11,f'=IF(OR(I{r}="-",J{r}="-"),"-",IF(I{r}>J{r}*1.02,"개선",IF(I{r}<J{r}*0.98,"악화","보합")))',align='center',bold=True)
    put(w8,r,12,f'=IF(K{r}="-","비교할 값이 부족합니다.",'
                f'IF(K{r}="개선","최근 3개월이 직전 3개월보다 확대 — 마진 방향은 우호적입니다.",'
                f'IF(K{r}="악화","최근 3개월이 직전 3개월보다 축소 — 금액이 늘어도 마진은 압박받고 있습니다.",'
                f'"뚜렷한 방향성이 없습니다.")))'
                f'&IF(H{r}="-",""," 연간으로는 "&TEXT(H{r},"+0.0%")&".")',sz=8,wrap=True)
    r+=1
E8B=r-1
w8.freeze_panes='D7'
print('분석8 완료',A8,B8)
