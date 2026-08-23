const fs=require('fs');
const D=require('docx');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,
       BorderStyle,ShadingType,HeadingLevel,PageOrientation,ImageRun,convertInchesToTwip}=D;
const F='맑은 고딕';
const W=10440;
const NAVY='1F3864', GRAY='595959', RED='C00000', LIGHT='DCE6F1', BOX='F2F5FA';

const t=(text,o={})=>new TextRun({text,font:F,size:o.sz||19,bold:!!o.b,color:o.c||'000000',italics:!!o.i});
const p=(text,o={})=>new Paragraph({
  alignment:o.al||AlignmentType.JUSTIFIED,
  spacing:{before:o.before===undefined?40:o.before,after:o.after===undefined?40:o.after,line:o.line||250},
  indent:o.ind?{left:o.ind}:undefined,
  border:o.border,
  children:Array.isArray(text)?text:[t(text,o)]});
const h1=(text)=>new Paragraph({spacing:{before:180,after:80},
  border:{bottom:{style:BorderStyle.SINGLE,size:8,color:NAVY,space:2}},
  children:[t(text,{sz:22,b:true,c:NAVY})]});
const bullet=(text,o={})=>new Paragraph({spacing:{before:20,after:20,line:250},indent:{left:220,hanging:150},
  children:[t('· ',{sz:19,b:true,c:NAVY}),...(Array.isArray(text)?text:[t(text,o)])]});

function cell(txt,o={}){
  const runs=Array.isArray(txt)?txt:[t(txt,{sz:o.sz||16,b:o.b,c:o.c})];
  return new TableCell({
    width:{size:o.w,type:WidthType.DXA},
    shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,
    margins:{top:40,bottom:40,left:70,right:70},
    verticalAlign:'center',
    children:[new Paragraph({alignment:o.al||AlignmentType.LEFT,spacing:{before:0,after:0,line:230},children:runs})]});
}
function table(cols,rows,opt={}){
  const head=new TableRow({tableHeader:true,children:cols.map((c,i)=>
    cell(c.h,{w:c.w,b:true,c:'FFFFFF',fill:NAVY,al:c.al||AlignmentType.CENTER,sz:opt.hsz||15}))});
  const body=rows.map((r,ri)=>new TableRow({children:r.map((v,i)=>{
    const isTot=opt.totalRows&&opt.totalRows.includes(ri);
    return cell(v,{w:cols[i].w,al:i===0?(cols[i].al||AlignmentType.LEFT):AlignmentType.RIGHT,
      b:isTot,fill:isTot?LIGHT:(ri%2?'F7F9FC':undefined),sz:opt.sz||16,
      c:(typeof v==='string'&&v.startsWith('−'))?RED:undefined});})}));
  return new Table({columnWidths:cols.map(c=>c.w),width:{size:W,type:WidthType.DXA},
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:'AAB7CF'},bottom:{style:BorderStyle.SINGLE,size:4,color:'AAB7CF'},
             left:{style:BorderStyle.SINGLE,size:4,color:'AAB7CF'},right:{style:BorderStyle.SINGLE,size:4,color:'AAB7CF'},
             insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:'D0D8E8'},insideVertical:{style:BorderStyle.SINGLE,size:2,color:'D0D8E8'}},
    rows:[head,...body]});
}
const cap=(s)=>new Paragraph({spacing:{before:30,after:110},children:[t(s,{sz:14,i:true,c:GRAY})]});
const gap=(n)=>new Paragraph({spacing:{before:0,after:0},children:[t('',{sz:n||10})]});

// ── 본문 ────────────────────────────────────────────────────────────
const children=[];
children.push(new Paragraph({spacing:{after:20},children:[t('2026년 1~7월 수출 심층 분석',{sz:30,b:true,c:NAVY})]}));
children.push(new Paragraph({spacing:{after:60},children:[t('무엇이 늘었나 · 왜 늘었나 · 무엇이 위험한가',{sz:20,c:GRAY})]}));
children.push(new Paragraph({spacing:{after:120},
  border:{bottom:{style:BorderStyle.SINGLE,size:12,color:NAVY,space:4}},
  children:[t('분석 대상 : 관세청 수출입무역통계 5종(총괄·품목별·성질별·신성질별·국가별), 2025.01~2026.07, 수리일 기준 · 비교 단위 : 2026.1~7 vs 2025.1~7 · 금액 단위 : 억 달러(원자료 천달러를 환산)',{sz:14,c:GRAY})]}));

children.push(new Table({columnWidths:[W],width:{size:W,type:WidthType.DXA},
  borders:{top:{style:BorderStyle.SINGLE,size:12,color:NAVY},bottom:{style:BorderStyle.SINGLE,size:12,color:NAVY},
           left:{style:BorderStyle.SINGLE,size:12,color:NAVY},right:{style:BorderStyle.SINGLE,size:12,color:NAVY}},
  rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},
    shading:{type:ShadingType.CLEAR,fill:BOX,color:'auto'},margins:{top:100,bottom:100,left:150,right:150},
    children:[
      new Paragraph({spacing:{after:50},children:[t('핵심 결론',{sz:19,b:true,c:NAVY})]}),
      bullet([t('수출은 +50.5% 늘었지만 중량은 −4.2% 줄었다. 금액 증가분 1,996.7억 달러의 ',{sz:17}),t('113%가 가격효과',{sz:17,b:true}),t('이고 물량효과는 −8.4%다. 더 많이 판 것이 아니라 같은 물량을 비싸게 판 결과다.',{sz:17})]),
      bullet([t('구성 변화(비싼 품목으로 이동)는 +3.8%p뿐이고 ',{sz:17}),t('순수 가격 상승이 +51.4%',{sz:17,b:true}),t('(피셔 가격지수 1.514, 물량지수 0.994). 수출 고도화가 아니라 가격 사이클이다.',{sz:17})]),
      bullet([t('HS85(전기기기·반도체)는 물량 −0.7%에 단가 +110.0%로 증가분의 ',{sz:17}),t('67.0%',{sz:17,b:true}),t('를 만들었다. HS84까지 더하면 85%. 반면 HS87(차량)은 −1.0%로 역성장했다.',{sz:17})]),
      bullet([t('상위 20개국 중 13개국이 HS85와 상관계수 0.8 이상이며 이들이 상위 20개국 수출의 ',{sz:17}),t('90%',{sz:17,b:true}),t('를 차지한다. ',{sz:17}),t('지역 분산이 품목 리스크를 상쇄하지 못한다',{sz:17,b:true}),t('.',{sz:17})]),
      bullet([t('IT 수출금액이 20% 줄면 수출 −9.8%, 무역수지 −34.7%. 흑자가 0이 되는 임계점은 ',{sz:17}),t('IT −57.7%',{sz:17,b:true}),t('. 지금의 흑자 1,677.6억 달러는 반도체 가격의 함수다.',{sz:17})]),
    ]})]})]}));
children.push(gap(120));

children.push(h1('1. 핵심 지표'));
children.push(table(
  [{h:'지표',w:2100},{h:'2025.1~7',w:1500},{h:'2026.1~7',w:1500},{h:'증감',w:1500},{h:'증감률',w:1140},{h:'메모',w:2700,al:AlignmentType.CENTER}],
  [['수출 금액','3,953.9억$','5,950.6억$','+1,996.7억$','+50.5%','2025년 연간의 83.9%를 7개월에 달성'],
   ['수출 중량','1억1,474만t','1억987만t','−487만t','−4.2%','물량은 오히려 감소'],
   ['수출 단가','3.446','5.416','+1.970','+57.2%','천달러/톤'],
   ['일평균 수출','25.4억$','38.0억$','+12.6억$','+49.5%','조업일수는 +0.6%(155.5→156.5일)'],
   ['건당 수출금액','50.7천$','75.1천$','+24.4천$','+48.3%','건수는 +1.5%에 그침'],
   ['수입 금액','3,614.6억$','4,273.0억$','+658.5억$','+18.2%','수입 단가 +16.8%'],
   ['무역수지','339.3억$','1,677.6억$','+1,338.3억$','+394.4%','증가폭의 대부분이 수출 단가에서 발생']],
  {totalRows:[6]}));
children.push(cap('출처 : 분석2_월별추이, 분석8_물량·가격분해. 원자료 천달러를 억 달러로 환산(1억 달러 = 100,000천달러).'));

children.push(h1('2. 수출 증가의 분해 — 물량인가 가격인가'));
children.push(p('금액은 물량 × 단가이므로 금액 증감은 물량효과·가격효과·교차효과로 정확히 분해된다. 세 값의 합은 실제 증감액과 일치하며(검증 완료), 결과는 한쪽으로 크게 기운다.'));
children.push(table(
  [{h:'효과',w:2600},{h:'금액',w:1700},{h:'증가분 대비',w:1400},{h:'해석',w:4740,al:AlignmentType.CENTER}],
  [['물량효과 (Q₁−Q₀)×P₀','−167.8억$','−8.4%','물량은 오히려 줄어 증가를 깎아먹었다'],
   ['가격효과 (P₁−P₀)×Q₀','+2,260.5억$','+113.2%','증가의 전부가 여기서 나왔다'],
   ['교차효과','−95.9억$','−4.8%','물량 감소 × 단가 상승의 상쇄분'],
   ['합계 = 실제 증감','+1,996.7억$','100.0%','분해 항등식 검증 OK']],
  {totalRows:[3]}));
children.push(gap(60));
children.push(p([t('구성효과 검증. ',{b:true}),t('단가 상승분이 "비싼 품목으로 갈아탄 결과"인지 확인하기 위해 신성질별 중분류 14개로 지수를 계산했다. 피셔 물량지수 0.994(−0.6%), 피셔 가격지수 1.514(+51.4%), 금액지수 1.505로 항등식이 성립한다. 중량 단순합 기준 단가 상승(+57.2%)에서 순수 가격효과(+51.4%)를 제외한 구성(믹스)효과는 +3.8%p에 불과하다. 즉 ')
  ,t('품목 구성 변화가 아니라 팔던 품목의 값이 오른 것',{b:true}),t('이다. 참고로 톤 단위 중량 단순합은 저가·중량물(광산물 등)에 지배되어 −4.2%로 나오지만, 금액 가중 물량지수로 보면 −0.6%로 사실상 보합이다. 물량이 급감했다는 해석은 과장이다.')]));

children.push(h1('3. 품목 — 어디가 값으로 올랐고 어디가 물량으로 늘었나'));
children.push(table(
  [{h:'HS',w:600},{h:'품목',w:2000},{h:'Δ금액',w:1300},{h:'금액',w:900},{h:'물량',w:900},{h:'단가',w:900},{h:'기여도',w:900},{h:'유형',w:2940,al:AlignmentType.CENTER}],
  [['85','전기기기·반도체','+1,337.4억$','+108.6%','−0.7%','+110.0%','67.0%','가격 주도 — 메모리 가격'],
   ['84','기계류(컴퓨터·SSD 포함)','+362.0억$','+75.5%','+0.9%','+74.0%','18.1%','가격 주도 — 저장장치 파급 추정'],
   ['27','광물성 연료','+92.8억$','+34.3%','−8.5%','+46.7%','4.6%','가격 주도 — 유가·정제마진'],
   ['71','귀금속·보석','+60.5억$','+154.9%','−19.2%','+215.6%','3.0%','가격 주도 — 금 가격'],
   ['89','선박','+35.9억$','+23.7%','−0.8%','+24.8%','1.8%','고선가 수주분 인도'],
   ['33','화장품·정유','+17.8억$','+28.3%','+20.4%','+6.6%','0.9%','물량 주도 — 실수요 확대'],
   ['87','차량','−5.5억$','−1.0%','−2.2%','+1.2%','−0.3%','감소 — 상위 10대 중 유일']],
  {sz:15}));
children.push(cap('출처 : 분석9_품목단가분해(HS). HS 2단위 97개를 유형 분류하면 가격 주도 32개, 물량 주도 11개, 혼합 12개, 감소 33개, 단가 불안정 9개다.'));
children.push(p([t('두 가지 반론 검증. ',{b:true}),t('첫째, "반도체만 좋다"는 통념은 절반만 맞다. 금액가중 확산지수는 2026.06 99.0%, 07월 98.7%로 ')
  ,t('금액이 큰 품목은 사실상 전부 증가',{b:true}),t('했다. 다만 개수 기준으로는 97개 중 33개가 감소해 온기가 고르지는 않다. 둘째, IT를 제외한 수출도 2,738억 → 3,043억 달러로 +11.1% 늘었고, 2026년 6·7월 전년동월비는 +17.7%, +17.8%로 오히려 가속 중이다. 반도체 외 수요도 살아 있다는 신호이므로, IT 사이클 반전 시 완충 여지를 판단할 때 이 지표를 함께 봐야 한다.')]));

children.push(h1('4. 국가 — 지역 분산이 품목 리스크를 대신하지 못한다'));
children.push(p('각국 월별 수출과 HS85류 월별 수출의 19개월 상관계수를 계산했다. 값이 1에 가까울수록 그 나라 수출이 반도체 사이클과 함께 움직였다는 뜻이다.'));
children.push(table(
  [{h:'국가',w:1150},{h:'Δ금액',w:1200},{h:'증가율',w:900},{h:'기여도',w:900},{h:'HS85 동조성',w:1150},{h:'건당금액 변화',w:1250},{h:'건수',w:900},{h:'비고',w:2990,al:AlignmentType.CENTER}],
  [['중국','+496.7억$','+69.4%','24.9%','0.988','+58.8%','+6.7%','2위→1위, 수지 −82.8억→+188.5억$ 흑자 전환'],
   ['미국','+384.8억$','+53.1%','19.3%','0.951','+72.0%','−11.0%','건수 감소·고가화, 흑자 301.8→584.6억$'],
   ['홍콩','+277.8억$','+161.9%','13.9%','0.949','+126.6%','+15.6%','중계·환적 성격, 월변동성 49.7%로 최대'],
   ['베트남','+197.0억$','+56.8%','9.9%','0.972','+55.1%','+1.1%','후공정 기지'],
   ['대만','+159.7억$','+63.6%','8.0%','0.873','+20.4%','+35.9%','건수 급증형'],
   ['독일','+11.0억$','+9.2%','0.6%','0.446','−11.8%','+23.8%','비동조 — 반도체와 무관하게 완만'],
   ['폴란드','−0.7억$','−0.6%','−0.0%','0.229','−3.5%','+3.0%','비동조 — 감소']],
  {sz:15}));
children.push(cap('출처 : 분석5_국가별, 분석12_국가심층. 기여도는 전체 수출 증가분 대비 비중.'));
children.push(p([t('구조적 함의. ',{b:true}),t('상위 20개국을 동조성 구간으로 묶으면 0.90 이상 6개국이 이들 수출의 67.1%, 0.80 이상 13개국이 90.1%를 차지한다. 반대로 동조성이 낮은 국가(독일 0.446, 튀르키예 0.395, 폴란드 0.229)는 규모도 작고 증가율도 한 자릿수다. ')
  ,t('결과적으로 수출처를 늘려도 대부분이 같은 반도체 사이클 위에 있어, 지역 다변화로는 품목 집중 리스크가 거의 줄지 않는다.',{b:true}),t(' 실질적인 분산은 "다른 나라"가 아니라 "다른 품목"에서만 나온다.')]));

children.push(h1('5. 집중도 · 모멘텀 — 좁아졌고, 아직 빠르다'));
children.push(table(
  [{h:'지표',w:2400},{h:'2025.01',w:1300},{h:'2026.07',w:1300},{h:'최고치',w:1500},{h:'해석',w:3940,al:AlignmentType.CENTER}],
  [['품목 HHI','1,334','2,444','2,458(26.06)','집중도 83% 상승'],
   ['유효 품목수','7.5개','4.1개','—','실질적으로 4개 품목이 수출을 지탱'],
   ['국가 HHI / 유효 국가수','896 / 11.2','1,028 / 9.7','—','국가도 좁아졌으나 품목만큼은 아님'],
   ['IT 비중','29.2%','51.3%','53.7%(26.06)','수출의 절반이 IT부품·IT제품'],
   ['중화권(중·홍·대) 비중','27.1%','34.4%','36.7%(26.06)','+7.3%p'],
   ['전년동월비','—','+63.0%','+70.4%(26.06)','3개월 이동평균 +62.2%로 추세 유지'],
   ['일평균 전월비(26.07)','—','−9.0%','—','월합 −2.9%보다 실제 둔화가 크다']],
  {sz:15}));
children.push(cap('출처 : 분석10_집중도·확산도, 분석11_모멘텀·계절성. HHI = Σ(점유율²)×10,000, 유효 품목수 = 10,000÷HHI.'));
children.push(p([t('캐리오버(이월) 효과. ',{b:true}),t('2026년 1~7월 누계만으로 이미 2025년 연간 수출의 83.9%에 도달했다. 8~12월이 최근 3개월 평균 대비 25% 급락해도 연간으로는 +34.7%, 40% 급락해도 +24.6% 증가가 유지된다. ')
  ,t('연간 증가율이 높게 나오는 것과 지금 모멘텀이 살아 있는 것은 별개',{b:true}),t('이므로, 하반기 판단은 연간 증가율이 아니라 3개월 이동평균 전년동월비와 조업일수 보정 일평균으로 해야 한다.')]));

children.push(h1('6. 원인 분석 — 왜 이런 숫자가 나왔는가'));
children.push(p('아래 「관측 사실」은 데이터로 확인된 것이고, 「유력 원인」은 데이터 밖 환경에 대한 해석이다. 확정 전 검증 경로를 함께 적었다.'));
children.push(table(
  [{h:'관측 사실(데이터)',w:2500},{h:'유력 원인(해석)',w:4300},{h:'검증 방법',w:3000},{h:'신뢰도',w:640,al:AlignmentType.CENTER}],
  [['HS85 물량 −0.7%, 단가 +110.0%\nIT부품 단가 +141.3%','AI 데이터센터 투자 확대에 따른 HBM·서버 DRAM 수요 급증과, 범용 DRAM 캐파의 HBM 전환에 따른 공급 부족이 가격을 끌어올린 것으로 판단','DRAM·NAND 고정거래가격, 반도체 수출 물량·단가 지수, 메모리 업체 분기 ASP','높음'],
   ['HS84 단가 +74.0%인데 성질별 「기계류·정밀기기」 단가는 −0.7%','HS84에 포함된 컴퓨터·SSD(8471)로 메모리 가격이 파급된 것으로 추정. 순수 일반기계는 정체','HS 4단위(8471·8473·8486) 재조회로 84류 내부 분리','중간'],
   ['HS71 단가 +215.6%, 물량 −19.2%','국제 금 가격 급등. 물량이 20% 줄었는데 금액이 155% 늘어난 패턴은 가격 외 설명이 어려움','금 현물가격, HS 7108 세부 실적','중간'],
   ['미국 금액 +53.1%, 건수 −11.0%, 건당금액 +72.0%','소액·다건 화물이 줄고 고가품 중심으로 재편. 관세 환경 변화와 소액 화물 취급 축소가 배경일 가능성','대미 수출의 HS별·금액구간별 분해, 관세 대상/예외 품목 구분','중간'],
   ['HS87 −1.0%, 내구소비재 −7.2%','주요 시장 관세 부담과 현지생산 확대. 반도체 호황이 자동차로 전이되지 않음','대미·대EU 자동차 단가/물량, 현지공장 생산 실적','중간'],
   ['홍콩 +161.9%, 변동성 49.7%','최종 소비지가 아닌 반도체 중계·환적 경로. 실수요와 재고 이동이 섞여 있음','홍콩 경유 물량의 최종 목적지, 중국 직수출과 합산 추이','중간']],
  {sz:14,hsz:14}));
children.push(cap('출처 : 분석14_원인분석. 2026년 6~7월의 개별 사건(가격 협상·관세 조치·대형 인도 건 등)은 확인하지 않았으므로 최근 2개월 해석은 잠정이다.'));

children.push(h1('7. 리스크 — 흑자의 57.7%가 반도체 가격에 걸려 있다'));
children.push(table(
  [{h:'시나리오 / 임계점',w:3400},{h:'수출',w:1400},{h:'수출 증감',w:1200},{h:'무역수지',w:1400},{h:'수지 증감',w:1200},{h:'전년동기비',w:1840}],
  [['현재(2026.1~7 실적)','5,950.6억$','—','1,677.6억$','—','+50.5%'],
   ['IT 수출금액 −10%','5,659.8억$','−4.9%','1,386.8억$','−17.3%','+43.1%'],
   ['IT 수출금액 −20%','5,369.0억$','−9.8%','1,096.0억$','−34.7%','+35.8%'],
   ['IT 수출금액 −30%','5,078.2억$','−14.7%','805.2억$','−52.0%','+28.4%'],
   ['중화권 수요 −20%','5,536.2억$','−7.0%','1,263.2억$','−24.7%','+40.0%'],
   ['임계점 : 흑자 소멸','—','IT −57.7%','0억$','−100%','+21.3%'],
   ['임계점 : 수출 증가율 0%','3,953.9억$','IT −68.7%','—','—','0.0%']],
  {sz:15,totalRows:[5,6]}));
children.push(cap('출처 : 분석13_민감도. 단가 충격은 물량 불변 가정의 단순 시뮬레이션이며 수요·환율의 2차 효과는 미반영. 엑셀의 노란색 셀에 자체 전망치를 넣으면 즉시 재계산된다.'));

children.push(h1('8. 실무 권고'));
children.push(bullet([t('매월 볼 지표 5개를 고정하라. ',{b:true}),t('① HS85 전년동월비(방향) ② 수출 단가(흑자의 지속성) ③ 조업일수 보정 일평균(달력 착시 제거) ④ IT 제외 3개월이평 전년동월비(확산 여부) ⑤ 상위 5개국 비중(집중도). 다섯 개 모두 엑셀에서 값만 갱신하면 자동 계산된다.')]));
children.push(bullet([t('예산·전망은 물량 기준으로 이중 작성하라. ',{b:true}),t('금액 기준 +50.5%는 단가가 만든 숫자다. 물량 기준(피셔 물량지수 −0.6%)으로 한 벌 더 만들어 두면, 가격이 되돌아설 때 계획을 처음부터 다시 짜지 않아도 된다.')]));
children.push(bullet([t('흑자를 항상 소득으로 보지 마라. ',{b:true}),t('IT 수출금액 −20%만으로 흑자가 34.7% 줄어든다. 환율·재정·투자 판단에서 1,677억 달러를 기준선으로 삼으면 위험하다. 보수 기준선은 IT −20~30% 시나리오다.')]));
children.push(bullet([t('분산은 국가가 아니라 품목에서 찾아라. ',{b:true}),t('상위 20개국 중 13개국이 반도체와 동조하고 이들이 90%를 차지한다. 신규 시장 개척은 동조성이 낮은 품목(화장품·정유 HS33은 물량 주도 +20.4%)과 묶어야 실제 분산 효과가 난다.')]));
children.push(bullet([t('계약 조건에 단가 연동을 점검하라. ',{b:true}),t('단가가 1년 새 두 배가 된 품목(HS85 +110%, HS71 +216%)은 장기계약이 고정가면 상승분을 놓치고, 반대로 하락 국면에서는 손실을 떠안는다. 지금이 조건 재설계 시점이다.')]));
children.push(bullet([t('품목별 시트를 총액 용도로 쓰지 마라. ',{b:true}),t('원본 품목별 조회에는 HS 98류(특수분류)가 빠져 있어 총괄 대비 수출 −0.038%, 수입 −0.036% 과소집계된다. 품목 간 비교에는 문제없지만 총액 인용 시 반드시 총괄 시트를 쓸 것.')]));

children.push(h1('9. 분석의 한계'));
children.push(p('① 2026년은 7개월치뿐이라 연간 비교는 성립하지 않는다(2025년은 12월이 연중 최고였다). ② 계절지수는 2025년 1년치로만 산출한 근사치다. ③ 단가는 톤당 금액이므로 같은 품목의 시점 비교에만 유효하다. ④ 국가별 자료에 중량이 없어 건당 금액을 대리지표로 사용했다. ⑤ 명목 달러 기준으로 환율·물가 조정이 되어 있지 않다. ⑥ 6장의 원인은 데이터 밖 해석이며 검증 경로를 거쳐 확정해야 한다.'));
children.push(new Paragraph({spacing:{before:70,after:10},children:[t('[참고 그림] 국가별 반도체 동조성과 수출 증가율 — 오른쪽 위로 몰릴수록 지역 분산 효과가 없다',{sz:16,b:true,c:NAVY})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:30},children:[
  new ImageRun({type:'png',data:fs.readFileSync('html2/final_html_3961d951.png'),transformation:{width:430,height:201}})]}));
children.push(new Paragraph({spacing:{before:10,after:40},children:[t('가로축 = 각국 월별 수출과 HS85류 월별 수출의 19개월 상관계수, 세로축 = 2026.1~7 수출 증가율. 상위 20개국. 출처 : 분석12·분석15.',{sz:14,i:true,c:GRAY})]}));
children.push(new Paragraph({spacing:{before:40},border:{top:{style:BorderStyle.SINGLE,size:8,color:NAVY,space:3}},
  children:[t('근거 파일 : 수출입실적_통합분석_2025-2026.xlsx (원본 5개 시트 무수정 · 정제 5개 · 분석 15개 · 그래프 23종 · 수식 71,296개, 오류 0). 본 보고서의 모든 수치는 해당 파일에서 재현·검증할 수 있다.',{sz:14,c:GRAY})]}));

const doc=new Document({
  styles:{default:{document:{run:{font:F,size:19}}}},
  sections:[{properties:{page:{margin:{top:700,right:640,bottom:620,left:640}}},children}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('수출_심층분석_보고서.docx',b);console.log('written',b.length);});
