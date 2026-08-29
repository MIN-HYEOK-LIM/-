const d=require('docx'); const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,HeadingLevel,AlignmentType,Table,TableRow,TableCell,
  WidthType,ShadingType,BorderStyle,ImageRun,PageBreak,Footer,PageNumber,LineRuleType,convertMillimetersToTwip}=d;
const D=require('./data.js');
const F='맑은 고딕';
const NAVY='1F3864', GREY='595959', RED='C00000', GREEN='1F6E43';
const W=9020;                    // 표 전체 폭 (DXA)
const NB={style:BorderStyle.NONE,size:0,color:'FFFFFF'};

function t(text,o={}){return new TextRun({text,font:F,size:o.sz||19,bold:o.b,italics:o.i,color:o.c||'000000'});}
function p(text,o={}){return new Paragraph({children:Array.isArray(text)?text:[t(text,o)],
  alignment:o.al,spacing:{before:o.sb??0,after:o.sa??90,line:o.line??252,lineRule:LineRuleType.AUTO},
  indent:o.ind?{left:o.ind}:undefined,
  border:o.bb?{bottom:{style:BorderStyle.SINGLE,size:6,color:o.bb}}:undefined});}
function h1(text){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:260,after:130},
  border:{bottom:{style:BorderStyle.SINGLE,size:10,color:NAVY}},
  children:[new TextRun({text,font:F,size:26,bold:true,color:NAVY})]});}
function h2(text){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:190,after:80},
  children:[new TextRun({text,font:F,size:21,bold:true,color:NAVY})]});}
function bullet(text,o={}){return new Paragraph({bullet:{level:0},spacing:{after:50,line:252,lineRule:LineRuleType.AUTO},
  children:Array.isArray(text)?text:[t(text,{sz:o.sz||19})]});}
function cell(v,o={}){
  return new TableCell({width:{size:o.w,type:WidthType.DXA},
    shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,
    margins:{top:40,bottom:40,left:70,right:70},
    children:[new Paragraph({alignment:o.al||AlignmentType.CENTER,spacing:{after:0,line:220,lineRule:LineRuleType.AUTO},
      children:[new TextRun({text:String(v),font:F,size:o.sz||15,bold:o.b,
        color:o.c||(o.hdr?'FFFFFF':'000000')})]})]});}
function table(cols,head,rows,o={}){
  const hr=new TableRow({tableHeader:true,children:head.map((h,i)=>
    cell(h,{w:cols[i],hdr:true,b:true,fill:o.hf||NAVY,sz:o.hsz||14}))});
  const rr=rows.map((r,ri)=>new TableRow({children:r.map((v,i)=>
    cell(v,{w:cols[i],sz:o.sz||15,al:(o.left&&o.left.includes(i))?AlignmentType.LEFT:AlignmentType.CENTER,
      fill:ri%2?'F2F5FA':undefined,b:o.bold&&o.bold.includes(i),
      c:(typeof v==='string'&&v.trim().startsWith('-')&&v.includes('%'))?RED:undefined}))}));
  return new Table({columnWidths:cols,width:{size:cols.reduce((a,b)=>a+b,0),type:WidthType.DXA},rows:[hr,...rr]});
}
function pngSize(buf){return {w:buf.readUInt32BE(16),h:buf.readUInt32BE(20)};}
function fig(file,capNo,cap,widthMm){
  const buf=fs.readFileSync('../figs/'+file); const sz=pngSize(buf);
  const wPx=widthMm, hPx=Math.round(widthMm*sz.h/sz.w);
  return [new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100,after:30,line:240,lineRule:LineRuleType.AUTO},
      children:[new ImageRun({type:'png',data:buf,transformation:{width:wPx,height:hPx}})]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:110,line:240,lineRule:LineRuleType.AUTO},
      children:[new TextRun({text:`[그림 ${capNo}] ${cap}`,font:F,size:15,color:GREY})]})];
}
function note(text){return new Paragraph({spacing:{before:40,after:130},
  children:[new TextRun({text,font:F,size:15,italics:true,color:GREY})]});}
function box(lines){
  return new Table({columnWidths:[W],width:{size:W,type:WidthType.DXA},
    borders:{top:{style:BorderStyle.SINGLE,size:12,color:NAVY},bottom:{style:BorderStyle.SINGLE,size:12,color:NAVY},
             left:NB,right:NB,insideHorizontal:NB,insideVertical:NB},
    rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},
      shading:{type:ShadingType.CLEAR,fill:'F2F5FA',color:'auto'},
      margins:{top:150,bottom:150,left:180,right:180},
      children:lines.map(l=>new Paragraph({spacing:{after:60,line:250,lineRule:LineRuleType.AUTO},
        children:[new TextRun({text:l.h?l.h:'',font:F,size:18,bold:true,color:NAVY}),
                  new TextRun({text:l.t,font:F,size:18})]}))})]})]});
}


const body=[];
body.push(new Paragraph({spacing:{after:30},children:[new TextRun({text:'대한민국 수출 데이터 분석',font:F,size:15,bold:true,color:GREY,characterSpacing:40})]}));
body.push(new Paragraph({spacing:{after:40},children:[new TextRun({text:'반도체 밖에서 알파를 찾는다',font:F,size:36,bold:true,color:NAVY})]}));
body.push(new Paragraph({spacing:{after:150},border:{bottom:{style:BorderStyle.SINGLE,size:14,color:NAVY}},
  children:[new TextRun({text:'2024.01~2026.07 관세청 수출입무역통계 월별 해부  |  바이사이드 투자 참고자료',font:F,size:17,color:GREY})]}));
body.push(box([
 {h:'무엇이 나왔나  ',t:'2026년 수출 급증의 82%는 반도체 3개 코드가 만들었고 이는 이미 알려진 사실이다. 그러나 반도체를 뺀 나머지도 7개월 중 6개월 플러스, 최근 3개월 +15.3%로 함께 좋아지고 있다.'},
 {h:'알파는 어디에  ',t:'금액 증가율이 아니라 (1) 최근 3개월 속도가 그 직전 3개월보다 빨라지고 (2) 그 성장이 단가가 아니라 물량에서 나오는 곳. 두 조건을 통과한 비반도체는 통신기기·계측검사장비·수동부품·정밀화학이다.'},
 {h:'피해야 할 곳  ',t:'정유·귀금속·합성수지·이차전지는 금액 증가율이 크지만 가속도가 죽었거나 물량이 없다. 금액만 보면 고점에서 잡는다.'}
]));

body.push(h1('1. 분석 방법론'));
body.push(h2('1.1  원자료와 가공'));
body.push(p('관세청 수출입무역통계(통관 수리일 기준 공식 집계)에서 받은 9개 파일이다. 2024년 1월~2026년 7월 31개월, 단위는 중량 톤·금액 천달러. 증권사 추정치나 잠정치가 섞이지 않은 원자료다.'));
body.push(bullet('총괄·품목별·성질별·신성질별 5개 파일 — HS 2단위 97개 류 전체. 국가 수출입 전 영역을 덮는다.'));
body.push(bullet('HS 4단위 4개 파일 — 27(석유)·29(유기화학)·39(플라스틱)·71(귀금속)·72(철강)·84(기계)·85(전기전자)·87(자동차)·89(선박)·90(정밀기기) 10개 류를 321개 코드로 쪼갠 것. 2026년 1~7월 국가 총수출의 84.6%를 차지한다. 나머지 영역(농수산·섬유 등)은 2단위 21개 부문으로 별도 확인했고 결론을 뒤집는 흐름은 없었다.'));
body.push(p('원자료는 숫자가 쉼표 포함 텍스트로 저장돼 그대로는 계산되지 않는다. 원본을 한 셀도 고치지 않고 참조만 하는 정제 계층을 만든 뒤, 그 위에 4단위 321개 코드 × 31개월 행렬을 쌓았다. 본문의 모든 숫자는 이 행렬의 계산 결과이며 손으로 입력한 값은 없다. 원본 총계행과 31개월 합계를 8개 항목 전부 대조했고 차이는 반올림 오차 수준(금액 ±1천달러)이다.'));
body.push(h2('1.2  분석 원칙 다섯 가지'));
body.push(bullet([t('① 누계가 아니라 월별로. ',{b:true}),t('"1~7월 누계 +50%"는 변화가 1월에 시작됐는지 6월에 시작됐는지를 감춘다. 투자에서 중요한 건 크기가 아니라 시점이므로 모든 증감을 전년 같은 달과 1:1로 비교했다.')]));
body.push(bullet([t('② 속도가 아니라 가속도를. ',{b:true}),t('최근 3개월(26.05~07) 전년동월비 − 직전 3개월(26.02~04) 전년동월비를 가속도로 정의했다. 성장률이 높아도 가속도가 마이너스면 정점을 지난 것이다. 두 부호의 조합으로 가속 성장·성장 둔화·바닥 통과·침체 심화 네 국면에 자동 배정했다.')]));
body.push(bullet([t('③ 금액을 단가와 물량으로. ',{b:true}),t('같은 +20%라도 단가가 올라서인지 물량이 늘어서인지에 따라 지속성이 다르다. 단가가 물량의 2배 이상 기여하면 가격 주도(사이클성), 반대면 물량 주도(실수요)로 분류했다.')]));
body.push(bullet([t('④ 조업일수를 보정. ',{b:true}),t('2026년 2월은 조업일수가 19.0일로 전년 22.0일보다 3일 적다. 이것만으로 명목 전년동월비가 실제보다 나쁘게 나온다. 금액÷조업일수의 일평균 기준을 함께 계산했다.')]));
body.push(bullet([t('⑤ 2단위가 아니라 4단위로. ',{b:true}),t('84류에는 일반기계와 메모리 모듈이, 72류에는 급증하는 반제품과 감소하는 판재가 함께 있다. 2단위로는 정반대 신호가 상쇄돼 사라진다. 321개 코드를 밸류체인 관점의 38개 투자 섹터로 다시 묶었다.')]));
body.push(p('이 지표들을 종합점수(성장 30 + 가속도 25 + 성장동력×증가월수 20 + 규모 15 + 변동성 역수 10)로 환산하고 국면과 결합해 비중확대 / 중립 / 매집 검토 / 회피 등급을 부여했다. 규모가 작으면서 월별 변동성이 큰 섹터는 한두 건의 선적으로 순위가 뒤집히므로 비중확대에서 제외했다.'));

body.push(h1('2. 데이터 해부'));
body.push(h2('2.1  반도체를 빼면 무엇이 남는가'));
body.push(p('2026년 수출은 매달 전년 대비 30~70% 늘었다. 대부분은 반도체가 만들었다 — 반도체 4개 코드(집적회로 8542, 메모리 모듈 8473, 저장장치 8523, 개별소자 8541)의 전년동월비는 7개월 내내 세 자릿수였고, 국가 수출에서 차지하는 몫이 1월 33.2%에서 7월 45.9%까지 올랐다. 여기까지는 시장이 아는 이야기다.'));
body.push(p([t('시장이 덜 보는 것은 나머지다. ',{b:true}),t('반도체를 뺀 수출은 7개월 중 6개월 플러스, 최근 3개월 +15.3%였다. 명목상 유일한 마이너스인 2월(-4.6%)조차 조업일수가 3일 적었던 달로, 일평균으로 보정하면 +10.4%로 뒤집힌다. 반대로 조업일수가 3.5일 늘었던 1월은 명목 +14.4%가 보정 후 -2.6%가 된다. "2026년은 반도체만 좋았다"는 통념은 데이터와 맞지 않는다.')]));
body.push(...fig('fig02.png',1,'반도체·비반도체 월별 전년동월비 (2025.01~2026.07)',400));
body.push(table([700,850,850,700,900,850,900,950,950,1370],
 ['월','총수출 억$','반도체 억$','반도체 비중','비반도체 억$','총수출 YoY','반도체 YoY','비반도체 YoY','비반도체 보정','조업일수 26/25'],
 D.macro,{hsz:12,sz:13}));
body.push(note('단위 억달러. 보정 YoY = 금액÷조업일수로 낸 일평균끼리의 전년동월비. 2월과 1월에서 명목과 보정의 부호가 뒤집히는 것이 조업일수 효과다.'));

body.push(h2('2.2  누가 빨라지고 있는가 — 가속도 순위'));
body.push(p('금액 증가율이 높은 섹터와 지금 빨라지고 있는 섹터는 다르다. 아래는 반도체를 뺀 규모 15억달러 이상 섹터를 가속도 순으로 세운 것이다. 오른쪽 두 열이 핵심으로, 7개월 중 몇 달이 플러스였는지(우연이 아닌지)와 성장이 단가에서 왔는지 물량에서 왔는지(지속되는지)를 본다. 상위 12개 중 11개가 비반도체이고, 합성수지(144억$)·디스플레이(104억$)·완성차(426억$)·수동부품(185억$) 같은 큰 축에서도 가속이 나타난다.'));
body.push(table([1900,900,700,950,950,850,850,1920],
 ['섹터','그룹','2026 억$','최근3M YoY','직전3M YoY','가속도 %p','증가 월수','성장동력'],
 D.accel.slice(0,11),{hsz:12,sz:13,left:[0],bold:[5]}));

body.push(h2('2.3  성장의 질 — 단가인가 물량인가'));
body.push(p('가속하는 섹터를 모두 사면 안 된다. 단가는 되돌려지지만 물량은 수요의 증거다. 아래 산점도에서 왼쪽 위(물량 주도)가 실수요형, 오른쪽 아래(가격 주도)가 사이클형이다.'));
body.push(p([t('두 조건을 모두 통과한 곳이 알파 후보다. ',{b:true}),t('가속도가 플러스이면서 물량이 함께 늘어난 비반도체는 계측·검사장비(물량 +13.3%), 정밀·기능성 화학(+12.2%), 철강 봉·형강(+41.4%), 철강 원료·반제품(+35.7%) 네 곳이다. 물량이 +12.0% 늘면서 단가까지 오른 통신기기를 더하면 다섯이다.')]));
body.push(...fig('fig06.png',2,'섹터별 단가 증감률(가로) vs 물량 증감률(세로) — 왼쪽 위가 실수요형',350));

body.push(h2('2.4  4단위로 내려가면 보이는 것'));
body.push(p('섹터 평균은 내부 분화를 감춘다. 아래는 알파 후보를 구성하는 개별 4단위 코드를 월별로 편 것이다. 오른쪽 7개 열이 2026년 1~7월의 전년동월비다.'));
body.push(table([560,1900,620,700,660,660,560,560,560,560,560,560,560],
 ['HS4','품목','2026 억$','26 YoY','단가Δ','물량Δ','1월','2월','3월','4월','5월','6월','7월'],
 D.alpha,{hsz:12,sz:12,left:[1]}));
body.push(bullet([t('휴대폰·통신장비(8517)는 6~7월에 다시 붙었다. ',{b:true}),t('4~5월 +13%, +7%로 식는 듯하다가 6월 +62%, 7월 +123%로 재가속. 단가(+42.2%)와 물량(+12.0%)이 함께 올라 믹스 개선과 수요 증가가 동시에 진행 중이다.')]));
body.push(bullet([t('측정·검사기기(9031)는 단가가 8.5% 내렸는데 물량이 18.4% 늘었다. ',{b:true}),t('가격을 깎아 파는 게 아니라 저가 물량이 늘어난 것으로, 7월 +42%는 가속을 보여준다. 오실로스코프·전기검사(9030)는 단가·물량이 함께 올라 7개월 중 6개월 30% 이상 증가했다.')]));
body.push(bullet([t('접속기기(8536)와 인쇄회로(8534)는 조용히 꾸준하다. ',{b:true}),t('8536은 물량 +15.6%로 순수 물량 성장, 8534는 7개월 내내 플러스. 둘 다 세트 수요를 후행하는 부품으로, 반도체 호황의 낙수가 실제로 도달했다는 증거다.')]));
body.push(bullet([t('철강 반제품(7207)·봉강(7214)은 극단적이다. ',{b:true}),t('7207은 6월 +649%, 7월 +278%, 7214는 물량이 301.9% 늘었다. 다만 규모가 2.8억·4.2억달러로 작아 대형 계약 한두 건에 좌우될 수 있다.')]));

body.push(h2('2.5  같은 산업 안에서 상류와 하류가 갈린다'));
body.push(p('2단위로 "화학"이나 "철강"을 하나로 보면 정반대로 움직이는 두 사이클을 섞어 사게 된다. 4단위로 체인을 펴면 이렇게 갈린다.'));
body.push(table([1050,2900,3900,1170],['체인','앞단(상류·성장)','뒷단(하류·부진)','격차'],D.chain,{hsz:12,sz:13,left:[1,2],bold:[3]}));
body.push(p([t('화학이 특히 중요하다. ',{b:true}),t('정유(2710)가 +38.3%인 동안 기초유분은 +2.5%, 합성수지는 +5.3%에 그쳤다. 같은 원유를 쓰지만 정제와 석유화학은 완전히 다른 사이클이며 2026년의 이익은 정제에 몰려 있다. "화학주"를 묶어 담으면 오르는 쪽과 내리는 쪽을 함께 사게 된다. 자동차에서는 전동화가 재개됐다 — 축전지(8507)÷가솔린엔진(8407) 비율이 5.43(24년)→4.03(25년)→6.34(26년)로 뛰었다.')]));

body.push(...fig('fig09.png',3,'철강 체인 월별 전년동월비 — 반제품·봉형강이 오르는 동안 판재는 내린다',360));
body.push(h2('2.6  금액과 마진이 반대로 가는 곳'));
body.push(p('수출 금액이 늘어도 남는 것이 줄어드는 경우가 있다. 수출 톤당 단가를 투입 원자재의 수입 톤당 단가로 나눠 스프레드를 근사했다(절대 수준이 아니라 방향을 본다).'));
body.push(table([2450,1450,1080,1180,1180,1680],['지표','구성 HS4','2026 평균','최근 3개월','직전 3개월','방향'],D.margin,{hsz:12,sz:13,left:[0]}));
body.push(p([t('정유가 대표적이다. ',{b:true}),t('수출 금액은 +39.3%인데 정제마진 프록시는 최근 3개월 1.33으로 직전 1.77보다 좁아졌다. 금액 증가율만 보고 들어가면 마진 정점을 지난 뒤에 잡는다. 롤마진과 석화 스프레드도 같은 방향이고, 반대로 의료기기·합성수지 단가는 개선됐다.')]));

body.push(h1('3. 바이사이드 관점의 투자 판단'));
body.push(p('반도체가 좋다는 것은 가격에 반영돼 있다. 이 데이터에서 얻을 알파는 두 종류다 — (가) 반도체 호황의 낙수가 실제로 도달했는데 아직 그렇게 취급받지 않는 후방 부품, (나) 반도체와 무관하게 자체 사이클이 돌아선 섹터.'));
body.push(h2('3.1  비중확대 — 가속과 물량이 모두 확인된 곳'));
const OW=[
 ['통신기기\n8517·8529','73','+48.1%','+23.7','단가 +42.2%\n물량 +12.0%','6월 +62%, 7월 +123%로 재가속. 7개월 내내 플러스. 단가와 물량이 함께 올라 믹스 개선과 수요 증가가 동시 진행.','8517 월별 물량. 단가만 남고 물량이 빠지면 축소.'],
 ['계측·검사장비\n9030·9031·9027','38','+21.3%','+15.3','단가 +2.2%\n물량 +13.3%','9031은 단가가 8.5% 내렸는데도 물량이 18.4% 증가, 7월 +42%. 반도체 후공정 검사 수요의 낙수가 도달한 증거.','소자−장비 격차. 좁혀지면 장비까지 확산.'],
 ['수동부품·회로\n8534·8536·8529','185','+13.7%','+6.1','단가 +9.5%\n물량 +3.8%','규모와 성장을 함께 가진 드문 섹터. 8536은 물량 +15.6%로 순수 물량 성장. 세트 수요 후행 구간의 초입.','반도체 소자 사이클과의 시차.'],
 ['정밀·기능성 화학\n2917 등','26','+18.6%','+10.8','단가 +5.7%\n물량 +12.2%','범용 석유화학(+5.3%)이 정체하는 동안 홀로 앞선다. 물량 주도, 7개월 중 6개월 증가.','범용과의 격차 유지 여부.']
];
body.push(table([1360,560,700,620,1080,2900,1800],
 ['섹터','2026 억$','최근3M','가속도','성장동력','투자 논거','확인 지표'],OW,{hsz:12,sz:12,left:[5,6],hf:'1F6E43'}));
body.push(h2('3.2  전술적 비중확대 · 관찰 · 회피'));
body.push(bullet([t('전술적 비중확대 — 철강 원료·반제품·봉형강. ',{b:true}),t('가속도 +65.6%p로 전 섹터 1위(7207 6월 +649%, 7214 물량 +301.9%). 판재(-6.5%)와 정반대다. 다만 합쳐 28억달러로 작아 개별 계약 리스크가 크므로 소규모·분산이 맞다. 확인 지표는 판재/봉형강 비율(10.18→6.16)이 계속 내려가는지다. 컴퓨터·서버(8471)는 4월 +73%, 7월 +68%로 가속 중이나 규모 6억달러라 단독 포지션보다 확인용 지표로 쓰는 편이 낫다.')]));
body.push(bullet([t('관찰 — 자동차 완성차·부품. ',{b:true}),t('비반도체 최대 섹터인 완성차(426억$)가 2년 부진 끝에 6월 +5%, 7월 +7%로 돌아섰다(가속도 +9.8%p). 다만 물량은 아직 -2.3%로 단가가 버티는 구조여서 물량이 함께 도는지를 8~9월에 확인해야 한다. 부품(105억$)은 7월에야 첫 플러스(+3%)로, 통상 완성차를 후행하므로 다음 순번일 수 있다. 가전은 감소폭이 5월 -27%에서 6~7월 -6%, -5%로 줄었으나 반등이 아니라 감소 둔화이므로 아직 매수 구간이 아니다.')]));
body.push(bullet([t('회피 — 금액에 속기 쉬운 함정. ',{b:true}),t('아래 다섯 섹터의 공통점은 금액 증가율이 두 자릿수 이상인데 물량이 없거나 가속도가 죽었다는 것이다. 특히 귀금속 소재는 증가율 1위(+145.7%)이자 가속도 최하위(-110.9%p)로, 3월 +431%에서 7월 +93%까지 다섯 달 연속 속도가 떨어졌다.')]));
body.push(table([1320,560,700,620,700,700,4420],
 ['섹터','2026 억$','최근3M','가속도','단가Δ','물량Δ','왜 조심해야 하나'],D.trap,{hsz:12,sz:12,left:[6],hf:'9C2B2B'}));

body.push(h2('3.3  포트폴리오 구성 제안'));
const PF=[
 ['코어 (유지)','반도체 소자·모듈·저장장치','국가 수출 증가의 82%. 다만 성장의 대부분이 단가여서 신규 자금은 단가 지표가 꺾이는지 보고 결정','비중 유지'],
 ['알파 (신규)','통신기기 · 계측검사장비 · 수동부품 · 정밀화학','가속 + 물량 두 조건을 모두 통과. 반도체 낙수가 실제 도달했으나 아직 그렇게 취급받지 않음','신규 편입'],
 ['위성 (소액)','철강 원료·반제품 · 봉형강 · 컴퓨터·서버','가속도 최상위이나 규모가 작아 변동성 큼. 분산 전제','소규모 분산'],
 ['관찰','자동차 완성차·부품 · 가전','턴어라운드 초입. 물량 전환 확인 후 진입','8~9월 재확인'],
 ['축소·회피','정유 · 귀금속 소재 · 합성수지 · 이차전지 · 철강 판재','금액은 크나 가속도 소멸 또는 물량 부재. 마진 프록시도 악화','비중 축소']
];
body.push(table([1150,2950,3720,1200],['구분','대상','근거','액션'],PF,{hsz:12,sz:13,left:[1,2]}));

body.push(h2('3.4  리스크와 이 자료의 한계'));
body.push(bullet('집중도 — 2026년 7월 기준 수출의 45.9%가 반도체 4개 코드다. 반도체 단가가 10% 되돌려지면 국가 수출 증가율은 4~5%p 깎인다. 위 알파 바스켓 중 통신기기·계측검사·수동부품은 반도체 전방 수요에 연동돼 있어 완전한 헤지가 아니다. 자동차·철강·의료기기가 상대적으로 독립적이다.'));
body.push(bullet('이 자료는 무역통계 기반 정량 스크리닝이다. 개별 기업의 실적·밸류에이션·수급은 반영돼 있지 않다. 수출이 좋아도 주가에 이미 반영됐거나 수출 주체가 상장사가 아닐 수 있으므로 종목 선택은 별도 검증이 필요하다.'));
body.push(bullet('단가는 톤당 단가다. 대당·개당 ASP의 방향을 가늠하는 근사치이지 ASP 자체가 아니며, 단가 상승의 일부는 고부가 품목 비중 확대일 수 있다.'));
body.push(bullet('조선·방산처럼 인도 시점에 금액이 몰리는 섹터는 단월 증감률이 추세가 아니다. 조선은 3개월 이동합 기준 2026년 내내 +11~30%로 안정적이나 단월로는 -6%~+64%로 널뛴다. 또한 통관 기준이므로 해외 현지생산·중계무역은 잡히지 않아, 완성차·가전처럼 생산 이전이 진행 중인 산업은 수출 지표가 실제 사업 규모를 과소평가할 수 있다.'));
body.push(new Paragraph({spacing:{before:150},border:{top:{style:BorderStyle.SINGLE,size:6,color:'BFBFBF'}},
  children:[new TextRun({text:'자료 : 관세청 수출입무역통계(2024.01~2026.07, 통관 수리일 기준). 원자료 9개 파일을 HS 4단위 321개 코드 × 31개월 행렬로 재구성해 분석했으며, 본문의 모든 수치와 그림은 그 계산 결과다. 본 자료는 정보 제공 목적이며 특정 종목의 매매를 권유하지 않는다.',font:F,size:13,italics:true,color:GREY})]}));

const doc=new Document({
  styles:{default:{document:{run:{font:F,size:19},paragraph:{spacing:{line:252,lineRule:LineRuleType.AUTO}}}}},
  sections:[{properties:{page:{margin:{top:convertMillimetersToTwip(16),bottom:convertMillimetersToTwip(14),
      left:convertMillimetersToTwip(17),right:convertMillimetersToTwip(17)}}},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({children:[PageNumber.CURRENT],font:F,size:15,color:GREY})]})]})},
    children:body}]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('수출데이터_분석보고서.docx',b);console.log('saved',b.length);});
