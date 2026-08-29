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
function fig(file,capNo,cap,widthMm){
  const wPx=widthMm, hPx=Math.round(widthMm*0.487);
  return [new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100,after:30,line:240,lineRule:LineRuleType.AUTO},
      children:[new ImageRun({type:'png',data:fs.readFileSync('../figs/'+file),
        transformation:{width:wPx,height:hPx}})]}),
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

