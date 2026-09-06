
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ dest: path.join(__dirname, "uploads") });
app.use(express.json({limit:"10mb"}));
app.use(express.static(path.join(__dirname,"public")));

function clean(v){ return v === null || v === undefined ? "" : String(v).trim(); }
function teacherIds(v){
  return clean(v).replace(/\s+/g,"").split(/[\\/,;|]+/).filter(Boolean);
}
function cellValue(cell){
  if(cell.value === null || cell.value === undefined) return "";
  if(typeof cell.value === "object"){
    if(cell.value.text) return cell.value.text;
    if(cell.value.result !== undefined) return String(cell.value.result);
  }
  return String(cell.value);
}
function colLetter(n){
  let s="";
  while(n){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}
  return s;
}
function rgb(color){
  if(!color) return null;
  if(color.argb) return "#"+color.argb.slice(-6);
  if(color.rgb) return "#"+color.rgb.slice(-6);
  return null;
}
function styleJSON(cell){
  const f=cell.font||{}, fill=cell.fill||{}, a=cell.alignment||{}, b=cell.border||{};
  const side=x=>x?{style:x.style||null,color:rgb(x.color)}:null;
  return {
    font:{name:f.name||"Arial",size:f.size||11,bold:!!f.bold,italic:!!f.italic,underline:!!f.underline,color:rgb(f.color)},
    fill:{type:fill.type||null,color:rgb(fill.fgColor),bg:rgb(fill.bgColor)},
    align:{horizontal:a.horizontal||null,vertical:a.vertical||null,wrap:!!a.wrapText,textRotation:a.textRotation||0,indent:a.indent||0},
    border:{top:side(b.top),bottom:side(b.bottom),left:side(b.left),right:side(b.right)}
  };
}
function sheetJSON(ws){
  const cells=[];
  const maxRow=ws.rowCount, maxCol=ws.columnCount;
  for(let r=1;r<=maxRow;r++){
    for(let c=1;c<=maxCol;c++){
      const cell=ws.getCell(r,c);
      const value=cellValue(cell);
      const hasStyle=cell.style && (cell.style.font || cell.style.fill || cell.style.border || cell.style.alignment);
      if(value!=="" || hasStyle){
        cells.push({
          r,c,v:value,s:styleJSON(cell),
          w:ws.getColumn(c).width||10,
          h:ws.getRow(r).height||15
        });
      }
    }
  }
  const merges=[...ws.model.merges||[]];
  return {
    name:ws.name,maxRow,maxCol,cells,merges,
    colWidths:Array.from({length:maxCol},(_,i)=>ws.getColumn(i+1).width||10),
    rowHeights:Array.from({length:maxRow},(_,i)=>ws.getRow(i+1).height||15),
    views:ws.views||[]
  };
}
function parseTeachers(wb){
  const ws=wb.getWorksheet("Лист2");
  const map={};
  if(!ws) return map;
  ws.eachRow(row=>{
    for(let c=1;c<row.cellCount;c++){
      const id=clean(row.getCell(c).value);
      const name=clean(row.getCell(c+1).value);
      if(/^\d+$/.test(id) && name && !map[id]) map[id]=name;
    }
  });
  return map;
}
function classColumns(ws){
  const out=[];
  for(let c=1;c<=ws.columnCount;c++){
    const v=clean(ws.getCell(3,c).value);
    if(v && /sinf/i.test(v)) out.push({col:c,className:v});
  }
  if(!out.length){
    for(let r=1;r<=Math.min(ws.rowCount,8);r++){
      for(let c=1;c<=ws.columnCount;c++){
        const v=clean(ws.getCell(r,c).value);
        if(v && /sinf/i.test(v)) out.push({col:c,className:v});
      }
      if(out.length) break;
    }
  }
  return out;
}
function parseLessons(ws, teachers){
  const cols=classColumns(ws), lessons=[];
  // The school files repeat blocks whose first two columns are day code + lesson number.
  // A lesson block starts whenever column B is 1 and the next rows contain 2 and 3.
  for(let r=1;r<=ws.rowCount-2;r++){
    const a=clean(ws.getCell(r,1).value), b=clean(ws.getCell(r,2).value);
    const b2=clean(ws.getCell(r+1,2).value), b3=clean(ws.getCell(r+2,2).value);
    if(b==="1" && b2==="2" && b3==="3"){
      // For each block use up to 7 visible periods, but only rows with real lesson numbers.
      let day = ({D:"Dushanba",U:"Seshanba",SH:"Chorshanba",A:"Payshanba",N:"Juma",B:"Shanba",S:"Yakshanba",E:"Yakshanba"})[a] || a || `Kun ${r}`;
      for(let i=0;i<7 && r+i<=ws.rowCount;i++){
        const rr=r+i, period=Number(clean(ws.getCell(rr,2).value));
        if(!period) continue;
        for(const cc of cols){
          const teacherRaw=clean(ws.getCell(rr,cc.col).value);
          const subject=clean(ws.getCell(rr,cc.col+1).value);
          if(!teacherRaw && !subject) continue;
          const ids=teacherIds(teacherRaw);
          lessons.push({
            id:`${ws.name}__${rr}__${cc.col}`,
            sheet:ws.name,row:rr,teacherCol:cc.col,subjectCol:cc.col+1,
            day,period,className:cc.className,subject,teacherRaw,teacherIds:ids,
            teacherNames:ids.map(id=>teachers[id]||`ID ${id}`)
          });
        }
      }
    }
  }
  return lessons;
}
async function readWorkbook(filePath){
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const teachers=parseTeachers(wb);
  const sheets=wb.worksheets.map(ws=>({name:ws.name, view:sheetJSON(ws), lessons:parseLessons(ws,teachers)}));
  return {sheets,teachers};
}
function detect(lessons){
  const conflicts=[];
  const byTeacher=new Map(), byClass=new Map();
  for(const l of lessons){
    for(const t of l.teacherIds){
      const k=`${t}|${l.day}|${l.period}`;
      if(!byTeacher.has(k))byTeacher.set(k,[]);
      byTeacher.get(k).push(l);
    }
    const k=`${l.className}|${l.day}|${l.period}`;
    if(!byClass.has(k))byClass.set(k,[]);
    byClass.get(k).push(l);
  }
  for(const [k,ls] of byTeacher){
    const classes=[...new Set(ls.map(x=>x.className))];
    if(classes.length>1){
      const tid=k.split("|")[0];
      conflicts.push({type:"teacher",day:ls[0].day,period:ls[0].period,teacherId:tid,
        teacherName:ls[0].teacherNames[ls[0].teacherIds.indexOf(tid)]||`ID ${tid}`,lessons:ls});
    }
  }
  for(const [k,ls] of byClass){
    if(ls.length>1) conflicts.push({type:"class",day:ls[0].day,period:ls[0].period,className:ls[0].className,lessons:ls});
  }
  return conflicts;
}
function dailyWarnings(lessons){
  const map=new Map(), warnings=[];
  for(const l of lessons){
    const k=`${l.className}|${l.day}`;
    if(!map.has(k))map.set(k,{className:l.className,day:l.day,set:new Set()});
    map.get(k).set.add(l.period);
  }
  for(const x of map.values()){
    const m=x.className.match(/(\d+)/), grade=m?Number(m[1]):99;
    const max=grade<=4?5:6;
    const count=x.set.size;
    if(count>max) warnings.push({...x,count,max,type:"daily-limit"});
  }
  return warnings;
}
function suggestions(lessons, c){
  const days=[...new Set(lessons.map(x=>x.day))], periods=[...new Set(lessons.map(x=>x.period))].sort((a,b)=>a-b);
  const out=[];
  for(const day of days) for(const period of periods){
    if(day===c.day&&period===c.period)continue;
    const occ=lessons.filter(x=>x.day===day&&x.period===period);
    const occC=new Set(occ.map(x=>x.className)), occT=new Set(occ.flatMap(x=>x.teacherIds));
    const classFree=c.lessons.every(x=>!occC.has(x.className));
    const teacherFree=c.type!=="teacher" || c.lessons.every(x=>x.teacherIds.every(t=>!occT.has(t)));
    if(classFree&&teacherFree)out.push({day,period});
  }
  return out.slice(0,10);
}
app.post("/api/upload",upload.single("file"),async(req,res)=>{
  try{
    if(!req.file)return res.status(400).json({error:"Excel fayl yuborilmadi."});
    const data=await readWorkbook(req.file.path);
    res.json(data);
    fs.unlink(req.file.path,()=>{});
  }catch(e){console.error(e);res.status(500).json({error:"Excel o‘qilmadi.",detail:e.message});}
});
app.post("/api/check",(req,res)=>{
  const lessons=req.body.lessons||[];
  const conflicts=detect(lessons).map(c=>({...c,suggestions:suggestions(lessons,c)}));
  res.json({conflicts,warnings:dailyWarnings(lessons),count:conflicts.length});
});
app.post("/api/export",upload.single("original"),async(req,res)=>{
  try{
    if(!req.file)return res.status(400).json({error:"Original Excel fayl kerak."});
    const moves=JSON.parse(req.body.moves||"[]");
    const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(req.file.path);
    for(const mv of moves){
      const ws=wb.getWorksheet(mv.sheet); if(!ws)continue;
      // Move the complete class lesson pair (teacher + subject) from source row to target row.
      const sourceRow=Number(mv.row), targetRow=Number(mv.targetRow);
      const tc=Number(mv.teacherCol), sc=Number(mv.subjectCol);
      if(!sourceRow||!targetRow||!tc||!sc)continue;
      const srcTeacher=ws.getCell(sourceRow,tc).value, srcSubject=ws.getCell(sourceRow,sc).value;
      ws.getCell(sourceRow,tc).value=null; ws.getCell(sourceRow,sc).value=null;
      ws.getCell(targetRow,tc).value=srcTeacher; ws.getCell(targetRow,sc).value=srcSubject;
    }
    const out=path.join(__dirname,"uploads",`dars_jadvali_${Date.now()}.xlsx`);
    await wb.xlsx.writeFile(out);
    res.download(out,"dars_jadvali_natija.xlsx",()=>fs.unlink(out,()=>{}));
    fs.unlink(req.file.path,()=>{});
  }catch(e){console.error(e);res.status(500).json({error:"Excel yaratishda xato.",detail:e.message});}
});
app.get("*",(req,res)=>res.sendFile(path.join(__dirname,"public","index.html")));
app.listen(PORT,()=>console.log(`Server: http://localhost:${PORT}`));
