const $ = (id) => document.getElementById(id);
let workbook = null;
let currentFile = null;

const pageNames = {dashboard:"대시보드",data:"데이터 탐색",analysis:"산업연관분석",simulation:"시뮬레이션",report:"결과 보고서"};
function showPage(name){
  document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));
  $("page-"+name).classList.add("active");
  document.querySelectorAll(".nav-item").forEach(b=>b.classList.toggle("active",b.dataset.page===name));
  $("page-title").textContent=pageNames[name];
}
document.querySelectorAll(".nav-item").forEach(b=>b.addEventListener("click",()=>showPage(b.dataset.page)));
$("go-upload").addEventListener("click",()=>showPage("data"));
document.querySelectorAll(".go-data").forEach(b=>b.addEventListener("click",()=>showPage("data")));

function toast(message){
  const t=$("toast"); t.textContent=message; t.classList.add("show");
  clearTimeout(toast.timer); toast.timer=setTimeout(()=>t.classList.remove("show"),2600);
}
function updateStats(){
  $("stat-file").textContent=currentFile?currentFile.name:"없음";
  $("stat-sheets").textContent=workbook?workbook.SheetNames.length:"0";
  $("stat-class").textContent=$("classification").value||"미선택";
  $("stat-state").textContent=workbook?"파일 인식 완료":"대기";
  $("stat-state").style.color=workbook?"#16866a":"#d18b31";
}
function loadFile(file){
  if(!file)return;
  if(!window.XLSX){toast("엑셀 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하세요.");return;}
  const ext=file.name.split(".").pop().toLowerCase();
  if(!["xlsx","xls","xlsm"].includes(ext)){toast("Excel 파일(.xlsx, .xls, .xlsm)을 선택해 주세요.");return;}
  const reader=new FileReader();
  reader.onload=(e)=>{
    try{
      workbook=XLSX.read(e.target.result,{type:"array",cellDates:true});
      currentFile=file;
      $("workbook-area").classList.remove("hidden");
      $("workbook-name").textContent=file.name;
      $("workbook-meta").textContent=`${workbook.SheetNames.length}개 시트 · ${(file.size/1024/1024).toFixed(2)} MB · 브라우저에서 읽음`;
      $("sheet-count").textContent=workbook.SheetNames.length;
      renderSheets();
      if(workbook.SheetNames.length)renderPreview(workbook.SheetNames[0]);
      updateStats();
      toast("엑셀 파일을 읽었습니다. 시트 구성을 확인해 주세요.");
    }catch(err){console.error(err);toast("파일을 읽지 못했습니다. 손상 여부와 파일 형식을 확인하세요.");}
  };
  reader.onerror=()=>toast("파일을 읽는 중 오류가 발생했습니다.");
  reader.readAsArrayBuffer(file);
}
function renderSheets(){
  const list=$("sheet-list");list.innerHTML="";
  workbook.SheetNames.forEach((name,i)=>{
    const btn=document.createElement("button");
    btn.className="sheet-item"+(i===0?" active":"");
    btn.textContent=name;
    btn.title=name;
    btn.addEventListener("click",()=>{
      document.querySelectorAll(".sheet-item").forEach(x=>x.classList.remove("active"));
      btn.classList.add("active");renderPreview(name);
    });
    list.appendChild(btn);
  });
}
function renderPreview(name){
  const sheet=workbook.Sheets[name];
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false});
  $("preview-title").textContent=name;
  $("preview-subtitle").textContent=`${rows.length.toLocaleString("ko-KR")}개 행 · 상위 30행 미리보기`;
  const wrap=$("preview-wrap");wrap.classList.remove("empty-preview");
  if(!rows.length){wrap.classList.add("empty-preview");wrap.textContent="이 시트에는 표시할 데이터가 없습니다.";return;}
  const maxCols=Math.min(Math.max(...rows.slice(0,30).map(r=>r.length),1),14);
  const table=document.createElement("table");
  const thead=document.createElement("thead"),tr=document.createElement("tr");
  for(let c=0;c<maxCols;c++){const th=document.createElement("th");th.textContent=c===0?"행 / 항목":`열 ${c+1}`;tr.appendChild(th);}
  thead.appendChild(tr);table.appendChild(thead);
  const tbody=document.createElement("tbody");
  rows.slice(0,30).forEach((row,ri)=>{
    const tr=document.createElement("tr");
    for(let c=0;c<maxCols;c++){const td=document.createElement("td");td.textContent=row[c]===undefined?"":String(row[c]);td.title=td.textContent;tr.appendChild(td);}
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);wrap.innerHTML="";wrap.appendChild(table);
}
$("file-input").addEventListener("change",e=>loadFile(e.target.files[0]));
$("clear-file").addEventListener("click",()=>{
  workbook=null;currentFile=null;$("file-input").value="";$("workbook-area").classList.add("hidden");
  $("classification").value="";updateStats();toast("업로드한 파일을 해제했습니다.");
});
$("classification").addEventListener("change",()=>{
  $("class-hint").textContent=$("classification").value?`선택한 기준: ${$("classification").value} · 실제 포함 여부는 시트 내용과 대조해야 합니다.`:"자료 구조 확인 후 선택하세요.";
  updateStats();
});
const drop=$("drop-zone");
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("drag");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("drag");}));
drop.addEventListener("drop",e=>loadFile(e.dataTransfer.files[0]));
updateStats();