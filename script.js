const $ = id => document.getElementById(id);
let rawRows = [];
let headers = [];
let ticketRows = [];
let charts = {};
let columnMap = {key:null,status:null,priority:null,brand:null,owner:null};
const priorityColors = {high:"#f04444",highest:"#ffad36",critical:"#8c5cf6",blocker:"#ed7393",major:"#f04444",medium:"#4da7e8",low:"#42c477",lowest:"#42c477"};
const palette = ["#0968dc","#13b5a5","#ffab37","#8c5cf6","#91a4ba","#43bd76","#ed7393","#4da7e8"];
const aliases = {
 key:["key","issue key","ticket key","ticket id","issue id"],
 status:["status","issue status"],
 priority:["priority","management priority","product team priority","shape priority"],
 brand:["bsh brand (multiple)","brand/country/language","manufacturer"],
 owner:["closed by","assignee","owner","resolved by","assigned to"]
};
function notify(message){const t=$("toast");t.textContent=message;t.style.display="block";clearTimeout(notify.timer);notify.timer=setTimeout(()=>t.style.display="none",3500)}
function message(text,type=""){const el=$("loadMessage");el.textContent=text;el.className="load-message "+type}
function clean(v){return v===null||v===undefined?"":String(v).trim()}
function normalize(s){return clean(s).toLowerCase().replace(/[\s_\-/]+/g," ").trim()}
function findHeaderRow(matrix){
  const limit=Math.min(matrix.length,40);
  for(let i=0;i<limit;i++){
    const row=(matrix[i]||[]).map(normalize);
    const hasKey=row.some(v=>["key","issue key","ticket key"].includes(v));
    const hasSummary=row.some(v=>["summary","issue summary","title"].includes(v));
    const hasStatus=row.some(v=>["status","issue status"].includes(v));
    if(hasKey && (hasSummary || hasStatus)) return i;
  }
  return -1;
}
function findColumn(aliasList){
  const normalized=headers.map(normalize);
  // Prefer exact header matches in the priority order provided.
  for(const alias of aliasList){const idx=normalized.findIndex(h=>h===normalize(alias));if(idx>=0)return idx}
  // Fallback for common Jira variations.
  for(const alias of aliasList){const idx=normalized.findIndex(h=>h.includes(normalize(alias)));if(idx>=0)return idx}
  return -1;
}
function buildMapping(){
  for(const field of Object.keys(aliases)) columnMap[field]=findColumn(aliases[field]);
  const ids={key:"keyColumn",status:"statusColumn",priority:"priorityColumn",brand:"brandColumn",owner:"ownerColumn"};
  for(const [field,id] of Object.entries(ids)){
    const select=$(id);
    select.innerHTML='<option value="-1">— Not available —</option>'+headers.map((h,i)=>`<option value="${i}">${escapeHtml(h||`Column ${i+1}`)}</option>`).join("");
    select.value=String(columnMap[field]);
  }
  $("mappingPanel").hidden=false;
}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function loadFile(file){
  if(!file)return;
  $("fileName").textContent=file.name;
  const reader=new FileReader();
  reader.onload=event=>{
    try{
      if(!window.XLSX)throw new Error("Excel reader library did not load. Check your internet connection and reload the page.");
      const isCsv=/\.csv$/i.test(file.name);
      const workbook=XLSX.read(event.target.result,{type:isCsv?"string":"array",cellDates:false});
      if(!workbook.SheetNames.length)throw new Error("The workbook contains no worksheets.");
      const sheet=workbook.Sheets[workbook.SheetNames[0]];
      const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false});
      const headerRow=findHeaderRow(matrix);
      if(headerRow<0)throw new Error("Could not detect a header row. Make sure the sheet contains a column such as Key / Issue key and Summary.");
      headers=(matrix[headerRow]||[]).map((h,i)=>clean(h)||`Column ${i+1}`);
      rawRows=matrix.slice(headerRow+1).filter(row=>row.some(v=>clean(v)!==""));
      if(!rawRows.length)throw new Error("The detected header row has no data rows below it.");
      buildMapping();
      applyMapping();
      message(`Loaded ${ticketRows.length.toLocaleString()} rows from sheet "${workbook.SheetNames[0]}". Header row: ${headerRow+1}.`, "success");
      notify("Excel data loaded successfully");
    }catch(err){rawRows=[];ticketRows=[];resetDashboard();message(err.message,"error");notify("Unable to load workbook")}
  };
  if(/\.csv$/i.test(file.name))reader.readAsText(file);else reader.readAsArrayBuffer(file);
}
function applyMapping(){
  const ids={key:"keyColumn",status:"statusColumn",priority:"priorityColumn",brand:"brandColumn",owner:"ownerColumn"};
  for(const [field,id] of Object.entries(ids))columnMap[field]=Number($(id).value);
  ticketRows=rawRows.map(row=>({
    key:valueAt(row,columnMap.key),
    status:valueAt(row,columnMap.status),
    priority:valueAt(row,columnMap.priority),
    brand:valueAt(row,columnMap.brand),
    owner:valueAt(row,columnMap.owner)
  })).filter(t=>Object.values(t).some(v=>v!==""));
  updateDashboard();
}
function valueAt(row,index){return index<0||index===undefined?"":clean(row[index])}
function counts(field){const out={};for(const row of ticketRows){const value=row[field]||"Unspecified";out[value]=(out[value]||0)+1}return out}
function updateDashboard(){
  $("totalKpi").textContent=ticketRows.length.toLocaleString();
  $("highKpi").textContent=ticketRows.filter(t=>["high","highest","critical","blocker"].includes(normalize(t.priority))).length.toLocaleString();
  $("brandsKpi").textContent=new Set(ticketRows.map(t=>t.brand).filter(v=>v&&normalize(v)!=="unspecified")).size.toLocaleString();
  $("ownersKpi").textContent=new Set(ticketRows.map(t=>t.owner).filter(v=>v&&normalize(v)!=="unassigned"&&normalize(v)!=="unspecified")).size.toLocaleString();
  renderCharts();
}
function resetDashboard(){
  ["totalKpi","highKpi","brandsKpi","ownersKpi"].forEach(id=>$(id).textContent="—");
  for(const id of ["priorityChart","statusChart","brandChart"]){if(charts[id]){charts[id].destroy();delete charts[id]}const canvas=$(id);canvas.parentElement.classList.remove("has-data");}
  $("mappingPanel").hidden=true;
}
function makeChart(id,type,data,options){
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart($(id),{type,data,options});
  $(id).parentElement.classList.add("has-data");
}
function renderCharts(){
  const priorities=counts("priority");
  const preferred=["Highest","High","Critical","Blocker","Major","Medium","Low","Lowest","Minor","Unspecified"];
  const pLabels=[...preferred.filter(k=>priorities[k]),...Object.keys(priorities).filter(k=>!preferred.includes(k)).sort()];
  if(columnMap.priority<0){showEmpty("priorityChart","priorityEmpty","No priority column was detected.");}
  else makeChart("priorityChart","bar",{labels:pLabels,datasets:[{data:pLabels.map(k=>priorities[k]),backgroundColor:pLabels.map(k=>priorityColors[normalize(k)]||"#94a3b8"),borderRadius:4,barThickness:20}]},{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,grid:{color:"#edf2f7"}},y:{grid:{display:false}}}});
  const statuses=counts("status"),sLabels=Object.keys(statuses).sort((a,b)=>statuses[b]-statuses[a]);
  if(columnMap.status<0){showEmpty("statusChart","statusEmpty","No status column was detected.");}
  else makeChart("statusChart","doughnut",{labels:sLabels,datasets:[{data:sLabels.map(k=>statuses[k]),backgroundColor:sLabels.map((_,i)=>palette[i%palette.length]),borderColor:"#fff",borderWidth:2}]},{responsive:true,maintainAspectRatio:false,cutout:"58%",plugins:{legend:{position:"right",labels:{boxWidth:12,usePointStyle:true,pointStyle:"rectRounded",padding:12}}}});
  const brands=counts("brand"),bLabels=Object.keys(brands).sort((a,b)=>brands[b]-brands[a]).slice(0,12);
  if(columnMap.brand<0){showEmpty("brandChart","brandEmpty","No brand column was detected.");}
  else makeChart("brandChart","bar",{labels:bLabels,datasets:[{data:bLabels.map(k=>brands[k]),backgroundColor:bLabels.map((_,i)=>palette[i%palette.length]),borderRadius:4,barThickness:28}]},{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:"#edf2f7"}},x:{grid:{display:false},ticks:{maxRotation:35,minRotation:0,font:{size:10}}}}});
}
function showEmpty(chartId,emptyId,text){
  if(charts[chartId]){charts[chartId].destroy();delete charts[chartId]}
  $(chartId).parentElement.classList.remove("has-data");
  $(emptyId).textContent=text;
}
$("fileInput").addEventListener("change",e=>loadFile(e.target.files[0]));
$("applyMapping").addEventListener("click",()=>{applyMapping();message(`Dashboard updated using ${ticketRows.length.toLocaleString()} rows and your selected Excel columns.`,"success")});
const dz=$("dropzone");
dz.addEventListener("dragover",e=>{e.preventDefault();dz.classList.add("dragover")});
dz.addEventListener("dragleave",()=>dz.classList.remove("dragover"));
dz.addEventListener("drop",e=>{e.preventDefault();dz.classList.remove("dragover");if(e.dataTransfer.files[0])loadFile(e.dataTransfer.files[0])});
document.querySelectorAll(".nav button").forEach(btn=>btn.addEventListener("click",()=>{
  document.querySelectorAll(".nav button").forEach(b=>b.classList.remove("active"));btn.classList.add("active");
  const view=btn.dataset.view;
  if(view==="Upload Data")$("uploadSection").scrollIntoView({behavior:"smooth",block:"center"});
  else if(view==="Reports")document.querySelector(".charts").scrollIntoView({behavior:"smooth",block:"center"});
  else if(view==="About")notify("Upload a Jira/Sherpa Excel export. Header row and field columns are detected from the workbook.");
  else window.scrollTo({top:0,behavior:"smooth"});
}));
resetDashboard();
