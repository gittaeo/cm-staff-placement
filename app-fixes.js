/* Staff placement reliability fixes: all imported data remains in this tab's memory. */
(function(){
"use strict";

const REQUIRED_EXPORT_COLUMNS=["직원번호","직원명","현장명","직무/직급","배치 시작일","배치 종료일","배치 구분","진행 상태","중복 배치 여부","날짜 오류 여부"];
const SOURCE_COLUMNS=["원본 파일명","원본 시트명","원본 행 번호"];
const HEADER_ALIASES={
  employeeId:["직원번호","사번","id","사원번호"], employeeName:["직원명","성명","이름","직원"],
  siteName:["현장명","현장","프로젝트명","사업명","근무현장"], role:["직무직급","직무","직급","분야","담당업무"],
  startDate:["배치시작일","시작일","착수일","투입일"], endDate:["배치종료일","종료일","완료일","철수일"],
  assignmentType:["배치구분","구분","배치상태","계약실제"]
};
let allAssignments=[];
let importReport={files:0,sheets:0,rawRows:0,accepted:0,excluded:[],duplicateSource:0,errors:[],details:[]};
let composing=false;
let searchTimer=0;

function safe(v,fallback=""){ return v===undefined||v===null?fallback:String(v); }
function normalizeText(v){ return safe(v).replace(/[\u00a0\u200b\u3000\ufeff]/g," ").replace(/[\r\n\t]+/g," ").replace(/\s+/g," ").trim(); }
function normalizeHeader(v){ return normalizeText(v).toLowerCase().replace(/[()（）\[\]{}\s_\-./\\:·]/g,""); }
function headerField(v){ const n=normalizeHeader(v); for(const [key,list] of Object.entries(HEADER_ALIASES)) if(list.some(x=>normalizeHeader(x)===n)) return key; return null; }
function employeeKey(r){ const id=normalizeText(r.employeeId); return id?"id:"+id.toLowerCase():"name:"+normalizeText(r.employeeName).toLowerCase(); }
function siteKey(v){ return normalizeText(v).toLowerCase(); }
function validISO(v){ return /^\d{4}-\d{2}-\d{2}$/.test(safe(v))&&!isNaN(new Date(v+"T00:00:00").getTime()); }
function parseDateValue(v,date1904){
  if(v instanceof Date&&!isNaN(v)){ return v.getFullYear()+"-"+pad2(v.getMonth()+1)+"-"+pad2(v.getDate()); }
  if(typeof v==="number"&&isFinite(v)) return dateFromExcel(v,!!date1904);
  const s=normalizeText(v); if(!s) return "";
  if(/^\d+(\.\d+)?$/.test(s)){ const n=Number(s); if(n>1000&&n<100000) return dateFromExcel(n,!!date1904); }
  let m=s.match(/(\d{4})\s*(?:년|[.\/-])\s*(\d{1,2})\s*(?:월|[.\/-])\s*(\d{1,2})(?:\s*일)?/);
  if(!m) m=s.match(/^(\d{4})(\d{2})(\d{2})/);
  if(!m) return "";
  const iso=m[1]+"-"+pad2(m[2])+"-"+pad2(m[3]); return validISO(iso)?iso:"";
}
function kindLabel(v){ const s=normalizeText(v); if(/실제/.test(s))return "실제배치"; if(/예정|변경|수정|조정/.test(s))return "변경 예정 배치"; if(/계약/.test(s)||!s)return "계약배치"; return s; }
function kindCode(v){ const s=kindLabel(v); return s==="실제배치"?"actual":s==="변경 예정 배치"?"pending":"contract"; }
function statusLabel(r){ if(!validISO(r.startDate)||!validISO(r.endDate)||r.endDate<r.startDate)return "확인 필요"; const t=todayISO(); return t<r.startDate?"배치 예정":t>r.endDate?"배치 종료":"배치 중"; }
function rowMap(row){ const o={}; (Array.isArray(row)?row:[]).forEach(c=>{if(c&&Number.isInteger(c.i))o[c.i]=c.val;}); return o; }
function findHeader(sheet){
  const rows=Array.isArray(sheet&&sheet.rows)?sheet.rows:[]; let best={index:-1,score:0,map:{},headers:{}};
  rows.slice(0,30).forEach((row,index)=>{ const vals=rowMap(row), map={},headers={}; let score=0;
    Object.keys(vals).forEach(i=>{ const h=normalizeText(vals[i]); headers[i]=h; const f=headerField(h); if(f&&!Object.prototype.hasOwnProperty.call(map,f)){map[f]=Number(i);score++;} });
    if(score>best.score)best={index,score,map,headers};
  }); return best;
}
function sourceFingerprint(r){ return [employeeKey(r),siteKey(r.siteName),r.startDate,r.endDate,normalizeText(r.assignmentType)].join("|"); }
function ensureLegacyState(){
  const sites=new Map(),staff=new Map(); state.sites=[];state.staff=[];state.placements=[];state.changes=Array.isArray(state.changes)?state.changes:[];
  allAssignments.filter(r=>!r.excludedReason).forEach(r=>{
    const sk=siteKey(r.siteName); if(!sites.has(sk)){const x={id:uid(),name:r.siteName,hostRole:"확인필요",consortium:[],consortiumStaff:[]};sites.set(sk,x);state.sites.push(x);}
    const ek=employeeKey(r); if(!staff.has(ek)){const x={id:uid(),employeeId:r.employeeId,name:r.employeeName,job:r.role};staff.set(ek,x);state.staff.push(x);}
    const p={id:r.rowId||uid(),siteId:sites.get(sk).id,staffId:staff.get(ek).id,kind:kindCode(r.assignmentType),start:r.startDate,end:r.endDate,sourceFile:r.sourceFile,sourceSheet:r.sourceSheet,sourceRow:r.sourceRow,originalData:r.originalData,rowRef:r};
    r.rowId=p.id; state.placements.push(p);
  });
}
function calculateAssignmentOverlaps(rows){
  const groups=new Map(),out=[]; rows.filter(r=>!r.excludedReason&&validISO(r.startDate)&&validISO(r.endDate)&&r.endDate>=r.startDate).forEach(r=>{const k=employeeKey(r);if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);});
  groups.forEach((list,key)=>{for(let i=0;i<list.length;i++)for(let j=i+1;j<list.length;j++){const a=list[i],b=list[j];if(siteKey(a.siteName)===siteKey(b.siteName))continue;if(a.startDate<=b.endDate&&b.startDate<=a.endDate)out.push({employeeKey:key,a,b,start:a.startDate>b.startDate?a.startDate:b.startDate,end:a.endDate<b.endDate?a.endDate:b.endDate});}}); return out;
}
function overlapSet(){ const s=new Set(); calculateAssignmentOverlaps(allAssignments).forEach(x=>{s.add(x.a.rowId);s.add(x.b.rowId);});return s; }

async function parseFiles(files,append){
  const list=Array.from(files||[]); if(!list.length)return; if(!append){allAssignments=[];importReport={files:0,sheets:0,rawRows:0,accepted:0,excluded:[],duplicateSource:0,errors:[],details:[]};}
  const seen=new Set(allAssignments.filter(r=>!r.excludedReason).map(sourceFingerprint));
  for(const file of list){ const fd={file:file.name,sheets:0,rawRows:0,accepted:0,excluded:0,error:""}; importReport.files++;
    try{
      if(/\.xls$/i.test(file.name)) throw new Error("구형 XLS 바이너리는 이 브라우저 내장 판독기가 지원하지 않습니다. XLSX 또는 CSV로 저장해 주세요.");
      const wb=await parseWorkbook(file); const sheets=Array.isArray(wb.sheets)?wb.sheets:[];
      for(const sh of sheets){ fd.sheets++;importReport.sheets++; const rows=Array.isArray(sh.rows)?sh.rows:[]; if(!rows.length)continue; const h=findHeader(sh);
        if(h.score<2){importReport.errors.push({file:file.name,sheet:sh.name,error:"헤더를 찾지 못했습니다."});continue;}
        for(let ri=h.index+1;ri<rows.length;ri++){ const vals=rowMap(rows[ri]); if(!Object.values(vals).some(v=>normalizeText(v)!==""))continue; fd.rawRows++;importReport.rawRows++;
          const originalData={}; Object.keys(vals).forEach(i=>{const name=normalizeText(h.headers[i])||("열 "+(Number(i)+1));originalData[name]=vals[i]===undefined||vals[i]===null?"":vals[i];});
          const get=k=>h.map[k]===undefined?"":vals[h.map[k]]; const start=parseDateValue(get("startDate"),wb.date1904),end=parseDateValue(get("endDate"),wb.date1904);
          const r={rowId:uid(),employeeId:normalizeText(get("employeeId")),employeeName:normalizeText(get("employeeName")),siteName:normalizeText(get("siteName")),role:normalizeText(get("role")),startDate:start,endDate:end,assignmentType:kindLabel(get("assignmentType")),sourceFile:file.name,sourceSheet:safe(sh.name,"시트"),sourceRow:ri+1,originalData,dateError:false,excludedReason:"",duplicateSource:false};
          const reasons=[]; if(!r.employeeName&&!r.employeeId)reasons.push("직원 식별값 없음");if(!r.siteName)reasons.push("현장명 없음");if(!start||!end){reasons.push("날짜 확인 필요");r.dateError=true;}else if(end<start){reasons.push("잘못된 기간");r.dateError=true;}
          r.excludedReason=reasons.join(", "); const fp=sourceFingerprint(r); if(!r.excludedReason&&seen.has(fp)){r.duplicateSource=true;r.excludedReason="중복 원본 데이터";importReport.duplicateSource++;}else if(!r.excludedReason)seen.add(fp);
          allAssignments.push(r); if(r.excludedReason){fd.excluded++;importReport.excluded.push(r);}else{fd.accepted++;importReport.accepted++;}
        }
      }
    }catch(e){fd.error=safe(e&&e.message,e);importReport.errors.push({file:file.name,sheet:"",error:fd.error});}
    importReport.details.push(fd);
  }
  ensureLegacyState(); ui.import={step:"done",result:{imported:importReport.accepted,skipped:importReport.excluded.length,dupCount:calculateAssignmentOverlaps(allAssignments).length,dupStaff:[]}}; renderOverview();toast(importReport.accepted+"건 처리 완료","ok");
}

window.handleUpload=function(files){return parseFiles(files,false);};
window.addUploadFiles=function(files){return parseFiles(files,true);};
window.resetImportedData=function(){if(!confirm("불러온 데이터를 모두 초기화할까요?"))return;allAssignments=[];importReport={files:0,sheets:0,rawRows:0,accepted:0,excluded:[],duplicateSource:0,errors:[],details:[]};ensureLegacyState();ui.import={step:"pick",result:null};renderOverview();};
window.dropUpload=function(ev){ev.preventDefault();parseFiles(ev.dataTransfer.files,allAssignments.length>0);};

window.renderUploadPanel=function(){ const host=document.getElementById("ov-upload");if(!host)return;const r=importReport; const has=allAssignments.length>0;
  host.innerHTML='<div class="panel"><div class="panel-head"><h3>엑셀 불러오기</h3><div class="spacer"></div><span class="chip green">브라우저 메모리에서만 처리</span></div><div class="panel-body">'+
  '<div class="upload-zone" ondragover="event.preventDefault()" ondrop="dropUpload(event)" onclick="document.getElementById(\'xlsx-file\').click()"><div class="big">엑셀 파일을 선택하거나 여러 파일을 끌어 놓으세요</div><div class="small muted">지원: .xlsx · .xlsm · .csv / 모든 데이터 시트와 자동 헤더 탐색</div></div>'+
  '<input type="file" multiple id="xlsx-file" accept=".xlsx,.xls,.xlsm,.csv" class="hidden" onchange="'+(has?'addUploadFiles':'handleUpload')+'(this.files)">'+
  '<div class="btn-row" style="margin-top:12px"><button class="btn primary sm" onclick="document.getElementById(\'xlsx-file\').click()">'+(has?'파일 추가':'파일 선택')+'</button>'+(has?'<button class="btn danger sm" onclick="resetImportedData()">불러온 데이터 초기화</button>':'')+'</div>'+
  '<div class="import-note" style="margin-top:12px">엑셀 데이터는 외부로 전송되지 않으며 현재 브라우저 메모리에서만 처리됩니다.</div>'+
  (has?'<div class="cards"><div class="card"><div class="num">'+r.files+'</div><div class="lbl">불러온 파일</div></div><div class="card"><div class="num">'+r.sheets+'</div><div class="lbl">확인한 시트</div></div><div class="card"><div class="num">'+r.rawRows+'</div><div class="lbl">전체 원본 행</div></div><div class="card ok"><div class="num">'+r.accepted+'</div><div class="lbl">정상 배치</div></div><div class="card amber"><div class="num">'+r.excluded.length+'</div><div class="lbl">제외 행</div></div><div class="card"><div class="num">'+r.duplicateSource+'</div><div class="lbl">중복 원본</div></div></div>':'')+
  (r.errors.length?'<div class="small" style="color:var(--red)">'+r.errors.map(x=>esc(x.file+' '+x.sheet+' · '+x.error)).join('<br>')+'</div>':'')+'</div></div>';
};

const oldRenderOverview=renderOverview;
window.handleSearchCompositionStart=function(){composing=true;clearTimeout(searchTimer);};
function applyOverviewSearch(value){
  const q=normalizeText(value).toLowerCase();ui.ovSearch=safe(value);
  document.querySelectorAll('#view-overview .tl[data-staff]').forEach(row=>{row.style.display=!q||normalizeText(row.textContent).toLowerCase().includes(q)?'':'none';});
  let visibleDetails=0;
  document.querySelectorAll('#view-overview .panel-body.tight tbody tr').forEach(row=>{const show=!q||normalizeText(row.textContent).toLowerCase().includes(q);row.style.display=show?'':'none';if(show)visibleDetails++;});
  let empty=document.getElementById('overview-search-empty');
  if(q&&!visibleDetails){if(!empty){empty=document.createElement('div');empty.id='overview-search-empty';empty.className='empty';empty.innerHTML='검색 조건과 일치하는 직원이 없습니다.<br>검색어 또는 필터를 변경해 주세요.';const table=document.querySelector('#view-overview .panel-body.tight table');if(table)table.after(empty);}}
  else if(empty)empty.remove();
}
window.handleSearchCompositionEnd=function(event){composing=false;clearTimeout(searchTimer);applyOverviewSearch(event&&event.target&&event.target.value);};
window.handleOverviewSearch=function(event){const value=safe(event&&event.target&&event.target.value);if(composing||event&&event.isComposing)return;clearTimeout(searchTimer);searchTimer=setTimeout(()=>applyOverviewSearch(value),40);};
window.renderOverview=function(){ oldRenderOverview(); const search=document.getElementById("ov-search");if(search){search.placeholder="직원명, 직원번호 또는 직무";}
  const view=document.getElementById("view-overview");const overlaps=calculateAssignmentOverlaps(allAssignments);const msg=document.createElement("div");msg.className="import-note";msg.style.background=overlaps.length?'var(--red-bg)':'var(--green-bg)';msg.textContent=overlaps.length?("서로 다른 현장 간 중복 배치 "+overlaps.length+"건을 확인해 주세요."):"현재 확인된 중복 배치가 없습니다. 전체 배치 현황은 아래에서 계속 확인할 수 있습니다.";const title=view.querySelector(".section-title");if(title)title.after(msg);
  document.querySelectorAll(".mono").forEach(x=>{x.style.whiteSpace="nowrap";x.style.minWidth="104px";}); document.querySelectorAll(".panel-body.tight").forEach(x=>x.style.overflowX="auto");
};
window.toggleOverlapOnly=function(){const n=calculateAssignmentOverlaps(allAssignments).length;if(!n){ui.ovOverlapOnly=false;toast("현재 확인된 중복 배치가 없습니다. 전체 현황을 계속 표시합니다.","ok");renderOverview();return;}ui.ovOverlapOnly=!ui.ovOverlapOnly;renderOverview();};

function assignmentForPlacement(p){return p&&p.rowRef?p.rowRef:allAssignments.find(r=>r.rowId===p.id);}
const oldStaffById=staffById;window.staffById=function(id){return oldStaffById(id)||{id,name:"[확인 필요]",employeeId:"",job:""};};
const oldSiteById=siteById;window.siteById=function(id){return oldSiteById(id)||{id,name:"[확인 필요]",hostRole:"확인필요",consortium:[]};};
window.fmt=function(d){return validISO(d)?d:"날짜 확인 필요";};
window.statusOf=function(p){const r=assignmentForPlacement(p)||{startDate:p&&p.start,endDate:p&&p.end};const t=statusLabel(r);return {t,cls:t==="배치 중"?"sd-on":t==="배치 예정"?"sd-will":"sd-end"};};

function matchesSearch(r,q){q=normalizeText(q).toLowerCase();if(!q)return true;return [r.employeeName,r.employeeId,r.role].some(v=>normalizeText(v).toLowerCase().includes(q));}
const originalDetailRows=detailRows;window.detailRows=function(filtered){const q=ui.ovSearch;const safeRows=(Array.isArray(filtered)?filtered:[]).filter(p=>{const r=assignmentForPlacement(p);return !r||matchesSearch(r,q);});if(!safeRows.length&&q)return '<tr><td colspan="8" class="empty">검색 조건과 일치하는 직원이 없습니다.<br>검색어 또는 필터를 변경해 주세요.</td></tr>';return originalDetailRows(safeRows);};

window.renderHistory=function(){
  const groups=new Map();allAssignments.filter(r=>!r.excludedReason).forEach(r=>{const k=employeeKey(r);if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);});
  const keys=[...groups.keys()].sort((a,b)=>safe(groups.get(a)[0].employeeName).localeCompare(safe(groups.get(b)[0].employeeName),"ko"));if(!ui.histKey||!groups.has(ui.histKey))ui.histKey=keys[0]||"";
  const options=keys.map(k=>{const r=groups.get(k)[0];return '<option value="'+esc(k)+'"'+(ui.histKey===k?' selected':'')+'>'+esc(r.employeeName||"[확인 필요]")+(r.employeeId?' ('+esc(r.employeeId)+')':'')+'</option>';}).join("");
  const selected=groups.get(ui.histKey)||[];const latest=[...selected].sort((a,b)=>safe(b.startDate).localeCompare(safe(a.startDate)))[0];
  const rows=[...selected].sort((a,b)=>safe(b.startDate).localeCompare(safe(a.startDate))).map(r=>'<tr><td>'+esc(r.siteName||"[확인 필요]")+'</td><td>'+esc(r.role||"[확인 필요]")+'</td><td>'+esc(r.assignmentType||"[확인 필요]")+'</td><td class="mono">'+esc(r.startDate||"날짜 확인 필요")+'</td><td class="mono">'+esc(r.endDate||"날짜 확인 필요")+'</td><td>'+statusLabel(r)+'</td><td>'+esc(r.sourceFile)+'</td><td>'+esc(r.sourceSheet)+'</td></tr>').join("");
  document.getElementById("view-manage").innerHTML='<div class="section-title"><h2>직원별 이력</h2><p>전체 정규화 배치 데이터에서 직원별 모든 이력을 확인합니다.</p></div><div class="sub-tabs">'+['sites','placements','history','export'].map(k=>'<button class="sub-tab '+(ui.mSub===k?'active':'')+'" onclick="ui.mSub=\''+k+'\';renderManage()">'+({sites:'현장 · 직원 관리',placements:'배치 관리',history:'직원별 이력',export:'엑셀 내려받기'})[k]+'</button>').join('')+'</div>'+
  (keys.length?'<div class="panel"><div class="panel-head"><h3>직원 선택</h3><div class="spacer"></div><select onchange="ui.histKey=this.value;renderHistory()">'+options+'</select></div><div class="panel-body"><b>'+esc(latest.employeeName||"[확인 필요]")+'</b> · 사번 '+esc(latest.employeeId||"[확인 필요]")+' · '+esc(latest.role||"[확인 필요]")+' · 전체 '+selected.length+'건 · 최근 '+esc(latest.siteName||"[확인 필요]")+' ('+esc(latest.startDate||"")+' ~ '+esc(latest.endDate||"")+')</div><div class="panel-body tight" style="overflow-x:auto"><table><thead><tr><th>현장명</th><th>직무/직급</th><th>배치 구분</th><th>배치 시작일</th><th>배치 종료일</th><th>진행 상태</th><th>원본 파일</th><th>원본 시트</th></tr></thead><tbody>'+rows+'</tbody></table></div></div>':'<div class="panel"><div class="empty">표시할 직원 배치 이력이 없습니다.<br>엑셀 파일을 불러온 뒤 다시 확인해 주세요.</div></div>');
};
const legacyRenderManage=renderManage;
window.renderManage=function(){
  if(ui.mSub==="history"){ window.renderHistory(); return; }
  return legacyRenderManage();
};

function exportColumns(rows){const originals=[];rows.forEach(r=>Object.keys(r.originalData||{}).forEach(k=>{if(!originals.includes(k)&&!REQUIRED_EXPORT_COLUMNS.includes(k)&&!SOURCE_COLUMNS.includes(k))originals.push(k);}));return REQUIRED_EXPORT_COLUMNS.concat(originals,SOURCE_COLUMNS);}
function exportObject(r,overlaps){const o={"직원번호":r.employeeId,"직원명":r.employeeName,"현장명":r.siteName,"직무/직급":r.role,"배치 시작일":r.startDate,"배치 종료일":r.endDate,"배치 구분":r.assignmentType,"진행 상태":statusLabel(r),"중복 배치 여부":overlaps.has(r.rowId)?"예":"아니오","날짜 오류 여부":r.dateError?"예":"아니오"};Object.entries(r.originalData||{}).forEach(([k,v])=>{if(!(k in o))o[k]=v===undefined||v===null?"":v;});o["원본 파일명"]=r.sourceFile;o["원본 시트명"]=r.sourceSheet;o["원본 행 번호"]=r.sourceRow;return o;}
async function exportRows(rows,name){
  rows=Array.isArray(rows)?rows:[];const cols=exportColumns(rows);const missing=REQUIRED_EXPORT_COLUMNS.filter(x=>!cols.includes(x));const ovs=overlapSet();const objects=rows.map(r=>exportObject(r,ovs));
  if(objects.length!==rows.length||missing.length){toast("엑셀 생성 중 일부 행 또는 항목이 누락되었습니다. 내려받기를 중단하고 데이터를 다시 확인합니다.","err");throw new Error("export validation failed");}
  const matrix=[cols].concat(objects.map(o=>cols.map(c=>{const v=o[c];return v===undefined||v===null?"":v;})));if(matrix.length-1!==rows.length)throw new Error("row count mismatch");
  const widths=cols.map(c=>/날짜|시작일|종료일/.test(c)?14:Math.min(40,Math.max(12,c.length+4)));const bytes=await buildXlsx([{name:"전체 배치 이력",widths,rows:matrix}]);downloadBytes(name+"_"+ymd()+".xlsx",bytes);toast(rows.length+"개 원본 행 · "+cols.length+"개 열을 내려받았습니다.","ok");return {rows:rows.length,columns:cols.length,bytes};
}
window.exportStaff=function(){const ids=[...document.querySelectorAll('.exp-staff:checked')].map(x=>x.value);const keys=new Set(ids.map(id=>{const s=staffById(id);return s.employeeId?"id:"+normalizeText(s.employeeId).toLowerCase():"name:"+normalizeText(s.name).toLowerCase();}));const rows=allAssignments.filter(r=>!r.excludedReason&&keys.has(employeeKey(r)));if(!rows.length){toast("선택한 직원의 원본 배치 행이 없습니다.","err");return;}return exportRows(rows,"직원별_전체원본_배치");};
window.exportSites=function(){const ids=[...document.querySelectorAll('.exp-site:checked')].map(x=>x.value);const names=new Set(ids.map(id=>siteKey(siteById(id).name)));const rows=allAssignments.filter(r=>!r.excludedReason&&names.has(siteKey(r.siteName)));if(!rows.length){toast("선택한 현장의 원본 배치 행이 없습니다.","err");return;}return exportRows(rows,"현장별_전체원본_배치");};

window.save=function(){};try{localStorage.removeItem(STORAGE_KEY);sessionStorage.clear();}catch(e){}
state.settings={apiKey:"",model:"",external:false};window.enhanceWithGemini=async result=>({result,used:false});
document.querySelector('.icon-btn[onclick="openSettings()"]')?.remove();
const style=document.createElement("style");style.textContent='.mono,td:nth-child(5),td:nth-child(6){white-space:nowrap;min-width:108px}.panel-body.tight{overflow-x:auto}.tl-chart{min-width:720px;overflow:visible}.tl{min-width:1020px}.panel:has(.tl){overflow-x:auto}.bar{min-width:3px}.bar.blue{background:var(--blue)}.bar.green{background:var(--green)}.bar.amber{background:var(--amber)}.ov-overlay{background:repeating-linear-gradient(45deg,#e04444,#e04444 7px,#b91c1c 7px,#b91c1c 14px);border:2px solid #991b1b;box-shadow:0 1px 3px rgba(127,29,29,.35);z-index:5;opacity:1;pointer-events:none}.tl-meta span{white-space:nowrap}';document.head.appendChild(style);

// Preserve seed/manual rows in memory so all existing screens continue to work before an upload.
window.__placementTest={normalizeText,normalizeHeader,parseDateValue,findHeader,employeeKey,siteKey,statusLabel,calculateAssignmentOverlaps,exportColumns,exportObject,getRows:()=>allAssignments,getReport:()=>importReport,parseFiles,exportRows};
const selfRows=Array.from({length:148},(_,i)=>({rowId:"self"+i,employeeId:"E"+(i%59),employeeName:"직원"+(i%59),siteName:"현장"+(i%59),role:"직무",startDate:"2025-01-01",endDate:"2025-01-01",assignmentType:"계약배치",originalData:i%2?{연락처:"",비고:false}:{소속:"A",점수:0},sourceFile:"test.xlsx",sourceSheet:"Sheet1",sourceRow:i+2,dateError:false,excludedReason:""}));
const selfCols=exportColumns(selfRows),selfMapped=selfRows.map(r=>exportObject(r,new Set()));
document.documentElement.dataset.placementSelfTest=JSON.stringify({inputRows:selfRows.length,exportRows:selfMapped.length,overlaps:calculateAssignmentOverlaps(selfRows).length,unionColumns:["연락처","비고","소속","점수"].every(x=>selfCols.includes(x)),zeroFalsePreserved:selfMapped.some(x=>x.점수===0)&&selfMapped.some(x=>x.비고===false),requiredColumns:REQUIRED_EXPORT_COLUMNS.every(x=>selfCols.includes(x))});
allAssignments=(Array.isArray(state.placements)?state.placements:[]).map(p=>{const st=staffById(p.staffId),si=siteById(p.siteId);return {rowId:p.id,employeeId:safe(st.employeeId),employeeName:safe(st.name),siteName:safe(si.name),role:safe(st.job),startDate:safe(p.start),endDate:safe(p.end),assignmentType:kindLabel(KIND[p.kind]),sourceFile:safe(p.sourceFile,"직접 입력"),sourceSheet:safe(p.sourceSheet,"배치 관리"),sourceRow:p.sourceRow||0,originalData:p.originalData||{},dateError:!validISO(p.start)||!validISO(p.end)||p.end<p.start,excludedReason:""};});
renderOverview();updatePrivacyBadge();
})();
