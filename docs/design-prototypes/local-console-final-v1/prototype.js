"use strict";

const pages = {
  today: {phase:"OPERATE · TODAY", title:"Daily Control", question:"Where does WFM need to intervene before the next checkpoint?"},
  staffing: {phase:"PLAN · PROTECT", title:"Staff Preparation", question:"Where will published net capacity fail to cover the requirement?"},
  service: {phase:"OPERATE · RECOVER", title:"Intraday Service", question:"Which configured queues explain the selected Management LOB result?"},
  attendance: {phase:"RECONCILE · POST-DAY", title:"Attendance & Schedule Review", question:"Which exact schedule interval remains unsupported after final Activities overlap?"},
  history: {phase:"REVIEW · IMPROVE", title:"Historical Review", question:"Where did demand, capacity and final workforce outcomes diverge from plan?"}
};

function showView(name, push=true){
  if(!pages[name]) name="today";
  document.querySelectorAll(".view").forEach(el=>el.classList.toggle("active",el.dataset.page===name));
  document.querySelectorAll(".cycle-nav button").forEach(el=>el.classList.toggle("active",el.dataset.view===name));
  document.getElementById("page-phase").textContent=pages[name].phase;
  document.getElementById("page-title").textContent=pages[name].title;
  document.getElementById("page-question").textContent=pages[name].question;
  const disabledByPage={service:["planning","staff","leader","agent"],staffing:["leader","agent"]};
  const disabled=new Set(disabledByPage[name]||[]);
  document.querySelectorAll(".filter").forEach(filter=>{
    const isDisabled=disabled.has(filter.dataset.filter);
    filter.classList.toggle("not-applied",isDisabled);
    filter.title=isDisabled?`Not applied to ${pages[name].title}`:"";
  });
  if(push){const url=new URL(location.href);url.searchParams.set("view",name);history.replaceState(null,"",url)}
  window.scrollTo(0,0);
}

function closeMenus(except=null){
  document.querySelectorAll(".filter.open").forEach(el=>{if(el!==except)el.classList.remove("open")});
}

function updateFilterLabel(filter){
  const button=filter.querySelector(".filter-button span");
  const checks=[...filter.querySelectorAll('input[type="checkbox"]')];
  if(!checks.length){const active=filter.querySelector('input[type="radio"]:checked');if(active)button.textContent=active.parentElement.textContent.trim();return}
  const selected=checks.filter(c=>c.checked);
  if(filter.dataset.filter==="lob") button.textContent=selected.length===checks.length?`All ${checks.length} selected`:selected.length?selected.map(c=>c.value).join(", "):"None selected";
  else button.textContent=selected.length===checks.length?button.dataset.all||button.textContent:selected.length===1?selected[0].parentElement.textContent.trim():`${selected.length} selected`;
}

function applyLobFilter(){
  const selected=[...document.querySelectorAll('[data-filter="lob"] input:checked')].map(c=>c.value);
  document.querySelectorAll('tr[data-lob],.heat-row[data-lob],.agent-case[data-lob]').forEach(row=>row.classList.toggle("filtered",!selected.includes(row.dataset.lob)));
  const label=selected.length===4?"all operational LOBs":selected.join(" + ")||"no LOB";
  document.getElementById("scope-copy").textContent=`14 Sep 2026 · ${label} · all active teams`;
}

document.addEventListener("DOMContentLoaded",()=>{
  const name=new URLSearchParams(location.search).get("view")||"today";
  showView(name,false);
  document.querySelectorAll(".cycle-nav button").forEach(btn=>btn.addEventListener("click",()=>showView(btn.dataset.view)));
  document.querySelectorAll(".filter-button").forEach(btn=>btn.addEventListener("click",event=>{event.stopPropagation();const filter=btn.closest(".filter");const willOpen=!filter.classList.contains("open");closeMenus();filter.classList.toggle("open",willOpen)}));
  document.querySelectorAll(".filter-menu").forEach(menu=>menu.addEventListener("click",event=>event.stopPropagation()));
  document.querySelectorAll(".filter-menu input").forEach(input=>input.addEventListener("change",()=>{const filter=input.closest(".filter");updateFilterLabel(filter);if(filter.dataset.filter==="lob")applyLobFilter()}));
  document.querySelector(".clear-filter").addEventListener("click",()=>{document.querySelectorAll('.multi-menu input[type="checkbox"]').forEach(c=>c.checked=true);document.querySelectorAll(".filter").forEach(updateFilterLabel);applyLobFilter()});
  document.addEventListener("click",()=>closeMenus());
  document.addEventListener("keydown",event=>{if(event.key==="Escape")closeMenus()});
});
