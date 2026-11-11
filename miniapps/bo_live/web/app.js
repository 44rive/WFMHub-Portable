'use strict';
const $ = id => document.getElementById(id);
let state = null, csrf = '', dirty = false, busy = false, page = 1, toastTimer;
const pageSize = 40;
const timingLabels = {past:'Past action time',upcoming:'Upcoming',future:'Future',unscheduled:'Unscheduled'};
function node(tag, cls, text) {const n=document.createElement(tag); if(cls)n.className=cls; if(text!==undefined)n.textContent=text; return n;}
function count(n) {return n===null || n===undefined ? '—' : n.toLocaleString();}
function shortGroup(name) {return name.replace(/^\(TOYOTA\)\s*/, '');}
function date(value, full=false) {
  if(!value)return 'No capture';
  const options={timeZone:state?.settings.source_timezone || 'UTC',day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit',hour12:false};
  if(full)options.year='numeric';
  return new Intl.DateTimeFormat('en-GB', options).format(new Date(value));
}
function badge(value) {return node('span',`badge ${value}`,value.toUpperCase());}
function emptyRow(columns,message) {const tr=node('tr'); const td=node('td','empty-cell',message);td.colSpan=columns;tr.append(td);return tr;}
function toast(message) {$('toast').textContent=message;$('toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('toast').hidden=true,6000);}
async function post(path,payload) {
  const response=await fetch(path,{method:'POST',headers:{'Content-Type':'application/json','X-BO-CSRF':csrf},body:JSON.stringify(payload)});
  const result=await response.json();if(!response.ok)throw Error(result.error || 'Request failed.');return result;
}
function navigate() {
  const requested=location.hash.slice(1) || 'overview';const view=['overview','tasks','captures','setup'].includes(requested)?requested:'overview';
  document.querySelectorAll('.view').forEach(n=>n.hidden=n.id!==`view-${view}`);
  document.querySelectorAll('[data-view]').forEach(n=>{if(n.dataset.view===view)n.setAttribute('aria-current','page');else n.removeAttribute('aria-current');});
}
function setOptions(id,values,label) {
  const select=$(id),current=select.value;
  select.replaceChildren(new Option(label,''));
  values.forEach(v=>select.add(new Option(id==='groupFilter'?shortGroup(v):v,v)));
  if(values.includes(current))select.value=current;
}
function priority(row) {return node('span',`priority rank-${row.priority_rank}`,row.priority);}
function scheduleCell(row) {
  const td=node('td');
  td.append(node('span','',row.scheduled_for?date(row.scheduled_for):'Not scheduled'));
  td.append(node('span',`cell-secondary task-timing ${row.timing}`,timingLabels[row.timing]));
  return td;
}
function assignment(row) {const named=!!row.assigned_to.trim();return node('span',named?'':'assignment-empty',named?row.assigned_to:(state.settings.assignment_blank_means_unassigned?'Unassigned':'Blank · not confirmed'));}
function renderMetrics() {
  const cards=[['Listed tasks','total','Captured task rows','☷'],['Past action time','past','At each group’s capture time','◷'],[state.settings.priority_order[0]+' priority','urgent','Exact source priority','↗'],['Unassigned tasks','unassigned',state.settings.assignment_blank_means_unassigned?'Assignment meaning confirmed':'Confirm blank assignment in Setup','○']];
  $('metricCards').replaceChildren(...cards.map(([label,key,note,symbol])=>{const card=node('div','metric-card');const head=node('div','metric-label',label);head.append(node('span','metric-symbol',symbol));card.append(head,node('div','metric-value',count(state.totals[key])),node('div','metric-note',note));return card;}));
  const hasData=state.groups.some(g=>g.observed_at);
  $('emptyState').hidden=hasData;
  const good=state.all_fresh && !state.inbox_attention && !state.inbox_error, banner=$('healthBanner');banner.className=`health-banner${good?'':' attention'}`;
  const content=node('div','health-content');content.append(node('strong','',`${state.fresh_groups} of ${state.group_count} configured groups have fresh evidence.`));
  content.append(node('span','',good?' Counts combine separate group captures.':' Totals cover the last available evidence; missing or old groups need attention.'));
  if(state.inbox_attention)content.append(node('span','', ` ${state.inbox_attention} rejected inbox file(s) need review.`));
  if(state.inbox_error)content.append(node('span','', ` ${state.inbox_error}`));
  if(!state.settings.scope_confirmed)content.append(node('span','', ' Pilot scope · full TOYOTA group coverage is not confirmed.'));
  const link=node('a','text-link','Review captures ↗');link.href='#captures';banner.replaceChildren(node('span','health-icon',good?'✓':'!'),content,link);
  $('evidenceLabel').textContent=good?'Configured groups · captured evidence':'Last available evidence · coverage needs attention';
  $('captureWindow').textContent=state.scan_start?`Capture window: ${date(state.scan_start)} — ${date(state.scan_end)}`:'Awaiting the first complete capture';
  $('scopeBadge').textContent=state.settings.scope_confirmed?'CONFIGURED TOYOTA SCOPE':'PILOT SCOPE';
  const timing=[['past','Past action time'],['upcoming',`Next ${state.settings.upcoming_minutes} minutes`],['future','Later'],['unscheduled','Unscheduled']];
  $('timingBars').replaceChildren(...timing.map(([key,label])=>{const wrap=node('div',`timing-${key}`),line=node('div','timing-line');line.append(node('span','',label),node('strong','',count(state.totals[key])));const bar=node('progress');bar.max=Math.max(state.totals.total || 0,1);bar.value=state.totals[key] || 0;bar.setAttribute('aria-label',label);wrap.append(line,bar);return wrap;}));
}
function renderGroups() {
  $('groupRows').replaceChildren(...state.groups.map(group=>{const tr=node('tr');const name=node('td');const button=node('button','group-link',shortGroup(group.name));button.type='button';button.addEventListener('click',()=>{$('groupFilter').value=group.name;page=1;renderTasks();location.hash='tasks';});name.append(button);tr.append(name);
    for(const key of ['total','past','unassigned'])tr.append(node('td','number',count(group.metrics[key])));
    tr.append(node('td','',group.observed_at?date(group.observed_at):'Awaiting capture'));
    const health=node('td');health.append(badge(group.health));if(group.error)health.title=group.error;tr.append(health);return tr;}));
}
function renderAttention() {
  const rows=state.tasks.slice(0,8).map(row=>{const tr=node('tr'),reference=node('td');reference.append(node('span','case-number',row.case_number),node('span','cell-secondary',shortGroup(row.task_group)));const pri=node('td');pri.append(priority(row));const owner=node('td');owner.append(assignment(row));const evidence=node('td');evidence.append(badge(row.health));tr.append(reference,pri,scheduleCell(row),owner,evidence);return tr;});
  $('attentionRows').replaceChildren(...(rows.length?rows:[emptyRow(5,'Your prioritized tasks will appear after the first complete capture.')]));
}
function renderTasks() {
  if(!state)return;
  const search=$('searchInput').value.trim().toLowerCase(),group=$('groupFilter').value,pri=$('priorityFilter').value,timing=$('timingFilter').value,owner=$('assignmentFilter').value;
  const language=$('languageFilter').value;
  const filtered=state.tasks.filter(row=>(!group || row.task_group===group)&&(!pri || row.priority===pri)&&(!language || row.language===language)&&(!timing || row.timing===timing)&&(!owner || (owner==='blank'?!row.assigned_to.trim():!!row.assigned_to.trim()))&&(!search || [row.case_number,row.task_type,row.assigned_to,row.task_group,row.language].some(v=>v.toLowerCase().includes(search))));
  const pages=Math.max(1,Math.ceil(filtered.length/pageSize));page=Math.min(page,pages);
  $('taskResultCount').textContent=`${count(filtered.length)} captured task${filtered.length===1?'':'s'}`;
  const rows=filtered.slice((page-1)*pageSize,page*pageSize).map(row=>{const tr=node('tr'),groupCell=node('td');groupCell.append(node('span','',shortGroup(row.task_group)),node('span','cell-secondary',row.task_type));const reference=node('td');reference.append(node('span','case-number',row.case_number),node('span','cell-secondary',`Captured row ${row.row_index}`));const priCell=node('td');priCell.append(priority(row));const ownerCell=node('td');ownerCell.append(assignment(row));const health=node('td');health.append(node('span','cell-secondary',date(row.observed_at)),badge(row.health));tr.append(reference,groupCell,priCell,node('td','',row.language || '—'),scheduleCell(row),ownerCell,health);return tr;});
  $('taskRows').replaceChildren(...(rows.length?rows:[emptyRow(7,state.tasks.length?'No tasks match these filters.':'No complete task capture yet. Open Setup to connect PAD.')]));
  const controls=node('div','pagination-controls');for(const [label,delta] of [['← Previous',-1],['Next →',1]]){const b=node('button','button subtle',label);b.type='button';b.disabled=delta<0?page===1:page===pages;b.addEventListener('click',()=>{page+=delta;renderTasks();});controls.append(b);}
  $('taskPagination').replaceChildren(node('span','',`Page ${page} of ${pages} · ${count(filtered.length)} rows`),controls);
}
function renderHistory() {
  $('acknowledgeButton').hidden=!state.inbox_attention;
  const rows=state.history.map(capture=>{const tr=node('tr');tr.append(node('td','',shortGroup(capture.task_group)),node('td','',date(capture.finished_at,true)));const status=node('td');status.append(badge(capture.status));tr.append(status,node('td','number',count(capture.source_total)),node('td','',capture.error || 'Complete capture retained'));return tr;});
  $('historyRows').replaceChildren(...(rows.length?rows:[emptyRow(5,'No captures received yet.')]));
  $('inboxEvents').replaceChildren(...(state.events.length?state.events.map(event=>{const n=node('div','event'),content=node('div');content.append(node('p','',event.message),node('small','',`${date(event.occurred_at,true)} · ${event.filename}`));n.append(badge(event.status),content);return n;}):[node('div','empty-cell','PAD inbox activity will appear here.')]));
}
function populateSettings() {
  if(dirty)return;const s=state.settings;
  $('sourceTimezone').value=s.source_timezone || '';$('groupNames').value=s.groups.join('\n');$('priorityOrder').value=s.priority_order.join('\n');
  $('scopeConfirmed').checked=s.scope_confirmed;$('blankAssignment').checked=s.assignment_blank_means_unassigned;
  $('upcomingMinutes').value=s.upcoming_minutes;$('pollMinutes').value=s.poll_minutes;$('staleMinutes').value=s.stale_minutes;
  $('settingsFields').disabled=state.demo;
}
async function load() {
  if(busy)return;busy=true;
  try {const response=await fetch('/api/state',{cache:'no-store'});if(!response.ok)throw Error('Cannot read the local workspace.');state=await response.json();csrf=state.csrf;$('connectionNotice').hidden=true;
    $('demoNotice').hidden=!state.demo;$('modeBadge').textContent=state.demo?'DEMO':'LOCAL LIVE';$('modeBadge').className=`badge ${state.demo?'incomplete':'fresh'}`;
    $('zoneLabel').textContent=state.settings.source_timezone || 'Time zone not set';$('navCount').textContent=count(state.totals.total);$('inboxPath').textContent=state.inbox_path;
    setOptions('groupFilter',state.settings.groups,'All configured groups');setOptions('priorityFilter',[...new Set([...state.settings.priority_order,...state.tasks.map(r=>r.priority)])],'All priorities');
    setOptions('languageFilter',[...new Set(state.tasks.map(r=>r.language).filter(Boolean))].sort(),'All languages');
    renderMetrics();renderGroups();renderAttention();renderTasks();renderHistory();populateSettings();
  } catch(error){$('connectionNotice').hidden=false;if(!state)toast(error.message);}finally{busy=false;}
}
window.addEventListener('hashchange',navigate);navigate();
document.querySelectorAll('.refresh-button').forEach(b=>b.addEventListener('click',async()=>{await load();if($('connectionNotice').hidden)toast('View refreshed. Capture times show the age of the evidence.');}));
for(const id of ['searchInput','groupFilter','priorityFilter','languageFilter','timingFilter','assignmentFilter'])$(id).addEventListener(id==='searchInput'?'input':'change',()=>{page=1;renderTasks();});
$('clearFilters').addEventListener('click',()=>{for(const id of ['searchInput','groupFilter','priorityFilter','languageFilter','timingFilter','assignmentFilter'])$(id).value='';page=1;renderTasks();});
$('settingsForm').addEventListener('input',()=>dirty=true);
$('settingsForm').addEventListener('submit',async event=>{event.preventDefault();const feedback=$('settingsFeedback');feedback.textContent='Saving…';feedback.className='form-feedback';try{
  const lines=id=>$(id).value.split('\n').map(v=>v.trim()).filter(Boolean);
  await post('/api/settings',{source_timezone:$('sourceTimezone').value.trim() || null,groups:lines('groupNames'),scope_confirmed:$('scopeConfirmed').checked,assignment_blank_means_unassigned:$('blankAssignment').checked,priority_order:lines('priorityOrder'),upcoming_minutes:Number($('upcomingMinutes').value),poll_minutes:Number($('pollMinutes').value),stale_minutes:Number($('staleMinutes').value)});
  dirty=false;feedback.textContent='Settings saved.';await load();toast('Scope and calculation settings saved.');
}catch(error){feedback.textContent=error.message;feedback.className='form-feedback error';}});
$('scanButton').addEventListener('click',async()=>{const button=$('scanButton');button.disabled=true;try{const result=await post('/api/scan',{});await load();toast(result.results.length?`Processed ${result.results.length} inbox file(s). Review their status below.`:'Inbox checked. No new capture files.');}catch(error){toast(error.message);}finally{button.disabled=false;}});
$('acknowledgeButton').addEventListener('click',async()=>{try{await post('/api/acknowledge',{});await load();toast('Inbox issues marked reviewed. Rejected evidence is retained.');}catch(error){toast(error.message);}});
$('copyInbox').addEventListener('click',async()=>{try{await navigator.clipboard.writeText(state.inbox_path);toast('PAD inbox path copied.');}catch{toast('Select and copy the displayed inbox path.');}});
$('captureFile').addEventListener('change',async event=>{const file=event.target.files[0],feedback=$('importFeedback');if(!file)return;feedback.className='form-feedback';feedback.textContent='Validating capture…';try{if(file.size>16*1024*1024)throw Error('Capture exceeds 16 MB.');const payload=JSON.parse(await file.text());const result=await post('/api/import',payload);feedback.textContent=result.duplicate?'This capture was already imported.':'Capture imported.';await load();toast(feedback.textContent);}catch(error){feedback.textContent=error.message;feedback.className='form-feedback error';}finally{event.target.value='';}});
load();setInterval(load,10000);
