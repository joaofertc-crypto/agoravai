'use strict';

const $ = id => document.getElementById(id);
const fmt = n => Number(n || 0).toLocaleString('pt-BR');
const pct = (n, total) => total ? (100 * n / total).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + '%' : '0,00%';
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const normalize = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const iso = date => date instanceof Date && !Number.isNaN(date.getTime()) ? `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}` : '';
const br = value => value ? value.split('-').reverse().join('/') : '';
let data = [], charts = {}, exportTables = [], sourceName = '';

function parseDate(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === 'number' && Number.isFinite(value)) {
    const d = XLSX.SSF.parse_date_code(value);
    return d ? new Date(d.y, d.m - 1, d.d) : null;
  }
  if (typeof value === 'string') {
    const v = value.trim();
    let m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return validDate(+m[3], +m[2], +m[1]);
    m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return validDate(+m[1], +m[2], +m[3]);
  }
  return null;
}
function validDate(y,m,d) {
  const date = new Date(y,m-1,d);
  return date.getFullYear() === y && date.getMonth() === m-1 && date.getDate() === d ? date : null;
}
function situation(row) {
  const status = normalize(row.status);
  if (status === 'resolvido') return 'resolvido';
  if (status === 'cancelado') return 'cancelado';
  return 'pendente';
}
function isReopened(value) {
  const s = normalize(value);
  return s === 'reaberto' || s === 'reaberta' || s === 'sim' || s === 'true' || s === '1';
}
function slaState(value) {
  const s = normalize(value);
  if (s.includes('nao expirad') || s.includes('sem expir') || s === 'nao' || s === 'false' || s === '0') return 'Em dia';
  if (s.includes('expirad') || s === 'sim' || s === 'true' || s === '1') return 'Expirado';
  return 'Não informado';
}
function counts(rows, field) {
  const map = new Map();
  rows.forEach(row => {
    const key = String(row[field] ?? '').trim() || 'Não informado';
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map].sort((a,b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
}
function makeChart(id, type, labels, values, colors) {
  const canvas = $(id);
  if (!canvas || typeof Chart === 'undefined') return;
  if (charts[id]) charts[id].destroy();
  const palette = colors || ['#247bc0','#f1a33c','#17a398','#8567d5','#e86a72','#7696b7','#a6b5ca'];
  const doughnut = type === 'doughnut';
  charts[id] = new Chart(canvas, {
    type,
    data: { labels, datasets: [{
      data: values,
      backgroundColor: palette,
      borderColor: doughnut ? '#fff' : (type === 'line' ? '#17a398' : 'transparent'),
      borderWidth: doughnut ? 3 : type === 'line' ? 2 : 0,
      borderRadius: type === 'bar' ? 5 : 0,
      tension: .3,
      fill: false,
      pointRadius: type === 'line' ? 3 : 0
    }]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: doughnut, position: 'bottom', labels: { boxWidth: 11, font: { size: 11 } } },
        tooltip: { callbacks: { label: ctx => `${ctx.label || ''}: ${fmt(ctx.parsed?.y ?? ctx.parsed)}` } }
      },
      scales: doughnut ? {} : {
        x: { ticks: { maxRotation: 35, minRotation: 0, font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { precision: 0 } }
      }
    }
  });
}
function table(title, headers, rows, total = null, wide = false) {
  exportTables.push({ title, headers, rows, total });
  return `<article class="panel ${wide ? 'wide' : ''}"><h3>${esc(title)}</h3><div class="scroll"><table><thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.length ? rows.map(row => `<tr>${row.map(v => `<td>${esc(v)}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${headers.length}">Sem dados</td></tr>`}</tbody>${total ? `<tfoot><tr>${total.map(v => `<td>${esc(v)}</td>`).join('')}</tr></tfoot>` : ''}</table></div></article>`;
}
function summary(title, rows, field, label = field) {
  const grouped = counts(rows, field);
  return table(title, [label, 'Total', '%'], grouped.map(([name,n]) => [name,fmt(n),pct(n,rows.length)]), ['Total geral',fmt(rows.length),pct(rows.length,rows.length)]);
}
function kpi(label, value, detail, style = '') {
  return `<div class="kpi ${style}"><span>${esc(label)}</span><strong>${esc(value)}</strong><small>${esc(detail)}</small></div>`;
}
function crossTable(title, rows, rowField, colField, rowLabel, colLabel) {
  const rowNames = [...new Set(rows.map(r => r[rowField] || 'Não informado'))].sort((a,b) => a.localeCompare(b,'pt-BR'));
  const colNames = [...new Set(rows.map(r => r[colField] || 'Não informado'))].sort((a,b) => a.localeCompare(b,'pt-BR'));
  const matrix = new Map();
  const rowTotals = new Map();
  const colTotals = new Map();
  for (const r of rows) {
    const a = r[rowField] || 'Não informado', b = r[colField] || 'Não informado';
    const key = JSON.stringify([a,b]);
    matrix.set(key, (matrix.get(key) || 0) + 1);
    rowTotals.set(a, (rowTotals.get(a) || 0) + 1);
    colTotals.set(b, (colTotals.get(b) || 0) + 1);
  }
  return table(title, [rowLabel,...colNames,'Total geral'], rowNames.map(a => [a,...colNames.map(b => fmt(matrix.get(JSON.stringify([a,b])) || 0)),fmt(rowTotals.get(a))]), ['Total geral',...colNames.map(b => fmt(colTotals.get(b) || 0)),fmt(rows.length)], true);
}
function render() {
  if (!data.length) {
    $('message').textContent = 'Nenhum registro válido encontrado.';
    $('cards').innerHTML = '';
    $('panels').innerHTML = '';
    $('export').disabled = true;
    Object.values(charts).forEach(c => c.destroy());
    charts = {};
    return;
  }
  const start = $('dateStart').value, end = $('dateEnd').value;
  const group = $('group').value, channel = $('channel').value;
  const selectedSituation = $('situation')?.value || '';
  const selectedReopened = $('reopened')?.value || '';
  const category = $('category')?.value || '';
  const subcategory = $('subcategory')?.value || '';
  const search = normalize($('searchTicket')?.value || '');
  if (start && end && start > end) {
    $('message').textContent = 'A data inicial não pode ser posterior à data final.';
    return;
  }
  // Critério temporal único: data de criação, para todas as situações.
  const all = data.filter(r =>
    (!group || r.grupo === group) &&
    (!channel || r.canal === channel) &&
    (!category || r.categoria === category) &&
    (!subcategory || r.subcategoria === subcategory) &&
    (!search || normalize([r.numero,r.solicitante,r.descricao,r.ultimaAcao,r.categoria,r.subcategoria,r.grupo].join(' ')).includes(search)) &&
    (!selectedSituation || r.situacao === selectedSituation) &&
    (!selectedReopened || (selectedReopened === 'sim' ? r.reaberto : !r.reaberto)) &&
    (!start || (r.inicio && iso(r.inicio) >= start)) &&
    (!end || (r.inicio && iso(r.inicio) <= end))
  );
  const resolved = all.filter(r => r.situacao === 'resolvido');
  const canceled = all.filter(r => r.situacao === 'cancelado');
  const pending = all.filter(r => r.situacao === 'pendente');
  const reopened = all.filter(r => r.reaberto);
  const reopenedResolved = reopened.filter(r => r.situacao === 'resolvido');
  const reopenedPending = reopened.filter(r => r.situacao === 'pendente');
  const reopenedCanceled = reopened.filter(r => r.situacao === 'cancelado');
  const expired = pending.filter(r => r.slaSituacao === 'Expirado');

  $('cards').innerHTML =
    kpi('Total no filtro',fmt(all.length),'Chamados únicos do relatório') +
    kpi('Resolvidos',fmt(resolved.length),pct(resolved.length,all.length)+' do total','success') +
    kpi('Pendentes',fmt(pending.length),pct(pending.length,all.length)+' do total','warning') +
    kpi('Cancelados',fmt(canceled.length),pct(canceled.length,all.length)+' do total','danger') +
    kpi('Reabertos',fmt(reopened.length),pct(reopened.length,all.length)+' do total','purple') +
    kpi('Reabertos resolvidos',fmt(reopenedResolved.length),pct(reopenedResolved.length,reopened.length)+' dos reabertos','success') +
    kpi('Reabertos pendentes',fmt(reopenedPending.length),pct(reopenedPending.length,reopened.length)+' dos reabertos','warning') +
    kpi('SLA expirado · pendentes',fmt(expired.length),pct(expired.length,pending.length)+' dos pendentes','danger');

  makeChart('statusChart','doughnut',['Resolvidos','Pendentes','Cancelados'],[resolved.length,pending.length,canceled.length],['#18a38b','#f1a33c','#e86a72']);
  const channels = counts(all,'canal').slice(0,8);
  makeChart('channelChart','bar',channels.map(x=>x[0]),channels.map(x=>x[1]),'#247bc0');
  const groups = counts(all,'grupo').slice(0,10);
  makeChart('groupChart','bar',groups.map(x=>x[0]),groups.map(x=>x[1]),'#8567d5');

  // Últimos 14 dias corridos dentro do período com finalizações registradas.
  const finishDates = resolved.filter(r=>r.fim).map(r=>iso(r.fim)).sort();
  const lastDay = end || finishDates.at(-1) || '';
  let days = [];
  if (lastDay) {
    const [y,m,d] = lastDay.split('-').map(Number);
    for (let i=13;i>=0;i--) {
      const dt = new Date(y,m-1,d-i);
      days.push(iso(dt));
    }
  }
  const perDay = new Map();
  resolved.forEach(r=>{const day=iso(r.fim);if(day)perDay.set(day,(perDay.get(day)||0)+1);});
  makeChart('trendChart','line',days.map(br),days.map(day=>perDay.get(day)||0),'#17a398');

  const slaLabels = ['Em dia','Expirado','Não informado'];
  makeChart('slaChart','doughnut',slaLabels,slaLabels.map(s=>pending.filter(r=>r.slaSituacao===s).length),['#18a38b','#e86a72','#a6b5ca']);
  const pendingStatuses = counts(pending,'status');
  makeChart('openChart','bar',pendingStatuses.map(x=>x[0]),pendingStatuses.map(x=>x[1]),'#f1a33c');
  makeChart('reopenedChart','doughnut',['Reabertos','Sem reabertura'],[reopened.length,all.length-reopened.length],['#8567d5','#a6b5ca']);
  makeChart('reopenedStatusChart','doughnut',['Resolvidos','Pendentes','Cancelados'],[reopenedResolved.length,reopenedPending.length,reopenedCanceled.length],['#18a38b','#f1a33c','#e86a72']);

  exportTables = [];
  let html =
    summary('Todos · Situação atual',all,'situacao','Situação') +
    summary('Todos · Status original',all,'status','Status') +
    summary('Todos · Grupo de atendimento',all,'grupo','Grupo') +
    summary('Todos · Solicitação / Canal',all,'canal','Solicitação / Canal') +
    summary('Todos · Reabertura',all,'reab','Reabertura') +
    summary('Todos · Categoria',all,'categoria','Categoria') +
    summary('Todos · Subcategoria',all,'subcategoria','Subcategoria') +
    crossTable('Categorias × Subcategorias',all,'categoria','subcategoria','Categoria','Subcategoria') +
    summary('Resolvidos · Grupo',resolved,'grupo','Grupo') +
    summary('Pendentes · Grupo',pending,'grupo','Grupo') +
    summary('Cancelados · Grupo',canceled,'grupo','Grupo') +
    summary('Reabertos · Grupo',reopened,'grupo','Grupo') +
    summary('Reabertos · Situação atual',reopened,'situacao','Situação') +
    summary('Pendentes · Status atual',pending,'status','Status') +
    summary('Pendentes · SLA de solução',pending,'slaSituacao','SLA de solução') +
    summary('Pendentes · SLA primeiro atendimento',pending,'slaPrimeiro','SLA primeiro atendimento') +
    summary('Resolvidos · SLA de solução',resolved,'slaSituacao','SLA de solução') +
    crossTable('Pendentes · Grupo × Status',pending,'grupo','status','Grupo','Status') +
    crossTable('Todos · Grupo × Situação',all,'grupo','situacao','Grupo','Situação');

  const lastSeven = [...new Set(resolved.map(r=>iso(r.fim)).filter(Boolean))].sort().slice(-7);
  const resolvedGroups = [...new Set(resolved.map(r=>r.grupo))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  const dayGroupCounts = new Map();
  resolved.forEach(r=>{const d=iso(r.fim);if(d){const key=JSON.stringify([r.grupo,d]);dayGroupCounts.set(key,(dayGroupCounts.get(key)||0)+1);}});
  html += table('Últimas datas de resolução por grupo',['Grupo',...lastSeven.map(br)],resolvedGroups.map(g=>[g,...lastSeven.map(d=>fmt(dayGroupCounts.get(JSON.stringify([g,d]))||0))]),['Total geral',...lastSeven.map(d=>fmt(resolved.filter(r=>iso(r.fim)===d).length))],true);
  $('panels').innerHTML = html;
  renderConversations(all);
  $('message').textContent = `${fmt(data.length)} registros importados de ${sourceName} · ${fmt(all.length)} exibidos · ${fmt(resolved.length + pending.length + canceled.length)} classificados`;
  $('export').disabled = false;
}
function setOptions(id,values) {
  const select = $(id);
  select.replaceChildren(new Option('Todos',''));
  [...new Set(values.filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR')).forEach(value=>select.add(new Option(value,value)));
}
async function loadFile(file) {
  $('message').textContent = 'Lendo arquivo Excel...';
  await new Promise(resolve=>setTimeout(resolve,30));
  if (typeof XLSX === 'undefined') throw new Error('SheetJS não carregou. Verifique sua conexão com a internet.');
  const wb = XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
  let matrix, headerIndex = -1;
  for (const name of wb.SheetNames) {
    const candidate = XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:''});
    const found = candidate.findIndex(row=>row.some(c=>normalize(c)==='nº chamado' || normalize(c)==='no chamado') && row.some(c=>normalize(c)==='nome do status'));
    if (found>=0) {matrix=candidate;headerIndex=found;break;}
  }
  if (headerIndex<0) throw new Error('Cabeçalho não encontrado. São necessárias as colunas Nº Chamado e Nome do Status.');
  const headers = matrix[headerIndex].map(x=>normalize(x));
  const get = (row,...names) => {
    for (const name of names) {
      const index=headers.indexOf(normalize(name));
      if(index>=0 && row[index]!=='' && row[index]!=null) return row[index];
    }
    return '';
  };
  const required = ['Nº Chamado','Nome do Status','Data de Criação','Reaberto'];
  const missing = required.filter(name=>!headers.includes(normalize(name)));
  if (missing.length) throw new Error('Colunas obrigatórias ausentes: '+missing.join(', '));
  const parsed = matrix.slice(headerIndex+1).filter(row=>String(get(row,'Nº Chamado')).trim()).map(row=>{
    const status=String(get(row,'Nome do Status')||'Não informado').trim();
    const reab=String(get(row,'Reaberto')||'Não informado').trim();
    const slaSol=String(get(row,'SLA de Solução Expirado')||'Não informado').trim();
    const item={
      numero:String(get(row,'Nº Chamado')).trim(),
      solicitante:String(get(row,'Nome do Solicitante')||'Não informado').trim(),
      operador:String(get(row,'Nome Completo do Operador')||'Não informado').trim(),
      categoria:String(get(row,'Nome da Categoria')||'Não informado').trim(),
      subcategoria:String(get(row,'Nome da Sub-Categoria')||'Não informado').trim(),
      tipo:String(get(row,'Nome do Tipo de Ocorrência')||'Não informado').trim(),
      descricao:String(get(row,'Descrição')||'').trim(),
      ultimaAcao:String(get(row,'Descrição da Última Ação')||'').trim(),
      canal:String(get(row,'Canal de abertura','Canal de atendimento','Nome da Solicitação')||'Não informado').trim(),
      grupo:String(get(row,'Nome do Grupo')||'Não informado').trim(),
      status, inicio:parseDate(get(row,'Data de Criação')),
      fim:parseDate(get(row,'Data de Finalização')),
      reab, reaberto:isReopened(reab),
      slaSol, slaSituacao:slaState(slaSol),
      slaPrimeiro:String(get(row,'SLA de Primeiro Atendimento Expirado')||'Não informado').trim()
    };
    item.situacao=situation(item);
    return item;
  });
  if (!parsed.length) throw new Error('Nenhum chamado encontrado na planilha.');
  data=parsed;
  sourceName=file.name;
  setOptions('group',data.map(r=>r.grupo));
  setOptions('channel',data.map(r=>r.canal));
  setOptions('category',data.map(r=>r.categoria));
  setOptions('subcategory',data.map(r=>r.subcategoria));
  if($('searchTicket'))$('searchTicket').value='';
  ['dateStart','dateEnd','situation','reopened','category','subcategory'].forEach(id=>{if($(id))$(id).value='';});
  render();
}
$('file').addEventListener('change',async event=>{
  const file=event.target.files?.[0];
  if (!file) return;
  try { await loadFile(file); }
  catch(error) { $('message').textContent='Erro: '+error.message; console.error(error); }
  finally {event.target.value='';}
});
['dateStart','dateEnd','group','channel','situation','reopened','subcategory'].forEach(id=>{
  if($(id))$(id).addEventListener('change',render);
});
$('clear').addEventListener('click',()=>{
  ['dateStart','dateEnd','group','channel','situation','reopened','category','subcategory'].forEach(id=>{if($(id))$(id).value='';});
  if($('searchTicket'))$('searchTicket').value='';
  if(data.length)render();
});
$('export').addEventListener('click',()=>{
  if(!data.length || !exportTables.length || typeof XLSX==='undefined')return;
  const wb=XLSX.utils.book_new();
  exportTables.forEach((t,i)=>{
    const rows=[[t.title],[],t.headers,...t.rows,...(t.total?[t.total]:[])];
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=t.headers.map((h,j)=>({wch:Math.min(45,Math.max(16,String(h).length+3,...t.rows.slice(0,100).map(r=>String(r[j]??'').length+2)))}));
    XLSX.utils.book_append_sheet(wb,ws,`Consulta ${i+1}`);
  });
  XLSX.writeFile(wb,'consultas_chamados.xlsx');
});

// Consultas de conversas e descrições; conteúdo exibido somente como texto seguro.
let visibleTickets=[];
function renderConversations(rows){
  visibleTickets=rows;
  const target=$('conversations');
  if(!target)return;
  const limit=200;
  $('conversationCount').textContent=`${fmt(rows.length)} chamados encontrados · exibindo até ${limit}`;
  target.innerHTML=rows.slice(0,limit).map((r,i)=>`<button type="button" class="ticket" data-ticket="${i}"><strong>${esc(r.numero)}</strong><span>${esc(r.status)} · ${esc(r.categoria)} / ${esc(r.subcategoria)}</span><small>${esc(r.solicitante)} · ${esc(r.grupo)}</small></button>`).join('') || '<p>Nenhum chamado encontrado.</p>';
}
function showTicket(r){
  if(!r)return;
  $('ticketTitle').textContent=`Chamado ${r.numero}`;
  $('ticketMeta').textContent=`${r.status} · ${r.grupo} · ${r.categoria} / ${r.subcategoria} · Solicitante: ${r.solicitante} · Operador: ${r.operador}`;
  $('ticketDescription').textContent=r.descricao||'Sem descrição registrada.';
  $('ticketLastAction').textContent=r.ultimaAcao||'Sem última ação registrada.';
  $('ticketDialog').showModal();
}
$('conversations')?.addEventListener('click',event=>{
  const item=event.target.closest('[data-ticket]');
  if(item)showTicket(visibleTickets[Number(item.dataset.ticket)]);
});
$('closeTicket')?.addEventListener('click',()=>$('ticketDialog').close());
$('searchTicket')?.addEventListener('input',()=>render());
$('category')?.addEventListener('change',()=>{
  const cat=$('category').value;
  setOptions('subcategory',data.filter(r=>!cat||r.categoria===cat).map(r=>r.subcategoria));
  render();
});
