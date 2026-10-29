const $ = id => document.getElementById(id);

const fmt = n => Number(n || 0).toLocaleString('pt-BR');
const pct = (n, total) => total
  ? (100 * n / total).toLocaleString('pt-BR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }) + '%'
  : '0,00%';

const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
}[c]));

const normalize = value => String(value ?? '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .trim()
  .toLowerCase()
  .replace(/\s+/g, '');

/*
 * Fantasia:
 * "000123 - KOP"       -> "KOP"
 * "000456 – EMPRESA A" -> "EMPRESA A"
 * O restante do nome é preservado.
 */
function extrairFantasia(valor) {
  if (valor == null) return '';
  return String(valor)
    .trim()
    .replace(/^\s*\d+\s*[-–—]\s*/, '')
    .trim()
    .replace(/\s+/g, ' ');
}

function chaveFantasia(valor) {
  return normalize(extrairFantasia(valor)).replace(/\s+/g, ' ');
}

const iso = date => date instanceof Date && !Number.isNaN(date.getTime())
  ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
  : '';

const br = value => value ? value.split('-').reverse().join('/') : '';

let data = [];
let charts = {};
let exportTables = [];
let sourceName = '';
let visibleTickets = [];

const activeKpis = new Set();

const MULTI = [
  ['group', r => r.grupo],
  ['channel', r => r.canal],
  ['category', r => r.categoria],
  ['subcategory', r => r.subcategoria],
  ['requester', r => r.solicitante],
  ['occurrenceType', r => r.tipo],
  ['statusName', r => r.status],
  ['operator', r => r.operador],
  ['company', r => extrairFantasia(r.fantasia)],
  ['situation', r => r.situacao],
  ['reopened', r => r.reaberto ? 'sim' : 'nao'],
  ['openChildren', r => r.filhos],
  ['slaFirst', r => r.slaPrimeiroEx ? 'sim' : 'nao'],
  ['slaSolution', r => r.slaSolEx ? 'sim' : 'nao']
];

const DYNAMIC = [
  ['group', r => r.grupo],
  ['channel', r => r.canal],
  ['category', r => r.categoria],
  ['subcategory', r => r.subcategoria],
  ['requester', r => r.solicitante],
  ['occurrenceType', r => r.tipo],
  ['statusName', r => r.status],
  ['operator', r => r.operador],
  ['company', r => extrairFantasia(r.fantasia)]
];

const SIMPLE_INPUTS = [
  'dateStart', 'dateEnd', 'finishStart', 'finishEnd',
  'ticketNumber', 'ageMin', 'ageMax'
];

const KPI_DEFS = {
  res: { group: 'sit', label: 'Resolvidos', test: r => r.situacao === 'resolvido' },
  pen: { group: 'sit', label: 'Pendentes', test: r => r.situacao === 'pendente' },
  can: { group: 'sit', label: 'Cancelados', test: r => r.situacao === 'cancelado' },
  reab: { group: 'reab', label: 'Reabertos', test: r => r.reaberto },
  reabRes: { group: 'reabRes', label: 'Reabertos resolvidos', test: r => r.reaberto && r.situacao === 'resolvido' },
  reabPen: { group: 'reabPen', label: 'Reabertos pendentes', test: r => r.reaberto && r.situacao === 'pendente' },
  sla: { group: 'sla', label: 'SLA expirado · pendentes', test: r => r.situacao === 'pendente' && r.slaSituacao === 'Expirado' },
  sla1: { group: 'sla1', label: 'SLA 1º atendimento expirado', test: r => r.slaPrimeiroEx },
  filhos: { group: 'filhos', label: 'Com filhos em aberto', test: r => r.filhos === 'sim' }
};

function passesKpis(row) {
  const groups = new Map();

  activeKpis.forEach(id => {
    const def = KPI_DEFS[id];
    if (!def) return;
    if (!groups.has(def.group)) groups.set(def.group, []);
    groups.get(def.group).push(def);
  });

  return [...groups.values()].every(list => list.some(def => def.test(row)));
}

/* ------------------------------ Dados ------------------------------ */

function parseDate(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;

  if (typeof value === 'number' && Number.isFinite(value) && typeof XLSX !== 'undefined') {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? validDate(parsed.y, parsed.m, parsed.d) : null;
  }

  if (typeof value === 'string') {
    const v = value.trim();
    let match = v.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
    if (match) return validDate(Number(match[3]), Number(match[2]), Number(match[1]));

    match = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) return validDate(Number(match[1]), Number(match[2]), Number(match[3]));
  }

  return null;
}

function validDate(year, month, day) {
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day ? date : null;
}

function situation(row) {
  const status = normalize(row.status);
  if (status.includes('cancel')) return 'cancelado';
  if (status.includes('resol') || status.includes('conclu') || status.includes('fechad')) return 'resolvido';
  return 'pendente';
}

function isReopened(value) {
  return ['reaberto', 'reaberta', 'sim', 'true', '1', 'yes'].includes(normalize(value));
}

function yesNo(value) {
  return ['sim', 's', 'true', '1', 'yes', 'verdadeiro'].includes(normalize(value)) ? 'sim' : 'nao';
}

function slaState(value) {
  const v = normalize(value);
  if (v.includes('naoexpirad') || v.includes('semexpir') || ['nao', 'false', '0'].includes(v)) return 'Em dia';
  if (v.includes('expirad') || ['sim', 'true', '1'].includes(v)) return 'Expirado';
  return 'Não informado';
}

function counts(rows, field) {
  const map = new Map();
  rows.forEach(row => {
    const key = String(row[field] ?? '').trim() || 'Não informado';
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
}

function selectedValues(id) {
  const select = $(id);
  if (!select) return [];
  if (select.multiple) {
    return Array.from(select.selectedOptions).map(option => option.value).filter(value => value !== '');
  }
  return select.value ? [select.value] : [];
}

function matchesSelected(selected, value) {
  return !selected.length || selected.includes(String(value ?? ''));
}

const inRange = (date, from, to) => {
  if (!from && !to) return true;
  const dateString = iso(date);
  return Boolean(dateString) && (!from || dateString >= from) && (!to || dateString <= to);
};

const numberOrNull = id => {
  const value = $(id)?.value;
  return value === '' || value == null ? null : Number(value);
};

function readFilters() {
  const filters = { selected: {} };
  MULTI.forEach(([id]) => { filters.selected[id] = selectedValues(id); });
  filters.start = $('dateStart')?.value || '';
  filters.end = $('dateEnd')?.value || '';
  filters.finishStart = $('finishStart')?.value || '';
  filters.finishEnd = $('finishEnd')?.value || '';
  filters.number = normalize($('ticketNumber')?.value || '');
  filters.ageMin = numberOrNull('ageMin');
  filters.ageMax = numberOrNull('ageMax');
  filters.search = normalize($('searchTicket')?.value || '');
  return filters;
}

function passesBase(row, filters) {
  for (const [id, getter] of MULTI) {
    const value = getter(row);
    const selected = filters.selected[id] || [];

    if (id === 'company') {
      const currentCompany = chaveFantasia(value);
      if (selected.length && !selected.some(option => chaveFantasia(option) === currentCompany)) return false;
    } else if (!matchesSelected(selected, value)) {
      return false;
    }
  }

  if (filters.number && !normalize(row.numero).includes(filters.number)) return false;
  if (!inRange(row.inicio, filters.start, filters.end)) return false;
  if (!inRange(row.fim, filters.finishStart, filters.finishEnd)) return false;

  if (filters.ageMin !== null || filters.ageMax !== null) {
    if (!Number.isFinite(row.idade)) return false;
    if (filters.ageMin !== null && row.idade < filters.ageMin) return false;
    if (filters.ageMax !== null && row.idade > filters.ageMax) return false;
  }

  if (filters.search && !normalize([
    row.numero, row.solicitante, row.descricao, row.ultimaAcao,
    row.categoria, row.subcategoria, row.grupo, row.canal, row.status, row.operador
  ].join(' ')).includes(filters.search)) return false;

  return true;
}

/* --------------------------- Gráficos/tabelas ----------------------- */

function makeChart(id, type, labels, values, colors) {
  const canvas = $(id);
  if (!canvas || typeof Chart === 'undefined') return;

  if (charts[id]) {
    charts[id].destroy();
    delete charts[id];
  }

  const palette = colors || ['#247bc0', '#f1a33c', '#17a398', '#8567d5', '#e86a72', '#7696b7', '#a6b5ca'];
  const doughnut = type === 'doughnut';

  charts[id] = new Chart(canvas, {
    type,
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: palette,
        borderColor: doughnut ? '#fff' : (type === 'line' ? '#17a398' : 'transparent'),
        borderWidth: doughnut ? 3 : type === 'line' ? 2 : 0,
        borderRadius: type === 'bar' ? 5 : 0,
        tension: 0.3,
        fill: false,
        pointRadius: type === 'line' ? 3 : 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: doughnut, position: 'bottom', labels: { boxWidth: 11, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: context => `${context.label || ''}: ${fmt(context.parsed?.y ?? context.parsed)}`
          }
        }
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
  return `<article class="panel ${wide ? 'wide' : ''}">
    <h3>${esc(title)}</h3>
    <div class="scroll">
      <table>
        <thead><tr>${headers.map(header => `<th>${esc(header)}</th>`).join('')}</tr></thead>
        <tbody>${rows.length
          ? rows.map(row => `<tr>${row.map(value => `<td>${esc(value)}</td>`).join('')}</tr>`).join('')
          : `<tr><td colspan="${headers.length}">Sem dados</td></tr>`}
        </tbody>
        ${total ? `<tfoot><tr>${total.map(value => `<td>${esc(value)}</td>`).join('')}</tr></tfoot>` : ''}
      </table>
    </div>
  </article>`;
}

function summary(title, rows, field, label = field) {
  const grouped = counts(rows, field);
  return table(title, [label, 'Total', '%'],
    grouped.map(([name, amount]) => [name, fmt(amount), pct(amount, rows.length)]),
    ['Total geral', fmt(rows.length), pct(rows.length, rows.length)]);
}

function kpi(id, label, value, detail, style = '', on = false) {
  return `<button type="button" class="kpi ${style} ${on ? 'on' : ''}" data-kpi="${id}" aria-pressed="${on}">
    <span>${esc(label)}</span><strong>${esc(value)}</strong><small>${esc(detail)}</small>
  </button>`;
}

function crossTable(title, rows, rowField, colField, rowLabel, colLabel) {
  const rowNames = [...new Set(rows.map(row => row[rowField] || 'Não informado'))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const colNames = [...new Set(rows.map(row => row[colField] || 'Não informado'))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const matrix = new Map(), rowTotals = new Map(), colTotals = new Map();

  for (const row of rows) {
    const a = row[rowField] || 'Não informado';
    const b = row[colField] || 'Não informado';
    const key = JSON.stringify([a, b]);
    matrix.set(key, (matrix.get(key) || 0) + 1);
    rowTotals.set(a, (rowTotals.get(a) || 0) + 1);
    colTotals.set(b, (colTotals.get(b) || 0) + 1);
  }

  return table(title, [rowLabel, ...colNames, 'Total geral'],
    rowNames.map(a => [a, ...colNames.map(b => fmt(matrix.get(JSON.stringify([a, b])) || 0)), fmt(rowTotals.get(a))]),
    ['Total geral', ...colNames.map(b => fmt(colTotals.get(b) || 0)), fmt(rows.length)], true);
}

/* ------------------------- Comparação mensal ------------------------ */

function monthKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function monthLabel(key) {
  if (!key) return '';
  const [year, month] = key.split('-').map(Number);
  const label = new Date(year, month - 1, 1).toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function makeMonthComparisonChart(rows) {
  const canvas = $('trendChart');
  if (!canvas || typeof Chart === 'undefined') return;

  if (charts.trendChart) {
    charts.trendChart.destroy();
    delete charts.trendChart;
  }

  const resolved = rows.filter(row => row.situacao === 'resolvido' && row.fim instanceof Date && !Number.isNaN(row.fim.getTime()));
  const totals = new Map();

  resolved.forEach(row => {
    const key = monthKey(row.fim);
    if (key) totals.set(key, (totals.get(key) || 0) + 1);
  });

  const months = [...totals.keys()].sort();

  charts.trendChart = new Chart(canvas, {
    type: 'line',
    data: {
      labels: months.map(monthLabel),
      datasets: [{
        label: 'Chamados resolvidos',
        data: months.map(month => totals.get(month) || 0),
        borderColor: '#17a398',
        backgroundColor: 'rgba(23, 163, 152, 0.16)',
        borderWidth: 3,
        tension: 0.3,
        fill: true,
        pointRadius: 4,
        pointHoverRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: true, position: 'bottom' },
        tooltip: { callbacks: { label: context => `${context.dataset.label}: ${fmt(context.parsed.y)}` } }
      },
      scales: {
        x: { title: { display: true, text: 'Mês de resolução' }, ticks: { maxRotation: 45, minRotation: 0 } },
        y: { beginAtZero: true, title: { display: true, text: 'Quantidade de chamados resolvidos' }, ticks: { precision: 0 } }
      }
    }
  });
}

/* ------------------------------ Render ------------------------------ */

function renderActiveKpis() {
  const target = $('activeKpis');
  if (!target) return;
  target.innerHTML = activeKpis.size
    ? 'Indicadores selecionados: ' + [...activeKpis]
      .filter(id => KPI_DEFS[id])
      .map(id => `<button type="button" class="chip" data-remove-kpi="${id}">${esc(KPI_DEFS[id].label)} ×</button>`)
      .join(' ')
    : '';
}

function render() {
  if (!data.length) {
    if ($('message')) $('message').textContent = 'Importe seu relatório Excel.';
    if ($('cards')) $('cards').innerHTML = '';
    if ($('panels')) $('panels').innerHTML = '';
    if ($('activeKpis')) $('activeKpis').innerHTML = '';
    if ($('export')) $('export').disabled = true;
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};
    return;
  }

  const filters = readFilters();

  if ((filters.start && filters.end && filters.start > filters.end) ||
      (filters.finishStart && filters.finishEnd && filters.finishStart > filters.finishEnd)) {
    if ($('message')) {
      $('message').textContent = 'A data inicial não pode ser posterior à data final.';
      $('message').className = 'message error';
    }
    return;
  }

  if (filters.ageMin !== null && filters.ageMax !== null && filters.ageMin > filters.ageMax) {
    if ($('message')) {
      $('message').textContent = 'A idade mínima não pode ser maior que a máxima.';
      $('message').className = 'message error';
    }
    return;
  }

  const base = data.filter(row => passesBase(row, filters));
  const all = base.filter(passesKpis);
  const baseCount = id => base.filter(KPI_DEFS[id].test).length;
  const baseReopened = baseCount('reab');
  const basePending = baseCount('pen');

  if ($('cards')) {
    $('cards').innerHTML =
      kpi('total', 'Total no filtro', fmt(base.length), 'Chamados do relatório', '', activeKpis.size === 0) +
      kpi('res', 'Resolvidos', fmt(baseCount('res')), pct(baseCount('res'), base.length) + ' do total', 'success', activeKpis.has('res')) +
      kpi('pen', 'Pendentes', fmt(basePending), pct(basePending, base.length) + ' do total', 'warning', activeKpis.has('pen')) +
      kpi('can', 'Cancelados', fmt(baseCount('can')), pct(baseCount('can'), base.length) + ' do total', 'danger', activeKpis.has('can')) +
      kpi('reab', 'Reabertos', fmt(baseReopened), pct(baseReopened, base.length) + ' do total', 'purple', activeKpis.has('reab')) +
      kpi('reabRes', 'Reabertos resolvidos', fmt(baseCount('reabRes')), pct(baseCount('reabRes'), baseReopened) + ' dos reabertos', 'success', activeKpis.has('reabRes')) +
      kpi('reabPen', 'Reabertos pendentes', fmt(baseCount('reabPen')), pct(baseCount('reabPen'), baseReopened) + ' dos reabertos', 'warning', activeKpis.has('reabPen')) +
      kpi('sla', 'SLA expirado · pendentes', fmt(baseCount('sla')), pct(baseCount('sla'), basePending) + ' dos pendentes', 'danger', activeKpis.has('sla')) +
      kpi('sla1', 'SLA 1º atend. expirado', fmt(baseCount('sla1')), pct(baseCount('sla1'), base.length) + ' do total', 'danger', activeKpis.has('sla1')) +
      kpi('filhos', 'Com filhos em aberto', fmt(baseCount('filhos')), pct(baseCount('filhos'), base.length) + ' do total', 'purple', activeKpis.has('filhos'));
  }

  renderActiveKpis();

  const resolved = all.filter(row => row.situacao === 'resolvido');
  const canceled = all.filter(row => row.situacao === 'cancelado');
  const pending = all.filter(row => row.situacao === 'pendente');
  const reopened = all.filter(row => row.reaberto);
  const reopenedResolved = reopened.filter(row => row.situacao === 'resolvido');
  const reopenedPending = reopened.filter(row => row.situacao === 'pendente');
  const reopenedCanceled = reopened.filter(row => row.situacao === 'cancelado');

  makeChart('statusChart', 'doughnut', ['Resolvidos', 'Pendentes', 'Cancelados'],
    [resolved.length, pending.length, canceled.length], ['#18a38b', '#f1a33c', '#e86a72']);

  const channelCounts = counts(all, 'canal').slice(0, 8);
  makeChart('channelChart', 'bar', channelCounts.map(item => item[0]), channelCounts.map(item => item[1]), '#247bc0');

  const groupCounts = counts(all, 'grupo').slice(0, 10);
  makeChart('groupChart', 'bar', groupCounts.map(item => item[0]), groupCounts.map(item => item[1]), '#8567d5');

  makeMonthComparisonChart(all);

  const slaLabels = ['Em dia', 'Expirado', 'Não informado'];
  makeChart('slaChart', 'doughnut', slaLabels,
    slaLabels.map(label => pending.filter(row => row.slaSituacao === label).length),
    ['#18a38b', '#e86a72', '#a6b5ca']);

  const pendingStatuses = counts(pending, 'status');
  makeChart('openChart', 'bar', pendingStatuses.map(item => item[0]), pendingStatuses.map(item => item[1]), '#f1a33c');

  makeChart('reopenedChart', 'doughnut', ['Reabertos', 'Sem reabertura'],
    [reopened.length, all.length - reopened.length], ['#8567d5', '#a6b5ca']);

  makeChart('reopenedStatusChart', 'doughnut', ['Resolvidos', 'Pendentes', 'Cancelados'],
    [reopenedResolved.length, reopenedPending.length, reopenedCanceled.length],
    ['#18a38b', '#f1a33c', '#e86a72']);

  exportTables = [];

  let html =
    summary('Todos · Situação atual', all, 'situacao', 'Situação') +
    summary('Todos · Status original', all, 'status', 'Status') +
    summary('Todos · Grupo de atendimento', all, 'grupo', 'Grupo') +
    summary('Todos · Solicitação / Canal', all, 'canal', 'Solicitação / Canal') +
    summary('Todos · Tipo de ocorrência', all, 'tipo', 'Tipo de ocorrência') +
    summary('Todos · Operador', all, 'operador', 'Operador') +
    summary('Todos · Fantasia', all.map(row => ({ ...row, fantasia: extrairFantasia(row.fantasia) })), 'fantasia', 'Fantasia') +
    summary('Todos · Reabertura', all, 'reab', 'Reabertura') +
    summary('Todos · Categoria', all, 'categoria', 'Categoria') +
    summary('Todos · Subcategoria', all, 'subcategoria', 'Subcategoria') +
    crossTable('Categorias × Subcategorias', all, 'categoria', 'subcategoria', 'Categoria', 'Subcategoria') +
    summary('Resolvidos · Grupo', resolved, 'grupo', 'Grupo') +
    summary('Pendentes · Grupo', pending, 'grupo', 'Grupo') +
    summary('Cancelados · Grupo', canceled, 'grupo', 'Grupo') +
    summary('Reabertos · Grupo', reopened, 'grupo', 'Grupo') +
    summary('Reabertos · Situação atual', reopened, 'situacao', 'Situação') +
    summary('Pendentes · Status atual', pending, 'status', 'Status') +
    summary('Pendentes · SLA de solução', pending, 'slaSituacao', 'SLA de solução') +
    summary('Pendentes · SLA primeiro atendimento', pending, 'slaPrimeiro', 'SLA primeiro atendimento') +
    summary('Resolvidos · SLA de solução', resolved, 'slaSituacao', 'SLA de solução') +
    crossTable('Pendentes · Grupo × Status', pending, 'grupo', 'status', 'Grupo', 'Status') +
    crossTable('Todos · Grupo × Situação', all, 'grupo', 'situacao', 'Grupo', 'Situação');

  const lastSeven = [...new Set(resolved.map(row => iso(row.fim)).filter(Boolean))].sort().slice(-7);
  const resolvedGroups = [...new Set(resolved.map(row => row.grupo))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const dayGroupCounts = new Map();

  resolved.forEach(row => {
    const day = iso(row.fim);
    if (day) {
      const key = JSON.stringify([row.grupo, day]);
      dayGroupCounts.set(key, (dayGroupCounts.get(key) || 0) + 1);
    }
  });

  html += table('Últimas datas de resolução por grupo',
    ['Grupo', ...lastSeven.map(br)],
    resolvedGroups.map(group => [group, ...lastSeven.map(day => fmt(dayGroupCounts.get(JSON.stringify([group, day])) || 0))]),
    ['Total geral', ...lastSeven.map(day => fmt(resolved.filter(row => iso(row.fim) === day).length))], true);

  if ($('panels')) $('panels').innerHTML = html;
  renderConversations(all);

  if ($('message')) {
    $('message').textContent = `${fmt(data.length)} registros importados de ${sourceName} · ${fmt(all.length)} exibidos · ${fmt(resolved.length + pending.length + canceled.length)} classificados`;
    $('message').className = 'message success';
  }
  if ($('export')) $('export').disabled = false;
}

/* --------------------------- Opções de filtros ---------------------- */

function setOptions(id, values, preserveSelection = false) {
  const select = $(id);
  if (!select) return;

  const previous = preserveSelection ? selectedValues(id) : [];
  const uniqueMap = new Map();

  values.forEach(value => {
    const original = String(value ?? '').trim();
    if (!original) return;

    const display = id === 'company' ? extrairFantasia(original) : original;
    if (!display) return;

    const key = id === 'company' ? chaveFantasia(display) : display;
    if (!uniqueMap.has(key)) uniqueMap.set(key, display);
  });

  const uniqueValues = [...uniqueMap.values()].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  select.replaceChildren();

  uniqueValues.forEach(value => {
    const option = new Option(value, value);
    option.selected = id === 'company'
      ? previous.some(oldValue => chaveFantasia(oldValue) === chaveFantasia(value))
      : previous.includes(value);
    select.add(option);
  });
}

function updateSubcategoryOptions() {
  const selectedCategories = selectedValues('category');
  const previousSubcategories = selectedValues('subcategory');
  const eligibleRows = data.filter(row => !selectedCategories.length || selectedCategories.includes(row.categoria));

  setOptions('subcategory', eligibleRows.map(row => row.subcategoria), true);

  const valid = new Set(eligibleRows.map(row => row.subcategoria));
  Array.from($('subcategory')?.options || []).forEach(option => {
    option.selected = previousSubcategories.includes(option.value) && valid.has(option.value);
  });
}

function resetFilters() {
  MULTI.forEach(([id]) => {
    const select = $(id);
    if (select) Array.from(select.options).forEach(option => { option.selected = false; });
  });
  SIMPLE_INPUTS.forEach(id => { if ($(id)) $(id).value = ''; });
  if ($('searchTicket')) $('searchTicket').value = '';
  activeKpis.clear();
}

/* ------------------------------ Importação -------------------------- */

async function loadFile(file) {
  if ($('message')) {
    $('message').textContent = 'Lendo arquivo Excel...';
    $('message').className = 'message';
  }

  await new Promise(resolve => setTimeout(resolve, 30));
  if (typeof XLSX === 'undefined') throw new Error('SheetJS não carregou. Verifique sua conexão com a internet.');

  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  let matrix;
  let headerIndex = -1;

  for (const sheetName of workbook.SheetNames) {
    const candidate = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    const found = candidate.findIndex(row =>
      row.some(cell => ['nº chamado', 'no chamado', 'numero do chamado'].includes(normalize(cell))) &&
      row.some(cell => normalize(cell) === 'nome do status')
    );
    if (found >= 0) {
      matrix = candidate;
      headerIndex = found;
      break;
    }
  }

  if (headerIndex < 0) throw new Error('Cabeçalho não encontrado. São necessárias as colunas Nº Chamado e Nome do Status.');

  const headers = matrix[headerIndex].map(normalize);
  const get = (row, ...names) => {
    for (const name of names) {
      const index = headers.indexOf(normalize(name));
      if (index >= 0 && row[index] !== '' && row[index] != null) return row[index];
    }
    return '';
  };

  const required = ['Nº Chamado', 'Nome do Status', 'Data de Criação', 'Reaberto'];
  const missing = required.filter(name => !headers.includes(normalize(name)));
  if (missing.length) throw new Error('Colunas obrigatórias ausentes: ' + missing.join(', '));

  const parsed = matrix.slice(headerIndex + 1)
    .filter(row => String(get(row, 'Nº Chamado', 'Nº do Chamado', 'No Chamado')).trim())
    .map(row => {
      const status = String(get(row, 'Nome do Status') || 'Não informado').trim();
      const reab = String(get(row, 'Reaberto') || 'Não informado').trim();
      const slaSol = String(get(row, 'SLA de Solução Expirado') || 'Não informado').trim();
      const slaPrimeiro = String(get(row, 'SLA de Primeiro Atendimento Expirado') || 'Não informado').trim();

      const idadeBruta = get(row, 'Idade do Chamado — Dias', 'Idade do Chamado - Dias', 'Idade do Chamado (Dias)', 'Idade do Chamado');
      const idade = idadeBruta === '' ? NaN : Number(String(idadeBruta).replace(',', '.'));

      const item = {
        numero: String(get(row, 'Nº Chamado', 'Nº do Chamado', 'No Chamado')).trim(),
        solicitante: String(get(row, 'Nome do Solicitante') || 'Não informado').trim(),
        operador: String(get(row, 'Nome Completo do Operador', 'Nome do Operador') || 'Não informado').trim(),
        categoria: String(get(row, 'Nome da Categoria') || 'Não informado').trim(),
        subcategoria: String(get(row, 'Nome da Sub-Categoria', 'Nome da Subcategoria') || 'Não informado').trim(),
        tipo: String(get(row, 'Nome do Tipo de Ocorrência', 'Tipo de Ocorrência') || 'Não informado').trim(),
        fantasia: String(get(row, 'Fantasia') || 'Não informado').trim(),
        descricao: String(get(row, 'Descrição') || '').trim(),
        ultimaAcao: String(get(row, 'Descrição da Última Ação') || '').trim(),
        canal: String(get(row, 'Canal de abertura', 'Canal de atendimento', 'Nome da Solicitação') || 'Não informado').trim(),
        grupo: String(get(row, 'Nome do Grupo') || 'Não informado').trim(),
        status,
        inicio: parseDate(get(row, 'Data de Criação')),
        fim: parseDate(get(row, 'Data de Finalização')),
        idade: Number.isFinite(idade) ? idade : null,
        reab,
        reaberto: isReopened(reab),
        filhos: yesNo(get(row, 'Possui Chamados Filhos em Aberto')),
        slaSol,
        slaSituacao: slaState(slaSol),
        slaSolEx: slaState(slaSol) === 'Expirado',
        slaPrimeiro,
        slaPrimeiroEx: slaState(slaPrimeiro) === 'Expirado'
      };

      item.situacao = situation(item);
      return item;
    });

  if (!parsed.length) throw new Error('Nenhum chamado encontrado na planilha.');

  data = parsed;
  sourceName = file.name;

  DYNAMIC.forEach(([id, getter]) => setOptions(id, data.map(getter)));
  resetFilters();
  updateSubcategoryOptions();
  render();
}

/* -------------------------------- Eventos ---------------------------- */

$('file')?.addEventListener('change', async event => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    await loadFile(file);
  } catch (error) {
    if ($('message')) {
      $('message').textContent = 'Erro: ' + error.message;
      $('message').className = 'message error';
    }
    console.error(error);
  } finally {
    event.target.value = '';
  }
});

[...MULTI.map(([id]) => id), ...SIMPLE_INPUTS].forEach(id => {
  const element = $(id);
  if (!element) return;
  const eventName = element.tagName === 'SELECT' || element.type === 'date' ? 'change' : 'input';

  element.addEventListener(eventName, () => {
    if (id === 'category') updateSubcategoryOptions();
    render();
  });
});

$('searchTicket')?.addEventListener('input', render);

$('cards')?.addEventListener('click', event => {
  const button = event.target.closest('[data-kpi]');
  if (!button) return;

  const id = button.dataset.kpi;
  if (id === 'total') activeKpis.clear();
  else if (activeKpis.has(id)) activeKpis.delete(id);
  else activeKpis.add(id);

  render();
});

$('activeKpis')?.addEventListener('click', event => {
  const chip = event.target.closest('[data-remove-kpi]');
  if (!chip) return;
  activeKpis.delete(chip.dataset.removeKpi);
  render();
});

$('clear')?.addEventListener('click', () => {
  resetFilters();
  updateSubcategoryOptions();
  if (data.length) render();
});

$('export')?.addEventListener('click', () => {
  if (!data.length || !exportTables.length || typeof XLSX === 'undefined') return;

  const workbook = XLSX.utils.book_new();
  exportTables.forEach((item, index) => {
    const rows = [[item.title], [], item.headers, ...item.rows, ...(item.total ? [item.total] : [])];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);

    worksheet['!cols'] = item.headers.map((header, columnIndex) => ({
      wch: Math.min(45, Math.max(
        16,
        String(header).length + 3,
        ...item.rows.slice(0, 100).map(row => String(row[columnIndex] ?? '').length + 2)
      ))
    }));

    XLSX.utils.book_append_sheet(workbook, worksheet, `Consulta ${index + 1}`);
  });

  XLSX.writeFile(workbook, 'consultas_chamados.xlsx');
});

/* ----------------------------- Chamados ----------------------------- */

function renderConversations(rows) {
  visibleTickets = rows;
  const target = $('conversations');
  if (!target) return;

  const limit = 200;
  if ($('conversationCount')) {
    $('conversationCount').textContent = `${fmt(rows.length)} chamados encontrados · exibindo até ${limit}`;
  }

  target.innerHTML = rows.slice(0, limit).map((row, index) => `
    <button type="button" class="ticket" data-ticket="${index}">
      <strong>${esc(row.numero)}</strong>
      <span>${esc(row.status)} · ${esc(row.categoria)} / ${esc(row.subcategoria)}</span>
      <small>${esc(row.solicitante)} · ${esc(row.grupo)}</small>
    </button>
  `).join('') || '<p class="empty">Nenhum chamado encontrado.</p>';
}

function showTicket(row) {
  if (!row) return;
  if ($('ticketTitle')) $('ticketTitle').textContent = `Chamado ${row.numero}`;

  if ($('ticketMeta')) {
    $('ticketMeta').textContent = `${row.status} · ${row.grupo} · ${row.categoria} / ${row.subcategoria} · Solicitante: ${row.solicitante} · Operador: ${row.operador}`;
  }
  if ($('ticketDescription')) $('ticketDescription').textContent = row.descricao || 'Sem descrição registrada.';
  if ($('ticketLastAction')) $('ticketLastAction').textContent = row.ultimaAcao || 'Sem última ação registrada.';
  $('ticketDialog')?.showModal();
}

$('conversations')?.addEventListener('click', event => {
  const button = event.target.closest('[data-ticket]');
  if (button) showTicket(visibleTickets[Number(button.dataset.ticket)]);
});

$('closeTicket')?.addEventListener('click', () => $('ticketDialog')?.close());
