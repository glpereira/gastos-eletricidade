import {
  DIA, MIN15, lerLinhas, serie, resumir, periodoDe, periodoAnterior, periodosComDados,
  juntar, periodosDoTipo, NOMES_PERIODO, diaKey, periodoTarifario, fixoPorDia,
} from './core.js';
import * as store from './store.js';

/* ------------------------------------------------------------------ */
/* Estado                                                              */
/* ------------------------------------------------------------------ */

const CFG_OMISSAO = {
  tipo: 'simples',
  ciclo: 'diario',
  precos: { simples: 0.16, vazio: 0.10, foraVazio: 0.19, cheias: 0.17, ponta: 0.24 },
  potenciaDia: 0.25,
  outrosMes: 3.0,
  diaInicio: 16,
  horaMarcaFim: true,
  editado: false,
};

const estado = {
  leituras: {},
  cfg: structuredClone(CFG_OMISSAO),
  pontos: [],
  periodos: [],
  idx: -1,
  tab: 'resumo',
  medida: 'kwh',
  filtroDias: 'todos',
  graficos: {},
};

const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];

const nf = (d) => new Intl.NumberFormat('pt-PT', { minimumFractionDigits: d, maximumFractionDigits: d });
const f0 = nf(0), f1 = nf(1), f2 = nf(2), f3 = nf(3);
const eur = new Intl.NumberFormat('pt-PT', { style: 'currency', currency: 'EUR' });
const kwh = (v) => `${(v >= 100 ? f0 : f1).format(v)} kWh`;
const DIAS_SEM = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
const MES_C = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
const dataCurta = (ms) => { const d = new Date(ms); return `${DIAS_SEM[d.getUTCDay()]} ${d.getUTCDate()} ${MES_C[d.getUTCMonth()]}`; };
const horaTxt = (ms) => { const d = new Date(ms); return `${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`; };

function cor(nome) {
  return getComputedStyle(document.documentElement).getPropertyValue(nome).trim();
}
const COR_PERIODO = { simples: '--s1', vazio: '--s1', foraVazio: '--s2', cheias: '--s2', ponta: '--s3' };

/* ------------------------------------------------------------------ */
/* Arranque                                                            */
/* ------------------------------------------------------------------ */

async function iniciar() {
  estado.leituras = await store.ler('leituras', {});
  const cfg = await store.ler('cfg', null);
  if (cfg) estado.cfg = { ...structuredClone(CFG_OMISSAO), ...cfg, precos: { ...CFG_OMISSAO.precos, ...cfg.precos } };
  recalcular();
  estado.idx = estado.periodos.length - 1;
  ligarEventos();
  preencherForm();
  render();
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => render());
}

function recalcular() {
  estado.pontos = serie(estado.leituras, estado.cfg.horaMarcaFim);
  estado.periodos = periodosComDados(estado.pontos, estado.cfg.diaInicio);
  if (estado.idx >= estado.periodos.length || estado.idx < 0) estado.idx = estado.periodos.length - 1;
}

/* ------------------------------------------------------------------ */
/* Eventos                                                             */
/* ------------------------------------------------------------------ */

function ligarEventos() {
  $$('.tabs button').forEach((b) => b.addEventListener('click', () => irPara(b.dataset.tab)));
  $$('[data-ir]').forEach((b) => b.addEventListener('click', () => irPara(b.dataset.ir)));
  $('#perAnt').addEventListener('click', () => { if (estado.idx > 0) { estado.idx--; render(); } });
  $('#perSeg').addEventListener('click', () => { if (estado.idx < estado.periodos.length - 1) { estado.idx++; render(); } });

  $$('#segMedida button').forEach((b) => b.addEventListener('click', () => {
    estado.medida = b.dataset.medida;
    $$('#segMedida button').forEach((x) => x.setAttribute('aria-pressed', x === b));
    render();
  }));
  $$('#segDias button').forEach((b) => b.addEventListener('click', () => {
    estado.filtroDias = b.dataset.dias;
    $$('#segDias button').forEach((x) => x.setAttribute('aria-pressed', x === b));
    render();
  }));

  $('#inFicheiro').addEventListener('change', async (e) => {
    await importar([...e.target.files]);
    e.target.value = '';
  });
  $('#fTarifa').addEventListener('input', guardarForm);
  $('#fTarifa').addEventListener('change', guardarForm);
  $('#btExportar').addEventListener('click', exportarCopia);
  $('#inBackup').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (f) await restaurarCopia(f);
  });
  $('#btApagar').addEventListener('click', async () => {
    if (!confirm('Apagar todas as leituras guardadas neste dispositivo? A tarifa é mantida.')) return;
    estado.leituras = {};
    await store.gravar('leituras', {});
    recalcular();
    render();
  });
  document.addEventListener('click', (e) => { if (!e.target.closest('.mapa td')) esconderDica(); });
}

function irPara(tab) {
  estado.tab = tab;
  $$('.tabs button').forEach((b) => b.setAttribute('aria-selected', b.dataset.tab === tab));
  render();
}

/* ------------------------------------------------------------------ */
/* Importação                                                          */
/* ------------------------------------------------------------------ */

async function lerFicheiro(f) {
  const buf = await f.arrayBuffer();
  const wb = /\.csv$/i.test(f.name)
    ? XLSX.read(new TextDecoder().decode(buf), { type: 'string', raw: true })
    : XLSX.read(buf, { type: 'array' }); // datas como números de série: não dependem do fuso horário
  let ultimoErro;
  for (const nome of wb.SheetNames) {
    const linhas = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, raw: true, defval: null });
    try { return lerLinhas(linhas); } catch (e) { ultimoErro = e; }
  }
  throw ultimoErro || new Error('Ficheiro vazio.');
}

async function importar(ficheiros) {
  const out = $('#resultadoImport');
  out.innerHTML = '';
  for (const f of ficheiros) {
    const linha = document.createElement('p');
    try {
      const r = await lerFicheiro(f);
      const { leituras, novas } = juntar(estado.leituras, r.leituras);
      estado.leituras = leituras;
      await store.gravar('leituras', leituras);
      const ks = Object.keys(r.leituras).sort();
      linha.innerHTML = `<span class="ok">✓ ${esc(f.name)}</span>: ${f0.format(ks.length)} leituras de ${ks[0].slice(0, 10)} a ${ks[ks.length - 1].slice(0, 10)} (${f0.format(novas)} novas). Coluna usada: “${esc(r.coluna)}”, valores em ${r.unidade}${r.unidade === 'kWh' ? '' : ' convertidos para kWh'}.`;
      for (const a of r.avisos) linha.innerHTML += `<br><span class="muted">${esc(a)}</span>`;
    } catch (e) {
      console.error(e);
      linha.innerHTML = `<span class="erro">✗ ${esc(f.name)}</span>: ${esc(e.message)}`;
    }
    out.append(linha);
  }
  store.pedirPersistencia();
  recalcular();
  estado.idx = estado.periodos.length - 1;
  render();
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function exportarCopia() {
  const blob = new Blob([JSON.stringify({ app: 'gastos-eletricidade', versao: 1, cfg: estado.cfg, leituras: estado.leituras })], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `eletricidade-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

async function restaurarCopia(f) {
  try {
    const j = JSON.parse(await f.text());
    if (!j.leituras) throw new Error('Não é uma cópia desta aplicação.');
    estado.leituras = juntar(estado.leituras, j.leituras).leituras;
    if (j.cfg) estado.cfg = { ...estado.cfg, ...j.cfg, precos: { ...estado.cfg.precos, ...j.cfg.precos } };
    await store.gravar('leituras', estado.leituras);
    await store.gravar('cfg', estado.cfg);
    preencherForm();
    recalcular();
    estado.idx = estado.periodos.length - 1;
    render();
    alert('Cópia restaurada.');
  } catch (e) {
    alert(`Não foi possível restaurar: ${e.message}`);
  }
}

/* ------------------------------------------------------------------ */
/* Tarifa                                                              */
/* ------------------------------------------------------------------ */

function preencherForm() {
  const f = $('#fTarifa'), c = estado.cfg;
  f.tipo.value = c.tipo;
  f.ciclo.value = c.tipo === 'trihorario' ? 'diario' : c.ciclo;
  for (const p of ['simples', 'vazio', 'foraVazio', 'cheias', 'ponta']) f[`p_${p}`].value = c.precos[p];
  f.potenciaDia.value = c.potenciaDia;
  f.outrosMes.value = c.outrosMes;
  f.diaInicio.value = c.diaInicio;
  f.horaMarcaFim.checked = c.horaMarcaFim;
  mostrarCamposTarifa();
}

function mostrarCamposTarifa() {
  const f = $('#fTarifa'), t = f.tipo.value;
  $$('#fTarifa [data-mostrar]').forEach((el) => { el.hidden = !el.dataset.mostrar.split(' ').includes(t); });
  const sem = f.querySelector('option[value="semanal"]');
  sem.disabled = t !== 'bihorario';
  if (sem.disabled && f.ciclo.value === 'semanal') f.ciclo.value = 'diario';
}

let tGuardar;
function guardarForm() {
  mostrarCamposTarifa();
  const f = $('#fTarifa'), c = estado.cfg;
  const num = (v, d) => (v === '' || !Number.isFinite(+v) ? d : +v);
  const antes = `${c.diaInicio}|${c.horaMarcaFim}`;
  c.tipo = f.tipo.value;
  c.ciclo = f.ciclo.value;
  for (const p of Object.keys(c.precos)) c.precos[p] = num(f[`p_${p}`].value, c.precos[p]);
  c.potenciaDia = num(f.potenciaDia.value, c.potenciaDia);
  c.outrosMes = num(f.outrosMes.value, c.outrosMes);
  c.diaInicio = Math.min(28, Math.max(1, Math.round(num(f.diaInicio.value, c.diaInicio))));
  c.horaMarcaFim = f.horaMarcaFim.checked;
  c.editado = true;
  if (antes !== `${c.diaInicio}|${c.horaMarcaFim}`) {
    const atual = estado.periodos[estado.idx];
    recalcular();
    if (atual) {
      const p = periodoDe(atual.inicio + 20 * DIA, c.diaInicio);
      const i = estado.periodos.findIndex((x) => x.id === p.id);
      estado.idx = i >= 0 ? i : estado.periodos.length - 1;
    }
  }
  clearTimeout(tGuardar);
  tGuardar = setTimeout(async () => {
    await store.gravar('cfg', c);
    $('#tarifaGuardada').textContent = 'Guardado ✓';
    setTimeout(() => { $('#tarifaGuardada').textContent = ''; }, 1500);
  }, 300);
}

/* ------------------------------------------------------------------ */
/* Render                                                              */
/* ------------------------------------------------------------------ */

function render() {
  const temDados = estado.pontos.length > 0;
  $$('[data-painel]').forEach((s) => { s.hidden = s.dataset.painel !== estado.tab || (!temDados && s.dataset.painel !== 'dados'); });
  $('#vazio').hidden = temDados || estado.tab === 'dados';
  $('#navper').hidden = !temDados || estado.tab === 'dados';
  renderInfoDados();
  if (!temDados) return;

  const p = estado.periodos[estado.idx];
  $('#perRotulo').textContent = p.rotulo;
  $('#perAnt').disabled = estado.idx <= 0;
  $('#perSeg').disabled = estado.idx >= estado.periodos.length - 1;

  if (estado.tab === 'resumo') renderResumo(p);
  if (estado.tab === 'comparar') renderComparar(p);
  if (estado.tab === 'horas') renderHoras(p);
}

function renderInfoDados() {
  const n = estado.pontos.length;
  $('#infoDados').textContent = n
    ? `Guardadas ${f0.format(n)} leituras de 15 minutos, de ${dataCurta(estado.pontos[0][0])} ${new Date(estado.pontos[0][0]).getUTCFullYear()} a ${dataCurta(estado.pontos[n - 1][0])} ${new Date(estado.pontos[n - 1][0]).getUTCFullYear()}.`
    : 'Ainda não há leituras guardadas.';
}

/** Fim efetivo dos dados dentro do período (para períodos ainda a decorrer). */
function limiteDados(p) {
  const ultimo = estado.pontos[estado.pontos.length - 1][0] + MIN15;
  return Math.min(p.fim, Math.max(p.inicio, ultimo));
}

function delta(atual, anterior, el, sufixo) {
  if (!anterior) { el.textContent = 'sem dados para comparar'; el.className = 'delta'; return; }
  const d = (atual - anterior) / anterior;
  const seta = d > 0.005 ? '▲' : d < -0.005 ? '▼' : '=';
  el.textContent = `${seta} ${f0.format(Math.abs(d * 100))}% ${sufixo}`;
  el.className = `delta ${d > 0.005 ? 'subiu' : d < -0.005 ? 'desceu' : ''}`;
}

function renderResumo(p) {
  const cfg = estado.cfg;
  const r = resumir(estado.pontos, p.inicio, p.fim, cfg);
  const lim = limiteDados(p);
  const emCurso = lim < p.fim;

  // comparação justa: se o período está a decorrer, comparar com os mesmos dias do anterior
  const pa = periodoAnterior(p, cfg.diaInicio);
  const ra = resumir(estado.pontos, pa.inicio, emCurso ? pa.inicio + (lim - p.inicio) : pa.fim, cfg);
  const custoAtual = r.custo.total;
  const custoAnt = ra.custo.total;
  const suf = emCurso ? `vs. mesmos dias de ${pa.curto}` : `vs. ${pa.curto}`;

  $('#kTotal').textContent = kwh(r.total);
  delta(r.total, ra.n ? ra.total : 0, $('#kTotalDelta'), suf);
  $('#kCusto').textContent = eur.format(custoAtual);
  delta(custoAtual, ra.n ? custoAnt : 0, $('#kCustoDelta'), suf);
  const media = r.diasComDados ? r.total / r.diasComDados : 0;
  $('#kMedia').textContent = kwh(media);
  $('#kMediaSub').textContent = emCurso
    ? `previsão: ${kwh(media * p.dias)} · ${eur.format(media * p.dias * precoMedio(r) + fixoPorDia(cfg) * p.dias)}`
    : `${r.diasComDados} dias com dados`;
  if (r.maxPonto) {
    $('#kPico').textContent = `${f1.format(r.maxPonto[1] * 4)} kW`;
    $('#kPicoSub').textContent = `${dataCurta(r.maxPonto[0])}, ${horaTxt(r.maxPonto[0])}`;
  }

  const av = [];
  if (emCurso) av.push(`Período ainda a decorrer: há dados até ${dataCurta(lim - MIN15)}. O custo mostrado corresponde aos ${r.diasComDados} dias já registados.`);
  else if (r.cobertura < 0.98) av.push(`Faltam leituras neste período (${f0.format(r.cobertura * 100)}% dos quartos de hora). Importe um ficheiro que cubra estas datas.`);
  if (!cfg.editado) av.push('Os preços são de exemplo: ajuste-os em Dados → Tarifa.');
  $('#avisoCobertura').hidden = !av.length;
  $('#avisoCobertura').textContent = av.join(' ');

  // gráfico por dia (empilhado por período tarifário)
  const periodos = periodosDoTipo(cfg.tipo);
  const dias = [];
  for (let t = p.inicio; t < p.fim; t += DIA) dias.push(t);
  $('#subDia').textContent = periodos.length > 1 ? 'kWh por dia, dividido por período tarifário.' : 'kWh por dia.';
    desenhar('gDia', {
    type: 'bar',
    data: {
      labels: dias.map((t) => new Date(t).getUTCDate()),
      datasets: periodos.map((per) => ({
        label: NOMES_PERIODO[per],
        data: dias.map((t) => r.porDiaPer[diaKey(t)]?.[per] ?? null),
        backgroundColor: cor(COR_PERIODO[per]),
        borderColor: cor('--surface'),
        borderWidth: { top: 0, bottom: periodos.length > 1 ? 2 : 0, left: 0, right: 0 },
        borderRadius: 4,
        borderSkipped: false,
        stack: 'd',
        maxBarThickness: 22,
      })),
    },
    options: opcoesBase({
      empilhado: true,
      legenda: periodos.length > 1,
      yTitulo: 'kWh',
      tooltipTitulo: (it) => dataCurta(dias[it[0].dataIndex]),
      tooltipRodape: (it) => `Total: ${kwh(r.porDia[diaKey(dias[it[0].dataIndex])] || 0)}`,
      linhaMedia: media,
    }),
  });

  // por período tarifário
  $('#subTarifa').textContent = cfg.tipo === 'simples'
    ? 'Tarifa simples: todo o consumo ao mesmo preço. Mude a tarifa em Dados para ver a divisão.'
    : `Tarifa ${cfg.tipo === 'bihorario' ? 'bi-horária' : 'tri-horária'}, ciclo ${cfg.ciclo === 'semanal' ? 'semanal' : 'diário'}.`;
  const tot = r.total || 1;
  $('#barrasTarifa').innerHTML = periodos
    .map((per) => `<span style="width:${(r.porPeriodo[per] / tot) * 100}%;background:${cor(COR_PERIODO[per])}" title="${NOMES_PERIODO[per]}"></span>`)
    .join('');
  $('#tabTarifa').innerHTML = `<tr><th>Período</th><th>kWh</th><th>%</th><th>Custo</th></tr>` + periodos
    .map((per) => `<tr><td><span class="chip" style="background:${cor(COR_PERIODO[per])}"></span>${NOMES_PERIODO[per]}</td><td>${f1.format(r.porPeriodo[per])}</td><td>${f0.format((r.porPeriodo[per] / tot) * 100)}%</td><td>${eur.format(r.porPeriodo[per] * (cfg.precos[per] || 0))}</td></tr>`)
    .join('');
  const c = r.custo;
  $('#custoDetalhe').innerHTML = `
    <div><span>Energia</span><span>${eur.format(c.energia)}</span></div>
    <div><span>Potência (${r.diasComDados} dias)</span><span>${eur.format(c.potencia)}</span></div>
    <div><span>Outros custos fixos</span><span>${eur.format(c.outros)}</span></div>
    <div class="tot"><span>Total estimado</span><span>${eur.format(custoAtual)}</span></div>`;
}

function precoMedio(r) {
  return r.total ? r.custo.energia / r.total : 0;
}

function renderComparar(p) {
  const cfg = estado.cfg;
  const ate = estado.idx;
  const desde = Math.max(0, ate - 12);
  const lista = estado.periodos.slice(desde, ate + 1);
  const res = lista.map((x) => ({ p: x, r: resumir(estado.pontos, x.inicio, x.fim, cfg) }));
  const eEur = estado.medida === 'eur';
  const valor = (o) => (eEur ? o.r.custo.total : o.r.total);
  const s1 = cor('--s1');

  desenhar('gMeses', {
    type: 'bar',
    data: {
      labels: res.map((o) => o.p.curto),
      datasets: [{
        label: eEur ? 'Custo estimado' : 'Consumo',
        data: res.map(valor),
        backgroundColor: res.map((o) => (o.p.id === p.id ? s1 : s1 + '80')),
        borderRadius: 4,
        maxBarThickness: 48,
      }],
    },
    options: opcoesBase({
      legenda: false,
      yTitulo: eEur ? '€' : 'kWh',
      fmt: eEur ? (v) => eur.format(v) : kwh,
      tooltipTitulo: (it) => res[it[0].dataIndex].p.rotulo,
      tooltipRodape: (it) => {
        const o = res[it[0].dataIndex];
        const extra = o.r.cobertura < 0.98 ? ` · dados de ${f0.format(o.r.cobertura * 100)}% do período` : '';
        return `${kwh(o.r.total / Math.max(1, o.r.diasComDados))} por dia${extra}`;
      },
      aoTocar: (i) => { estado.idx = desde + i; render(); },
    }),
  });

  // acumulado
  const pa = periodoAnterior(p, cfg.diaInicio);
  const acum = (per) => {
    const r = resumir(estado.pontos, per.inicio, per.fim, cfg);
    let s = 0;
    const out = [];
    for (let t = per.inicio; t < per.fim; t += DIA) {
      const v = r.porDia[diaKey(t)];
      if (v == null) { out.push(null); continue; }
      s += v;
      out.push(s);
    }
    return out;
  };
  const a1 = acum(p), a0 = acum(pa);
  const n = Math.max(a1.length, a0.length);
  desenhar('gAcum', {
    type: 'line',
    data: {
      labels: Array.from({ length: n }, (_, i) => `dia ${i + 1}`),
      datasets: [
        { label: p.curto, pointStyle: 'line', data: a1, borderColor: cor('--s1'), backgroundColor: cor('--s1'), borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, spanGaps: false },
        { label: pa.curto, pointStyle: 'line', data: a0, borderColor: cor('--s2'), backgroundColor: cor('--s2'), borderWidth: 2, borderDash: [5, 4], pointRadius: 0, pointHoverRadius: 5 },
      ],
    },
    options: opcoesBase({ legenda: true, linhas: true, yTitulo: 'kWh', modoIndice: true, xMaxTicks: 8 }),
  });

  // tabela (inclui comparação com o mesmo período do ano anterior)
  const linhas = [...res].reverse().map((o) => {
    const pAno = periodoAnterior(o.p, cfg.diaInicio, 12);
    const ant = resumir(estado.pontos, pAno.inicio, pAno.fim, cfg);
    // compara a média por dia, para ser justo com períodos incompletos
    const dAno = ant.diasComDados >= 7 && o.r.diasComDados >= 3
      ? (o.r.total / o.r.diasComDados) / (ant.total / ant.diasComDados) - 1 : null;
    const dTxt = dAno == null ? '–' : `<span class="${dAno > 0 ? 'subiu' : 'desceu'}">${dAno > 0 ? '▲' : '▼'} ${f0.format(Math.abs(dAno * 100))}%</span>`;
    return `<tr class="${o.p.id === p.id ? 'atual' : ''}"><td>${o.p.rotulo}${o.r.cobertura < 0.98 ? ' <span class="muted">*</span>' : ''}</td><td>${f1.format(o.r.total)}</td><td>${f1.format(o.r.total / Math.max(1, o.r.diasComDados))}</td><td>${eur.format(o.r.custo.total)}</td><td>${dTxt}</td></tr>`;
  });
  $('#tabMeses').innerHTML = `<tr><th>Período</th><th>kWh</th><th>kWh/dia</th><th>Custo</th><th>kWh/dia vs. ano ant.</th></tr>${linhas.join('')}`
    + `<tr><td colspan="5" class="muted small" style="text-align:left">* período com dados incompletos (custos fixos contados só nos dias com dados)</td></tr>`;
}

function perfil(inicio, fim, filtro) {
  const soma = new Array(24).fill(0);
  const dias = new Set();
  for (const [ms, v] of estado.pontos) {
    if (ms < inicio || ms >= fim) continue;
    const d = new Date(ms), dow = d.getUTCDay();
    const fds = dow === 0 || dow === 6;
    if (filtro === 'uteis' && fds) continue;
    if (filtro === 'fds' && !fds) continue;
    soma[d.getUTCHours()] += v;
    dias.add(diaKey(ms));
  }
  return dias.size ? soma.map((v) => v / dias.size) : null;
}

function renderHoras(p) {
  const cfg = estado.cfg;
  const pa = periodoAnterior(p, cfg.diaInicio);
  const h1 = perfil(p.inicio, p.fim, estado.filtroDias);
  const h0 = perfil(pa.inicio, pa.fim, estado.filtroDias);
  const horas = Array.from({ length: 24 }, (_, h) => `${String(h).padStart(2, '0')}h`);
  const ds = [];
  if (h1) ds.push({ label: p.curto, data: h1, backgroundColor: cor('--s1'), borderRadius: 4, maxBarThickness: 18 });
  if (h0) ds.push({ label: pa.curto, data: h0, backgroundColor: cor('--s2'), borderRadius: 4, maxBarThickness: 18 });
  desenhar('gHora', {
    type: 'bar',
    data: { labels: horas, datasets: ds },
    options: opcoesBase({
      legenda: ds.length > 1,
      yTitulo: 'kWh médio',
      fmt: (v) => `${f2.format(v)} kWh`,
      modoIndice: true,
      tooltipTitulo: (it) => `${horas[it[0].dataIndex]}–${String(it[0].dataIndex + 1).padStart(2, '0')}h`,
    }),
  });

  const dest = $('#destaquesHora');
  if (h1) {
    const max = h1.indexOf(Math.max(...h1));
    const base = Math.min(...h1);
    const totDia = h1.reduce((a, b) => a + b, 0);
    const noite = [18, 19, 20, 21].reduce((a, h) => a + h1[h], 0);
    const vaz = h1.filter((_, h) => h >= 22 || h < 8).reduce((a, b) => a + b, 0);
    dest.innerHTML = `
      <div>Hora com mais consumo<strong>${String(max).padStart(2, '0')}h–${String(max + 1).padStart(2, '0')}h</strong>${f2.format(h1[max])} kWh em média</div>
      <div>Consumo de base (sempre ligado)<strong>${f0.format(base * 1000)} W</strong>a hora mais baixa do dia</div>
      <div>Entre as 18h e as 22h<strong>${f0.format((noite / totDia) * 100)}%</strong>do consumo diário</div>
      <div>Entre as 22h e as 8h<strong>${f0.format((vaz / totDia) * 100)}%</strong>do consumo (horário de vazio)</div>`;
  } else dest.innerHTML = '<div>Sem dados para este filtro.</div>';

  renderMapa(p);

  // picos
  const top = estado.pontos.filter(([ms]) => ms >= p.inicio && ms < p.fim).sort((a, b) => b[1] - a[1]).slice(0, 10);
  $('#tabPicos').innerHTML = `<tr><th>Quando</th><th>kWh</th><th>Potência média</th><th>Período</th></tr>` + top
    .map(([ms, v]) => `<tr><td>${dataCurta(ms)}, ${horaTxt(ms)}–${horaTxt(ms + MIN15)}</td><td>${f3.format(v)}</td><td>${f1.format(v * 4)} kW</td><td>${NOMES_PERIODO[periodoTarifario(ms, cfg.tipo, cfg.ciclo)]}</td></tr>`)
    .join('');
}

function renderMapa(p) {
  const r = resumir(estado.pontos, p.inicio, p.fim, estado.cfg);
  const vals = Object.values(r.grelha).flat().filter((v) => v > 0).sort((a, b) => a - b);
  const q = (f) => vals[Math.min(vals.length - 1, Math.floor(f * vals.length))] ?? 0;
  const limites = [0.15, 0.3, 0.45, 0.6, 0.75, 0.9].map(q);
  const passo = (v) => { let i = 1; for (const l of limites) if (v > l) i++; return i; };

  let h = '<table><thead><tr><th></th>';
  for (let i = 0; i < 24; i++) h += `<th>${i % 3 === 0 ? String(i).padStart(2, '0') : ''}</th>`;
  h += '</tr></thead><tbody>';
  for (let t = p.inicio; t < p.fim; t += DIA) {
    const dk = diaKey(t);
    const d = new Date(t), dow = d.getUTCDay();
    const linha = r.grelha[dk];
    h += `<tr class="${dow === 0 || dow === 6 ? 'fds' : ''}"><th>${DIAS_SEM[dow]} ${d.getUTCDate()}</th>`;
    for (let hr = 0; hr < 24; hr++) {
      if (!linha) { h += '<td class="sem"></td>'; continue; }
      const v = linha[hr];
      h += `<td style="background:var(--seq-${v > 0 ? passo(v) : 0})" data-t="${t + hr * 3600000}" data-v="${v}"></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  const mapa = $('#mapa');
  mapa.innerHTML = h;
  mapa.onpointerover = mapa.onclick = (e) => {
    const td = e.target.closest('td[data-t]');
    if (!td) return esconderDica();
    const t = +td.dataset.t;
    mostrarDica(td, `${dataCurta(t)}, ${horaTxt(t)}–${horaTxt(t + 3600000)}: ${f2.format(+td.dataset.v)} kWh`);
  };
  mapa.onpointerleave = (e) => { if (e.pointerType === 'mouse') esconderDica(); };

  const fmt = (v) => f2.format(v);
  $('#legMapa').innerHTML = `<span>kWh por hora:</span><span class="cx" style="background:var(--seq-1)"></span><span>&lt; ${fmt(limites[0])}</span>`
    + limites.slice(1).map((l, i) => `<span class="cx" style="background:var(--seq-${i + 2})"></span>`).join('')
    + `<span class="cx" style="background:var(--seq-7)"></span><span>&gt; ${fmt(limites[5])}</span>`
    + `<span class="cx" style="margin-left:12px;outline:1px dashed var(--grid);outline-offset:-1px"></span><span>sem dados</span>`;
}

function mostrarDica(el, txt) {
  const d = $('#dica');
  d.textContent = txt;
  d.hidden = false;
  const r = el.getBoundingClientRect();
  const w = d.offsetWidth;
  d.style.left = `${Math.max(8, Math.min(innerWidth - w - 8, r.left + r.width / 2 - w / 2))}px`;
  d.style.top = `${Math.max(8, r.top - d.offsetHeight - 8)}px`;
}
function esconderDica() { $('#dica').hidden = true; }

/* ------------------------------------------------------------------ */
/* Chart.js                                                            */
/* ------------------------------------------------------------------ */

function opcoesBase(o) {
  const ink2 = cor('--ink-2'), muted = cor('--muted'), grid = cor('--grid'), axis = cor('--axis');
  const fmt = o.fmt || kwh;
  const opts = {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 250 },
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: {
        display: !!o.legenda,
        position: 'top',
        align: 'start',
        labels: { color: ink2, usePointStyle: !!o.linhas, pointStyleWidth: 28, boxWidth: 12, boxHeight: 12, useBorderRadius: true, borderRadius: 3, font: { size: 13 } },
      },
      tooltip: {
        backgroundColor: cor('--ink'),
        titleColor: cor('--surface'),
        bodyColor: cor('--surface'),
        footerColor: cor('--surface'),
        padding: 10,
        boxPadding: 4,
        usePointStyle: true,
        callbacks: {
          label: (c) => (c.raw == null ? null : `${c.dataset.label}: ${fmt(c.raw)}`),
          ...(o.tooltipTitulo ? { title: o.tooltipTitulo } : {}),
          ...(o.tooltipRodape ? { footer: o.tooltipRodape } : {}),
        },
        filter: (c) => c.raw != null,
      },
    },
    scales: {
      x: {
        stacked: !!o.empilhado,
        grid: { display: false },
        border: { color: axis },
        ticks: { color: muted, autoSkip: true, maxTicksLimit: o.xMaxTicks || 16, maxRotation: 0 },
      },
      y: {
        stacked: !!o.empilhado,
        beginAtZero: true,
        grid: { color: grid },
        border: { display: false },
        ticks: { color: muted, maxTicksLimit: 6 },
        title: { display: !!o.yTitulo, text: o.yTitulo, color: muted },
      },
    },
  };
  if (o.aoTocar) opts.onClick = (_, els) => { if (els.length) o.aoTocar(els[0].index); };
  if (o.linhaMedia) opts.plugins.linhaMedia = { valor: o.linhaMedia, cor: ink2, fundo: cor('--surface'), texto: `média ${kwh(o.linhaMedia)}` };
  return opts;
}

const pluginLinhaMedia = {
  id: 'linhaMedia',
  afterDatasetsDraw(chart, _a, opt) {
    if (!opt?.valor) return;
    const { ctx, chartArea: ca, scales: { y } } = chart;
    const py = y.getPixelForValue(opt.valor);
    if (py < ca.top || py > ca.bottom) return;
    ctx.save();
    ctx.strokeStyle = opt.cor;
    ctx.setLineDash([4, 4]);
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(ca.left, py);
    ctx.lineTo(ca.right, py);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.font = '12px system-ui, -apple-system, sans-serif';
    ctx.textAlign = 'right';
    const w = ctx.measureText(opt.texto).width;
    ctx.fillStyle = opt.fundo;
    ctx.fillRect(ca.right - w - 6, py - 19, w + 6, 16);
    ctx.fillStyle = opt.cor;
    ctx.fillText(opt.texto, ca.right - 2, py - 6);
    ctx.restore();
  },
};

function desenhar(id, config) {
  if (!window.Chart) return;
  Chart.defaults.font.family = 'system-ui, -apple-system, "Segoe UI", sans-serif';
  estado.graficos[id]?.destroy();
  estado.graficos[id] = new Chart(document.getElementById(id), { ...config, plugins: [pluginLinhaMedia] });
}

iniciar();
