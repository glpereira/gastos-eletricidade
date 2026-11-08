// Lógica pura (sem DOM): leitura dos ficheiros da E-Redes, tarifas, períodos e agregações.
// Todas as datas são tratadas como "hora local ingénua" codificada em UTC,
// para que as mudanças de hora não baralhem as contas.

export const MIN15 = 15 * 60 * 1000;
export const DIA = 24 * 60 * 60 * 1000;

const pad = (n) => String(n).padStart(2, '0');

/** Chave "AAAA-MM-DDTHH:MM" a partir de um instante ingénuo (ms). */
export function chave(ms) {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
}

export function deChave(k) {
  return Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10), +k.slice(11, 13), +k.slice(14, 16));
}

export function diaKey(ms) {
  return chave(ms).slice(0, 10);
}

/* ------------------------------------------------------------------ */
/* Leitura do ficheiro                                                 */
/* ------------------------------------------------------------------ */

const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

/** Converte o valor de uma célula de data em {y,m,d} ou null. */
function lerData(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  if (typeof v === 'number') {
    // número de série do Excel (dias desde 1899-12-30)
    const ms = Math.round((v - 25569) * DIA);
    const d = new Date(ms);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }
  const s = String(v).trim();
  let r = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (r) return { y: +r[1], m: +r[2], d: +r[3] };
  r = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (r) return { y: +r[3], m: +r[2], d: +r[1] };
  return null;
}

/** Converte o valor de uma célula de hora em minutos desde a meia-noite, ou null. */
function lerHora(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  if (typeof v === 'number') {
    const frac = v - Math.floor(v);
    if (v >= 1 && frac === 0) return 24 * 60; // "24:00"
    return Math.round(frac * 24 * 60);
  }
  const r = String(v).match(/(\d{1,2}):(\d{2})/);
  return r ? +r[1] * 60 + +r[2] : null;
}

function lerNumero(v) {
  if (typeof v === 'number') return v;
  if (v == null) return NaN;
  const s = String(v).trim().replace(/\s/g, '');
  if (!s) return NaN;
  // "1.234,56" ou "0,123" (formato português) ou "0.123"
  const t = s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s;
  return Number(t);
}

/**
 * Recebe as linhas de uma folha (array de arrays) e devolve
 * { leituras: {chave: kWh}, unidade, coluna, avisos }.
 * A "chave" é a data/hora tal como aparece no ficheiro.
 */
export function lerLinhas(linhas) {
  // 1. encontrar a linha de cabeçalho
  let h = -1, cData = -1, cHora = -1, cValor = -1, cabecalho = null;
  for (let i = 0; i < Math.min(linhas.length, 60); i++) {
    const row = (linhas[i] || []).map(norm);
    const iData = row.findIndex((c) => /^data/.test(c));
    const consumo = row
      .map((c, j) => ({ c, j }))
      .filter(({ c }) => /consumo|ativa|energia|potencia/.test(c) && !/injet|estado|data|hora/.test(c));
    if (iData >= 0 && consumo.length) {
      h = i;
      cData = iData;
      cHora = row.findIndex((c, j) => j !== iData && /^hora/.test(c));
      // preferir "registado"
      const pref = consumo.find(({ c }) => /registad/.test(c)) || consumo[0];
      cValor = pref.j;
      cabecalho = linhas[i][cValor];
      break;
    }
  }
  if (h < 0) {
    throw new Error('Não encontrei as colunas "Data" e "Consumo" no ficheiro. Confirme que exportou o diagrama de carga (consumos de 15 em 15 minutos) do Balcão Digital.');
  }

  const cab = norm(cabecalho);
  // A E-Redes indica potência média (kW) em cada quarto de hora: kWh = kW / 4.
  const unidade = /kwh/.test(cab) ? 'kWh' : /\bw\b|\(w\)/.test(cab) && !/kw/.test(cab) ? 'W' : 'kW';
  const fator = unidade === 'kWh' ? 1 : unidade === 'W' ? 0.25 / 1000 : 0.25;

  const leituras = {};
  let ignoradas = 0;
  for (let i = h + 1; i < linhas.length; i++) {
    const row = linhas[i];
    if (!row) continue;
    const vData = row[cData];
    let data = lerData(vData);
    let min = cHora >= 0 ? lerHora(row[cHora]) : null;
    if (min == null) min = lerHora(vData instanceof Date || typeof vData === 'string' ? vData : null);
    if (min == null && typeof vData === 'number') min = lerHora(vData);
    const val = lerNumero(row[cValor]);
    if (!data || min == null || !Number.isFinite(val)) {
      if (row.some((c) => c != null && c !== '')) ignoradas++;
      continue;
    }
    const ms = Date.UTC(data.y, data.m - 1, data.d) + min * 60000;
    const k = chave(ms);
    // Na mudança da hora de inverno a mesma hora aparece duas vezes: somar.
    leituras[k] = (leituras[k] || 0) + val * fator;
  }

  const n = Object.keys(leituras).length;
  if (!n) throw new Error('O ficheiro não tem leituras válidas.');
  const avisos = [];
  if (ignoradas > 5) avisos.push(`${ignoradas} linhas foram ignoradas por não terem data, hora ou valor válidos.`);
  return { leituras, unidade, coluna: String(cabecalho), avisos };
}

/* ------------------------------------------------------------------ */
/* Tarifas (Portugal continental)                                      */
/* ------------------------------------------------------------------ */

function ultimoDomingo(ano, mes /* 0-11 */) {
  const d = new Date(Date.UTC(ano, mes + 1, 0));
  d.setUTCDate(d.getUTCDate() - d.getUTCDay());
  return Date.UTC(ano, mes, d.getUTCDate());
}

/** Hora legal de verão: do último domingo de março ao último domingo de outubro. */
export function eVerao(ms) {
  const ano = new Date(ms).getUTCFullYear();
  return ms >= ultimoDomingo(ano, 2) && ms < ultimoDomingo(ano, 9);
}

const entre = (m, a, b) => m >= a && m < b;
const hm = (h, m = 0) => h * 60 + m;

/**
 * Classifica o início de um quarto de hora em 'vazio' | 'cheias' | 'ponta' | 'simples' | 'foraVazio'.
 */
export function periodoTarifario(ms, tipo, ciclo) {
  if (tipo === 'simples') return 'simples';
  const d = new Date(ms);
  const m = d.getUTCHours() * 60 + d.getUTCMinutes();
  const dow = d.getUTCDay(); // 0 domingo
  const verao = eVerao(ms);

  if (tipo === 'bihorario') {
    if (ciclo === 'semanal') {
      if (dow === 0) return 'vazio';
      if (dow === 6) {
        const fora = verao
          ? entre(m, hm(9), hm(14)) || entre(m, hm(20), hm(22))
          : entre(m, hm(9, 30), hm(13)) || entre(m, hm(18, 30), hm(22));
        return fora ? 'foraVazio' : 'vazio';
      }
      return m < hm(7) ? 'vazio' : 'foraVazio';
    }
    return m >= hm(22) || m < hm(8) ? 'vazio' : 'foraVazio';
  }

  // tri-horário, ciclo diário
  if (m >= hm(22) || m < hm(8)) return 'vazio';
  const ponta = verao
    ? entre(m, hm(10, 30), hm(13)) || entre(m, hm(19, 30), hm(21))
    : entre(m, hm(9), hm(10, 30)) || entre(m, hm(18), hm(20, 30));
  return ponta ? 'ponta' : 'cheias';
}

export const NOMES_PERIODO = {
  simples: 'Simples',
  vazio: 'Vazio',
  foraVazio: 'Fora de vazio',
  cheias: 'Cheias',
  ponta: 'Ponta',
};

export function periodosDoTipo(tipo) {
  if (tipo === 'simples') return ['simples'];
  if (tipo === 'bihorario') return ['vazio', 'foraVazio'];
  return ['vazio', 'cheias', 'ponta'];
}

/* ------------------------------------------------------------------ */
/* Períodos de faturação (dia 16 → dia 15)                             */
/* ------------------------------------------------------------------ */

export const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const MES_CURTO = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

/**
 * Período de faturação que contém o dia `ms`. É identificado pelo mês em que termina:
 * com início no dia 16, o período "novembro 2026" vai de 16 out a 15 nov.
 */
export function periodoDe(ms, diaInicio = 16) {
  const d = new Date(ms);
  let y = d.getUTCFullYear(), mo = d.getUTCMonth();
  if (diaInicio > 1 && d.getUTCDate() >= diaInicio) mo += 1;
  if (mo > 11) { mo = 0; y += 1; }
  return periodoPorFim(y, mo, diaInicio);
}

export function periodoPorFim(y, mo, diaInicio = 16) {
  y += Math.floor(mo / 12);
  mo = ((mo % 12) + 12) % 12;
  // com início no dia 16, o período termina no dia 15 (fim exclusivo = dia 16)
  const inicio = diaInicio > 1 ? Date.UTC(y, mo - 1, diaInicio) : Date.UTC(y, mo, 1);
  const fim = diaInicio > 1 ? Date.UTC(y, mo, diaInicio) : Date.UTC(y, mo + 1, 1);
  const id = `${y}-${pad(mo + 1)}`;
  const ult = new Date(fim - DIA), ini = new Date(inicio);
  const rotulo = `${ini.getUTCDate()} ${MES_CURTO[ini.getUTCMonth()]} – ${ult.getUTCDate()} ${MES_CURTO[ult.getUTCMonth()]} ${ult.getUTCFullYear()}`;
  const curto = `${MES_CURTO[ult.getUTCMonth()]} ${String(ult.getUTCFullYear()).slice(2)}`;
  return { id, y, mo, inicio, fim, dias: Math.round((fim - inicio) / DIA), rotulo, curto };
}

export function periodoAnterior(p, diaInicio = 16, n = 1) {
  return periodoPorFim(p.y, p.mo - n, diaInicio);
}

/* ------------------------------------------------------------------ */
/* Agregações                                                          */
/* ------------------------------------------------------------------ */

/**
 * Converte o mapa de leituras em lista ordenada de [inicioIntervalo(ms), kWh].
 * Se as horas do ficheiro marcam o fim do intervalo, recua 15 minutos.
 */
export function serie(leituras, horaMarcaFim = true) {
  const desloc = horaMarcaFim ? MIN15 : 0;
  return Object.entries(leituras)
    .map(([k, v]) => [deChave(k) - desloc, v])
    .sort((a, b) => a[0] - b[0]);
}

/** Estatísticas de um intervalo [inicio, fim). */
export function resumir(pontos, inicio, fim, cfg) {
  const periodos = periodosDoTipo(cfg.tipo);
  const porPeriodo = Object.fromEntries(periodos.map((p) => [p, 0]));
  const porDia = {};
  const porDiaPer = {}; // dia -> {periodo: kWh}
  const porHora = new Array(24).fill(0);
  const grelha = {}; // dia -> [24]
  let total = 0, n = 0, maxPonto = null;

  for (const [ms, kwh] of pontos) {
    if (ms < inicio || ms >= fim) continue;
    total += kwh;
    n++;
    const per = periodoTarifario(ms, cfg.tipo, cfg.ciclo);
    porPeriodo[per] = (porPeriodo[per] || 0) + kwh;
    const dk = diaKey(ms);
    porDia[dk] = (porDia[dk] || 0) + kwh;
    const pd = (porDiaPer[dk] ||= {});
    pd[per] = (pd[per] || 0) + kwh;
    const hr = new Date(ms).getUTCHours();
    porHora[hr] += kwh;
    (grelha[dk] ||= new Array(24).fill(0))[hr] += kwh;
    if (!maxPonto || kwh > maxPonto[1]) maxPonto = [ms, kwh];
  }

  const diasComDados = Object.keys(porDia).length;
  const diasTotal = Math.round((fim - inicio) / DIA);
  // Custos fixos contados só nos dias com dados (num período a decorrer = dias já passados).
  const custo = calcularCusto(porPeriodo, diasComDados, cfg);
  return {
    total, n, porPeriodo, porDia, porDiaPer, porHora, grelha, maxPonto,
    diasComDados, diasTotal,
    cobertura: n / (diasTotal * 96),
    mediaHora: porHora.map((v) => (diasComDados ? v / diasComDados : 0)),
    custo,
  };
}

export function fixoPorDia(cfg) {
  return (Number(cfg.potenciaDia) || 0) + (Number(cfg.outrosMes) || 0) / 30.4375;
}

export function calcularCusto(porPeriodo, dias, cfg) {
  let energia = 0;
  for (const [p, kwh] of Object.entries(porPeriodo)) energia += kwh * (Number(cfg.precos?.[p]) || 0);
  const potencia = dias * (Number(cfg.potenciaDia) || 0);
  const outros = (Number(cfg.outrosMes) || 0) * (dias / 30.4375);
  return { energia, potencia, outros, total: energia + potencia + outros };
}

/** Lista dos períodos de faturação que têm dados, do mais antigo para o mais recente. */
export function periodosComDados(pontos, diaInicio = 16) {
  if (!pontos.length) return [];
  const ult = periodoDe(pontos[pontos.length - 1][0], diaInicio);
  let p = periodoDe(pontos[0][0], diaInicio);
  const lista = [];
  while (p.inicio <= ult.inicio) {
    lista.push(p);
    p = periodoPorFim(p.y, p.mo + 1, diaInicio);
  }
  return lista;
}

/** Junta leituras novas às existentes (as novas substituem). */
export function juntar(atual, novas) {
  let novasN = 0;
  const r = { ...atual };
  for (const [k, v] of Object.entries(novas)) {
    if (!(k in r)) novasN++;
    r[k] = v;
  }
  return { leituras: r, novas: novasN };
}
