// Executar com: node --test tests/
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  lerLinhas, serie, resumir, periodoDe, periodoPorFim, periodoAnterior,
  periodoTarifario, eVerao, periodosComDados, juntar,
} from '../js/core.js';

const U = (y, m, d, h = 0, mi = 0) => Date.UTC(y, m - 1, d, h, mi);

test('período de faturação vai do dia 16 ao dia 15', () => {
  const p = periodoDe(U(2026, 11, 8));
  assert.equal(p.id, '2026-11');
  assert.equal(p.inicio, U(2026, 10, 16));
  assert.equal(p.fim, U(2026, 11, 16));
  assert.equal(p.dias, 31);
  assert.equal(p.rotulo, '16 out – 15 nov 2026');
  assert.equal(periodoDe(U(2026, 11, 16)).id, '2026-12');
  assert.equal(periodoDe(U(2026, 12, 20)).id, '2027-01');
  assert.equal(periodoDe(U(2026, 1, 3)).inicio, U(2025, 12, 16));
  assert.equal(periodoAnterior(periodoPorFim(2026, 0)).id, '2025-12');
  assert.equal(periodoPorFim(2026, -1).id, '2025-12');
});

test('dia de início configurável (1 = mês civil)', () => {
  const p = periodoDe(U(2026, 2, 10), 1);
  assert.equal(p.inicio, U(2026, 2, 1));
  assert.equal(p.fim, U(2026, 3, 1));
  assert.equal(p.dias, 28);
});

test('horário de verão', () => {
  assert.equal(eVerao(U(2026, 3, 28)), false);
  assert.equal(eVerao(U(2026, 3, 29)), true); // último domingo de março
  assert.equal(eVerao(U(2026, 10, 24)), true);
  assert.equal(eVerao(U(2026, 10, 25)), false); // último domingo de outubro
});

test('períodos tarifários', () => {
  assert.equal(periodoTarifario(U(2026, 11, 9, 7, 45), 'bihorario', 'diario'), 'vazio');
  assert.equal(periodoTarifario(U(2026, 11, 9, 8, 0), 'bihorario', 'diario'), 'foraVazio');
  assert.equal(periodoTarifario(U(2026, 11, 9, 22, 0), 'bihorario', 'diario'), 'vazio');
  // semanal: dia útil vazio só até às 7h; domingo sempre vazio
  assert.equal(periodoTarifario(U(2026, 11, 9, 7, 30), 'bihorario', 'semanal'), 'foraVazio');
  assert.equal(periodoTarifario(U(2026, 11, 8, 12, 0), 'bihorario', 'semanal'), 'vazio');
  assert.equal(periodoTarifario(U(2026, 11, 7, 10, 0), 'bihorario', 'semanal'), 'foraVazio'); // sábado inverno
  assert.equal(periodoTarifario(U(2026, 11, 7, 14, 0), 'bihorario', 'semanal'), 'vazio');
  // tri-horário inverno: ponta 9h-10h30 e 18h-20h30
  assert.equal(periodoTarifario(U(2026, 11, 9, 9, 15), 'trihorario', 'diario'), 'ponta');
  assert.equal(periodoTarifario(U(2026, 11, 9, 10, 30), 'trihorario', 'diario'), 'cheias');
  assert.equal(periodoTarifario(U(2026, 11, 9, 20, 15), 'trihorario', 'diario'), 'ponta');
  // verão: ponta 10h30-13h e 19h30-21h
  assert.equal(periodoTarifario(U(2026, 7, 9, 9, 15), 'trihorario', 'diario'), 'cheias');
  assert.equal(periodoTarifario(U(2026, 7, 9, 12, 45), 'trihorario', 'diario'), 'ponta');
  assert.equal(periodoTarifario(U(2026, 7, 9, 3, 0), 'trihorario', 'diario'), 'vazio');
});

test('lê folha no formato da E-Redes (kW por quarto de hora)', () => {
  const linhas = [
    ['Código do Ponto de Entrega', 'PT0002000000000000XX'],
    [],
    ['Data', 'Hora', 'Consumo registado (kW)', 'Estado'],
    ['2026/10/16', '00:15', '0,400', 'Real'],
    ['2026/10/16', '00:30', 0.8, 'Real'],
    ['16/10/2026', '00:45', '1.2', 'Real'],
    ['', '', '', ''],
  ];
  const r = lerLinhas(linhas);
  assert.equal(r.unidade, 'kW');
  assert.deepEqual(Object.keys(r.leituras), ['2026-10-16T00:15', '2026-10-16T00:30', '2026-10-16T00:45']);
  assert.ok(Math.abs(r.leituras['2026-10-16T00:15'] - 0.1) < 1e-9);
  const s = serie(r.leituras, true);
  assert.equal(s[0][0], U(2026, 10, 16, 0, 0)); // fim do intervalo -> início
  assert.ok(Math.abs(s.reduce((a, [, v]) => a + v, 0) - 0.6) < 1e-9);
});

test('lê kWh, datas do Excel e coluna Data/Hora única', () => {
  const r = lerLinhas([
    ['Data/Hora', 'Consumo (kWh)'],
    [new Date(2026, 9, 16, 0, 15), 0.25],
    [new Date(2026, 9, 16, 0, 30), 0.5],
  ]);
  assert.equal(r.unidade, 'kWh');
  assert.equal(r.leituras['2026-10-16T00:30'], 0.5);
});

test('prefere "registado" a "medido" e falha sem colunas', () => {
  const r = lerLinhas([['Data', 'Hora', 'Consumo medido (kW)', 'Consumo registado (kW)'], ['2026-10-16', '01:00', 4, 2]]);
  assert.equal(r.leituras['2026-10-16T01:00'], 0.5);
  assert.throws(() => lerLinhas([['Nome', 'Morada'], ['a', 'b']]), /Data/);
});

test('resumo, custo e lista de períodos', () => {
  const leituras = {};
  // 1 kWh por hora (0,25 por quarto de hora) de 16 out a 15 nov
  for (let t = U(2026, 10, 16); t < U(2026, 11, 16); t += 15 * 60000) {
    leituras[new Date(t + 15 * 60000).toISOString().slice(0, 16)] = 0.25;
  }
  const pts = serie(leituras, true);
  const p = periodoDe(U(2026, 11, 1));
  const cfg = { tipo: 'bihorario', ciclo: 'diario', precos: { vazio: 0.1, foraVazio: 0.2 }, potenciaDia: 0.3, outrosMes: 0 };
  const r = resumir(pts, p.inicio, p.fim, cfg);
  assert.ok(Math.abs(r.total - 31 * 24) < 1e-6);
  assert.ok(Math.abs(r.porPeriodo.vazio - 31 * 10) < 1e-6);
  assert.ok(Math.abs(r.porPeriodo.foraVazio - 31 * 14) < 1e-6);
  assert.equal(r.diasComDados, 31);
  assert.ok(Math.abs(r.custo.total - (310 * 0.1 + 434 * 0.2 + 31 * 0.3)) < 1e-6);
  assert.ok(Math.abs(r.mediaHora[3] - 1) < 1e-9);
  assert.deepEqual(periodosComDados(pts).map((x) => x.id), ['2026-11']);
});

test('juntar substitui repetidas e conta novas', () => {
  const r = juntar({ a: 1, b: 2 }, { b: 3, c: 4 });
  assert.deepEqual(r.leituras, { a: 1, b: 3, c: 4 });
  assert.equal(r.novas, 1);
});
