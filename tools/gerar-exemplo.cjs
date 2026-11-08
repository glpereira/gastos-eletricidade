// Gera exemplo/consumos-exemplo.xlsx com dados fictícios no formato do Balcão Digital.
// Uso: node tools/gerar-exemplo.cjs  (requer `npm i xlsx` numa pasta temporária, ou NODE_PATH)
const XLSX = require('xlsx');
let seed = 7;
const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
const linhas = [
  ['Código do Ponto de Entrega', 'PT0002000000000000XX (exemplo)'],
  ['Período', 'Diagrama de carga - 15 minutos'],
  [],
  ['Data', 'Hora', 'Consumo registado, Ativa (kW)', 'Estado'],
];
const pad = (n) => String(n).padStart(2, '0');
const inicio = Date.UTC(2025, 9, 16), fim = Date.UTC(2026, 10, 8);
for (let t = inicio + 15 * 60000; t <= fim; t += 15 * 60000) {
  const d = new Date(t - 15 * 60000);
  const h = d.getUTCHours() + d.getUTCMinutes() / 60;
  const mes = d.getUTCMonth(), dow = d.getUTCDay();
  const inverno = [10, 11, 0, 1, 2].includes(mes) ? 1 : [3, 9].includes(mes) ? 0.5 : 0;
  let kw = 0.12 + 0.03 * rnd(); // base: frigorífico, router...
  if (h >= 7 && h < 8.5) kw += 0.6 * rnd();
  if (h >= 12.5 && h < 14 && (dow === 0 || dow === 6)) kw += 1.2 * rnd();
  if (h >= 19 && h < 21) kw += 1.0 + 1.4 * rnd(); // jantar
  if (h >= 18 && h < 23) kw += inverno * (0.8 + 0.6 * rnd()); // aquecimento
  if (h >= 21.5 && h < 23 && rnd() < 0.3) kw += 1.8; // máquinas
  if (h >= 2 && h < 3.5 && dow === 3) kw += 1.6; // termoacumulador / máquina no vazio
  if (rnd() < 0.01) kw += 2.5 * rnd(); // forno, chaleira
  const e = new Date(t);
  const hora = e.getUTCHours() === 0 && e.getUTCMinutes() === 0 ? '00:00' : `${pad(e.getUTCHours())}:${pad(e.getUTCMinutes())}`;
  linhas.push([`${e.getUTCFullYear()}/${pad(e.getUTCMonth() + 1)}/${pad(e.getUTCDate())}`, hora, kw.toFixed(3).replace('.', ','), 'Real']);
}
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(linhas), 'Consumos');
XLSX.writeFile(wb, require('path').join(__dirname, '..', 'exemplo', 'consumos-exemplo.xlsx'), { compression: true });
console.log(linhas.length - 4, 'leituras');
