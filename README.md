# Gastos de Eletricidade

Aplicação web para acompanhar os gastos de eletricidade da casa a partir dos
consumos de 15 em 15 minutos que a **E-Redes** disponibiliza no Balcão Digital.

- **Período de faturação do dia 16 ao dia 15** do mês seguinte (configurável).
- **Resumo**: consumo, custo estimado, média diária, previsão para o período e pico de potência.
- **Comparar**: últimos 13 períodos (kWh ou €), consumo acumulado vs. período anterior,
  tabela com variação face ao mesmo período do ano anterior.
- **Horas do dia**: perfil médio por hora (todos os dias, dias úteis ou fim de semana),
  mapa de calor dia × hora, consumo de base e os 10 quartos de hora de maior consumo.
- **Tarifas** simples, bi-horária (ciclo diário ou semanal) e tri-horária (ciclo diário),
  com horário de verão/inverno.
- Funciona **no iPad, sem servidor**: os dados ficam guardados só no próprio dispositivo
  (IndexedDB) e a aplicação funciona sem rede depois de aberta a primeira vez.

## Usar no iPad

1. Abra o endereço da aplicação no Safari (ver *Publicar* abaixo).
2. Toque em **Partilhar → Adicionar ao ecrã principal**. Passa a abrir como uma app
   e o iPadOS não apaga os dados por falta de uso.
3. No Balcão Digital da E-Redes (balcaodigital.e-redes.pt), em **Consumos**, escolha a
   vista de 15 minutos e o intervalo de datas e exporte o Excel.
4. Na app, em **Dados → Escolher ficheiro**, selecione o ficheiro em *Ficheiros → Transferências*.
5. Em **Dados → Tarifa**, introduza os preços da sua fatura (com IVA).

Pode importar ficheiros novos sempre que quiser: as leituras repetidas são substituídas.
Use **Exportar cópia** de vez em quando para guardar os dados no iCloud Drive.

> A E-Redes não tem uma API pública e o login usa CAPTCHA, por isso a importação é feita
> a partir do ficheiro exportado. A E-Redes indica a potência média (kW) de cada
> quarto de hora; a app converte para kWh (kW ÷ 4).

Para experimentar sem dados reais, importe `exemplo/consumos-exemplo.xlsx` (dados fictícios).

## Publicar (GitHub Pages)

*Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`*.
O endereço fica `https://<utilizador>.github.io/gastos-eletricidade/`.
Os seus consumos **nunca** são enviados para o GitHub: só o código da aplicação é publicado.

## Desenvolvimento

Sem dependências nem compilação: HTML, CSS e JavaScript (módulos ES).
As bibliotecas [Chart.js](https://www.chartjs.org/) 4.4.1 e
[SheetJS](https://sheetjs.com/) 0.18.5 estão em `vendor/`.

```sh
npm start   # servidor local em http://localhost:8080
npm test    # testes da lógica (períodos, tarifas, leitura de ficheiros, custos)
```

| Ficheiro | Conteúdo |
|---|---|
| `js/core.js` | leitura do Excel, tarifas, períodos de faturação, agregações (sem DOM, testável) |
| `js/app.js` | interface e gráficos |
| `js/store.js` | armazenamento local (IndexedDB) |
| `sw.js` | cache para uso sem rede (altere `VERSAO` ao publicar alterações) |
| `tools/gerar-exemplo.cjs` | gera o ficheiro de exemplo |
