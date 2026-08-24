r"""
Gera o painel "Melhoria Salarial" (painel.html) a partir de dados.json.

Ao contrário dos outros painéis (4 Pilares, Departamentos), este é
INTERATIVO: a taxa média de comissão sobre pedidos e os checkboxes de
"bateu Industrializado/Thermo/Dia 15/Dia 30/Campanha" são editáveis
direto no navegador (o Edmar decide RCA por RCA, igual fazia na
planilha) e o Salário Atual/Potencial recalcula na hora, em JS. O estado
fica salvo no localStorage do navegador de quem estiver usando.
"""

import json
import os

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
CAMINHO_DADOS = os.path.join(PASTA_BASE, "dados.json")
CAMINHO_SAIDA = os.path.join(PASTA_BASE, "painel.html")

TEMPLATE = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Melhoria Salarial — Equipe GYN</title>
<style>
:root {
  --bg: #F2F4F0;
  --surface: #FFFFFF;
  --surface-2: #F7F8F5;
  --border: #E2E5DD;
  --ink: #1C231C;
  --ink-soft: #5B655A;
  --ink-faint: #8B948A;
  --good: #1D9A5D;
  --good-soft: #E4F5EC;
  --bad: #C23B3B;
  --bad-soft: #FBEAEA;
  --warn: #B9790F;
  --warn-soft: #FBF0DC;
  --accent: #1F7A5C;
  --accent-soft: #E1F1EA;
  --shadow: 0 1px 2px rgba(20,25,20,0.05), 0 10px 24px -14px rgba(20,25,20,0.18);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --bg: #12160F;
    --surface: #1B211A;
    --surface-2: #212820;
    --border: #303A2E;
    --ink: #E9EEE6;
    --ink-soft: #AEBAA9;
    --ink-faint: #7C887A;
    --good: #3FC17F;
    --good-soft: #123625;
    --bad: #E2685F;
    --bad-soft: #3A1D1B;
    --warn: #E0A73C;
    --warn-soft: #3A2E12;
    --accent: #3FC17F;
    --accent-soft: #17301F;
    --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 12px 28px -14px rgba(0,0,0,0.6);
  }
}
:root[data-theme="dark"] {
  --bg: #12160F;
  --surface: #1B211A;
  --surface-2: #212820;
  --border: #303A2E;
  --ink: #E9EEE6;
  --ink-soft: #AEBAA9;
  --ink-faint: #7C887A;
  --good: #3FC17F;
  --good-soft: #123625;
  --bad: #E2685F;
  --bad-soft: #3A1D1B;
  --warn: #E0A73C;
  --warn-soft: #3A2E12;
  --accent: #3FC17F;
  --accent-soft: #17301F;
  --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 12px 28px -14px rgba(0,0,0,0.6);
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--ink); font-family: ui-sans-serif, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; -webkit-font-smoothing: antialiased; }
.wrap { max-width: 1400px; margin: 0 auto; padding: 24px 20px 64px; }
header.top { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; flex-wrap: wrap; margin-bottom: 18px; }
header.top h1 { font-size: 24px; margin: 0 0 4px; }
header.top p { margin: 0; color: var(--ink-soft); font-size: 13.5px; }

.config-bar {
  background: var(--surface); border: 1px solid var(--border); border-radius: 14px;
  box-shadow: var(--shadow); padding: 14px 18px; display: flex; gap: 24px; align-items: center;
  flex-wrap: wrap; margin-bottom: 18px;
}
.config-item { display: flex; flex-direction: column; gap: 3px; }
.config-item label { font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .03em; color: var(--ink-faint); }
.config-item .val-row { display: flex; align-items: center; gap: 6px; }
.config-item input[type=number] {
  width: 90px; font-size: 15px; font-weight: 800; padding: 5px 8px; border-radius: 8px;
  border: 1px solid var(--border); background: var(--surface-2); color: var(--ink);
  font-variant-numeric: tabular-nums;
}
.config-item .fixo { font-size: 15px; font-weight: 800; color: var(--ink); }
.config-note { font-size: 12px; color: var(--ink-faint); max-width: 320px; }

.tabs { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 16px; }
.tab {
  appearance: none; border: 1px solid var(--border); background: var(--surface); color: var(--ink-soft);
  font: inherit; font-size: 12.5px; font-weight: 700; padding: 6px 13px; border-radius: 999px; cursor: pointer;
}
.tab.active { background: var(--accent); border-color: var(--accent); color: #fff; }
.tab .count { opacity: .75; margin-left: 4px; }

.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 14px; }

.card { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; box-shadow: var(--shadow); padding: 16px; display: flex; flex-direction: column; gap: 10px; }
.card-head { display: flex; justify-content: space-between; align-items: baseline; gap: 8px; }
.card-head h3 { margin: 0; font-size: 15.5px; }
.card-head .meta { font-size: 11.5px; color: var(--ink-faint); }

.checks { display: flex; flex-wrap: wrap; gap: 6px; }
.chk {
  display: inline-flex; align-items: center; gap: 5px; font-size: 11px; font-weight: 700;
  padding: 4px 8px; border-radius: 8px; background: var(--surface-2); border: 1px solid var(--border);
  cursor: pointer; user-select: none; color: var(--ink-soft);
}
.chk input { accent-color: var(--accent); }
.chk.on { background: var(--good-soft); color: var(--good); border-color: transparent; }

table.breakdown { width: 100%; border-collapse: collapse; font-size: 12px; }
table.breakdown th { text-align: right; font-size: 10px; text-transform: uppercase; color: var(--ink-faint); font-weight: 800; padding: 4px 4px; border-bottom: 1px solid var(--border); }
table.breakdown th:first-child, table.breakdown td:first-child { text-align: left; }
table.breakdown td { text-align: right; padding: 4px 4px; font-variant-numeric: tabular-nums; font-weight: 700; }
table.breakdown td.dif-pos { color: var(--good); }
table.breakdown td.dif-zero { color: var(--ink-faint); font-weight: 500; }

.summary { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 4px; }
.summary .box { border-radius: 10px; padding: 9px 10px; }
.summary .box .l { font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: .03em; margin-bottom: 2px; }
.summary .box .v { font-size: 16px; font-weight: 800; font-variant-numeric: tabular-nums; }
.summary .soma { background: var(--good-soft); color: var(--good); grid-column: 1 / -1; }
.summary .atual { background: var(--surface-2); color: var(--ink); }
.summary .total { background: var(--accent-soft); color: var(--accent); }

.foot { text-align: center; color: var(--ink-faint); font-size: 12px; margin-top: 36px; }
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div>
      <h1>Melhoria Salarial</h1>
      <p id="subtitle"></p>
    </div>
  </header>

  <div class="config-bar">
    <div class="config-item">
      <label>Taxa média de comissão</label>
      <div class="val-row"><input type="number" id="taxaInput" step="0.01" min="0"><span class="fixo">%</span></div>
    </div>
    <div class="config-item">
      <label>Dias úteis no mês</label>
      <div class="fixo" id="diasUteis"></div>
    </div>
    <div class="config-item">
      <label>Meta de pedidos/dia</label>
      <div class="fixo" id="metaPedidosDia"></div>
    </div>
    <p class="config-note">A taxa é usada só pra calcular o componente "Pedidos" do salário — os outros
    tópicos (Departamentos, Industrializado, Thermo, bônus) já vêm calculados dos dados reais.
    Marque os checkboxes de cada vendedor conforme ele for batendo a meta.</p>
  </div>

  <div class="tabs" id="tabs"></div>
  <div class="grid" id="grid"></div>

  <p class="foot">Dados extraídos de RESULTADO.xlsm (MELHORIA SALARIO) · gerado automaticamente</p>
</div>

<script>
const DADOS = __DADOS_JSON__;

function fmtMoeda(v) {
  return "R$ " + Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function chaveEstado(codigo) { return "melhoria_salarial_" + codigo; }

function lerEstado(codigo) {
  const padrao = { industrializado: false, thermo: false, dia15: false, dia30: false, campanha: false };
  try {
    const salvo = localStorage.getItem(chaveEstado(codigo));
    return salvo ? Object.assign(padrao, JSON.parse(salvo)) : padrao;
  } catch (e) { return padrao; }
}

function salvarEstado(codigo, estado) {
  try { localStorage.setItem(chaveEstado(codigo), JSON.stringify(estado)); } catch (e) {}
}

function lerTaxa() {
  try {
    const salva = localStorage.getItem("melhoria_salarial_taxa");
    return salva !== null ? parseFloat(salva) : DADOS.constantes.taxa_padrao * 100;
  } catch (e) { return DADOS.constantes.taxa_padrao * 100; }
}

function salvarTaxa(pct) {
  try { localStorage.setItem("melhoria_salarial_taxa", String(pct)); } catch (e) {}
}

function calcular(rca, taxaPct) {
  const taxa = taxaPct / 100;
  const estado = lerEstado(rca.codigo);
  const { dias_uteis, meta_pedidos_dia } = DADOS.constantes;

  const pedidosAtual = taxa * rca.valor_vendido;
  const ticketMedio = rca.total_pedidos ? rca.valor_vendido / rca.total_pedidos : 0;
  const pedidosPotencial = taxa * ticketMedio * meta_pedidos_dia * dias_uteis;

  const linhas = [
    { label: "Departamentos", atual: rca.departamentos.atual, potencial: rca.departamentos.potencial },
    { label: "Industrializado", atual: estado.industrializado ? rca.industrializado_potencial : 0, potencial: rca.industrializado_potencial },
    { label: "Thermo", atual: estado.thermo ? rca.thermo_potencial : 0, potencial: rca.thermo_potencial },
    { label: "Pedidos", atual: pedidosAtual, potencial: pedidosPotencial },
    { label: "Bônus Dia 15", atual: estado.dia15 ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
    { label: "Bônus Dia 30", atual: estado.dia30 ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
    { label: "Prêmio Campanha", atual: estado.campanha ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
  ];

  const soma = linhas.reduce((s, l) => s + (l.potencial - l.atual), 0);
  const salarioAtual = pedidosAtual;
  const salarioTotal = salarioAtual + soma;

  return { linhas, soma, salarioAtual, salarioTotal, estado };
}

function montarCard(rca, taxaPct) {
  const r = calcular(rca, taxaPct);

  const linhasHtml = r.linhas.map(l => {
    const dif = l.potencial - l.atual;
    const classeDif = dif > 0.005 ? "dif-pos" : "dif-zero";
    return `<tr>
      <td>${l.label}</td>
      <td>${fmtMoeda(l.atual)}</td>
      <td>${fmtMoeda(l.potencial)}</td>
      <td class="${classeDif}">${fmtMoeda(dif)}</td>
    </tr>`;
  }).join("");

  const checks = [
    ["industrializado", "Industrializado"],
    ["thermo", "Thermo"],
    ["dia15", "Dia 15"],
    ["dia30", "Dia 30"],
    ["campanha", "Campanha"],
  ].map(([chave, label]) => {
    const on = r.estado[chave] ? "on" : "";
    return `<label class="chk ${on}" data-chave="${chave}">
      <input type="checkbox" ${r.estado[chave] ? "checked" : ""} data-chave="${chave}"> ${label}
    </label>`;
  }).join("");

  return `
  <article class="card" data-codigo="${rca.codigo}" data-supervisor="${rca.supervisor}">
    <div class="card-head">
      <h3>${rca.nome}</h3>
      <span class="meta">RCA ${rca.codigo} · ${rca.supervisor}</span>
    </div>
    <div class="checks">${checks}</div>
    <table class="breakdown">
      <thead><tr><th></th><th>Atual</th><th>Potencial</th><th>Diferença</th></tr></thead>
      <tbody>${linhasHtml}</tbody>
    </table>
    <div class="summary">
      <div class="box soma"><div class="l">Soma (upside)</div><div class="v">${fmtMoeda(r.soma)}</div></div>
      <div class="box atual"><div class="l">Salário Atual</div><div class="v">${fmtMoeda(r.salarioAtual)}</div></div>
      <div class="box total"><div class="l">Salário Total Potencial</div><div class="v">${fmtMoeda(r.salarioTotal)}</div></div>
    </div>
  </article>`;
}

function montarTabs() {
  const supervisores = [...new Set(DADOS.rcas.map(r => r.supervisor))].sort();
  const contagem = s => DADOS.rcas.filter(r => r.supervisor === s).length;
  const tabsEl = document.getElementById("tabs");
  tabsEl.innerHTML =
    `<button class="tab active" data-sup="__todos__">Todos <span class="count">${DADOS.rcas.length}</span></button>` +
    supervisores.map(s => `<button class="tab" data-sup="${s}">${s} <span class="count">${contagem(s)}</span></button>`).join("");
  tabsEl.addEventListener("click", (e) => {
    const btn = e.target.closest(".tab");
    if (!btn) return;
    tabsEl.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    btn.classList.add("active");
    const sup = btn.dataset.sup;
    document.querySelectorAll(".card").forEach(c => {
      c.style.display = (sup === "__todos__" || c.dataset.supervisor === sup) ? "" : "none";
    });
  });
}

function renderizar() {
  const taxaPct = lerTaxa();
  document.getElementById("taxaInput").value = taxaPct.toFixed(2);
  document.getElementById("diasUteis").textContent = DADOS.constantes.dias_uteis;
  document.getElementById("metaPedidosDia").textContent = DADOS.constantes.meta_pedidos_dia;
  document.getElementById("subtitle").textContent =
    `${DADOS.rcas.length} RCAs — simulação de salário atual x potencial`;

  const ordenados = DADOS.rcas.slice().sort((a, b) => a.nome.localeCompare(b.nome));
  document.getElementById("grid").innerHTML = ordenados.map(r => montarCard(r, taxaPct)).join("");
}

document.getElementById("taxaInput").addEventListener("input", (e) => {
  const v = parseFloat(e.target.value);
  if (!isNaN(v)) { salvarTaxa(v); renderizar(); }
});

document.getElementById("grid").addEventListener("change", (e) => {
  const input = e.target.closest("input[type=checkbox]");
  if (!input) return;
  const card = e.target.closest(".card");
  const codigo = card.dataset.codigo;
  const chave = input.dataset.chave;
  const estado = lerEstado(codigo);
  estado[chave] = input.checked;
  salvarEstado(codigo, estado);
  renderizar();
});

montarTabs();
renderizar();
</script>
</body>
</html>
"""


def main():
    with open(CAMINHO_DADOS, "r", encoding="utf-8") as f:
        dados = json.load(f)

    html = TEMPLATE.replace("__DADOS_JSON__", json.dumps(dados, ensure_ascii=False))
    with open(CAMINHO_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Painel gerado em: {CAMINHO_SAIDA}")


if __name__ == "__main__":
    main()
