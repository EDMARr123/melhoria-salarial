r"""
Gera o painel "Melhoria Salarial" (painel.html) a partir de dados.json.

Modelo igual à planilha original: trata 1 RCA por vez — digita o código,
puxa nome e todos os dados daquele vendedor. Metas (positivação por
categoria e pedidos/dia) e taxa média de comissão ficam editáveis num
painel de configuração (valem pra qualquer vendedor que você olhar,
igual eram fixas em ITENS FOCO na planilha); os checkboxes de
Industrializado/Thermo/Dia 15/Dia 30/Campanha ficam salvos por RCA no
localStorage do navegador. Tudo recalcula ao vivo, sem precisar reabrir
o Excel.
"""

import base64
import json
import os

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
CAMINHO_DADOS = os.path.join(PASTA_BASE, "dados.json")
CAMINHO_SAIDA = os.path.join(PASTA_BASE, "painel.html")

# Reaproveita as mesmas fotos já cadastradas no painel_pilares — mesmos
# vendedores/supervisores, não faz sentido duplicar os arquivos aqui.
PASTA_FOTOS_SUPERVISORES = os.path.join(PASTA_BASE, "..", "painel_pilares", "fotos_supervisores")
PASTA_FOTOS_RCAS = os.path.join(PASTA_BASE, "..", "painel_pilares", "fotos_rcas")


def _fotos_supervisores_json():
    fotos = {}
    if os.path.isdir(PASTA_FOTOS_SUPERVISORES):
        for nome_arquivo in os.listdir(PASTA_FOTOS_SUPERVISORES):
            nome, ext = os.path.splitext(nome_arquivo)
            if ext.lower() not in (".jpg", ".jpeg", ".png"):
                continue
            tipo_mime = "image/png" if ext.lower() == ".png" else "image/jpeg"
            with open(os.path.join(PASTA_FOTOS_SUPERVISORES, nome_arquivo), "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            fotos[nome.upper()] = f"data:{tipo_mime};base64,{b64}"
    return json.dumps(fotos, ensure_ascii=False)


def _fotos_rcas_json():
    """fotos_rcas/{SUPERVISOR}/{NOME_RCA}.jpg — chave = nome do RCA em
    maiúsculas (mesmo texto usado em dados.json), não precisa bater com o
    código."""
    fotos = {}
    if os.path.isdir(PASTA_FOTOS_RCAS):
        for nome_pasta_supervisor in os.listdir(PASTA_FOTOS_RCAS):
            pasta_supervisor = os.path.join(PASTA_FOTOS_RCAS, nome_pasta_supervisor)
            if not os.path.isdir(pasta_supervisor):
                continue
            for nome_arquivo in os.listdir(pasta_supervisor):
                nome, ext = os.path.splitext(nome_arquivo)
                if ext.lower() not in (".jpg", ".jpeg", ".png"):
                    continue
                tipo_mime = "image/png" if ext.lower() == ".png" else "image/jpeg"
                with open(os.path.join(pasta_supervisor, nome_arquivo), "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("ascii")
                fotos[nome.upper().strip()] = f"data:{tipo_mime};base64,{b64}"
    return json.dumps(fotos, ensure_ascii=False)


_FOTOS_SUPERVISORES_JSON = _fotos_supervisores_json()
_FOTOS_RCAS_JSON = _fotos_rcas_json()

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
    --bg: #12160F; --surface: #1B211A; --surface-2: #212820; --border: #303A2E;
    --ink: #E9EEE6; --ink-soft: #AEBAA9; --ink-faint: #7C887A;
    --good: #3FC17F; --good-soft: #123625; --bad: #E2685F; --bad-soft: #3A1D1B;
    --warn: #E0A73C; --warn-soft: #3A2E12; --accent: #3FC17F; --accent-soft: #17301F;
    --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 12px 28px -14px rgba(0,0,0,0.6);
  }
}
:root[data-theme="dark"] {
  --bg: #12160F; --surface: #1B211A; --surface-2: #212820; --border: #303A2E;
  --ink: #E9EEE6; --ink-soft: #AEBAA9; --ink-faint: #7C887A;
  --good: #3FC17F; --good-soft: #123625; --bad: #E2685F; --bad-soft: #3A1D1B;
  --warn: #E0A73C; --warn-soft: #3A2E12; --accent: #3FC17F; --accent-soft: #17301F;
  --shadow: 0 1px 2px rgba(0,0,0,0.35), 0 12px 28px -14px rgba(0,0,0,0.6);
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--ink); font-family: ui-sans-serif, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; -webkit-font-smoothing: antialiased; }
.wrap { max-width: 1180px; margin: 0 auto; padding: 24px 20px 64px; }
header.top h1 { font-size: 22px; margin: 0 0 4px; }
header.top p { margin: 0 0 18px; color: var(--ink-soft); font-size: 13.5px; }

.panel { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; box-shadow: var(--shadow); padding: 18px 20px; margin-bottom: 16px; }
.panel h2 { font-size: 13px; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-faint); margin: 0 0 14px; font-weight: 800; }

.busca-row { display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap; }
.campo { display: flex; flex-direction: column; gap: 4px; }
.campo label { font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .03em; color: var(--ink-faint); }
.campo input, .campo select {
  font: inherit; font-size: 15px; font-weight: 700; padding: 8px 10px; border-radius: 9px;
  border: 1px solid var(--border); background: var(--surface-2); color: var(--ink);
  font-variant-numeric: tabular-nums;
}
.campo input:disabled { color: var(--accent); border-style: dashed; opacity: 1; }
.media-pedidos-linha {
  align-self: center; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .02em;
  color: var(--accent); line-height: 1.8;
}
#codigoInput { width: 110px; }
#nomeAtual { font-size: 15px; font-weight: 800; color: var(--accent); align-self: center; padding-bottom: 8px; }
.rota-atual { font-size: 12.5px; color: var(--ink-faint); align-self: center; padding-bottom: 8px; }

.linha-cabecalho { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 4px; }
.linha-cabecalho .box { background: var(--surface-2); border-radius: 10px; padding: 10px 12px; }
.linha-cabecalho .box .l { font-size: 10.5px; font-weight: 800; text-transform: uppercase; color: var(--ink-faint); margin-bottom: 2px; }
.linha-cabecalho .box .v { font-size: 16px; font-weight: 800; font-variant-numeric: tabular-nums; }

table.breakdown { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 6px; }
table.breakdown th { text-align: right; font-size: 10.5px; text-transform: uppercase; color: var(--ink-faint); font-weight: 800; padding: 6px 6px; border-bottom: 1px solid var(--border); }
table.breakdown th:first-child, table.breakdown td:first-child { text-align: left; }
table.breakdown td { text-align: right; padding: 6px 6px; font-variant-numeric: tabular-nums; font-weight: 700; border-bottom: 1px solid var(--border); }
table.breakdown tr:last-child td { border-bottom: none; }
table.breakdown td.dif-pos { color: var(--good); }
table.breakdown td.dif-zero { color: var(--ink-faint); font-weight: 500; }
table.breakdown input.meta-posit-input {
  width: 56px; font: inherit; font-size: 12.5px; font-weight: 800; text-align: center;
  padding: 4px 4px; border-radius: 6px; border: 1px solid var(--border); background: var(--surface-2); color: var(--accent);
}

.checks { display: flex; flex-wrap: wrap; gap: 8px; margin: 12px 0 4px; }
.chk {
  display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 700;
  padding: 6px 11px; border-radius: 9px; background: var(--surface-2); border: 1px solid var(--border);
  cursor: pointer; user-select: none; color: var(--ink-soft);
}
.chk input { accent-color: var(--accent); width: 15px; height: 15px; }
.chk.on { background: var(--good-soft); color: var(--good); border-color: transparent; }

.resumo-grid { display: grid; grid-template-columns: 1fr; gap: 10px; margin-top: 14px; }
.resumo-box { border-radius: 12px; padding: 14px 18px; display: flex; justify-content: space-between; align-items: center; }
.resumo-box .l { font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: .03em; }
.resumo-box .v { font-size: 22px; font-weight: 800; font-variant-numeric: tabular-nums; }
.resumo-box.soma { background: var(--good-soft); color: var(--good); }
.resumo-box.atual { background: var(--surface-2); color: var(--ink); }
.resumo-box.total { background: var(--accent-soft); color: var(--accent); }

.btn-limpar {
  margin-top: 12px; background: var(--bad); color: #fff; border: none; border-radius: 10px;
  padding: 11px 20px; font: inherit; font-size: 13.5px; font-weight: 800; cursor: pointer; width: 100%;
}
.btn-limpar:hover { opacity: .9; }

.vazio { text-align: center; padding: 60px 20px; color: var(--ink-faint); font-size: 14px; }
.foot { text-align: center; color: var(--ink-faint); font-size: 12px; margin-top: 30px; }
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <h1>Melhoria Salarial</h1>
    <p>Digite o código do RCA pra ver o salário atual x potencial dele — igual à planilha, 1 vendedor por vez.</p>
  </header>

  <div class="panel">
    <div class="busca-row">
      <div class="campo">
        <label>Código do RCA</label>
        <input type="number" id="codigoInput" list="rcaList" placeholder="ex: 15">
        <datalist id="rcaList"></datalist>
      </div>
      <div id="avatarAtual" style="display:flex;align-self:center"></div>
      <div id="nomeAtual"></div>
      <div class="rota-atual" id="rotaAtual"></div>
    </div>
  </div>

  <div id="conteudo"></div>

  <p class="foot">Dados extraídos de RESULTADO.xlsm (MELHORIA SALARIO) · gerado automaticamente</p>
</div>

<script>
const DADOS = __DADOS_JSON__;
const FOTOS_SUPERVISORES = __FOTOS_SUPERVISORES_JSON__;
const FOTOS_RCAS = __FOTOS_RCAS_JSON__;
const RCAS_POR_CODIGO = Object.fromEntries(DADOS.rcas.map(r => [r.codigo, r]));

function fmtMoeda(v) {
  return "R$ " + Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtPct(v) {
  return (v * 100).toLocaleString("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
}

function normalizarNomeFoto(nome) {
  return nome.replace(/\s*-\s*$/, "").trim().toUpperCase();
}

function iniciais(nome) {
  return nome.split(/\s+/).filter(Boolean).slice(0, 2).map(p => p[0]).join("").toUpperCase();
}

// Cor de avatar determinística a partir do nome, pra quando não tem foto cadastrada.
const PALETA_AVATAR = ["#0E7C86", "#7A5CC7", "#C0672B", "#3D6FB4", "#1D9A5D", "#B4740A", "#B4406B"];
function corAvatar(nome) {
  let h = 0;
  for (const c of nome) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return PALETA_AVATAR[h % PALETA_AVATAR.length];
}

function avatarHtml(nome, foto, tamanho) {
  return foto
    ? `<img src="${foto}" alt="${nome}" style="width:${tamanho}px;height:${tamanho}px;object-fit:cover;border-radius:50%;flex:none;">`
    : `<div style="width:${tamanho}px;height:${tamanho}px;border-radius:50%;flex:none;background:${corAvatar(nome)};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:${Math.round(tamanho * 0.36)}px;">${iniciais(nome)}</div>`;
}

// Campos numéricos editáveis são type="text" (não type="number") — assim
// dá pra devolver o cursor no lugar certo depois de cada tecla, coisa que
// o Chrome não deixa fazer em input number (setSelectionRange trava nele
// e faz o cursor voltar pro começo, invertendo os dígitos digitados).
// Por isso aceita vírgula OU ponto na entrada, e mostra sempre com vírgula.
function parseNum(str) {
  const v = parseFloat(String(str).replace(",", "."));
  return isNaN(v) ? 0 : v;
}
function fmtInput(v) {
  return String(v).replace(".", ",");
}

// ---- Configuração global (só taxa + meta de pedidos/dia — valem pra
// qualquer vendedor) ----
function lerConfig() {
  const padrao = {
    taxaPct: DADOS.constantes.taxa_padrao * 100,
    metaPedidosDia: DADOS.constantes.meta_pedidos_dia,
  };
  try {
    const salvo = localStorage.getItem("melhoria_salarial_config");
    if (!salvo) return padrao;
    const parsed = JSON.parse(salvo);
    return {
      taxaPct: parsed.taxaPct ?? padrao.taxaPct,
      metaPedidosDia: parsed.metaPedidosDia ?? padrao.metaPedidosDia,
    };
  } catch (e) { return padrao; }
}

function salvarConfig(config) {
  try { localStorage.setItem("melhoria_salarial_config", JSON.stringify(config)); } catch (e) {}
}

// ---- Estado por RCA (checkboxes + Meta Posit de cada categoria) — Meta
// Posit é por vendedor, não global, porque "Limpar marcações deste
// vendedor" precisa zerar só a meta DELE, sem mexer nos outros.
function chaveEstado(codigo) { return "melhoria_salarial_estado_" + codigo; }

function estadoPadrao() {
  return {
    industrializado: false, thermo: false, dia15: false, dia30: false, campanha: false,
    metasCategoria: Object.assign({}, DADOS.constantes.metas_categoria_padrao),
  };
}

function lerEstado(codigo) {
  const padrao = estadoPadrao();
  try {
    const salvo = localStorage.getItem(chaveEstado(codigo));
    if (!salvo) return padrao;
    const parsed = JSON.parse(salvo);
    return {
      industrializado: parsed.industrializado ?? padrao.industrializado,
      thermo: parsed.thermo ?? padrao.thermo,
      dia15: parsed.dia15 ?? padrao.dia15,
      dia30: parsed.dia30 ?? padrao.dia30,
      campanha: parsed.campanha ?? padrao.campanha,
      metasCategoria: Object.assign({}, padrao.metasCategoria, parsed.metasCategoria || {}),
    };
  } catch (e) { return padrao; }
}

function salvarEstado(codigo, estado) {
  try { localStorage.setItem(chaveEstado(codigo), JSON.stringify(estado)); } catch (e) {}
}

function limparEstado(codigo) {
  const zerado = {
    industrializado: false, thermo: false, dia15: false, dia30: false, campanha: false,
    metasCategoria: Object.fromEntries(DADOS.constantes.ordem_categorias.map(c => [c, 0])),
  };
  salvarEstado(codigo, zerado);
}

// ---- Cálculo ----
function calcular(rca, config) {
  const taxa = config.taxaPct / 100;
  const estado = lerEstado(rca.codigo);
  const { dias_uteis, taxas_categoria, labels_categoria, ordem_categorias } = DADOS.constantes;

  const linhasCategorias = ordem_categorias.map(chave => {
    const cat = rca.categorias[chave];
    const metaPosit = estado.metasCategoria[chave];
    const comissaoAtual = cat.valor * taxas_categoria[chave];
    const comissaoPotencial = cat.positivacao ? (comissaoAtual / cat.positivacao) * metaPosit : 0;
    return { chave, label: labels_categoria[chave], atual: comissaoAtual, potencial: comissaoPotencial, ...cat, metaPosit };
  });
  const departamentosAtual = linhasCategorias.reduce((s, l) => s + l.atual, 0);
  const departamentosPotencial = linhasCategorias.reduce((s, l) => s + l.potencial, 0);

  const pedidosAtual = taxa * rca.valor_vendido;
  const ticketMedio = rca.total_pedidos ? rca.valor_vendido / rca.total_pedidos : 0;
  const pedidosPotencial = taxa * ticketMedio * config.metaPedidosDia * dias_uteis;

  const linhasResumo = [
    { label: "Departamentos", atual: departamentosAtual, potencial: departamentosPotencial },
    { label: "Industrializado", atual: estado.industrializado ? rca.industrializado_potencial : 0, potencial: rca.industrializado_potencial },
    { label: "Thermo", atual: estado.thermo ? rca.thermo_potencial : 0, potencial: rca.thermo_potencial },
    { label: "Pedidos", atual: pedidosAtual, potencial: pedidosPotencial },
    { label: "Bônus Dia 15", atual: estado.dia15 ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
    { label: "Bônus Dia 30", atual: estado.dia30 ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
    { label: "Prêmio Campanha", atual: estado.campanha ? rca.premio_fixo : 0, potencial: rca.premio_fixo },
  ];

  const soma = linhasResumo.reduce((s, l) => s + (l.potencial - l.atual), 0);
  const salarioAtual = pedidosAtual;
  const salarioTotal = salarioAtual + soma;

  return { linhasCategorias, linhasResumo, soma, salarioAtual, salarioTotal, estado, ticketMedio };
}

function montarConteudo(rca) {
  const config = lerConfig();
  const r = calcular(rca, config);

  const linhasCatHtml = r.linhasCategorias.map(l => {
    const dif = l.potencial - l.atual;
    const classeDif = dif > 0.005 ? "dif-pos" : "dif-zero";
    return `<tr>
      <td>${l.label}</td>
      <td>${l.peso.toLocaleString("pt-BR", {maximumFractionDigits:2})}</td>
      <td>${fmtMoeda(l.valor)}</td>
      <td>${l.positivacao}</td>
      <td><input type="text" inputmode="numeric" class="meta-posit-input" data-chave="${l.chave}" value="${fmtInput(l.metaPosit)}"></td>
      <td>${fmtMoeda(l.atual)}</td>
      <td>${fmtMoeda(l.potencial)}</td>
      <td class="${classeDif}">${fmtMoeda(dif)}</td>
    </tr>`;
  }).join("");

  const linhasResumoHtml = r.linhasResumo.map(l => {
    const dif = l.potencial - l.atual;
    const classeDif = dif > 0.005 ? "dif-pos" : "dif-zero";
    return `<tr><td>${l.label}</td><td>${fmtMoeda(l.atual)}</td><td>${fmtMoeda(l.potencial)}</td><td class="${classeDif}">${fmtMoeda(dif)}</td></tr>`;
  }).join("");

  const checks = [
    ["industrializado", "Industrializado bateu"],
    ["thermo", "Thermo bateu"],
    ["dia15", "Positivação Dia 15 bateu"],
    ["dia30", "Positivação Dia 30 bateu"],
    ["campanha", "Prêmio Campanha bateu"],
  ].map(([chave, label]) => {
    const on = r.estado[chave] ? "on" : "";
    return `<label class="chk ${on}" data-chave="${chave}">
      <input type="checkbox" ${r.estado[chave] ? "checked" : ""} data-chave="${chave}"> ${label}
    </label>`;
  }).join("");

  return `
    <div class="panel">
      <div class="linha-cabecalho">
        <div class="box"><div class="l">Valor Vendido</div><div class="v">${fmtMoeda(rca.valor_vendido)}</div></div>
        <div class="box"><div class="l">Ticket Médio</div><div class="v">${fmtMoeda(r.ticketMedio)}</div></div>
        <div class="box" style="display:flex;align-items:center;gap:10px">
          ${avatarHtml(rca.supervisor, FOTOS_SUPERVISORES[rca.supervisor], 36)}
          <div><div class="l">Supervisor</div><div class="v">${rca.supervisor}</div></div>
        </div>
      </div>
    </div>

    <div class="panel">
      <div class="busca-row" style="align-items:flex-start">
        <div class="campo">
          <label>Taxa média comissão (%)</label>
          <input type="text" inputmode="decimal" id="cfgTaxaInline" value="${fmtInput(config.taxaPct)}">
        </div>
        <div class="campo">
          <label>Meta de pedidos/dia</label>
          <input type="text" inputmode="numeric" id="cfgMetaPedidosInline" value="${fmtInput(config.metaPedidosDia)}">
        </div>
        <div class="media-pedidos-linha">
          QTD NO MÊS: ${rca.total_pedidos}<br>
          MÉDIA DE PEDIDOS: ${Math.round(rca.total_pedidos / DADOS.constantes.dias_uteis)}
        </div>
        <div class="media-pedidos-linha" style="margin-left:auto;text-align:right">
          PARTICIPAÇÃO INDUSTRIALIZADO: ${fmtPct(rca.industrializado_participacao_pct)}<br>
          MARGEM INDUSTRIALIZADO: ${fmtPct(rca.industrializado_margem_pct)}
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>Por categoria</h2>
      <div style="overflow-x:auto">
      <table class="breakdown">
        <thead><tr><th>Categoria</th><th>Peso</th><th>Valor</th><th>Positivação</th><th>Meta Posit</th><th>Comissão Atual</th><th>Comissão Potencial</th><th>Diferença</th></tr></thead>
        <tbody>${linhasCatHtml}</tbody>
      </table>
      </div>
    </div>

    <div class="panel">
      <h2>Bônus e desafios</h2>
      <div class="checks">${checks}</div>
    </div>

    <div class="panel">
      <h2>Resumo</h2>
      <table class="breakdown">
        <thead><tr><th>Tópico</th><th>Atual</th><th>Potencial</th><th>Diferença</th></tr></thead>
        <tbody>${linhasResumoHtml}</tbody>
      </table>
      <div class="resumo-grid">
        <div class="resumo-box soma"><span class="l">Somar (upside)</span><span class="v">${fmtMoeda(r.soma)}</span></div>
        <div class="resumo-box atual"><span class="l">Salário Atual</span><span class="v">${fmtMoeda(r.salarioAtual)}</span></div>
        <div class="resumo-box total"><span class="l">Salário Total Potencial</span><span class="v">${fmtMoeda(r.salarioTotal)}</span></div>
      </div>
      <button class="btn-limpar" id="btnLimpar">Limpar marcações deste vendedor</button>
    </div>
  `;
}

function renderizarRca() {
  const codigo = parseInt(document.getElementById("codigoInput").value);
  const conteudo = document.getElementById("conteudo");
  const avatarAtual = document.getElementById("avatarAtual");
  const nomeAtual = document.getElementById("nomeAtual");
  const rotaAtual = document.getElementById("rotaAtual");

  // Campos de meta são editados ao vivo (digitando) — como o innerHTML
  // inteiro é reconstruído a cada tecla, sem isto o campo perderia o foco
  // e o cursor a cada dígito. Guarda quem estava focado e devolve o foco
  // depois de redesenhar.
  const ativo = document.activeElement;
  const focoAntes = (ativo && conteudo.contains(ativo) && ativo.tagName === "INPUT")
    ? { id: ativo.id, chave: ativo.dataset.chave, inicio: ativo.selectionStart, fim: ativo.selectionEnd, valorBruto: ativo.value }
    : null;

  const rca = RCAS_POR_CODIGO[codigo];
  if (!rca) {
    avatarAtual.innerHTML = "";
    nomeAtual.textContent = "";
    rotaAtual.textContent = "";
    conteudo.innerHTML = codigo ? `<div class="vazio">RCA ${codigo} não encontrado.</div>` : `<div class="vazio">Digite um código de RCA acima pra começar.</div>`;
    return;
  }

  avatarAtual.innerHTML = avatarHtml(rca.nome, FOTOS_RCAS[normalizarNomeFoto(rca.nome)], 40);
  nomeAtual.textContent = rca.nome;
  rotaAtual.textContent = `RCA ${rca.codigo} · ${rca.rota}`;
  conteudo.innerHTML = montarConteudo(rca);

  if (focoAntes) {
    const seletor = focoAntes.chave
      ? `[data-chave="${focoAntes.chave}"]`
      : (focoAntes.id ? `#${focoAntes.id}` : null);
    const novo = seletor && conteudo.querySelector(seletor);
    if (novo) {
      // Devolve exatamente o texto que a pessoa tinha digitado (não o
      // número já formatado/arredondado) — senão uma vírgula digitada na
      // hora certa some antes de dar tempo de digitar a casa decimal.
      novo.value = focoAntes.valorBruto;
      novo.focus();
      try { novo.setSelectionRange(focoAntes.inicio, focoAntes.fim); } catch (e) {}
    }
  }
}

// Delegação anexada 1 vez só (o #conteudo em si nunca é recriado, só o
// innerHTML dele — anexar listener dentro de renderizarRca() empilharia
// um novo listener duplicado a cada troca de RCA/checkbox).
document.getElementById("conteudo").addEventListener("change", (e) => {
  const chk = e.target.closest("input[type=checkbox][data-chave]");
  if (chk) {
    const codigo = parseInt(document.getElementById("codigoInput").value);
    const chave = chk.dataset.chave;
    const estado = lerEstado(codigo);
    estado[chave] = chk.checked;
    salvarEstado(codigo, estado);
    renderizarRca();
    return;
  }
});

document.getElementById("conteudo").addEventListener("input", (e) => {
  if (e.target.classList.contains("meta-posit-input")) {
    const codigo = parseInt(document.getElementById("codigoInput").value);
    const estado = lerEstado(codigo);
    estado.metasCategoria[e.target.dataset.chave] = parseNum(e.target.value);
    salvarEstado(codigo, estado);
    renderizarRca();
  } else if (e.target.id === "cfgMetaPedidosInline") {
    const cfg = lerConfig();
    cfg.metaPedidosDia = parseNum(e.target.value);
    salvarConfig(cfg);
    renderizarRca();
  } else if (e.target.id === "cfgTaxaInline") {
    const cfg = lerConfig();
    cfg.taxaPct = parseNum(e.target.value);
    salvarConfig(cfg);
    renderizarRca();
  }
});

document.getElementById("conteudo").addEventListener("click", (e) => {
  if (!e.target.closest("#btnLimpar")) return;
  const codigo = parseInt(document.getElementById("codigoInput").value);
  limparEstado(codigo);
  renderizarRca();
});

function montarDatalist() {
  const opts = DADOS.rcas
    .slice()
    .sort((a, b) => a.nome.localeCompare(b.nome))
    .map(r => `<option value="${r.codigo}">${r.nome} — RCA ${r.codigo}</option>`)
    .join("");
  document.getElementById("rcaList").innerHTML = opts;
}

document.getElementById("codigoInput").addEventListener("input", renderizarRca);

montarDatalist();
renderizarRca();
</script>
</body>
</html>
"""


def main():
    with open(CAMINHO_DADOS, "r", encoding="utf-8") as f:
        dados = json.load(f)

    html = TEMPLATE.replace("__DADOS_JSON__", json.dumps(dados, ensure_ascii=False))
    html = html.replace("__FOTOS_SUPERVISORES_JSON__", _FOTOS_SUPERVISORES_JSON)
    html = html.replace("__FOTOS_RCAS_JSON__", _FOTOS_RCAS_JSON)
    with open(CAMINHO_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Painel gerado em: {CAMINHO_SAIDA}")


if __name__ == "__main__":
    main()
