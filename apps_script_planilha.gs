// Script da planilha Google que recebe o "Feito" dos vendedores
// (Melhoria Salarial). Colar em Extensões > Apps Script e implantar como
// App da Web (Executar como: Eu; Quem pode acessar: Qualquer pessoa).

const ABA = "Respostas";

function _aba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(["Data/hora", "RCA", "Vendedor", "Supervisor", "Meta pedidos/dia",
                  "Salário atual", "Salário potencial", "Preenchimento (JSON)"]);
  }
  return sh;
}

// Vendedor clicou em "Feito"
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  _aba().appendRow([new Date(), d.rca, d.nome, d.supervisor, d.metaPedidosDia,
                    d.salarioAtual, d.salarioPotencial, JSON.stringify(d.estado)]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Painel do gerente lê a última resposta de cada RCA
function doGet() {
  const linhas = _aba().getDataRange().getValues().slice(1);
  const ultimo = {};
  linhas.forEach(l => {
    ultimo[l[1]] = { data: l[0], rca: l[1], nome: l[2], supervisor: l[3], metaPedidosDia: l[4],
                     salarioAtual: l[5], salarioPotencial: l[6], estado: JSON.parse(l[7] || "{}") };
  });
  return ContentService.createTextOutput(JSON.stringify(Object.values(ultimo)))
    .setMimeType(ContentService.MimeType.JSON);
}
