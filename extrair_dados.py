r"""
Extrai os dados do painel "Melhoria Salarial" a partir de RESULTADO.xlsm
(pasta MELHORIA SALARIO) e salva um dados.json pronto pro gerador de HTML
consumir.

A planilha original (abas ITENS FOCO / RESUMO) só calcula 1 RCA por vez —
o código digitado em ITENS FOCO!C8 decide quem aparece. Este script
reproduz a MESMA lógica de fórmulas da planilha, mas rodando pra TODOS os
RCAs de uma vez, lendo direto das abas de origem:

- '315' (ROTINA315): 1 linha por RCA — nome (col B), total de pedidos
  (col G), valor vendido (col J).
- BACON / CALABRESA / FRESCAIS / PAES / LACTIOS / BATATA / BOVINO / SUINO:
  1 linha por RCA — peso (col K), valor realizado (col L), positivação
  (col D).

Por categoria, meta de positivação (I na ITENS FOCO) e taxa de comissão
(G na ITENS FOCO) são fixas pra empresa toda (confirmado lendo a fórmula
de cada bloco, 21/08):
  bacon=3% calabresa=3% frescais=3% paes=3% lactios=3% batata=3%
  bovino=1,5% suino=3%
  metas: bacon=5 calabresa=22 frescais=15 paes=17 lactios=23 batata=15
         bovino=10 suino=6

Industrializado/Thermo/Positivação Dia 15/Dia 30/Prêmio Campanha usam
CHECKBOX MANUAL na planilha original (M17/M20/S11/S14/V11) — não tem
fórmula, é o Edmar revisando RCA por RCA. Aqui a gente só extrai o valor
"potencial" (o teto de cada bônus, se bater a meta); o "bateu ou não" fica
editável no próprio painel (checkbox no navegador, guardado no
localStorage — ver gerar_painel.py).

A taxa média de comissão sobre pedidos (ITENS FOCO!O11, hoje 0,0178) e o
Salário Atual/Total que dependem dela também ficam editáveis no painel —
aqui só exportamos valor_vendido e total_pedidos pra recalcular ao vivo.
"""

import json
import os

import openpyxl

CAMINHO_RESULTADO = r"C:\Users\edmar\Desktop\MELHORIA SALARIO\RESULTADO.xlsm"
CAMINHO_PAINEL_PILARES = r"c:\AutomacaoMaxGestao\painel_pilares\dados.json"

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
CAMINHO_SAIDA = os.path.join(PASTA_BASE, "dados.json")

# (chave, rótulo, aba de origem, taxa de comissão, meta de positivação)
CATEGORIAS = [
    ("bacon", "Bacon", "BACON", 0.03, 5),
    ("calabresa", "Calabresas", "CALABRESA", 0.03, 22),
    ("frescais", "Linguiças/Frescais", "FRESCAIS", 0.03, 15),
    ("paes", "Pães", "PAES", 0.03, 17),
    ("lactios", "Lácteos", "LACTIOS", 0.03, 23),
    ("batata", "Batata", "BATATA", 0.03, 15),
    ("bovino", "Bovino", "BOVINO", 0.015, 10),
    ("suino", "Suíno", "SUINO", 0.03, 6),
]

PREMIO_FIXO = 150  # teto de Dia 15 / Dia 30 / Campanha (mesmo valor pra todo mundo, confirmado em ITENS FOCO)


def _num(v):
    return v if isinstance(v, (int, float)) else 0


def _ler_categoria(wb, nome_aba):
    """{codigo: {"peso":..,"valor":..,"positivacao":..}} — cada aba de
    categoria tem 1 linha por RCA, sem cabeçalho: A=código, K=peso,
    L=valor realizado, D=positivação."""
    ws = wb[nome_aba]
    out = {}
    for row in ws.iter_rows(values_only=True):
        codigo = row[0]
        if codigo is None:
            continue
        out[int(codigo)] = {
            "peso": _num(row[10]),      # K
            "valor": _num(row[11]),     # L
            "positivacao": _num(row[3]),  # D
        }
    return out


def _ler_supervisores():
    """Cross-referencia com o painel 4 Pilares pra saber o supervisor de
    cada RCA (essa planilha de melhoria salarial não tem essa coluna)."""
    if not os.path.exists(CAMINHO_PAINEL_PILARES):
        return {}
    with open(CAMINHO_PAINEL_PILARES, "r", encoding="utf-8") as f:
        dados = json.load(f)
    return {int(r["codigo"]): r["supervisor"] for r in dados}


def extrair():
    wb = openpyxl.load_workbook(CAMINHO_RESULTADO, data_only=True)
    ws_geral = wb["ITENS FOCO"]

    dias_uteis = _num(ws_geral["O14"].value) or 23
    meta_pedidos_dia = _num(ws_geral["P17"].value) or 15
    taxa_padrao = _num(ws_geral["O11"].value) or 0.0178

    cache_categorias = {chave: _ler_categoria(wb, aba) for chave, _, aba, _, _ in CATEGORIAS}
    supervisores = _ler_supervisores()

    ws315 = wb["315"]
    rcas = []
    for row in ws315.iter_rows(values_only=True):
        codigo = row[0]
        if codigo is None:
            continue
        codigo = int(codigo)
        nome_bruto = str(row[1] or "").strip()
        if " - " in nome_bruto:
            nome_rca, rota = nome_bruto.split(" - ", 1)
        else:
            nome_rca, rota = nome_bruto, ""

        total_pedidos = _num(row[6])   # G
        valor_vendido = _num(row[9])   # J

        categorias_dados = {}
        soma_atual = soma_potencial = 0.0
        for chave, label, _aba, taxa, meta_posit in CATEGORIAS:
            info = cache_categorias[chave].get(codigo, {"peso": 0, "valor": 0, "positivacao": 0})
            comissao_atual = info["valor"] * taxa
            comissao_potencial = (
                (comissao_atual / info["positivacao"]) * meta_posit if info["positivacao"] else 0
            )
            categorias_dados[chave] = {
                "label": label,
                "peso": info["peso"],
                "valor": info["valor"],
                "positivacao": info["positivacao"],
                "meta_positivacao": meta_posit,
                "comissao_atual": comissao_atual,
                "comissao_potencial": comissao_potencial,
            }
            soma_atual += comissao_atual
            soma_potencial += comissao_potencial

        # Industrializado/Thermo: mesma base de cálculo da planilha —
        # 25% do valor vendido é a meta industrializado, 3% é a meta
        # thermo; a comissão-teto é 1% e 2% em cima dessas metas.
        meta_industrializado = valor_vendido * 0.25
        meta_thermo = valor_vendido * 0.03
        industrializado_potencial = meta_industrializado * 0.01
        thermo_potencial = meta_thermo * 0.02

        rcas.append({
            "codigo": codigo,
            "nome": nome_rca.strip(),
            "rota": rota.strip(),
            "supervisor": supervisores.get(codigo, "SEM SUPERVISOR"),
            "valor_vendido": valor_vendido,
            "total_pedidos": total_pedidos,
            "categorias": categorias_dados,
            "departamentos": {"atual": soma_atual, "potencial": soma_potencial},
            "industrializado_potencial": industrializado_potencial,
            "thermo_potencial": thermo_potencial,
            "premio_fixo": PREMIO_FIXO,
        })

    return rcas, {"dias_uteis": dias_uteis, "meta_pedidos_dia": meta_pedidos_dia, "taxa_padrao": taxa_padrao}


def main():
    rcas, constantes = extrair()
    saida = {"rcas": rcas, "constantes": constantes}
    with open(CAMINHO_SAIDA, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    print(f"{len(rcas)} RCAs extraídos. Salvo em: {CAMINHO_SAIDA}")


if __name__ == "__main__":
    main()
