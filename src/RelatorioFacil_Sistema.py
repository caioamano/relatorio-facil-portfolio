
import argparse
from datetime import datetime
import sys
import traceback
import re
import unicodedata
from pathlib import Path
from collections import defaultdict

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, Image, PageBreak
)


# ============================================================
# CONFIGURACAO VISUAL
# ============================================================

NAVY = colors.HexColor("#16324F")
TEAL = colors.HexColor("#0F766E")
LIGHT = colors.HexColor("#F3F7FA")
GRAY = colors.HexColor("#64748B")
BORDER = colors.HexColor("#D9E2EC")
WHITE = colors.white
RED = colors.HexColor("#B91C1C")
GREEN = colors.HexColor("#166534")


# ============================================================
# IMPORTACAO E VALIDACAO
# ============================================================

ALIASES = {
    "sku": ["sku", "codigo", "codigoproduto"],
    "produto": ["produto", "nomeproduto", "descricao"],
    "categoria": ["categoria"],
    "custounitario": ["custounitario", "custo", "custoindividual"],
    "precovenda": ["precovenda", "precodevenda", "preco"],
    "estoqueinicial": ["estoqueinicial"],
    "estoqueminimo": ["estoqueminimo"],
    "datacadastro": ["datacadastro"],
    "data": ["data", "datavenda"],
    "vendaid": ["vendaid", "idvenda"],
    "quantidade": ["quantidade", "qtd"],
    "precounitario": ["precounitario", "precopeca"],
    "desconto": ["desconto"],
    "totalvenda": ["totalvenda", "valorvenda", "total"],
    "formapagamento": ["formapagamento", "pagamento", "metodopagamento"]
}

MAPA_COLUNAS = {
    unicodedata.normalize("NFKD", nome).encode("ascii", "ignore")
    .decode().lower(): canonico
    for canonico, aliases in ALIASES.items()
    for nome in aliases
}


def normalizar_nome(nome):
    texto = unicodedata.normalize("NFKD", str(nome))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", texto.lower())


def converter_numero(valor):
    if pd.isna(valor) or valor == "":
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    texto = texto.replace("R$", "").replace(" ", "")

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    return float(texto)


def padronizar_colunas(df):
    df = df.copy()
    novas = []

    for coluna in df.columns:
        chave = normalizar_nome(coluna)
        novas.append(MAPA_COLUNAS.get(chave, chave))

    df.columns = novas
    return df


def validar_colunas(df, obrigatorias, origem):
    faltantes = [c for c in obrigatorias if c not in df.columns]

    if faltantes:
        raise ValueError(
            f"Faltam colunas em {origem}: {', '.join(faltantes)}"
        )


def ler_csv(caminho):
    return pd.read_csv(
        caminho,
        encoding="utf-8-sig",
        sep=None,
        engine="python"
    )


def carregar_entradas(excel=None, csv_produtos=None, csv_vendas=None):
    if excel:
        caminho = Path(excel)

        if not caminho.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

        produtos = pd.read_excel(
            caminho, sheet_name="BASE_PRODUTOS"
        )
        vendas = pd.read_excel(
            caminho, sheet_name="BASE_VENDAS"
        )

    elif csv_produtos and csv_vendas:
        produtos = ler_csv(csv_produtos)
        vendas = ler_csv(csv_vendas)

    else:
        raise ValueError(
            "Informe um arquivo Excel ou os dois arquivos CSV."
        )

    produtos = preparar_produtos(produtos)
    vendas = preparar_vendas(vendas, produtos)

    return produtos, vendas


def preparar_produtos(df):
    df = padronizar_colunas(df)

    validar_colunas(
        df,
        [
            "sku", "produto", "custounitario",
            "precovenda", "estoqueinicial", "estoqueminimo"
        ],
        "produtos"
    )

    if "categoria" not in df.columns:
        df["categoria"] = "Sem categoria"

    df["sku"] = df["sku"].fillna("").astype(str).str.strip()
    df["produto"] = df["produto"].fillna("").astype(str).str.strip()

    df = df[
        (df["sku"] != "")
        & (df["sku"].str.lower() != "nan")
        & (df["produto"] != "")
        & (df["produto"].str.lower() != "nan")
    ].copy()

    df["categoria"] = df["categoria"].fillna("Sem categoria").astype(str)

    if df["sku"].duplicated().any():
        raise ValueError("Existem SKUs duplicados na base de produtos.")

    for coluna in [
        "custounitario", "precovenda",
        "estoqueinicial", "estoqueminimo"
    ]:
        df[coluna] = df[coluna].apply(converter_numero)

        if (df[coluna] < 0).any():
            raise ValueError(f"Valores negativos em {coluna}.")

    for coluna in ["estoqueinicial", "estoqueminimo"]:
        if (df[coluna] % 1 != 0).any():
            raise ValueError(f"{coluna} precisa usar números inteiros.")

    return df.reset_index(drop=True)


def preparar_vendas(df, produtos):
    df = padronizar_colunas(df)

    validar_colunas(
        df,
        [
            "data", "sku", "quantidade", "precounitario",
            "custounitario", "desconto", "formapagamento"
        ],
        "vendas"
    )

    df["data"] = pd.to_datetime(
        df["data"], dayfirst=True, errors="coerce"
    )

    if df["data"].isna().any():
        raise ValueError("Existem datas inválidas na base de vendas.")

    df["sku"] = df["sku"].fillna("").astype(str).str.strip()

    # Recusar vendas sem produto correspondente no cadastro.
    skus_conhecidos = set(produtos["sku"].astype(str))

    skus_desconhecidos = sorted({
        sku for sku in df["sku"].unique()
        if sku not in skus_conhecidos
    })

    if skus_desconhecidos:
        lista = ", ".join(skus_desconhecidos[:10])
        raise ValueError(
            "Existem vendas com SKU sem cadastro de produto: "
            + lista
        )


    for coluna in [
        "quantidade", "precounitario",
        "custounitario", "desconto"
    ]:
        df[coluna] = df[coluna].apply(converter_numero)

    if (
        (df["quantidade"] <= 0).any()
        or (df["quantidade"] % 1 != 0).any()
        or (df["precounitario"] < 0).any()
        or (df["custounitario"] < 0).any()
        or (df["desconto"] < 0).any()
    ):
        raise ValueError("Existem valores inválidos na base de vendas.")

    subtotal = df["quantidade"] * df["precounitario"]

    if (df["desconto"] > subtotal).any():
        raise ValueError("Há um desconto maior que o subtotal.")

    # Não confiar em fórmulas preexistentes no Excel.
    df["total_calculado"] = subtotal - df["desconto"]

    if "vendaid" not in df.columns:
        df["vendaid"] = ""

    df["vendaid"] = df["vendaid"].fillna("").astype(str).str.strip()

    # Gerar IDs para linhas que não possuem identificador.
    for indice in df.index:
        if not df.at[indice, "vendaid"]:
            df.at[indice, "vendaid"] = f"REGISTRO-{indice + 1}"

    if "produto" not in df.columns:
        df["produto"] = ""

    nomes_por_sku = dict(zip(produtos["sku"], produtos["produto"]))

    df["produto"] = df["produto"].fillna("").astype(str).str.strip()

    for indice in df.index:
        sku = df.at[indice, "sku"]

        if not df.at[indice, "produto"]:
            df.at[indice, "produto"] = nomes_por_sku.get(sku, sku)

    df["formapagamento"] = (
        df["formapagamento"].fillna("Não informado").astype(str)
    )

    return df.reset_index(drop=True)


# ============================================================
# FORMATACAO DO RELATORIO
# ============================================================

def moeda(valor):
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def gerar_relatorio(
    produtos,
    todas_vendas,
    inicio,
    fim,
    caminho_pdf
):
    inicio = pd.Timestamp(inicio)
    fim = pd.Timestamp(fim)

    if fim < inicio:
        raise ValueError("A data final é anterior à data inicial.")

    # O período final inclui todo o dia selecionado.
    periodo = todas_vendas[
        (todas_vendas["data"] >= inicio)
        & (todas_vendas["data"] < fim + pd.Timedelta(days=1))
    ].copy()

    faturamento = periodo["total_calculado"].sum()
    custo = (periodo["quantidade"] * periodo["custounitario"]).sum()
    lucro = faturamento - custo
    unidades = int(periodo["quantidade"].sum())
    quantidade_vendas = periodo["vendaid"].nunique()
    ticket = faturamento / quantidade_vendas if quantidade_vendas else 0
    margem = lucro / faturamento if faturamento else 0

    # Agrupar produtos.
    resumo = defaultdict(
        lambda: {"unidades": 0, "faturamento": 0.0, "custo": 0.0}
    )

    for _, v in periodo.iterrows():
        sku = v["sku"]
        resumo[sku]["unidades"] += int(v["quantidade"])
        resumo[sku]["faturamento"] += float(v["total_calculado"])
        resumo[sku]["custo"] += (
            float(v["quantidade"]) * float(v["custounitario"])
        )

    nomes = dict(zip(produtos["sku"], produtos["produto"]))

    ranking = []

    for sku, info in resumo.items():
        ranking.append({
            "sku": sku,
            "produto": nomes.get(sku, sku),
            "unidades": info["unidades"],
            "faturamento": info["faturamento"],
            "lucro": info["faturamento"] - info["custo"]
        })

    ranking.sort(
        key=lambda x: (
            -x["unidades"], -x["faturamento"], x["produto"]
        )
    )
    ranking = ranking[:10]

    pagamentos = (
        periodo.groupby("formapagamento")["total_calculado"]
        .sum().sort_values(ascending=False)
    )

    por_dia = periodo.groupby(
        periodo["data"].dt.strftime("%d/%m")
    )["total_calculado"].sum()

    # Estoque calculado com todas as vendas fornecidas no arquivo.
    vendido_historico = (
        todas_vendas.groupby("sku")["quantidade"].sum().to_dict()
    )

    estoques = []

    for _, p in produtos.iterrows():
        sku = p["sku"]
        saldo = int(p["estoqueinicial"] - vendido_historico.get(sku, 0))
        minimo = int(p["estoqueminimo"])

        estoques.append({
            "produto": p["produto"],
            "saldo": saldo,
            "minimo": minimo,
            "situacao": "ABAIXO DO MÍNIMO" if saldo <= minimo else "OK"
        })

    estoques.sort(
        key=lambda x: (
            x["situacao"] != "ABAIXO DO MÍNIMO",
            x["produto"]
        )
    )

    # Criar gráficos temporários.
    pasta_temp = Path(caminho_pdf).parent / "_graficos_relatorio_facil"
    pasta_temp.mkdir(parents=True, exist_ok=True)

    grafico_dia = pasta_temp / "vendas_diarias.png"
    grafico_pagamento = pasta_temp / "pagamentos.png"
    grafico_top = pasta_temp / "top_produtos.png"

    fig, ax = plt.subplots(figsize=(8, 3.4))

    if not por_dia.empty:
        ax.bar(por_dia.index, por_dia.values, color="#0F766E")
        ax.set_title("Faturamento por dia")
        ax.set_ylabel("Valor (R$)")
        ax.grid(axis="y", alpha=0.2)
        ax.tick_params(axis="x", rotation=35)
    else:
        ax.text(0.5, 0.5, "Sem vendas no período",
                ha="center", va="center")
        ax.set_axis_off()

    fig.tight_layout()
    fig.savefig(grafico_dia, dpi=160, bbox_inches="tight")
    plt.close(fig)

    
    # Grafico de faturamento por forma de pagamento
    # Figura quadrada para manter a rosca circular.
    fig, ax = plt.subplots(figsize=(4.4, 4.4))

    if not pagamentos.empty and pagamentos.sum() > 0:
        ax.pie(
            pagamentos.values,
            labels=pagamentos.index,
            autopct="%1.0f%%",
            pctdistance=0.78,
            startangle=90,
            textprops={
                "color": "white",
                "weight": "bold",
                "fontsize": 10
            },
            wedgeprops={
                "width": 0.42,
                "edgecolor": "white",
                "linewidth": 1
            }
        )

        ax.set_title(
            "Por forma de pagamento",
            pad=14
        )

        ax.set_aspect("equal", adjustable="box")

    else:
        ax.text(
            0.5, 0.5,
            "Sem dados",
            ha="center",
            va="center"
        )
        ax.set_axis_off()

    fig.tight_layout()

    fig.savefig(
        grafico_pagamento,
        dpi=160,
        bbox_inches="tight"
    )

    plt.close(fig)


    fig, ax = plt.subplots(figsize=(8, 3.7))

    if ranking:
        ax.barh(
            [x["produto"] for x in ranking][::-1],
            [x["unidades"] for x in ranking][::-1],
            color="#2563EB"
        )
        ax.set_xlabel("Unidades vendidas")
        ax.set_title("Top 10 produtos mais vendidos")
        ax.grid(axis="x", alpha=0.2)
    else:
        ax.text(0.5, 0.5, "Sem vendas no período",
                ha="center", va="center")
        ax.set_axis_off()

    fig.tight_layout()
    fig.savefig(grafico_top, dpi=160, bbox_inches="tight")
    plt.close(fig)

    # Criar PDF.
    caminho_pdf = Path(caminho_pdf)
    caminho_pdf.parent.mkdir(parents=True, exist_ok=True)

    pagina = landscape(A4)
    largura_util = pagina[0] - 28 * mm

    estilos = getSampleStyleSheet()
    estilos.add(ParagraphStyle(
        name="TituloRF", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=21,
        leading=25, textColor=NAVY, alignment=0
    ))
    estilos.add(ParagraphStyle(
        name="SubRF", parent=estilos["Normal"],
        fontSize=9, leading=12, textColor=GRAY
    ))
    estilos.add(ParagraphStyle(
        name="SecRF", parent=estilos["Heading2"],
        fontSize=12, textColor=TEAL, spaceBefore=5, spaceAfter=7
    ))
    estilos.add(ParagraphStyle(
        name="KpiLabRF", parent=estilos["Normal"],
        fontSize=7, textColor=GRAY, alignment=TA_CENTER
    ))
    estilos.add(ParagraphStyle(
        name="KpiValRF", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=12,
        textColor=NAVY, alignment=TA_CENTER
    ))
    estilos.add(ParagraphStyle(
        name="TabelaRF", parent=estilos["Normal"],
        fontSize=8, leading=10
    ))
    estilos.add(ParagraphStyle(
        name="NotaRF", parent=estilos["Normal"],
        fontSize=8, leading=11, textColor=GRAY
    ))

    story = [
        Paragraph("RELATÓRIO FÁCIL", estilos["TituloRF"]),
        Paragraph("RELATÓRIO GERENCIAL DE VENDAS", estilos["SecRF"]),
        Paragraph(
            f"Período: {inicio.strftime('%d/%m/%Y')} "
            f"a {fim.strftime('%d/%m/%Y')}",
            estilos["SubRF"]
        ),
        Spacer(1, 5 * mm)
    ]

    indicadores = [
        ("FATURAMENTO", moeda(faturamento)),
        ("LUCRO BRUTO ESTIMADO", moeda(lucro)),
        ("UNIDADES VENDIDAS", str(unidades)),
        ("VENDAS REGISTRADAS", str(quantidade_vendas)),
        ("TICKET MÉDIO", moeda(ticket)),
        ("MARGEM BRUTA", f"{margem:.1%}".replace(".", ","))
    ]

    labels = [
        Paragraph(a, estilos["KpiLabRF"]) for a, _ in indicadores
    ]
    valores = [
        Paragraph(b, estilos["KpiValRF"]) for _, b in indicadores
    ]

    tabela_kpi = Table(
        [labels, valores],
        colWidths=[largura_util / 6] * 6,
        rowHeights=[22, 34]
    )
    tabela_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.5, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    story += [tabela_kpi, Spacer(1, 5 * mm)]

    # Graficos da primeira pagina: limitar dimensoes sem deformar.
    img_dia = Image(
        str(grafico_dia),
        width=145 * mm,
        height=58 * mm,
        kind="proportional"
    )

    img_pag = Image(
        str(grafico_pagamento),
        width=80 * mm,
        height=58 * mm,
        kind="proportional"
    )


    tabela_graficos = Table(
        [[img_dia, img_pag]],
        colWidths=[largura_util * 0.64, largura_util * 0.36]
    )
    tabela_graficos.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))

    story += [
        tabela_graficos,
        Spacer(1, 3 * mm),
        Paragraph(
            "Lucro bruto estimado: faturamento menos custo dos produtos "
            "vendidos. Não inclui despesas operacionais, impostos ou "
            "outros custos. O estoque considera o histórico de vendas "
            "fornecido e não inclui reposições, devoluções ou ajustes.",
            estilos["NotaRF"]
        ),
        PageBreak(),
        Paragraph("PRODUTOS E ESTOQUE", estilos["TituloRF"]),
        Paragraph("Ranking das vendas no período selecionado.", estilos["SubRF"]),
        Spacer(1, 4 * mm)
    ]

    dados = [[
        "Produto", "Unidades", "Faturamento", "Lucro bruto estimado"
    ]]

    for item in ranking:
        dados.append([
            Paragraph(escape(str(item["produto"])), estilos["TabelaRF"]),
            str(item["unidades"]),
            moeda(item["faturamento"]),
            moeda(item["lucro"])
        ])

    if not ranking:
        dados.append(["Nenhuma venda no período", "-", moeda(0), moeda(0)])

    tabela_ranking = Table(
        dados,
        colWidths=[
            largura_util * 0.38,
            largura_util * 0.14,
            largura_util * 0.22,
            largura_util * 0.26
        ],
        repeatRows=1
    )
    tabela_ranking.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT]),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story += [tabela_ranking, Spacer(1, 5 * mm)]

    img_top = Image(str(grafico_top), width=155 * mm, height=69 * mm)

    dados_estoque = [["Produto", "Saldo", "Mínimo", "Situação"]]

    for item in estoques[:12]:
        dados_estoque.append([
            Paragraph(escape(str(item["produto"])), estilos["TabelaRF"]),
            str(item["saldo"]),
            str(item["minimo"]),
            item["situacao"]
        ])

    if not estoques:
        dados_estoque.append(["Nenhum produto", "-", "-", "-"])

    tabela_estoque = Table(
        dados_estoque,
        colWidths=[42 * mm, 13 * mm, 13 * mm, 29 * mm],
        repeatRows=1
    )

    regras = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.35, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (2, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]

    for linha, item in enumerate(estoques[:12], start=1):
        cor = RED if item["situacao"] == "ABAIXO DO MÍNIMO" else GREEN
        regras.append(("TEXTCOLOR", (3, linha), (3, linha), cor))

    tabela_estoque.setStyle(TableStyle(regras))

    detalhes = Table(
        [[img_top, tabela_estoque]],
        colWidths=[largura_util * 0.64, largura_util * 0.36]
    )
    detalhes.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))

    story += [
        detalhes,
        Spacer(1, 3 * mm),
        Paragraph(
            "O saldo de estoque é estimado por estoque inicial menos "
            "todas as vendas presentes no arquivo. São exibidos até "
            "12 produtos nesta tabela de estoque.",
            estilos["NotaRF"]
        )
    ]

    def rodape(canvas, documento):
        canvas.saveState()
        canvas.setStrokeColor(BORDER)
        canvas.line(
            14 * mm, 10 * mm,
            pagina[0] - 14 * mm, 10 * mm
        )
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRAY)
        canvas.drawString(
            14 * mm, 6 * mm,
            "Relatório Fácil | Valores estimados a partir dos dados informados"
        )
        canvas.drawRightString(
            pagina[0] - 14 * mm, 6 * mm,
            f"Página {documento.page}"
        )
        canvas.restoreState()

    doc = SimpleDocTemplate(
        str(caminho_pdf),
        pagesize=pagina,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm,
        title="Relatório Fácil - Relatório gerencial",
        author="Relatório Fácil"
    )

    doc.build(story, onFirstPage=rodape, onLaterPages=rodape)

    # O Excel verificara o codigo de saida e a existencia do PDF.


def main():
    parser = argparse.ArgumentParser(
        description="Relatório Fácil - gerador de relatórios comerciais"
    )

    parser.add_argument("--excel", help="Arquivo Excel com as duas bases")
    parser.add_argument("--csv-produtos", help="Arquivo CSV de produtos")
    parser.add_argument("--csv-vendas", help="Arquivo CSV de vendas")
    parser.add_argument("--inicio", help="Data inicial YYYY-MM-DD")
    parser.add_argument("--fim", help="Data final YYYY-MM-DD")
    parser.add_argument(
        "--saida",
        default="RelatorioFacil_Relatorio.pdf",
        help="Caminho do PDF gerado"
    )

    args = parser.parse_args()

    produtos, vendas = carregar_entradas(
        excel=args.excel,
        csv_produtos=args.csv_produtos,
        csv_vendas=args.csv_vendas
    )

    if bool(args.inicio) != bool(args.fim):
        raise ValueError(
            "Informe as duas datas: --inicio e --fim."
        )

    if args.inicio and args.fim:
        inicio = pd.Timestamp(args.inicio)
        fim = pd.Timestamp(args.fim)
    elif not vendas.empty:
        inicio = vendas["data"].min().normalize()
        fim = vendas["data"].max().normalize()
    else:
        raise ValueError(
            "Não existem vendas. Informe um período com --inicio e --fim."
        )

    gerar_relatorio(
        produtos, vendas, inicio, fim, args.saida
    )



def registrar_erro(erro, detalhes):
    """Salva detalhes tecnicos para facilitar o diagnostico."""

    destino = Path.cwd() / "RelatorioFacil_erro.log"

    try:
        argumentos = sys.argv[1:]

        if "--saida" in argumentos:
            indice = argumentos.index("--saida")

            if indice + 1 < len(argumentos):
                caminho_pdf = Path(argumentos[indice + 1])
                destino = (
                    caminho_pdf.parent / "RelatorioFacil_erro.log"
                )

    except Exception:
        pass

    try:
        destino.parent.mkdir(parents=True, exist_ok=True)

        with destino.open("a", encoding="utf-8") as arquivo:
            arquivo.write("\n" + "=" * 70 + "\n")
            arquivo.write(
                "Data e hora: "
                + datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                + "\n"
            )
            arquivo.write(
                f"Erro: {type(erro).__name__}: {erro}\n\n"
            )
            arquivo.write(detalhes)
            arquivo.write("\n")

    except Exception:
        # O programa ainda deve terminar com codigo diferente de zero.
        pass


if __name__ == "__main__":
    try:
        main()

    except Exception as erro:
        registrar_erro(erro, traceback.format_exc())
        raise SystemExit(1)
