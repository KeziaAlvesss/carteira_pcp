"""
PCP Bonsono - Painel de Programação da Carteira
================================================
Uso:
    pip install pandas openpyxl dash dash-bootstrap-components plotly
    python pcp_bonsono.py

Os dois arquivos Excel (Cabeçalho da Nota e Resultado da Query) são
carregados diretamente pela tela, usando os botões de upload no topo
do painel. Não é necessário colocar os arquivos na pasta do servidor.

Pedidos excluídos ficam salvos em pedidos_excluidos.json (mesma pasta).
"""

import sys, json, base64, io
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, State, ctx, dash_table, no_update, ALL
import dash
import dash_bootstrap_components as dbc
from datetime import date

# ── Configuração ───────────────────────────────────────────────────
# Arquivos locais usados apenas como carga inicial automática quando
# rodando no seu computador (opcional). Em produção (nuvem) eles não
# existirão e o painel vai aguardar o upload manual pela tela.
ARQUIVO_CABECALHO  = "Cabecalho_da_Nota.xlsx"
ARQUIVO_ITENS      = "Resultado_da_Query.xlsx"
ARQUIVO_EXCLUIDOS  = "pedidos_excluidos.json"
SKIPROWS           = 2
# ──────────────────────────────────────────────────────────────────

BASE = Path(__file__).parent

# ── Persistência de exclusões ──────────────────────────────────────
def carregar_excluidos():
    p = BASE / ARQUIVO_EXCLUIDOS
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
            return {str(n): "" for n in data}
        except:
            return {}
    return {}

def salvar_excluidos(excluidos: dict):
    (BASE / ARQUIVO_EXCLUIDOS).write_text(
        json.dumps(excluidos, ensure_ascii=False, indent=2), encoding="utf-8"
    )

# ── Processamento dos dados ─────────────────────────────────────────
def processar_dados(cab: pd.DataFrame, it: pd.DataFrame) -> pd.DataFrame:
    """Recebe os dois DataFrames brutos (já lidos do Excel) e devolve
    a base unificada e pronta para o painel."""
    cab["Nro. Único"] = cab["Nro. Único"].astype("Int64")
    it["NRO_UNICO"]   = it["NRO_UNICO"].astype("Int64")

    cab_cols = {
        "Nro. Único":                    "NRO_UNICO",
        "Nro. Nota":                     "NRO_NOTA_CAB",
        "Nome Parceiro (Parceiro)":      "PARCEIRO",
        "Previsão de entrega":           "PREV_ENTREGA",
        "Data Agendamento Entrega":      "DT_AGENDAMENTO",
        "Dt. Neg.":                      "DT_NEGOCIACAO",
        "Vlr. Nota":                     "VLR_NOTA",
        "Metro Cúbico":                  "M3",
        "Regiao Vendedor":               "REGIAO",
        "Apelido (Vendedor)":            "VENDEDOR",
        "Descrição (Tipo de Operação)":  "TIPO_OP",
        "Descrição (Centro de Resultado)": "CENTRO",
        "Status conferência":            "STATUS_CONF",
        "Status NF-e":                   "STATUS_NFE",
        "Análise Financeira":            "FINANCEIRO",
        "observação":                    "OBS",
    }
    cab_sel = cab[[c for c in cab_cols if c in cab.columns]].rename(columns=cab_cols)

    df = it.merge(cab_sel, on="NRO_UNICO", how="left")

    for col in ["PREV_ENTREGA", "DT_AGENDAMENTO", "DT_NEGOCIACAO"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    hoje = pd.Timestamp(date.today())
    df["PREV_ENTREGA_STR"]   = df["PREV_ENTREGA"].dt.strftime("%d/%m/%Y").fillna("Sem previsão")
    df["DT_AGENDAMENTO_STR"] = df["DT_AGENDAMENTO"].dt.strftime("%d/%m/%Y").fillna("Sem agendamento")
    df["MES_PREV"] = df["PREV_ENTREGA"].dt.to_period("M").astype(str).fillna("Sem previsão")
    df["ATRASO"]   = df["PREV_ENTREGA"] < hoje

    return df


def parse_upload_excel(contents: str) -> pd.DataFrame:
    """Decodifica o conteúdo de um dcc.Upload (base64) e lê como Excel."""
    _, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)
    return pd.read_excel(io.BytesIO(decoded), skiprows=SKIPROWS)


def df_to_json(df: pd.DataFrame) -> str:
    return df.to_json(date_format="iso", orient="split")


def df_from_json(data_json: str) -> pd.DataFrame:
    df = pd.read_json(io.StringIO(data_json), orient="split")
    for col in ["PREV_ENTREGA", "DT_AGENDAMENTO", "DT_NEGOCIACAO"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "NRO_UNICO" in df.columns:
        df["NRO_UNICO"] = df["NRO_UNICO"].astype("Int64")
    return df


def carregar_dados_locais():
    """Tenta carregar os arquivos da pasta local (uso apenas em
    desenvolvimento). Retorna None se não encontrar."""
    cab_path = BASE / ARQUIVO_CABECALHO
    it_path  = BASE / ARQUIVO_ITENS
    if not cab_path.exists() or not it_path.exists():
        return None
    try:
        cab = pd.read_excel(cab_path, skiprows=SKIPROWS)
        it  = pd.read_excel(it_path,  skiprows=SKIPROWS)
        return processar_dados(cab, it)
    except Exception:
        return None


CORES_GRUPO = {
    "CAMA BOX":              "#378ADD",
    "CAIXA BOX":             "#185FA5",
    "COLC ESPUMA COMERCIAL": "#1D9E75",
    "COLC MOLA COMERCIAL":   "#0F6E56",
    "COLC MOLA PREMIUM":     "#7F77DD",
    "BOX BAU":               "#EF9F27",
    "BOX BIPARTIDO":         "#BA7517",
    "TRAVESSEIROS":          "#D85A30",
    "PROTETORES":            "#888780",
    "SAIAS":                 "#B4B2A9",
    "BLOCO LAMINADO":        "#5F5E5A",
}

def fmt_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except:
        return "—"

def card_metric(titulo, valor, sub=None, cor="#378ADD"):
    return dbc.Card(dbc.CardBody([
        html.P(titulo, className="text-muted mb-1", style={"fontSize":"12px"}),
        html.H4(valor, style={"fontWeight":"500","color":cor}),
        html.P(sub, className="text-muted mb-0", style={"fontSize":"11px"}) if sub else None,
    ]), className="shadow-sm border-0 bg-light")

def grafico_vazio(texto):
    fig = go.Figure()
    fig.update_layout(
        height=320, plot_bgcolor="white", paper_bgcolor="white",
        xaxis={"visible": False}, yaxis={"visible": False},
        annotations=[dict(text=texto, showarrow=False, font=dict(size=13, color="#aaa"))],
    )
    return fig

# ── Carga inicial (só funciona localmente, se os arquivos existirem) ─
df_inicial = carregar_dados_locais()
DADOS_INICIAIS = df_to_json(df_inicial) if df_inicial is not None else None
REGIOES_INICIAIS  = sorted(df_inicial["REGIAO"].dropna().unique())     if df_inicial is not None else []
GRUPOS_INICIAIS   = sorted(df_inicial["GRUPO_PROD"].dropna().unique()) if df_inicial is not None else []
TIPOS_INICIAIS    = sorted(df_inicial["TIPO_OP"].dropna().unique())    if df_inicial is not None else []

UPLOAD_STYLE = {
    "width": "100%", "height": "54px", "lineHeight": "54px",
    "borderWidth": "1px", "borderStyle": "dashed", "borderRadius": "6px",
    "borderColor": "#ccc", "textAlign": "center", "fontSize": "13px",
    "color": "#888", "cursor": "pointer",
}

# ── App ────────────────────────────────────────────────────────────
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY], title="PCP Bonsono")
server = app.server  # necessário para o gunicorn (Render) enxergar o app

app.layout = dbc.Container([

    dbc.Row(dbc.Col(html.H4("📦 PCP Bonsono — Painel da Carteira",
                            className="my-3 fw-normal text-secondary"))),

    # ── Upload dos arquivos ──────────────────────────────────────────
    dbc.Card(dbc.CardBody([
        dbc.Row([
            dbc.Col([
                html.Label("Cabeçalho da Nota (.xlsx)", className="small text-muted"),
                dcc.Upload(
                    id="upload-cab",
                    children=html.Div(id="upload-cab-label", children="Arraste o arquivo ou clique para selecionar"),
                    style=UPLOAD_STYLE, multiple=False,
                ),
            ], md=6),
            dbc.Col([
                html.Label("Resultado da Query (.xlsx)", className="small text-muted"),
                dcc.Upload(
                    id="upload-itens",
                    children=html.Div(id="upload-itens-label", children="Arraste o arquivo ou clique para selecionar"),
                    style=UPLOAD_STYLE, multiple=False,
                ),
            ], md=6),
        ]),
        html.Div(id="upload-status", className="mt-2"),
    ]), className="mb-3 border-0 shadow-sm"),

    # ── Filtros ────────────────────────────────────────────────────
    dbc.Card(dbc.CardBody(dbc.Row([
        dbc.Col([
            html.Label("Região do Vendedor", className="small text-muted"),
            dcc.Dropdown(REGIOES_INICIAIS, multi=True, placeholder="Todas", id="f-regiao"),
        ], md=3),
        dbc.Col([
            html.Label("Grupo de Produto", className="small text-muted"),
            dcc.Dropdown(GRUPOS_INICIAIS, multi=True, placeholder="Todos", id="f-grupo"),
        ], md=3),
        dbc.Col([
            html.Label("Tipo de Operação", className="small text-muted"),
            dcc.Dropdown(TIPOS_INICIAIS, multi=True, placeholder="Todos", id="f-tipo"),
        ], md=3),
        dbc.Col([
            html.Label("Data Agendamento Entrega", className="small text-muted"),
            dcc.DatePickerRange(id="f-data", display_format="DD/MM/YYYY",
                                start_date_placeholder_text="De",
                                end_date_placeholder_text="Até"),
        ], md=3),
    ])), className="mb-3 border-0 shadow-sm"),

    # ── KPIs ───────────────────────────────────────────────────────
    dbc.Row([
        dbc.Col(id="kpi-pedidos", md=2),
        dbc.Col(id="kpi-itens",   md=2),
        dbc.Col(id="kpi-valor",   md=3),
        dbc.Col(id="kpi-m3",      md=2),
        dbc.Col(id="kpi-atraso",  md=3),
    ], className="mb-3 g-2"),

    # ── Gráficos ───────────────────────────────────────────────────
    dbc.Row([
        dbc.Col(dcc.Graph(id="g-grupo",   config={"displayModeBar":False}), md=6),
        dbc.Col(dcc.Graph(id="g-produto", config={"displayModeBar":False}), md=6),
    ], className="mb-3"),

    # ── Tabela ─────────────────────────────────────────────────────
    dbc.Card([
        dbc.CardHeader(html.B("📋 Pedidos detalhados")),
        dbc.CardBody(html.Div(id="tabela-pedidos")),
    ], className="mb-4 border-0 shadow-sm"),

    # Stores
    dcc.Store(id="store-excluidos"),
    dcc.Store(id="store-dados", data=DADOS_INICIAIS),

], fluid=True)


# ── Callback: processa os uploads e gera a base unificada ──────────
@app.callback(
    Output("store-dados",       "data"),
    Output("upload-status",     "children"),
    Output("upload-cab-label",  "children"),
    Output("upload-itens-label","children"),
    Output("f-regiao",          "options"),
    Output("f-grupo",           "options"),
    Output("f-tipo",            "options"),
    Input("upload-cab",         "contents"),
    Input("upload-itens",       "contents"),
    State("upload-cab",         "filename"),
    State("upload-itens",       "filename"),
    prevent_initial_call=True,
)
def processar_uploads(cab_contents, it_contents, cab_name, it_name):
    label_cab = f"✓ {cab_name}" if cab_name else no_update
    label_it  = f"✓ {it_name}"  if it_name  else no_update

    if not cab_contents or not it_contents:
        status = dbc.Alert(
            "Carregue os dois arquivos (Cabeçalho da Nota e Resultado da Query) para gerar o painel.",
            color="warning", className="py-2 mb-0", style={"fontSize":"13px"},
        )
        return no_update, status, label_cab, label_it, no_update, no_update, no_update

    try:
        cab = parse_upload_excel(cab_contents)
        it  = parse_upload_excel(it_contents)
        df  = processar_dados(cab, it)
    except Exception as e:
        status = dbc.Alert(f"Erro ao processar os arquivos: {e}",
                            color="danger", className="py-2 mb-0", style={"fontSize":"13px"})
        return no_update, status, label_cab, label_it, no_update, no_update, no_update

    regioes = sorted(df["REGIAO"].dropna().unique())
    grupos  = sorted(df["GRUPO_PROD"].dropna().unique())
    tipos   = sorted(df["TIPO_OP"].dropna().unique())

    status = dbc.Alert(
        f"✅ Dados carregados: {df['NRO_UNICO'].nunique()} pedidos, {len(df)} itens.",
        color="success", className="py-2 mb-0", style={"fontSize":"13px"},
    )
    return df_to_json(df), status, label_cab, label_it, regioes, grupos, tipos


# ── Callback: remover pedido individual da lista ───────────────────
@app.callback(
    Output("store-excluidos", "data", allow_duplicate=True),
    Input({"type": "btn-remover", "index": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def remover_individual(n_clicks):
    triggered = ctx.triggered_id
    if not triggered or not any(n_clicks):
        return no_update
    nro_remover = str(triggered["index"])
    excluidos = carregar_excluidos()
    excluidos.pop(nro_remover, None)
    salvar_excluidos(excluidos)
    return {"ts": pd.Timestamp.now().isoformat()}


# ── Callback: painel principal ─────────────────────────────────────
@app.callback(
    Output("kpi-pedidos",    "children"),
    Output("kpi-itens",      "children"),
    Output("kpi-valor",      "children"),
    Output("kpi-m3",         "children"),
    Output("kpi-atraso",     "children"),
    Output("g-grupo",        "figure"),
    Output("g-produto",      "figure"),
    Output("tabela-pedidos", "children"),
    Input("f-regiao",        "value"),
    Input("f-grupo",         "value"),
    Input("f-tipo",          "value"),
    Input("f-data",          "start_date"),
    Input("f-data",          "end_date"),
    Input("store-excluidos", "data"),
    Input("store-dados",     "data"),
)
def atualizar_painel(regioes, grupos, tipos, dt_ini, dt_fim, _store_excl, data_json):
    if not data_json:
        vazio = card_metric("—", "—")
        msg = html.Div("Carregue os dois arquivos Excel acima para visualizar o painel.",
                        className="text-muted text-center p-4")
        return vazio, vazio, vazio, vazio, vazio, grafico_vazio("Aguardando dados"), grafico_vazio("Aguardando dados"), msg

    df_global = df_from_json(data_json)
    excluidos = set(carregar_excluidos().keys())

    d = df_global.copy()

    # Aplica exclusões
    if excluidos:
        d = d[~d["NRO_UNICO"].astype(str).isin(excluidos)]

    # Aplica filtros
    if regioes: d = d[d["REGIAO"].isin(regioes)]
    if grupos:  d = d[d["GRUPO_PROD"].isin(grupos)]
    if tipos:   d = d[d["TIPO_OP"].isin(tipos)]
    if dt_ini:  d = d[d["DT_AGENDAMENTO"] >= pd.Timestamp(dt_ini)]
    if dt_fim:  d = d[d["DT_AGENDAMENTO"] <= pd.Timestamp(dt_fim)]

    ped = d.drop_duplicates("NRO_UNICO")

    n_ped     = ped["NRO_UNICO"].nunique()
    n_it      = len(d)
    valor_tot = ped["VLR_NOTA"].sum()
    m3_tot    = ped["M3"].sum()
    atrasados = int(ped["ATRASO"].sum())
    excl_badge = f" ({len(excluidos)} em fab.)" if excluidos else ""

    kpi_ped  = card_metric("Pedidos",       str(n_ped) + excl_badge, f"{n_it} linhas de item")
    kpi_it   = card_metric("Itens",         str(n_it), f"{d['PRODUTO'].nunique()} produtos distintos")
    kpi_val  = card_metric("Valor total",   fmt_brl(valor_tot), "soma das notas")
    kpi_m3   = card_metric("Volume",        f"{m3_tot:,.1f} m³")
    kpi_atr  = card_metric("Atrasados",     str(atrasados), "previsão vencida", cor="#E24B4A")

    # Grupo de produto
    grp = d.groupby("GRUPO_PROD")["QTDNEG"].sum().reset_index().sort_values("QTDNEG", ascending=True)
    fig_grupo = go.Figure(go.Bar(
        x=grp["QTDNEG"], y=grp["GRUPO_PROD"], orientation="h",
        marker_color=[CORES_GRUPO.get(g,"#888780") for g in grp["GRUPO_PROD"]],
        text=grp["QTDNEG"].astype(int), textposition="outside",
    ))
    fig_grupo.update_layout(title="Qtd. por grupo de produto", height=320,
        margin=dict(l=10,r=30,t=40,b=10), plot_bgcolor="white", paper_bgcolor="white")

    # Top 10 produtos
    top = (d.groupby("PRODUTO")["QTDNEG"].sum().reset_index()
             .sort_values("QTDNEG", ascending=False).head(10).sort_values("QTDNEG"))
    fig_prod = go.Figure(go.Bar(
        x=top["QTDNEG"], y=top["PRODUTO"], orientation="h",
        marker_color="#7F77DD", text=top["QTDNEG"].astype(int), textposition="outside",
    ))
    fig_prod.update_layout(title="Top 10 produtos", height=320,
        margin=dict(l=10,r=40,t=40,b=10), plot_bgcolor="white", paper_bgcolor="white")

    # Tabela
    tab = (ped[["NRO_UNICO","PARCEIRO","REGIAO","TIPO_OP",
                "DT_AGENDAMENTO_STR","VLR_NOTA","M3","STATUS_CONF"]]
           .rename(columns={
               "NRO_UNICO":"Nro. Único","PARCEIRO":"Parceiro","REGIAO":"Região",
               "TIPO_OP":"Tipo de Operação","DT_AGENDAMENTO_STR":"Dt. Agend. Entrega",
               "VLR_NOTA":"Valor (R$)","M3":"m³","STATUS_CONF":"Status Conf.",
           }).sort_values("Dt. Agend. Entrega"))
    tab["Valor (R$)"] = tab["Valor (R$)"].apply(lambda x: fmt_brl(x) if pd.notna(x) else "")
    tab["m³"] = tab["m³"].apply(lambda x: f"{x:,.3f}".replace(",","X").replace(".",",").replace("X",".") if pd.notna(x) and x>0 else "")

    tabela = dash_table.DataTable(
        data=tab.fillna("").to_dict("records"),
        columns=[{"name":c,"id":c} for c in tab.columns],
        page_size=15, sort_action="native", filter_action="native",
        style_table={"overflowX":"auto"},
        style_header={"backgroundColor":"#f8f9fa","fontWeight":"500","fontSize":"12px","border":"1px solid #dee2e6"},
        style_cell={"fontSize":"12px","padding":"8px 10px","border":"1px solid #dee2e6",
                    "maxWidth":"200px","overflow":"hidden","textOverflow":"ellipsis"},
        style_data_conditional=[{
            "if":{"filter_query":'{Dt. Agend. Entrega} < "' + date.today().strftime("%d/%m/%Y") + '"'},
            "backgroundColor":"#fff5f5","color":"#A32D2D",
        }],
    )

    return kpi_ped, kpi_it, kpi_val, kpi_m3, kpi_atr, fig_grupo, fig_prod, tabela


if __name__ == "__main__":
    import os
    porta = int(os.environ.get("PORT", 8050))
    print("\n✅ PCP Bonsono iniciado!")
    print(f"   Acesse: http://127.0.0.1:{porta}\n")
    app.run(debug=False, host="0.0.0.0", port=porta)