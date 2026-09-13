# -*- coding: utf-8 -*-
"""BI Dashboard - Investigacao de Perdas (BR) - GCMS / LYNCEUS / BR CONCESSIONS"""

from __future__ import annotations

import datetime as dt
import warnings
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Configuracao
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="BI Investigacao - BR",
    page_icon=":bar_chart:",
    layout="wide",
)

PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = PROJECT_DIR / "data"
DEFAULT_FILE = DATA_DIR / "geral_01.xlsx"

LYN_DATE_COLS = [
    "Creation Date UTC",
    "Submitted Time",
    " Inv. Start Date UTC",
    "Inv. End Date UTC",
    "Approved Time",
    "Closed Date UTC",
    "Last Update Date UTC",
    "First Closed Date",
    "Event start timestamp UTC",
    "Event end timestamp UTC",
]

TEMPLATE = "plotly_white"
PALETTE = px.colors.qualitative.Set2

def parse_mixed_dates(s: pd.Series) -> pd.Series:
    """Converte datas da LYNCEUS (textos como 'Mar 11, 2026 3:43pm' ou 'Set 10, 2026')."""
    s = pd.Series(s).astype("string")
    out = pd.to_datetime(s, errors="coerce", format="%b %d, %Y %I:%M%p")
    faltantes = out.isna() & s.notna()
    if faltantes.any():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            out.loc[faltantes] = pd.to_datetime(s[faltantes], errors="coerce")
    return out


# ----------------------------------------------------------------------------
# Carga de dados (com cache)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Carregando base de dados...")
def load_data(filepath: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Le as 3 abas do arquivo Excel e aplica limpeza padrao.

    Retorna (df_bc, df_lyn, df_gc) com ids normalizados e datas convertidas.
    """
    df_bc = pd.read_excel(
        filepath,
        sheet_name="BR CONCESSIONS",
        dtype={
            "lt_dispatch_transporter_id": str,
            "tracking_id": str,
            "asin": str,
            "delivery_station_code": str,
            "delivery_station_country_code": str,
        },
    )
    df_lyn = pd.read_excel(filepath, sheet_name="LYNCEUS", dtype={"Transporter Id": str, "Order ID/TID": str})
    df_gc = pd.read_excel(filepath, sheet_name="GCMS", dtype={"Driver ID": str, "DSP Short Name": str})

    # -- ID normalizado (sem espacos, maiusculo)
    def clean_id(s: pd.Series) -> pd.Series:
        out = s.astype("string").str.strip().str.upper()
        out = out.replace({"<NA>": None, "NAN": None, "NONE": None})
        return out

    df_bc["id_transporter"] = clean_id(df_bc["lt_dispatch_transporter_id"])
    df_bc["id_tracking"] = clean_id(df_bc["tracking_id"])
    df_bc["id_asin"] = clean_id(df_bc["asin"])

    df_lyn["id_transporter"] = clean_id(df_lyn["Transporter Id"])
    df_lyn["id_tracking"] = clean_id(df_lyn["Order ID/TID"])

    df_gc["id_transporter"] = clean_id(df_gc["Driver ID"])

    # -- Datas LYNCEUS (texto livre -> datetime)
    for c in LYN_DATE_COLS:
        if c in df_lyn.columns:
            df_lyn[c] = parse_mixed_dates(df_lyn[c])

    df_lyn["Creation Date"] = df_lyn["Creation Date UTC"]

    # -- Colunas numericas
    df_bc["Concession_Cost"] = pd.to_numeric(df_bc["Concession_Cost"], errors="coerce")

    return df_bc, df_lyn, df_gc


def get_source_file() -> Path:
    if "source_file" in st.session_state and st.session_state["source_file"] is not None:
        return Path(st.session_state["source_file"])
    return DEFAULT_FILE


def current_source_label(source: Path) -> str:
    return source.name


# ----------------------------------------------------------------------------
# Helpers de visualizacao
# ----------------------------------------------------------------------------
def fmt_money(v: float) -> str:
    if pd.isna(v):
        return "-"
    return f"$ {v:,.2f}"


def fmt_int(v: float) -> str:
    if pd.isna(v):
        return "-"
    return f"{v:,.0f}"


def title_bar(title: str, height: int = 430) -> None:
    st.subheader(title)
    st.caption("Fonte: arquivo Excel das abas GCMS / LYNCEUS / BR CONCESSIONS.")


def kpi_cols(items: list[tuple[str, str, str]]) -> None:
    cols = st.columns(len(items))
    for col, (rotulo, valor, dica) in zip(cols, items):
        col.metric(label=rotulo, value=valor, help=dica)


def sort_widget(df: pd.DataFrame, key: str, default_col: str | None = None,
                default_desc: bool = True, label: str = "Ordenar por") -> pd.DataFrame:
    """Mostra controles de classificacao (coluna + direcao) e retorna o df ordenado."""
    cols = list(df.columns)
    idx = cols.index(default_col) if default_col in cols else 0
    c1, c2, _ = st.columns([3, 2, 4])
    sort_col = c1.selectbox(label, cols, index=idx, key=f"sortcol_{key}")
    direcao = c2.selectbox("Direcao", ["Descrescente", "Crescente"],
                           index=0 if default_desc else 1, key=f"sortdir_{key}")
    return df.sort_values(sort_col, ascending=(direcao == "Crescente"), na_position="last")


def sortable_table(df: pd.DataFrame, key: str, default_col: str | None = None,
                   default_desc: bool = True, limit: int | None = None,
                   label: str = "Ordenar tabela por") -> pd.DataFrame:
    """Exibe uma tabela com controle de classificacao embutido."""
    out = sort_widget(df, key, default_col, default_desc, label)
    if limit is not None:
        out = out.head(limit)
    st.dataframe(out, width="stretch", hide_index=True)
    return out


# ----------------------------------------------------------------------------
# Pagina 1 - Visao Geral
# ----------------------------------------------------------------------------
def page_overview(df_bc, df_lyn, df_gc):
    st.title("Visao Geral de Investigacao - Brasil")

    total_cost = df_bc["Concession_Cost"].sum()
    total_rows = len(df_bc)
    n_transp_bc = df_bc["id_transporter"].nunique()

    n_inv = len(df_lyn)
    loss = df_lyn["Loss Value $"].sum()
    recovered = df_lyn["Recovered Value $"].sum()
    avoidance = df_lyn["Avoidance Value $"].sum()
    recovery_rate = (recovered / loss * 100) if loss else 0.0

    n_drivers = df_gc["id_transporter"].nunique()
    avg_dpm = df_gc["Loss Rate DPM"].mean()

    kpi_cols([
        ("Registros de concessao (BR)", fmt_int(total_rows), "Total de linhas na aba BR CONCESSIONS"),
        ("Custo total de concessao", fmt_money(total_cost), "Soma de Concession_Cost"),
        ("Transportadores unicos (BR)", fmt_int(n_transp_bc), "IDs distintos em lt_dispatch_transporter_id"),
        ("Investigacoes (LYNCEUS)", fmt_int(n_inv), "Total de casos na aba LYNCEUS"),
    ])
    kpi_cols([
        ("Perda (Loss Value)", fmt_money(loss), "Soma da perda confirmada nas investigacoes"),
        ("Recuperado (Recovered)", fmt_money(recovered), "Valor recuperado nas investigacoes"),
        ("Prevenido (Avoidance)", fmt_money(avoidance), "Evitacao de perda apontada nas investigacoes"),
        ("Taxa de recuperacao", f"{recovery_rate:.1f}%", "Recuperado / Perda"),
    ])
    kpi_cols([
        ("Motoristas rastreados (GCMS)", fmt_int(n_drivers), "Drivers unicos na aba GCMS"),
        ("Loss Rate medio (GCMS)", f"{avg_dpm:,.0f} DPM", "Media de perdas por milhao"),
        ("Casos abertos (LYNCEUS)", fmt_int((df_lyn['Status'] == 'Open').sum()), "Investigacoes em andamento"),
        ("Motoristas suspeitos/PRD (GCMS)", fmt_int(((df_gc['Driver Risk'].isin(['PRD', 'PAD']))).sum()), "PRD + PAD na GCMS"),
    ])

    st.divider()

# --- Cobertura / relacionamento entre as abas
    st.subheader("Relacionamento entre as abas (ID comum)")
    st.caption("Use os controles 'Ordenar por' para classificar os dados em qualquer grafico ou tabela do dashboard.")

    bc_ids = set(df_bc["id_transporter"].dropna())
    lyn_ids = set(df_lyn["id_transporter"].dropna())
    gc_ids = set(df_gc["id_transporter"].dropna())

    nivel = pd.DataFrame(
        [
            {
                "Grupo": "Transportadores com concedoes (BR)",
                "Qtd": len(bc_ids),
            },
            {
                "Grupo": "Transportadores investigados (LYNCEUS)",
                "Qtd": len(lyn_ids),
            },
            {
                "Grupo": "Motoristas com perfil (GCMS)",
                "Qtd": len(gc_ids),
            },
            {
                "Grupo": "Investigados E com concessao",
                "Qtd": len(lyn_ids & bc_ids),
            },
            {
                "Grupo": "Investigados SEM concessao registrada",
                "Qtd": len(lyn_ids - bc_ids),
            },
            {
                "Grupo": "Investigados COM perfil GCMS",
                "Qtd": len(lyn_ids & gc_ids),
            },
            {
                "Grupo": "Investigados SEM registro na GCMS",
                "Qtd": len(lyn_ids - gc_ids),
            },
        ]
)
    nivel = sort_widget(nivel, "ov_rel", default_col="Qtd", default_desc=True, label="Ordenar relacionamento")
    fig = px.bar(
        nivel,
        x="Qtd",
        y="Grupo",
        orientation="h",
        color="Qtd",
        color_continuous_scale="Tealgrn",
        text="Qtd",
        template=TEMPLATE,
        height=400,
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(showlegend=False, coloraxis_showscale=False, margin=dict(l=10, r=40, t=10, b=10))
    st.plotly_chart(fig, width="stretch")

    st.caption(
        "Leitura: todo registro de concessao pode ser vinculado ao motorista (lt_dispatch_transporter_id), "
        "o qual possui perfil de risco na GCMS (Driver ID) e pode ter investigacao na LYNCEUS (Transporter Id). "
        "O mesmo vinculo existe em nivel de pacote: tracking_id (BR) = Order ID/TID (LYNCEUS)."
    )
    st.divider()

    left, right = st.columns(2)
    with left:
        st.subheader("Correlacao entre abas - pacote")
        br_tid = set(df_bc["id_tracking"].dropna())
        lyn_tid = set(df_lyn["id_tracking"].dropna())
        n_ordem = df_lyn["id_tracking"].nunique()
        st.markdown(
            f"- Order IDs na LYNCEUS: **{n_ordem:,.0f}**"
            f"\n- Pacotes BR CONCESSIONS que batem com um pedido investigado: "
            f"**{df_bc['id_tracking'].isin(lyn_tid).sum():,.0f}** de **{len(df_bc):,.0f}** registros"
            f"\n- Custo dessas concessoes vinculadas: "
            f"**{fmt_money(df_bc.loc[df_bc['id_tracking'].isin(lyn_tid), 'Concession_Cost'].sum())}**"
        )
    with right:
        st.subheader("Cobertura de custo por motorista conhecido")
        known = df_bc["id_transporter"].isin(gc_ids)
        cost_known = df_bc.loc[known, "Concession_Cost"].sum()
        st.markdown(
            f"- {known.sum():,.0f} de {len(df_bc):,.0f} registros pertencem a motoristas com perfil na GCMS "
            f"({known.mean()*100:.1f}%)."
            f"\n- Custo coberto por motorista conhecido: {fmt_money(cost_known)} de {fmt_money(total_cost)} "
            f"({cost_known/total_cost*100:.1f}%)."
        )

    st.divider()

    # --- Insights automaticos
    st.subheader("Pontos de atencao (calculados dos dados)")
    insights = []
    if recovery_rate < 25:
        insights.append(f"A taxa de recuperacao e apenas {recovery_rate:.1f}% - gerar acoes para recuperar perdas.")
    prin_mo = df_lyn["MO"].value_counts(dropna=False).index[0] if len(df_lyn) else None
    if prin_mo:
        insights.append(f"Modus operandi dominante nas investigacoes: **{prin_mo}**.")
    top_anomalia = "not_delivered"
    vc = df_bc["concession_event_name"].value_counts(dropna=False)
    if len(vc):
        insights.append(
            f"Evento de concessao mais frequente: **{vc.index[0]}** ({vc.iloc[0]:,.0f} ocorrencias, "
            f"{((vc.iloc[0]/total_rows)*100):.1f}% do total)."
        )
    state = df_lyn["State Province"].value_counts().index[0] if df_lyn["State Province"].nunique() else None
    if state:
        insights.append(f"Estado com mais investigacoes: **{state}**.")
    pad_prd = df_gc[df_gc["Driver Risk"].isin(["PRD", "PAD"])]
    if len(pad_prd):
        payload = df_lyn[df_lyn["id_transporter"].isin(set(pad_prd["id_transporter"]))]
        insights.append(
            f"{len(pad_prd)} motoristas marcados como PRD/PAD na GCMS. Desses, "
            f"{payload['id_transporter'].nunique()} possuem investigacao registrada (perda de {fmt_money(payload['Loss Value $'].sum())})."
        )
    for i in insights:
        st.markdown(f"- {i}")
    st.caption("Instrucoes: os insights sao recalculados automaticamente a cada nova base enviada.")


# ----------------------------------------------------------------------------
# Pagina 2 - BR CONCESSIONS
# ----------------------------------------------------------------------------
def page_bc(df_bc):
    st.title("BR CONCESSIONS - Analise de Concessoes")
    st.caption("Registro operacional de concessoes (perdas) por transportador e pacote.")

    total_cost = df_bc["Concession_Cost"].sum()
    rows = len(df_bc)
    kpi_cols([
        ("Registros", fmt_int(rows), "Total de linhas"),
        ("Custo total", fmt_money(total_cost), "Soma de Concession_Cost"),
        ("Transportadores", fmt_int(df_bc["id_transporter"].nunique()), "IDs unicos"),
        ("Produtos distintos (ASIN)", fmt_int(df_bc["id_asin"].nunique()), "ASINs unicos"),
    ])
    st.divider()

    left, right = st.columns(2)
    with left:
        vc = df_bc["concession_event_name"].value_counts(dropna=False).reset_index()
        vc.columns = ["Evento", "Qtd"]
        vc = sort_widget(vc, "bc_evt", default_col="Qtd", default_desc=True, label="Ordenar eventos")
        fig = px.bar(vc, x="Evento", y="Qtd", color="Qtd", color_continuous_scale="Tealrose",
                     title="Eventos de concessao", template=TEMPLATE, height=420)
        fig.update_layout(showlegend=False, coloraxis_showscale=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, width="stretch")
    with right:
        top_t_all = df_bc.groupby("id_transporter")["Concession_Cost"].sum().reset_index()
        top_t_all.columns = ["Transportador", "Custo"]
        top_t_all = sort_widget(top_t_all, "bc_top", default_col="Custo", default_desc=True, label="Ordenar transportadores")
        top_t = top_t_all.head(15)[::-1]
        fig = px.bar(top_t, x="Custo", y="Transportador", orientation="h",
                     color="Custo", color_continuous_scale="Tealrose",
                     title="Top 15 transportadores por custo", template=TEMPLATE, height=420)
        fig.update_layout(showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")

    left, right = st.columns(2)
    with left:
        top_a_all = df_bc.groupby(["id_asin", "asin_name"]).agg(
            Custo=("Concession_Cost", "sum"), Qtd=("Concession_Cost", "count")
        ).reset_index()
        top_a_all = sort_widget(top_a_all, "bc_asin", default_col="Custo", default_desc=True, label="Ordenar ASINs")
        top_a = top_a_all.head(12)
        fig = px.bar(top_a, x="Custo", y="Qtd", color="Qtd", color_continuous_scale="Blues",
                     title="Top ASINs por custo de concessao (hover p/ nome)", hover_data={"asin_name": True},
                     labels={"Qtd": "Ocorrencias"}, template=TEMPLATE, height=430)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")
    with right:
        fig = px.histogram(df_bc, x="Concession_Cost", nbins=50,
                           title="Distribuicao do custo por concessao",
                           labels={"Concession_Cost": "Custo de concessao"},
                           template=TEMPLATE, height=430)
        fig.update_layout(yaxis_title="Frequencia")
        st.plotly_chart(fig, width="stretch")

    st.divider()
    st.subheader("Distribuicao geografica (entrega)")
    geo = df_bc.groupby("delivery_station_state").agg(
        Custo=("Concession_Cost", "sum"), Registros=("Concession_Cost", "count")
    ).reset_index()
    geo = sort_widget(geo, "bc_geo", default_col="Custo", default_desc=True, label="Ordenar estados")
    fig = px.bar(geo, x="delivery_station_state", y="Custo", color="Registros",
                 color_continuous_scale="Tealgrn",
                 title="Custo de concessao por estado", template=TEMPLATE, height=420)
    fig.update_layout(coloraxis_showscale=True, xaxis_tickangle=-45)
    st.plotly_chart(fig, width="stretch")

    st.divider()
    with st.expander("Ver tabela - maiores transportadores por custo"):
        tb = df_bc.groupby(["id_transporter"]).agg(
            Registros=("Concession_Cost", "count"),
            Custo=("Concession_Cost", "sum"),
            Custo_medio=("Concession_Cost", "mean"),
            Produtos=("id_asin", "nunique"),
        ).reset_index()
        sortable_table(tb, "bc_tbtab", default_col="Custo", default_desc=True, label="Ordenar tabela")


# ----------------------------------------------------------------------------
# Pagina 3 - LYNCEUS
# ----------------------------------------------------------------------------
def page_lyn(df_lyn):
    st.title("LYNCEUS - Analise de Investigacoes")
    st.caption("Casos de investigacao de perda (aberta, encerrada ou pendente).")

    loss = df_lyn["Loss Value $"].sum()
    recovered = df_lyn["Recovered Value $"].sum()
    avoidance = df_lyn["Avoidance Value $"].sum()
    kpi_cols([
        ("Investigacoes", fmt_int(len(df_lyn)), "Total de casos"),
        ("Loss Value total", fmt_money(loss), "Perda confirmada"),
        ("Recovered total", fmt_money(recovered), "Valor recuperado"),
        ("Avoidance total", fmt_money(avoidance), "Perda evitada"),
    ])
    kpi_cols([
        ("Statuses abertos", fmt_int((df_lyn["Status"] == "Open").sum()), "Casos abertos"),
        ("SEV2/SEV3", fmt_int(df_lyn["SEV"].isin(["SEV2", "SEV3"]).sum()), "Casos de severidade alta"),
        ("Closures", fmt_int(df_lyn["Closures"].sum()), "Total de fechamentos (dematge/deboard)"),
        ("Separations", fmt_int(df_lyn["Separations"].sum()), "Total de separacoes de associado"),
    ])
    st.divider()

    left, right = st.columns(2)
    with left:
        v1 = df_lyn["Status"].value_counts(dropna=False).reset_index()
        v1.columns = ["Status", "Qtd"]
        v1 = sort_widget(v1, "lyn_status", default_col="Qtd", default_desc=True, label="Ordenar status")
        fig = px.bar(v1, x="Status", y="Qtd", color="Status",
                     color_discrete_sequence=PALETTE, title="Status dos casos", template=TEMPLATE, height=380)
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, width="stretch")
    with right:
        v2 = df_lyn["SEV"].value_counts(dropna=False).reset_index()
        v2.columns = ["SEV", "Qtd"]
        v2 = sort_widget(v2, "lyn_sev", default_col="Qtd", default_desc=True, label="Ordenar SEV")
        fig = px.bar(v2, x="SEV", y="Qtd", color="SEV",
                     color_discrete_sequence=PALETTE, title="Distribuicao de severidade (SEV)", template=TEMPLATE, height=380)
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, width="stretch")

    left, right = st.columns(2)
    with left:
        v3_all = df_lyn.groupby("MO")["Loss Value $"].sum().reset_index()
        v3_all.columns = ["MO", "Perda"]
        v3_all = sort_widget(v3_all, "lyn_mo", default_col="Perda", default_desc=True, label="Ordenar MO")
        v3 = v3_all.head(12)
        fig = px.bar(v3, x="Perda", y="MO", orientation="h", color="Perda",
                     color_continuous_scale="Tealrose",
                     title="Perda por modus operandi (MO)", template=TEMPLATE, height=430)
        fig.update_layout(showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")
    with right:
        v4 = df_lyn["Loss Bucket"].value_counts(dropna=False).reset_index()
        v4.columns = ["Bucket", "Qtd"]
        v4 = sort_widget(v4, "lyn_bucket", default_col="Qtd", default_desc=True, label="Ordenar buckets")
        fig = px.bar(v4, x="Bucket", y="Qtd", color="Qtd", color_continuous_scale="Tealgrn",
                     title="Casos por Loss Bucket", template=TEMPLATE, height=430)
        fig.update_layout(coloraxis_showscale=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, width="stretch")

    st.divider()
    left, right = st.columns(2)
    with left:
        df_lyn["Mes"] = df_lyn["Creation Date"].dt.to_period("M").astype(str)
        tserie = df_lyn.groupby("Mes").agg(
            Casos=("Inv. Id", "count"),
            Perda=("Loss Value $", "sum"),
        ).reset_index()
        fig = px.bar(tserie, x="Mes", y="Casos", color="Perda", color_continuous_scale="Tealgrn",
                     title="Casos por mes de criacao", template=TEMPLATE, height=400)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")
    with right:
        v5 = df_lyn["State Province"].value_counts(dropna=False).reset_index()
        v5.columns = ["Estado", "Qtd"]
        v5 = sort_widget(v5, "lyn_estado", default_col="Qtd", default_desc=True, label="Ordenar estados")
        fig = px.bar(v5, x="Estado", y="Qtd", color="Qtd", color_continuous_scale="Blues",
                     title="Casos por estado", template=TEMPLATE, height=400)
        fig.update_layout(coloraxis_showscale=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, width="stretch")

    st.divider()
    st.subheader("Maiores perdas por investigacao")
    top_all = df_lyn[
        ["Inv. Id", "Transporter Id", "Status", "SEV", "MO", "Loss Bucket",
         "Loss Value $", "Recovered Value $", "Avoidance Value $", "State Province", "City"]
    ]
    sortable_table(top_all, "lyn_top20", default_col="Loss Value $", default_desc=True,
                   limit=20, label="Ordenar perdas")

    with st.expander("Ver todas as investigacoes (resumo)"):
        cols = [c for c in df_lyn.columns if c not in ("Open Case", "Additional Investigator",
                                                       "Supporting 3rd Party Tool", "Unsolvable Reason Other Details",
                                                       "Non Security Reason Other Details", "Reason for PII Info Access")]
        sortable_table(df_lyn[cols], "lyn_totab", default_col="Loss Value $", default_desc=True,
                       label="Ordenar investigacoes")


# ----------------------------------------------------------------------------
# Pagina 4 - GCMS
# ----------------------------------------------------------------------------
def page_gcms(df_gc):
    st.title("GCMS - Perfil de Motoristas / Transportadores")
    st.caption("Base de perfil, risco e desempenho dos motoristas (DSP).")

    kpi_cols([
        ("Motoristas", fmt_int(df_gc["id_transporter"].nunique()), "Drivers unicos"),
        ("Loss Rate medio", f"{df_gc['Loss Rate DPM'].mean():,.0f} DPM", "Perdas por milhao"),
        ("Concession Cost medio", fmt_money(df_gc["Concession Cost"].mean()), "Media por motorista"),
        ("Ship Units total", fmt_int(df_gc["Total Ship Units"].sum()), "Total de entregas"),
    ])
    risk = df_gc["Driver Risk"].value_counts(dropna=False).reset_index()
    risk.columns = ["Risco", "Qtd"]
    r_known = df_gc["Driver Risk"].notna().sum()
    risk_prec = [
        ("Trusted Transporter", ((df_gc["Driver Risk"] == "Trusted Transporter").sum())),
        ("PRD (Problema recorrente)", ((df_gc["Driver Risk"] == "PRD").sum())),
        ("PAD (Acao de perda)", ((df_gc["Driver Risk"] == "PAD").sum())),
        ("Sem classificacao", int(r_known) - int((df_gc["Driver Risk"] == "Trusted Transporter").sum()) - int((df_gc["Driver Risk"] == "PRD").sum()) - int((df_gc["Driver Risk"] == "PAD").sum()) if r_known else 0),
    ]
    kpi_cols([
        ("Trusted Transporter", fmt_int(risk_prec[0][1]), "Motoristas confiaveis"),
        ("PRD", fmt_int(risk_prec[1][1]), "Perda recorrente"),
        ("PAD", fmt_int(risk_prec[2][1]), "Perda com acao"),
        ("Sem classificacao", fmt_int(int(df_gc['Driver Risk'].isna().sum())), "Nao classificados"),
    ])
    st.divider()

    left, right = st.columns(2)
    with left:
        emp = df_gc["Employee Status [:Offboard Reason]"].astype(str).value_counts(dropna=False).reset_index()
        emp.columns = ["Status", "Qtd"]
        emp = sort_widget(emp, "gc_emp", default_col="Qtd", default_desc=True, label="Ordenar status")
        fig = px.bar(emp, x="Qtd", y="Status", orientation="h", color="Qtd",
                     color_continuous_scale="Tealrose", title="Situacao do empregado (offboard)", template=TEMPLATE, height=430)
        fig.update_layout(showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")
    with right:
        dsp = df_gc["DSP Short Name"].value_counts(dropna=False).reset_index()
        dsp.columns = ["DSP", "Qtd"]
        dsp = sort_widget(dsp, "gc_dsp", default_col="Qtd", default_desc=True, label="Ordenar DSPs")
        fig = px.bar(dsp, x="Qtd", y="DSP", orientation="h", color="Qtd",
                     color_continuous_scale="Blues", title="Top DSPs por quantidade de motoristas", template=TEMPLATE, height=430)
        fig.update_layout(showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, width="stretch")

    st.divider()
    st.subheader("Motoristas de maior risco / perda")
    gc_plot = df_gc.copy()
    fig = px.scatter(
        gc_plot, x="Total Ship Value", y="Concession Cost", color="Driver Risk",
        size="Total Ship Units", hover_name="id_transporter", hover_data=["DSP Short Name", "Loss Rate DPM"],
        color_discrete_map={"Trusted Transporter": "#2ca02c", "PRD": "#d62728", "PAD": "#ff7f0e"},
        title="Valor entregue x custo de concessao (bolha = unidades; cor = risco)",
        template=TEMPLATE, height=520,
    )
    fig.update_layout(legend_title="Driver Risk")
    st.plotly_chart(fig, width="stretch")

    st.divider()
    st.subheader("Top motoristas por custo de concessao")
    sortable_table(
        df_gc[["id_transporter", "DSP Short Name", "Driver Risk",
               "Loss Rate DPM", "Concession Cost", "Total Ship Units", "Total Ship Value"]],
        "gc_top", default_col="Concession Cost", default_desc=True, limit=25,
        label="Ordenar motoristas",
    )


# ----------------------------------------------------------------------------
# Pagina 5 - Cruzamento (BI otimizado)
# ----------------------------------------------------------------------------
def page_cross(df_bc, df_lyn, df_gc):
    st.title("Cruzamento de dados - GCMS x LYNCEUS x BR CONCESSIONS")
    st.caption(
        "O ID do motorista e a chave comum: Barra de concessao (`lt_dispatch_transporter_id`), investigacao "
        "(`Transporter Id`) e perfil de risco (`Driver ID`). Completa-se no nivel de pacote: `tracking_id` = `Order ID/TID`."
    )

    bc_ids = set(df_bc["id_transporter"].dropna())
    lyn_ids = set(df_lyn["id_transporter"].dropna())
    gc_ids = set(df_gc["id_transporter"].dropna())

    # ---- Conjuntos de relacionamento
    sets_df = pd.DataFrame(
        [
            {"Grupo": "SOMENTE BR", "Qtd": len(bc_ids - lyn_ids - gc_ids), "Cor": "BR CONCESSIONS"},
            {"Grupo": "BR + GCMS", "Qtd": len((bc_ids & gc_ids) - lyn_ids), "Cor": "BR + GCMS"},
            {"Grupo": "BR + LYNCEUS", "Qtd": len((bc_ids & lyn_ids) - gc_ids), "Cor": "BR + LYNCEUS"},
            {"Grupo": "LYNCEUS + GCMS", "Qtd": len((lyn_ids & gc_ids) - bc_ids), "Cor": "LYNCEUS + GCMS"},
            {"Grupo": "TODAS (BR+LYN+GCMS)", "Qtd": len(bc_ids & lyn_ids & gc_ids), "Cor": "TODAS"},
            {"Grupo": "SOMENTE LYNCEUS", "Qtd": len(lyn_ids - bc_ids - gc_ids), "Cor": "LYNCEUS"},
            {"Grupo": "SOMENTE GCMS", "Qtd": len(gc_ids - bc_ids - lyn_ids), "Cor": "GCMS"},
        ]
    )
    sets_df = sort_widget(sets_df, "cr_sets", default_col="Qtd", default_desc=True, label="Ordenar classificacao")
    fig = px.bar(sets_df, x="Qtd", y="Grupo", orientation="h", color="Cor",
                 color_discrete_sequence=["#4c72b0", "#9ecae1", "#dd8452", "#c5b0d5", "#55a868", "#ff9896", "#98df8a"],
                 text="Qtd", template=TEMPLATE, height=460, title="Classificacao dos transportadores por presenca em cada base")
    fig.update_traces(textposition="outside")
    fig.update_layout(showlegend=False, margin=dict(l=10, r=40, t=50, b=10))
    st.plotly_chart(fig, width="stretch")

    st.divider()

    # ---- Tabela mestra
    st.subheader("Tabela mestra por transportador")
    br_agg = df_bc.groupby("id_transporter").agg(
        n_concessoes=("Concession_Cost", "count"),
        custo_concessao=("Concession_Cost", "sum"),
        produtos=("id_asin", "nunique"),
    ).reset_index().rename(columns={"id_transporter": "id"})

    lyn_agg = df_lyn.groupby("id_transporter").agg(
        n_investigacoes=("Inv. Id", "count"),
        loss=("Loss Value $", "sum"),
        recovered=("Recovered Value $", "sum"),
        avoidance=("Avoidance Value $", "sum"),
        sev_max=("SEV", "max"),
    ).reset_index().rename(columns={"id_transporter": "id"})

    gc_agg = df_gc.groupby("id_transporter").agg(
        driver_risk=("Driver Risk", "first"),
        loss_rate_dpm=("Loss Rate DPM", "first"),
        dsp=("DSP Short Name", "first"),
        status_emp=("Employee Status [:Offboard Reason]", "first"),
    ).reset_index().rename(columns={"id_transporter": "id"})

    master = (br_agg.merge(lyn_agg, on="id", how="outer")
                    .merge(gc_agg, on="id", how="outer"))
    master["n_investigacoes"] = master["n_investigacoes"].fillna(0).astype(int)
    master["n_concessoes"] = master["n_concessoes"].fillna(0).astype(int)

# filtros
    with st.container():
        f1, f2, f3 = st.columns(3)
        investigados = f1.selectbox("Somente investigados", ["Todos", "Sim, com investigacao", "Nao investigados"], key="f_inv")
        tem_concessao = f2.selectbox("Com concessao", ["Todos", "Com concessao", "Sem concessao"], key="f_conc")
        risco = f3.selectbox("Driver Risk (GCMS)", ["Todos"] + sorted(df_gc["Driver Risk"].dropna().unique().tolist()), key="f_risk")

    if investigados == "Sim, com investigacao":
        master = master[master["n_investigacoes"] > 0]
    elif investigados == "Nao investigados":
        master = master[master["n_investigacoes"] == 0]
    if tem_concessao == "Com concessao":
        master = master[master["n_concessoes"] > 0]
    elif tem_concessao == "Sem concessao":
        master = master[master["n_concessoes"] == 0]
    if risco != "Todos":
        master = master[master["driver_risk"] == risco]

    master = sort_widget(master, "cr_master", default_col="custo_concessao", default_desc=True,
                         label="Ordenar tabela mestra")

    st.dataframe(
        master[["id", "dsp", "driver_risk", "status_emp", "loss_rate_dpm",
                "n_concessoes", "custo_concessao", "produtos",
                "n_investigacoes", "loss", "recovered", "avoidance", "sev_max"]],
        width="stretch", hide_index=True,
    )

    st.divider()

    # ---- Cross de nivel de pacote
    left, right = st.columns(2)
    with left:
        st.subheader("Vincular pacote investigado x concessao")
        lyn_tid = set(df_lyn["id_tracking"].dropna())
        matched = df_bc["id_tracking"].isin(lyn_tid)
        st.markdown(
            f"- Pacotes investigados na LYNCEUS: **{df_lyn['id_tracking'].nunique():,.0f}**"
            f"\n- **{matched.sum():,.0f}** registros de concessao (BR) referem-se a um pacote investigado."
            f"\n- Custo dessas ocorrencias: **{fmt_money(df_bc.loc[matched, 'Concession_Cost'].sum())}**."
        )
        st.caption("Ex.: concession `dÃ©livrÃ© nÃ£t reÃ§u` com `Order ID/TID` investigado revela o quanto da "
                   "indaixa de 'nÃ£o entregue' foi objeto de investigacao formal.")
    with right:
        st.subheader("Investigados sem perfil na GCMS")
        sem_gc = lyn_ids - gc_ids
        st.markdown(
            f"- **{len(sem_gc):,}** transportadores investigados nao possuem registro na GCMS "
            f"({len(sem_gc)/max(len(lyn_ids),1)*100:.1f}% das investigacoes).\n"
            f"- O exemplo mais comum: casos de **Delivered Not Received** sem alocacao de perfil de risco."
        )

    st.divider()
    st.subheader("Insights de cruzamento")
    hl = []
    custo_inv = df_bc[df_bc["id_transporter"].isin(lyn_ids)]
    hl.append(
        f"**{fmt_money(custo_inv['Concession_Cost'].sum())}** de {fmt_money(df_bc['Concession_Cost'].sum())} "
        f"({custo_inv['Concession_Cost'].sum()/df_bc['Concession_Cost'].sum()*100:.1f}%) em concessoes vem de transportadores "
        f"com investigacao formal na LYNCEUS."
    )
    custo_gc = df_bc[df_bc["id_transporter"].isin(gc_ids)]
    hl.append(
        f"**{custo_gc['Concession_Cost'].mean():,.2f}** e o custo medio por concessao para motoristas com perfil GCMS "
        f"(comparar com a media geral {df_bc['Concession_Cost'].mean():,.2f})."
    )
    risk_prd = set(df_gc.loc[df_gc["Driver Risk"].isin(["PRD", "PAD"]), "id_transporter"])
    perda_prd = df_lyn[df_lyn["id_transporter"].isin(risk_prd)]["Loss Value $"].sum()
    hl.append(
        f"Motoristas PRD/PAD respondem por **{fmt_money(perda_prd)}** em perdas investigadas "
        f"({perda_prd/df_lyn['Loss Value $'].sum()*100:.1f}% do valor total de loss)."
    )
    for h in hl:
        st.markdown(f"- {h}")


# ----------------------------------------------------------------------------
# Pagina 6 - Dados brutos / exportacoes
# ----------------------------------------------------------------------------
def page_raw(df_bc, df_lyn, df_gc, source_name):
    st.title("Dados brutos e exportacao")
    st.caption(f"Fonte carregada: **{source_name}**")

    want = st.selectbox("Aba", ["BR CONCESSIONS", "LYNCEUS", "GCMS"], key="raw_select")
    df = {"BR CONCESSIONS": df_bc, "LYNCEUS": df_lyn, "GCMS": df_gc}[want]
    st.caption("Classifique por qualquer coluna. Escolha o numero de linhas abaixo apos ordenar.")
    df = sort_widget(df, "raw_sort", default_col=df.columns[0], default_desc=False, label="Ordenar dados brutos")
    n = st.slider("Linhas a exibir", 10, min(len(df), 500), 50, key="raw_n")
    st.dataframe(df.head(n), width="stretch", hide_index=True)

    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        f"Baixar {want} (CSV)",
        data=csv,
        file_name=f"{want.replace(' ', '_')}.csv",
        mime="text/csv",
    )


# ----------------------------------------------------------------------------
# Sidebar + navegacao
# ----------------------------------------------------------------------------
def sidebar():
    st.sidebar.title("BI Investigacao - BR")
    st.sidebar.caption("Fontes: abas GCMS, LYNCEUS e BR CONCESSIONS de um mesmo Excel.")

    uploaded = st.sidebar.file_uploader(
        "Atualizar base de dados",
        type=["xlsx", "xls"],
        help="Envie um novo Excel com as 3 abas (GCMS, LYNCEUS, BR CONCESSIONS) para substituir a base atual.",
    )
    if uploaded is not None:
        target = DATA_DIR / uploaded.name
        target.write_bytes(uploaded.getbuffer())
        st.session_state["source_file"] = str(target)
        st.sidebar.success(f"Base atualizada: {uploaded.name}")
        st.sidebar.caption("Recarregue os dados abaixo para aplicar.")

    source = get_source_file()
    if st.sidebar.button("Recarregar dados atuais"):
        st.cache_data.clear()
        st.rerun()
    if source != DEFAULT_FILE:
        if st.sidebar.button("Restaurar base padrao (geral_01.xlsx)"):
            st.session_state["source_file"] = str(DEFAULT_FILE)
            st.cache_data.clear()
            st.rerun()

    st.sidebar.divider()
    st.sidebar.markdown(f"Arquivo ativo: **{current_source_label(source)}**")
    st.sidebar.caption(
        "Para atualizar, envie um novo arquivo acima. O arquivo e salvo na pasta `data/` "
        "do projeto e passa a ser a base de origem ate que a base padrao seja restaurada."
    )

    nav = st.sidebar.radio(
        "Navegacao",
        ["Visao Geral", "BR CONCESSIONS", "LYNCEUS", "GCMS", "Cruzamento de dados", "Dados brutos"],
    )
    return nav, source


# ----------------------------------------------------------------------------
def main():
    nav, source = sidebar()

    if not source.exists():
        st.error("Base de dados nao encontrada. Coloque o Excel em `data/` ou envie um arquivo no menu lateral.")
        st.stop()

    df_bc, df_lyn, df_gc = load_data(str(source))

    if nav == "Visao Geral":
        page_overview(df_bc, df_lyn, df_gc)
    elif nav == "BR CONCESSIONS":
        page_bc(df_bc)
    elif nav == "LYNCEUS":
        page_lyn(df_lyn)
    elif nav == "GCMS":
        page_gcms(df_gc)
    elif nav == "Cruzamento de dados":
        page_cross(df_bc, df_lyn, df_gc)
    elif nav == "Dados brutos":
        page_raw(df_bc, df_lyn, df_gc, current_source_label(source))

    st.divider()
    st.caption(f"Dashboard gerado em {dt.datetime.now():%d/%m/%Y %H:%M} - base: {current_source_label(source)}")


if __name__ == "__main__":
    main()
