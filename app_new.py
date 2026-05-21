import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

APP_TITLE = "Publications"
DEFAULT_XLSX = "publicacoes.xlsx"
FALLBACK_XLSX = "889a07e4-87ea-4ae9-822a-25b5baf10492.xlsx"
SHEET_NAME = "vlistapubs"

# ── Google Sheets via opensheet.elk.sh ───────────────────────────────────────
# URL de exportação direta do Google Sheets (não depende de serviços terceiros)
# A sheet tem de estar partilhada como "Qualquer pessoa com o link pode ver"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1CXoOMiWEg_UnpoE1xNnt1axDzka0snsUj-A-1Q7IjvE/gviz/tq?tqx=out:csv&sheet=vlistapubs"

st.set_page_config(page_title=APP_TITLE, layout="wide")

# ─────────────────────────── helpers ────────────────────────────────────────

def find_excel_file() -> Path | None:
    candidates = [Path(__file__).with_name(DEFAULT_XLSX), Path(__file__).with_name(FALLBACK_XLSX)]
    for c in candidates:
        if c.exists():
            return c
    xlsx_files = sorted(Path(__file__).parent.glob("*.xlsx"))
    return xlsx_files[0] if xlsx_files else None


def normalize_type(value: str) -> str:
    txt = str(value).strip()
    mapping = {
        "Jornal": "Article",
        "Journal": "Article",
        "Proceedings": "Conference Paper",
        "Conference Paper": "Conference Paper",
        "Book": "Book",
        "Book Chapter": "Book Chapter",
    }
    return mapping.get(txt, txt if txt else "Other")


def parse_numeric(x):
    if pd.isna(x):
        return None
    txt = str(x).strip()
    if not txt:
        return None
    txt = (
        txt.replace(".", "").replace(",", ".")
        if txt.count(",") == 1 and txt.count(".") <= 1
        else txt.replace(",", ".")
    )
    try:
        return float(txt)
    except Exception:
        return None


def extract_quartiles(text: str) -> list[str]:
    if text is None:
        return []
    found = re.findall(r"Q\s*([1-4])", str(text).upper())
    return [f"Q{x}" for x in found]


def best_quartile_list(qs: list[str]) -> str:
    order = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    qs = [q for q in qs if q in order]
    if not qs:
        return ""
    return min(qs, key=lambda q: order[q])


def has_value(x) -> bool:
    if pd.isna(x):
        return False
    return str(x).strip() != ""


def compute_h_index(citations) -> int:
    citations = sorted(
        [int(x) for x in citations if x is not None and not pd.isna(x)],
        reverse=True,
    )
    h = 0
    for i, c in enumerate(citations, start=1):
        if c >= i:
            h = i
        else:
            break
    return h


def get_col(df: pd.DataFrame, col: str, default="") -> pd.Series:
    """Versão segura de df.get() para colunas — evita comportamento inesperado."""
    return df[col] if col in df.columns else pd.Series(default, index=df.index)


# ─────────────────────────── data loading ───────────────────────────────────

@st.cache_data(ttl=300)  # cache de 5 minutos
def load_data(uploaded_bytes: bytes | None = None) -> pd.DataFrame:
    if uploaded_bytes is not None:
        # Upload manual tem prioridade
        excel_path = io.BytesIO(uploaded_bytes)
        xl = pd.ExcelFile(excel_path)
        sheet = SHEET_NAME if SHEET_NAME in xl.sheet_names else xl.sheet_names[-1]
        df = pd.read_excel(excel_path, sheet_name=sheet)
    else:
        # Carregar da Google Sheet via opensheet.elk.sh
        try:
            df = pd.read_csv(GSHEET_URL)
        except Exception as e:
            raise RuntimeError(
                f"Não foi possível carregar os dados da Google Sheet: {e}"
            )

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].fillna("").astype(str).str.strip()

    df["ANO"] = pd.to_numeric(get_col(df, "ANO"), errors="coerce")

    df["AUTOR"] = get_col(df, "T1")
    df["TITULO_CURTO"] = get_col(df, "T2")
    df["FONTE_EXIBICAO"] = get_col(df, "T3")
    df["FONTE_PESQ"] = get_col(df, "SOURCE")
    df["SJR_CATEGORIA"] = get_col(df, "SJRCATEGORIA")
    df["JCR_CATEGORIA"] = get_col(df, "JCRCATEGORIA")
    df["JCR"] = get_col(df, "JCR")
    df["JCRANO"] = pd.to_numeric(get_col(df, "JCRANO"), errors="coerce").apply(
        lambda x: str(int(x)) if pd.notna(x) else ""
    )
    df["SJR"] = get_col(df, "SJR")
    df["SJRANO"] = pd.to_numeric(get_col(df, "SJRANO"), errors="coerce").apply(
        lambda x: str(int(x)) if pd.notna(x) else ""
    )
    df["TIPO"] = get_col(df, "TIPO").apply(normalize_type)
    df["URL"] = get_col(df, "URL")

    if "TITULO" in df.columns:
        mask_title = df["TITULO_CURTO"].eq("")
        df.loc[mask_title, "TITULO_CURTO"] = df.loc[mask_title, "TITULO"]

    df["CIT_ISI"] = get_col(df, "CITISI").apply(parse_numeric)
    df["CIT_SCOPUS"] = get_col(df, "CITSCOPUS").apply(parse_numeric)
    df["CIT_SCHOLAR"] = get_col(df, "CITSCHOLAR").apply(parse_numeric)

    df["SCOPUS_VAL"] = get_col(df, "SCOPUS")
    df["DBLP_VAL"] = get_col(df, "DBLP")
    df["ISI_VAL"] = get_col(df, "ISI")

    df["TEM_SCOPUS"] = df["SCOPUS_VAL"].apply(has_value)
    df["TEM_DBLP"] = df["DBLP_VAL"].apply(has_value)
    df["TEM_ISI"] = df["ISI_VAL"].apply(has_value)

    df["JCR_Q"] = df["JCR_CATEGORIA"].apply(lambda x: best_quartile_list(extract_quartiles(x)))
    df["SJR_Q"] = df["SJR_CATEGORIA"].apply(lambda x: best_quartile_list(extract_quartiles(x)))
    df["MELHOR_Q"] = df.apply(
        lambda row: best_quartile_list(
            extract_quartiles(row.get("JCR_CATEGORIA", ""))
            + extract_quartiles(row.get("SJR_CATEGORIA", ""))
        ),
        axis=1,
    )

    search_cols = ["TITULO_CURTO", "AUTOR", "FONTE_PESQ", "TIPO", "MELHOR_Q"]
    existing = [c for c in search_cols if c in df.columns]
    df["SEARCH_TEXT"] = df[existing].fillna("").astype(str).agg(" ".join, axis=1).str.lower()

    return df


# ─────────────────────────── URL state helpers ──────────────────────────────

def read_query_params() -> dict:
    """Lê os filtros guardados nos query params do URL."""
    params = st.query_params
    result = {}
    if "search" in params:
        result["search"] = params["search"]
    if "ano_min" in params and "ano_max" in params:
        try:
            result["ano_min"] = int(params["ano_min"])
            result["ano_max"] = int(params["ano_max"])
        except ValueError:
            pass
    if "tipos" in params:
        result["tipos"] = params.get_all("tipos")
    if "quartil" in params:
        result["quartil"] = params.get_all("quartil")
    if "bases" in params:
        result["bases"] = params.get_all("bases")
    for flag in ("only_url", "only_q1", "only_cit"):
        if flag in params:
            result[flag] = params[flag] == "1"
    return result


def write_query_params(filters: dict):
    """Escreve os filtros activos nos query params do URL."""
    new_params = {}
    if filters.get("search"):
        new_params["search"] = filters["search"]
    if "ano_min" in filters and "ano_max" in filters:
        new_params["ano_min"] = str(filters["ano_min"])
        new_params["ano_max"] = str(filters["ano_max"])
    if filters.get("tipos"):
        new_params["tipos"] = filters["tipos"]
    if filters.get("quartil"):
        new_params["quartil"] = filters["quartil"]
    if filters.get("bases"):
        new_params["bases"] = filters["bases"]
    for flag in ("only_url", "only_q1", "only_cit"):
        if filters.get(flag):
            new_params[flag] = "1"
    st.query_params.update(new_params)
    # Limpar params que já não estão activos
    for key in list(st.query_params.keys()):
        if key not in new_params:
            del st.query_params[key]


# ─────────────────────────── UI helpers ─────────────────────────────────────

def metric_card(title: str, value: str):
    st.markdown(
        f"""
        <div style="padding:0.9rem 1rem;border:1px solid #ddd;border-radius:14px;margin-bottom:0.4rem;">
            <div style="font-size:0.85rem;color:#666;">{title}</div>
            <div style="font-size:1.45rem;font-weight:700;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# Paleta de quartis partilhada — usada nos badges e nos gráficos
Q_COLORS = {
    "Q1": {"bg": "#e8f5e9", "fg": "#1b5e20", "border": "#c8e6c9", "hex": "#4caf50"},
    "Q2": {"bg": "#e3f2fd", "fg": "#0d47a1", "border": "#bbdefb", "hex": "#2196f3"},
    "Q3": {"bg": "#fff8e1", "fg": "#8d6e00", "border": "#ffecb3", "hex": "#ffc107"},
    "Q4": {"bg": "#fbe9e7", "fg": "#bf360c", "border": "#ffccbc", "hex": "#ff5722"},
}
Q_DEFAULT = {"bg": "#f3f4f6", "fg": "#374151", "border": "#d1d5db", "hex": "#9e9e9e"}


def q_badge(q: str, label: str | None = None) -> str:
    q = str(q).strip().upper()
    c = Q_COLORS.get(q, Q_DEFAULT)
    shown = label or (q if q else "—")
    return (
        f'<span style="display:inline-block;padding:0.20rem 0.62rem;'
        f"border-radius:999px;background:{c['bg']};color:{c['fg']};font-weight:700;"
        f'font-size:0.82rem;margin-right:0.35rem;border:1px solid {c["border"]};">'
        f"{shown}</span>"
    )


def yes_no_badge(flag: bool, label: str) -> str:
    bg = "#e8f5e9" if flag else "#f3f4f6"
    fg = "#1b5e20" if flag else "#6b7280"
    border = "#c8e6c9" if flag else "#d1d5db"
    symbol = "✓" if flag else "—"
    return (
        f'<span style="display:inline-block;padding:0.20rem 0.62rem;'
        f"border-radius:999px;background:{bg};color:{fg};font-weight:700;"
        f'font-size:0.82rem;margin-right:0.35rem;border:1px solid {border};">'
        f"{label} {symbol}</span>"
    )


def format_year(y) -> str:
    try:
        if pd.isna(y):
            return "—"
        return str(int(float(y)))
    except Exception:
        return str(y) if y else "—"


def format_citations(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    try:
        return str(int(x))
    except Exception:
        return str(x)


# ─────────────────────────── filters ────────────────────────────────────────

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros")

    url_params = read_query_params()

    # Reset button
    if st.sidebar.button("🔄 Repor filtros"):
        for key in list(st.session_state.keys()):
            if key.startswith("filter_"):
                del st.session_state[key]
        st.query_params.clear()
        st.rerun()

    anos_validos = df["ANO"].dropna().astype(int)
    if len(anos_validos) >= 2:
        ano_min_data, ano_max_data = int(anos_validos.min()), int(anos_validos.max())
        default_ano = (
            url_params.get("ano_min", ano_min_data),
            url_params.get("ano_max", ano_max_data),
        )
        ano_range = st.sidebar.slider(
            "Intervalo de anos",
            min_value=ano_min_data,
            max_value=ano_max_data,
            value=default_ano,
            key="filter_ano_range",
        )
    else:
        ano_range = None

    tipos = sorted([x for x in df["TIPO"].dropna().astype(str).unique() if x])
    tipos_sel = st.sidebar.multiselect(
        "Tipo de publicação",
        tipos,
        default=url_params.get("tipos", tipos),
        key="filter_tipos",
    )

    fontes = sorted([x for x in df["FONTE_PESQ"].dropna().astype(str).unique() if x])
    fontes_sel = st.sidebar.multiselect("Fonte / Venue", fontes, key="filter_fontes")

    best_q_opts = sorted([x for x in df["MELHOR_Q"].dropna().astype(str).unique() if x])
    melhor_q_sel = st.sidebar.multiselect(
        "Quartil",
        best_q_opts,
        default=url_params.get("quartil", []),
        key="filter_quartil",
    )

    bases_sel = st.sidebar.multiselect(
        "Indexado em",
        ["SCOPUS", "DBLP", "ISI"],
        default=url_params.get("bases", []),
        key="filter_bases",
    )

    st.sidebar.markdown("---")
    only_url = st.sidebar.checkbox("Apenas com URL", value=url_params.get("only_url", False), key="filter_url")
    only_q1 = st.sidebar.checkbox("Apenas Q1", value=url_params.get("only_q1", False), key="filter_q1")
    only_with_citations = st.sidebar.checkbox(
        "Apenas com alguma citação",
        value=url_params.get("only_cit", False),
        key="filter_cit",
    )

    # Guardar estado nos query params
    write_query_params({
        "search": st.session_state.get("filter_search", ""),
        "ano_min": ano_range[0] if ano_range else None,
        "ano_max": ano_range[1] if ano_range else None,
        "tipos": tipos_sel,
        "quartil": melhor_q_sel,
        "bases": bases_sel,
        "only_url": only_url,
        "only_q1": only_q1,
        "only_cit": only_with_citations,
    })

    out = df.copy()

    if ano_range is not None:
        out = out[
            (out["ANO"].fillna(-1).astype(int) >= ano_range[0])
            & (out["ANO"].fillna(-1).astype(int) <= ano_range[1])
        ]

    if tipos_sel:
        out = out[out["TIPO"].isin(tipos_sel)]
    if fontes_sel:
        out = out[out["FONTE_PESQ"].isin(fontes_sel)]
    if melhor_q_sel:
        out = out[out["MELHOR_Q"].isin(melhor_q_sel)]

    if "SCOPUS" in bases_sel:
        out = out[out["TEM_SCOPUS"]]
    if "DBLP" in bases_sel:
        out = out[out["TEM_DBLP"]]
    if "ISI" in bases_sel:
        out = out[out["TEM_ISI"]]

    search = st.session_state.get("filter_search", "")
    if search:
        out = out[out["SEARCH_TEXT"].str.contains(re.escape(search.lower().strip()), na=False)]
    if only_url:
        out = out[out["URL"].astype(str).str.strip().ne("")]
    if only_q1:
        out = out[out["MELHOR_Q"] == "Q1"]
    if only_with_citations:
        out = out[out[["CIT_SCOPUS", "CIT_SCHOLAR"]].fillna(0).sum(axis=1) > 0]

    return out.sort_values(
        by=["ANO", "TIPO", "TITULO_CURTO"], ascending=[False, True, True], na_position="last"
    )


# ─────────────────────────── publication renderers ──────────────────────────

def render_publication(pub: pd.Series):
    title = pub.get("TITULO_CURTO", "") or "Sem título"
    authors = pub.get("AUTOR", "") or "—"
    source = pub.get("FONTE_EXIBICAO", "") or "—"
    year = format_year(pub.get("ANO", ""))
    ptype = pub.get("TIPO", "") or "—"
    melhor_q = pub.get("MELHOR_Q", "") or ""
    jcr_q = pub.get("JCR_Q", "") or ""
    sjr_q = pub.get("SJR_Q", "") or ""
    url = str(pub.get("URL", "")).strip()

    def _clean(v):
        s = str(v).strip() if v is not None else ""
        return "" if s.lower() in ("nan", "none", "") else s
    jcr_val = _clean(pub.get("JCR", ""))
    jcr_ano = _clean(pub.get("JCRANO", ""))
    jcr_cat = _clean(pub.get("JCR_CATEGORIA", ""))
    sjr_val = _clean(pub.get("SJR", ""))
    sjr_ano = _clean(pub.get("SJRANO", ""))
    sjr_cat = _clean(pub.get("SJR_CATEGORIA", ""))

    st.markdown(f"### {title}")
    st.markdown(f"**Autor(es):** {authors}")
    st.markdown(f"**Fonte:** {source} &nbsp;&nbsp; **Ano:** {year} &nbsp;&nbsp; **Tipo:** {ptype}")

    badges_html = ""

    quartil_badges = ""
    if melhor_q:
        quartil_badges += q_badge(melhor_q, melhor_q)
    if jcr_q:
        quartil_badges += q_badge(jcr_q, f"JCR {jcr_q}")
    if sjr_q:
        quartil_badges += q_badge(sjr_q, f"SJR {sjr_q}")
    if quartil_badges:
        badges_html += f'<div style="margin:0.4rem 0 0.5rem;"><strong>Quartis:</strong> {quartil_badges}</div>'

    idx_badges = ""
    if pub.get("TEM_SCOPUS", False):
        idx_badges += yes_no_badge(True, "SCOPUS")
    if pub.get("TEM_DBLP", False):
        idx_badges += yes_no_badge(True, "DBLP")
    if pub.get("TEM_ISI", False):
        idx_badges += yes_no_badge(True, "ISI")
    if idx_badges:
        badges_html += f'<div style="margin-bottom:0.6rem;"><strong>Indexação:</strong> {idx_badges}</div>'

    if badges_html:
        st.markdown(badges_html, unsafe_allow_html=True)

    st.markdown(
        f"**Citações:** SCOPUS {format_citations(pub.get('CIT_SCOPUS'))} | Scholar {format_citations(pub.get('CIT_SCHOLAR'))}"
    )

    # Bloco JCR / SJR
    metrics_html = ""
    if jcr_val or jcr_cat:
        jcr_label = f"JCR {jcr_val}" if jcr_val else "JCR"
        jcr_label += f" ({jcr_ano})" if jcr_ano else ""
        jcr_detail = f"<br><span style='font-size:0.82rem;color:#555;'>{jcr_cat}</span>" if jcr_cat else ""
        metrics_html += (
            f'<div style="flex:1;min-width:220px;padding:0.6rem 0.8rem;border:1px solid #e2e8f0;'
            f'border-radius:10px;background:#f8fafc;">'
            f'<div style="font-size:0.78rem;color:#64748b;font-weight:600;text-transform:uppercase;'
            f'letter-spacing:0.05em;">JCR</div>'
            f'<div style="font-size:1rem;font-weight:700;color:#1e293b;">{jcr_label}</div>'
            f'{jcr_detail}</div>'
        )
    if sjr_val or sjr_cat:
        sjr_label = f"SJR {sjr_val}" if sjr_val else "SJR"
        sjr_label += f" ({sjr_ano})" if sjr_ano else ""
        sjr_detail = f"<br><span style='font-size:0.82rem;color:#555;'>{sjr_cat}</span>" if sjr_cat else ""
        metrics_html += (
            f'<div style="flex:1;min-width:220px;padding:0.6rem 0.8rem;border:1px solid #e2e8f0;'
            f'border-radius:10px;background:#f8fafc;">'
            f'<div style="font-size:0.78rem;color:#64748b;font-weight:600;text-transform:uppercase;'
            f'letter-spacing:0.05em;">SJR</div>'
            f'<div style="font-size:1rem;font-weight:700;color:#1e293b;">{sjr_label}</div>'
            f'{sjr_detail}</div>'
        )
    if metrics_html:
        st.markdown(
            f'<div style="display:flex;flex-wrap:wrap;gap:0.6rem;margin:0.5rem 0;">{metrics_html}</div>',
            unsafe_allow_html=True,
        )

    if url:
        st.markdown(f"**Ligação:** [abrir publicação]({url})")
    st.markdown("---")


def render_card_list_paginated(df: pd.DataFrame, page_size: int = 20):
    total = len(df)
    n_pages = max(1, (total + page_size - 1) // page_size)

    if n_pages > 1:
        col_left, col_mid, col_right = st.columns([1, 3, 1])
        with col_mid:
            page = st.number_input(
                f"Página (de {n_pages})",
                min_value=1,
                max_value=n_pages,
                value=1,
                step=1,
                key="page_number",
            )
    else:
        page = 1

    start = (page - 1) * page_size
    end = start + page_size
    page_df = df.iloc[start:end]

    for _, row in page_df.iterrows():
        render_publication(row)

    if n_pages > 1:
        st.caption(f"A mostrar {start + 1}–{min(end, total)} de {total} publicações")


def render_table_view(df: pd.DataFrame):
    """Tabela compacta e ordenável com células de quartil coloridas."""
    # Garantir colunas JCR/SJR mesmo que ausentes
    for col in ["JCR", "JCRANO", "JCR_CATEGORIA", "SJR", "SJRANO", "SJR_CATEGORIA"]:
        if col not in df.columns:
            df = df.copy()
            df[col] = ""

    display = df[
        ["ANO", "TITULO_CURTO", "AUTOR", "TIPO", "FONTE_EXIBICAO", "MELHOR_Q",
         "TEM_SCOPUS", "TEM_DBLP", "TEM_ISI", "CIT_SCOPUS", "CIT_SCHOLAR",
         "JCR", "JCRANO", "JCR_CATEGORIA", "SJR", "SJRANO", "SJR_CATEGORIA", "URL"]
    ].copy()

    display.columns = [
        "Ano", "Título", "Autor(es)", "Tipo", "Fonte", "Q",
        "SCOPUS", "DBLP", "ISI", "Cit. SCOPUS", "Cit. Scholar",
        "JCR", "JCR Ano", "JCR Categoria", "SJR", "SJR Ano", "SJR Categoria", "URL",
    ]
    display["Ano"] = display["Ano"].apply(format_year)
    # Manter como numérico (Int64) para permitir ordenação correta na tabela
    display["Cit. SCOPUS"] = pd.to_numeric(display["Cit. SCOPUS"], errors="coerce").astype("Int64")
    display["Cit. Scholar"] = pd.to_numeric(display["Cit. Scholar"], errors="coerce").astype("Int64")
    display["SCOPUS"] = display["SCOPUS"].map({True: "✓", False: "—"})
    display["DBLP"] = display["DBLP"].map({True: "✓", False: "—"})
    display["ISI"] = display["ISI"].map({True: "✓", False: "—"})
    display["URL"] = display["URL"].apply(
        lambda u: f"[link]({u})" if str(u).strip() else ""
    )

    def style_q(val):
        # Corrigido: usar .map() em vez do depreciado .applymap()
        c = Q_COLORS.get(str(val).strip(), Q_DEFAULT)
        return f"background-color: {c['bg']}; color: {c['fg']}; font-weight: bold"

    styled = display.style.map(style_q, subset=["Q", "JCR", "SJR"])
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ─────────────────────────── charts ─────────────────────────────────────────

def render_charts(df: pd.DataFrame):
    try:
        import plotly.express as px
        import plotly.graph_objects as go
    except ImportError:
        st.info("Instale `plotly` para ver gráficos: `pip install plotly`")
        return

    with st.expander("📊 Gráficos e análise", expanded=False):
        tab1, tab2, tab3, tab4 = st.tabs([
            "Evolução temporal", "Por tipo", "Por quartil", "Citações acumuladas"
        ])

        with tab1:
            year_counts = (
                df.dropna(subset=["ANO"])
                .groupby(["ANO", "TIPO"])
                .size()
                .reset_index(name="n")
            )
            if not year_counts.empty:
                fig = px.bar(
                    year_counts,
                    x="ANO",
                    y="n",
                    color="TIPO",
                    labels={"ANO": "Ano", "n": "Publicações", "TIPO": "Tipo"},
                    title="Publicações por ano e tipo",
                )
                fig.update_layout(legend_title_text="Tipo")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sem dados de ano disponíveis.")

        with tab2:
            tipo_counts = df["TIPO"].value_counts().reset_index()
            tipo_counts.columns = ["Tipo", "n"]
            fig2 = px.pie(tipo_counts, names="Tipo", values="n", title="Distribuição por tipo")
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            q_counts = df["MELHOR_Q"].replace("", "Sem quartil").value_counts().reset_index()
            q_counts.columns = ["Quartil", "n"]
            color_map = {k: v["hex"] for k, v in Q_COLORS.items()}
            color_map["Sem quartil"] = Q_DEFAULT["hex"]
            fig3 = px.bar(
                q_counts.sort_values("Quartil"),
                x="Quartil",
                y="n",
                color="Quartil",
                color_discrete_map=color_map,
                title="Publicações por quartil",
                labels={"n": "Publicações"},
            )
            fig3.update_layout(showlegend=False)
            st.plotly_chart(fig3, use_container_width=True)

        with tab4:
            _render_cumulative_citations(df, px, go)


def _render_cumulative_citations(df: pd.DataFrame, px, go):
    """Gráfico de citações acumuladas ao longo do tempo."""
    df_ano = df.dropna(subset=["ANO"]).copy()
    df_ano["ANO"] = df_ano["ANO"].astype(int)

    if df_ano.empty:
        st.info("Sem dados de ano disponíveis para calcular citações acumuladas.")
        return

    # Agregar citações por ano
    agg = (
        df_ano.groupby("ANO")
        .agg(
            CIT_SCOPUS=("CIT_SCOPUS", "sum"),
            CIT_SCHOLAR=("CIT_SCHOLAR", "sum"),
            N_PUBS=("TITULO_CURTO", "count"),
        )
        .reset_index()
        .sort_values("ANO")
    )

    agg["CIT_SCOPUS_CUM"] = agg["CIT_SCOPUS"].cumsum()
    agg["CIT_SCHOLAR_CUM"] = agg["CIT_SCHOLAR"].cumsum()
    agg["N_PUBS_CUM"] = agg["N_PUBS"].cumsum()

    col_a, col_b = st.columns(2)

    with col_a:
        fonte = st.selectbox(
            "Fonte de citações",
            ["SCOPUS", "Scholar", "Ambas"],
            key="chart_cit_fonte",
        )

    with col_b:
        modo = st.radio(
            "Modo",
            ["Acumulado", "Por ano"],
            horizontal=True,
            key="chart_cit_modo",
        )

    fig = go.Figure()

    if modo == "Acumulado":
        y_scopus = "CIT_SCOPUS_CUM"
        y_scholar = "CIT_SCHOLAR_CUM"
        titulo = "Citações acumuladas ao longo do tempo"
        y_label = "Citações (acumulado)"
    else:
        y_scopus = "CIT_SCOPUS"
        y_scholar = "CIT_SCHOLAR"
        titulo = "Citações por ano"
        y_label = "Citações"

    if fonte in ("SCOPUS", "Ambas"):
        fig.add_trace(go.Scatter(
            x=agg["ANO"], y=agg[y_scopus],
            mode="lines+markers",
            name="SCOPUS",
            line=dict(color=Q_COLORS["Q2"]["hex"], width=2),
            marker=dict(size=6),
        ))

    if fonte in ("Scholar", "Ambas"):
        fig.add_trace(go.Scatter(
            x=agg["ANO"], y=agg[y_scholar],
            mode="lines+markers",
            name="Scholar",
            line=dict(color=Q_COLORS["Q1"]["hex"], width=2),
            marker=dict(size=6),
        ))

    # Barra secundária com número de publicações por ano
    fig.add_trace(go.Bar(
        x=agg["ANO"],
        y=agg["N_PUBS"] if modo == "Por ano" else agg["N_PUBS_CUM"],
        name="Publicações",
        yaxis="y2",
        opacity=0.25,
        marker_color="#9e9e9e",
    ))

    fig.update_layout(
        title=titulo,
        xaxis_title="Ano",
        yaxis_title=y_label,
        yaxis2=dict(
            title="Publicações",
            overlaying="y",
            side="right",
            showgrid=False,
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
    )

    st.plotly_chart(fig, use_container_width=True)

    # Tabela resumo
    with st.expander("Ver dados detalhados"):
        resumo = agg[["ANO", "N_PUBS", "CIT_SCOPUS", "CIT_SCHOLAR",
                       "N_PUBS_CUM", "CIT_SCOPUS_CUM", "CIT_SCHOLAR_CUM"]].copy()
        resumo.columns = [
            "Ano", "Pubs/ano", "Cit. SCOPUS/ano", "Cit. Scholar/ano",
            "Pubs acum.", "Cit. SCOPUS acum.", "Cit. Scholar acum.",
        ]
        st.dataframe(resumo, use_container_width=True, hide_index=True)


# ─────────────────────────── export ─────────────────────────────────────────

def render_export(df: pd.DataFrame):
    export_cols = [
        "ANO", "TIPO", "TITULO_CURTO", "AUTOR", "FONTE_EXIBICAO",
        "MELHOR_Q", "JCR_Q", "SJR_Q",
        "TEM_SCOPUS", "TEM_DBLP", "TEM_ISI",
        "CIT_SCOPUS", "CIT_SCHOLAR", "CIT_ISI", "URL",
    ]
    export_df = df[[c for c in export_cols if c in df.columns]].copy()
    export_df.rename(columns={
        "ANO": "Ano", "TIPO": "Tipo", "TITULO_CURTO": "Título",
        "AUTOR": "Autor(es)", "FONTE_EXIBICAO": "Fonte",
        "MELHOR_Q": "Melhor Q", "JCR_Q": "JCR Q", "SJR_Q": "SJR Q",
        "TEM_SCOPUS": "SCOPUS", "TEM_DBLP": "DBLP", "TEM_ISI": "ISI",
        "CIT_SCOPUS": "Cit. SCOPUS", "CIT_SCHOLAR": "Cit. Scholar",
        "CIT_ISI": "Cit. ISI", "URL": "URL",
    }, inplace=True)

    col1, col2 = st.columns(2)
    with col1:
        csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇️ Exportar CSV",
            data=csv_bytes,
            file_name="publicacoes_filtradas.csv",
            mime="text/csv",
        )
    with col2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Publicações")
        st.download_button(
            label="⬇️ Exportar Excel",
            data=buf.getvalue(),
            file_name="publicacoes_filtradas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ─────────────────────────── main ───────────────────────────────────────────

def main():
    st.title(APP_TITLE)
    st.markdown("""
       <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .stDeployButton {display: none;}
        .main .block-container {
                max-width: 100%;
                padding-top: 1rem;
                padding-bottom: 1rem;
       </style>
    """, unsafe_allow_html=True)

    # ── Upload opcional do ficheiro Excel ────────────────────────────────────
    with st.sidebar.expander("📂 Carregar ficheiro Excel", expanded=False):
        uploaded = st.file_uploader(
            "Escolha um ficheiro .xlsx",
            type=["xlsx"],
            key="uploaded_excel",
            label_visibility="collapsed",
        )

    uploaded_bytes = uploaded.read() if uploaded is not None else None

    try:
        df = load_data(uploaded_bytes)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    # ── Barra de pesquisa no topo da página ─────────────────────────────────
    url_params = read_query_params()
    st.text_input(
        "🔍 Pesquisa rápida",
        value=url_params.get("search", ""),
        placeholder="Pesquisar por título, autores, fonte…",
        key="filter_search",
        label_visibility="collapsed",
    )

    filtered = apply_filters(df)

    # ── KPI metrics ──────────────────────────────────────────────────────────
    total_pubs = len(filtered)
    total_q1 = int((filtered["MELHOR_Q"] == "Q1").sum())
    total_cit_scopus = int(filtered["CIT_SCOPUS"].fillna(0).sum())
    total_cit_scholar = int(filtered["CIT_SCHOLAR"].fillna(0).sum())
    h_scopus = compute_h_index(filtered["CIT_SCOPUS"].tolist())
    h_scholar = compute_h_index(filtered["CIT_SCHOLAR"].tolist())

    cols = st.columns(6)
    for col, title, value in zip(
        cols,
        ["Total", "Q1", "Citations SCOPUS", "Citations Scholar",
         "h-index SCOPUS", "h-index Scholar"],
        [total_pubs, total_q1, total_cit_scopus, total_cit_scholar, h_scopus, h_scholar],
    ):
        with col:
            metric_card(title, str(value))

    # ── Charts ───────────────────────────────────────────────────────────────
    render_charts(filtered)

    # ── View toggle + export ─────────────────────────────────────────────────
    st.markdown("## Publications")

    top_left, top_right = st.columns([3, 1])
    with top_left:
        view = st.radio(
            "Vista",
            ["🃏 Cartões", "📋 Tabela"],
            horizontal=True,
            key="view_mode",
            label_visibility="collapsed",
        )
    with top_right:
        render_export(filtered)

    if filtered.empty:
        # Mostrar quais os filtros activos quando não há resultados
        active = []
        if st.session_state.get("filter_search", ""):
            active.append(f"pesquisa «{st.session_state['filter_search']}»")
        if st.session_state.get("filter_quartil"):
            active.append(f"quartil {st.session_state['filter_quartil']}")
        if st.session_state.get("filter_bases"):
            active.append(f"indexado em {st.session_state['filter_bases']}")
        msg = "Não existem publicações com os filtros atuais."
        if active:
            msg += f" Filtros ativos: {', '.join(active)}."
        st.info(msg)
        return

    st.caption(f"{total_pubs} publicação(ões) encontrada(s)")

    if view == "🃏 Cartões":
        sort_col, sort_dir_col = st.columns([2, 1])
        with sort_col:
            sort_by = st.selectbox(
                "Ordenar por",
                ["Ano", "Citações SCOPUS", "Citações Scholar"],
                key="sort_by",
                label_visibility="collapsed",
            )
        with sort_dir_col:
            sort_asc = st.radio(
                "Direção",
                ["↓ Desc", "↑ Asc"],
                horizontal=True,
                key="sort_dir",
                label_visibility="collapsed",
            ) == "↑ Asc"

        sort_map = {
            "Ano": "ANO",
            "Citações SCOPUS": "CIT_SCOPUS",
            "Citações Scholar": "CIT_SCHOLAR",
        }
        sort_field = sort_map[sort_by]
        sorted_filtered = filtered.sort_values(
            by=sort_field, ascending=sort_asc, na_position="last"
        )
        render_card_list_paginated(sorted_filtered, page_size=20)
    else:
        render_table_view(filtered)


if __name__ == "__main__":
    main()
