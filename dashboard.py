import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import json
import io
from datetime import datetime

st.set_page_config(
    page_title="Dashboard de Acessos",
    page_icon="📊",
    layout="wide",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  .main { background: #0f1117; }

  /* Header */
  .dash-header {
    background: linear-gradient(135deg, #1a1f2e 0%, #0f1117 100%);
    border: 1px solid #2a2f45;
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 20px;
  }
  .dash-title { font-size: 26px; font-weight: 700; color: #e2e8f0; margin: 0; }
  .dash-sub   { font-size: 13px; color: #64748b; margin: 4px 0 0 0; font-family: 'JetBrains Mono'; }
  .dash-badge {
    background: #1e40af22;
    border: 1px solid #1e40af;
    color: #60a5fa;
    font-size: 11px;
    font-weight: 600;
    padding: 4px 12px;
    border-radius: 99px;
    font-family: 'JetBrains Mono';
  }

  /* KPI cards */
  .kpi-row { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
  .kpi-card {
    flex: 1; min-width: 140px;
    background: #1a1f2e;
    border: 1px solid #2a2f45;
    border-radius: 10px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
  }
  .kpi-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0;
    height: 3px;
  }
  .kpi-card.blue::before  { background: #3b82f6; }
  .kpi-card.green::before { background: #22c55e; }
  .kpi-card.amber::before { background: #f59e0b; }
  .kpi-card.red::before   { background: #ef4444; }
  .kpi-card.teal::before  { background: #14b8a6; }
  .kpi-card.purple::before{ background: #a855f7; }

  .kpi-label { font-size: 11px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: .08em; }
  .kpi-value { font-size: 32px; font-weight: 700; color: #e2e8f0; font-family: 'JetBrains Mono'; line-height: 1.1; margin: 6px 0 2px; }
  .kpi-icon  { font-size: 22px; position: absolute; top: 16px; right: 16px; opacity: .25; }

  /* Section title */
  .sec-title {
    font-size: 13px; font-weight: 600; color: #94a3b8;
    text-transform: uppercase; letter-spacing: .1em;
    margin: 0 0 12px 0; border-left: 3px solid #3b82f6;
    padding-left: 10px;
  }

  /* Peak chip */
  .peak-chip {
    display: inline-block;
    background: #1e40af22; border: 1px solid #1e40af55;
    color: #93c5fd; font-family: 'JetBrains Mono';
    font-size: 12px; padding: 4px 12px; border-radius: 99px;
    margin: 4px 4px 4px 0;
  }

  /* JSON area */
  .json-box {
    background: #0d1117; border: 1px solid #2a2f45; border-radius: 8px;
    padding: 16px; font-family: 'JetBrains Mono'; font-size: 12px;
    color: #7dd3fc; max-height: 340px; overflow-y: auto;
    white-space: pre-wrap; word-break: break-all;
  }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_excel(file) -> dict | None:
    """Parse the specific hourly-report Excel format and return a structured dict."""
    try:
        df_raw = pd.read_excel(file, sheet_name=0, header=None)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None

    # Locate date (row 1, col 0) and header rows (rows 2 & 3)
    date_val = None
    for r in range(min(5, len(df_raw))):
        v = df_raw.iloc[r, 0]
        if isinstance(v, (datetime, pd.Timestamp)):
            date_val = pd.Timestamp(v).strftime("%d/%m/%Y")
            break
        if isinstance(v, str) and v.strip() and v.strip()[0].isdigit():
            date_val = v.strip()
            break

    # Find the "De:" / "Ate:" row — that is the real data header
    header_row = None
    for r in range(len(df_raw)):
        cell = str(df_raw.iloc[r, 0]).strip()
        if cell.lower() in ("de:", "de"):
            header_row = r
            break

    if header_row is None:
        st.error("Não foi possível identificar a estrutura de cabeçalho na planilha.")
        return None

    # Build clean column names from the two-row header (category + sub-label)
    cat_row = df_raw.iloc[header_row - 1]  # Entradas / Total Pagam. / ...
    sub_row = df_raw.iloc[header_row]       # De: / Ate: / Qtd. / [%] ...

    cats = []
    current_cat = "Período"
    for i, val in enumerate(sub_row):
        v = str(val).strip() if pd.notna(val) else ""
        c = str(cat_row.iloc[i]).strip() if pd.notna(cat_row.iloc[i]) else ""
        if c and c.lower() not in ("nan", "de:", "ate:"):
            current_cat = c
        cats.append((current_cat, v if v else f"col{i}"))

    # Data rows: from header_row+1 until a row that starts with "TOTAL" or NaN cascade
    rows = []
    for r in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[r]
        first = str(row.iloc[0]).strip()
        if first.lower().startswith("total") or first.lower() == "nan":
            # grab totals if this is the TOTAL row
            if first.lower().startswith("total"):
                total_row = row
            continue
        rows.append(row)

    if not rows:
        st.error("Nenhuma linha de dados encontrada.")
        return None

    data_df = pd.DataFrame([r.values for r in rows], columns=[f"{c[0]}|{c[1]}" for c in cats])
    data_df = data_df.reset_index(drop=True)

    # Clean up: extract hour blocks as strings
    def fmt_time(v):
        import datetime as _dt
        if isinstance(v, _dt.time):
            return v.strftime("%H:%M:%S")
        return str(v).strip()

    data_df["hora_inicio"] = data_df["Período|De:"].apply(fmt_time)
    data_df["hora_fim"]    = data_df["Período|Ate:"].apply(fmt_time)

    def to_num(val):
        """Convert a scalar or Series to a numeric value, defaulting to 0."""
        import math
        try:
            result = pd.to_numeric(val, errors="coerce")
            # If it's a Series, fill NaN; if scalar, handle NaN manually
            if isinstance(result, pd.Series):
                return result.fillna(0)
            return 0 if (result is None or (isinstance(result, float) and math.isnan(result))) else result
        except Exception:
            return 0

    import re as _re
    records = []
    for _, row in data_df.iterrows():
        h_start = str(row["hora_inicio"]).strip()
        h_end   = str(row["hora_fim"]).strip()

        # Skip any row whose time fields don't look like HH:MM:SS or HH:MM
        if not _re.match(r"^\d{1,2}:\d{2}", h_start):
            continue

        record = {
            "periodo": {"de": h_start, "ate": h_end},
            "entradas":     {"qtd": int(to_num(row.get("Entradas|Qtd.",  0))), "pct": float(to_num(row.get("Entradas|[%]",    0)))},
            "total_pagam":  {"qtd": int(to_num(row.get("Total Pagam.|Qtd.", 0))), "pct": float(to_num(row.get("Total Pagam.|[%]", 0)))},
            "valores":      {"val": float(to_num(row.get("Valores|Val.",  0))), "pct": float(to_num(row.get("Valores|[%]",    0)))},
            "saidas":       {"qtd": int(to_num(row.get("Saídas|Qtd.",   0))), "pct": float(to_num(row.get("Saídas|[%]",   0)))},
            "cancelados":   {"qtd": int(to_num(row.get("Cancelados|Qtd.", 0))), "pct": float(to_num(row.get("Cancelados|[%]", 0)))},
            "carencia":     {"qtd": int(to_num(row.get("Carência|Qtd.", 0))), "pct": float(to_num(row.get("Carência|[%]",  0)))},
        }
        records.append(record)

    # Totals row (best-effort)
    totals = {}
    try:
        tr = df_raw[df_raw.iloc[:, 0].astype(str).str.lower().str.startswith("total")]
        if not tr.empty:
            tr = tr.iloc[0]
            cols = [f"{c[0]}|{c[1]}" for c in cats]
            tr_s = pd.Series(tr.values, index=cols)
            totals = {
                "entradas":    int(to_num(tr_s.get("Entradas|Qtd.",      0))),
                "total_pagam": int(to_num(tr_s.get("Total Pagam.|Qtd.",  0))),
                "valores":     float(to_num(tr_s.get("Valores|Val.",     0))),
                "saidas":      int(to_num(tr_s.get("Saídas|Qtd.",        0))),
                "cancelados":  int(to_num(tr_s.get("Cancelados|Qtd.",    0))),
                "carencia":    int(to_num(tr_s.get("Carência|Qtd.",      0))),
            }
    except Exception:
        pass

    return {
        "data": date_val or "—",
        "registros": records,
        "totais": totals,
    }


def make_plotly_theme():
    return dict(
        paper_bgcolor="#1a1f2e",
        plot_bgcolor="#1a1f2e",
        font=dict(family="Inter, sans-serif", color="#94a3b8", size=12),
        xaxis=dict(gridcolor="#2a2f45", zerolinecolor="#2a2f45", tickcolor="#64748b"),
        yaxis=dict(gridcolor="#2a2f45", zerolinecolor="#2a2f45", tickcolor="#64748b"),
        margin=dict(t=30, b=40, l=40, r=20),
    )


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="dash-header">
  <div>
    <p class="dash-title">📊 Dashboard de Acessos</p>
    <p class="dash-sub">Carregue um arquivo Excel (.xlsx) para visualizar os dados</p>
  </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Selecione o arquivo Excel",
    type=["xlsx", "xls"],
    help="Planilha com dados de controle de acesso por período",
)

if not uploaded:
    st.info("⬆️  Nenhum arquivo carregado. Selecione um arquivo .xlsx para começar.")
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
with st.spinner("Lendo e convertendo planilha para JSON…"):
    result = parse_excel(uploaded)

if result is None:
    st.stop()

records = result["registros"]
totais  = result["totais"]
data    = result["data"]

# Build DataFrame from JSON records — guard against any malformed records
records = [r for r in records if isinstance(r, dict) and isinstance(r.get("periodo"), dict)]
active = [r for r in records if r["entradas"]["qtd"] > 0 or r["saidas"]["qtd"] > 0]
if not active:
    active = records  # fallback

horas       = [str(r["periodo"]["de"])[:5] for r in active]
entradas    = [r["entradas"]["qtd"]    for r in active]
pagamentos  = [r["total_pagam"]["qtd"] for r in active]
valores_val = [r["valores"]["val"]     for r in active]
saidas      = [r["saidas"]["qtd"]      for r in active]
cancelados  = [r["cancelados"]["qtd"]  for r in active]
carencia    = [r["carencia"]["qtd"]    for r in active]

pico_hora = horas[entradas.index(max(entradas))] if entradas else "—"
pico_saida = horas[saidas.index(max(saidas))] if saidas else "—"

# ── KPI row ───────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="display:flex; align-items:center; gap:12px; margin-bottom:16px;">
  <p class="sec-title" style="margin:0">Resumo do dia</p>
  <span class="dash-badge">{data}</span>
</div>
<div class="kpi-row">
  <div class="kpi-card blue">
    <span class="kpi-icon">🚪</span>
    <div class="kpi-label">Entradas</div>
    <div class="kpi-value">{totais.get('entradas', sum(entradas)):,}</div>
  </div>
  <div class="kpi-card green">
    <span class="kpi-icon">💳</span>
    <div class="kpi-label">Pagamentos</div>
    <div class="kpi-value">{totais.get('total_pagam', sum(pagamentos)):,}</div>
  </div>
  <div class="kpi-card amber">
    <span class="kpi-icon">💰</span>
    <div class="kpi-label">Valor Total (R$)</div>
    <div class="kpi-value">{totais.get('valores', sum(valores_val)):,.0f}</div>
  </div>
  <div class="kpi-card teal">
    <span class="kpi-icon">🏃</span>
    <div class="kpi-label">Saídas</div>
    <div class="kpi-value">{totais.get('saidas', sum(saidas)):,}</div>
  </div>
  <div class="kpi-card red">
    <span class="kpi-icon">❌</span>
    <div class="kpi-label">Cancelados</div>
    <div class="kpi-value">{totais.get('cancelados', sum(cancelados)):,}</div>
  </div>
  <div class="kpi-card purple">
    <span class="kpi-icon">⏳</span>
    <div class="kpi-label">Carência</div>
    <div class="kpi-value">{totais.get('carencia', sum(carencia)):,}</div>
  </div>
</div>
<div style="margin-bottom:24px;">
  <span style="font-size:12px; color:#64748b;">Pico de entradas: </span>
  <span class="peak-chip">{pico_hora}h</span>
  <span style="font-size:12px; color:#64748b; margin-left:8px;">Pico de saídas: </span>
  <span class="peak-chip">{pico_saida}h</span>
</div>
""", unsafe_allow_html=True)

# ── Charts ────────────────────────────────────────────────────────────────────
theme = make_plotly_theme()

col1, col2 = st.columns([3, 2])

with col1:
    st.markdown('<p class="sec-title">Fluxo por hora</p>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Entradas", x=horas, y=entradas,
                         marker_color="#3b82f6", opacity=.85))
    fig.add_trace(go.Bar(name="Saídas",   x=horas, y=saidas,
                         marker_color="#14b8a6", opacity=.85))
    fig.add_trace(go.Scatter(name="Pagamentos", x=horas, y=pagamentos,
                             mode="lines+markers", line=dict(color="#f59e0b", width=2),
                             marker=dict(size=5)))
    fig.update_layout(**theme, barmode="group", height=300,
                      legend=dict(orientation="h", y=1.12, x=0, bgcolor="#0f1117", bordercolor="#2a2f45", borderwidth=1))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.markdown('<p class="sec-title">Distribuição geral</p>', unsafe_allow_html=True)
    labels = ["Entradas", "Pagamentos", "Saídas", "Carência", "Cancelados"]
    values = [
        totais.get("entradas", sum(entradas)),
        totais.get("total_pagam", sum(pagamentos)),
        totais.get("saidas", sum(saidas)),
        totais.get("carencia", sum(carencia)),
        totais.get("cancelados", sum(cancelados)),
    ]
    colors = ["#3b82f6", "#f59e0b", "#14b8a6", "#a855f7", "#ef4444"]
    fig2 = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=.55, marker_colors=colors,
        textinfo="label+percent",
        textfont=dict(size=11, color="#e2e8f0"),
    ))
    fig2.update_layout(**theme, height=300, showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

# ── Valores por hora ──────────────────────────────────────────────────────────
st.markdown('<p class="sec-title">Faturamento por hora (R$)</p>', unsafe_allow_html=True)
fig3 = go.Figure(go.Bar(
    x=horas, y=valores_val,
    marker=dict(
        color=valores_val,
        colorscale=[[0, "#1e3a5f"], [0.5, "#2563eb"], [1, "#93c5fd"]],
        showscale=False,
    ),
    text=[f"R${v:,.0f}" if v > 0 else "" for v in valores_val],
    textposition="outside",
    textfont=dict(size=10, color="#94a3b8"),
))
fig3.update_layout(**theme, height=240,
)
fig3.update_yaxes(tickprefix="R$ ", gridcolor="#2a2f45", zerolinecolor="#2a2f45", tickcolor="#64748b")
st.plotly_chart(fig3, use_container_width=True)

# ── Data table ────────────────────────────────────────────────────────────────
with st.expander("📋  Ver tabela completa", expanded=False):
    table_df = pd.DataFrame({
        "Hora":        horas,
        "Entradas":    entradas,
        "Ent. %":      [r["entradas"]["pct"]   for r in active],
        "Pagamentos":  pagamentos,
        "Pag. %":      [r["total_pagam"]["pct"] for r in active],
        "Valores R$":  valores_val,
        "Saídas":      saidas,
        "Cancelados":  cancelados,
        "Carência":    carencia,
    })
    st.dataframe(table_df, use_container_width=True, hide_index=True)

# ── JSON viewer / download ────────────────────────────────────────────────────
with st.expander("🔧  Ver / baixar JSON gerado", expanded=False):
    json_str = json.dumps(result, ensure_ascii=False, indent=2)
    st.markdown(f'<div class="json-box">{json_str}</div>', unsafe_allow_html=True)
    st.download_button(
        "⬇️  Baixar JSON",
        data=json_str.encode("utf-8"),
        file_name=f"acessos_{data.replace('/', '-')}.json",
        mime="application/json",
    )
