import re
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Dashboard Consultórios", layout="wide")

# --- Corporate styling ---
st.markdown("""
<style>
.block-container {padding-top: 1.5rem;}
div[data-testid="stMetricValue"] {color:#0F4C81;}
h1, h2, h3 { color:#1f2a44; }
section[data-testid="stSidebar"] {background-color:#f5f7fb}
</style>
""", unsafe_allow_html=True)

st.title("🏥 Dashboard de Ocupação dos Consultórios")
st.caption("Lendo somente as abas **CONSULTÓRIO** (ignorando 'OCUPAÇÃO DAS SALAS'). Integra automaticamente TODAS as abas **MÉDICOS** (ex.: 'MÉDICOS 1', 'MÉDICOS 2', 'MÉDICOS 3').")

DEFAULT_PATH = Path("/mnt/data/ESCALA DOS CONSULTORIOS DEFINITIVO.xlsx")

# ---------- Sidebar: Upload ----------
st.sidebar.header("📂 Fonte de Dados")
uploaded = st.sidebar.file_uploader("Envie o Excel (.xlsx)", type=["xlsx"], key="main_xlsx")

def load_excel(file_like):
    try:
        return pd.ExcelFile(file_like)
    except Exception as e:
        st.error(f"Não foi possível abrir o arquivo: {e}")
        return None

excel = None
if uploaded is not None:
    excel = load_excel(uploaded)
    fonte = "Upload do usuário"
elif DEFAULT_PATH.exists():
    excel = load_excel(DEFAULT_PATH)
    fonte = f"Arquivo padrão: {DEFAULT_PATH.name}"
else:
    st.error("Nenhum arquivo encontrado. Envie um Excel com as abas de CONSULTÓRIO.")
    st.stop()

st.sidebar.success(f"Usando dados de: {fonte}")

# ---------- Utilitários ----------
def _normalize_col(col):
    c = str(col).strip().lower()
    c = (c
         .replace("á","a").replace("ã","a").replace("â","a")
         .replace("é","e").replace("ê","e")
         .replace("í","i").replace("î","i")
         .replace("ó","o").replace("õ","o").replace("ô","o")
         .replace("ú","u").replace("ü","u")
         .replace("ç","c"))
    c = re.sub(r"\s+", " ", c)
    return c

def detect_header_and_parse(excel, sheet_name):
    for header in [0,1,2,3,4]:
        try:
            df = excel.parse(sheet_name, header=header)
        except Exception:
            continue
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            continue

        cols_norm = [_normalize_col(c) for c in df.columns]
        col_dia = None; col_manha=None; col_tarde=None

        for i, cn in enumerate(cols_norm):
            if col_dia is None:
                if "dia" in cn or any(d in cn for d in ["segunda","terca","terça","quarta","quinta","sexta","sabado","sábado"]):
                    col_dia = df.columns[i]
            if any(k in cn for k in ["manha","manhã"]): col_manha = df.columns[i]
            if "tarde" in cn: col_tarde = df.columns[i]

        # fallback: primeira coluna contém dias
        if col_dia is None and len(df.columns) >= 1:
            first_col = df.columns[0]
            sample = df[first_col].astype(str).str.lower()
            if sample.str.contains("segunda|terca|terça|quarta|quinta|sexta|sabado|sábado").any():
                col_dia = first_col

        if col_dia is not None and (col_manha is not None or col_tarde is not None):
            use_cols = [c for c in [col_dia, col_manha, col_tarde] if c is not None]
            df = df[use_cols].copy()
            rename = {col_dia:"Dia"}
            if col_manha is not None: rename[col_manha]="Manhã"
            if col_tarde is not None: rename[col_tarde]="Tarde"
            df = df.rename(columns=rename)
            df["Dia"] = df["Dia"].astype(str).str.strip()
            df = df[df["Dia"].str.len()>0]
            return df
    return None

def tidy_from_sheets(excel):
    frames = []
    for sheet in excel.sheet_names:
        s_norm = _normalize_col(sheet)
        if ("consult" in s_norm) and ("ocupa" not in s_norm):
            df = detect_header_and_parse(excel, sheet)
            if df is None or df.empty:
                continue
            df["Dia"] = (df["Dia"].astype(str).str.strip()
                         .str.replace("terca","terça", case=False)
                         .str.replace("sabado","sábado", case=False)
                         .str.capitalize())
            df.insert(0, "Sala", sheet.strip())
            long = df.melt(id_vars=["Sala","Dia"], value_vars=[c for c in ["Manhã","Tarde"] if c in df.columns],
                           var_name="Turno", value_name="Médico")
            long["Médico"] = long["Médico"].astype(str).replace({"nan":"","None":""}).str.strip()
            frames.append(long)
    if not frames:
        return pd.DataFrame(columns=["Sala","Dia","Turno","Médico"])
    full = pd.concat(frames, ignore_index=True)
    full["Dia"] = pd.Categorical(full["Dia"], categories=["Segunda","Terça","Quarta","Quinta","Sexta","Sábado"], ordered=True)
    full["Ocupado"] = full["Médico"].str.len() > 0
    return full

df = tidy_from_sheets(excel)
if df.empty:
    st.error("Não foram encontrados dados nas abas 'CONSULTÓRIO'.")
    st.stop()

# ---------- Filtros ----------
st.sidebar.header("🔎 Filtros")
salas = sorted(df["Sala"].dropna().unique().tolist())
dias = [d for d in ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado"] if d in df["Dia"].astype(str).unique()]
turnos = sorted(df["Turno"].dropna().unique().tolist())
medicos = sorted([m for m in df["Médico"].dropna().unique().tolist() if m])

sel_salas = st.sidebar.multiselect("Consultório(s)", salas, default=salas)
sel_dias = st.sidebar.multiselect("Dia(s)", dias, default=dias)
sel_turnos = st.sidebar.multiselect("Turno(s)", turnos, default=turnos)
sel_medicos = st.sidebar.multiselect("Médico(s)", medicos, default=[], help="Deixe vazio para não filtrar por médico.")

# Base para KPIs (NÃO filtra por médico)
mask_base = (df["Sala"].isin(sel_salas) & df["Dia"].astype(str).isin(sel_dias) & df["Turno"].isin(sel_turnos))
fdf_base = df[mask_base].copy()

# Aplicar filtro de médico apenas onde fizer sentido
mask_medico = df["Médico"].isin(sel_medicos) if sel_medicos else True
fdf = df[mask_base & mask_medico].copy()

# ---------- KPIs ----------
total_salas = len(set(sel_salas))
total_slots = len(fdf_base)
ocupados = int(fdf_base["Ocupado"].sum())
tx_ocup = (ocupados / total_slots * 100) if total_slots > 0 else 0
slots_livres = max(total_slots - ocupados, 0)
medicos_distintos = fdf_base.loc[fdf_base["Ocupado"], "Médico"].nunique()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Consultórios selecionados", total_salas)
c2.metric("Slots (dia x turno x sala)", total_slots)
c3.metric("Slots livres", slots_livres)
c4.metric("Ocupados", ocupados)

kc1, kc2 = st.columns(2)
kc1.metric("Taxa de ocupação", f"{tx_ocup:.1f}%")
kc2.metric("Médicos distintos (no filtro de sala/dia/turno)", medicos_distintos)

# ---------- Gráficos de ocupação (sem heatmap) com porcentagens nas barras ----------
colA, colB = st.columns(2)
with colA:
    by_sala = fdf_base.groupby("Sala")["Ocupado"].mean().reset_index()
    by_sala["Taxa de Ocupação (%)"] = (by_sala["Ocupado"]*100).round(1)
    fig1 = px.bar(by_sala, x="Sala", y="Taxa de Ocupação (%)", title="Ocupação por Consultório (%)", text="Taxa de Ocupação (%)")
    fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig1.update_yaxes(range=[0,100])
    st.plotly_chart(fig1, use_container_width=True)

with colB:
    by_dia = fdf_base.groupby("Dia")["Ocupado"].mean().reset_index()
    by_dia["Taxa de Ocupação (%)"] = (by_dia["Ocupado"]*100).round(1)
    fig2 = px.bar(by_dia, x="Dia", y="Taxa de Ocupação (%)", title="Ocupação por Dia da Semana (%)", text="Taxa de Ocupação (%)")
    fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig2.update_yaxes(range=[0,100])
    st.plotly_chart(fig2, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    by_turno = fdf_base.groupby("Turno")["Ocupado"].mean().reset_index()
    by_turno["Taxa de Ocupação (%)"] = (by_turno["Ocupado"]*100).round(1)
    fig3 = px.bar(by_turno, x="Turno", y="Taxa de Ocupação (%)", title="Ocupação por Turno (%)", text="Taxa de Ocupação (%)")
    fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig3.update_yaxes(range=[0,100])
    st.plotly_chart(fig3, use_container_width=True)

with colD:
    top_med = (fdf[fdf["Ocupado"]]
               .groupby("Médico")
               .size()
               .reset_index(name="Turnos Utilizados")
               .sort_values("Turnos Utilizados", ascending=False)
               .head(15))
    if not top_med.empty:
        fig4 = px.bar(top_med, x="Turnos Utilizados", y="Médico", orientation="h", title="Top Médicos por Nº de Turnos", text="Turnos Utilizados")
        fig4.update_traces(textposition="outside")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Sem médicos ocupando slots nos filtros atuais.")

# ---------- Integração das abas MÉDICOS (1, 2, 3...) ----------
def _to_number(x):
    import numpy as np, re as _re
    if pd.isna(x):
        return np.nan
    txt = str(x)
    txt = _re.sub(r"[^\d,.-]", "", txt)
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt and "." not in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except:
        return pd.NA

def load_medicos_from_excel(excel: pd.ExcelFile):
    frames = []
    for s in excel.sheet_names:
        sn = _normalize_col(s)
        if "medic" in sn:  # captura "médicos", "medicos"
            try:
                dfm = excel.parse(s, header=0)
            except Exception:
                continue
            if dfm is None or dfm.empty:
                continue
            # normaliza colunas
            norm = {c:_normalize_col(c) for c in dfm.columns}
            dfm.columns = [norm[c] for c in dfm.columns]
            rename = {}
            for c in dfm.columns:
                if "nome" in c or "medico" in c: rename[c]="Médico"
                if c=="crm" or "crm" in c: rename[c]="CRM"
                if "especial" in c: rename[c]="Especialidade"
                if "planos" in c or c=="plano": rename[c]="Planos"
                if "valor" in c or "aluguel" in c or "negoci" in c: rename[c]="Valor Aluguel"
                if "exclus" in c: rename[c]="Sala Exclusiva"
                if "divid" in c: rename[c]="Sala Dividida"
            dfm = dfm.rename(columns=rename)
            keep = [c for c in ["Médico","CRM","Especialidade","Planos","Sala Exclusiva","Sala Dividida","Valor Aluguel"] if c in dfm.columns]
            if not keep:
                continue
            dfm = dfm[keep].copy()
            frames.append(dfm)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    # normalizações finais
    if "Médico" in out.columns: out["Médico"] = out["Médico"].astype(str).str.strip()
    if "Planos" in out.columns: out["Planos"] = out["Planos"].astype(str).str.strip()
    if "Valor Aluguel" in out.columns: out["Valor Aluguel"] = out["Valor Aluguel"].apply(_to_number)
    for c in ["Sala Exclusiva","Sala Dividida"]:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip().str.upper().replace({"X":"Sim","":""})
    return out

med_df = load_medicos_from_excel(excel)

if med_df.empty:
    st.warning("Não foram encontradas abas de **MÉDICOS** no arquivo. Os indicadores de plano/aluguel ficarão ocultos.")
else:
    # Enriquecer com turnos utilizados
    usos = fdf_base.groupby("Médico").size().reset_index(name="Turnos Utilizados")
    med_enriched = med_df.merge(usos, on="Médico", how="left")

    st.markdown("---")
    st.subheader("💼 Indicador: PLANOS × Aluguel × Profissionais")

    # KPIs deste bloco
    tot_prof = med_enriched["Médico"].nunique()
    categorias_planos = med_enriched["Planos"].nunique() if "Planos" in med_enriched.columns else 0
    cpa, cpb, cpc = st.columns(3)
    cpa.metric("Profissionais (total)", tot_prof)
    cpb.metric("Categorias em PLANOS", categorias_planos)
    if "Valor Aluguel" in med_enriched.columns:
        media_valor = med_enriched["Valor Aluguel"].dropna().mean()
        cpc.metric("Valor médio de aluguel (R$)", f"{media_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
    else:
        cpc.metric("Valor médio de aluguel (R$)", "—")

    g1, g2 = st.columns(2)
    with g1:
        if "Planos" in med_enriched.columns:
            cont = med_enriched.groupby("Planos")["Médico"].nunique().reset_index(name="Profissionais")
            fig7 = px.bar(cont, x="Planos", y="Profissionais", title="Profissionais por PLANOS", text="Profissionais")
            fig7.update_traces(textposition="outside")
            st.plotly_chart(fig7, use_container_width=True)
        else:
            st.info("Coluna PLANOS não encontrada.")

    with g2:
        if "Valor Aluguel" in med_enriched.columns and "Planos" in med_enriched.columns:
            avgv = med_enriched.groupby("Planos")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)")
            avgv["Valor médio (R$)"] = avgv["Valor médio (R$)"].round(2)
            fig8 = px.bar(avgv, x="Planos", y="Valor médio (R$)", title="Valor médio de aluguel por PLANOS", text="Valor médio (R$)")
            fig8.update_traces(texttemplate="R$ %{y:.2f}", textposition="outside")
            st.plotly_chart(fig8, use_container_width=True)
        else:
            st.info("Inclua as colunas PLANOS e Valor Aluguel.")

    if "Valor Aluguel" in med_enriched.columns:
        st.markdown("##### Distribuição de profissionais por faixa de aluguel × PLANOS")
        bins = [0,500,1000,1500,2000,3000,9999999]
        labels = ["até 500","501–1000","1001–1500","1501–2000","2001–3000","3000+"]
        med_enriched["Faixa Aluguel"] = pd.cut(med_enriched["Valor Aluguel"], bins=bins, labels=labels, include_lowest=True)
        dist = med_enriched.groupby(["Planos","Faixa Aluguel"])["Médico"].nunique().reset_index(name="Profissionais")
        fig9 = px.bar(dist, x="Faixa Aluguel", y="Profissionais", color="Planos", barmode="group",
                      title="Profissionais por faixa de aluguel × PLANOS", text="Profissionais")
        fig9.update_traces(textposition="outside")
        st.plotly_chart(fig9, use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        if "Especialidade" in med_enriched.columns and "Valor Aluguel" in med_enriched.columns:
            esp_avg = med_enriched.groupby("Especialidade")["Valor Aluguel"].mean().reset_index(name="Valor médio (R$)").sort_values("Valor médio (R$)", ascending=False)
            fig10 = px.bar(esp_avg, x="Valor médio (R$)", y="Especialidade", orientation="h", title="Valor médio de aluguel por especialidade", text="Valor médio (R$)")
            fig10.update_traces(texttemplate="R$ %{x:.2f}", textposition="outside")
            st.plotly_chart(fig10, use_container_width=True)
        else:
            st.info("Inclua 'Especialidade' e 'Valor Aluguel'.")
    with g4:
        if "Planos" in med_enriched.columns and "Especialidade" in med_enriched.columns:
            plano_esp = med_enriched.groupby(["Especialidade","Planos"])["Médico"].nunique().reset_index(name="Profissionais")
            fig11 = px.bar(plano_esp, x="Especialidade", y="Profissionais", color="Planos", barmode="group",
                           title="Profissionais por especialidade × PLANOS", text="Profissionais")
            fig11.update_traces(textposition="outside")
            st.plotly_chart(fig11, use_container_width=True)
        else:
            st.info("Inclua 'Especialidade' e 'PLANOS'.")

    g5, g6 = st.columns(2)
    with g5:
        if "Sala Exclusiva" in med_enriched.columns or "Sala Dividida" in med_enriched.columns:
            ts = med_enriched.copy()
            ts["Tipo de Sala"] = None
            if "Sala Exclusiva" in ts.columns:
                ts.loc[ts["Sala Exclusiva"].eq("Sim"), "Tipo de Sala"] = "Exclusiva"
            if "Sala Dividida" in ts.columns:
                ts.loc[ts["Sala Dividida"].eq("Sim"), "Tipo de Sala"] = ts["Tipo de Sala"].fillna("Dividida")
            ts = ts.dropna(subset=["Tipo de Sala"])
            if not ts.empty:
                dist_ts = ts.groupby("Tipo de Sala")["Médico"].nunique().reset_index(name="Profissionais")
                fig12 = px.bar(dist_ts, x="Tipo de Sala", y="Profissionais", title="Profissionais por tipo de sala", text="Profissionais")
                fig12.update_traces(textposition="outside")
                st.plotly_chart(fig12, use_container_width=True)
            else:
                st.info("Sem marcações de sala exclusiva/dividida para analisar.")
        else:
            st.info("Inclua colunas 'Sala Exclusiva' e/ou 'Sala Dividida'.")

    st.markdown("##### Tabela (Médico × CRM × Especialidade × PLANOS × Valor × Tipo de Sala × Turnos)")
    cols_show = [c for c in ["Médico","CRM","Especialidade","Planos","Valor Aluguel","Sala Exclusiva","Sala Dividida","Turnos Utilizados"] if c in med_enriched.columns]
    st.dataframe(med_enriched[cols_show].sort_values(["Planos","Especialidade","Valor Aluguel","Médico"], na_position="last"), use_container_width=True)

# ---------- Detalhamento ----------
st.subheader("📋 Agenda Detalhada (Tabela)")
st.dataframe(
    fdf.sort_values(["Sala","Dia","Turno"]).reset_index(drop=True)[["Sala","Dia","Turno","Médico"]],
    use_container_width=True
)
csv = fdf.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Baixar dados filtrados (CSV)", data=csv, file_name="agenda_filtrada.csv", mime="text/csv")
