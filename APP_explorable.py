# APP_explorable.py
# Streamlit dashboard – KPI e-commerce Gedimat (explorable UI/UX)
# Refactor of the "TV slideshow" version into drill-down menus + deeper analysis panels.
#
# Key changes vs TV mode:
# - Sidebar navigation (menus)
# - Drill-down pages for Stores and Products (detail views)
# - Expanders for "deeper analysis" (UX hypotheses, checks, action lists)
# - No auto-refresh loop / no global overflow locking (standard dashboard ergonomics)

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------
# App configuration
# -----------------------------
st.set_page_config(
    page_title="Gedimat – KPI E-commerce",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      /* Slightly tighter padding than Streamlit default */
      .block-container {padding-top: 1.0rem; padding-bottom: 1.0rem; padding-left: 1.2rem; padding-right: 1.2rem;}

      /* Card styling helpers */
      .kpi-card {border: 1px solid rgba(255,255,255,0.12); border-radius: 16px; padding: 14px 16px; height: 100%;}
      .kpi-label {font-size: 13px; opacity: 0.78; margin-bottom: 6px;}
      .kpi-value {font-size: 34px; font-weight: 850; line-height: 1.1;}
      .kpi-delta {font-size: 13px; opacity: 0.80; margin-top: 6px;}
      .page-title {font-size: 34px; font-weight: 900; line-height: 1.1; margin: 0 0 2px 0;}
      .page-subtitle {font-size: 14px; opacity: 0.75; margin: 0 0 10px 0;}
      .chip {display:inline-block; padding: 5px 9px; border-radius: 999px; margin-right: 8px; font-size: 12px; opacity: 0.92;}
      .chip-ok {background: rgba(0, 200, 0, 0.12);}
      .chip-warn {background: rgba(255, 165, 0, 0.14);}
      .chip-bad {background: rgba(255, 0, 0, 0.12);}
      .small {font-size: 12px; opacity: 0.75;}
      .section-title {font-size: 18px; font-weight: 800; margin-top: 8px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Files & folders (repo layout)
# -----------------------------
ROOT = Path(__file__).resolve().parent
CONST_DIR = ROOT / "Constantes"
VAR_DIR = ROOT / "Variables"

FILES = {
    "visitors_overview": VAR_DIR / "Visitors_Overview_-_Gedimat.csv",
    "visits_store_ecom": VAR_DIR / "Visites_par_magasin_e-commerce_-_Gedimat.csv",
    "visits_store_adherents": VAR_DIR / "Visites_Magasins__Adhérents_-_Gedimat (1).csv",
    "magasin_ecommerce": CONST_DIR / "Magasin_ecommerce_-_Gedimat.csv",
    "pages_famille": CONST_DIR / "Pages_famille_-_Gedimat.csv",
    "stock_consult": VAR_DIR / "Disponibilité_-_Consultations_stock_par_magasin_-_Gedimat.csv",
    "clients_connexion": VAR_DIR / "Connexion_de_clients_-_Gedimat.csv",
    "top_products": VAR_DIR / "Produits_les_plus_visités_-_Gedimat.csv",
    "commandes": VAR_DIR / "commandes.xls",
    "magasins": CONST_DIR / "magasins.xls",
}


# -----------------------------
# Helpers: robust loading
# -----------------------------
@st.cache_data(ttl=180)
def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin-1")
    except Exception:
        try:
            return pd.read_csv(path, sep=";")
        except Exception:
            return pd.DataFrame()


@st.cache_data(ttl=180)
def load_excel_any(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception:
        pass
    try:
        return pd.read_excel(path, engine="xlrd")  # .xls
    except Exception:
        return pd.DataFrame()


def first_existing(paths: Iterable[Path]) -> Optional[Path]:
    for p in paths:
        if p.exists():
            return p
    return None


def safe_num(x, digits=0, suffix="") -> str:
    if x is None:
        return "—"
    try:
        if pd.isna(x):
            return "—"
    except Exception:
        pass
    try:
        x = float(x)
    except Exception:
        return str(x)
    if digits == 0:
        s = f"{x:,.0f}".replace(",", " ")
    else:
        s = f"{x:,.{digits}f}".replace(",", " ")
    return f"{s}{suffix}"


def pct(x, digits=1) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return safe_num(100 * x, digits=digits, suffix=" %")


def header(title: str, subtitle: str = ""):
    st.markdown(f"<div class='page-title'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='page-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def kpi_card(label: str, value: str, delta: Optional[str] = None):
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.markdown(f"<div class='kpi-label'>{label}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='kpi-value'>{value}</div>", unsafe_allow_html=True)
    if delta:
        st.markdown(f"<div class='kpi-delta'>{delta}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def chip(text: str, level: str = "ok"):
    cls = {"ok": "chip-ok", "warn": "chip-warn", "bad": "chip-bad"}.get(level, "chip-ok")
    st.markdown(f"<span class='chip {cls}'>{text}</span>", unsafe_allow_html=True)


def guess_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df.empty:
        return None
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    low = [(c.lower(), c) for c in df.columns]
    for cand in candidates:
        for cl, orig in low:
            if cand.lower() in cl:
                return orig
    return None


def to_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.date


# -----------------------------
# Load data (graceful)
# -----------------------------
visitors_overview = load_csv(FILES["visitors_overview"])
visits_store_ecom = load_csv(FILES["visits_store_ecom"])
visits_store_adherents = load_csv(FILES["visits_store_adherents"])
magasin_ecommerce = load_csv(FILES["magasin_ecommerce"])
pages_famille = load_csv(FILES["pages_famille"])
stock_consult = load_csv(FILES["stock_consult"])
clients_connexion = load_csv(FILES["clients_connexion"])
top_products = load_csv(FILES["top_products"])

commandes = load_excel_any(FILES["commandes"])
magasins_ref = load_excel_any(FILES["magasins"])


# -----------------------------
# KPI derivations (best-effort)
# -----------------------------
@dataclass
class GlobalKPIs:
    visitors: float = np.nan
    sessions: float = np.nan
    pageviews: float = np.nan
    unique_pages: float = np.nan
    bounce_rate: float = np.nan
    avg_session_sec: float = np.nan
    connected_sessions: float = np.nan
    connected_visitors: float = np.nan
    pro_sessions: float = np.nan
    pro_visitors: float = np.nan
    part_connected_visitors: float = np.nan


def compute_global_kpis() -> GlobalKPIs:
    out = GlobalKPIs()

    if not visitors_overview.empty:
        vcol = guess_col(visitors_overview, ["Visiteurs", "Visitors"])
        scol = guess_col(visitors_overview, ["Sessions"])
        pvcol = guess_col(visitors_overview, ["Pages consultées", "Pageviews", "Pages vues"])
        upcol = guess_col(visitors_overview, ["Pages uniques", "Unique pages"])
        brcol = guess_col(visitors_overview, ["Taux de rebond", "Bounce rate"])
        durcol = guess_col(visitors_overview, ["durée de la session moyenne", "Avg session", "session moyenne"])

        row = visitors_overview.iloc[-1]
        if vcol:
            out.visitors = float(row.get(vcol, np.nan))
        if scol:
            out.sessions = float(row.get(scol, np.nan))
        if pvcol:
            out.pageviews = float(row.get(pvcol, np.nan))
        if upcol:
            out.unique_pages = float(row.get(upcol, np.nan))
        if brcol:
            out.bounce_rate = float(row.get(brcol, np.nan))
        if durcol:
            out.avg_session_sec = float(row.get(durcol, np.nan))

    if not clients_connexion.empty:
        tcol = guess_col(clients_connexion, ["Type Client", "Type", "Client"])
        scol = guess_col(clients_connexion, ["Sessions"])
        vcol = guess_col(clients_connexion, ["Visiteurs", "Visitors"])
        if tcol and scol and vcol:
            tmp = clients_connexion.copy()
            tmp[tcol] = tmp[tcol].fillna("total")
            total = tmp[tmp[tcol].astype(str).str.lower().eq("total")].head(1)
            if not total.empty:
                out.connected_sessions = float(total.iloc[0][scol])
                out.connected_visitors = float(total.iloc[0][vcol])
            pro = tmp[tmp[tcol].astype(str).str.lower().str.contains("pro")]
            if not pro.empty:
                out.pro_sessions = float(pd.to_numeric(pro[scol], errors="coerce").fillna(0).sum())
                out.pro_visitors = float(pd.to_numeric(pro[vcol], errors="coerce").fillna(0).sum())

    if not np.isnan(out.connected_visitors) and not np.isnan(out.visitors) and out.visitors:
        out.part_connected_visitors = out.connected_visitors / out.visitors

    return out


def compute_store_table() -> pd.DataFrame:
    base = pd.DataFrame()
    if not magasin_ecommerce.empty:
        base = magasin_ecommerce.copy()
    elif not visits_store_ecom.empty:
        base = visits_store_ecom.copy()

    if base.empty:
        return base

    name_col = guess_col(base, ["Nom Magasin", "Magasin", "Store"])
    if not name_col:
        return pd.DataFrame()

    sess_col = guess_col(base, ["Sessions"])
    vis_col = guess_col(base, ["Visiteurs", "Visitors"])
    br_col = guess_col(base, ["Taux de rebond", "Bounce rate"])
    ord_col = guess_col(base, ["Commandes", "Orders"])

    out = pd.DataFrame({
        "Magasin": base[name_col].astype(str),
        "Sessions": base[sess_col] if sess_col else np.nan,
        "Visiteurs": base[vis_col] if vis_col else np.nan,
        "Taux de rebond": base[br_col] if br_col else np.nan,
        "Commandes": base[ord_col] if ord_col else np.nan,
    })

    out["Conv. (Cmd/Sess)"] = np.where(
        (pd.to_numeric(out["Sessions"], errors="coerce") > 0) & pd.notna(out["Commandes"]),
        pd.to_numeric(out["Commandes"], errors="coerce") / pd.to_numeric(out["Sessions"], errors="coerce"),
        np.nan,
    )

    for c in ["Sessions", "Visiteurs", "Taux de rebond", "Commandes", "Conv. (Cmd/Sess)"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    return out


def compute_stock_signals() -> Dict[str, float]:
    res = {"sessions_stock": np.nan, "pages_stock": np.nan, "events_stock": np.nan}
    if stock_consult.empty:
        return res

    sess_col = guess_col(stock_consult, ["Sessions"])
    pv_col = guess_col(stock_consult, ["Pages consultées", "Pageviews", "Pages vues"])
    ev_col = guess_col(stock_consult, ["Événements personnalisés", "Events"])

    if sess_col:
        res["sessions_stock"] = float(pd.to_numeric(stock_consult[sess_col], errors="coerce").fillna(0).sum())
    if pv_col:
        res["pages_stock"] = float(pd.to_numeric(stock_consult[pv_col], errors="coerce").fillna(0).sum())
    if ev_col:
        res["events_stock"] = float(pd.to_numeric(stock_consult[ev_col], errors="coerce").fillna(0).sum())

    return res


def compute_alerts(store_tbl: pd.DataFrame, g: GlobalKPIs) -> List[Tuple[str, str]]:
    alerts: List[Tuple[str, str]] = []

    if not np.isnan(g.bounce_rate) and g.bounce_rate > 0.65:
        alerts.append(("warn", f"Taux de rebond élevé: {pct(g.bounce_rate, 1)}"))

    if not store_tbl.empty and "Commandes" in store_tbl.columns:
        zero_orders = store_tbl[pd.to_numeric(store_tbl["Commandes"], errors="coerce").fillna(0) == 0]
        if len(zero_orders) >= 10:
            alerts.append(("bad", f"{len(zero_orders)} magasins à 0 commande (période du fichier)."))
        elif 1 <= len(zero_orders) < 10:
            alerts.append(("warn", f"{len(zero_orders)} magasins à 0 commande (période du fichier)."))

    if not np.isnan(g.part_connected_visitors):
        if g.part_connected_visitors < 0.20:
            alerts.append(("warn", f"Part visiteurs connectés faible: {pct(g.part_connected_visitors, 1)}"))

    if not commandes.empty:
        dcol = guess_col(commandes, ["date", "Date", "Date de commande", "Commande date", "Created"])
        store_col = guess_col(commandes, ["Nom Magasin", "Magasin", "Store"])
        id_col = guess_col(commandes, ["Commande", "Order", "Order ID", "N° commande", "Numero"])
        if dcol and store_col:
            tmp = commandes.copy()
            tmp[dcol] = to_date(tmp[dcol])
            tmp = tmp[pd.notna(tmp[dcol])]
            if not tmp.empty:
                last_date = max(tmp[dcol])
                cutoff = pd.Timestamp(last_date) - pd.Timedelta(days=30)
                tmp["_d"] = pd.to_datetime(tmp[dcol])
                recent = tmp[tmp["_d"] >= cutoff]
                all_stores = set(tmp[store_col].astype(str).unique())
                recent_stores = set(recent[store_col].astype(str).unique())
                dormant = sorted(list(all_stores - recent_stores))
                if len(dormant) > 0:
                    level = "bad" if len(dormant) >= 10 else "warn"
                    alerts.append((level, f"{len(dormant)} magasins sans commande depuis 30 jours (base dernière date: {last_date})."))
                if id_col:
                    tmp_day = tmp[tmp[dcol] == last_date]
                    n_orders = tmp_day[id_col].nunique() if not tmp_day.empty else 0
                    if n_orders == 0:
                        alerts.append(("bad", f"Aucune commande sur la dernière date disponible ({last_date})."))
        else:
            alerts.append(("warn", "commandes.xls chargé mais colonnes date/magasin non reconnues → alertes avancées limitées."))

    if not alerts:
        alerts.append(("ok", "Aucun signal critique détecté avec les données disponibles."))
    return alerts


GLOBAL = compute_global_kpis()
STORE_TBL = compute_store_table()
STOCK_SIG = compute_stock_signals()
ALERTS = compute_alerts(STORE_TBL, GLOBAL)


# -----------------------------
# Sidebar: navigation + state
# -----------------------------
st.sidebar.markdown("## Navigation")
page = st.sidebar.radio(
    "Aller à",
    [
        "Vue d'ensemble",
        "Trafic & audiences",
        "Magasins",
        "Produits",
        "Stock",
        "Tunnel",
        "Alertes & actions",
        "Données (debug)",
    ],
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.markdown("### Contexte")
st.sidebar.markdown(
    f"<div class='small'>Fichiers trouvés : <b>{sum(1 for p in FILES.values() if p.exists())}</b> / {len(FILES)}</div>",
    unsafe_allow_html=True,
)

with st.sidebar.expander("📦 Voir les fichiers attendus"):
    for k, p in FILES.items():
        ok = "✅" if p.exists() else "⚠️"
        st.write(f"{ok} {k}: {p.name}")


# Drill-down selections
if "selected_store" not in st.session_state:
    st.session_state.selected_store = None
if "selected_sku" not in st.session_state:
    st.session_state.selected_sku = None
if "selected_product_name" not in st.session_state:
    st.session_state.selected_product_name = None


# -----------------------------
# Shared drill-down builders
# -----------------------------
def store_detail_block(store_name: str):
    header(f"Magasin : {store_name}", "Détail & diagnostic (drill-down)")

    row = STORE_TBL[STORE_TBL["Magasin"].astype(str) == str(store_name)]
    if row.empty:
        st.info("Aucune ligne magasin correspondante dans les exports disponibles.")
        return
    r = row.iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Sessions", safe_num(r.get("Sessions", np.nan)))
    with c2:
        kpi_card("Visiteurs", safe_num(r.get("Visiteurs", np.nan)))
    with c3:
        kpi_card("Commandes", safe_num(r.get("Commandes", np.nan)))
    with c4:
        kpi_card("Conversion", pct(r.get("Conv. (Cmd/Sess)", np.nan), 2), f"Rebond: {pct(r.get('Taux de rebond', np.nan), 1)}")

    st.markdown("<div class='section-title'>Pistes UX / process</div>", unsafe_allow_html=True)


    # Stock consultations for that store (if available)
    if not stock_consult.empty:
        name_col = guess_col(stock_consult, ["Nom Magasin", "Magasin"])
        sess_col = guess_col(stock_consult, ["Sessions"])
        disp_col = guess_col(stock_consult, ["Disponibilite", "Disponibilité"])
        if name_col and sess_col:
            tmp = stock_consult.copy()
            tmp = tmp[tmp[name_col].astype(str) == str(store_name)]
            if not tmp.empty:
                tmp[sess_col] = pd.to_numeric(tmp[sess_col], errors="coerce").fillna(0)
                st.markdown("<div class='section-title'>Consultations stock</div>", unsafe_allow_html=True)
                if disp_col and disp_col in tmp.columns:
                    agg = tmp.groupby(disp_col, as_index=False)[sess_col].sum().sort_values(sess_col, ascending=False)
                    fig = px.bar(agg, x=disp_col, y=sess_col, title="Répartition des consultations par disponibilité")
                    fig.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Donnée stock présente mais dimension 'Disponibilité' non exploitable pour ce magasin.")

    # Orders recency for that store (if available)
    if not commandes.empty:
        dcol = guess_col(commandes, ["date", "Date", "Date de commande", "Commande date", "Created"])
        store_col = guess_col(commandes, ["Nom Magasin", "Magasin", "Store"])
        id_col = guess_col(commandes, ["Commande", "Order", "Order ID", "N° commande", "Numero"])
        if dcol and store_col and id_col:
            tmp = commandes.copy()
            tmp[dcol] = pd.to_datetime(tmp[dcol], errors="coerce")
            tmp = tmp[pd.notna(tmp[dcol])]
            tmp = tmp[tmp[store_col].astype(str) == str(store_name)]
            if not tmp.empty:
                daily = tmp.groupby(tmp[dcol].dt.date)[id_col].nunique().reset_index().rename(columns={id_col: "Commandes"})
                daily = daily.sort_values(dcol)
                fig = px.line(daily, x=dcol, y="Commandes", title="Commandes (par jour) – selon commandes.xls")
                fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)


def product_detail_block(product_name: str, sku: Optional[str] = None):
    header("Produit : " + (product_name[:80] if product_name else "—"), "Détail & diagnostic (drill-down)")

    # base demand row from top_products
    pcol = guess_col(top_products, ["Nom des produits", "Nom du produit", "Produit", "Product", "Libellé", "Désignation"])
    vcol = guess_col(top_products, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
    skcol = guess_col(top_products, ["SKU", "SKU de produit", "Référence", "Reference", "Ref", "Code produit"])

    if not vcol:
        numeric_cols = []
        for c in top_products.columns:
            if c in (pcol, skcol):
                continue
            s = pd.to_numeric(top_products[c], errors="coerce")
            if s.notna().sum() > 0:
                numeric_cols.append(c)
        vcol = numeric_cols[0] if numeric_cols else None

    if pcol and vcol:
        tmp = top_products.copy()
        tmp[vcol] = pd.to_numeric(tmp[vcol], errors="coerce").fillna(0)
        filt = tmp[tmp[pcol].astype(str) == str(product_name)]
        if not filt.empty:
            r = filt.iloc[0]
            demand = float(r.get(vcol, 0))
            sku_found = str(r.get(skcol, "")).strip() if skcol else ""
        else:
            demand = np.nan
            sku_found = ""
    else:
        demand = np.nan
        sku_found = ""

    sku_effective = (sku or sku_found or "").strip()

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Demandes (vues/sessions)", safe_num(demand))
    with c2:
        kpi_card("SKU", sku_effective if sku_effective else "—")
    with c3:
        # Orders for this SKU (if possible)
        orders = np.nan
        ca = np.nan
        if not commandes.empty and sku_effective:
            c_sku = guess_col(commandes, ["SKU", "sku", "Produit", "product", "Référence", "reference", "Code produit"])
            c_id = guess_col(commandes, ["Commande", "Order", "Order ID", "N° commande", "Numero", "Numéro"])
            c_amt = guess_col(commandes, ["CA", "Montant", "Total", "revenue", "Chiffre", "Net", "Amount"])
            if c_sku and c_id:
                sales = commandes.copy()
                sales[c_sku] = sales[c_sku].astype(str).str.strip()
                sub = sales[sales[c_sku] == sku_effective]
                orders = float(sub[c_id].nunique()) if not sub.empty else 0.0
                if c_amt:
                    ca = float(pd.to_numeric(sub[c_amt], errors="coerce").fillna(0).sum()) if not sub.empty else 0.0

        if not np.isnan(orders):
            delta = f"CA: {safe_num(ca, 0, ' €')}" if not np.isnan(ca) else None
            kpi_card("Commandes (commandes.xls)", safe_num(orders), delta)
        else:
            kpi_card("Commandes (commandes.xls)", "—", "Jointure SKU indisponible")

    st.markdown("<div class='section-title'>Analyse UX / e-commerce</div>", unsafe_allow_html=True)



    if not commandes.empty and sku_effective:
        # Basic time series for this SKU (if dates exist)
        dcol = guess_col(commandes, ["date", "Date", "Date de commande", "Commande date", "Created"])
        c_sku = guess_col(commandes, ["SKU", "sku", "Produit", "product", "Référence", "reference", "Code produit"])
        c_id = guess_col(commandes, ["Commande", "Order", "Order ID", "N° commande", "Numero", "Numéro"])
        if dcol and c_sku and c_id:
            sales = commandes.copy()
            sales[dcol] = pd.to_datetime(sales[dcol], errors="coerce")
            sales = sales[pd.notna(sales[dcol])]
            sales[c_sku] = sales[c_sku].astype(str).str.strip()
            sub = sales[sales[c_sku] == sku_effective]
            if not sub.empty:
                daily = sub.groupby(sub[dcol].dt.date)[c_id].nunique().reset_index().rename(columns={c_id: "Commandes"})
                fig = px.line(daily, x=dcol, y="Commandes", title="Commandes (par jour) – ce SKU (commandes.xls)")
                fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)


# -----------------------------
# Pages
# -----------------------------
def page_overview():
    header("Vue d'ensemble", "KPIs globaux + signaux rapides")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Visiteurs", safe_num(GLOBAL.visitors))
    with c2:
        kpi_card("Sessions", safe_num(GLOBAL.sessions))
    with c3:
        kpi_card("Pages consultées", safe_num(GLOBAL.pageviews))
    with c4:
        kpi_card("Taux de rebond", pct(GLOBAL.bounce_rate, 1), f"Durée moy.: {safe_num(GLOBAL.avg_session_sec, 0)} s")

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    left, right = st.columns([2, 1])
    with left:
        if not STORE_TBL.empty:
            chip(f"{len(STORE_TBL)} magasins e-commerce", "ok")
            if STORE_TBL["Commandes"].notna().any():
                chip(f"Commandes total: {safe_num(STORE_TBL['Commandes'].sum())}", "ok")
            if STORE_TBL["Conv. (Cmd/Sess)"].notna().any():
                chip(f"Conv. moyenne: {pct(STORE_TBL['Conv. (Cmd/Sess)'].mean(), 2)}", "warn")
        else:
            chip("Données magasins e-commerce indisponibles", "warn")

    with right:
        lvl = "ok"
        if any(a[0] == "bad" for a in ALERTS):
            lvl = "bad"
        elif any(a[0] == "warn" for a in ALERTS):
            lvl = "warn"
        chip("STATUT: " + ("CRITIQUE" if lvl == "bad" else "À SURVEILLER" if lvl == "warn" else "OK"), lvl)

    st.markdown("<div class='section-title'>Lecture rapide</div>", unsafe_allow_html=True)
    colA, colB = st.columns([1.1, 1])
    with colA:
        if not pages_famille.empty:
            rcol = guess_col(pages_famille, ["rayon", "famille", "catégorie", "categorie"])
            scol = guess_col(pages_famille, ["Sessions"])
            if rcol and scol:
                tmp = pages_famille.copy()
                tmp[scol] = pd.to_numeric(tmp[scol], errors="coerce").fillna(0)
                top = tmp.groupby(rcol, as_index=False)[scol].sum().sort_values(scol, ascending=False).head(10)
                fig = px.bar(top, x=rcol, y=scol, title="Top 10 familles (sessions)")
                fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Pages familles: colonnes non reconnues.")
        else:
            st.info("Ajoute Pages_famille_-_Gedimat.csv pour afficher les familles les plus consultées.")

    with colB:
        if not top_products.empty:
            pcol = guess_col(top_products, ["Nom des produits", "Nom du produit", "Produit", "Product", "Libellé", "Désignation"])
            vcol = guess_col(top_products, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
            if not vcol:
                numeric_cols = []
                for c in top_products.columns:
                    if c == pcol:
                        continue
                    s = pd.to_numeric(top_products[c], errors="coerce")
                    if s.notna().sum() > 0:
                        numeric_cols.append(c)
                vcol = numeric_cols[0] if numeric_cols else None
            if pcol and vcol:
                tmp = top_products.copy()
                tmp[vcol] = pd.to_numeric(tmp[vcol], errors="coerce").fillna(0)
                top10 = tmp.sort_values(vcol, ascending=False).head(10)
                fig = px.bar(top10.sort_values(vcol, ascending=True), x=vcol, y=pcol, orientation="h", title="Top 10 produits consultés")
                fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Top produits: colonnes non reconnues.")
        else:
            st.info("Ajoute Produits_les_plus_visités_-_Gedimat.csv pour afficher les top produits.")


def page_traffic():
    header("Trafic & audiences", "Segmentation connecté / Pro + pistes UX")

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Visiteurs connectés", safe_num(GLOBAL.connected_visitors), f"Part: {pct(GLOBAL.part_connected_visitors, 1)}")
    with c2:
        kpi_card("Sessions connectées", safe_num(GLOBAL.connected_sessions), f"Sessions Pro: {safe_num(GLOBAL.pro_sessions)}")
    with c3:
        kpi_card("Visiteurs Pro", safe_num(GLOBAL.pro_visitors))



    if not pages_famille.empty:
        rcol = guess_col(pages_famille, ["rayon", "famille", "catégorie", "categorie"])
        scol = guess_col(pages_famille, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
        if rcol and scol:
            tmp = pages_famille.copy()
            tmp[scol] = pd.to_numeric(tmp[scol], errors="coerce").fillna(0)
            top = tmp.groupby(rcol, as_index=False)[scol].sum().sort_values(scol, ascending=False).head(12)
            fig = px.bar(top, x=rcol, y=scol, title="Familles les plus consultées")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)


def page_stores():
    header("Magasins", "Liste + drill-down par magasin")

    if STORE_TBL.empty:
        st.warning("Pas de données magasin (Magasin_ecommerce_-_Gedimat.csv / Visites_par_magasin_e-commerce_-_Gedimat.csv).")
        return

    # Search / sort controls
    left, right = st.columns([1.1, 1])
    with left:
        query = st.text_input("Rechercher un magasin", value="", placeholder="Ex: Lyon, Bordeaux, ...")
    with right:
        sort_by = st.selectbox("Trier par", ["Sessions", "Commandes", "Conv. (Cmd/Sess)", "Taux de rebond", "Visiteurs"], index=0)

    tbl = STORE_TBL.copy()
    if query.strip():
        q = query.strip().lower()
        tbl = tbl[tbl["Magasin"].astype(str).str.lower().str.contains(q)]

    if sort_by in tbl.columns:
        tbl = tbl.sort_values(sort_by, ascending=False, na_position="last")

    st.dataframe(
        tbl,
        use_container_width=True,
        height=360,
        hide_index=True,
    )

    st.markdown("<div class='section-title'>Drill-down</div>", unsafe_allow_html=True)
    store_list = sorted(STORE_TBL["Magasin"].astype(str).unique().tolist())
    default_store = st.session_state.selected_store if st.session_state.selected_store in store_list else (store_list[0] if store_list else None)
    chosen = st.selectbox("Choisir un magasin", store_list, index=store_list.index(default_store) if default_store in store_list else 0)
    st.session_state.selected_store = chosen

    store_detail_block(chosen)


def page_products():
    header("Produits", "Liste + drill-down produit (vues vs commandes si possible)")

    if top_products.empty:
        st.warning("Fichier Produits_les_plus_visités_-_Gedimat.csv manquant.")
        return

    pcol = guess_col(top_products, ["Nom des produits", "Nom du produit", "Produit", "Product", "Libellé", "Désignation"])
    vcol = guess_col(top_products, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
    skcol = guess_col(top_products, ["SKU", "SKU de produit", "Référence", "Reference", "Ref", "Code produit"])

    if not vcol:
        numeric_cols = []
        for c in top_products.columns:
            if c in (pcol, skcol):
                continue
            s = pd.to_numeric(top_products[c], errors="coerce")
            if s.notna().sum() > 0:
                numeric_cols.append(c)
        vcol = numeric_cols[0] if numeric_cols else None

    if not pcol or not vcol:
        st.error("Colonnes produit / volume non reconnues dans l'export.")
        st.write("Colonnes disponibles:", list(top_products.columns))
        return

    tmp = top_products.copy()
    tmp[vcol] = pd.to_numeric(tmp[vcol], errors="coerce").fillna(0)

    # Controls
    c1, c2, c3 = st.columns([1.2, 1, 1])
    with c1:
        search = st.text_input("Rechercher un produit", value="", placeholder="Ex: perceuse, cheville, ...")
    with c2:
        topn = st.selectbox("Top N", [10, 20, 50, 100], index=1)
    with c3:
        sort_dir = st.selectbox("Tri", ["Desc", "Asc"], index=0)

    view = tmp
    if search.strip():
        q = search.strip().lower()
        view = view[view[pcol].astype(str).str.lower().str.contains(q)]

    view = view.sort_values(vcol, ascending=(sort_dir == "Asc")).head(int(topn))

    # Show table
    show_cols = [c for c in [pcol, skcol, vcol] if c and c in view.columns]
    st.dataframe(view[show_cols], use_container_width=True, height=360, hide_index=True)

    st.markdown("<div class='section-title'>Drill-down</div>", unsafe_allow_html=True)

    # pick product from full list (not only filtered) to keep selection stable
    products_list = top_products[pcol].astype(str).unique().tolist()
    products_list = [p for p in products_list if p and p.strip()]
    products_list = sorted(products_list)

    # default selection
    default_name = st.session_state.selected_product_name if st.session_state.selected_product_name in products_list else products_list[0]
    chosen_name = st.selectbox("Choisir un produit", products_list, index=products_list.index(default_name))
    st.session_state.selected_product_name = chosen_name

    sku_default = ""
    if skcol:
        sku_row = top_products[top_products[pcol].astype(str) == str(chosen_name)]
        if not sku_row.empty:
            sku_default = str(sku_row.iloc[0].get(skcol, "")).strip()

    chosen_sku = st.text_input("SKU (pour joindre commandes.xls si nécessaire)", value=sku_default)
    st.session_state.selected_sku = chosen_sku.strip() if chosen_sku else None

    product_detail_block(chosen_name, sku=st.session_state.selected_sku)


def page_stock():
    header("Stock", "Consultations stock + disponibilité (si dimension dispo présente)")

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Sessions (consultations stock)", safe_num(STOCK_SIG["sessions_stock"]))
    with c2:
        kpi_card("Pages consultées (stock)", safe_num(STOCK_SIG["pages_stock"]))
    with c3:
        kpi_card("Événements stock", safe_num(STOCK_SIG["events_stock"]))

    if stock_consult.empty:
        st.warning("Fichier stock consulté manquant.")
        return

    name_col = guess_col(stock_consult, ["Nom Magasin", "Magasin"])
    sess_col = guess_col(stock_consult, ["Sessions"])
    disp_col = guess_col(stock_consult, ["Disponibilite", "Disponibilité"])

    st.markdown("<div class='section-title'>Exploration</div>", unsafe_allow_html=True)

    if name_col and sess_col:
        tmp = stock_consult.copy()
        tmp = tmp[pd.notna(tmp[name_col])]
        tmp[sess_col] = pd.to_numeric(tmp[sess_col], errors="coerce").fillna(0)

        left, right = st.columns([1.1, 1])
        with left:
            top = tmp.groupby(name_col, as_index=False)[sess_col].sum().sort_values(sess_col, ascending=False).head(15)
            fig = px.bar(top, x=name_col, y=sess_col, title="Magasins : consultations stock (sessions)")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)

        with right:
            if disp_col and disp_col in tmp.columns:
                agg = tmp.groupby(disp_col, as_index=False)[sess_col].sum().sort_values(sess_col, ascending=False)
                fig = px.pie(agg, names=disp_col, values=sess_col, title="Répartition par disponibilité (sessions)")
                fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10), showlegend=True)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Colonne 'Disponibilité' non exploitable dans ce fichier.")



def page_funnel():
    header("Tunnel", "Lecture simplifiée + actions d'investigation")

    visitors = GLOBAL.visitors
    connected = GLOBAL.connected_visitors

    total_orders = np.nan
    if not STORE_TBL.empty and STORE_TBL["Commandes"].notna().any():
        total_orders = float(STORE_TBL["Commandes"].sum())

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Visiteurs", safe_num(visitors))
    with c2:
        kpi_card("Visiteurs connectés", safe_num(connected), f"Part: {pct(GLOBAL.part_connected_visitors, 1)}")
    with c3:
        kpi_card("Commandes", safe_num(total_orders), "Source: Magasin_ecommerce" if not np.isnan(total_orders) else "—")

    steps = ["Visiteurs", "Connectés", "Commandes"]
    vals = [
        float(visitors) if not np.isnan(visitors) else 0.0,
        float(connected) if not np.isnan(connected) else 0.0,
        float(total_orders) if not np.isnan(total_orders) else 0.0,
    ]
    fig = go.Figure(go.Funnel(y=steps, x=vals))
    fig.update_layout(height=460, margin=dict(l=10, r=10, t=20, b=10))
    st.plotly_chart(fig, use_container_width=True)




def page_alerts():
    header("Alertes & actions", "Priorisation + checklists")

    # Render alerts
    for level, text in ALERTS[:12]:
        icon = {"ok": "🟢", "warn": "🟠", "bad": "🔴"}.get(level, "🟢")
        st.markdown(f"### {icon} {text}")

    st.markdown("<div class='section-title'>Actions recommandées</div>", unsafe_allow_html=True)
    st.markdown(
        """
- Vérifier les **magasins à 0 commande** : stock, livraison, paiement, horaires, params click&collect.
- Contrôler les **produits très consultés** : ruptures, prix, infos produit, frais/délais.
- Si rebond élevé : performance, erreurs, tracking, pages d’atterrissage.
- Mettre en place un **diagnostic funnel** (micro-conversions) pour localiser la chute.
        """.strip()
    )



def page_debug():
    header("Données (debug)", "Voir les tables brutes pour vérifier colonnes / jointures")

    st.write("visitors_overview", visitors_overview.head(10))
    st.write("clients_connexion", clients_connexion.head(10))
    st.write("store table (computed)", STORE_TBL.head(20))
    st.write("pages_famille", pages_famille.head(10))
    st.write("top_products", top_products.head(10))
    st.write("stock_consult", stock_consult.head(10))
    st.write("commandes", commandes.head(10))
    st.write("magasins_ref", magasins_ref.head(10))


# -----------------------------
# Router
# -----------------------------
if page == "Vue d'ensemble":
    page_overview()
elif page == "Trafic & audiences":
    page_traffic()
elif page == "Magasins":
    page_stores()
elif page == "Produits":
    page_products()
elif page == "Stock":
    page_stock()
elif page == "Tunnel":
    page_funnel()
elif page == "Alertes & actions":
    page_alerts()
else:
    page_debug()
