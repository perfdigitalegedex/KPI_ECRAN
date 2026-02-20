# APP.py
# Streamlit "TV dashboard" – KPI e-commerce Gedimat
# - 7 slides
# - 2 minutes total loop (120s)
# - no user interaction
# - attempts to avoid page scrolling (fit to screen)

from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, Dict, List

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
except Exception:  # pragma: no cover
    st_autorefresh = None


# -----------------------------
# App configuration (TV mode)
# -----------------------------
st.set_page_config(
    page_title="Gedimat – KPI E-commerce (TV)",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hard "TV" CSS: hide Streamlit chrome, enlarge typography, and avoid global scrolling.
st.markdown(
    """
    <style>
      /* Hide Streamlit default UI */
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}
      header {visibility: hidden;}
      [data-testid="stSidebar"] {display:none;}

      /* Tighten page padding */
      .block-container {padding-top: 0.8rem; padding-bottom: 0.8rem; padding-left: 1.2rem; padding-right: 1.2rem;}

      /* Avoid overall page scroll where possible */
      html, body {height: 100%; overflow: hidden;}
      section.main {height: 100vh; overflow: hidden;}

      /* Metrics style */
      .tv-kpi-label {font-size: 18px; opacity: 0.78; margin-bottom: 4px;}
      .tv-kpi-value {font-size: 56px; font-weight: 800; line-height: 1.05;}
      .tv-kpi-delta {font-size: 18px; opacity: 0.85; margin-top: 6px;}
      .tv-title {font-size: 52px; font-weight: 900; line-height: 1.0; margin: 0;}
      .tv-subtitle {font-size: 18px; opacity: 0.75; margin-top: 10px;}
      .tv-chip {display:inline-block; padding: 6px 10px; border-radius: 999px; margin-right: 8px; font-size: 14px; opacity: 0.9;}
      .tv-chip-ok {background: rgba(0, 200, 0, 0.12);}
      .tv-chip-warn {background: rgba(255, 165, 0, 0.14);}
      .tv-chip-bad {background: rgba(255, 0, 0, 0.12);}
      .tv-card {border: 1px solid rgba(255,255,255,0.12); border-radius: 16px; padding: 14px 16px; height: 100%;}
      .tv-small {font-size: 14px; opacity: 0.75;}
      .tv-alert {font-size: 22px; font-weight: 750; margin: 8px 0;}
      .tv-divider {height: 10px;}
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

# Expected filenames (you can rename in your repo; then update here)
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
@st.cache_data(ttl=180)  # refresh every 3 minutes
def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin-1")
    except Exception:
        # try separator sniff (rare)
        try:
            return pd.read_csv(path, sep=";")
        except Exception:
            return pd.DataFrame()


@st.cache_data(ttl=180)
def load_excel_any(path: Path) -> pd.DataFrame:
    """
    Reads xlsx/xls if possible.
    Note: .xls requires xlrd in requirements. If missing, we fall back gracefully.
    """
    if not path.exists():
        return pd.DataFrame()

    # Try default engine(s)
    try:
        return pd.read_excel(path)
    except Exception:
        pass

    # Try xlrd explicitly (xls)
    try:
        return pd.read_excel(path, engine="xlrd")
    except Exception:
        return pd.DataFrame()


def first_existing(paths: Iterable[Path]) -> Optional[Path]:
    for p in paths:
        if p.exists():
            return p
    return None


def safe_num(x, digits=0, suffix=""):
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


def pct(x, digits=1):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return safe_num(100 * x, digits=digits, suffix=" %")


def tv_title(title: str, subtitle: str = ""):
    st.markdown(f"<div class='tv-title'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='tv-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def kpi_tile(label: str, value: str, delta: Optional[str] = None):
    st.markdown("<div class='tv-card'>", unsafe_allow_html=True)
    st.markdown(f"<div class='tv-kpi-label'>{label}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='tv-kpi-value'>{value}</div>", unsafe_allow_html=True)
    if delta:
        st.markdown(f"<div class='tv-kpi-delta'>{delta}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def chip(text: str, level: str = "ok"):
    cls = {"ok": "tv-chip-ok", "warn": "tv-chip-warn", "bad": "tv-chip-bad"}.get(level, "tv-chip-ok")
    st.markdown(f"<span class='tv-chip {cls}'>{text}</span>", unsafe_allow_html=True)


def guess_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df.empty:
        return None
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    # fuzzy contains
    low = [(c.lower(), c) for c in df.columns]
    for cand in candidates:
        for cl, orig in low:
            if cand.lower() in cl:
                return orig
    return None


def to_date(series: pd.Series) -> pd.Series:
    # tolerant parsing
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

# Excel (optional but recommended for advanced alerts)
commandes = load_excel_any(FILES["commandes"])
magasins_ref = load_excel_any(FILES["magasins"])


# -----------------------------
# Slide timing (2 min loop / 7 slides)
# -----------------------------
TOTAL_SECONDS = 120
N_SLIDES = 7
SLIDE_SECONDS = TOTAL_SECONDS / N_SLIDES  # ~17.142s

# Auto-refresh (keeps TV moving)
if st_autorefresh is not None:
    st_autorefresh(interval=int(SLIDE_SECONDS * 1000), key="tv_refresh")
else:
    st.info("Dépendance manquante: streamlit-autorefresh. Ajoute-la dans requirements.txt.")

# Stable indexing based on start time (prevents drift)
if "tv_t0" not in st.session_state:
    st.session_state.tv_t0 = time.time()

elapsed = time.time() - st.session_state.tv_t0
slide_index = int((elapsed % TOTAL_SECONDS) // SLIDE_SECONDS)  # 0..6


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
        if vcol: out.visitors = float(row.get(vcol, np.nan))
        if scol: out.sessions = float(row.get(scol, np.nan))
        if pvcol: out.pageviews = float(row.get(pvcol, np.nan))
        if upcol: out.unique_pages = float(row.get(upcol, np.nan))
        if brcol: out.bounce_rate = float(row.get(brcol, np.nan))
        if durcol: out.avg_session_sec = float(row.get(durcol, np.nan))

    if not clients_connexion.empty:
        # Expected columns: Type Client, Sessions, Visiteurs
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
                out.pro_sessions = float(pro[scol].sum())
                out.pro_visitors = float(pro[vcol].sum())

    if not np.isnan(out.connected_visitors) and not np.isnan(out.visitors) and out.visitors:
        out.part_connected_visitors = out.connected_visitors / out.visitors

    return out


def compute_store_table() -> pd.DataFrame:
    """
    Builds a store performance table from available sources:
    - Magasin_ecommerce_-_Gedimat.csv (sessions/visitors/bounce/orders)
    - Visites_par_magasin_e-commerce_-_Gedimat.csv (sessions/visitors)
    """
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

    # Ensure columns exist
    def ensure(colname: str):
        if colname not in base.columns:
            base[colname] = np.nan

    # Map likely columns
    sess_col = guess_col(base, ["Sessions"])
    vis_col = guess_col(base, ["Visiteurs", "Visitors"])
    br_col = guess_col(base, ["Taux de rebond", "Bounce rate"])
    ord_col = guess_col(base, ["Commandes", "Orders"])

    # Standardize names
    out = pd.DataFrame({
        "Magasin": base[name_col].astype(str),
        "Sessions": base[sess_col] if sess_col else np.nan,
        "Visiteurs": base[vis_col] if vis_col else np.nan,
        "Taux de rebond": base[br_col] if br_col else np.nan,
        "Commandes": base[ord_col] if ord_col else np.nan,
    })

    # Add conversion if possible
    out["Conv. (Cmd/Sess)"] = np.where(
        (out["Sessions"].astype(float) > 0) & pd.notna(out["Commandes"]),
        out["Commandes"].astype(float) / out["Sessions"].astype(float),
        np.nan,
    )

    # clean numeric
    for c in ["Sessions", "Visiteurs", "Taux de rebond", "Commandes", "Conv. (Cmd/Sess)"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    return out


def compute_stock_signals() -> Dict[str, float]:
    """
    From stock consultations: best-effort signals about stock-related interest.
    We don't have explicit 'in stock' ratio here; we use sessions/pages/events.
    """
    res = {"sessions_stock": np.nan, "pages_stock": np.nan, "events_stock": np.nan}
    if stock_consult.empty:
        return res

    sess_col = guess_col(stock_consult, ["Sessions"])
    pv_col = guess_col(stock_consult, ["Pages consultées", "Pageviews", "Pages vues"])
    ev_col = guess_col(stock_consult, ["Événements personnalisés", "Events"])

    if sess_col: res["sessions_stock"] = float(pd.to_numeric(stock_consult[sess_col], errors="coerce").sum())
    if pv_col: res["pages_stock"] = float(pd.to_numeric(stock_consult[pv_col], errors="coerce").sum())
    if ev_col: res["events_stock"] = float(pd.to_numeric(stock_consult[ev_col], errors="coerce").sum())

    return res


def compute_alerts(store_tbl: pd.DataFrame, g: GlobalKPIs) -> List[Tuple[str, str]]:
    """
    Returns list of (level, text) where level in {"bad","warn","ok"}.
    Alerts are based on whatever is available.
    """
    alerts: List[Tuple[str, str]] = []

    # Traffic drop can't be computed with 1 row. Still flag if bounce is high.
    if not np.isnan(g.bounce_rate) and g.bounce_rate > 0.65:
        alerts.append(("warn", f"Taux de rebond élevé: {pct(g.bounce_rate, 1)}"))

    # No-orders stores
    if not store_tbl.empty and "Commandes" in store_tbl.columns:
        zero_orders = store_tbl[pd.to_numeric(store_tbl["Commandes"], errors="coerce").fillna(0) == 0]
        if len(zero_orders) >= 10:
            alerts.append(("bad", f"{len(zero_orders)} magasins à 0 commande (période du fichier)."))
        elif 1 <= len(zero_orders) < 10:
            alerts.append(("warn", f"{len(zero_orders)} magasins à 0 commande (période du fichier)."))

    # Connected visitors share
    if not np.isnan(g.part_connected_visitors):
        if g.part_connected_visitors < 0.20:
            alerts.append(("warn", f"Part visiteurs connectés faible: {pct(g.part_connected_visitors, 1)}"))

    # Advanced alerts from commandes (if available & parsable)
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
                # stores with no order in last 30 days
                all_stores = set(tmp[store_col].astype(str).unique())
                recent_stores = set(recent[store_col].astype(str).unique())
                dormant = sorted(list(all_stores - recent_stores))
                if len(dormant) > 0:
                    level = "bad" if len(dormant) >= 10 else "warn"
                    alerts.append((level, f"{len(dormant)} magasins sans commande depuis 30 jours (base dernière date: {last_date})."))

                # if also have order id, check no orders yesterday etc (best-effort)
                if id_col:
                    tmp_day = tmp[tmp[dcol] == last_date]
                    n_orders = tmp_day[id_col].nunique() if not tmp_day.empty else 0
                    if n_orders == 0:
                        alerts.append(("bad", f"Aucune commande sur la dernière date disponible ({last_date})."))
        else:
            alerts.append(("warn", "commandes.xls chargé mais colonnes date/magasin non reconnues → alertes avancées désactivées."))

    if not alerts:
        alerts.append(("ok", "Aucun signal critique détecté avec les données disponibles."))
    return alerts


# -----------------------------
# Slides (7)
# -----------------------------
GLOBAL = compute_global_kpis()
STORE_TBL = compute_store_table()
STOCK_SIG = compute_stock_signals()
ALERTS = compute_alerts(STORE_TBL, GLOBAL)

NOW_STR = time.strftime("%d/%m/%Y %H:%M:%S")


def slide_1_health():
    tv_title("E-commerce – Santé globale", f"Mise à jour: {NOW_STR}  •  Slide 1/7")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_tile("Visiteurs", safe_num(GLOBAL.visitors), None)
    with c2:
        kpi_tile("Sessions", safe_num(GLOBAL.sessions), None)
    with c3:
        kpi_tile("Pages consultées", safe_num(GLOBAL.pageviews), None)
    with c4:
        kpi_tile("Taux de rebond", pct(GLOBAL.bounce_rate, 1), f"Durée moy.: {safe_num(GLOBAL.avg_session_sec, 0)} s")

    st.markdown("<div class='tv-divider'></div>", unsafe_allow_html=True)

    # Store summary (top line chips)
    left, right = st.columns([2, 1])
    with left:
        if not STORE_TBL.empty:
            n = len(STORE_TBL)
            chip(f"{n} magasins e-commerce", "ok")
            if "Commandes" in STORE_TBL.columns and STORE_TBL["Commandes"].notna().any():
                chip(f"Commandes total (période): {safe_num(STORE_TBL['Commandes'].sum())}", "ok")
            if "Conv. (Cmd/Sess)" in STORE_TBL.columns and STORE_TBL["Conv. (Cmd/Sess)"].notna().any():
                chip(f"Conv. moyenne: {pct(STORE_TBL['Conv. (Cmd/Sess)'].mean(), 2)}", "warn")
        else:
            chip("Données magasins e-commerce indisponibles", "warn")

    with right:
        # Quick status based on alerts
        lvl = "ok"
        if any(a[0] == "bad" for a in ALERTS):
            lvl = "bad"
        elif any(a[0] == "warn" for a in ALERTS):
            lvl = "warn"
        chip("STATUT: " + ("CRITIQUE" if lvl == "bad" else "À SURVEILLER" if lvl == "warn" else "OK"), lvl)

    # A compact chart if possible
    if not pages_famille.empty:
        rcol = guess_col(pages_famille, ["rayon", "famille", "catégorie", "categorie"])
        scol = guess_col(pages_famille, ["Sessions"])
        if rcol and scol:
            tmp = pages_famille.copy()
            tmp[scol] = pd.to_numeric(tmp[scol], errors="coerce")
            top = tmp.groupby(rcol, as_index=False)[scol].sum().sort_values(scol, ascending=False).head(10)
            fig = px.bar(top, x=rcol, y=scol, title="Top 10 familles (sessions)")
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.markdown("<div class='tv-small'>Astuce: ajoute Pages_famille_-_Gedimat.csv pour afficher les familles les plus consultées.</div>", unsafe_allow_html=True)


def slide_2_traffic_interest():
    tv_title("Trafic & intérêt client", f"Mise à jour: {NOW_STR}  •  Slide 2/7")

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_tile("Visiteurs connectés", safe_num(GLOBAL.connected_visitors), f"Part: {pct(GLOBAL.part_connected_visitors, 1)}")
    with c2:
        kpi_tile("Sessions connectées", safe_num(GLOBAL.connected_sessions), f"Pro: {safe_num(GLOBAL.pro_sessions)} sess.")
    with c3:
        kpi_tile("Visiteurs Pro", safe_num(GLOBAL.pro_visitors), None)

    st.markdown("<div class='tv-divider'></div>", unsafe_allow_html=True)

    left, right = st.columns([1.2, 1])

    with left:
        if not pages_famille.empty:
            rcol = guess_col(pages_famille, ["rayon", "famille", "catégorie", "categorie"])
            scol = guess_col(pages_famille, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
            if rcol and scol:
                tmp = pages_famille.copy()
                tmp[scol] = pd.to_numeric(tmp[scol], errors="coerce")
                top = tmp.groupby(rcol, as_index=False)[scol].sum().sort_values(scol, ascending=False).head(12)
                fig = px.bar(top, x=rcol, y=scol, title="Familles les plus consultées")
                fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.markdown(
                    "<div class='tv-card'><b>Pages familles</b><div class='tv-small'>Colonnes non reconnues pour afficher les familles.</div></div>",
                    unsafe_allow_html=True
                )
        else:
            st.markdown("<div class='tv-card'><b>Pages familles</b><div class='tv-small'>Fichier Pages_famille_-_Gedimat.csv manquant.</div></div>", unsafe_allow_html=True)

    with right:
        st.markdown("<div class='tv-card'><b>Top produits consultés</b>", unsafe_allow_html=True)

        if top_products.empty:
            st.markdown("<div class='tv-small'>Fichier Produits_les_plus_visités_-_Gedimat.csv manquant.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
            return

        pcol = guess_col(top_products, ["Nom des produits", "Nom du produit", "Produit", "Product", "Libellé", "Désignation"])
        scol = guess_col(top_products, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
        skcol = guess_col(top_products, ["SKU", "SKU de produit", "Référence", "Reference", "Ref", "Code produit"])

        if not scol:
            numeric_cols = []
            for c in top_products.columns:
                if c == pcol or c == skcol:
                    continue
                s = pd.to_numeric(top_products[c], errors="coerce")
                if s.notna().sum() > 0:
                    numeric_cols.append(c)
            scol = numeric_cols[0] if numeric_cols else None

        if not pcol or not scol:
            st.markdown("<div class='tv-small'>⚠️ Colonnes produit / volume non reconnues dans l'export.</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='tv-small'>Colonnes disponibles: {', '.join(list(top_products.columns)[:10])}...</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
            return

        tmp = top_products.copy()
        tmp[scol] = pd.to_numeric(tmp[scol], errors="coerce").fillna(0)
        show = tmp.sort_values(scol, ascending=False).head(10)

        for i, (_, row) in enumerate(show.iterrows(), start=1):
            name = str(row.get(pcol, "")).strip()
            vol = row.get(scol, np.nan)
            sku = str(row.get(skcol, "")).strip() if skcol else ""
            sku_txt = f" <span style='opacity:.6'>({sku})</span>" if sku else ""
            st.markdown(
                f"<div class='tv-small'>{i}. <b>{name[:60]}</b>{sku_txt} — <b>{safe_num(vol)}</b></div>",
                unsafe_allow_html=True
            )

        st.markdown("</div>", unsafe_allow_html=True)


def slide_3_store_difficulty():
    tv_title("Magasins – pilotage & difficultés", f"Mise à jour: {NOW_STR}  •  Slide 3/7")

    if STORE_TBL.empty:
        st.warning("Pas de données magasin (Magasin_ecommerce_-_Gedimat.csv / Visites_par_magasin_e-commerce_-_Gedimat.csv).")
        return

    has_orders = ("Commandes" in STORE_TBL.columns) and STORE_TBL["Commandes"].notna().any()
    metric = "Commandes" if has_orders else "Sessions"

    conv_mean = float(STORE_TBL["Conv. (Cmd/Sess)"].mean()) if ("Conv. (Cmd/Sess)" in STORE_TBL.columns and STORE_TBL["Conv. (Cmd/Sess)"].notna().any()) else np.nan

    tbl = STORE_TBL.copy()
    tbl["__metric"] = pd.to_numeric(tbl[metric], errors="coerce").fillna(0)

    def status_row(r):
        cmd = r.get("Commandes", np.nan)
        conv = r.get("Conv. (Cmd/Sess)", np.nan)
        sess = r.get("Sessions", np.nan)

        if has_orders and pd.notna(cmd) and float(cmd) == 0 and pd.to_numeric(sess, errors="coerce") > 0:
            return "🔴"
        if has_orders and pd.notna(cmd) and float(cmd) == 0:
            return "🟠"
        if pd.notna(conv_mean) and pd.notna(conv) and conv_mean > 0 and float(conv) < 0.6 * conv_mean:
            return "🟠"
        return "🟢"

    tbl["Statut"] = tbl.apply(status_row, axis=1)

    top = tbl.sort_values("__metric", ascending=False).head(12)
    watch = tbl.sort_values("__metric", ascending=True).head(12)

    left, right = st.columns([1.25, 1])

    with left:
        fig = px.bar(
            top.sort_values("__metric", ascending=True),
            x="__metric",
            y="Magasin",
            orientation="h",
            title=f"Top 12 magasins ({metric})",
        )
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=50, b=10))
        fig.update_xaxes(title=None)
        fig.update_yaxes(title=None)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("<div class='tv-card'><b>À surveiller (12)</b>", unsafe_allow_html=True)
        for _, r in watch.iterrows():
            st.markdown(
                "<div class='tv-small'>"
                f"{r['Statut']} <b>{str(r['Magasin'])[:40]}</b> "
                f"| {safe_num(r.get('Sessions', np.nan))} sess. "
                f"| {safe_num(r.get('Commandes', np.nan))} cmd "
                f"| Conv {pct(r.get('Conv. (Cmd/Sess)', np.nan), 2)}"
                "</div>",
                unsafe_allow_html=True
            )
        st.markdown("</div>", unsafe_allow_html=True)


def slide_4_products_frustration():
    tv_title("Produits – demande & blocages", f"Mise à jour: {NOW_STR}  •  Slide 4/7")

    left, right = st.columns([1.25, 1])

    if top_products.empty:
        with left:
            st.markdown("<div class='tv-card'><b>Demande (produits)</b><div class='tv-small'>Fichier top produits manquant.</div></div>", unsafe_allow_html=True)
        with right:
            st.markdown("<div class='tv-card'><b>Blocages</b><div class='tv-small'>Impossible d'analyser sans top produits.</div></div>", unsafe_allow_html=True)
        return

    pcol = guess_col(top_products, ["Nom des produits", "Nom du produit", "Produit", "Product", "Libellé", "Désignation"])
    vcol = guess_col(top_products, ["Sessions", "Consultations", "Pages consultées", "Vues", "Views"])
    skcol = guess_col(top_products, ["SKU", "SKU de produit", "Référence", "Reference", "Ref", "Code produit"])

    if not vcol:
        numeric_cols = []
        for c in top_products.columns:
            if c == pcol or c == skcol:
                continue
            s = pd.to_numeric(top_products[c], errors="coerce")
            if s.notna().sum() > 0:
                numeric_cols.append(c)
        vcol = numeric_cols[0] if numeric_cols else None

    if not pcol or not vcol:
        with left:
            st.markdown("<div class='tv-card'><b>Demande (produits)</b><div class='tv-small'>Colonnes produit/volume non reconnues.</div></div>", unsafe_allow_html=True)
        with right:
            st.markdown("<div class='tv-card'><b>Blocages</b><div class='tv-small'>Impossible de calculer sans colonnes identifiées.</div></div>", unsafe_allow_html=True)
        return

    tmp = top_products.copy()
    tmp[vcol] = pd.to_numeric(tmp[vcol], errors="coerce").fillna(0)

    top10 = tmp.sort_values(vcol, ascending=False).head(10)

    with left:
        fig = px.bar(
            top10.sort_values(vcol, ascending=True),
            x=vcol, y=pcol,
            orientation="h",
            title="Demande: Top 10 produits consultés",
        )
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=50, b=10))
        fig.update_xaxes(title=None)
        fig.update_yaxes(title=None)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("<div class='tv-card'><b>Blocages probables (à vérifier)</b>", unsafe_allow_html=True)
        st.markdown("<div class='tv-small'>Très consulté ≠ vendu → vérifier stock, prix, panier, paiement, livraison.</div>", unsafe_allow_html=True)

        blockers = []

        if not commandes.empty:
            c_sku = guess_col(commandes, ["SKU", "sku", "Produit", "product", "Référence", "reference", "Code produit"])
            c_amt = guess_col(commandes, ["CA", "Montant", "Total", "revenue", "Chiffre", "Net", "Amount"])
            c_id = guess_col(commandes, ["Commande", "Order", "Order ID", "N° commande", "Numero", "Numéro"])

            if skcol and c_sku:
                sales = commandes.copy()
                sales[c_sku] = sales[c_sku].astype(str).str.strip()

                if c_amt:
                    sales[c_amt] = pd.to_numeric(sales[c_amt], errors="coerce").fillna(0)
                    agg = sales.groupby(c_sku, as_index=False)[c_amt].sum().rename(columns={c_amt: "CA"})
                else:
                    agg = sales.groupby(c_sku, as_index=False).size()
                    agg = agg.rename(columns={"size": "NbLignes"})
                    agg["CA"] = np.nan

                if c_id:
                    oc = sales.groupby(c_sku)[c_id].nunique().reset_index().rename(columns={c_id: "Commandes"})
                    agg = agg.merge(oc, on=c_sku, how="left")
                else:
                    agg["Commandes"] = np.nan

                demand = top10.copy()
                demand[skcol] = demand[skcol].astype(str).str.strip()
                merged = demand.merge(agg, left_on=skcol, right_on=c_sku, how="left")

                merged["Commandes"] = pd.to_numeric(merged["Commandes"], errors="coerce").fillna(0)
                suspects = merged.sort_values(vcol, ascending=False)
                suspects = suspects[suspects["Commandes"] == 0].head(6)

                if not suspects.empty:
                    blockers.append("🔴 Top produits consultés mais 0 commande (période commandes):")
                    for _, r in suspects.iterrows():
                        blockers.append(f"• {str(r[pcol])[:52]} ({safe_num(r[vcol])} vues)")
                else:
                    blockers.append("🟢 Pas de produit 'très vu / 0 commande' détecté sur le top 10.")
            else:
                blockers.append("🟠 commandes.xls lisible, mais jointure produit impossible (SKU manquant).")
        else:
            blockers.append("🟠 commandes.xls non exploitable → pas de comparaison vues/commandes.")

        if not stock_consult.empty:
            blockers.append("🟠 Stock: fichier présent → vérifier ruptures sur ces top produits.")

        for line in blockers[:8]:
            st.markdown(f"<div class='tv-small'>{line}</div>", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)


def slide_5_stock_availability():
    tv_title("Stock & disponibilité", f"Mise à jour: {NOW_STR}  •  Slide 5/7")

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_tile("Sessions (consultations stock)", safe_num(STOCK_SIG["sessions_stock"]), None)
    with c2:
        kpi_tile("Pages consultées (stock)", safe_num(STOCK_SIG["pages_stock"]), None)
    with c3:
        kpi_tile("Événements stock", safe_num(STOCK_SIG["events_stock"]), None)

    st.markdown("<div class='tv-divider'></div>", unsafe_allow_html=True)

    if stock_consult.empty:
        st.warning("Fichier stock consulté manquant.")
        return

    name_col = guess_col(stock_consult, ["Nom Magasin", "Magasin"])
    sess_col = guess_col(stock_consult, ["Sessions"])
    disp_col = guess_col(stock_consult, ["Disponibilite", "Disponibilité"])

    left, right = st.columns([1.1, 1])

    with left:
        if name_col and sess_col:
            tmp = stock_consult.copy()
            tmp[sess_col] = pd.to_numeric(tmp[sess_col], errors="coerce")
            # Keep only meaningful rows (some exports include a NaN header line)
            tmp = tmp[pd.notna(tmp[name_col])]
            top = tmp.groupby(name_col, as_index=False)[sess_col].sum().sort_values(sess_col, ascending=False).head(12)
            fig = px.bar(top, x=name_col, y=sess_col, title="Magasins: consultations stock (sessions)")
            fig.update_layout(height=440, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)

    with right:
        # If availability dimension exists, show breakdown
        if disp_col and sess_col:
            tmp = stock_consult.copy()
            tmp[sess_col] = pd.to_numeric(tmp[sess_col], errors="coerce")
            tmp = tmp[pd.notna(tmp[disp_col])]
            top = tmp.groupby(disp_col, as_index=False)[sess_col].sum().sort_values(sess_col, ascending=False).head(10)
            fig = px.pie(top, names=disp_col, values=sess_col, title="Répartition par disponibilité (sessions)")
            fig.update_layout(height=440, margin=dict(l=10, r=10, t=50, b=10), showlegend=True)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.markdown("<div class='tv-card'><b>Détail disponibilité</b><div class='tv-small'>Colonne 'Disponibilite' non exploitable dans ce fichier.</div></div>", unsafe_allow_html=True)


def slide_6_funnel_simple():
    tv_title("Tunnel simplifié", f"Mise à jour: {NOW_STR}  •  Slide 6/7")

    # Best-effort funnel: visitors -> connected visitors -> orders (from magasin_ecommerce if available)
    visitors = GLOBAL.visitors
    connected = GLOBAL.connected_visitors

    total_orders = np.nan
    if not STORE_TBL.empty and STORE_TBL["Commandes"].notna().any():
        total_orders = float(STORE_TBL["Commandes"].sum())

    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_tile("Visiteurs", safe_num(visitors), None)
    with c2:
        kpi_tile("Visiteurs connectés", safe_num(connected), f"Part: {pct(GLOBAL.part_connected_visitors, 1)}")
    with c3:
        kpi_tile("Commandes", safe_num(total_orders), "Source: Magasin_ecommerce" if not np.isnan(total_orders) else "—")

    st.markdown("<div class='tv-divider'></div>", unsafe_allow_html=True)

    # Funnel chart
    steps = ["Visiteurs", "Connectés", "Commandes"]
    vals = [
        float(visitors) if not np.isnan(visitors) else 0.0,
        float(connected) if not np.isnan(connected) else 0.0,
        float(total_orders) if not np.isnan(total_orders) else 0.0,
    ]
    fig = go.Figure(go.Funnel(y=steps, x=vals))
    fig.update_layout(height=480, margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown(
        "<div class='tv-small'>Lecture: si 'Connectés' baisse, problème d'auth / UX / trafic qualifié. "
        "Si 'Commandes' baisse à trafic stable, suspecter stock, prix, panier, paiement, logistique.</div>",
        unsafe_allow_html=True
    )


def slide_7_alerts():
    tv_title("Alertes & actions", f"Mise à jour: {NOW_STR}  •  Slide 7/7")

    # Render up to 6 alerts (no scroll)
    max_alerts = 6
    shown = ALERTS[:max_alerts]

    for level, text in shown:
        icon = {"ok": "🟢", "warn": "🟠", "bad": "🔴"}.get(level, "🟢")
        st.markdown(f"<div class='tv-alert'>{icon} {text}</div>", unsafe_allow_html=True)

    st.markdown("<div class='tv-divider'></div>", unsafe_allow_html=True)

    # Suggested next actions (static, but tied to signals)
    st.markdown("<div class='tv-card'><b>Actions recommandées (checklist)</b>", unsafe_allow_html=True)
    actions = [
        "Vérifier top magasins à 0 commande (ruptures / livraison / paiement / horaires).",
        "Contrôler disponibilité des produits très consultés (priorité réassort / substitution).",
        "Si trafic baisse: vérifier campagnes / SEO / incidents site.",
        "Si rebond élevé: vérifier vitesse, pages d’atterrissage, erreurs 404, tracking.",
    ]
    for a in actions:
        st.markdown(f"<div class='tv-small'>• {a}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Small note if excel not loaded
    if FILES["commandes"].exists() and commandes.empty:
        st.markdown(
            "<div class='tv-small'>⚠️ commandes.xls détecté mais non lisible. "
            "Ajoute <b>xlrd</b> dans requirements.txt ou convertis le fichier en .xlsx/.csv.</div>",
            unsafe_allow_html=True
        )


SLIDES = [
    slide_1_health,
    slide_2_traffic_interest,
    slide_3_store_difficulty,
    slide_4_products_frustration,
    slide_5_stock_availability,
    slide_6_funnel_simple,
    slide_7_alerts,
]


# -----------------------------
# Render selected slide
# -----------------------------
# Keep everything within a single viewport height (best-effort)
with st.container():
    SLIDES[slide_index]()

# Footer progress line (compact)
progress = (slide_index + 1) / N_SLIDES
st.progress(progress, text=f"Lecture TV • {slide_index+1}/{N_SLIDES} • boucle {TOTAL_SECONDS}s")
