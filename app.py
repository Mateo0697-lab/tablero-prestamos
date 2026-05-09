import os
import json
from datetime import datetime, date
from functools import wraps

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "clave-local-cambiar")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.getenv("EXCEL_PATH", "data/Prestamos 04-2026 - 03-2027.xlsm")
EXCEL_FILE = EXCEL_PATH if os.path.isabs(EXCEL_PATH) else os.path.join(BASE_DIR, EXCEL_PATH)
USERS_FILE = os.path.join(BASE_DIR, "users.json")

COLORS = {
    "primary": "#0b2d4d",
    "secondary": "#155e9f",
    "accent": "#00a3e0",
    "success": "#1f8f5f",
    "warning": "#f2b705",
    "danger": "#c0392b",
}


def ensure_admin_user():
    """Crea un usuario admin inicial si no existe users.json."""
    if os.path.exists(USERS_FILE):
        return
    users = [{"username": "admin", "password_hash": generate_password_hash("admin123"), "role": "admin"}]
    save_users(users)


def load_users():
    ensure_admin_user()
    if not os.path.exists(USERS_FILE):
        return []
    with open(USERS_FILE, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)


def current_user():
    username = session.get("username")
    if not username:
        return None
    return next((u for u in load_users() if u["username"] == username), None)


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("username"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user or user.get("role") != "admin":
            flash("No tenés permisos para acceder a esta sección.", "danger")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper


def money(value):
    try:
        value = float(value)
    except Exception:
        value = 0
    return "$ {:,.2f}".format(value).replace(",", "X").replace(".", ",").replace("X", ".")


def money_short(value):
    try:
        value = float(value)
    except Exception:
        value = 0
    abs_value = abs(value)
    if abs_value >= 1_000_000_000:
        return "$ {:.1f} MM".format(value / 1_000_000_000).replace(".", ",")
    if abs_value >= 1_000_000:
        return "$ {:.1f} M".format(value / 1_000_000).replace(".", ",")
    return money(value)


def pct(value):
    try:
        value = float(value)
    except Exception:
        value = 0
    return "{:.2f}%".format(value).replace(".", ",")


def normalize_columns(df):
    mapping = {}
    for col in df.columns:
        text = str(col).strip()
        lower = text.lower()
        if ("id" in lower and "préstamo" in lower) or ("id" in lower and "prestamo" in lower):
            mapping[col] = "id_prestamo"
        elif "número de cuota" in lower or "numero de cuota" in lower:
            mapping[col] = "nro_cuota"
        elif "fecha de inicio" in lower:
            mapping[col] = "fecha_inicio"
        elif "plazo" in lower:
            mapping[col] = "plazo"
        elif lower.startswith("c: capital") or lower == "capital" or lower.endswith("capital"):
            mapping[col] = "capital"
        elif "tna" in lower:
            mapping[col] = "tna"
        elif "tea" in lower:
            mapping[col] = "tea"
        elif "intereses" in lower:
            mapping[col] = "intereses"
        elif lower.startswith("e: iva") or lower == "iva" or lower.endswith("iva"):
            mapping[col] = "iva"
        elif "total a pagar" in lower:
            mapping[col] = "total_pagar"
        elif "fecha de vencimiento" in lower:
            mapping[col] = "fecha_vencimiento"
        elif "entidad" in lower:
            mapping[col] = "entidad"
        elif lower == "estado":
            mapping[col] = "estado"
        elif "amortización" in lower or "amortizacion" in lower:
            mapping[col] = "tipo_amortizacion"
        elif "restructurado" in lower or "reestructurado" in lower:
            mapping[col] = "reestructurado"
    return df.rename(columns=mapping)


def load_cuotas():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"No se encontró el Excel en: {EXCEL_FILE}")
    df = pd.read_excel(EXCEL_FILE, sheet_name="Cuotas", engine="openpyxl")
    df = normalize_columns(df)
    required = ["id_prestamo", "nro_cuota", "capital", "intereses", "iva", "total_pagar", "fecha_vencimiento", "entidad", "estado", "reestructurado"]
    for c in required:
        if c not in df.columns:
            df[c] = None
    df = df.dropna(subset=["id_prestamo"], how="all").copy()
    df["fecha_vencimiento"] = pd.to_datetime(df["fecha_vencimiento"], errors="coerce")
    for c in ["capital", "intereses", "iva", "total_pagar"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["estado"] = df["estado"].fillna("Pendiente").astype(str)
    df["entidad"] = df["entidad"].fillna("Sin entidad").astype(str).str.strip()
    df["reestructurado"] = df["reestructurado"].fillna("").astype(str)
    hoy = pd.Timestamp(date.today())
    df["dias_vencimiento"] = (df["fecha_vencimiento"] - hoy).dt.days
    df["estado_calculado"] = "Al día"
    pendiente = ~df["estado"].str.lower().str.contains("pag", na=False)
    df.loc[pendiente & (df["dias_vencimiento"] < 0), "estado_calculado"] = "Vencida"
    df.loc[pendiente & (df["dias_vencimiento"].between(0, 7, inclusive="both")), "estado_calculado"] = "Próxima"
    df.loc[~pendiente, "estado_calculado"] = "Pagada"
    return df


def load_prestamos():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Seguimiento", engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    df = normalize_columns(df)
    required = ["id_prestamo", "fecha_inicio", "plazo", "capital", "tna", "tea", "iva", "tipo_amortizacion", "entidad"]
    for c in required:
        if c not in df.columns:
            df[c] = None
    df = df.dropna(subset=["id_prestamo"], how="all").copy()
    df["fecha_inicio"] = pd.to_datetime(df["fecha_inicio"], errors="coerce")
    for c in ["capital", "tna", "tea", "plazo"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["entidad"] = df["entidad"].fillna("Sin entidad").astype(str).str.strip()
    return df


def weighted_average(df, value_col, weight_col):
    try:
        data = df[[value_col, weight_col]].dropna()
        data = data[data[weight_col] > 0]
        if data.empty:
            return 0
        return (data[value_col] * data[weight_col]).sum() / data[weight_col].sum()
    except Exception:
        return 0


def build_metrics(cuotas, prestamos):
    pendientes = ~cuotas["estado_calculado"].eq("Pagada")
    vencidas = cuotas["estado_calculado"].eq("Vencida")
    proximas = cuotas["estado_calculado"].eq("Próxima")
    pagadas = cuotas["estado_calculado"].eq("Pagada")
    hoy = pd.Timestamp(date.today())
    mes_actual = cuotas[cuotas["fecha_vencimiento"].dt.to_period("M") == hoy.to_period("M")]
    mes_pendiente = mes_actual[~mes_actual["estado_calculado"].eq("Pagada")]
    entidad_vigente = cuotas.loc[pendientes, "entidad"].nunique()

    if not prestamos.empty:
        tna_prom = weighted_average(prestamos, "tna", "capital")
        tea_prom = weighted_average(prestamos, "tea", "capital")
        ult_fecha = prestamos["fecha_inicio"].max()
        ult_adjudicacion = ult_fecha.strftime("%d/%m/%Y") if pd.notna(ult_fecha) else "-"
        adjudicados_mes = int((prestamos["fecha_inicio"].dt.to_period("M") == hoy.to_period("M")).sum())
    else:
        tna_prom = tea_prom = 0
        ult_adjudicacion = "-"
        adjudicados_mes = 0

    return {
        "total_pendiente": money(cuotas.loc[pendientes, "total_pagar"].sum()),
        "capital_pendiente": money(cuotas.loc[pendientes, "capital"].sum()),
        "intereses_pendientes": money(cuotas.loc[pendientes, "intereses"].sum()),
        "iva_pendiente": money(cuotas.loc[pendientes, "iva"].sum()),
        "cuotas_faltantes": int(pendientes.sum()),
        "cuotas_pagadas": int(pagadas.sum()),
        "cuotas_vencidas": int(vencidas.sum()),
        "monto_vencido": money(cuotas.loc[vencidas, "total_pagar"].sum()),
        "cuotas_proximas": int(proximas.sum()),
        "prestamos": int(cuotas["id_prestamo"].nunique()),
        "entidades_activas": int(entidad_vigente),
        "cuota_mes_total": money(mes_actual["total_pagar"].sum()),
        "cuota_mes_pendiente": money(mes_pendiente["total_pagar"].sum()),
        "cuota_mes_promedio": money(mes_actual["total_pagar"].mean() if len(mes_actual) else 0),
        "cuota_mes_cantidad": int(len(mes_actual)),
        "tna_promedio": pct(tna_prom),
        "tea_promedio": pct(tea_prom),
        "ultima_adjudicacion": ult_adjudicacion,
        "adjudicados_mes": adjudicados_mes,
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }


def chart_data(cuotas):
    pendientes = cuotas[~cuotas["estado_calculado"].eq("Pagada")].copy()
    by_entidad = pendientes.groupby("entidad", dropna=False)["total_pagar"].sum().sort_values(ascending=False).head(10)
    mensual = pendientes.dropna(subset=["fecha_vencimiento"]).copy()
    mensual["periodo"] = mensual["fecha_vencimiento"].dt.strftime("%m/%Y")
    by_month = mensual.groupby("periodo", sort=False)["total_pagar"].sum().head(24)
    estados = cuotas.groupby("estado_calculado")["total_pagar"].sum().reindex(["Vencida", "Próxima", "Al día", "Pagada"]).fillna(0)
    estados_count = cuotas.groupby("estado_calculado")["id_prestamo"].count().reindex(["Vencida", "Próxima", "Al día", "Pagada"]).fillna(0)
    return {
        "entidades_labels": list(by_entidad.index.astype(str)),
        "entidades_values": [round(x, 2) for x in by_entidad.values],
        "meses_labels": list(by_month.index.astype(str)),
        "meses_values": [round(x, 2) for x in by_month.values],
        "estados_labels": list(estados.index.astype(str)),
        "estados_values": [round(x, 2) for x in estados.values],
        "estados_count": [int(x) for x in estados_count.values],
    }


def prepare_cuotas_table(cuotas, prestamos):
    """Enriquece la hoja Cuotas con datos de Seguimiento y cálculos para la vista Cuotas."""
    df = cuotas.copy()

    if not prestamos.empty:
        seg_cols = ["id_prestamo", "fecha_inicio", "plazo", "tna", "tea", "tipo_amortizacion"]
        seg = prestamos[[c for c in seg_cols if c in prestamos.columns]].copy()
        seg = seg.drop_duplicates(subset=["id_prestamo"], keep="last")
        df = df.merge(seg, on="id_prestamo", how="left", suffixes=("", "_seg"))
    else:
        df["fecha_inicio"] = pd.NaT
        df["plazo"] = 0
        df["tna"] = 0
        df["tea"] = 0
        df["tipo_amortizacion"] = ""

    if "plazo" not in df.columns:
        df["plazo"] = 0
    df["plazo"] = pd.to_numeric(df["plazo"], errors="coerce").fillna(0).astype(int)

    # Cuotas pendientes por préstamo: toda cuota no pagada según el estado calculado.
    pendientes_por_id = (
        df[~df["estado_calculado"].eq("Pagada")]
        .groupby("id_prestamo")["nro_cuota"]
        .count()
        .to_dict()
    )
    df["cuotas_pendientes"] = df["id_prestamo"].map(pendientes_por_id).fillna(0).astype(int)

    # Período de vencimiento para filtro y visualización.
    df["periodo_key"] = df["fecha_vencimiento"].dt.strftime("%Y-%m")
    df["periodo_label"] = df["fecha_vencimiento"].dt.strftime("%m/%Y")

    return df


def build_cuotas_summary(df):
    pendientes = ~df["estado_calculado"].eq("Pagada")
    vencidas = df["estado_calculado"].eq("Vencida")
    return {
        "cuotas_pendientes": int(df.loc[pendientes, "nro_cuota"].count()),
        "total_periodo": money(df["total_pagar"].sum()),
        "cuota_promedio": money(df["total_pagar"].mean() if len(df) else 0),
        "vencidas": int(vencidas.sum()),
    }


@app.context_processor
def inject_globals():
    return {"current_user": current_user(), "colors": COLORS}


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = next((u for u in load_users() if u["username"] == username), None)
        if user and check_password_hash(user["password_hash"], password):
            session["username"] = username
            return redirect(url_for("dashboard"))
        flash("Usuario o contraseña incorrectos.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    try:
        cuotas = load_cuotas()
        prestamos = load_prestamos()
        metrics = build_metrics(cuotas, prestamos)
        charts = chart_data(cuotas)
        vencidas = cuotas[cuotas["estado_calculado"].eq("Vencida")].sort_values("fecha_vencimiento").head(12)
        proximas = cuotas[cuotas["estado_calculado"].eq("Próxima")].sort_values("fecha_vencimiento").head(12)
        error = None
    except Exception as e:
        metrics = {}
        charts = {}
        vencidas = pd.DataFrame()
        proximas = pd.DataFrame()
        error = str(e)
    return render_template("dashboard.html", metrics=metrics, charts=charts, vencidas=vencidas, proximas=proximas, money=money, error=error)


@app.route("/cuotas")
@login_required
def cuotas():
    cuotas_base = load_cuotas()
    prestamos = load_prestamos()
    df_all = prepare_cuotas_table(cuotas_base, prestamos).sort_values(["fecha_vencimiento", "entidad", "id_prestamo"])
    df = df_all.copy()

    entidad = request.args.get("entidad", "")
    estado = request.args.get("estado", "")
    periodo = request.args.get("periodo", "")
    buscar = request.args.get("buscar", "").strip()

    if periodo:
        df = df[df["periodo_key"].eq(periodo)]
    if entidad:
        df = df[df["entidad"].eq(entidad)]
    if estado:
        df = df[df["estado_calculado"].eq(estado)]
    if buscar:
        mask = (
            df["id_prestamo"].astype(str).str.contains(buscar, case=False, na=False)
            | df["entidad"].astype(str).str.contains(buscar, case=False, na=False)
        )
        df = df[mask]

    entidades = sorted(df_all["entidad"].dropna().unique().tolist())
    periodos_df = df_all.dropna(subset=["periodo_key"]).drop_duplicates("periodo_key").sort_values("periodo_key")
    periodos = [(r["periodo_key"], r["periodo_label"]) for _, r in periodos_df.iterrows()]
    summary = build_cuotas_summary(df)
    rows = df.head(1000).to_dict("records")
    return render_template(
        "cuotas.html",
        rows=rows,
        entidades=entidades,
        periodos=periodos,
        entidad=entidad,
        estado=estado,
        periodo=periodo,
        buscar=buscar,
        summary=summary,
        money=money,
        pct=pct,
    )


@app.route("/usuarios", methods=["GET", "POST"])
@login_required
@admin_required
def usuarios():
    users = load_users()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "viewer")
        if not username or not password:
            flash("Completá usuario y contraseña.", "warning")
        elif any(u["username"] == username for u in users):
            flash("Ese usuario ya existe.", "warning")
        else:
            users.append({"username": username, "password_hash": generate_password_hash(password), "role": role})
            save_users(users)
            flash("Usuario creado correctamente.", "success")
            return redirect(url_for("usuarios"))
    return render_template("usuarios.html", users=users)


@app.route("/health")
def health():
    return {"status": "ok", "excel": EXCEL_FILE, "exists": os.path.exists(EXCEL_FILE)}


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
