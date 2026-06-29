from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import os
import unicodedata
from datetime import datetime
from io import BytesIO

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "tablero-prestamos-local")

USERS_FILE = "users.json"
SPREADSHEET_ID = "177cRLYfGVEi0d67FB68xashEOB_hXR-tIxsl6EiIPA8"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


def normalizar_texto(valor):
    valor = str(valor).strip().lower()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    valor = valor.replace(" ", "_")
    valor = valor.replace(".", "")
    valor = valor.replace("/", "_")
    valor = valor.replace("-", "_")
    return valor


def formato_numero(valor):
    try:
        return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"


def formato_entero(valor):
    try:
        return f"{int(valor):,}".replace(",", ".")
    except:
        return "0"


def formato_fecha(valor):
    try:
        if pd.isna(valor) or valor == "":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return ""


def parse_numero(valor):
    try:
        if pd.isna(valor) or valor == "":
            return 0

        valor = str(valor).strip()
        valor = valor.replace("$", "")
        valor = valor.replace(" ", "")
        valor = valor.replace(".", "")
        valor = valor.replace(",", ".")

        return float(valor)
    except:
        return 0


def cargar_hoja(nombre_hoja):
    creds = Credentials.from_service_account_file(
        "google_service_account.json",
        scopes=SCOPES
    )

    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    ws = spreadsheet.worksheet(nombre_hoja)

    values = ws.get_all_values()

    if not values:
        return pd.DataFrame()

    headers = values[0]
    rows = values[1:]

    clean_headers = []
    valid_indexes = []

    for i, header in enumerate(headers):
        header_limpio = normalizar_texto(header)
        if header_limpio != "":
            clean_headers.append(header_limpio)
            valid_indexes.append(i)

    clean_rows = []

    for row in rows:
        clean_row = []
        for i in valid_indexes:
            clean_row.append(row[i] if i < len(row) else "")
        clean_rows.append(clean_row)

    df = pd.DataFrame(clean_rows, columns=clean_headers)
    df = df.dropna(how="all")

    return df


def cargar_usuarios():
    if not os.path.exists(USERS_FILE):
        usuarios_default = [
            {
                "username": "admin",
                "password_hash": generate_password_hash("admin123"),
                "role": "admin"
            }
        ]
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(usuarios_default, f, indent=2, ensure_ascii=False)

    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_usuarios(usuarios):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, indent=2, ensure_ascii=False)


class UsuarioActual:
    def __init__(self, username, role):
        self.username = username
        self.role = role


@app.context_processor
def inject_user():
    if "user" in session:
        return {
            "current_user": UsuarioActual(
                session.get("user"),
                session.get("role", "viewer")
            )
        }
    return {"current_user": None}


@app.after_request
def no_cache(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


def load_cuotas():
    df = cargar_hoja("Cuotas")

    if df.empty:
        return df

    required = [
        "id_prestamo",
        "nro_cuota",
        "capital",
        "intereses",
        "iva",
        "total",
        "fecha_vencimiento",
        "estado",
        "entidad"
    ]

    for c in required:
        if c not in df.columns:
            df[c] = None

    df = df.dropna(subset=["id_prestamo"], how="all").copy()

    df["fecha_vencimiento"] = pd.to_datetime(
        df["fecha_vencimiento"],
        errors="coerce",
        dayfirst=True
    )

    for col in ["capital", "intereses", "iva", "total"]:
        df[col] = df[col].apply(parse_numero)

    df["estado"] = df["estado"].fillna("Pendiente").astype(str)

    return df


def load_prestamos():
    df = cargar_hoja("Seguimiento")

    if df.empty:
        return df

    required = [
        "id_prestamo",
        "fecha_inicio",
        "plazo",
        "capital",
        "tna",
        "tea",
        "iva",
        "tipo_amortizacion",
        "entidad"
    ]

    for c in required:
        if c not in df.columns:
            df[c] = None

    df = df.dropna(subset=["id_prestamo"], how="all").copy()

    df["fecha_inicio"] = pd.to_datetime(
        df["fecha_inicio"],
        errors="coerce",
        dayfirst=True
    )

    for col in ["capital", "tna", "tea", "iva", "plazo"]:
        df[col] = df[col].apply(parse_numero)

    return df


def preparar_cuotas_tabla():
    cuotas = load_cuotas()
    prestamos = load_prestamos()

    if cuotas.empty:
        return []

    if not prestamos.empty:
        columnas_merge = [
            "id_prestamo",
            "fecha_inicio",
            "tna",
            "tea",
            "tipo_amortizacion",
            "plazo"
        ]

        columnas_merge = [c for c in columnas_merge if c in prestamos.columns]
        prestamos_merge = prestamos[columnas_merge].drop_duplicates("id_prestamo")

        cuotas = cuotas.merge(
            prestamos_merge,
            on="id_prestamo",
            how="left"
        )

    cuotas_pendientes_por_id = cuotas[
        ~cuotas["estado"].str.lower().str.contains("pag|cancel", na=False)
    ].groupby("id_prestamo").size().to_dict()

    resultado = []

    for _, row in cuotas.iterrows():
        estado = str(row.get("estado", "Pendiente"))
        estado_lower = estado.lower()

        if "pag" in estado_lower:
            estado_color = "pagado"
        elif "cancel" in estado_lower:
            estado_color = "cancelado"
        elif "venc" in estado_lower:
            estado_color = "vencida"
        else:
            estado_color = "pendiente"

        resultado.append({
            "id": row.get("id_prestamo", ""),
            "id_prestamo": row.get("id_prestamo", ""),
            "entidad": row.get("entidad", ""),
            "numero": row.get("nro_cuota", ""),
            "nro_cuota": row.get("nro_cuota", ""),
            "capital": formato_numero(row.get("capital", 0)),
            "intereses": formato_numero(row.get("intereses", 0)),
            "iva": formato_numero(row.get("iva", 0)),
            "total": formato_numero(row.get("total", 0)),
            "vencimiento": formato_fecha(row.get("fecha_vencimiento", "")),
            "fecha_vencimiento": formato_fecha(row.get("fecha_vencimiento", "")),
            "estado": estado,
            "estado_color": estado_color,
            "cuotas_pendientes": cuotas_pendientes_por_id.get(row.get("id_prestamo", ""), 0),
            "tna": formato_numero(row.get("tna", 0)),
            "tea": formato_numero(row.get("tea", 0)),
            "fecha_toma": formato_fecha(row.get("fecha_inicio", "")),
            "fecha_inicio": formato_fecha(row.get("fecha_inicio", "")),
            "tipo_amortizacion": row.get("tipo_amortizacion", ""),
            "plazo": row.get("plazo", "")
        })

    return resultado


def calcular_indicadores():
    cuotas = load_cuotas()
    prestamos = load_prestamos()

    base_vacia = {
        "fecha_actualizacion": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "ultima_lectura": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total_prestamos": "0",
        "cuotas_faltantes": "0",
        "cuotas_pendientes": "0",
        "cuotas_pagadas": "0",
        "cuotas_vencidas": "0",
        "cuota_estimada_mes": "0,00",
        "bancos_entidades": "0",
        "entidades_activas": "0",
        "total_pendiente": "0,00",
        "capital_pendiente": "0,00",
        "monto_vencido": "0,00",
        "tasa_credito": "0,00",
        "tasa_tna": "0,00",
        "tasa_tea": "0,00",
        "total_cuotas_mes": "0,00",
        "pendiente_mes": "0,00",
        "fecha_adjudicacion": "",
        "deuda_por_entidad": [],
        "proximas": []
    }

    if cuotas.empty:
        return base_vacia

    hoy = pd.Timestamp.today().normalize()
    mes_actual = hoy.month
    anio_actual = hoy.year

    cuotas_pagadas_df = cuotas[
        cuotas["estado"].str.lower().str.contains("pag", na=False)
    ]

    cuotas_abiertas_df = cuotas[
        ~cuotas["estado"].str.lower().str.contains("pag|cancel", na=False)
    ]

    cuotas_vencidas_df = cuotas_abiertas_df[
        cuotas_abiertas_df["fecha_vencimiento"] < hoy
    ]

    cuotas_mes_df = cuotas[
        (cuotas["fecha_vencimiento"].dt.month == mes_actual) &
        (cuotas["fecha_vencimiento"].dt.year == anio_actual)
    ]

    total_pendiente = cuotas_abiertas_df["total"].sum()
    capital_pendiente = cuotas_abiertas_df["capital"].sum()
    monto_vencido = cuotas_vencidas_df["total"].sum()
    total_mes = cuotas_mes_df["total"].sum()

    pendiente_mes = cuotas_mes_df[
        ~cuotas_mes_df["estado"].str.lower().str.contains("pag|cancel", na=False)
    ]["total"].sum()

    cuota_promedio_mes = cuotas_mes_df["total"].mean() if len(cuotas_mes_df) > 0 else 0
    entidades_activas = cuotas_abiertas_df["entidad"].nunique()

    tasa_tna = 0
    tasa_tea = 0
    fecha_adjudicacion = ""

    if not prestamos.empty:
        tasa_tna = prestamos["tna"].mean()
        tasa_tea = prestamos["tea"].mean()

        fechas_validas = prestamos["fecha_inicio"].dropna()
        if len(fechas_validas) > 0:
            fecha_adjudicacion = formato_fecha(fechas_validas.max())

    deuda_entidad = (
        cuotas_abiertas_df
        .groupby("entidad")["total"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    deuda_por_entidad = [
        {
            "entidad": row["entidad"],
            "total": formato_numero(row["total"])
        }
        for _, row in deuda_entidad.iterrows()
    ]

    proximas_df = cuotas_abiertas_df[
        cuotas_abiertas_df["fecha_vencimiento"] >= hoy
    ].sort_values("fecha_vencimiento").head(10)

    proximas = []

    for _, row in proximas_df.iterrows():
        proximas.append({
            "id": row.get("id_prestamo", ""),
            "entidad": row.get("entidad", ""),
            "nro_cuota": row.get("nro_cuota", ""),
            "vencimiento": formato_fecha(row.get("fecha_vencimiento", "")),
            "fecha_vencimiento": formato_fecha(row.get("fecha_vencimiento", "")),
            "total": formato_numero(row.get("total", 0)),
            "estado": row.get("estado", "Pendiente"),
            "estado_color": "pendiente"
        })

    return {
        "fecha_actualizacion": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "ultima_lectura": datetime.now().strftime("%d/%m/%Y %H:%M"),

        "total_prestamos": formato_entero(cuotas["id_prestamo"].nunique()),
        "cuotas_faltantes": formato_entero(len(cuotas_abiertas_df)),
        "cuotas_pendientes": formato_entero(len(cuotas_abiertas_df)),
        "cuotas_pagadas": formato_entero(len(cuotas_pagadas_df)),
        "cuotas_vencidas": formato_entero(len(cuotas_vencidas_df)),

        "cuota_estimada_mes": formato_numero(cuota_promedio_mes),
        "bancos_entidades": formato_entero(entidades_activas),
        "entidades_activas": formato_entero(entidades_activas),

        "total_pendiente": formato_numero(total_pendiente),
        "capital_pendiente": formato_numero(capital_pendiente),
        "monto_vencido": formato_numero(monto_vencido),
        "tasa_credito": formato_numero(tasa_tna),
        "tasa_tna": formato_numero(tasa_tna),
        "tasa_tea": formato_numero(tasa_tea),

        "total_cuotas_mes": formato_numero(total_mes),
        "pendiente_mes": formato_numero(pendiente_mes),
        "fecha_adjudicacion": fecha_adjudicacion,

        "deuda_por_entidad": deuda_por_entidad,
        "proximas": proximas
    }


@app.route("/")
def index():
    if "user" not in session:
        return redirect(url_for("login"))
    return redirect(url_for("dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        usuarios = cargar_usuarios()

        for usuario in usuarios:
            if usuario["username"] == username and check_password_hash(usuario["password_hash"], password):
                session["user"] = usuario["username"]
                session["role"] = usuario.get("role", "viewer")
                return redirect(url_for("dashboard"))

        flash("Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    indicadores = calcular_indicadores()
    cuotas = load_cuotas()
    hoy = pd.Timestamp.today().normalize()

    if cuotas.empty:
        vencidas = pd.DataFrame()
    else:
        abiertas = cuotas[
            ~cuotas["estado"].astype(str).str.lower().str.strip().str.contains("pag|cancel", na=False)
        ].copy()

        vencidas = abiertas[
            abiertas["fecha_vencimiento"] < hoy
        ].copy()

    metrics = indicadores.copy()
    metrics["actualizado"] = indicadores.get("ultima_lectura", "")
    metrics["cuota_mes_promedio"] = indicadores.get("cuota_estimada_mes", "0,00")
    metrics["cuota_mes_cantidad"] = indicadores.get("cuotas_pendientes", "0")
    metrics["tna_promedio"] = indicadores.get("tasa_tna", "0,00")
    metrics["tea_promedio"] = indicadores.get("tasa_tea", "0,00")
    metrics["cuota_mes_total"] = indicadores.get("total_cuotas_mes", "0,00")
    metrics["cuota_mes_pendiente"] = indicadores.get("pendiente_mes", "0,00")
    metrics["ultima_adjudicacion"] = indicadores.get("fecha_adjudicacion", "")
    metrics["adjudicados_mes"] = indicadores.get("total_prestamos", "0")

    charts = {
        "entidades_labels": [],
        "entidades_values": [],
        "estados_labels": [],
        "estados_count": [],
        "meses_labels": [],
        "meses_values": []
    }

    capital_por_entidad = []
    capital_corriente_no_corriente = []
    top_proximos_vencimientos = []

    totales_corriente_no_corriente = {
        "corriente": "0,00",
        "no_corriente": "0,00",
        "total": "0,00",
        "pct_corriente": "0,00%",
        "pct_no_corriente": "0,00%"
    }

    if not cuotas.empty:
        abiertas = cuotas[
            ~cuotas["estado"].astype(str).str.lower().str.strip().str.contains("pag|cancel", na=False)
        ].copy()

        deuda_entidad = (
            abiertas.groupby("entidad")["total"]
            .sum()
            .sort_values(ascending=False)
        )

        charts["entidades_labels"] = deuda_entidad.index.tolist()
        charts["entidades_values"] = deuda_entidad.values.tolist()

        capital_entidad = (
            abiertas.groupby("entidad")["capital"]
            .sum()
            .sort_values(ascending=False)
        )

        capital_por_entidad = [
            {
                "entidad": entidad,
                "capital": formato_numero(capital)
            }
            for entidad, capital in capital_entidad.items()
        ]

        fecha_corte_corriente = hoy + pd.DateOffset(months=12)

        abiertas["clasificacion"] = abiertas["fecha_vencimiento"].apply(
            lambda x: "Corriente" if pd.notna(x) and x <= fecha_corte_corriente else "No corriente"
        )

        capital_clasificado = (
            abiertas
            .groupby(["entidad", "clasificacion"])["capital"]
            .sum()
            .unstack(fill_value=0)
        )

        capital_clasificado["Total"] = capital_clasificado.sum(axis=1)
        capital_clasificado = capital_clasificado.sort_values("Total", ascending=False)

        total_corriente_general = capital_clasificado["Corriente"].sum() if "Corriente" in capital_clasificado.columns else 0
        total_no_corriente_general = capital_clasificado["No corriente"].sum() if "No corriente" in capital_clasificado.columns else 0
        total_capital_general = total_corriente_general + total_no_corriente_general

        pct_corriente_general = (total_corriente_general / total_capital_general * 100) if total_capital_general else 0
        pct_no_corriente_general = (total_no_corriente_general / total_capital_general * 100) if total_capital_general else 0

        totales_corriente_no_corriente = {
            "corriente": formato_numero(total_corriente_general),
            "no_corriente": formato_numero(total_no_corriente_general),
            "total": formato_numero(total_capital_general),
            "pct_corriente": f"{pct_corriente_general:.2f}%".replace(".", ","),
            "pct_no_corriente": f"{pct_no_corriente_general:.2f}%".replace(".", ",")
        }

        for entidad, row in capital_clasificado.iterrows():
            corriente = row.get("Corriente", 0)
            no_corriente = row.get("No corriente", 0)
            total = row.get("Total", 0)

            pct_corriente = (corriente / total * 100) if total else 0
            pct_no_corriente = (no_corriente / total * 100) if total else 0

            capital_corriente_no_corriente.append({
                "entidad": entidad,
                "corriente": formato_numero(corriente),
                "no_corriente": formato_numero(no_corriente),
                "total": formato_numero(total),
                "pct_corriente": f"{pct_corriente:.2f}%".replace(".", ","),
                "pct_no_corriente": f"{pct_no_corriente:.2f}%".replace(".", ",")
            })

        estados = cuotas["estado"].value_counts()

        charts["estados_labels"] = estados.index.tolist()
        charts["estados_count"] = estados.values.tolist()

        pendientes_flujo = cuotas[
            (~cuotas["estado"].astype(str).str.lower().str.strip().str.contains("pag|cancel", na=False)) &
            (cuotas["fecha_vencimiento"] >= hoy)
        ].copy()

        pendientes_flujo["mes_orden"] = pendientes_flujo["fecha_vencimiento"].dt.to_period("M")

        flujo = (
            pendientes_flujo
            .groupby("mes_orden")["total"]
            .sum()
            .sort_index()
        )

        charts["meses_labels"] = [m.strftime("%m/%Y") for m in flujo.index]
        charts["meses_values"] = flujo.values.tolist()

        top_proximos_df = pendientes_flujo.sort_values("fecha_vencimiento").head(10)

        top_proximos_vencimientos = [
            {
                "fecha": formato_fecha(row.get("fecha_vencimiento", "")),
                "entidad": row.get("entidad", ""),
                "cuota": row.get("nro_cuota", ""),
                "capital": formato_numero(row.get("capital", 0)),
                "total": formato_numero(row.get("total", 0))
            }
            for _, row in top_proximos_df.iterrows()
        ]

    return render_template(
        "dashboard.html",
        metrics=metrics,
        vencidas=vencidas,
        charts=charts,
        capital_por_entidad=capital_por_entidad,
        capital_corriente_no_corriente=capital_corriente_no_corriente,
        totales_corriente_no_corriente=totales_corriente_no_corriente,
        top_proximos_vencimientos=top_proximos_vencimientos,
        **indicadores
    )

@app.route("/cuotas")
def cuotas():
    if "user" not in session:
        return redirect(url_for("login"))

    cuotas_lista = preparar_cuotas_tabla()

    estado = request.args.get("estado", "")
    entidad = request.args.get("entidad", "")
    buscar = request.args.get("buscar", "")
    fecha_desde = request.args.get("fecha_desde", "")
    fecha_hasta = request.args.get("fecha_hasta", "")

    if estado:
        cuotas_lista = [
            c for c in cuotas_lista
            if estado.lower() in str(c.get("estado", "")).lower()
        ]

    if entidad:
        cuotas_lista = [
            c for c in cuotas_lista
            if entidad.lower() in str(c.get("entidad", "")).lower()
        ]

    if buscar:
        cuotas_lista = [
            c for c in cuotas_lista
            if buscar.lower() in str(c.get("id", "")).lower()
            or buscar.lower() in str(c.get("id_prestamo", "")).lower()
            or buscar.lower() in str(c.get("entidad", "")).lower()
        ]

    if fecha_desde:
        fecha_desde_dt = pd.to_datetime(fecha_desde, errors="coerce")
        cuotas_lista = [
            c for c in cuotas_lista
            if pd.to_datetime(c.get("fecha_vencimiento", ""), dayfirst=True, errors="coerce") >= fecha_desde_dt
        ]

    if fecha_hasta:
        fecha_hasta_dt = pd.to_datetime(fecha_hasta, errors="coerce")
        cuotas_lista = [
            c for c in cuotas_lista
            if pd.to_datetime(c.get("fecha_vencimiento", ""), dayfirst=True, errors="coerce") <= fecha_hasta_dt
        ]

    entidades = sorted(list(set([c["entidad"] for c in cuotas_lista if c.get("entidad")])))

    cuotas_pendientes = len([
        c for c in cuotas_lista
        if "pag" not in str(c.get("estado", "")).lower()
    ])

    cuotas_pagadas = len([
        c for c in cuotas_lista
        if "pag" in str(c.get("estado", "")).lower()
    ])

    total_cuotas = len(cuotas_lista)

    total_periodo = sum([
        parse_numero(c.get("total", "0"))
        for c in cuotas_lista
    ])

    cuota_promedio = total_periodo / total_cuotas if total_cuotas > 0 else 0

    vencidas_count = len([
        c for c in cuotas_lista
        if "venc" in str(c.get("estado", "")).lower()
    ])

    summary = {
        "cuotas_pendientes": cuotas_pendientes,
        "cuotas_pagadas": cuotas_pagadas,
        "total_cuotas": total_cuotas,
        "entidades": len(entidades),
        "total_periodo": formato_numero(total_periodo),
        "cuota_promedio": formato_numero(cuota_promedio),
        "vencidas": vencidas_count
    }

    return render_template(
        "cuotas.html",
        rows=cuotas_lista,
        entidades=entidades,
        summary=summary,
        estado=estado,
        entidad=entidad,
        buscar=buscar,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta
    )


@app.route("/cuotas/exportar")
def exportar_cuotas():
    if "user" not in session:
        return redirect(url_for("login"))

    cuotas_lista = preparar_cuotas_tabla()

    df = pd.DataFrame(cuotas_lista)

    columnas = [
        "id_prestamo",
        "entidad",
        "nro_cuota",
        "fecha_vencimiento",
        "estado",
        "capital",
        "intereses",
        "iva",
        "total",
        "cuotas_pendientes",
        "tna",
        "tea",
        "fecha_inicio",
        "tipo_amortizacion"
    ]

    df = df[[c for c in columnas if c in df.columns]]

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Detalle Cuotas")

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="detalle_cuotas.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/usuarios", methods=["GET", "POST"])
def usuarios():
    if "user" not in session:
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        flash("No tenés permisos para administrar usuarios.", "danger")
        return redirect(url_for("dashboard"))

    usuarios_data = cargar_usuarios()

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "viewer").strip()

        if not username or not password:
            flash("Usuario y contraseña son obligatorios.", "danger")
            return redirect(url_for("usuarios"))

        existe = any(u["username"] == username for u in usuarios_data)

        if existe:
            flash("El usuario ya existe.", "warning")
            return redirect(url_for("usuarios"))

        usuarios_data.append({
            "username": username,
            "password_hash": generate_password_hash(password),
            "role": role
        })

        guardar_usuarios(usuarios_data)
        flash("Usuario creado correctamente.", "success")

        return redirect(url_for("usuarios"))

    return render_template("usuarios.html", usuarios=usuarios_data)


if __name__ == "__main__":
    app.run(debug=True)