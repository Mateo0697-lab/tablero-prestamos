from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import os
import unicodedata
from datetime import datetime

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


def normalize_columns(df):
    df.columns = [normalizar_texto(c) for c in df.columns]
    return df


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
            if i < len(row):
                clean_row.append(row[i])
            else:
                clean_row.append("")

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
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

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
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

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
            "tipo_amortizacion"
        ]

        prestamos_merge = prestamos[columnas_merge].drop_duplicates("id_prestamo")

        cuotas = cuotas.merge(
            prestamos_merge,
            on="id_prestamo",
            how="left"
        )

    cuotas_pendientes_por_id = cuotas[
        ~cuotas["estado"].str.lower().str.contains("pag", na=False)
    ].groupby("id_prestamo").size().to_dict()

    resultado = []

    for _, row in cuotas.iterrows():
        estado = str(row.get("estado", "Pendiente"))
        estado_lower = estado.lower()

        if "pag" in estado_lower:
            estado_color = "pagada"
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
            "tipo_amortizacion": row.get("tipo_amortizacion", "")
        })

    return resultado


def calcular_indicadores():
    cuotas = load_cuotas()
    prestamos = load_prestamos()

    if cuotas.empty:
        return {}

    hoy = pd.Timestamp.today().normalize()
    mes_actual = hoy.month
    anio_actual = hoy.year

    cuotas_pagadas_df = cuotas[
        cuotas["estado"].str.lower().str.contains("pag", na=False)
    ]

    cuotas_abiertas_df = cuotas[
        ~cuotas["estado"].str.lower().str.contains("pag", na=False)
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
        ~cuotas_mes_df["estado"].str.lower().str.contains("pag", na=False)
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

    proximas_df = cuotas_abiertas_df.sort_values("fecha_vencimiento").head(10)

    proximas = []

    for _, row in proximas_df.iterrows():
        proximas.append({
            "id": row.get("id_prestamo", ""),
            "entidad": row.get("entidad", ""),
            "vencimiento": formato_fecha(row.get("fecha_vencimiento", "")),
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
            ~cuotas["estado"].astype(str).str.lower().str.contains("pag", na=False)
        ].copy()

        vencidas = abiertas[
            abiertas["fecha_vencimiento"] < hoy
        ].copy()

    metrics = indicadores.copy()
    metrics["actualizado"] = indicadores.get("ultima_lectura", "")

    return render_template(
        "dashboard.html",
        metrics=metrics,
        vencidas=vencidas,
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

    if estado:
        cuotas_lista = [
            c for c in cuotas_lista
            if estado.lower() in str(c["estado"]).lower()
        ]

    if entidad:
        cuotas_lista = [
            c for c in cuotas_lista
            if entidad.lower() in str(c["entidad"]).lower()
        ]

    if buscar:
        cuotas_lista = [
            c for c in cuotas_lista
            if buscar.lower() in str(c["id"]).lower()
            or buscar.lower() in str(c["entidad"]).lower()
        ]

    entidades = sorted(list(set([c["entidad"] for c in cuotas_lista if c["entidad"]])))

    return render_template(
        "cuotas.html",
        cuotas=cuotas_lista,
        entidades=entidades
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