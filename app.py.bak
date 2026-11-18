
import calendar
import hashlib
import html
import random
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import holidays
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="El Luchador | Recursos Humanos",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).resolve().parent
RUTA_EXCEL = BASE / "data" / "sueldos.xlsx"

NEGRO = "#090A10"
AZUL = "#2F6BFF"
NARANJA = "#FF7A1A"
BLANCO = "#F5F7FF"
GRIS = "#AAB4CC"
SUPERFICIE = "#151824"

MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre",
    "Diciembre",
]

DIAS = [
    "Lunes", "Martes", "Miércoles",
    "Jueves", "Viernes", "Sábado", "Domingo",
]

HOY_AR = datetime.now(
    ZoneInfo("America/Argentina/Buenos_Aires")
).date()


# ============================================================
# 2. ESTILOS
# ============================================================

st.markdown(
    """
<style>
.stApp {
    background: #090A10;
    color: #F5F7FF;
}

[data-testid="stSidebar"] {
    background: #10131E;
    border-right: 1px solid #303648;
}

.block-container {
    padding-top: 1.5rem;
    max-width: 1550px;
}

h1, h2, h3 {
    color: #F5F7FF !important;
}

[data-testid="stMetric"] {
    border-radius: 14px;
    background: #151824;
    border: 1px solid #2C3145;
    padding: 13px 17px;
}

[data-testid="stMetricLabel"] p {
    color: #B0B9CE !important;
}

[data-testid="stMetricValue"] {
    color: #F5F7FF !important;
}

[data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p {
    color: #C4CCE0 !important;
}

.stButton > button[kind="primary"] {
    background: #2F6BFF;
    color: #FFFFFF;
    border: 0;
}

.stButton > button {
    border-radius: 9px;
}

.hero {
    border: 1px solid #344C8C;
    border-radius: 16px;
    padding: 22px 27px;
    background: linear-gradient(
        110deg, #1B2B56, #151824
    );
    margin-bottom: 16px;
}

.hero .eyebrow {
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.5px;
    color: #9BB8FF;
}

.hero h1 {
    font-size: 31px !important;
    margin: 7px 0;
}

.hero p {
    margin: 0;
    color: #CBD5EF;
}

.notice {
    padding: 11px 15px;
    border-radius: 10px;
    border: 1px solid #A45F2A;
    background: #282018;
    color: #FFCEAA;
    margin: 11px 0 16px;
}

.calbox {
    border: 1px solid #343A50;
    background: #151824;
    border-radius: 10px;
    padding: 7px;
    min-height: 178px;
    margin: 3px 0;
    overflow: hidden;
}

.calbox.holiday {
    border-color: #FF7A1A;
    background: #251C19;
}

.calbox.leave {
    border-color: #2F6BFF;
    background: #172446;
}

.calbox.risk {
    border: 2px solid #FF7A1A;
    background: #302018;
}

.calbox.empty {
    opacity: 0.16;
}

.calnum {
    font-weight: 800;
    font-size: 15px;
}

.calholiday {
    font-size: 10px;
    color: #FFBA82;
    line-height: 1.25;
    margin-top: 5px;
}

.calshift {
    font-size: 11px;
    color: #8CACFF;
    font-weight: 800;
    margin-top: 8px;
}

.calname {
    font-size: 10px;
    color: #E2E8F7;
    line-height: 1.35;
    overflow-wrap: anywhere;
}

.calstatus {
    font-size: 11px;
    color: #FFB77B;
    margin-top: 7px;
}

.legend {
    font-size: 12px;
    color: #BFC9DC;
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# 3. UTILIDADES GENERALES
# ============================================================

def norm(txt):
    txt = unicodedata.normalize("NFKD", str(txt))

    return "".join(
        x for x in txt
        if not unicodedata.combining(x)
    ).strip().lower()


def entera(valor):
    return f"{float(valor):,.0f}".replace(",", ".")


def ars(valor):
    return f"$ {entera(valor)} ARS"


def millones(valor):
    numero = f"{float(valor) / 1e6:.2f}"
    return "$ " + numero.replace(".", ",") + " M ARS"


def semilla(valor):
    texto = str(valor).encode("utf-8")
    digest = hashlib.sha256(texto).hexdigest()
    return int(digest[:12], 16)


def area_de_rol(rol):
    x = norm(rol)

    if "cortador" in x:
        return "Carnicería"

    if "cajer" in x:
        return "Caja"

    if "verdul" in x:
        return "Verdulería"

    if "produccion" in x:
        return "Producción"

    if "administrativ" in x:
        return "Administración"

    return "Otros"


def cabecera(nombre, descripcion):
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">
                EL LUCHADOR / PEOPLE OPERATIONS
            </div>
            <h1>{html.escape(nombre)}</h1>
            <p>{html.escape(descripcion)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def nota():
    st.markdown(
        """
        <div class="notice">
            <b>ENTORNO DEMO:</b>
            los salarios, identificadores, roles y sucursales
            proceden de <b>sueldos.xlsx</b>.
            Las fechas de ingreso, nacimiento, turnos,
            vacaciones y reemplazos son
            <b>SIMULADOS</b>.
            Los cambios no modifican el Excel original.
        </div>
        """,
        unsafe_allow_html=True,
    )


def grafico(fig):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=BLANCO,
        margin=dict(
            l=15,
            r=15,
            t=45,
            b=15,
        ),
        legend_title_text="",
    )

    return fig


# ============================================================
# 4. CARGA Y FILTRADO DE SUELDOS
# ============================================================

def desde_excel(origen):
    df = pd.read_excel(
        origen,
        engine="openpyxl",
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    requeridas = [
        "Razon social",
        "Sucursal",
        "Local",
        "Empleados",
        "Rol",
        "Semanal",
        "Salario mensual",
    ]

    faltan = [
        c for c in requeridas
        if c not in df.columns
    ]

    if faltan:
        raise ValueError(
            "Faltan columnas en el Excel: "
            + ", ".join(faltan)
        )

    df["Fila Excel"] = df.index + 2

    columnas_numericas = [
        "Salario mensual",
        "Semanal",
        "semanal real",
        "Aguinaldo junio",
        "Aguinaldo diciembre",
    ]

    for columna in columnas_numericas:
        if columna not in df.columns:
            df[columna] = None

        df[columna] = pd.to_numeric(
            df[columna],
            errors="coerce",
        )

    # Registros con campos esenciales completos
    df = df.dropna(
        subset=requeridas
    ).copy()

    # Exclusiones expresamente solicitadas
    excluidos = (
        df["Empleados"]
        .astype(str)
        .map(norm)
        .str.fullmatch(
            "compras verduleria|encargado|natalia"
        )
    )

    df = df.loc[
        (~excluidos)
        & (df["Salario mensual"] > 0)
    ].copy()

    # Datos que sí proceden del Excel
    df["Nombre"] = (
        df["Empleados"]
        .astype(str)
        .str.strip()
    )

    def razon_social(valor):
        x = norm(valor)

        if x == "mision":
            return "Misión"

        if x == "ilcomendatori":
            return "Il Comendatori"

        return str(valor)

    df["Razón social"] = (
        df["Razon social"].map(razon_social)
    )

    num = (
        df["Sucursal"]
        .astype(str)
        .str.extract(r"(\d+)")[0]
        .fillna("0")
    )

    df["ID"] = (
        "S"
        + num
        + "-F"
        + df["Fila Excel"]
        .astype(int)
        .astype(str)
        .str.zfill(2)
    )

    df["Área"] = (
        df["Rol"].map(area_de_rol)
    )

    # Regla salarial principal
    df["Mensual ARS"] = df["Salario mensual"]

    df["Semanal equivalente ARS"] = (
        df["Mensual ARS"] * 12 / 52
    )

    df["Anual base ARS"] = (
        df["Mensual ARS"] * 12
    )

    # Datos adicionales ficticios
    ingresos = []
    nacimientos = []

    for uid in df["ID"]:
        rng = random.Random(semilla(uid))

        ingresos.append(
            date(
                rng.randint(2014, 2025),
                rng.randint(1, 12),
                rng.randint(1, 28),
            )
        )

        nacimientos.append(
            date(
                rng.randint(1968, 2003),
                rng.randint(1, 12),
                rng.randint(1, 28),
            )
        )

    df["Ingreso DEMO"] = ingresos
    df["Nacimiento DEMO"] = nacimientos

    df["Documento"] = (
        "NO REGISTRADO — DEMO"
    )

    df["ART"] = "Pendiente de cargar"

    df["Domicilio"] = "Pendiente de cargar"

    return df.reset_index(drop=True)


# ============================================================
# 5. CALENDARIO ARGENTINO
# ============================================================

def feriados_argentina(anio):
    nacionales = holidays.country_holidays(
        "AR",
        years=[anio],
        language="es",
    )

    result = {
        fecha: (
            str(nombre),
            "Feriado nacional",
        )
        for fecha, nombre in nacionales.items()
    }

    # Fechas especiales de 2026.
    # Se distinguen los días no laborables
    # de los feriados nacionales y provinciales.

    if anio == 2026:

        dias_turisticos = [
            (3, 23),
            (7, 10),
            (12, 7),
        ]

        for mes, dia in dias_turisticos:
            result[date(2026, mes, dia)] = (
                "Día no laborable con fines turísticos",
                "No laborable",
            )

        result[date(2026, 11, 9)] = (
            "Visita de Su Santidad el Papa León XIV",
            "Feriado nacional",
        )

        result[date(2026, 11, 11)] = (
            "Feriado extraordinario — Provincia de Buenos Aires",
            "Feriado provincial",
        )

    return result


# ============================================================
# 6. TURNOS SIMULADOS
# ============================================================

def turno_base(uid, dia):
    """
    Turno ficticio y reproducible.

    Se establece un día de descanso por semana
    y turnos rotativos por semana.

    No representa un horario laboral validado.
    """

    descanso = semilla(
        "rest-" + uid
    ) % 7

    if dia.weekday() == descanso:
        return "Descanso"

    alternancia = (
        semilla("shift-" + uid)
        + dia.isocalendar().week
    ) % 2

    if alternancia == 0:
        return "Mañana"

    return "Tarde"


def ausente(uid, dia):
    return any(
        s["ID"] == uid
        and s["Estado"] == "Aprobada"
        and s["Inicio"] <= dia <= s["Fin"]

        for s in st.session_state.rrhh_vacaciones
    )


def turno(uid, dia, escenario=None):
    if ausente(uid, dia):
        return "Vacaciones"

    if escenario:
        empleado, inicio, fin = escenario

        if (
            uid == empleado
            and inicio <= dia <= fin
        ):
            return "Vacaciones simuladas"

    return turno_base(uid, dia)


def plan_turnos(plantilla, dia):
    aux = plantilla[
        [
            "ID",
            "Nombre",
            "Rol",
            "Área",
            "Local",
        ]
    ].copy()

    aux["Turno"] = aux["ID"].map(
        lambda uid: turno(uid, dia)
    )

    return aux


# ============================================================
# 7. GESTIÓN DE VACACIONES
# ============================================================

def antiguedad_a_cierre(ingreso, anio):
    corte = date(anio, 12, 31)

    anios = (
        corte.year
        - ingreso.year
        - (
            (corte.month, corte.day)
            < (ingreso.month, ingreso.day)
        )
    )

    return max(0, anios)


def cupo_orientativo(ingreso, anio):
    antiguedad = antiguedad_a_cierre(
        ingreso,
        anio,
    )

    if antiguedad <= 5:
        return 14

    if antiguedad <= 10:
        return 21

    if antiguedad <= 20:
        return 28

    return 35


def dias_en_anio(inicio, fin, anio):
    a = max(
        inicio,
        date(anio, 1, 1),
    )

    b = min(
        fin,
        date(anio, 12, 31),
    )

    return max(
        0,
        (b - a).days + 1,
    )


def dias_aprobados(uid, anio):
    return sum(
        dias_en_anio(
            s["Inicio"],
            s["Fin"],
            anio,
        )

        for s in st.session_state.rrhh_vacaciones

        if (
            s["ID"] == uid
            and s["Estado"] == "Aprobada"
        )
    )


def solapada(
    uid,
    inicio,
    fin,
    except_id=None,
):
    return any(
        s["ID"] == uid
        and s["Estado"] != "Rechazada"
        and s["N"] != except_id
        and inicio <= s["Fin"]
        and fin >= s["Inicio"]

        for s in st.session_state.rrhh_vacaciones
    )


# ============================================================
# 8. SIMULADOR DE COBERTURA
# ============================================================

def estudiar_cobertura(
    plantilla,
    uid,
    inicio,
    fin,
    minimo,
):
    objetivo = plantilla.loc[
        plantilla["ID"] == uid
    ].iloc[0]

    colegas = plantilla.loc[
        (plantilla["Local"] == objetivo["Local"])
        & (plantilla["Área"] == objetivo["Área"])
        & (plantilla["ID"] != uid)
    ]

    filas = []

    for delta in range(
        (fin - inicio).days + 1
    ):
        dia = inicio + timedelta(days=delta)

        base = turno_base(uid, dia)

        # No es necesario cubrir un descanso
        if base == "Descanso":
            continue

        quedan = sum(
            turno(r["ID"], dia) == base

            for _, r in colegas.iterrows()
        )

        faltan = max(
            0,
            int(minimo) - quedan,
        )

        filas.append(
            {
                "Fecha": dia,
                "Turno": base,
                "Personal restante": quedan,
                "Mínimo": int(minimo),
                "Faltantes": faltan,
                "Estado": (
                    "Revisar"
                    if faltan
                    else "Cubierto"
                ),
            }
        )

    columnas = [
        "Fecha",
        "Turno",
        "Personal restante",
        "Mínimo",
        "Faltantes",
        "Estado",
    ]

    return pd.DataFrame(
        filas,
        columns=columnas,
    )


# ============================================================
# 9. CANDIDATOS DE REEMPLAZO
# ============================================================

def candidatos_reemplazo(
    plantilla,
    uid,
    impacto,
    incluir_otras=True,
):
    objetivo = plantilla.loc[
        plantilla["ID"] == uid
    ].iloc[0]

    candidatos = plantilla.loc[
        (plantilla["Área"] == objetivo["Área"])
        & (plantilla["ID"] != uid)
    ]

    filas = []

    fechas_riesgo = impacto.loc[
        impacto["Faltantes"] > 0
    ]

    for _, candidato in candidatos.iterrows():

        otra = (
            candidato["Local"]
            != objetivo["Local"]
        )

        if otra and not incluir_otras:
            continue

        opciones = []

        for _, problema in fechas_riesgo.iterrows():

            dia = problema["Fecha"]

            turno_necesario = (
                problema["Turno"]
            )

            actual = turno(
                candidato["ID"],
                dia,
            )

            if actual.startswith("Vacaciones"):
                continue

            # Ya está trabajando en ese turno
            # y su presencia ya fue contada.
            if (
                not otra
                and actual == turno_necesario
            ):
                continue

            if otra:
                motivo = "Traslado de sucursal"
                prioridad = 3

            elif actual == "Descanso":
                motivo = "Reprogramar descanso"
                prioridad = 2

            else:
                motivo = "Cambiar turno"
                prioridad = 1

            opciones.append(
                (
                    dia,
                    motivo,
                    prioridad,
                )
            )

        if opciones:
            filas.append(
                {
                    "ID": candidato["ID"],
                    "Trabajador": candidato["Nombre"],
                    "Rol": candidato["Rol"],
                    "Sucursal": candidato["Local"],
                    "Días posibles": len(opciones),
                    "Condición": ", ".join(
                        sorted(
                            set(
                                o[1]
                                for o in opciones
                            )
                        )
                    ),
                    "Prioridad": min(
                        o[2]
                        for o in opciones
                    ),
                }
            )

    columnas = [
        "ID",
        "Trabajador",
        "Rol",
        "Sucursal",
        "Días posibles",
        "Condición",
        "Prioridad",
    ]

    if not filas:
        return pd.DataFrame(columns=columnas)

    return (
        pd.DataFrame(filas)
        .sort_values(
            [
                "Días posibles",
                "Prioridad",
            ],
            ascending=[
                False,
                True,
            ],
        )
        .reset_index(drop=True)
    )


# ============================================================
# 10. VISUALIZACIÓN DE CALENDARIOS
# ============================================================

def lista_resumida(nombres):
    contenido = ""

    for nombre in nombres[:2]:
        contenido += (
            '<div class="calname">• '
            + html.escape(nombre[:34])
            + "</div>"
        )

    if len(nombres) > 2:
        contenido += (
            '<div class="calname">'
            f"+{len(nombres) - 2} más"
            "</div>"
        )

    return contenido


def calendario_mensual(
    anio,
    mes,
    pintar,
):
    nombres = [
        "Lun", "Mar", "Mié", "Jue",
        "Vie", "Sáb", "Dom",
    ]

    for col, dia in zip(
        st.columns(7),
        nombres,
    ):
        col.markdown(f"**{dia}**")

    calendario = calendar.Calendar(
        firstweekday=0
    )

    semanas = calendario.monthdatescalendar(
        anio,
        mes,
    )

    for semana in semanas:

        columnas = st.columns(7)

        for col, dia in zip(
            columnas,
            semana,
        ):

            if dia.month != mes:
                col.markdown(
                    '<div class="calbox empty"></div>',
                    unsafe_allow_html=True,
                )

            else:
                contenido, css = pintar(dia)

                col.markdown(
                    f"""
                    <div class="calbox {css}">
                        <div class="calnum">
                            {dia.day}
                        </div>
                        {contenido}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def etiqueta(uid, plantilla):
    r = plantilla.loc[
        plantilla["ID"] == uid
    ].iloc[0]

    return (
        f'{uid} · {r["Nombre"]} · '
        f'{r["Local"]} · {r["Rol"]}'
    )


# ============================================================
# 11. CARGA DE ARCHIVO Y ESTADO DE SESIÓN
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            background:linear-gradient(
                125deg,#1D346F,#141827
            );
            padding:22px;
            border:1px solid #4066B8;
            border-radius:14px;
        ">
            <h3 style="margin:0">
                EL LUCHADOR
            </h3>
            <p style="
                margin:6px 0 0;
                color:#B1C8FF;
            ">
                RECURSOS HUMANOS · DEMO
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if RUTA_EXCEL.exists():
        fuente = RUTA_EXCEL

    else:
        fuente = st.file_uploader(
            "Cargar sueldos.xlsx",
            type=["xlsx"],
        )

        if fuente is None:
            st.info(
                "Coloca sueldos.xlsx dentro de data/ "
                "o cárgalo aquí."
            )
            st.stop()


try:
    personal = desde_excel(fuente)

except Exception as ex:
    st.error(
        f"No se pudo leer sueldos.xlsx: {ex}"
    )
    st.stop()


if personal.empty:
    st.error(
        "No se encontraron trabajadores válidos."
    )
    st.stop()


# ============================================================
# 12. DATOS INICIALES SIMULADOS
# ============================================================

if "rrhh_vacaciones" not in st.session_state:

    st.session_state.rrhh_vacaciones = []

    ejemplo = personal.sample(
        n=min(6, len(personal)),
        random_state=19,
    )

    for i, (_, r) in enumerate(
        ejemplo.iterrows()
    ):
        inicio = date(
            2026,
            2 + i,
            9,
        )

        st.session_state.rrhh_vacaciones.append(
            {
                "N": i + 1,
                "ID": r["ID"],
                "Inicio": inicio,
                "Fin": inicio + timedelta(
                    days=5 + (i % 3)
                ),
                "Estado": (
                    "Aprobada"
                    if i % 2 == 0
                    else "Pendiente"
                ),
                "Origen": "DEMO",
            }
        )


if "rrhh_anticipos" not in st.session_state:
    st.session_state.rrhh_anticipos = []


if "rrhh_alertas_vistas" not in st.session_state:
    st.session_state.rrhh_alertas_vistas = set()


if "rrhh_reemplazos" not in st.session_state:
    st.session_state.rrhh_reemplazos = {}


# ============================================================
# 13. NAVEGACIÓN
# ============================================================

with st.sidebar:

    st.markdown("### Navegación")

    pagina = st.radio(
        "Sección",
        [
            "Resumen ejecutivo",
            "Legajos digitales",
            "Calendario y turnos",
            "Vacaciones",
            "Remuneraciones",
            "Centro de alertas",
        ],
        label_visibility="collapsed",
    )

    st.divider()

    st.markdown("### Filtros generales")

    sedes = sorted(
        personal["Local"].unique().tolist()
    )

    sedes_activas = st.multiselect(
        "Sucursales",
        sedes,
        default=sedes,
    )

    anio = st.selectbox(
        "Año",
        [
            2026,
            2027,
            2028,
        ],
        index=0,
    )

    st.caption(
        "Salario: columna 'Salario mensual' "
        "· Moneda: ARS"
    )

    st.caption(
        "Turnos, fechas de ingreso y vacaciones: "
        "datos simulados"
    )


base = personal.loc[
    personal["Local"].isin(sedes_activas)
].copy()


if base.empty:
    st.warning(
        "Selecciona al menos una sucursal."
    )
    st.stop()


# ============================================================
# 14. DASHBOARD EJECUTIVO
# ============================================================

if pagina == "Resumen ejecutivo":

    cabecera(
        "Panel de Recursos Humanos",
        "Dotación, presupuesto salarial, cobertura "
        "y tendencias por sucursal, área y trabajador.",
    )

    nota()

    if "rh_sede" not in st.session_state:
        st.session_state.rh_sede = "Todas"

    opciones = (
        ["Todas"]
        + sorted(base["Local"].unique().tolist())
    )

    if st.session_state.rh_sede not in opciones:
        st.session_state.rh_sede = "Todas"

    st.markdown(
        "#### Filtrar rápidamente por sucursal"
    )

    columnas = st.columns(len(opciones))

    for col, opcion in zip(
        columnas,
        opciones,
    ):

        seleccionado = (
            opcion == st.session_state.rh_sede
        )

        if col.button(
            opcion,
            key="btn_" + opcion,
            type=(
                "primary"
                if seleccionado
                else "secondary"
            ),
            use_container_width=True,
        ):
            st.session_state.rh_sede = opcion
            st.rerun()

    muestra = base.copy()

    if st.session_state.rh_sede != "Todas":
        muestra = muestra.loc[
            muestra["Local"]
            == st.session_state.rh_sede
        ]

    c1, c2 = st.columns(2)

    area = c1.selectbox(
        "Área",
        ["Todas"]
        + sorted(
            muestra["Área"].unique().tolist()
        ),
    )

    desglose = c2.radio(
        "Agrupar por",
        [
            "Sucursal",
            "Razón social",
            "Área",
            "Trabajador",
        ],
        horizontal=True,
    )

    if area != "Todas":
        muestra = muestra.loc[
            muestra["Área"] == area
        ]

    salario = float(
        muestra["Mensual ARS"].sum()
    )

    sac_jun = float(
        muestra["Aguinaldo junio"]
        .fillna(0)
        .sum()
    )

    sac_dic = float(
        muestra["Aguinaldo diciembre"]
        .fillna(0)
        .sum()
    )

    ids = set(muestra["ID"])

    pendientes = sum(
        s["Estado"] == "Pendiente"
        and s["ID"] in ids

        for s in st.session_state.rrhh_vacaciones
    )

    cols = st.columns(4)

    indicadores = [
        (
            "Trabajadores",
            str(len(muestra)),
        ),
        (
            "Costo mensual base",
            millones(salario),
        ),
        (
            "Costo anual base",
            millones(salario * 12),
        ),
        (
            "Vacaciones pendientes",
            str(pendientes),
        ),
    ]

    for columna, indicador in zip(
        cols,
        indicadores,
    ):
        columna.metric(
            indicador[0],
            indicador[1],
        )

    t1, t2, t3, t4 = st.tabs(
        [
            "Distribución",
            "Tendencias",
            "Cobertura diaria",
            "Detalle por trabajador",
        ]
    )

    # --------------------------------------------------------
    # DISTRIBUCIÓN
    # --------------------------------------------------------

    with t1:

        var = {
            "Sucursal": "Local",
            "Razón social": "Razón social",
            "Área": "Área",
            "Trabajador": "ID",
        }[desglose]

        agrup = (
            muestra.groupby(
                var,
                as_index=False,
            )
            .agg(
                Trabajadores=("ID", "count"),
                Costo=("Mensual ARS", "sum"),
            )
            .rename(
                columns={var: "Grupo"}
            )
        )

        if desglose == "Trabajador":

            nombres = {
                r["ID"]: (
                    f'{r["ID"]} · '
                    f'{r["Nombre"]} · '
                    f'{r["Rol"]}'
                )

                for _, r in muestra.iterrows()
            }

            agrup["Grupo"] = (
                agrup["Grupo"].map(nombres)
            )

        a, b = st.columns([1.4, 1])

        with a:

            fig = px.bar(
                agrup.sort_values(
                    "Costo",
                    ascending=False,
                ),
                x="Grupo",
                y="Costo",
                color="Grupo",
                color_discrete_sequence=[
                    AZUL,
                    NARANJA,
                    "#7187C1",
                    "#B86A47",
                ],
                title="Costo mensual por grupo (ARS)",
            )

            fig.update_layout(
                showlegend=False,
                xaxis_title="",
                yaxis_title="ARS",
            )

            st.plotly_chart(
                grafico(fig),
                use_container_width=True,
            )

        with b:

            fig = px.pie(
                agrup,
                names="Grupo",
                values="Trabajadores",
                hole=0.62,
                color_discrete_sequence=[
                    AZUL,
                    NARANJA,
                    "#7187C1",
                    "#B86A47",
                ],
                title="Proporción de empleados",
            )

            st.plotly_chart(
                grafico(fig),
                use_container_width=True,
            )

        st.dataframe(
            agrup.rename(
                columns={
                    "Costo": "Costo mensual ARS"
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

    # --------------------------------------------------------
    # TENDENCIAS
    # --------------------------------------------------------

    with t2:

        st.info(
            "Esta es una proyección, no un historial "
            "real de nómina. El Excel contiene un "
            "salario mensual por trabajador, no doce "
            "remuneraciones mensuales diferentes."
        )

        serie = pd.DataFrame(
            {
                "Mes": MESES,
                "Base": [salario] * 12,
                "SAC": [
                    sac_jun if m == 6
                    else sac_dic if m == 12
                    else 0

                    for m in range(1, 13)
                ],
            }
        )

        serie["Base + SAC"] = (
            serie["Base"]
            + serie["SAC"]
        )

        fig = go.Figure()

        fig.add_trace(
            go.Scatter(
                x=serie["Mes"],
                y=serie["Base"],
                name="Salario base",
                mode="lines+markers",
                line=dict(
                    color=AZUL,
                    width=3,
                ),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=serie["Mes"],
                y=serie["Base + SAC"],
                name="Base + SAC Excel",
                mode="lines+markers",
                line=dict(
                    color=NARANJA,
                    width=3,
                ),
            )
        )

        fig.update_layout(
            title="Proyección anual del costo salarial (ARS)",
            yaxis_title="ARS",
        )

        st.plotly_chart(
            grafico(fig),
            use_container_width=True,
        )

        # Tendencia de vacaciones aprobadas
        vac_meses = [0] * 12

        for solicitud in st.session_state.rrhh_vacaciones:

            if (
                solicitud["ID"] not in ids
                or solicitud["Estado"] != "Aprobada"
            ):
                continue

            total_dias = (
                solicitud["Fin"]
                - solicitud["Inicio"]
            ).days + 1

            for n in range(total_dias):

                dia = (
                    solicitud["Inicio"]
                    + timedelta(days=n)
                )

                if dia.year == anio:
                    vac_meses[dia.month - 1] += 1

        fig2 = px.line(
            x=MESES,
            y=vac_meses,
            markers=True,
            title=(
                "Días corridos de vacaciones "
                "aprobadas (DEMO)"
            ),
            labels={
                "x": "Mes",
                "y": "Días",
            },
        )

        fig2.update_traces(
            line_color=NARANJA,
        )

        st.plotly_chart(
            grafico(fig2),
            use_container_width=True,
        )

    # --------------------------------------------------------
    # COBERTURA DIARIA
    # --------------------------------------------------------

    with t3:

        dia = st.date_input(
            "Fecha de cobertura",
            value=date(anio, 11, 11),
            key="fecha_dashboard",
        )

        tabla = plan_turnos(
            muestra,
            dia,
        )

        st.dataframe(
            tabla.rename(
                columns={"ID": "Código"}
            ),
            hide_index=True,
            use_container_width=True,
        )

        conteos = (
            tabla.groupby(
                [
                    "Local",
                    "Área",
                    "Turno",
                ],
                as_index=False,
            )
            .size()
            .rename(
                columns={
                    "size": "Personas"
                }
            )
        )

        st.markdown(
            "#### Dotación por sucursal, área y turno"
        )

        st.dataframe(
            conteos,
            hide_index=True,
            use_container_width=True,
        )

    # --------------------------------------------------------
    # DETALLE INDIVIDUAL
    # --------------------------------------------------------

    with t4:

        detalle = muestra[
            [
                "ID",
                "Nombre",
                "Rol",
                "Local",
                "Razón social",
                "Mensual ARS",
                "Anual base ARS",
            ]
        ].copy()

        st.dataframe(
            detalle,
            hide_index=True,
            use_container_width=True,
        )


# ============================================================
# 15. LEGAJOS DIGITALES
# ============================================================

elif pagina == "Legajos digitales":

    cabecera(
        "Legajos digitales",
        "Información salarial del Excel y ficha "
        "de prueba con datos complementarios.",
    )

    nota()

    uid = st.selectbox(
        "Trabajador",
        base["ID"].tolist(),
        format_func=lambda x: etiqueta(
            x,
            personal,
        ),
    )

    r = personal.loc[
        personal["ID"] == uid
    ].iloc[0]

    a, b, c = st.columns(3)

    a.metric(
        "Salario mensual",
        ars(r["Mensual ARS"]),
    )

    b.metric(
        "Semanal equivalente",
        ars(r["Semanal equivalente ARS"]),
    )

    c.metric(
        "Antigüedad simulada",
        (
            f'{antiguedad_a_cierre(r["Ingreso DEMO"], anio)} '
            "años"
        ),
    )

    izq, der = st.columns(2)

    with izq:

        st.markdown(
            "#### Información original del Excel"
        )

        st.write(
            f'**Nombre:** {r["Nombre"]}'
        )

        st.write(
            f'**Rol:** {r["Rol"]}'
        )

        st.write(
            f'**Sucursal:** {r["Local"]}'
        )

        st.write(
            f'**Razón social:** {r["Razón social"]}'
        )

        st.write(
            f'**Salario mensual:** {ars(r["Mensual ARS"])}'
        )

        st.write(
            "**Semanal indicado en planilla:** "
            + ars(r["Semanal"])
        )

        st.write(
            "**Semanal real en planilla:** "
            + ars(r["semanal real"])
        )

    with der:

        st.markdown(
            "#### Información de demostración"
        )

        st.write(
            "**Fecha de ingreso (DEMO):** "
            + r["Ingreso DEMO"].strftime(
                "%d/%m/%Y"
            )
        )

        st.write(
            "**Nacimiento (DEMO):** "
            + r["Nacimiento DEMO"].strftime(
                "%d/%m/%Y"
            )
        )

        st.write(
            f'**Documento:** {r["Documento"]}'
        )

        st.write(
            f'**Domicilio:** {r["Domicilio"]}'
        )

        st.write(
            f'**ART:** {r["ART"]}'
        )

    cupo = cupo_orientativo(
        r["Ingreso DEMO"],
        anio,
    )

    usados = dias_aprobados(
        uid,
        anio,
    )

    st.progress(
        min(
            1.0,
            usados / max(1, cupo),
        ),
        text=(
            f"Vacaciones DEMO: "
            f"{usados}/{cupo} días corridos"
        ),
    )

    st.caption(
        "Las fechas y los saldos vacacionales "
        "son simulados. Deben verificarse según "
        "la fecha real de ingreso y el régimen "
        "laboral aplicable."
    )


# ============================================================
# 16. CALENDARIO Y TURNOS
# ============================================================

elif pagina == "Calendario y turnos":

    cabecera(
        "Calendario laboral argentino",
        "Nombres y roles de los empleados "
        "asignados a cada turno.",
    )

    nota()

    a, b, c, d = st.columns(
        [1, 1, 1.1, 1]
    )

    mes = a.selectbox(
        "Mes",
        range(1, 13),
        index=(
            HOY_AR.month - 1
            if anio == HOY_AR.year
            else 0
        ),
        format_func=lambda x: MESES[x - 1],
    )

    sede = b.selectbox(
        "Sucursal",
        sorted(
            base["Local"].unique()
        ),
    )

    muestra = base.loc[
        base["Local"] == sede
    ].copy()

    area = c.selectbox(
        "Área",
        ["Todas"]
        + sorted(
            muestra["Área"].unique().tolist()
        ),
    )

    vista = d.selectbox(
        "Turnos visibles",
        [
            "Ambos",
            "Mañana",
            "Tarde",
        ],
    )

    if area != "Todas":
        muestra = muestra.loc[
            muestra["Área"] == area
        ]

    festivos = feriados_argentina(anio)

    st.markdown(
        f"### {MESES[mes - 1]} {anio} · {sede}"
    )

    def dibujar_dia(dia):

        personal_dia = plan_turnos(
            muestra,
            dia,
        )

        turno_manana = personal_dia.loc[
            personal_dia["Turno"] == "Mañana"
        ]

        turno_tarde = personal_dia.loc[
            personal_dia["Turno"] == "Tarde"
        ]

        man = (
            turno_manana["Nombre"]
            + " · "
            + turno_manana["Rol"]
        ).tolist()

        tar = (
            turno_tarde["Nombre"]
            + " · "
            + turno_tarde["Rol"]
        ).tolist()

        nv = int(
            personal_dia["Turno"]
            .eq("Vacaciones")
            .sum()
        )

        if dia in festivos:
            css = "holiday"

        elif nv:
            css = "leave"

        else:
            css = ""

        contenido = ""

        if dia in festivos:
            contenido += (
                '<div class="calholiday">'
                + html.escape(
                    festivos[dia][0]
                )
                + "</div>"
            )

        if vista in (
            "Ambos",
            "Mañana",
        ):
            contenido += (
                '<div class="calshift">'
                f"☀ Mañana ({len(man)})"
                "</div>"
                + lista_resumida(man)
            )

        if vista in (
            "Ambos",
            "Tarde",
        ):
            contenido += (
                '<div class="calshift">'
                f"◕ Tarde ({len(tar)})"
                "</div>"
                + lista_resumida(tar)
            )

        if nv:
            contenido += (
                '<div class="calstatus">'
                f"Vacaciones: {nv}"
                "</div>"
            )

        return contenido, css

    calendario_mensual(
        anio,
        mes,
        dibujar_dia,
    )

    st.caption(
        "Se muestran hasta dos nombres por turno "
        "dentro de cada día. Los turnos son "
        "SIMULADOS. Un feriado no implica "
        "automáticamente el cierre del establecimiento."
    )

    # --------------------------------------------------------
    # CONSULTA POR DÍA
    # --------------------------------------------------------

    st.markdown(
        "### Consultar dotación completa de un día"
    )

    limite_dia = calendar.monthrange(
        anio,
        mes,
    )[1]

    dia_consulta = st.selectbox(
        "Día del mes",
        range(
            1,
            limite_dia + 1,
        ),
        index=min(
            10,
            limite_dia - 1,
        ),
    )

    fecha = date(
        anio,
        mes,
        dia_consulta,
    )

    tabla_dia = plan_turnos(
        muestra,
        fecha,
    )

    k1, k2, k3 = st.columns(3)

    k1.metric(
        "Turno mañana",
        int(
            tabla_dia["Turno"]
            .eq("Mañana")
            .sum()
        ),
    )

    k2.metric(
        "Turno tarde",
        int(
            tabla_dia["Turno"]
            .eq("Tarde")
            .sum()
        ),
    )

    k3.metric(
        "De vacaciones",
        int(
            tabla_dia["Turno"]
            .eq("Vacaciones")
            .sum()
        ),
    )

    st.dataframe(
        tabla_dia,
        hide_index=True,
        use_container_width=True,
    )

    # --------------------------------------------------------
    # FERIADOS
    # --------------------------------------------------------

    st.markdown(
        "### Feriados y días no laborables del mes"
    )

    especiales = [
        {
            "Fecha": d.strftime("%d/%m/%Y"),
            "Descripción": nombre,
            "Tipo": tipo,
        }
        for d, (nombre, tipo) in sorted(
            festivos.items()
        )
        if d.month == mes
    ]

    if especiales:
        st.dataframe(
            pd.DataFrame(especiales),
            hide_index=True,
            use_container_width=True,
        )

    else:
        st.info(
            "No se registran fechas especiales "
            "nacionales o provinciales cargadas "
            "para este mes."
        )

    st.caption(
        "Calendario de referencia: "
        "argentina.gob.ar/feriados. "
        "La fecha especial del 11/11/2026 "
        "corresponde a la Provincia de Buenos Aires."
    )


# ============================================================
# 17. VACACIONES Y SIMULADOR
# ============================================================

elif pagina == "Vacaciones":

    cabecera(
        "Vacaciones y simulador de reemplazos",
        "Evalúa el impacto de ausencias antes "
        "de registrar y aprobar una solicitud.",
    )

    nota()

    saldos = []

    for _, r in base.iterrows():

        cupo = cupo_orientativo(
            r["Ingreso DEMO"],
            anio,
        )

        usados = dias_aprobados(
            r["ID"],
            anio,
        )

        saldos.append(
            {
                "ID": r["ID"],
                "Trabajador": r["Nombre"],
                "Sucursal": r["Local"],
                "Rol": r["Rol"],
                "Días de referencia": cupo,
                "Aprobados": usados,
                "Saldo orientativo": max(
                    0,
                    cupo - usados,
                ),
            }
        )

    saldos_df = pd.DataFrame(saldos)

    a, b, c = st.columns(3)

    a.metric(
        "Días de referencia",
        int(
            saldos_df["Días de referencia"]
            .sum()
        ),
    )

    b.metric(
        "Aprobados (DEMO)",
        int(
            saldos_df["Aprobados"].sum()
        ),
    )

    c.metric(
        "Saldo orientativo",
        int(
            saldos_df["Saldo orientativo"]
            .sum()
        ),
    )

    t1, t2, t3 = st.tabs(
        [
            "Resumen de saldos",
            "Simular vacaciones",
            "Solicitudes y aprobaciones",
        ]
    )

    # --------------------------------------------------------
    # SALDOS
    # --------------------------------------------------------

    with t1:

        st.dataframe(
            saldos_df,
            hide_index=True,
            use_container_width=True,
        )

        st.caption(
            "El saldo es orientativo porque las "
            "fechas de ingreso no constan "
            "en la planilla."
        )

    # --------------------------------------------------------
    # SIMULACIÓN
    # --------------------------------------------------------

    with t2:

        uid = st.selectbox(
            "Trabajador que solicita vacaciones",
            base["ID"].tolist(),
            format_func=lambda x: etiqueta(
                x,
                personal,
            ),
            key="sim_uid",
        )

        c1, c2, c3 = st.columns(3)

        inicio = c1.date_input(
            "Inicio",
            value=date(anio, 11, 16),
            key=f"sim_inicio_{anio}",
        )

        fin = c2.date_input(
            "Fin",
            value=date(anio, 11, 29),
            key=f"sim_fin_{anio}",
        )

        minimo = c3.number_input(
            "Mínimo por área y turno (DEMO)",
            min_value=1,
            max_value=15,
            value=1,
        )

        otras = st.checkbox(
            "Considerar personal de otras sucursales",
            value=True,
        )

        if fin < inicio:

            st.error(
                "La fecha de fin no puede ser "
                "anterior al inicio."
            )

        elif (
            inicio.year != anio
            or fin.year != anio
        ):

            st.error(
                "La simulación debe quedar dentro "
                "del año seleccionado."
            )

        elif (fin - inicio).days > 60:

            st.error(
                "El máximo de esta demo es "
                "61 días corridos por solicitud."
            )

        else:

            impacto = estudiar_cobertura(
                personal,
                uid,
                inicio,
                fin,
                minimo,
            )

            riesgo = int(
                impacto["Faltantes"]
                .gt(0)
                .sum()
            )

            a, b, c = st.columns(3)

            a.metric(
                "Días solicitados",
                (fin - inicio).days + 1,
            )

            b.metric(
                "Turnos a cubrir",
                len(impacto),
            )

            c.metric(
                "Días con déficit",
                riesgo,
            )

            if riesgo:

                st.warning(
                    f"La simulación detecta {riesgo} "
                    "turnos con déficit para el área "
                    "y sucursal del trabajador, según "
                    "el mínimo configurado."
                )

            else:

                st.success(
                    "No se detectan déficits bajo "
                    "el mínimo configurado."
                )

            st.markdown(
                "### Cobertura por día afectado"
            )

            st.dataframe(
                impacto,
                use_container_width=True,
                hide_index=True,
            )

            # ------------------------------------------------
            # CALENDARIO DE IMPACTO
            # ------------------------------------------------

            st.markdown(
                "### Calendario de impacto"
            )

            meses_involucrados = sorted(
                {
                    (
                        inicio
                        + timedelta(days=i)
                    ).replace(day=1)

                    for i in range(
                        (fin - inicio).days + 1
                    )
                }
            )

            mes_pintar = st.selectbox(
                "Ver mes del período",
                meses_involucrados,
                format_func=lambda d: (
                    f"{MESES[d.month - 1]} {d.year}"
                ),
            )

            riesgos_por_dia = {
                r["Fecha"]: int(r["Faltantes"])

                for _, r in impacto.iterrows()
            }

            festivos = feriados_argentina(
                mes_pintar.year
            )

            def pintar_impacto(d):

                dentro = (
                    inicio <= d <= fin
                )

                faltantes = (
                    riesgos_por_dia.get(d, 0)
                )

                if dentro:

                    if faltantes:
                        return (
                            '<div class="calstatus">'
                            f"⚠ Faltan {faltantes}"
                            "</div>",
                            "risk",
                        )

                    return (
                        '<div class="calshift">'
                        "✓ Sin déficit detectado"
                        "</div>",
                        "leave",
                    )

                if d in festivos:
                    return (
                        '<div class="calholiday">'
                        + html.escape(
                            festivos[d][0]
                        )
                        + "</div>",
                        "holiday",
                    )

                return "", ""

            calendario_mensual(
                mes_pintar.year,
                mes_pintar.month,
                pintar_impacto,
            )

            st.caption(
                "Azul: sin déficit. "
                "Naranja: falta personal. "
                "La simulación no asigna "
                "reemplazos automáticamente."
            )

            # ------------------------------------------------
            # REEMPLAZOS
            # ------------------------------------------------

            st.markdown(
                "### Candidatos a reemplazar"
            )

            candidatos = candidatos_reemplazo(
                personal,
                uid,
                impacto,
                otras,
            )

            if candidatos.empty:

                st.info(
                    "No se encontraron candidatos "
                    "adicionales de la misma área "
                    "para los turnos con déficit."
                )

            else:

                st.dataframe(
                    candidatos.drop(
                        columns=[
                            "ID",
                            "Prioridad",
                        ]
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

                st.caption(
                    "La propuesta puede requerir "
                    "cambiar turno, descanso o sucursal. "
                    "No es una disponibilidad confirmada."
                )

            trabajador = personal.loc[
                personal["ID"] == uid
            ].iloc[0]

            saldo = (
                cupo_orientativo(
                    trabajador["Ingreso DEMO"],
                    anio,
                )
                - dias_aprobados(
                    uid,
                    anio,
                )
            )

            dias_solicitados = (
                (fin - inicio).days + 1
            )

            if dias_solicitados > saldo:

                st.warning(
                    "El período excede el saldo "
                    f"orientativo de {saldo} días."
                )

            if st.button(
                "Registrar solicitud para evaluación",
                type="primary",
            ):

                if solapada(
                    uid,
                    inicio,
                    fin,
                ):

                    st.error(
                        "El trabajador ya tiene una "
                        "solicitud aprobada o pendiente "
                        "que se superpone."
                    )

                else:

                    numero = max(
                        (
                            s["N"]
                            for s in st.session_state.rrhh_vacaciones
                        ),
                        default=0,
                    ) + 1

                    st.session_state.rrhh_vacaciones.append(
                        {
                            "N": numero,
                            "ID": uid,
                            "Inicio": inicio,
                            "Fin": fin,
                            "Estado": "Pendiente",
                            "Origen": "DEMO",
                        }
                    )

                    st.success(
                        "Solicitud creada. Ahora podrá "
                        "revisarse en la pestaña "
                        "de aprobaciones."
                    )

                    st.rerun()

    # --------------------------------------------------------
    # APROBACIONES
    # --------------------------------------------------------

    with t3:

        listado = [
            s

            for s in st.session_state.rrhh_vacaciones

            if (
                s["ID"] in set(base["ID"])
                and (
                    s["Inicio"].year == anio
                    or s["Fin"].year == anio
                )
            )
        ]

        if listado:

            tabla_solicitudes = pd.DataFrame(
                [
                    {
                        "N°": s["N"],
                        "Empleado": etiqueta(
                            s["ID"],
                            personal,
                        ),
                        "Inicio": s["Inicio"].strftime(
                            "%d/%m/%Y"
                        ),
                        "Fin": s["Fin"].strftime(
                            "%d/%m/%Y"
                        ),
                        "Días": (
                            s["Fin"]
                            - s["Inicio"]
                        ).days + 1,
                        "Estado": s["Estado"],
                        "Origen": s["Origen"],
                    }

                    for s in listado
                ]
            )

            st.dataframe(
                tabla_solicitudes,
                hide_index=True,
                use_container_width=True,
            )

        pendientes = [
            s for s in listado
            if s["Estado"] == "Pendiente"
        ]

        if not pendientes:

            st.success(
                "No existen solicitudes pendientes "
                "para la selección actual."
            )

        else:

            numero = st.selectbox(
                "Solicitud a revisar",
                [
                    s["N"]
                    for s in pendientes
                ],
                format_func=lambda n: (
                    f"#{n} · "
                    + etiqueta(
                        next(
                            s["ID"]
                            for s in pendientes
                            if s["N"] == n
                        ),
                        personal,
                    )
                ),
            )

            solicitud = next(
                s for s in pendientes
                if s["N"] == numero
            )

            min_apr = st.number_input(
                "Mínimo requerido por área y turno",
                min_value=1,
                max_value=15,
                value=1,
                key="min_apr",
            )

            evaluacion = estudiar_cobertura(
                personal,
                solicitud["ID"],
                solicitud["Inicio"],
                solicitud["Fin"],
                min_apr,
            )

            dias_riesgo = int(
                evaluacion["Faltantes"]
                .gt(0)
                .sum()
            )

            st.metric(
                "Turnos con déficit",
                dias_riesgo,
            )

            if dias_riesgo:

                st.warning(
                    "La aprobación no implica que "
                    "exista cobertura confirmada. "
                    "Se deben revisar los reemplazos."
                )

            candidatos = candidatos_reemplazo(
                personal,
                solicitud["ID"],
                evaluacion,
                True,
            )

            propuestas = [
                "Sin reemplazo propuesto"
            ] + candidatos["ID"].tolist()

            reemplazo = st.selectbox(
                "Reemplazo sugerido "
                "(pendiente de confirmar)",
                propuestas,
                format_func=lambda x: (
                    etiqueta(x, personal)
                    if x != "Sin reemplazo propuesto"
                    else x
                ),
            )

            fila = personal.loc[
                personal["ID"] == solicitud["ID"]
            ].iloc[0]

            saldo = (
                cupo_orientativo(
                    fila["Ingreso DEMO"],
                    anio,
                )
                - dias_aprobados(
                    solicitud["ID"],
                    anio,
                )
            )

            dias = (
                solicitud["Fin"]
                - solicitud["Inicio"]
            ).days + 1

            ok_saldo = dias <= saldo

            if not ok_saldo:

                st.error(
                    "Saldo orientativo insuficiente: "
                    f"{saldo} días frente a "
                    f"{dias} solicitados."
                )

            confirmar = st.checkbox(
                "He revisado el impacto, las condiciones "
                "del reemplazo y los descansos aplicables.",
                key=f"confirmacion_{numero}",
            )

            c_apro, c_rech = st.columns(2)

            if c_apro.button(
                "Aprobar (solo DEMO)",
                type="primary",
                disabled=not (
                    confirmar
                    and ok_saldo
                ),
                use_container_width=True,
            ):

                if solapada(
                    solicitud["ID"],
                    solicitud["Inicio"],
                    solicitud["Fin"],
                    except_id=numero,
                ):

                    st.error(
                        "Existen solicitudes superpuestas "
                        "para este trabajador."
                    )

                else:

                    solicitud["Estado"] = "Aprobada"

                    st.session_state.rrhh_reemplazos[
                        numero
                    ] = {
                        "Candidato": reemplazo,
                        "Estado": (
                            "Propuesta, NO asignación confirmada"
                        ),
                        "Turnos con déficit": dias_riesgo,
                    }

                    st.success(
                        "Vacaciones aprobadas para la demo. "
                        "La asignación del reemplazo "
                        "queda pendiente."
                    )

                    st.rerun()

            if c_rech.button(
                "Rechazar solicitud",
                use_container_width=True,
            ):

                solicitud["Estado"] = "Rechazada"
                st.rerun()


# ============================================================
# 18. REMUNERACIONES
# ============================================================

elif pagina == "Remuneraciones":

    cabecera(
        "Remuneraciones y anticipos",
        "Presupuesto calculado desde la columna "
        "'Salario mensual', sin multiplicar "
        "por las semanas del mes.",
    )

    st.info(
        "La base salarial proviene exclusivamente "
        "de 'Salario mensual'. "
        "'Semanal' y 'semanal real' se muestran "
        "para comparación y no determinan "
        "el presupuesto."
    )

    base_mes = float(
        base["Mensual ARS"].sum()
    )

    jun = float(
        base["Aguinaldo junio"]
        .fillna(0)
        .sum()
    )

    dic = float(
        base["Aguinaldo diciembre"]
        .fillna(0)
        .sum()
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "Mensual base",
        millones(base_mes),
    )

    b.metric(
        "Anual base (12 meses)",
        millones(base_mes * 12),
    )

    c.metric(
        "SAC junio (Excel)",
        millones(jun),
    )

    d.metric(
        "SAC diciembre (Excel)",
        millones(dic),
    )

    serie = pd.DataFrame(
        {
            "Mes": MESES,
            "Salario base": [base_mes] * 12,
            "SAC de planilla": [
                jun if i == 6
                else dic if i == 12
                else 0

                for i in range(1, 13)
            ],
        }
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=serie["Mes"],
            y=serie["Salario base"],
            name="Salario mensual",
            marker_color=AZUL,
        )
    )

    fig.add_trace(
        go.Bar(
            x=serie["Mes"],
            y=serie["SAC de planilla"],
            name="SAC",
            marker_color=NARANJA,
        )
    )

    fig.update_layout(
        title=(
            "Proyección salarial base "
            "+ aguinaldo (ARS)"
        ),
        barmode="stack",
        yaxis_title="ARS",
    )

    st.plotly_chart(
        grafico(fig),
        use_container_width=True,
    )

    nomina = base[
        [
            "ID",
            "Nombre",
            "Local",
            "Rol",
            "Mensual ARS",
            "Semanal equivalente ARS",
            "Semanal",
            "semanal real",
            "Anual base ARS",
        ]
    ].copy()

    st.markdown(
        "### Detalle de salarios"
    )

    st.dataframe(
        nomina,
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------
    # ANTICIPOS
    # --------------------------------------------------------

    st.markdown(
        "### Registrar anticipo — simulación"
    )

    with st.form("form_anticipo"):

        uid = st.selectbox(
            "Empleado",
            base["ID"].tolist(),
            format_func=lambda x: etiqueta(
                x,
                personal,
            ),
        )

        monto = st.number_input(
            "Anticipo ARS",
            min_value=0.0,
            value=50000.0,
            step=10000.0,
        )

        fecha = st.date_input(
            "Fecha",
            value=HOY_AR,
        )

        limite = st.slider(
            "Límite DEMO sobre salario mensual",
            10,
            50,
            30,
            5,
        )

        enviado = st.form_submit_button(
            "Registrar anticipo",
            type="primary",
        )

    if enviado:

        salario = float(
            personal.loc[
                personal["ID"] == uid,
                "Mensual ARS",
            ].iloc[0]
        )

        acumulado = sum(
            a["Monto"]

            for a in st.session_state.rrhh_anticipos

            if (
                a["ID"] == uid
                and (
                    a["Fecha"].year,
                    a["Fecha"].month,
                ) == (
                    fecha.year,
                    fecha.month,
                )
            )
        )

        if monto <= 0:

            st.error(
                "El importe debe ser mayor a cero."
            )

        elif (
            acumulado + monto
            > salario * limite / 100
        ):

            st.error(
                "El anticipo supera el límite "
                "administrativo simulado."
            )

        else:

            st.session_state.rrhh_anticipos.append(
                {
                    "ID": uid,
                    "Monto": monto,
                    "Fecha": fecha,
                    "Origen": "DEMO",
                }
            )

            st.success(
                "Anticipo incorporado en la sesión "
                "de prueba. El salario base "
                "permanece intacto."
            )

            st.rerun()

    filas = [
        {
            "Empleado": etiqueta(
                a["ID"],
                personal,
            ),
            "Fecha": a["Fecha"].strftime(
                "%d/%m/%Y"
            ),
            "Anticipo ARS": a["Monto"],
            "Origen": a["Origen"],
        }

        for a in st.session_state.rrhh_anticipos

        if a["ID"] in set(base["ID"])
    ]

    if filas:

        st.dataframe(
            pd.DataFrame(filas),
            hide_index=True,
            use_container_width=True,
        )

    st.caption(
        "No incluye cargas sociales ni revisiones "
        "salariales futuras. La lógica de anticipos "
        "es ilustrativa y debe validarse "
        "con asesoría laboral."
    )


# ============================================================
# 19. CENTRO DE ALERTAS
# ============================================================

elif pagina == "Centro de alertas":

    cabecera(
        "Centro de alertas",
        "Solicitudes pendientes y vista previa "
        "de notificaciones administrativas por WhatsApp.",
    )

    nota()

    alertas = []

    for s in st.session_state.rrhh_vacaciones:

        if (
            s["Estado"] == "Pendiente"
            and s["ID"] in set(base["ID"])
        ):

            alertas.append(
                {
                    "Clave": f'VAC-{s["N"]}',
                    "Tipo": "Vacaciones",
                    "Referencia": etiqueta(
                        s["ID"],
                        personal,
                    ),
                    "Descripción": (
                        f'Solicitud del '
                        f'{s["Inicio"]:%d/%m} '
                        f'al {s["Fin"]:%d/%m} '
                        "pendiente de revisión"
                    ),
                    "Prioridad": "Media",
                }
            )

    alertas.append(
        {
            "Clave": "ART-REV",
            "Tipo": "Documentación",
            "Referencia": "Administración",
            "Descripción": (
                "Cargar ART y legajos reales: "
                "todavía no existen en el Excel"
            ),
            "Prioridad": "Alta",
        }
    )

    visibles = [
        a for a in alertas

        if (
            a["Clave"]
            not in st.session_state.rrhh_alertas_vistas
        )
    ]

    a, b, c = st.columns(3)

    a.metric(
        "Alertas abiertas",
        len(visibles),
    )

    b.metric(
        "Solicitudes pendientes",
        sum(
            a["Tipo"] == "Vacaciones"
            for a in visibles
        ),
    )

    c.metric(
        "Marcadas revisadas",
        len(
            st.session_state.rrhh_alertas_vistas
        ),
    )

    if visibles:

        st.dataframe(
            pd.DataFrame(visibles).drop(
                columns="Clave"
            ),
            hide_index=True,
            use_container_width=True,
        )

        i = st.selectbox(
            "Revisar alerta",
            range(len(visibles)),
            format_func=lambda j: (
                visibles[j]["Tipo"]
                + " · "
                + visibles[j]["Descripción"]
            ),
        )

        alerta = visibles[i]

        mensaje = (
            "EL LUCHADOR | RR. HH.\n"
            f'Tipo: {alerta["Tipo"]}\n'
            f'Referencia: {alerta["Referencia"]}\n'
            f'Situación: {alerta["Descripción"]}\n'
            "Ingresá a la plataforma para revisar."
        )

        st.code(
            mensaje,
            language=None,
        )

        if st.button(
            "Marcar como revisada",
            type="primary",
        ):

            st.session_state.rrhh_alertas_vistas.add(
                alerta["Clave"]
            )

            st.rerun()

    else:

        st.success(
            "No hay alertas pendientes de revisión."
        )

    st.warning(
        "WhatsApp está en modo DEMO. "
        "No se envían mensajes reales. "
        "Se requiere integrar WhatsApp Business "
        "Platform, disponer de credenciales, "
        "plantillas aprobadas y permisos."
    )


# ============================================================
# 20. PIE DE PÁGINA
# ============================================================

st.divider()

st.caption(
    "EL LUCHADOR · Demo local de Recursos Humanos "
    "· Datos salariales del Excel "
    "· Planificación sintética "
    "· Valores en ARS"
)
