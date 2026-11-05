# -*- coding: utf-8 -*-
"""
Dashboard de Seguimiento Académico UCC  (v3 – flujo guiado)
===========================================================
Instalación:
    pip install streamlit pandas plotly openpyxl
Ejecución:
    streamlit run dashboard_academico.py

Flujo: 1) modalidad -> 2) ciclo -> 3) programa -> informe.

Lógica de cumplimiento (según la Lista de seguimiento):
- Hay 3 revisiones (Fase 1, 2 y 3) y CADA UNA vale 100 %.
- El cumplimiento acumulado es el PROMEDIO de las fases ya revisadas del aula
  (ej.: F1 = 100 % y F2 = 50 %  ->  75 %).
- Cada criterio pertenece a una fase (M1 -> F1, M2 -> F2, M3 -> F3; "Calificó/Retroalimentó Mx"
  se evalúa en la fase siguiente, salvo M3).
- Una fase está REVISADA en un aula si tiene criterios diligenciados; una fase sin revisar
  no cuenta como 0 %.
- Tema claro: el script crea .streamlit/config.toml si no existe (reinicie la app la primera vez).
"""
from __future__ import annotations

import re
import unicodedata
import warnings
from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

TEMA_CLARO = """[theme]
base = "light"
primaryColor = "#1d4ed8"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f1f5f9"
textColor = "#1f2937"
"""
try:  # tema claro institucional (se aplica al reiniciar la app)
    _cfg = Path.cwd() / ".streamlit" / "config.toml"
    if not _cfg.exists():
        _cfg.parent.mkdir(exist_ok=True)
        _cfg.write_text(TEMA_CLARO, encoding="utf-8")
except Exception:
    pass

st.set_page_config(page_title="Seguimiento Académico UCC", page_icon="🎓",
                   layout="wide", initial_sidebar_state="expanded")

# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
ARCHIVO_POR_DEFECTO = "seguimiento_academico.xlsx"   # se usa si no se sube nada
TODOS = "Todos los programas"
NOMBRES_MOD = {"Extendida": "Aulas extendidas (presenciales)", "Virtual": "Aulas virtuales"}
COLORES_ESTADO = {"Si": "#16a34a", "Incompleto": "#f59e0b", "No": "#dc2626"}
MAPA_ESTADO = {"si": "Si", "s": "Si", "no": "No", "incompleto": "Incompleto",
               "incompleta": "Incompleto", "parcial": "Incompleto", "intermedio": "Incompleto"}
ALIAS = {  # nombres normalizados (minúsculas, sin tildes); coincidencia exacta
    "id": ["id curso", "id del curso", "nrc", "crn"],
    "curso": ["descripcion", "nombre del curso", "nombre curso", "curso", "materia", "asignatura"],
    "profesor": ["nombre completo", "nombre del profesor", "nombre profesor", "profesor", "docente"],
    "programa": ["org acad", "organizacion academica", "programa academico", "programa", "facultad", "escuela"],
    "seccion": ["seccion", "grupo"],
    "asignacion": ["asignacion profesor", "asignacion"],
}
# Niveles de cumplimiento (cada aula cae en uno solo). clave -> (icono, título, color)
NIVELES = {  # clave -> (icono, título, color de acento, color de fondo)
    "completas": ("✅", "Aulas completas", "#15803d", "#f0fdf4"),
    "avance": ("🟡", "Aulas en avance", "#d97706", "#fffbeb"),
    "criticas": ("🔴", "Aulas críticas", "#dc2626", "#fef2f2"),
    "vacias": ("⚫", "Aulas vacías", "#64748b", "#f1f5f9"),
}
TXT_NIVEL = {"completas": "Completa (100%)", "avance": "En avance (50% a 99%)",
             "criticas": "Crítica (menos de 50%)", "vacias": "Vacía (0%)"}
# Nombres claros para las columnas de las tablas
COL = {"Profesor": "Profesor", "Curso": "Curso", "Seccion": "Sección", "Programa": "Programa",
       "Ciclo": "Ciclo", "Modalidad": "Modalidad", "ID_Curso": "ID del curso",
       "Cumplimiento": "Cumplimiento (%)", "Pendientes": "Qué falta por corregir",
       "Criterios_No": "Criterios NO cumplidos (cantidad)",
       "Criterios_Inc": "Criterios incompletos (cantidad)",
       "Observacion": "Observación del revisor", "Ultima_Revision": "Fecha de última revisión",
       "Hasta fase": "Revisada hasta la fase", "Acum_100": "Promedio de las fases revisadas (%)",
       "Nivel_txt": "Nivel de cumplimiento"}

CSS = """
<style>
@keyframes fadeUp {from {opacity:0; transform:translateY(14px);} to {opacity:1; transform:none;}}
.hero {background:linear-gradient(120deg,#1e3a8a,#2563eb); padding:1.6rem 2rem; border-radius:18px;
       margin-bottom:1.2rem; animation:fadeUp .6s ease; box-shadow:0 4px 14px rgba(30,58,138,.25);}
.hero h1 {margin:0; font-size:1.9rem; color:#ffffff !important;}
.hero p {margin:.4rem 0 0; color:#dbeafe !important; font-size:1.05rem;}
.kpi {background:var(--bg); border:1px solid rgba(15,23,42,.08); border-top:6px solid var(--c);
      border-radius:16px; padding:1rem 1.2rem; box-shadow:0 2px 10px rgba(15,23,42,.08);
      animation:fadeUp .6s ease; min-height:190px;}
.kpi.act {box-shadow:0 0 0 3px var(--c), 0 6px 16px rgba(15,23,42,.18);}
.kpi-t {font-size:1.05rem; font-weight:700; color:#111827 !important;}
.kpi-d {font-size:.82rem; color:#4b5563 !important; min-height:2.4em;}
.kpi-v {font-size:2.6rem; font-weight:800; line-height:1.2;}
.kpi-s {font-size:.8rem; color:#6b7280 !important;}
.pill {display:inline-block; padding:.3rem .8rem; margin:.15rem .3rem .15rem 0; border-radius:999px;
       background:#eff6ff; border:1px solid #bfdbfe; color:#1e293b !important; font-size:.9rem;}
[data-testid="stFileUploader"] section {border:2px dashed #2563eb; border-radius:14px;
      background:rgba(37,99,235,.05);}
.block-container {padding-top:1.6rem;}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def norm(txt) -> str:
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", t).strip()


def buscar_col(cols, alias, usados=()):
    disp = [(c, norm(c)) for c in cols if c not in usados]
    for a in alias:
        for c, n in disp:
            if n == a:
                return c
    return None


def partes_nombre(cols, usados=()):
    """Nombre, 2º nombre, apellido y 2º apellido (los '2º' pueden escribirse de varias formas)."""
    reglas = (lambda n: n == "nombre", lambda n: n.startswith("2") and n.endswith("nombre"),
              lambda n: n == "apellido", lambda n: n.startswith("2") and n.endswith("apellido"))
    out = [next((c for c in cols if c not in usados and r(norm(c))), None) for r in reglas]
    return [c for c in out if c]


def estado(v):
    return None if v is None or pd.isna(v) else MAPA_ESTADO.get(norm(v))


def fase_de_criterio(nombre: str) -> int:
    """M1->F1, M2->F2, M3->F3; 'Calificó/Retroalimentó Mx' se evalúa en la fase siguiente (máx. F3)."""
    n = norm(nombre)
    m = re.search(r"(?<![a-z])m\s*([123])(?![0-9])", n)
    if not m:
        return 1
    k = int(m.group(1))
    return min(k + 1, 3) if n.startswith(("calific", "retroalim")) else k


def a_pct(s: pd.Series) -> pd.Series:
    x = pd.to_numeric(s.astype(str).str.replace("%", "").str.replace(",", ".").str.strip(), errors="coerce")
    return x * 100 if x.notna().any() and x.max() <= 1.0 else x


def ciclo_txt(s: pd.Series) -> pd.Series:
    n = pd.to_numeric(s, errors="coerce")
    return n.round().astype("Int64").astype("string").where(n.notna(), s.astype("string").str.strip())


def leer_hoja(xls: pd.ExcelFile, nombre: str):
    """Lee la hoja detectando la fila de encabezados (la fila con más textos)."""
    muestra = xls.parse(nombre, header=None, nrows=15)
    conteo = muestra.apply(lambda r: sum(isinstance(v, str) for v in r), axis=1)
    h = int(conteo.idxmax()) if len(conteo) else 0
    df = xls.parse(nombre, header=h)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def quitar_fase(c: str) -> str:
    return re.sub(r"\s*\[F\d\]$", "", c)


def pendientes(row, cs):
    no = [quitar_fase(c) for c in cs if row[c] == "No"]
    inc = [quitar_fase(c) for c in cs if row[c] == "Incompleto"]
    return " | ".join(p for p in (("NO cumple: " + ", ".join(no)) if no else "",
                                  ("INCOMPLETO: " + ", ".join(inc)) if inc else "") if p)


# ----------------------------------------------------------------------------
# Carga y procesamiento (cacheado)
# ----------------------------------------------------------------------------
def procesar_hoja(raw: pd.DataFrame, modalidad: str, canon: dict):
    df = raw.dropna(how="all").copy()
    cols = list(df.columns)
    usados = {cols[0]}  # Columna A = ciclo

    def tomar(alias):
        c = buscar_col(cols, alias, usados)
        if c:
            usados.add(c)
        return c

    c_id, c_cur, c_pro = tomar(ALIAS["id"]), tomar(ALIAS["curso"]), tomar(ALIAS["profesor"])
    c_prg, c_sec, c_asg = tomar(ALIAS["programa"]), tomar(ALIAS["seccion"]), tomar(ALIAS["asignacion"])
    partes = [] if c_pro else partes_nombre(cols, usados)
    usados.update(partes)

    # Criterios = columnas con mayoría de valores Si / No / Incompleto (se detectan por contenido)
    crit = []
    for c in cols:
        if c in usados or df[c].notna().sum() == 0:
            continue
        est = df[c].map(estado)
        if est.notna().sum() / df[c].notna().sum() >= 0.6:
            df[c] = est
            crit.append(c)
    if not crit:
        raise ValueError("no se detectaron columnas de criterios (Si/No/Incompleto)")

    # Nombre canónico del criterio (unifica mayúsculas/typos entre hojas) + fase a la que pertenece
    ren, fase_c = {}, {}
    for c in crit:
        key = norm(c).replace("calificaco", "califico")
        nombre = canon.setdefault(key, re.sub(r"\s+", " ", c).strip().replace("Calificacó", "Calificó"))
        k = fase_de_criterio(c)
        ren[c] = f"{nombre} [F{k}]"
        fase_c[ren[c]] = k

    # % reportados, observaciones y fechas por fase
    rep, obs, fec, rep_total = {}, {}, {}, None
    for c in cols:
        if c in ren or c in usados:
            continue
        n = norm(c)
        m = re.search(r"(?<![a-z])(?:fase|f|m)\s*([123])(?![0-9])", n)
        if "%" in n and m:
            rep[int(m.group(1))] = a_pct(df[c])
        elif "%" in n and "cumplimiento" in n:
            rep_total = a_pct(df[c])
        elif m and "observ" in n:
            obs.setdefault(int(m.group(1)), df[c])
        elif m and "fecha" in n:
            fec.setdefault(int(m.group(1)), pd.to_datetime(df[c], errors="coerce"))

    df = df.rename(columns=ren)
    crit = list(ren.values())

    def t(c, defecto):
        if not c:
            return pd.Series(defecto, index=df.index, dtype="string")
        return df[c].astype("string").str.strip().fillna(defecto)

    if c_pro:
        prof = df[c_pro].astype("string").str.strip().fillna("")
    elif partes:
        prof = df[partes].astype("string").fillna("").apply(lambda r: " ".join(x.strip() for x in r if x.strip()), axis=1)
    else:
        prof = pd.Series("", index=df.index)
    prof = prof.astype(str).str.title().replace("", "Sin asignar")

    out = pd.DataFrame({
        "Modalidad": modalidad, "Ciclo": ciclo_txt(df[cols[0]]),
        "Curso": t(c_cur or c_id, "Sin nombre"), "ID_Curso": t(c_id, ""), "Seccion": t(c_sec, ""),
        "Profesor": prof, "Programa": t(c_prg, "Sin especificar"), "Asignacion": t(c_asg, "—"),
    }, index=df.index)

    nan = float("nan")
    for k in (1, 2, 3):
        cs = [c for c in crit if fase_c[c] == k]
        sub = df[cs]
        rev = sub.notna().any(axis=1) if cs else pd.Series(False, index=df.index)
        si, no, inc = ((sub == v).sum(axis=1) for v in ("Si", "No", "Incompleto"))
        out[f"Rev_F{k}"] = rev
        out[f"No_F{k}"], out[f"Inc_F{k}"] = no, inc
        fr = ((si + 0.5 * inc) / len(cs)).where(rev) if cs else pd.Series(nan, index=df.index)
        out[f"Frac_F{k}"] = fr
        rk = rep.get(k, pd.Series(nan, index=df.index))
        if k in rep and fr.max() >= 0.999 and 0 < rk.max() < 99:
            rk = rk * 100 / rk.max()  # Excel antiguo (F1=50, F2=20, F3=30): se lleva cada fase a escala 0-100
        out[f"Rep_F{k}"] = rk
        out[f"Obs_F{k}"] = obs[k].astype("string").str.strip().replace("", pd.NA) if k in obs else pd.NA
        out[f"Fecha_F{k}"] = fec.get(k, pd.NaT)
        out[f"Pend_F{k}"] = sub.apply(lambda r, cs=cs: pendientes(r, cs), axis=1) if cs else ""
    out["Rep_Total"] = rep_total if rep_total is not None else nan
    out = pd.concat([out, df[crit]], axis=1)

    out = out[out["Ciclo"].notna() & ~((out["Curso"] == "Sin nombre") & (out["ID_Curso"] == ""))]
    diag = {"Hoja": modalidad, "Filas": len(out), "Curso": c_cur, "Profesor": c_pro or partes,
            "Programa": c_prg, "% reportados": [f"F{k}" for k in rep],
            "Criterios por fase": {f"F{k}": sum(v == k for v in fase_c.values()) for k in (1, 2, 3)}}
    return out.reset_index(drop=True), crit, fase_c, diag


@st.cache_data(show_spinner="Procesando archivo…")
def cargar_datos(data: bytes):
    xls = pd.ExcelFile(BytesIO(data))
    hojas = {norm(n): n for n in xls.sheet_names}
    buscar = lambda clave: next((v for k, v in hojas.items() if clave in k), None)
    avisos, diag, partes, crit_mod, fase_crit, canon = [], [], [], {}, {}, {}
    for clave, modalidad in (("virtual", "Virtual"), ("extendida", "Extendida")):
        nombre = buscar(clave)
        if not nombre:
            avisos.append(f"No se encontró una hoja de aulas {clave}s.")
            continue
        try:
            d, crit, fc, dg = procesar_hoja(leer_hoja(xls, nombre), modalidad, canon)
            partes.append(d)
            crit_mod[modalidad] = crit
            fase_crit.update(fc)
            diag.append(dg)
        except Exception as e:
            avisos.append(f"Hoja '{nombre}': {e}.")
    if not partes:
        return pd.DataFrame(), {"avisos": avisos, "diag": diag, "criterios": {}, "fase_crit": {}}
    cursos = pd.concat(partes, ignore_index=True)
    ref = cursos["Seccion"].where(cursos["Seccion"] != "", cursos["ID_Curso"])
    cursos["Etiqueta"] = cursos["Curso"].str.slice(0, 34) + " (" + ref + ") · " + cursos["Ciclo"]
    cursos["Etiqueta"] += cursos.groupby("Etiqueta").cumcount().map(lambda i: "" if i == 0 else f" #{i + 1}")
    return cursos, {"avisos": avisos, "diag": diag, "criterios": crit_mod, "fase_crit": fase_crit}


def calcular(base: pd.DataFrame, usar_excel: bool, fase: int) -> pd.DataFrame:
    """Cada fase vale 100 %. fase=1..3 -> cumplimiento de esa fase; fase=0 -> PROMEDIO de las fases revisadas."""
    df = base.copy()
    for k in (1, 2, 3):
        calc = df[f"Frac_F{k}"] * 100
        pct = df[f"Rep_F{k}"].fillna(calc) if usar_excel else calc
        df[f"Pct_F{k}"] = pct.clip(0, 100).where(df[f"Rev_F{k}"])  # fase sin revisar -> NaN (no es un 0 real)
    fases = [1, 2, 3] if fase == 0 else [fase]
    df["Revisado"] = df[[f"Rev_F{k}" for k in fases]].any(axis=1)
    df["Cumplimiento"] = df[[f"Pct_F{k}" for k in fases]].mean(axis=1).where(df["Revisado"]).round(1)
    df["Criterios_No"] = sum(df[f"No_F{k}"] for k in fases)
    df["Criterios_Inc"] = sum(df[f"Inc_F{k}"] for k in fases)
    if len(fases) == 1:
        df["Pendientes"] = df[f"Pend_F{fases[0]}"]
    else:
        df["Pendientes"] = df.apply(lambda r: " ‖ ".join(f"Fase {k}: {r[f'Pend_F{k}']}" for k in fases
                                                         if r[f"Pend_F{k}"]), axis=1)
    vig = pd.Series(0, index=df.index)
    for k in (1, 2, 3):
        vig = vig.where(~df[f"Rev_F{k}"], k)
    df["Hasta fase"] = vig.map({0: "Sin revisar", 1: "Fase 1", 2: "Fase 2", 3: "Fase 3"})
    df["Acum_100"] = df[["Pct_F1", "Pct_F2", "Pct_F3"]].mean(axis=1).round(1)  # promedio de fases revisadas
    obs = pd.Series(pd.NA, index=df.index, dtype="object")
    for k in (1, 2, 3):
        ref = (vig == k) if fase == 0 else pd.Series(fase == k, index=df.index)
        obs = obs.where(~ref, df[f"Obs_F{k}"].astype(object))
    df["Observacion"] = obs
    df["Ultima_Revision"] = df[[f"Fecha_F{k}" for k in fases]].max(axis=1)
    c = df["Cumplimiento"]
    df["Nivel"] = np.select([c >= 99.95, c >= 50, c > 0], ["completas", "avance", "criticas"], default="vacias")
    df["Nivel_txt"] = df["Nivel"].map(TXT_NIVEL)
    return df


def estado_fases(d: pd.DataFrame):
    """Estado de cada fase en el conjunto de aulas elegido + fase actual revisada."""
    n = len(d)
    cov = {k: int(d[f"Rev_F{k}"].sum()) for k in (1, 2, 3)}
    ultima = max([k for k in cov if cov[k] > 0], default=0)
    completas = [k for k in cov if n and cov[k] >= 0.95 * n]
    actual = max(completas) if completas else ultima
    filas = {}
    for k in (1, 2, 3):
        r = cov[k]
        if n and r >= 0.95 * n:
            filas[k] = ("✅", "Revisada")
        elif r > 0:
            filas[k] = ("🔄", f"En revisión ({r} de {n} aulas)")
        elif k == ultima + 1:
            filas[k] = ("🔄", "En revisión (aún sin aulas revisadas)")
        else:
            filas[k] = ("⏳", "Pendiente")
    return cov, actual, filas


# ----------------------------------------------------------------------------
# Helpers de interfaz
# ----------------------------------------------------------------------------
def mostrar_df(df, **kw):
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kw)
    except Exception:  # versiones antiguas de Streamlit
        st.dataframe(df, use_container_width=True, hide_index=True, **kw)


def mostrar_fig(fig):
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), legend_title_text="")
    try:
        st.plotly_chart(fig, width="stretch")
    except Exception:
        st.plotly_chart(fig, use_container_width=True)


def tabla_aulas(d: pd.DataFrame, extra=()):
    cols = ["Profesor", "Curso", "Seccion", "Programa", "Cumplimiento", "Pendientes", *extra,
            "Criterios_No", "Criterios_Inc", "Observacion", "Ultima_Revision"]
    return d.sort_values(["Cumplimiento", "Profesor"])[cols].rename(columns=COL)


CFG_TABLA = {
    COL["Cumplimiento"]: st.column_config.ProgressColumn(COL["Cumplimiento"], min_value=0, max_value=100,
                                                         format="%.1f%%"),
    COL["Ultima_Revision"]: st.column_config.DateColumn(COL["Ultima_Revision"], format="DD/MM/YYYY"),
    COL["Acum_100"]: st.column_config.ProgressColumn(COL["Acum_100"], min_value=0, max_value=100, format="%.1f"),
}


def a_excel(hojas: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, d in hojas.items():
            d.to_excel(w, sheet_name=nombre[:31], index=False)
            ws = w.sheets[nombre[:31]]
            for i, c in enumerate(d.columns, 1):
                largo = max([len(str(c))] + [len(str(v)) for v in d[c].head(200)])
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(60, largo + 2)
    return buf.getvalue()


def resumen_profesor(d: pd.DataFrame, por="Profesor") -> pd.DataFrame:
    return d.groupby(por).agg(**{
        "Aulas revisadas": ("Curso", "size"),
        "Cumplimiento promedio (%)": ("Cumplimiento", "mean"),
        "Aulas completas (100%)": ("Nivel", lambda s: int((s == "completas").sum())),
        "Aulas críticas o vacías (menos de 50%)": ("Nivel", lambda s: int(s.isin(["criticas", "vacias"]).sum())),
        "Criterios NO cumplidos (total)": ("Criterios_No", "sum"),
        "Criterios incompletos (total)": ("Criterios_Inc", "sum"),
    }).reset_index().round({"Cumplimiento promedio (%)": 1}).sort_values("Cumplimiento promedio (%)")


# ----------------------------------------------------------------------------
# Bienvenida y carga del archivo
# ----------------------------------------------------------------------------
st.markdown('<div class="hero"><h1>🎓 Bienvenido al dashboard de seguimiento académico de la UCC</h1>'
            '<p>Siga los tres pasos para ver el informe: modalidad, ciclo y programa.</p></div>',
            unsafe_allow_html=True)

with st.sidebar:
    st.header("📂 Archivo")
    archivo = st.file_uploader("Sube el Excel de seguimiento", type=["xlsx", "xlsm"])

if archivo is not None:
    contenido = archivo.getvalue()
elif Path(ARCHIVO_POR_DEFECTO).exists():
    contenido = Path(ARCHIVO_POR_DEFECTO).read_bytes()
else:
    st.info("👈 Primero suba el archivo Excel de seguimiento desde la barra lateral.")
    st.stop()

try:
    base, meta = cargar_datos(contenido)
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()
for a in meta["avisos"]:
    st.warning(a)
if base.empty:
    st.error("No se encontraron aulas. Revise el panel de diagnóstico de la barra lateral.")
    st.stop()

with st.sidebar:
    with st.expander("⚙️ Ajustes avanzados"):
        fuente = st.radio("Fuente de los % por fase", ["Columnas del Excel", "Recalculados desde los criterios"],
                          help="El Excel no cuenta 'Sí' con tilde ni 'Intermedio' como cumplidos; al recalcular, "
                               "'Sí' = Si e 'Intermedio' = Incompleto (medio punto).")
    with st.expander("🧪 Diagnóstico de columnas"):
        st.json(meta["diag"])
usar_excel = fuente.startswith("Columnas")

# ----------------------------------------------------------------------------
# Pasos 1-3: modalidad -> ciclo -> programa
# ----------------------------------------------------------------------------
st.markdown("### 1️⃣ Seleccione la modalidad a revisar")
opciones_mod = [m for m in ("Extendida", "Virtual") if m in set(base["Modalidad"])]
mod = st.radio("Modalidad", opciones_mod, index=None, horizontal=True, label_visibility="collapsed",
               format_func=lambda m: NOMBRES_MOD.get(m, m))
if mod is None:
    st.info("👆 Elija una modalidad para continuar.")
    st.stop()

base_m = base[base["Modalidad"] == mod]
n_ciclo = base_m["Ciclo"].value_counts()
st.markdown("### 2️⃣ Seleccione el ciclo a revisar")
ciclo = st.selectbox("Ciclo", sorted(n_ciclo.index, reverse=True), index=None, key=f"ciclo_{mod}",
                     placeholder="Elija un ciclo", label_visibility="collapsed",
                     format_func=lambda c: f"{c}  ·  {n_ciclo[c]} aulas")
if ciclo is None:
    st.info("👆 Elija un ciclo para ver el informe.")
    st.stop()

base_c = base_m[base_m["Ciclo"] == ciclo]
st.markdown("### 3️⃣ Seleccione el programa (opcional)")
prog = st.selectbox("Programa", [TODOS] + sorted(map(str, base_c["Programa"].unique())),
                    key=f"prog_{mod}_{ciclo}", label_visibility="collapsed")
d_p = base_c if prog == TODOS else base_c[base_c["Programa"] == prog]
prog_txt = "todos los programas" if prog == TODOS else f"el programa {prog}"
st.divider()

# ----------------------------------------------------------------------------
# Estado de las fases
# ----------------------------------------------------------------------------
cov, actual, est_f = estado_fases(d_p)
if actual == 0:
    st.warning(f"Las {len(d_p)} aulas de {prog_txt} aún no tienen ninguna fase revisada en el ciclo {ciclo}.")
    st.stop()
pills = "".join(f'<span class="pill">{ico} <b>Fase {k}:</b> {txt}</span>' for k, (ico, txt) in est_f.items())
st.markdown(f"#### 📅 Fase actual revisada: Fase {actual}", help="Última fase con revisión en (casi) todas las aulas.")
st.markdown(pills, unsafe_allow_html=True)

revisadas = [k for k in (1, 2, 3) if cov[k] > 0]
if len(revisadas) > 1:
    opc = [f"Fase {k}" for k in revisadas] + ["Todas las fases revisadas (acumulado)"]
    eleg = st.radio("Fase que desea analizar", opc, horizontal=True, index=opc.index(f"Fase {actual}")
                    if actual in revisadas else 0, key=f"fase_{mod}_{ciclo}_{prog}")
    fase_vista = 0 if eleg.startswith("Todas") else int(eleg[-1])
else:
    fase_vista = revisadas[0]
fase_txt = "todas las fases revisadas" if fase_vista == 0 else f"Fase {fase_vista}"
fase_ref = fase_txt if fase_vista == 0 else f"la {fase_txt}"

df = calcular(d_p, usar_excel, fase_vista)
dr = df[df["Revisado"]].copy()
sin_rev = df[~df["Revisado"]]
if dr.empty:
    st.warning(f"No hay aulas revisadas en {fase_ref} para {prog_txt}.")
    st.stop()

# ----------------------------------------------------------------------------
# Mensaje informativo + tarjetas (KPI)
# ----------------------------------------------------------------------------
prom = dr["Cumplimiento"].mean()
cuenta = dr["Nivel"].value_counts()
st.info(f"📌 En **{prog_txt}** hay **{len(dr)} aula{'s' if len(dr) != 1 else ''}** revisada{'s' if len(dr) != 1 else ''} en **{fase_ref}** (ciclo {ciclo}) con un "
        f"cumplimiento promedio de **{prom:.1f}%**. "
        + (f"Otras {len(sin_rev)} aulas todavía no se han revisado en esta fase." if len(sin_rev) else
           "Todas las aulas del ciclo ya fueron revisadas en esta fase."))

sig = (mod, ciclo, prog, fase_vista)
if st.session_state.get("sig") != sig:           # al cambiar la selección se limpia el detalle
    st.session_state["sig"], st.session_state["nivel"] = sig, None

DESC = {"completas": f"Cumplen el 100% de los criterios de {fase_ref}",
        "avance": f"Cumplen entre 50% y 99% de los criterios de {fase_ref}",
        "criticas": "Cumplen menos del 50% (sin contar las aulas vacías)",
        "vacias": "0% de cumplimiento: ningún criterio cumplido"}
cols_k = st.columns(4)
for col, (clave, (ico, titulo, color, fondo)) in zip(cols_k, NIVELES.items()):
    n = int(cuenta.get(clave, 0))
    act = " act" if st.session_state.get("nivel") == clave else ""
    with col:
        st.markdown(f'<div class="kpi{act}" style="--c:{color}; --bg:{fondo}">'
                    f'<div class="kpi-t">{ico} {titulo}</div><div class="kpi-d">{DESC[clave]}</div>'
                    f'<div class="kpi-v" style="color:{color}">{n}</div>'
                    f'<div class="kpi-s">{n / len(dr):.0%} de las aulas revisadas</div></div>',
                    unsafe_allow_html=True)
        st.button("Ver estas aulas ▸", key=f"btn_{clave}", on_click=lambda c=clave: st.session_state.update(nivel=c))

nivel = st.session_state.get("nivel")
if nivel is None:
    st.caption("👆 Haga clic en «Ver estas aulas» de una tarjeta para ver el detalle de esos cursos.")
else:
    ico, titulo = NIVELES[nivel][:2]
    sel = dr[dr["Nivel"] == nivel]
    st.subheader(f"{ico} {titulo}: {len(sel)} aulas · {prog_txt} · ciclo {ciclo} · {fase_txt}")
    if sel.empty:
        st.success("No hay aulas en este nivel.")
    else:
        mostrar_df(tabla_aulas(sel, extra=(["Hasta fase"] if fase_vista == 0 else [])), column_config=CFG_TABLA)
if len(sin_rev):
    with st.expander(f"🕓 {len(sin_rev)} aula(s) sin revisión en {fase_ref}"):
        mostrar_df(sin_rev[["Profesor", "Curso", "Seccion", "Programa"]].rename(columns=COL))
st.divider()

tab1, tab2, tab3 = st.tabs(["📊 Análisis por criterio", "👩‍🏫 Cumplimiento por profesor", "📥 Exportar"])

# ----------------------------------------------------------------------------
# Análisis por criterio
# ----------------------------------------------------------------------------
with tab1:
    fc = meta["fase_crit"]
    crits = [c for m in dr["Modalidad"].unique() for c in meta["criterios"].get(m, [])]
    crits = sorted(dict.fromkeys(crits), key=lambda c: fc[c])
    if fase_vista:
        crits = [c for c in crits if fc[c] == fase_vista]
    crits = [c for c in crits if c in dr.columns and dr[c].notna().any()]
    if not crits:
        st.info("No hay criterios evaluados con la selección actual.")
    else:
        ids = ["Modalidad", "Ciclo", "Etiqueta", "Curso", "Profesor", "Programa"]
        largo = dr.melt(id_vars=ids, value_vars=crits, var_name="Criterio", value_name="Estado").dropna(subset=["Estado"])
        cnt = largo.groupby(["Criterio", "Estado"]).size().reset_index(name="Cursos")
        cnt["% de cursos"] = cnt["Cursos"] / cnt.groupby("Criterio")["Cursos"].transform("sum") * 100
        fig = px.bar(cnt, x="% de cursos", y="Criterio", color="Estado", orientation="h", hover_data=["Cursos"],
                     color_discrete_map=COLORES_ESTADO,
                     category_orders={"Estado": ["Si", "Incompleto", "No"], "Criterio": crits[::-1]},
                     title=f"% de cursos por estado en cada criterio · {fase_txt} · {prog_txt}")
        fig.update_layout(height=max(380, 30 * len(crits) + 120), yaxis_title=None, xaxis_range=[0, 100])
        mostrar_fig(fig)

        st.subheader("¿Qué cursos carecen de un criterio específico?")
        c1, c2, c3 = st.columns([2, 2, 2])
        criterio = c1.selectbox("Criterio", crits)
        estados = c2.multiselect("Estados a revisar", ["No", "Incompleto", "Si"], default=["No", "Incompleto"])
        agrupar = c3.radio("Agrupar gráfico por", ["Programa", "Profesor"], horizontal=True)
        sub = largo[(largo["Criterio"] == criterio) & largo["Estado"].isin(estados)]
        if sub.empty:
            st.success(f"Todos los cursos revisados cumplen con «{criterio}».")
        else:
            g = sub.groupby([agrupar, "Estado"]).size().reset_index(name="Cursos")
            orden = g.groupby(agrupar)["Cursos"].sum().nlargest(20).index
            fig = px.bar(g[g[agrupar].isin(orden)], x="Cursos", y=agrupar, color="Estado", orientation="h",
                         color_discrete_map=COLORES_ESTADO, title=f"«{criterio}» por {agrupar.lower()} (top 20)")
            fig.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title=None)
            mostrar_fig(fig)
            mostrar_df(sub[["Profesor", "Curso", "Programa", "Estado"]].rename(columns={"Estado": "Estado del criterio"}))

# ----------------------------------------------------------------------------
# Cumplimiento por profesor (depende del programa elegido arriba)
# ----------------------------------------------------------------------------
with tab2:
    vista = "Profesor"
    if prog == TODOS:
        vista = st.radio("Agrupar por", ["Profesor", "Programa"], horizontal=True)
    st.caption(f"Cumplimiento en {fase_ref} · {prog_txt} · ciclo {ciclo}")
    res = resumen_profesor(dr, vista)
    peor = res.head(20).sort_values("Cumplimiento promedio (%)", ascending=False)
    fig = px.bar(peor, x="Cumplimiento promedio (%)", y=vista, orientation="h", color="Cumplimiento promedio (%)",
                 color_continuous_scale=["#dc2626", "#f59e0b", "#16a34a"], range_color=[0, 100],
                 title=f"Menor cumplimiento promedio por {vista.lower()} (máx. 20)")
    fig.update_layout(height=max(350, 28 * len(peor) + 120), yaxis_title=None, xaxis_range=[0, 100],
                      coloraxis_showscale=False)
    mostrar_fig(fig)
    mostrar_df(res, column_config={"Cumplimiento promedio (%)": st.column_config.ProgressColumn(
        "Cumplimiento promedio (%)", min_value=0, max_value=100, format="%.1f%%")})

    st.subheader(f"Detalle de aulas por {vista.lower()}")
    elegido = st.selectbox(f"Seleccione un {vista.lower()}", res[vista].tolist())
    mostrar_df(tabla_aulas(dr[dr[vista] == elegido]), column_config=CFG_TABLA)

# ----------------------------------------------------------------------------
# Exportar
# ----------------------------------------------------------------------------
with tab3:
    st.subheader("Exportar reporte de aulas")
    st.caption("Elija qué incluir. Si deja un filtro vacío, se incluyen todas las opciones.")
    sg = f"{mod}_{ciclo}_{prog}"
    e1, e2, e3 = st.columns(3)
    m_sel = e1.multiselect("Modalidad", opciones_mod, default=[mod], key=f"em_{sg}",
                           format_func=lambda m: NOMBRES_MOD.get(m, m)) or opciones_mod
    d_e = base[base["Modalidad"].isin(m_sel)]
    ciclos_e = sorted(d_e["Ciclo"].unique(), reverse=True)
    c_sel = e2.multiselect("Ciclo", ciclos_e, default=[ciclo] if ciclo in ciclos_e else [],
                           key=f"ec_{sg}_{'-'.join(m_sel)}", placeholder="Todos") or ciclos_e
    d_e = d_e[d_e["Ciclo"].isin(c_sel)]
    progs_e = sorted(map(str, d_e["Programa"].unique()))
    p_sel = e3.multiselect("Programa", progs_e, default=[prog] if prog in progs_e else [],
                           key=f"ep_{sg}_{'-'.join(m_sel)}_{'-'.join(c_sel)}", placeholder="Todos") or progs_e
    d_e = d_e[d_e["Programa"].isin(p_sel)]

    e4, e5 = st.columns(2)
    op_f = ["Todas las fases revisadas (acumulado)", "Fase 1", "Fase 2", "Fase 3"]
    f_sel = e4.selectbox("Fase a evaluar", op_f, index=fase_vista, key=f"ef_{sg}_{fase_vista}")
    fase_e = op_f.index(f_sel)
    n_sel = e5.multiselect("Niveles de cumplimiento a incluir", list(NIVELES), key="en",
                           default=["avance", "criticas", "vacias"], format_func=lambda k: TXT_NIVEL[k])

    ex_all = calcular(d_e, usar_excel, fase_e)
    ex = ex_all[ex_all["Revisado"] & ex_all["Nivel"].isin(n_sel)].sort_values(["Cumplimiento", "Profesor"])
    st.caption(f"{len(ex)} aula(s) cumplen los filtros elegidos.")
    if ex.empty:
        st.info("No hay aulas para exportar con esos filtros.")
    else:
        reporte = ex[["Modalidad", "Ciclo", "Programa", "Profesor", "Curso", "ID_Curso", "Seccion", "Nivel_txt",
                      "Cumplimiento", "Pendientes", "Criterios_No", "Criterios_Inc", "Hasta fase", "Acum_100",
                      "Observacion", "Ultima_Revision"]].rename(columns=COL)
        reporte["Modalidad"] = reporte["Modalidad"].map(NOMBRES_MOD)
        resumen = resumen_profesor(ex)
        sin = ex_all[~ex_all["Revisado"]][["Modalidad", "Ciclo", "Programa", "Profesor", "Curso", "Seccion"]].rename(columns=COL)
        mostrar_df(reporte, column_config=CFG_TABLA)
        sello = datetime.now().strftime("%Y%m%d_%H%M")
        b1, b2, _ = st.columns([1, 1, 3])
        b1.download_button("⬇️ Descargar Excel",
                           a_excel({"Aulas": reporte, "Resumen por profesor": resumen, "Sin revisar": sin}),
                           f"reporte_aulas_{sello}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        b2.download_button("⬇️ Descargar CSV", reporte.to_csv(index=False).encode("utf-8-sig"),
                           f"reporte_aulas_{sello}.csv", "text/csv")
