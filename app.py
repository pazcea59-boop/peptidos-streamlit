"""
Base de datos de péptidos (Streamlit)
================================================================
Aplicación para registrar, editar, guardar y exportar (Excel / PDF / CSV) una
base de datos de péptidos, con:
  - acceso mediante RUT chileno (validación módulo 11),
  - persistencia por usuario: los registros se guardan en disco y se recuperan
    al volver a ingresar con el mismo RUT,
  - visualización 3D interactiva de estructuras desde RCSB PDB,
  - verificación científica automática contra fuentes públicas (RCSB PDB,
    UniProt y PubMed) y comprobación de coherencia interna de cada registro.

Ejecución local:
    streamlit run app.py

Datos persistentes (junto a app.py, sin importar desde dónde se ejecute):
    data/usuarios/<huella_del_RUT>/peptidos_guardados.csv
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import time
import urllib.error
import urllib.request
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

TITULO_APP = "Base de datos de péptidos"

st.set_page_config(page_title=TITULO_APP, layout="wide")

# ---------------------------------------------------------------------------
# 1. CONSTANTES Y DATOS DE REFERENCIA
# ---------------------------------------------------------------------------
# La carpeta de datos se ancla a la ubicación de app.py (no al directorio de
# trabajo), así los registros se encuentran siempre en el mismo lugar.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data" / "usuarios"

COL_ID = "Identificador"
COL_OWNER = "RUT propietario"
COL_PDB = "Estructura 3D (PDB / enlace)"
COL_VERIF_ESTADO = "Estado de verificación"
COL_VERIF_NOTAS = "Notas de verificación"
COL_VERIF_FECHA = "Fecha de verificación"
COLS_VERIF = [COL_VERIF_ESTADO, COL_VERIF_NOTAS, COL_VERIF_FECHA]
COLUMNS = [
    COL_ID,
    "Secuencia",
    "Longitud",
    "Masa molecular (Da)",
    "Polaridad",
    "% Apolares",
    "% Polares sin carga",
    "% Carga positiva",
    "% Carga negativa",
    "Composición de aminoácidos",
    "Hidrofobicidad (GRAVY)",
    "Carga neta aprox. (pH 7)",
    COL_PDB,
    "Notas estructurales",
    "Propiedades complementarias",
    "Notas",
    COL_VERIF_ESTADO,
    COL_VERIF_NOTAS,
    COL_VERIF_FECHA,
    "Fecha de registro",
    COL_OWNER,
]
COLS_NUMERICAS = ["Longitud", "Masa molecular (Da)", "% Apolares", "% Polares sin carga",
                  "% Carga positiva", "% Carga negativa", "Hidrofobicidad (GRAVY)",
                  "Carga neta aprox. (pH 7)"]
COLS_TEXTO = [c for c in COLUMNS if c not in COLS_NUMERICAS]
# Columnas que el usuario puede elegir incluir en un reporte (el RUT va en el encabezado)
COLUMNS_EXPORTABLES = [c for c in COLUMNS if c not in (COL_ID, COL_OWNER)]
# El RUT propietario se gestiona internamente: se oculta en las tablas
CONFIG_OCULTA = {COL_OWNER: None}

# Escala de hidropatía de Kyte-Doolittle (1982)
KYTE_DOOLITTLE = {
    "A": 1.8, "R": -4.5, "N": -3.5, "D": -3.5, "C": 2.5,
    "Q": -3.5, "E": -3.5, "G": -0.4, "H": -3.2, "I": 4.5,
    "L": 3.8, "K": -3.9, "M": 1.9, "F": 2.8, "P": -1.6,
    "S": -0.8, "T": -0.7, "W": -0.9, "Y": -1.3, "V": 4.2,
}

# Masas promedio de residuos (Da); la masa del péptido suma además una molécula de agua
MASA_RESIDUO = {
    "A": 71.0788, "R": 156.1875, "N": 114.1038, "D": 115.0886, "C": 103.1388,
    "E": 129.1155, "Q": 128.1307, "G": 57.0519, "H": 137.1411, "I": 113.1594,
    "L": 113.1594, "K": 128.1741, "M": 131.1926, "F": 147.1766, "P": 97.1167,
    "S": 87.0782, "T": 101.1051, "W": 186.2132, "Y": 163.1760, "V": 99.1326,
}
MASA_AGUA = 18.01524

TRES_A_UNA = {
    "ala": "A", "arg": "R", "asn": "N", "asp": "D", "cys": "C",
    "gln": "Q", "glu": "E", "gly": "G", "his": "H", "ile": "I",
    "leu": "L", "lys": "K", "met": "M", "phe": "F", "pro": "P",
    "ser": "S", "thr": "T", "trp": "W", "tyr": "Y", "val": "V",
}

# Clasificación simplificada de residuos (heurística, ver README)
APOLARES = set("GAVLIMFWP")
POLARES_SIN_CARGA = set("STCYNQ")
POSITIVOS = set("KRH")
NEGATIVOS = set("DE")

FORMATO_1 = "Una letra (ej. ACDK)"
FORMATO_3 = "Tres letras (ej. Ala-Cys-Asp)"
AUTO = "Calcular automáticamente"
MANUAL = "Ingresar manualmente"
POLARIDADES = ["Polar", "Apolar (hidrofóbico)", "Anfipático / mixto",
               "Cargado positivo", "Cargado negativo"]

OPCIONES_MENU = [
    "Crear nuevo registro de péptido",
    "Gestionar mis registros (editar / guardar / exportar)",
    "Visualizar estructura 3D (RCSB PDB)",
    "Importar registros desde CSV / Excel",
]

TITULO_REPORTE_DEFECTO = TITULO_APP

# Valores por defecto de los widgets del formulario (también sirven para limpiarlo)
DEFAULTS_FORM = {
    "f_id": "", "f_formato": FORMATO_1, "f_seq": "",
    "f_pol_modo": AUTO, "f_pol_manual": POLARIDADES[0],
    "f_comp_modo": AUTO, "f_comp_manual": "",
    "f_hid_modo": AUTO, "f_hid_manual": 0.0,
    "f_pdb": "", "f_ver3d": False, "f_estruc": "", "f_props": "", "f_notas": "",
    "f_verif_auto": True,
}


# ---------------------------------------------------------------------------
# 2. VALIDACIÓN DE RUT (MÓDULO 11)
# ---------------------------------------------------------------------------
def calcular_dv(cuerpo: str) -> str:
    """Dígito verificador de un RUT por módulo 11.

    Se recorren los dígitos del cuerpo de derecha a izquierda multiplicándolos
    por la serie 2, 3, 4, 5, 6, 7, 2, 3... La suma se divide por 11 y el
    dígito es 11 - resto (11 -> "0", 10 -> "K").
    """
    suma, factor = 0, 2
    for digito in reversed(cuerpo):
        suma += int(digito) * factor
        factor = factor + 1 if factor < 7 else 2
    resultado = 11 - (suma % 11)
    if resultado == 11:
        return "0"
    if resultado == 10:
        return "K"
    return str(resultado)


def validar_rut(raw: str) -> tuple[str | None, str | None]:
    """Valida un RUT chileno. Devuelve (rut_normalizado, mensaje_de_error).

    El RUT normalizado tiene el formato «12345678-5» (sin puntos).
    """
    limpio = re.sub(r"[.\-\s]", "", raw or "").upper()
    if not limpio:
        return None, "Ingresa tu RUT."
    if not re.fullmatch(r"\d{7,8}[\dK]", limpio):
        return None, "Formato de RUT no válido. Ejemplo: 12.345.678-5"
    cuerpo, dv = limpio[:-1], limpio[-1]
    if calcular_dv(cuerpo) != dv:
        return None, "El dígito verificador no corresponde. Revisa el RUT ingresado."
    return f"{cuerpo}-{dv}", None


def formatear_rut(rut: str) -> str:
    """12345678-5 -> 12.345.678-5"""
    cuerpo, dv = rut.split("-")
    return f"{int(cuerpo):,}".replace(",", ".") + f"-{dv}"


# ---------------------------------------------------------------------------
# 3. LÓGICA DE PROPIEDADES FISICOQUÍMICAS
# ---------------------------------------------------------------------------
def parse_sequence(raw: str, formato: str) -> tuple[str | None, str | None]:
    """Normaliza una secuencia a código de una letra.

    Devuelve (secuencia, mensaje_de_error). Si hay error, secuencia es None.
    """
    limpio = re.sub(r"[\s\-_,.;]", "", raw or "")
    if not limpio:
        return None, "La secuencia está vacía."

    if formato == FORMATO_3:
        limpio = limpio.lower()
        if len(limpio) % 3 != 0:
            return None, "En código de tres letras, la longitud debe ser múltiplo de 3 (ej. AlaGly)."
        codigos = [limpio[i:i + 3] for i in range(0, len(limpio), 3)]
        invalidos = sorted({c for c in codigos if c not in TRES_A_UNA})
        if invalidos:
            return None, f"Códigos de tres letras no reconocidos: {', '.join(invalidos)}"
        return "".join(TRES_A_UNA[c] for c in codigos), None

    limpio = limpio.upper()
    invalidos = sorted(set(limpio) - set(KYTE_DOOLITTLE))
    if invalidos:
        return None, (f"Caracteres no válidos: {', '.join(invalidos)}. "
                      "Solo se admiten los 20 aminoácidos estándar.")
    return limpio, None


def compute_properties(seq: str) -> dict:
    """Calcula propiedades a partir de la secuencia (código de una letra)."""
    n = len(seq)
    cnt = Counter(seq)

    def pct(grupo: set) -> float:
        return round(100 * sum(cnt[a] for a in grupo) / n, 1)

    p_apolar = pct(APOLARES)
    if p_apolar >= 60:
        polaridad = "Apolar (hidrofóbico)"
    elif p_apolar <= 40:
        polaridad = "Polar"
    else:
        polaridad = "Anfipático / mixto"

    return {
        "longitud": n,
        "masa": round(sum(MASA_RESIDUO[a] * c for a, c in cnt.items()) + MASA_AGUA, 2),
        "gravy": round(sum(KYTE_DOOLITTLE[a] * c for a, c in cnt.items()) / n, 3),
        "p_apolar": p_apolar,
        "p_polar": pct(POLARES_SIN_CARGA),
        "p_pos": pct(POSITIVOS),
        "p_neg": pct(NEGATIVOS),
        "carga": cnt["K"] + cnt["R"] - cnt["D"] - cnt["E"],  # aproximación; ignora His y extremos
        "polaridad": polaridad,
        "composicion": "; ".join(f"{a}: {cnt[a]} ({100 * cnt[a] / n:.1f}%)" for a in sorted(cnt)),
    }


# ---------------------------------------------------------------------------
# 4. SESIÓN Y PERSISTENCIA POR USUARIO
# ---------------------------------------------------------------------------
def rut_actual() -> str:
    return st.session_state["rut"]


def limpiar(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza una tabla: columnas, tipos, filas sin ID y vínculo con el usuario."""
    df = df.copy()
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[COLUMNS]
    for c in COLS_TEXTO:  # evita columnas de texto inferidas como numéricas (NaN)
        df[c] = df[c].fillna("").astype(str)
    for c in COLS_NUMERICAS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df[COL_ID].str.strip() != ""].reset_index(drop=True)
    df[COL_OWNER] = rut_actual()  # todo registro queda asociado al RUT de la sesión
    return df


def ruta_csv_usuario() -> Path:
    """Ruta del CSV del usuario. La carpeta usa una huella (hash) del RUT, no el RUT."""
    huella = hashlib.sha256(rut_actual().encode("utf-8")).hexdigest()[:24]
    return DATA_DIR / huella / "peptidos_guardados.csv"


def cargar_guardados() -> pd.DataFrame:
    """Lee del disco los registros del usuario actual (y solo los suyos)."""
    ruta = ruta_csv_usuario()
    if not ruta.exists():
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str)
    except (OSError, pd.errors.ParserError, pd.errors.EmptyDataError):
        return pd.DataFrame(columns=COLUMNS)
    if COL_OWNER in df.columns:  # doble protección: solo filas del propietario
        df = df[df[COL_OWNER].astype(str) == rut_actual()]
    return limpiar(df)


def guardar_csv(df: pd.DataFrame) -> None:
    """Escritura atómica: se escribe un temporal y luego se reemplaza el archivo."""
    ruta = ruta_csv_usuario()
    ruta.parent.mkdir(parents=True, exist_ok=True)
    temporal = ruta.with_suffix(".tmp")
    df.to_csv(temporal, index=False, encoding="utf-8-sig")
    os.replace(temporal, ruta)


def persistir() -> bool:
    """Guarda en disco todos los registros de la sesión. Devuelve True si tuvo éxito."""
    try:
        guardar_csv(limpiar(st.session_state.registros))
        return True
    except OSError:
        return False


def init_state() -> None:
    # Al iniciar sesión se recuperan los registros guardados del usuario
    if "registros" not in st.session_state:
        st.session_state.registros = cargar_guardados()
    st.session_state.setdefault("editor_version", 0)
    for k, v in DEFAULTS_FORM.items():
        st.session_state.setdefault(k, v)


def cerrar_sesion() -> None:
    """Callback: borra el estado de la sesión. Los datos guardados en disco se conservan."""
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def flash(tipo: str, mensaje: str) -> None:
    """Guarda un mensaje para mostrarlo tras el siguiente rerun."""
    st.session_state["_flash"] = (tipo, mensaje)


def show_flash() -> None:
    if "_flash" in st.session_state:
        tipo, mensaje = st.session_state.pop("_flash")
        getattr(st, tipo)(mensaje)


def concat_seguro(frames: list[pd.DataFrame]) -> pd.DataFrame:
    frames = [f for f in frames if f is not None and not f.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLUMNS)


# ---------------------------------------------------------------------------
# 5. ESTRUCTURA 3D DESDE RCSB PDB
# ---------------------------------------------------------------------------
_PDB_ID = re.compile(r"[0-9][A-Za-z0-9]{3}")
_EXTENSIONES = re.compile(r"\.(pdb|cif|ent|pdb1|cif\.gz|pdb\.gz|pdb1\.gz)$", re.IGNORECASE)
MAX_BYTES_ESTRUCTURA = 25 * 1024 * 1024

REPRESENTACIONES = {
    "Cinta (cartoon)": "cartoon",
    "Bastones": "stick",
    "Líneas": "line",
    "Esferas": "sphere",
}
COLORES_3D = {
    "Espectro (extremo N a C)": {"color": "spectrum"},
    "Por cadena": {"colorscheme": "chain"},
    "Estructura secundaria": {"colorscheme": "ssJmol"},
    "Uniforme": {"color": "#1F4E78"},
}


def extraer_pdb_id(raw) -> str | None:
    """Obtiene un código PDB de 4 caracteres desde un código o un enlace de RCSB.

    Acepta, por ejemplo: «1CRN», «https://www.rcsb.org/structure/1CRN»,
    «https://www.rcsb.org/3d-view/1CRN», «https://files.rcsb.org/download/1CRN.pdb».
    """
    texto = (raw or "").strip() if isinstance(raw, str) else ""
    if not texto:
        return None
    if _PDB_ID.fullmatch(texto):
        return texto.upper()

    if not re.match(r"^https?://", texto, re.IGNORECASE):
        if "rcsb.org" not in texto.lower():
            return None
        texto = "https://" + texto
    url = urlparse(texto)
    host = (url.hostname or "").lower()
    if host != "rcsb.org" and not host.endswith(".rcsb.org"):
        return None

    consulta = parse_qs(url.query)
    for clave in ("structureId", "pdbId", "id"):
        for valor in consulta.get(clave, []):
            if _PDB_ID.fullmatch(valor):
                return valor.upper()
    for segmento in reversed([s for s in url.path.split("/") if s]):
        candidato = _EXTENSIONES.sub("", segmento)
        if _PDB_ID.fullmatch(candidato):
            return candidato.upper()
    return None


def _descargar(url: str, max_bytes: int = MAX_BYTES_ESTRUCTURA) -> bytes | None:
    """Descarga una URL. Devuelve None si no existe (404); lanza RuntimeError en otros fallos."""
    peticion = urllib.request.Request(url, headers={"User-Agent": "peptidos-streamlit/1.0"})
    try:
        with urllib.request.urlopen(peticion, timeout=20) as respuesta:
            datos = respuesta.read(max_bytes + 1)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        raise RuntimeError(f"RCSB respondió con error HTTP {exc.code}.") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise RuntimeError("No se pudo conectar con RCSB PDB. Revisa la conexión a internet.") from exc
    if len(datos) > max_bytes:
        raise RuntimeError("El archivo de la estructura es demasiado grande para visualizarlo.")
    return datos


# Los errores de red lanzan excepciones, que cache_data no almacena; así un fallo
# transitorio no queda en caché.
@st.cache_data(ttl=86400, show_spinner=False)
def descargar_estructura(pdb_id: str) -> tuple[str, str] | None:
    """Devuelve (contenido, formato) con formato 'pdb' o 'cif'; None si el código no existe."""
    for extension, formato in (("pdb", "pdb"), ("cif", "cif")):
        datos = _descargar(f"https://files.rcsb.org/download/{pdb_id}.{extension}")
        if datos is not None:
            return datos.decode("utf-8", errors="ignore"), formato
    return None


@st.cache_data(ttl=86400, show_spinner=False)
def info_rcsb(pdb_id: str) -> dict:
    """Metadatos básicos de la entrada (título, método, resolución)."""
    datos = _descargar(f"https://data.rcsb.org/rest/v1/core/entry/{pdb_id}", max_bytes=2_000_000)
    if datos is None:
        return {}
    j = json.loads(datos.decode("utf-8"))
    resolucion = (j.get("rcsb_entry_info") or {}).get("resolution_combined") or []
    exptl = j.get("exptl") or [{}]
    return {
        "titulo": (j.get("struct") or {}).get("title", ""),
        "metodo": exptl[0].get("method", ""),
        "resolucion": resolucion[0] if resolucion else None,
    }


_PLANTILLA_VISOR = """<div id="visor" style="width:100%;height:__ALTO__px;position:relative;
border:1px solid #c9ced4;border-radius:4px;"></div>
<script src="https://3dmol.org/build/3Dmol-min.js"></script>
<script>
(function () {
  var el = document.getElementById("visor");
  if (typeof $3Dmol === "undefined") {
    el.innerHTML = "<p style='font-family:sans-serif;padding:12px'>No se pudo cargar el visor " +
                   "3D (requiere acceso a 3dmol.org desde el navegador).</p>";
    return;
  }
  var datos = __DATOS__;
  var viewer = $3Dmol.createViewer(el, { backgroundColor: "white" });
  viewer.addModel(datos, "__FORMATO__");
  viewer.setStyle({}, __ESTILO__);
  viewer.setStyle({ resn: "HOH" }, {});
  viewer.zoomTo();
  viewer.render();
  if (__ROTAR__) { viewer.spin(true); }
})();
</script>"""


def html_visor(contenido: str, formato: str, representacion: str, color: dict,
               rotar: bool, alto: int = 520) -> str:
    """Genera el HTML autocontenido del visor (3Dmol.js, el motor de py3Dmol)."""
    datos = json.dumps(contenido).replace("</", "<\\/")  # evita cerrar el <script> por error
    estilo = json.dumps({representacion: color})
    return (_PLANTILLA_VISOR
            .replace("__ALTO__", str(alto))
            .replace("__DATOS__", datos)
            .replace("__FORMATO__", "pdb" if formato == "pdb" else "mmcif")
            .replace("__ESTILO__", estilo)
            .replace("__ROTAR__", "true" if rotar else "false"))


def mostrar_visor_3d(pdb_id: str, clave: str) -> None:
    """Descarga la estructura desde RCSB y la muestra de forma interactiva."""
    try:
        with st.spinner(f"Descargando la estructura {pdb_id} desde RCSB PDB..."):
            resultado = descargar_estructura(pdb_id)
    except RuntimeError as exc:
        st.error(str(exc))
        return
    if resultado is None:
        st.error(f"El código «{pdb_id}» no existe en RCSB PDB.")
        return
    contenido, formato = resultado

    c1, c2, c3 = st.columns(3)
    rep = c1.selectbox("Representación", list(REPRESENTACIONES), key=f"{clave}_rep")
    col = c2.selectbox("Coloreado", list(COLORES_3D), key=f"{clave}_col")
    rotar = c3.checkbox("Rotación automática", key=f"{clave}_rot")

    components.html(
        html_visor(contenido, formato, REPRESENTACIONES[rep], COLORES_3D[col], rotar),
        height=540,
    )

    try:
        info = info_rcsb(pdb_id)
    except (RuntimeError, ValueError):
        info = {}
    partes = [f"Código: {pdb_id}"]
    if info.get("titulo"):
        partes.append(info["titulo"])
    if info.get("metodo"):
        partes.append(f"Método: {info['metodo']}")
    if info.get("resolucion"):
        partes.append(f"Resolución: {info['resolucion']} Å")
    st.caption(" | ".join(partes))
    st.markdown(f"[Abrir la entrada en RCSB PDB](https://www.rcsb.org/structure/{pdb_id})")
    st.caption("Arrastra para rotar, rueda para acercar y Ctrl + arrastrar para desplazar.")


# ---------------------------------------------------------------------------
# 6. VERIFICACIÓN CIENTÍFICA (RCSB PDB, UNIPROT, PUBMED)
# ---------------------------------------------------------------------------
# Niveles de resultado por fuente
NIVEL_OK = "Coincide"
NIVEL_PARCIAL = "Coincidencia parcial"
NIVEL_DISC = "Discrepancia"
NIVEL_NADA = "Sin coincidencias"
NIVEL_ERR = "No disponible"
NIVEL_OMIT = "No aplicable"

# Estados consolidados del registro
ESTADO_VERIFICADO = "Verificado"
ESTADO_AVISOS = "Verificado con avisos"
ESTADO_PARCIAL = "Verificación parcial"
ESTADO_DISC = "Discrepancia detectada"
ESTADO_SIN = "Sin coincidencias externas"
ESTADO_NC = "No concluyente (sin conexión)"
ESTADO_PEND = "Pendiente de verificación"

# Tipo de mensaje de Streamlit según el estado
ESTADO_TIPO = {
    ESTADO_VERIFICADO: "success",
    ESTADO_AVISOS: "warning",
    ESTADO_PARCIAL: "warning",
    ESTADO_DISC: "error",
    ESTADO_SIN: "info",
    ESTADO_NC: "info",
    ESTADO_PEND: "info",
}

F_INTERNA = "Coherencia interna"
F_RCSB = "RCSB PDB (entrada)"
F_RCSB_BUSQ = "RCSB PDB (búsqueda por secuencia)"
F_UNIPROT = "UniProt"
F_PUBMED = "PubMed"
FUENTES_SECUENCIA = (F_RCSB, F_RCSB_BUSQ, F_UNIPROT)

URL_RCSB_GRAPHQL = "https://data.rcsb.org/graphql"
URL_RCSB_BUSQUEDA = "https://search.rcsb.org/rcsbsearch/v2/query"
URL_UNIPROT = "https://rest.uniprot.org/uniprotkb/search"
URL_EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"

AVISO_VERIFICACION = ("La verificación automática respalda la información con fuentes públicas, "
                      "pero no sustituye la revisión por un especialista.")

# Caché en memoria de las consultas externas (solo se guardan respuestas sin errores)
_CACHE_VERIF: dict[tuple, tuple[float, list]] = {}
_TTL_VERIF = 3600


def _res(fuente: str, nivel: str, detalle: str, **extra) -> dict:
    return {"fuente": fuente, "nivel": nivel, "detalle": detalle, **extra}


def _http_json(url: str, payload: dict | None = None, timeout: int = 12):
    """GET (o POST si hay payload) que devuelve JSON. None si no hay contenido o 404.

    Lanza RuntimeError ante fallos de red, errores HTTP o respuestas no válidas.
    """
    cabeceras = {"User-Agent": "peptidos-streamlit/1.0", "Accept": "application/json"}
    cuerpo = None
    if payload is not None:
        cuerpo = json.dumps(payload).encode("utf-8")
        cabeceras["Content-Type"] = "application/json"
    peticion = urllib.request.Request(url, data=cuerpo, headers=cabeceras)
    try:
        with urllib.request.urlopen(peticion, timeout=timeout) as respuesta:
            datos = respuesta.read(5_000_001)
    except urllib.error.HTTPError as exc:
        if exc.code in (204, 404):
            return None
        raise RuntimeError(f"El servicio respondió con error HTTP {exc.code}.") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise RuntimeError("No se pudo conectar con el servicio externo.") from exc
    if not datos:
        return None
    if len(datos) > 5_000_000:
        raise RuntimeError("La respuesta del servicio es demasiado grande.")
    try:
        return json.loads(datos.decode("utf-8"))
    except ValueError as exc:
        raise RuntimeError("Respuesta no válida del servicio.") from exc


def _a_float(valor) -> float | None:
    try:
        x = float(valor)
    except (TypeError, ValueError):
        return None
    return None if x != x else x


def nombre_informativo(nombre: str) -> str | None:
    """Devuelve el nombre si sirve para buscar en literatura; None si es un código genérico."""
    n = re.sub(r'["\[\]()]', " ", nombre or "")
    n = re.sub(r"\s+", " ", n).strip()
    if len(n) < 3 or not re.search(r"[A-Za-z]{3,}", n):
        return None
    if re.fullmatch(r"(?i)(pep|peptido|péptido|seq|sec|muestra|sample|test|prueba)[\s_\-]*\d*[a-z]?", n):
        return None
    return n


# --- Comparación de secuencias ---------------------------------------------
_RANGO = {"idéntica": 4, "contenida": 3, "fragmento": 2, "similar": 1}


def comparar_secuencias(reg: str, ref: str) -> tuple[str, float]:
    """Compara la secuencia registrada con una de referencia.

    Devuelve (tipo, valor): idéntica, contenida (el péptido forma parte de la
    referencia), fragmento (la referencia es parte del péptido; valor = fracción
    cubierta) o similar (valor = fracción aproximada de residuos coincidentes).
    """
    reg = reg.upper()
    ref = re.sub(r"[^A-Za-z]", "", ref or "").upper()
    if not ref:
        return "similar", 0.0
    if reg == ref:
        return "idéntica", 1.0
    if reg in ref:
        # Una secuencia muy corta está contenida en casi cualquier proteína: no es evidencia
        return ("contenida", 1.0) if len(reg) >= 6 else ("similar", 1.0)
    if ref in reg and len(ref) >= 6:
        return "fragmento", len(ref) / len(reg)
    if len(reg) * len(ref) > 4_000_000:
        return "similar", 0.0
    coincidencias = SequenceMatcher(None, reg, ref, autojunk=False).get_matching_blocks()
    return "similar", sum(b.size for b in coincidencias) / len(reg)


def _texto_comparacion(tipo: str, valor: float, largo_ref: int) -> str:
    if tipo == "idéntica":
        return "la secuencia registrada es idéntica a la de la referencia"
    if tipo == "contenida":
        return f"la secuencia registrada está contenida en la referencia (longitud {largo_ref})"
    if tipo == "fragmento":
        return f"la referencia cubre solo un fragmento ({valor:.0%}) de la secuencia registrada"
    return f"coincidencia aproximada de {valor:.0%} de los residuos"


# --- Coherencia interna (sin red) -------------------------------------------
def verificar_coherencia(seq: str, valores: dict) -> dict:
    """Comprueba que los valores registrados sean coherentes con la secuencia."""
    p = compute_properties(seq)
    graves, avisos = [], []

    longitud = _a_float(valores.get("Longitud"))
    if longitud is not None and int(longitud) != p["longitud"]:
        graves.append(f"longitud registrada ({longitud:.0f}) distinta de la calculada ({p['longitud']})")

    masa = _a_float(valores.get("Masa molecular (Da)"))
    if masa is not None and abs(masa - p["masa"]) > max(1.0, 0.005 * p["masa"]):
        graves.append(f"masa registrada ({masa:.2f} Da) distinta de la calculada ({p['masa']:.2f} Da)")

    gravy = _a_float(valores.get("Hidrofobicidad (GRAVY)"))
    if gravy is not None and abs(gravy - p["gravy"]) > 0.01:
        avisos.append(f"GRAVY registrado ({gravy:.3f}) difiere del calculado ({p['gravy']:.3f}); "
                      "puede ser un valor manual u otra escala")

    carga = _a_float(valores.get("Carga neta aprox. (pH 7)"))
    if carga is not None and int(round(carga)) != p["carga"]:
        avisos.append(f"carga neta registrada ({carga:.0f}) difiere de la aproximada ({p['carga']})")

    polaridad = str(valores.get("Polaridad") or "").strip()
    if polaridad and polaridad != p["polaridad"]:
        avisos.append(f"polaridad registrada («{polaridad}») distinta de la estimada («{p['polaridad']}»)")

    for col, clave in (("% Apolares", "p_apolar"), ("% Polares sin carga", "p_polar"),
                       ("% Carga positiva", "p_pos"), ("% Carga negativa", "p_neg")):
        v = _a_float(valores.get(col))
        if v is not None and abs(v - p[clave]) > 0.2:
            avisos.append(f"{col} registrado ({v:.1f}) difiere del calculado ({p[clave]:.1f})")

    if p["longitud"] < 2:
        avisos.append("la secuencia tiene un solo residuo")
    elif p["longitud"] > 50:
        avisos.append(f"longitud de {p['longitud']} residuos: supera el rango habitual de péptidos "
                      "(hasta unos 50) y podría tratarse de una proteína")

    if graves:
        return _res(F_INTERNA, NIVEL_DISC, "; ".join(graves + avisos))
    if avisos:
        return _res(F_INTERNA, NIVEL_PARCIAL, "; ".join(avisos))
    return _res(F_INTERNA, NIVEL_OK,
                "longitud, masa y composición consistentes con la secuencia registrada")


# --- RCSB PDB ---------------------------------------------------------------
_GQL_ENTRADA = """query ($id: String!) {
  entry(entry_id: $id) {
    struct { title }
    rcsb_primary_citation { title year journal_abbrev pdbx_database_id_PubMed pdbx_database_id_DOI }
    polymer_entities {
      rcsb_polymer_entity { pdbx_description }
      entity_poly { pdbx_seq_one_letter_code_can rcsb_entity_polymer_type }
    }
  }
}"""


def consultar_rcsb_entrada(pdb_id: str, seq: str) -> dict:
    """Comprueba que la entrada exista y que su secuencia sea coherente con la registrada."""
    j = _http_json(URL_RCSB_GRAPHQL, {"query": _GQL_ENTRADA, "variables": {"id": pdb_id}})
    entrada = ((j or {}).get("data") or {}).get("entry")
    if not entrada:
        return _res(F_RCSB, NIVEL_DISC, f"El código {pdb_id} no existe en RCSB PDB.")

    titulo = (entrada.get("struct") or {}).get("title", "")
    cita = entrada.get("rcsb_primary_citation") or {}
    pmid = cita.get("pdbx_database_id_PubMed")

    candidatos = []
    for n, ent in enumerate(entrada.get("polymer_entities") or [], start=1):
        poly = ent.get("entity_poly") or {}
        if poly.get("rcsb_entity_polymer_type") != "Protein":
            continue
        ref = poly.get("pdbx_seq_one_letter_code_can") or ""
        descripcion = (ent.get("rcsb_polymer_entity") or {}).get("pdbx_description", "")
        tipo, valor = comparar_secuencias(seq, ref)
        largo = len(re.sub(r"[^A-Za-z]", "", ref))
        candidatos.append((_RANGO[tipo], valor, n, descripcion, tipo, largo))

    extra = {"pmid": pmid}
    if not candidatos:
        return _res(F_RCSB, NIVEL_DISC,
                    f"La entrada {pdb_id} existe ({titulo}) pero no contiene cadenas polipeptídicas.",
                    **extra)

    _, valor, n, descripcion, tipo, largo = max(candidatos, key=lambda c: (c[0], c[1]))
    comparacion = _texto_comparacion(tipo, valor, largo)
    detalle = f"La entrada {pdb_id} existe ({titulo}). Entidad {n} ({descripcion}): {comparacion}."
    if tipo in ("idéntica", "contenida"):
        nivel = NIVEL_OK
    elif tipo == "fragmento" or valor >= 0.9:
        nivel = NIVEL_PARCIAL
    else:
        nivel = NIVEL_DISC
        detalle += " Revisa que el código PDB corresponda a este péptido."
    return _res(F_RCSB, nivel, detalle, **extra)


def consultar_rcsb_busqueda(seq: str) -> dict:
    """Busca en RCSB estructuras con secuencia idéntica o muy similar (sin código PDB)."""
    consulta = {
        "query": {
            "type": "terminal",
            "service": "sequence",
            "parameters": {"evalue_cutoff": 0.1, "identity_cutoff": 0.9,
                           "sequence_type": "protein", "value": seq},
        },
        "request_options": {"paginate": {"start": 0, "rows": 5}, "scoring_strategy": "sequence"},
        "return_type": "polymer_entity",
    }
    j = _http_json(URL_RCSB_BUSQUEDA, consulta, timeout=20)
    resultados = (j or {}).get("result_set") or []
    if not resultados:
        return _res(F_RCSB_BUSQ, NIVEL_NADA,
                    "No se hallaron estructuras en RCSB PDB con secuencia similar (identidad mínima 90 %).")

    aciertos, mejor = [], 0.0
    for item in resultados[:5]:
        identidad = None
        for servicio in item.get("services") or []:
            for nodo in servicio.get("nodes") or []:
                for ctx in nodo.get("match_context") or []:
                    v = _a_float(ctx.get("sequence_identity"))
                    if v is not None:
                        identidad = v
        if identidad is not None:
            mejor = max(mejor, identidad)
            aciertos.append(f"{item.get('identifier', '?')} ({identidad:.0%})")
        else:
            aciertos.append(str(item.get("identifier", "?")))
    total = (j or {}).get("total_count", len(resultados))
    detalle = (f"{total} entidades con secuencia similar; las mejores: {', '.join(aciertos)}. "
               "Es una coincidencia de secuencia; no confirma el nombre del péptido.")
    return _res(F_RCSB_BUSQ, NIVEL_OK if mejor >= 0.99 else NIVEL_PARCIAL, detalle)


# --- UniProt ----------------------------------------------------------------
def consultar_uniprot(nombre: str, seq: str) -> dict:
    """Busca el nombre en UniProt (entradas revisadas) y compara las secuencias."""
    params = urlencode({
        "query": f"({nombre}) AND (reviewed:true)",
        "fields": "accession,protein_name,organism_name,sequence",
        "size": 10,
        "format": "json",
    })
    j = _http_json(f"{URL_UNIPROT}?{params}")
    resultados = (j or {}).get("results") or []
    if not resultados:
        return _res(F_UNIPROT, NIVEL_NADA,
                    f"No se hallaron entradas revisadas (Swiss-Prot) para «{nombre}».")

    mejor = None
    for r in resultados:
        ref = (r.get("sequence") or {}).get("value", "")
        tipo, valor = comparar_secuencias(seq, ref)
        clave = (_RANGO[tipo], valor)
        if mejor is None or clave > mejor[0]:
            nombre_prot = (((r.get("proteinDescription") or {}).get("recommendedName") or {})
                           .get("fullName") or {}).get("value", "")
            organismo = (r.get("organism") or {}).get("scientificName", "")
            mejor = (clave, r.get("primaryAccession", "?"), nombre_prot, organismo,
                     tipo, valor, len(ref))

    _, acc, nombre_prot, organismo, tipo, valor, largo = mejor
    ficha = f"{acc} ({nombre_prot}; {organismo})"
    comparacion = _texto_comparacion(tipo, valor, largo)
    if tipo in ("idéntica", "contenida"):
        return _res(F_UNIPROT, NIVEL_OK, f"Entrada {ficha}: {comparacion}.")
    if tipo == "fragmento":
        return _res(F_UNIPROT, NIVEL_PARCIAL, f"Entrada {ficha}: {comparacion}.")
    return _res(F_UNIPROT, NIVEL_NADA,
                f"Hay {len(resultados)} entradas con ese nombre, pero ninguna contiene la secuencia "
                f"registrada (mejor coincidencia: {ficha}, {comparacion}).")


# --- PubMed -----------------------------------------------------------------
def consultar_pubmed(nombre: str | None, pmid_pdb) -> dict:
    """Busca respaldo bibliográfico: la cita primaria de la entrada PDB y el nombre del péptido."""
    ids: list[str] = []
    total = None
    if pmid_pdb:
        ids.append(str(pmid_pdb))
    if nombre:
        termino = f'"{nombre}"[Title/Abstract] AND peptide*[Title/Abstract]'
        j = _http_json(URL_EUTILS + "esearch.fcgi?" + urlencode({
            "db": "pubmed", "retmode": "json", "retmax": 3, "sort": "relevance", "term": termino}))
        er = (j or {}).get("esearchresult") or {}
        total = int(er.get("count", 0) or 0)
        ids += [i for i in er.get("idlist", []) if i not in ids]

    if not ids:
        if nombre is None:
            return _res(F_PUBMED, NIVEL_OMIT,
                        "El identificador es un código genérico; no hay nombre que buscar en la literatura.")
        return _res(F_PUBMED, NIVEL_NADA, f"No se hallaron publicaciones sobre «{nombre}» y péptidos.")

    j = _http_json(URL_EUTILS + "esummary.fcgi?" + urlencode({
        "db": "pubmed", "retmode": "json", "id": ",".join(ids[:4])}))
    resumen = (j or {}).get("result") or {}
    referencias = []
    for uid in ids[:4]:
        r = resumen.get(uid)
        if not isinstance(r, dict) or not r.get("title"):
            continue
        titulo = r["title"].strip()
        titulo = titulo[:110] + "..." if len(titulo) > 110 else titulo
        anio = str(r.get("pubdate", ""))[:4]
        origen = "cita de la entrada PDB" if str(uid) == str(pmid_pdb) else "coincidencia por nombre"
        referencias.append(f"PMID {uid}, {titulo} ({r.get('source', '')}, {anio}; {origen})")

    if not referencias:
        return _res(F_PUBMED, NIVEL_NADA, "No se pudo obtener el detalle de las publicaciones.")
    encabezado = f"{total} publicaciones para «{nombre}». " if total is not None else ""
    return _res(F_PUBMED, NIVEL_OK, encabezado + "Referencias: " + "; ".join(referencias) + ".")


# --- Orquestación -----------------------------------------------------------
def _seguro(fuente: str, funcion, *args) -> dict:
    """Ejecuta una consulta y convierte cualquier fallo en un resultado «No disponible»."""
    try:
        return funcion(*args)
    except RuntimeError as exc:
        return _res(fuente, NIVEL_ERR, str(exc))
    except (ValueError, KeyError, TypeError, AttributeError):
        return _res(fuente, NIVEL_ERR, "Respuesta inesperada del servicio.")


def _consultas_externas(seq: str, nombre: str | None, pdb_id: str | None, hay_texto_pdb: bool) -> list:
    """Consulta RCSB, UniProt y PubMed (las dos primeras en paralelo)."""
    fuentes_omitidas = []
    tareas = []
    if pdb_id:
        tareas.append((F_RCSB, consultar_rcsb_entrada, (pdb_id, seq)))
    elif hay_texto_pdb:
        fuentes_omitidas.append(_res(F_RCSB, NIVEL_OMIT,
                                     "El texto del campo de estructura no es un código ni un enlace de RCSB."))
    elif len(seq) >= 25:
        tareas.append((F_RCSB_BUSQ, consultar_rcsb_busqueda, (seq,)))
    else:
        fuentes_omitidas.append(_res(F_RCSB_BUSQ, NIVEL_OMIT,
                                     "Sin código PDB; la búsqueda por secuencia en RCSB requiere "
                                     "al menos 25 residuos."))
    if nombre:
        tareas.append((F_UNIPROT, consultar_uniprot, (nombre, seq)))
    else:
        fuentes_omitidas.append(_res(F_UNIPROT, NIVEL_OMIT,
                                     "El identificador es un código genérico; no hay nombre que buscar."))

    with ThreadPoolExecutor(max_workers=3) as pool:
        futuros = [pool.submit(_seguro, f, fn, *args) for f, fn, args in tareas]
        resultados = [fu.result() for fu in futuros]

    pmid = next((r.get("pmid") for r in resultados if r["fuente"] == F_RCSB), None)
    resultados.append(_seguro(F_PUBMED, consultar_pubmed, nombre, pmid))

    orden = [F_RCSB, F_RCSB_BUSQ, F_UNIPROT, F_PUBMED]
    todos = resultados + fuentes_omitidas
    return sorted(todos, key=lambda r: orden.index(r["fuente"]))


def consolidar_estado(fuentes: list) -> str:
    """Resume los resultados por fuente en un único estado del registro."""
    niveles = [f["nivel"] for f in fuentes]
    interna = next((f["nivel"] for f in fuentes if f["fuente"] == F_INTERNA), NIVEL_OK)
    externas = [f for f in fuentes if f["fuente"] != F_INTERNA]
    niv_ext = [f["nivel"] for f in externas]

    if NIVEL_DISC in niveles:
        return ESTADO_DISC
    if any(f["nivel"] == NIVEL_OK for f in externas if f["fuente"] in FUENTES_SECUENCIA):
        return ESTADO_VERIFICADO if interna == NIVEL_OK else ESTADO_AVISOS
    if any(n in (NIVEL_OK, NIVEL_PARCIAL) for n in niv_ext):
        return ESTADO_PARCIAL
    if NIVEL_ERR in niv_ext and all(n in (NIVEL_ERR, NIVEL_OMIT) for n in niv_ext):
        return ESTADO_NC
    return ESTADO_SIN


def verificar_registro(nombre: str, seq, pdb_texto, valores: dict | None = None) -> dict:
    """Verifica un registro contra fuentes externas y contra sí mismo.

    Devuelve {"estado", "fuentes", "fecha"}. Las consultas externas se guardan en
    caché una hora (solo si ninguna falló).
    """
    valores = valores or {}
    pdb_texto = pdb_texto.strip() if isinstance(pdb_texto, str) else ""
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

    seq_norm, error = parse_sequence(seq if isinstance(seq, str) else "", FORMATO_1)
    if error:
        fuentes = [_res(F_INTERNA, NIVEL_DISC, f"El registro no tiene una secuencia válida: {error}")]
        return {"estado": ESTADO_DISC, "fuentes": fuentes, "fecha": fecha}

    nombre_inf = nombre_informativo(nombre)
    pdb_id = extraer_pdb_id(pdb_texto)
    clave = (seq_norm, nombre_inf or "", pdb_id or "", bool(pdb_texto))

    en_cache = _CACHE_VERIF.get(clave)
    if en_cache and time.time() - en_cache[0] < _TTL_VERIF:
        externas = en_cache[1]
    else:
        externas = _consultas_externas(seq_norm, nombre_inf, pdb_id, bool(pdb_texto))
        if all(f["nivel"] != NIVEL_ERR for f in externas):
            _CACHE_VERIF[clave] = (time.time(), externas)

    fuentes = [verificar_coherencia(seq_norm, valores)] + externas
    return {"estado": consolidar_estado(fuentes), "fuentes": fuentes, "fecha": fecha}


def notas_verificacion(res: dict) -> str:
    """Texto compacto para guardar en el registro."""
    return " | ".join(f"[{f['fuente']}] {f['nivel']}: {f['detalle']}" for f in res["fuentes"])


def _md(texto: str) -> str:
    """Evita que Markdown interprete símbolos de las notas."""
    return str(texto).replace("$", "\\$").replace("*", "\\*").replace("_", "\\_")


def mostrar_verificacion(res: dict) -> None:
    """Muestra el estado consolidado y el detalle por fuente."""
    getattr(st, ESTADO_TIPO.get(res["estado"], "info"))(f"Estado de verificación: {res['estado']}")
    for f in res["fuentes"]:
        st.markdown(f"**{_md(f['fuente'])}** - {_md(f['nivel'])}: {_md(f['detalle'])}")
    st.caption(f"Consulta realizada el {res['fecha']}. {AVISO_VERIFICACION}")


def mostrar_notas_guardadas(notas: str) -> None:
    for linea in [x for x in str(notas).split(" | ") if x.strip()]:
        st.markdown(f"- {_md(linea)}")


# ---------------------------------------------------------------------------
# 7. GENERACIÓN DE REPORTES (EXCEL, PDF, CSV)
# ---------------------------------------------------------------------------
ESTADISTICAS_EXCEL = ["Longitud", "Masa molecular (Da)", "Hidrofobicidad (GRAVY)",
                      "Carga neta aprox. (pH 7)", "% Apolares", "% Polares sin carga",
                      "% Carga positiva", "% Carga negativa"]


@st.cache_data(show_spinner=False)
def build_excel(df: pd.DataFrame, titulo: str, rut_fmt: str, incluir_resumen: bool) -> bytes:
    buffer = BytesIO()
    azul = PatternFill("solid", fgColor="1F4E78")
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Peptidos")
        ws = writer.sheets["Peptidos"]

        for celda in ws[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = azul
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for i, col in enumerate(df.columns, start=1):
            ancho = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str)])
            ws.column_dimensions[get_column_letter(i)].width = min(ancho + 2, 60)

        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                celda.alignment = Alignment(vertical="top", wrap_text=True)

        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions

        if incluir_resumen:
            hoja = writer.book.create_sheet("Resumen")
            hoja.append(["Título", titulo])
            hoja.append(["RUT del usuario", rut_fmt])
            hoja.append(["Fecha de generación", datetime.now().strftime("%Y-%m-%d %H:%M")])
            hoja.append(["Total de registros", len(df)])
            hoja.append([])
            hoja.append(["Variable", "Promedio", "Mínimo", "Máximo"])
            fila_cab = hoja.max_row
            for col in ESTADISTICAS_EXCEL:
                if col in df.columns:
                    serie = pd.to_numeric(df[col], errors="coerce").dropna()
                    if not serie.empty:
                        hoja.append([col, round(float(serie.mean()), 3),
                                     float(serie.min()), float(serie.max())])
            for celda in hoja[fila_cab]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = azul

            if COL_VERIF_ESTADO in df.columns:
                hoja.append([])
                hoja.append(["Estado de verificación", "Registros"])
                fila_verif = hoja.max_row
                estados = df[COL_VERIF_ESTADO].replace("", ESTADO_PEND).value_counts()
                for estado, n in estados.items():
                    hoja.append([estado, int(n)])
                for celda in hoja[fila_verif]:
                    celda.font = Font(bold=True, color="FFFFFF")
                    celda.fill = azul

            for fila in range(1, 5):
                hoja.cell(row=fila, column=1).font = Font(bold=True)
            hoja.column_dimensions["A"].width = 30
            for letra in "BCD":
                hoja.column_dimensions[letra].width = 18
            hoja.column_dimensions["B"].width = 45
    return buffer.getvalue()


# Las fuentes estándar del PDF (Helvetica) solo cubren Latin-1; se sustituyen
# símbolos comunes en notas estructurales (alfa-hélice, beta-lámina, etc.).
_SUSTITUCIONES_PDF = {
    "α": "alfa", "β": "beta", "γ": "gamma", "δ": "delta", "κ": "kappa",
    "−": "-", "–": "-", "—": "-", "≥": ">=", "≤": "<=", "→": "->",
    "µ": "u", "μ": "u",
}


def _pdf_txt(valor) -> str:
    """Convierte un valor en texto seguro para un Paragraph de ReportLab."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)) or str(valor).strip() == "":
        return "-"
    texto = str(valor)
    for origen, destino in _SUSTITUCIONES_PDF.items():
        texto = texto.replace(origen, destino)
    texto = texto.encode("cp1252", "replace").decode("cp1252")
    return escape(texto).replace("\n", "<br/>")


def _num(valor, formato: str) -> str:
    try:
        return format(float(valor), formato)
    except (TypeError, ValueError):
        return "-"


def _secuencia_en_bloques(seq, tam: int = 10) -> str:
    if not isinstance(seq, str) or not seq:
        return "-"
    bloques = " ".join(seq[i:i + tam] for i in range(0, len(seq), tam))
    return f'<font name="Courier">{escape(bloques)}</font>'


# (encabezado, columna, formato, ancho relativo en cm) de la tabla resumen del PDF
_RESUMEN_PDF = [
    ("Identificador", COL_ID, None, 5.0),
    ("Longitud", "Longitud", ".0f", 2.5),
    ("Masa (Da)", "Masa molecular (Da)", ".2f", 3.0),
    ("Polaridad", "Polaridad", None, 4.0),
    ("GRAVY", "Hidrofobicidad (GRAVY)", ".3f", 2.5),
    ("Carga neta", "Carga neta aprox. (pH 7)", ".0f", 2.5),
    ("Verificación", COL_VERIF_ESTADO, None, 4.2),
    ("PDB / enlace", COL_PDB, None, 4.5),
]


@st.cache_data(show_spinner=False)
def build_pdf(df: pd.DataFrame, titulo: str, rut_fmt: str,
              incluir_resumen: bool, incluir_fichas: bool) -> bytes:
    buffer = BytesIO()
    pagina = landscape(A4)
    ancho_util = pagina[0] - 3 * cm
    doc = SimpleDocTemplate(
        buffer, pagesize=pagina, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.8 * cm,
        title=titulo, author=f"RUT {rut_fmt}",
    )
    estilos = getSampleStyleSheet()
    celda = ParagraphStyle("celda", parent=estilos["BodyText"], fontSize=8, leading=10)
    cabecera = ParagraphStyle("cabecera", parent=celda, fontName="Helvetica-Bold",
                              textColor=colors.white)

    historia = [
        Paragraph(_pdf_txt(titulo), estilos["Title"]),
        Paragraph(f"RUT del usuario: {escape(rut_fmt)} &nbsp;|&nbsp; "
                  f"Generado el {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; "
                  f"Total de registros: {len(df)}", estilos["Normal"]),
        Spacer(1, 0.5 * cm),
    ]

    # --- Tabla resumen -----------------------------------------------------
    if incluir_resumen:
        spec = [s for s in _RESUMEN_PDF if s[1] in df.columns]
        factor = ancho_util / (sum(s[3] for s in spec) * cm)
        anchos = [s[3] * cm * factor for s in spec]
        filas = [[Paragraph(s[0], cabecera) for s in spec]]
        for _, r in df.iterrows():
            filas.append([
                Paragraph(_pdf_txt(r[s[1]]) if s[2] is None else _num(r[s[1]], s[2]), celda)
                for s in spec
            ])
        tabla = Table(filas, colWidths=anchos, repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3F8")]),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B0B7BF")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        historia += [Paragraph("Resumen", estilos["Heading2"]), tabla]

    # --- Una ficha por péptido --------------------------------------------
    if incluir_fichas:
        if incluir_resumen:
            historia.append(PageBreak())
        historia.append(Paragraph("Fichas detalladas", estilos["Heading2"]))
        campos = [c for c in df.columns if c != COL_ID]
        for i, (_, r) in enumerate(df.iterrows(), start=1):
            datos = []
            for campo in campos:
                valor = (_secuencia_en_bloques(r[campo]) if campo == "Secuencia"
                         else _pdf_txt(r[campo]))
                datos.append([Paragraph(f"<b>{escape(campo)}</b>", celda), Paragraph(valor, celda)])
            if not datos:  # solo se pidió el identificador
                datos = [[Paragraph("<b>Identificador</b>", celda),
                          Paragraph(_pdf_txt(r[COL_ID]), celda)]]
            ficha = Table(datos, colWidths=[6 * cm, ancho_util - 6 * cm])
            ficha.setStyle(TableStyle([
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C9CED4")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))
            historia.append(KeepTogether([
                Paragraph(f"{i}. {_pdf_txt(r[COL_ID])}", estilos["Heading3"]),
                ficha,
                Spacer(1, 0.4 * cm),
            ]))

    def pie(canvas, documento):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.drawString(1.5 * cm, 1 * cm, f"RUT {rut_fmt}")
        canvas.drawRightString(pagina[0] - 1.5 * cm, 1 * cm, f"Página {documento.page}")
        canvas.restoreState()

    doc.build(historia, onFirstPage=pie, onLaterPages=pie)
    return buffer.getvalue()


def seccion_exportacion(df: pd.DataFrame, prefijo: str, nombre_base: str) -> None:
    """Sección final de descarga: opciones del reporte y botones por formato."""
    st.subheader("Exportación de reportes")
    if df.empty:
        st.info("No hay registros para exportar.")
        return

    with st.expander("Opciones del reporte", expanded=True):
        titulo = st.text_input("Título del reporte", value=TITULO_REPORTE_DEFECTO,
                               key=f"{prefijo}_titulo").strip() or TITULO_REPORTE_DEFECTO
        columnas = st.multiselect(
            "Columnas a incluir (el identificador siempre se incluye)",
            COLUMNS_EXPORTABLES, default=COLUMNS_EXPORTABLES, key=f"{prefijo}_columnas",
        )
        o1, o2, o3 = st.columns(3)
        pdf_resumen = o1.checkbox("PDF: incluir tabla resumen", value=True, key=f"{prefijo}_pdf_res")
        pdf_fichas = o2.checkbox("PDF: incluir fichas detalladas", value=True, key=f"{prefijo}_pdf_fic")
        xls_resumen = o3.checkbox("Excel: incluir hoja de estadísticas", value=True,
                                  key=f"{prefijo}_xls_res")

    df_exp = df[[COL_ID] + columnas].reset_index(drop=True)
    rut_fmt = formatear_rut(rut_actual())
    marca = datetime.now().strftime("%Y%m%d_%H%M")
    st.caption(f"Se exportarán {len(df_exp)} registros guardados y {len(df_exp.columns)} columnas.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Excel (.xlsx)**")
        st.caption("Hoja de datos con formato y filtros, más hoja opcional de estadísticas.")
        st.download_button(
            "Descargar Excel", data=build_excel(df_exp, titulo, rut_fmt, xls_resumen),
            file_name=f"{nombre_base}_{marca}.xlsx", key=f"{prefijo}_xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with c2:
        st.markdown("**PDF**")
        st.caption("Informe estructurado con tabla resumen y una ficha por péptido.")
        if pdf_resumen or pdf_fichas:
            st.download_button(
                "Descargar PDF",
                data=build_pdf(df_exp, titulo, rut_fmt, pdf_resumen, pdf_fichas),
                file_name=f"{nombre_base}_{marca}.pdf", key=f"{prefijo}_pdf",
                mime="application/pdf",
            )
        else:
            st.warning("Selecciona la tabla resumen y/o las fichas para generar el PDF.")
    with c3:
        st.markdown("**CSV (.csv)**")
        st.caption("Datos planos, compatibles con R, Python y otras herramientas.")
        st.download_button(
            "Descargar CSV", data=df_exp.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{nombre_base}_{marca}.csv", key=f"{prefijo}_csv", mime="text/csv",
        )


# ---------------------------------------------------------------------------
# 8. PÁGINAS DE LA APLICACIÓN
# ---------------------------------------------------------------------------
def pantalla_login() -> None:
    st.title(TITULO_APP)
    st.subheader("Acceso")
    st.write("Ingresa tu RUT para continuar. Tus registros quedarán vinculados a él, "
             "se guardarán de forma permanente y solo tú podrás verlos y gestionarlos.")
    with st.form("form_login"):
        rut_in = st.text_input("RUT", placeholder="12.345.678-5")
        enviado = st.form_submit_button("Ingresar", type="primary")
    if enviado:
        rut, error = validar_rut(rut_in)
        if error:
            st.error(error)
        else:
            st.session_state["rut"] = rut
            st.rerun()


def pantalla_bienvenida() -> None:
    st.markdown(
        "Registra péptidos con sus propiedades fisicoquímicas, verifica la información con "
        "fuentes científicas públicas (**RCSB PDB**, **UniProt** y **PubMed**) y visualiza su "
        "estructura 3D. Cada registro se **guarda de forma permanente** en tu base de datos "
        "personal y puedes exportarlo a **Excel**, **PDF** o **CSV**.\n\n"
        "Elige una opción en el menú de arriba para comenzar."
    )
    st.metric("Registros guardados en tu base de datos", len(st.session_state.registros))


def construir_fila(s, guardar: bool) -> tuple[dict | None, str | None]:
    """Valida el formulario y arma la fila. Devuelve (fila, mensaje_de_error).

    Con guardar=False (vista previa de la verificación) no se exige un
    identificador único.
    """
    ident = s.f_id.strip()
    if guardar:
        if not ident:
            return None, "El nombre / identificador es obligatorio."
        if ident in set(s.registros[COL_ID].astype(str)):
            return None, f"Ya existe un registro con el identificador «{ident}»."
    seq, error = parse_sequence(s.f_seq, s.f_formato)
    if error:
        return None, f"Secuencia inválida: {error}"

    p = compute_properties(seq)
    comp_manual = s.f_comp_modo == MANUAL and s.f_comp_manual.strip()
    if s.f_comp_modo == MANUAL and not comp_manual:
        return None, "Has elegido composición manual, pero el campo está vacío."

    # Si el campo PDB contiene un código o enlace de RCSB, se guarda el código normalizado
    pdb_texto = s.f_pdb.strip()
    pdb_valor = extraer_pdb_id(pdb_texto) or pdb_texto

    fila = {
        COL_ID: ident or "(sin nombre)",
        "Secuencia": seq,
        "Longitud": p["longitud"],
        "Masa molecular (Da)": p["masa"],
        "Polaridad": p["polaridad"] if s.f_pol_modo == AUTO else s.f_pol_manual,
        "% Apolares": p["p_apolar"],
        "% Polares sin carga": p["p_polar"],
        "% Carga positiva": p["p_pos"],
        "% Carga negativa": p["p_neg"],
        "Composición de aminoácidos": p["composicion"] if s.f_comp_modo == AUTO else comp_manual,
        "Hidrofobicidad (GRAVY)": p["gravy"] if s.f_hid_modo == AUTO else float(s.f_hid_manual),
        "Carga neta aprox. (pH 7)": p["carga"],
        COL_PDB: pdb_valor,
        "Notas estructurales": s.f_estruc.strip(),
        "Propiedades complementarias": s.f_props.strip(),
        "Notas": s.f_notas.strip(),
        COL_VERIF_ESTADO: ESTADO_PEND,
        COL_VERIF_NOTAS: "",
        COL_VERIF_FECHA: "",
        "Fecha de registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
        COL_OWNER: rut_actual(),
    }
    return fila, None


def agregar_registro() -> None:
    """Callback del botón 'Guardar registro': valida, verifica, agrega y guarda en disco."""
    s = st.session_state
    fila, error = construir_fila(s, guardar=True)
    if error:
        flash("error", error)
        return

    estado = ESTADO_PEND
    if s.f_verif_auto:
        res = verificar_registro(fila[COL_ID], fila["Secuencia"], fila[COL_PDB], fila)
        fila[COL_VERIF_ESTADO] = estado = res["estado"]
        fila[COL_VERIF_NOTAS] = notas_verificacion(res)
        fila[COL_VERIF_FECHA] = res["fecha"]

    ident = fila[COL_ID]
    s.registros = concat_seguro([s.registros, pd.DataFrame([fila])])

    if not persistir():
        flash("warning", f"El registro «{ident}» se agregó a la sesión, pero no se pudo escribir "
                         "en disco. Revisa los permisos de la carpeta de la aplicación.")
    elif estado == ESTADO_DISC:
        flash("warning", f"Registro «{ident}» guardado, pero la verificación detectó discrepancias. "
                         "Revisa las notas en «Gestionar mis registros».")
    else:
        flash("success", f"Registro «{ident}» guardado de forma permanente. "
                         f"Estado de verificación: {estado}.")
    for k, v in DEFAULTS_FORM.items():  # limpiar formulario
        s[k] = v
    s.pop("_verif_prev", None)


def pagina_crear() -> None:
    st.header("Nuevo registro de péptido")
    show_flash()
    s = st.session_state

    # --- Identificación y secuencia ---------------------------------------
    st.subheader("1. Identificación y secuencia")
    st.text_input("Nombre / Identificador *", key="f_id",
                  placeholder="Ej. PEP-001 o Magainin 2")
    st.radio("Formato de la secuencia", [FORMATO_1, FORMATO_3], key="f_formato", horizontal=True)
    st.text_area("Secuencia de aminoácidos *", key="f_seq", height=100,
                 help="Se ignoran espacios, guiones y comas. Solo los 20 aminoácidos estándar.")

    seq, error = (None, None)
    if s.f_seq.strip():
        seq, error = parse_sequence(s.f_seq, s.f_formato)
    props = compute_properties(seq) if seq else None

    if error:
        st.warning(error)
    elif props:
        st.caption("Vista previa de valores calculados a partir de la secuencia")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Longitud", props["longitud"])
        m2.metric("Masa (Da)", f"{props['masa']:.2f}")
        m3.metric("GRAVY", f"{props['gravy']:.3f}")
        m4.metric("Carga neta aprox.", f"{props['carga']:+d}")
        st.caption(f"Secuencia normalizada (una letra): {seq}")

    # --- Propiedades fisicoquímicas ---------------------------------------
    st.subheader("2. Propiedades fisicoquímicas")
    col_pol, col_comp, col_hid = st.columns(3)

    with col_pol:
        st.radio("Polaridad", [AUTO, MANUAL], key="f_pol_modo")
        if s.f_pol_modo == AUTO:
            st.info(props["polaridad"] if props else "Se calculará con una secuencia válida.")
        else:
            st.selectbox("Polaridad (manual)", POLARIDADES, key="f_pol_manual")

    with col_comp:
        st.radio("Composición de aminoácidos", [AUTO, MANUAL], key="f_comp_modo")
        if s.f_comp_modo == AUTO:
            st.info(props["composicion"] if props else "Se calculará con una secuencia válida.")
        else:
            st.text_area("Composición (manual)", key="f_comp_manual", height=100,
                         placeholder="Ej. A: 3 (15%); L: 5 (25%) ...")

    with col_hid:
        st.radio("Hidrofobicidad (índice hidropático)", [AUTO, MANUAL], key="f_hid_modo")
        if s.f_hid_modo == AUTO:
            st.info(f"GRAVY (Kyte-Doolittle): {props['gravy']:.3f}" if props
                    else "Se calculará con una secuencia válida.")
        else:
            st.number_input("Índice hidropático (manual)", key="f_hid_manual", step=0.01, format="%.3f")

    # --- Estructura 3D -----------------------------------------------------
    st.subheader("3. Estructura tridimensional")
    st.text_input("Código o enlace de RCSB PDB (opcional)", key="f_pdb",
                  placeholder="Ej. 1CRN o https://www.rcsb.org/structure/1CRN")
    pdb_id = extraer_pdb_id(s.f_pdb)
    if pdb_id:
        st.caption(f"Código RCSB PDB reconocido: {pdb_id}")
        st.checkbox("Previsualizar la estructura 3D", key="f_ver3d")
        if s.f_ver3d:
            mostrar_visor_3d(pdb_id, "form")
    elif s.f_pdb.strip():
        st.warning("El texto no se reconoce como código o enlace de RCSB PDB; "
                   "se guardará tal cual, sin visualización 3D.")
    st.text_area("Descripción / notas estructurales (opcional)", key="f_estruc", height=80,
                 placeholder="Ej. alfa-hélice anfipática en medio lipídico; método RMN...")

    # --- Campos adicionales -----------------------------------------------
    st.subheader("4. Información complementaria")
    st.text_area("Propiedades complementarias (opcional)", key="f_props", height=80,
                 placeholder="Ej. punto isoeléctrico, solubilidad, actividad biológica...")
    st.text_area("Notas (opcional)", key="f_notas", height=80)

    # --- Verificación científica -------------------------------------------
    st.subheader("5. Verificación con fuentes científicas")
    st.caption("Contrasta el registro con RCSB PDB (estructura y secuencia), UniProt (secuencia "
               "por nombre) y PubMed (respaldo bibliográfico), y comprueba que los valores "
               "calculados sean coherentes con la secuencia.")
    st.checkbox("Verificar automáticamente al guardar (puede tardar unos segundos)",
                key="f_verif_auto")

    clave_actual = (seq or "", s.f_id.strip(), s.f_pdb.strip())
    if st.button("Verificar ahora (vista previa)"):
        fila_prev, err_prev = construir_fila(s, guardar=False)
        if err_prev:
            st.warning(err_prev)
        else:
            with st.spinner("Consultando RCSB PDB, UniProt y PubMed..."):
                res_prev = verificar_registro(s.f_id.strip(), fila_prev["Secuencia"],
                                              fila_prev[COL_PDB], fila_prev)
            s["_verif_prev"] = {"clave": clave_actual, "res": res_prev}
    previa = s.get("_verif_prev")
    if previa:
        if previa["clave"] == clave_actual:
            mostrar_verificacion(previa["res"])
        else:
            st.caption("La verificación mostrada quedó desactualizada por cambios en el nombre, "
                       "la secuencia o el código PDB. Vuelve a ejecutarla.")

    st.divider()
    st.button("Guardar registro en mi base de datos", type="primary", on_click=agregar_registro)
    st.caption("Al guardar, el registro se almacena de forma permanente y estará disponible "
               "la próxima vez que ingreses con tu RUT.")

    if not s.registros.empty:
        st.divider()
        st.caption(f"Últimos registros guardados ({len(s.registros)} en total)")
        st.dataframe(s.registros.tail(5), width="stretch", hide_index=True,
                     column_config=CONFIG_OCULTA)


def invalidar_verificaciones(nuevo: pd.DataFrame, previo: pd.DataFrame) -> pd.DataFrame:
    """Si cambió la secuencia o el código PDB de un registro, su verificación queda pendiente."""
    if previo.empty:
        return nuevo
    base = previo.drop_duplicates(COL_ID).set_index(COL_ID)[["Secuencia", COL_PDB]].to_dict("index")
    nuevo = nuevo.copy()
    for idx, r in nuevo.iterrows():
        anterior = base.get(r[COL_ID])
        if anterior and (r["Secuencia"] != anterior["Secuencia"] or r[COL_PDB] != anterior[COL_PDB]):
            nuevo.loc[idx, COL_VERIF_ESTADO] = ESTADO_PEND
            nuevo.loc[idx, COL_VERIF_NOTAS] = ""
            nuevo.loc[idx, COL_VERIF_FECHA] = ""
    return nuevo


def ejecutar_verificacion(ids: list[str]) -> None:
    """Verifica los registros indicados, guarda el resultado en disco y recarga la página."""
    s = st.session_state
    df = s.registros.copy()
    barra = st.progress(0.0, text="Verificando registros...")
    for i, ident in enumerate(ids, start=1):
        pos = df.index[df[COL_ID] == ident]
        if len(pos) == 0:
            continue
        r = df.loc[pos[0]]
        res = verificar_registro(r[COL_ID], r["Secuencia"], r[COL_PDB], r.to_dict())
        df.loc[pos[0], COL_VERIF_ESTADO] = res["estado"]
        df.loc[pos[0], COL_VERIF_NOTAS] = notas_verificacion(res)
        df.loc[pos[0], COL_VERIF_FECHA] = res["fecha"]
        barra.progress(i / len(ids), text=f"Verificando {i} de {len(ids)}")
        if i < len(ids):
            time.sleep(0.4)  # respeta el límite de solicitudes de NCBI (PubMed)
    s.registros = limpiar(df)
    s.editor_version += 1
    if persistir():
        flash("success", f"Verificación completada para {len(ids)} registro(s).")
    else:
        flash("error", "Se verificaron los registros, pero no se pudo escribir en disco.")
    st.rerun()


def seccion_verificacion() -> None:
    s = st.session_state
    st.subheader("Verificación científica")
    st.caption("Contrasta los registros con RCSB PDB, UniProt y PubMed. Guarda antes los cambios "
               "pendientes de la tabla, porque la verificación trabaja sobre los datos guardados. "
               + AVISO_VERIFICACION)

    estados = s.registros[COL_VERIF_ESTADO].replace("", ESTADO_PEND).value_counts()
    st.caption("Estados actuales: " + " | ".join(f"{k}: {v}" for k, v in estados.items()))

    ids = st.multiselect("Registros a verificar (si no eliges ninguno, se verifican los pendientes)",
                         list(s.registros[COL_ID]))
    if st.button("Ejecutar verificación"):
        if ids:
            objetivo = ids
        else:
            pendientes = s.registros[s.registros[COL_VERIF_ESTADO].isin(["", ESTADO_PEND])]
            objetivo = list(pendientes[COL_ID])
        if not objetivo:
            st.info("No hay registros pendientes de verificación.")
        else:
            ejecutar_verificacion(objetivo)

    ver = st.selectbox("Ver las notas de verificación de un registro", list(s.registros[COL_ID]),
                       index=None, placeholder="Selecciona un registro…")
    if ver:
        fila = s.registros[s.registros[COL_ID] == ver].iloc[0]
        estado = fila[COL_VERIF_ESTADO] or ESTADO_PEND
        getattr(st, ESTADO_TIPO.get(estado, "info"))(f"Estado de verificación: {estado}")
        if fila[COL_VERIF_NOTAS]:
            mostrar_notas_guardadas(fila[COL_VERIF_NOTAS])
            st.caption(f"Verificado el {fila[COL_VERIF_FECHA]}.")
        else:
            st.caption("Este registro aún no tiene notas de verificación.")


def pagina_gestionar() -> None:
    st.header("Mis registros")
    show_flash()
    s = st.session_state
    if s.registros.empty:
        st.info("Aún no tienes registros. Crea uno nuevo o importa un archivo.")
        return

    st.caption("Doble clic en una celda para editar. Selecciona filas y pulsa Supr para eliminarlas. "
               "Los cambios se guardan de forma permanente al pulsar «Guardar cambios». "
               "Las columnas de verificación se actualizan solo mediante la verificación automática; "
               "si cambias la secuencia o el código PDB, la verificación del registro queda pendiente.")
    editado = st.data_editor(
        s.registros, num_rows="dynamic", width="stretch", hide_index=True,
        key=f"editor_{s.editor_version}", column_config=CONFIG_OCULTA, disabled=COLS_VERIF,
    )
    editado = limpiar(editado)

    if st.button("Guardar cambios", type="primary"):
        if editado[COL_ID].duplicated().any():
            st.error("Hay identificadores duplicados. Corrígelos antes de guardar.")
        else:
            s.registros = invalidar_verificaciones(editado, s.registros)
            s.editor_version += 1  # reinicia el editor para evitar ediciones "fantasma"
            if persistir():
                flash("success", f"Cambios guardados: {len(editado)} registros en tu base de datos.")
            else:
                flash("error", "No se pudo escribir en disco. Revisa los permisos de la carpeta.")
            st.rerun()

    with st.expander("Eliminar todos mis registros"):
        if st.checkbox("Confirmo que quiero eliminar de forma permanente todos mis registros"):
            if st.button("Eliminar todo"):
                s.registros = pd.DataFrame(columns=COLUMNS)
                s.editor_version += 1
                persistir()
                flash("success", "Todos los registros fueron eliminados.")
                st.rerun()

    st.divider()
    seccion_verificacion()

    st.divider()
    seccion_exportacion(s.registros, "mis_registros", "peptidos_dataset")


def pagina_visor() -> None:
    st.header("Visualización de estructura 3D (RCSB PDB)")
    s = st.session_state
    origen = st.radio("Origen de la estructura",
                      ["Registro guardado", "Código o enlace manual"], horizontal=True)

    pdb_id = None
    if origen == "Registro guardado":
        con_pdb = s.registros[s.registros[COL_PDB].map(extraer_pdb_id).notna()]
        if con_pdb.empty:
            st.info("Ninguno de tus registros tiene un código RCSB PDB válido. "
                    "Agrega uno al crear o editar un registro, o usa la opción manual.")
            return
        etiquetas = {f"{r[COL_ID]}  ({extraer_pdb_id(r[COL_PDB])})": extraer_pdb_id(r[COL_PDB])
                     for _, r in con_pdb.iterrows()}
        eleccion = st.selectbox("Registro", list(etiquetas))
        pdb_id = etiquetas[eleccion]
    else:
        texto = st.text_input("Código o enlace de RCSB PDB",
                              placeholder="Ej. 1CRN o https://www.rcsb.org/structure/1CRN")
        if texto.strip():
            pdb_id = extraer_pdb_id(texto)
            if pdb_id is None:
                st.error("No se reconoce el código o enlace. Usa un código de 4 caracteres "
                         "(ej. 1CRN) o un enlace de rcsb.org.")

    if pdb_id:
        mostrar_visor_3d(pdb_id, "visor")


def pagina_importar() -> None:
    st.header("Importar registros")
    st.caption(f"El archivo debe contener al menos la columna «{COL_ID}». "
               "Las columnas ausentes se dejarán vacías. Los registros importados quedarán "
               "vinculados a tu RUT y se guardarán de forma permanente. Lo ideal es importar un "
               "archivo exportado por esta aplicación. Los registros importados sin estado de "
               "verificación quedan pendientes; puedes verificarlos en «Gestionar mis registros».")
    archivo = st.file_uploader("Archivo CSV o Excel", type=["csv", "xlsx"])
    if archivo is None:
        return
    try:
        df = (pd.read_csv(archivo, encoding="utf-8-sig") if archivo.name.lower().endswith(".csv")
              else pd.read_excel(archivo))
    except Exception as exc:  # noqa: BLE001 - mostrar cualquier error de lectura al usuario
        st.error(f"No se pudo leer el archivo: {exc}")
        return
    if COL_ID not in df.columns:
        st.error(f"No se encontró la columna «{COL_ID}».")
        return

    df = limpiar(df)
    st.dataframe(df, width="stretch", hide_index=True, column_config=CONFIG_OCULTA)
    if st.button("Importar y guardar en mi base de datos", type="primary"):
        actuales = set(st.session_state.registros[COL_ID].astype(str))
        nuevos = df[~df[COL_ID].astype(str).isin(actuales)]
        st.session_state.registros = concat_seguro([st.session_state.registros, nuevos])
        st.session_state.editor_version += 1
        if persistir():
            st.success(f"{len(nuevos)} registros importados y guardados; "
                       f"{len(df) - len(nuevos)} omitidos por identificador repetido.")
        else:
            st.error("Los registros se importaron en la sesión, pero no se pudo escribir en disco.")


# ---------------------------------------------------------------------------
# 9. PUNTO DE ENTRADA
# ---------------------------------------------------------------------------
def main() -> None:
    # Puerta de acceso: sin RUT válido no se muestra nada más
    if "rut" not in st.session_state:
        pantalla_login()
        return

    init_state()
    st.title(TITULO_APP)

    with st.sidebar:
        st.subheader("Sesión")
        st.write(f"RUT: {formatear_rut(rut_actual())}")
        st.metric("Registros guardados", len(st.session_state.registros))
        st.caption("Tus registros se guardan en disco y se recuperan al volver a ingresar.")
        st.button("Cerrar sesión", on_click=cerrar_sesion)

    opcion = st.selectbox("¿Qué deseas realizar?", OPCIONES_MENU, index=None,
                          placeholder="Selecciona una opción…", key="menu")

    st.divider()
    if opcion is None:
        pantalla_bienvenida()
    elif opcion == OPCIONES_MENU[0]:
        pagina_crear()
    elif opcion == OPCIONES_MENU[1]:
        pagina_gestionar()
    elif opcion == OPCIONES_MENU[2]:
        pagina_visor()
    else:
        pagina_importar()


main()