
"""
Registro de propiedades fisicoquímicas de péptidos (Streamlit)
================================================================
Aplicación para registrar, editar, guardar y exportar (Excel / PDF / CSV) un
dataset de péptidos, con:
  - acceso mediante RUT chileno (validación módulo 11),
  - persistencia por usuario: los registros se guardan en disco y se recuperan
    al volver a ingresar con el mismo RUT,
  - visualización 3D interactiva de estructuras desde RCSB PDB.
 
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
import urllib.error
import urllib.request
from collections import Counter
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse
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
 
st.set_page_config(page_title="Dataset de péptidos", layout="wide")
 
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
 
TITULO_REPORTE_DEFECTO = "Dataset de propiedades fisicoquímicas de péptidos"
 
# Valores por defecto de los widgets del formulario (también sirven para limpiarlo)
DEFAULTS_FORM = {
    "f_id": "", "f_formato": FORMATO_1, "f_seq": "",
    "f_pol_modo": AUTO, "f_pol_manual": POLARIDADES[0],
    "f_comp_modo": AUTO, "f_comp_manual": "",
    "f_hid_modo": AUTO, "f_hid_manual": 0.0,
    "f_pdb": "", "f_ver3d": False, "f_estruc": "", "f_props": "", "f_notas": "",
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
# 6. GENERACIÓN DE REPORTES (EXCEL, PDF, CSV)
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
    ("Polaridad", "Polaridad", None, 4.5),
    ("GRAVY", "Hidrofobicidad (GRAVY)", ".3f", 2.5),
    ("Carga neta", "Carga neta aprox. (pH 7)", ".0f", 2.5),
    ("PDB / enlace", COL_PDB, None, 5.2),
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
# 7. PÁGINAS DE LA APLICACIÓN
# ---------------------------------------------------------------------------
def pantalla_login() -> None:
    st.title("Dataset de propiedades fisicoquímicas de péptidos")
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
        "Registra péptidos con sus propiedades fisicoquímicas y visualiza su estructura 3D. "
        "Cada registro se **guarda de forma permanente** en tu dataset personal y puedes "
        "exportarlo a **Excel**, **PDF** o **CSV**.\n\n"
        "Elige una opción en el menú de arriba para comenzar."
    )
    st.metric("Registros guardados en tu dataset", len(st.session_state.registros))
 
 
def agregar_registro() -> None:
    """Callback del botón 'Guardar registro': valida, calcula, agrega y guarda en disco."""
    s = st.session_state
    ident = s.f_id.strip()
    if not ident:
        flash("error", "El nombre / identificador es obligatorio.")
        return
    if ident in set(s.registros[COL_ID].astype(str)):
        flash("error", f"Ya existe un registro con el identificador «{ident}».")
        return
    seq, error = parse_sequence(s.f_seq, s.f_formato)
    if error:
        flash("error", f"Secuencia inválida: {error}")
        return
 
    p = compute_properties(seq)
    comp_manual = s.f_comp_manual.strip()
    if s.f_comp_modo == MANUAL and not comp_manual:
        flash("error", "Has elegido composición manual, pero el campo está vacío.")
        return
 
    # Si el campo PDB contiene un código o enlace de RCSB, se guarda el código normalizado
    pdb_texto = s.f_pdb.strip()
    pdb_valor = extraer_pdb_id(pdb_texto) or pdb_texto
 
    fila = {
        COL_ID: ident,
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
        "Fecha de registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
        COL_OWNER: rut_actual(),
    }
    s.registros = concat_seguro([s.registros, pd.DataFrame([fila])])
 
    if persistir():
        flash("success", f"Registro «{ident}» guardado de forma permanente en tu dataset.")
    else:
        flash("warning", f"El registro «{ident}» se agregó a la sesión, pero no se pudo escribir "
                         "en disco. Revisa los permisos de la carpeta de la aplicación.")
    for k, v in DEFAULTS_FORM.items():  # limpiar formulario
        s[k] = v
 
 
def pagina_crear() -> None:
    st.header("Nuevo registro de péptido")
    show_flash()
    s = st.session_state
 
    # --- Identificación y secuencia ---------------------------------------
    st.subheader("1. Identificación y secuencia")
    st.text_input("Nombre / Identificador *", key="f_id",
                  placeholder="Ej. PEP-001 o Magainina 2")
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
 
    st.button("Guardar registro en mi dataset", type="primary", on_click=agregar_registro)
    st.caption("Al guardar, el registro se almacena de forma permanente y estará disponible "
               "la próxima vez que ingreses con tu RUT.")
 
    if not s.registros.empty:
        st.divider()
        st.caption(f"Últimos registros guardados ({len(s.registros)} en total)")
        st.dataframe(s.registros.tail(5), width="stretch", hide_index=True,
                     column_config=CONFIG_OCULTA)
 
 
def pagina_gestionar() -> None:
    st.header("Mis registros")
    show_flash()
    s = st.session_state
    if s.registros.empty:
        st.info("Aún no tienes registros. Crea uno nuevo o importa un archivo.")
        return
 
    st.caption("Doble clic en una celda para editar. Selecciona filas y pulsa Supr para eliminarlas. "
               "Los cambios se guardan de forma permanente al pulsar «Guardar cambios».")
    editado = st.data_editor(
        s.registros, num_rows="dynamic", width="stretch", hide_index=True,
        key=f"editor_{s.editor_version}", column_config=CONFIG_OCULTA,
    )
    editado = limpiar(editado)
 
    if st.button("Guardar cambios", type="primary"):
        if editado[COL_ID].duplicated().any():
            st.error("Hay identificadores duplicados. Corrígelos antes de guardar.")
        else:
            s.registros = editado
            s.editor_version += 1  # reinicia el editor para evitar ediciones "fantasma"
            if persistir():
                flash("success", f"Cambios guardados: {len(editado)} registros en tu dataset.")
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
               "archivo exportado por esta aplicación.")
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
    if st.button("Importar y guardar en mi dataset", type="primary"):
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
# 8. PUNTO DE ENTRADA
# ---------------------------------------------------------------------------
def main() -> None:
    # Puerta de acceso: sin RUT válido no se muestra nada más
    if "rut" not in st.session_state:
        pantalla_login()
        return
 
    init_state()
    st.title("Dataset de propiedades fisicoquímicas de péptidos")
 
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
 