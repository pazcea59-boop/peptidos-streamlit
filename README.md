# 🧬 Dataset de propiedades fisicoquímicas de péptidos

Aplicación Streamlit para registrar péptidos, editarlos en una tabla interactiva,
guardarlos en un dataset local (CSV) y exportarlos a **Excel** y **PDF**.

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

```
peptidos-streamlit/
├── app.py              # aplicación completa
├── requirements.txt
├── README.md
├── .gitignore
└── data/               # se crea sola; contiene peptidos_guardados.csv
```

## Flujo de uso

1. **Crear nuevo registro**: nombre y secuencia (1 o 3 letras) son obligatorios; el resto es opcional.
2. **Gestionar tabla de trabajo**: editar, eliminar, guardar en el CSV y exportar.
3. **Ver registros guardados**: buscar, recargar en la tabla de trabajo, eliminar y exportar.
4. **Importar**: cargar un CSV/Excel (por ejemplo, uno exportado antes).

## Cálculos automáticos

- **Hidrofobicidad**: GRAVY con la escala de Kyte-Doolittle.
- **Masa molecular**: suma de masas promedio de residuos + agua.
- **Carga neta aprox. (pH 7)**: (K + R) − (D + E); ignora His y los extremos N/C.
- **Polaridad** (heurística): ≥ 60 % residuos apolares → apolar; ≤ 40 % → polar; en medio → anfipático/mixto.
- Clases: apolares `GAVLIMFWP`, polares sin carga `STCYNQ`, positivos `KRH`, negativos `DE`.

Todos los campos calculados pueden sustituirse con entrada manual en el formulario.

## Notas

- La tabla de trabajo es temporal (sesión del navegador). Usa «Guardar definitivamente» para persistirla.
- En Streamlit Community Cloud el disco es efímero: descarga tu Excel/CSV con regularidad.
- El PDF usa fuentes estándar (Latin-1): símbolos como α o β se transcriben como «alfa» / «beta».
