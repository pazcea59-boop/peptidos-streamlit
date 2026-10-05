# Dataset de propiedades fisicoquímicas de péptidos

Aplicación Streamlit para registrar péptidos, editarlos en una tabla interactiva,
guardarlos en un dataset local por usuario y exportarlos a Excel, PDF y CSV.
El acceso requiere un RUT chileno válido (módulo 11).

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

No se agregaron librerías nuevas: la validación del RUT usa solo la biblioteca estándar de Python.

## Estructura

```
peptidos-streamlit/
├── app.py              # aplicación completa
├── requirements.txt
├── README.md
├── .gitignore          # excluye data/ (datos vinculados a RUT)
└── data/usuarios/      # se crea sola: una carpeta por usuario
```

## Acceso por RUT

- Al iniciar se solicita el RUT; se acepta con o sin puntos y guion, y con K mayúscula o minúscula.
- El dígito verificador se valida con módulo 11 (factores 2 a 7, de derecha a izquierda).
- Cada registro lleva el RUT de su propietario y el dataset se guarda en una carpeta separada por usuario
  (nombrada con una huella del RUT, no con el RUT). Cada usuario solo ve y gestiona sus propios registros.
- «Cerrar sesión» borra la sesión, incluida la tabla de trabajo temporal.

Limitación: el RUT valida el formato, no la identidad. Quien conozca un RUT podría ingresar con él.
Si la app manejará datos sensibles, agrega una clave o un proveedor de autenticación (por ejemplo OIDC).

## Flujo de uso

1. **Crear nuevo registro**: nombre y secuencia (1 o 3 letras) son obligatorios; el resto es opcional.
2. **Gestionar tabla de trabajo**: editar, eliminar, guardar en el dataset y exportar.
3. **Ver registros guardados**: buscar, recargar en la tabla de trabajo, eliminar y exportar.
4. **Importar**: cargar un CSV o Excel (los registros quedan vinculados al RUT de la sesión).

## Exportación de reportes

Sección disponible al final de la tabla de trabajo y de los registros guardados:

- Título del reporte y selección de columnas.
- **Excel (.xlsx)**: hoja de datos con formato y filtros, más hoja opcional de estadísticas.
- **PDF**: tabla resumen y/o fichas detalladas por péptido, con el RUT en encabezado y pie.
- **CSV**: datos planos.

## Cálculos automáticos

- **Hidrofobicidad**: GRAVY con la escala de Kyte-Doolittle.
- **Masa molecular**: suma de masas promedio de residuos más agua.
- **Carga neta aprox. (pH 7)**: (K + R) - (D + E); ignora His y los extremos N/C.
- **Polaridad** (heurística): 60 % o más de residuos apolares, apolar; 40 % o menos, polar; en medio, anfipático/mixto.
- Clases: apolares `GAVLIMFWP`, polares sin carga `STCYNQ`, positivos `KRH`, negativos `DE`.

Todos los campos calculados pueden sustituirse con entrada manual.

## Notas

- La tabla de trabajo es temporal. Usa «Guardar definitivamente» para persistirla.
- En Streamlit Community Cloud el disco es efímero: descarga tu Excel o CSV con regularidad.
- El PDF usa fuentes estándar (Latin-1): símbolos como alfa o beta se transcriben como «alfa» y «beta».
