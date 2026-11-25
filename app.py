"""Aplicación Streamlit para calcular la duración entre dos columnas de Excel."""

from datetime import date, datetime, time, timedelta, timezone
from io import BytesIO
from zipfile import BadZipFile

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException


st.set_page_config(page_title="Duración de partidas", page_icon="⏱️", layout="wide")


@st.cache_data(show_spinner=False)
def obtener_hojas(contenido):
    """Devuelve los nombres de las hojas usando exclusivamente openpyxl."""
    libro = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    hojas = libro.sheetnames
    libro.close()
    return hojas


@st.cache_data(show_spinner="Leyendo hoja con openpyxl…")
def leer_hoja(contenido, nombre_hoja):
    """Carga los valores de la hoja seleccionada, sin usar pandas para leer Excel."""
    libro = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    hoja = libro[nombre_hoja]
    filas = tuple(tuple(fila) for fila in hoja.iter_rows(values_only=True))
    libro.close()
    return filas


def nombres_de_columnas(encabezados):
    """Completa encabezados vacíos y hace únicos los nombres repetidos."""
    nombres = []
    cantidades = {}
    for indice, valor in enumerate(encabezados, start=1):
        base = str(valor).strip() if valor is not None else ""
        base = base or f"Columna {indice}"
        cantidades[base] = cantidades.get(base, 0) + 1
        nombre = base if cantidades[base] == 1 else f"{base} ({cantidades[base]})"
        nombres.append(nombre)
    return nombres


def convertir_fecha_hora(valor):
    """Convierte valores de Excel o texto a datetime e indica si solo tenían hora."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        raise ValueError("celda vacía")

    if isinstance(valor, timedelta):
        referencia = datetime(1900, 1, 1)
        return referencia + valor, True
    if isinstance(valor, datetime):
        return valor, False
    if isinstance(valor, date):
        return datetime.combine(valor, time.min), False
    if isinstance(valor, time):
        return datetime.combine(date(1900, 1, 1), valor), True

    texto = str(valor).strip()
    try:
        fecha_hora = datetime.fromisoformat(texto.replace("Z", "+00:00"))
        if fecha_hora.tzinfo is not None:
            fecha_hora = fecha_hora.astimezone(timezone.utc).replace(tzinfo=None)
        return fecha_hora, False
    except ValueError:
        pass

    formatos_hora = ("%H:%M:%S.%f", "%H:%M:%S", "%H:%M")
    for formato in formatos_hora:
        try:
            return datetime.combine(date(1900, 1, 1), datetime.strptime(texto, formato).time()), True
        except ValueError:
            continue

    formatos_fecha = (
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
        "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y",
    )
    for formato in formatos_fecha:
        try:
            return datetime.strptime(texto, formato), False
        except ValueError:
            continue

    raise ValueError(f"formato de fecha/hora no reconocido: {texto!r}")


def calcular_duracion(inicio, fin):
    """Calcula la duración; para horas sin fecha interpreta el cruce de medianoche."""
    fecha_inicio, inicio_solo_hora = convertir_fecha_hora(inicio)
    fecha_fin, fin_solo_hora = convertir_fecha_hora(fin)

    if inicio_solo_hora and not fin_solo_hora:
        fecha_inicio = datetime.combine(fecha_fin.date(), fecha_inicio.time())
        if fecha_inicio > fecha_fin:
            fecha_inicio -= timedelta(days=1)
    elif fin_solo_hora and not inicio_solo_hora:
        fecha_fin = datetime.combine(fecha_inicio.date(), fecha_fin.time())
        if fecha_fin < fecha_inicio:
            fecha_fin += timedelta(days=1)
    elif inicio_solo_hora and fin_solo_hora and fecha_fin < fecha_inicio:
        fecha_fin += timedelta(days=1)

    duracion = fecha_fin - fecha_inicio
    if duracion.total_seconds() < 0:
        raise ValueError("la hora/fecha de fin es anterior a la de inicio")
    return duracion


def formato_hh_mm_ss(segundos):
    """Formatea segundos sin limitar las horas a 24."""
    segundos = int(round(segundos))
    horas, resto = divmod(segundos, 3600)
    minutos, segundos = divmod(resto, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"


def sugerir_columna(columnas, palabras, excluir=None):
    for columna in columnas:
        if columna != excluir and any(palabra in columna.lower() for palabra in palabras):
            return columna
    return next((columna for columna in columnas if columna != excluir), columnas[0])


def crear_excel(columnas, filas):
    """Genera un nuevo libro de resultados con openpyxl."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Resultados"
    hoja.append(columnas)
    for fila in filas:
        hoja.append(fila)
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.columns:
        letra = columna[0].column_letter
        ancho = min(max(max(len(str(celda.value or "")) for celda in columna) + 2, 12), 40)
        hoja.column_dimensions[letra].width = ancho
    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()


st.title("⏱️ Calculadora de duración de partidas")
st.write(
    "Carga un archivo Excel, selecciona las columnas de inicio y fin, "
    "y obtén la duración de cada partida o sesión."
)

archivo = st.file_uploader("Selecciona un archivo Excel (.xlsx)", type=["xlsx"])
if archivo is None:
    st.info("Carga un archivo .xlsx para comenzar.")
    st.stop()

contenido = archivo.getvalue()
try:
    hojas = obtener_hojas(contenido)
except (BadZipFile, InvalidFileException, OSError, KeyError) as error:
    st.error(f"No se pudo abrir el archivo Excel. Comprueba que no esté dañado: {error}")
    st.stop()
except Exception as error:
    st.error(f"Ocurrió un error al leer el archivo: {error}")
    st.stop()

if not hojas:
    st.error("El archivo no contiene hojas de cálculo.")
    st.stop()

nombre_hoja = st.selectbox("Hoja de cálculo", hojas)
try:
    filas_excel = leer_hoja(contenido, nombre_hoja)
except (BadZipFile, InvalidFileException, OSError, KeyError) as error:
    st.error(f"No se pudo leer la hoja seleccionada: {error}")
    st.stop()
except Exception as error:
    st.error(f"Ocurrió un error al leer los datos de la hoja: {error}")
    st.stop()

if not filas_excel:
    st.warning("La hoja seleccionada está vacía.")
    st.stop()

filas_muestra = min(len(filas_excel), 50)
def etiqueta_fila(numero):
    valores = filas_excel[numero - 1]
    muestra = " | ".join(str(valor) for valor in valores[:4] if valor is not None)
    return f"Fila {numero}" + (f": {muestra[:100]}" if muestra else "")

fila_encabezados = st.selectbox(
    "Fila que contiene los encabezados",
    options=list(range(1, filas_muestra + 1)),
    index=0,
    format_func=etiqueta_fila,
)

encabezados_originales = filas_excel[fila_encabezados - 1]
cantidad_columnas = max((len(fila) for fila in filas_excel), default=0)
encabezados_originales = tuple(encabezados_originales) + (None,) * (
    cantidad_columnas - len(encabezados_originales)
)
columnas = nombres_de_columnas(encabezados_originales)
if not columnas:
    st.error("No se encontraron columnas en la hoja.")
    st.stop()

st.subheader("Configura las columnas")
col_inicio, col_fin, col_id = st.columns(3)
with col_inicio:
    inicio_columna = st.selectbox(
        "Hora/fecha de inicio",
        columnas,
        index=columnas.index(sugerir_columna(columnas, ("inicio", "start", "comienzo"))),
    )
with col_fin:
    fin_predeterminada = sugerir_columna(columnas, ("fin", "final", "end", "termino"), inicio_columna)
    fin_columna = st.selectbox(
        "Hora/fecha de fin",
        columnas,
        index=columnas.index(fin_predeterminada),
    )
with col_id:
    opciones_id = ["(Sin identificador)"] + columnas
    id_columna = st.selectbox("Identificador (opcional)", opciones_id, index=0)

if inicio_columna == fin_columna:
    st.error("Selecciona columnas diferentes para inicio y fin.")
    st.stop()

indice_inicio = columnas.index(inicio_columna)
indice_fin = columnas.index(fin_columna)
indice_id = columnas.index(id_columna) if id_columna != "(Sin identificador)" else None
filas_datos = []
for numero_fila, fila in enumerate(filas_excel[fila_encabezados:], start=fila_encabezados + 1):
    fila = tuple(fila) + (None,) * (cantidad_columnas - len(fila))
    if all(valor is None or (isinstance(valor, str) and not valor.strip()) for valor in fila):
        continue
    filas_datos.append((numero_fila, fila))

if not filas_datos:
    st.warning("No hay filas de datos debajo de los encabezados seleccionados.")
    st.stop()

resultados = []
filas_excel_salida = []
errores = []
duraciones = []
for numero_fila, fila in filas_datos:
    try:
        duracion = calcular_duracion(fila[indice_inicio], fila[indice_fin])
        segundos = duracion.total_seconds()
        minutos = segundos / 60
        duraciones.append((numero_fila, fila, duracion, segundos, minutos))
        formato = formato_hh_mm_ss(segundos)
        fila_resultado = list(fila) + [formato, minutos, segundos]
        filas_excel_salida.append(fila_resultado)
        registro = {columnas[i]: fila[i] for i in range(len(columnas))}
        registro.update({
            "Duración (HH:MM:SS)": formato,
            "Duración (minutos)": round(minutos, 2),
            "Duración (segundos)": round(segundos, 2),
        })
        resultados.append(registro)
    except (ValueError, TypeError, IndexError) as error:
        errores.append((numero_fila, str(error)))
        filas_excel_salida.append(list(fila) + [None, None, None])
        registro = {columnas[i]: fila[i] for i in range(len(columnas))}
        registro.update({
            "Duración (HH:MM:SS)": None,
            "Duración (minutos)": None,
            "Duración (segundos)": None,
        })
        resultados.append(registro)

if errores:
    st.warning(
        f"{len(errores)} fila(s) no pudieron calcularse; se mantienen en la tabla "
        "con la duración vacía."
    )
    with st.expander("Ver filas con datos vacíos o formatos inválidos"):
        for numero, motivo in errores[:20]:
            st.write(f"Fila {numero}: {motivo}")
        if len(errores) > 20:
            st.write(f"… y {len(errores) - 20} fila(s) más.")

if not duraciones:
    st.error("No se pudo calcular ninguna duración. Revisa las columnas y los valores.")
    st.dataframe(pd.DataFrame(resultados), use_container_width=True, hide_index=True)
    st.stop()

total_segundos = sum(registro[3] for registro in duraciones)
promedio_segundos = total_segundos / len(duraciones)
mas_larga = max(duraciones, key=lambda registro: registro[3])
mas_corta = min(duraciones, key=lambda registro: registro[3])

st.subheader("Resumen")
metricas = st.columns(4)
metricas[0].metric("Duración total", formato_hh_mm_ss(total_segundos), f"{total_segundos / 60:,.2f} min")
metricas[1].metric("Duración promedio", formato_hh_mm_ss(promedio_segundos), f"{promedio_segundos / 60:,.2f} min")
metricas[2].metric("Partida más larga", formato_hh_mm_ss(mas_larga[3]), f"Fila {mas_larga[0]}")
metricas[3].metric("Partida más corta", formato_hh_mm_ss(mas_corta[3]), f"Fila {mas_corta[0]}")
st.caption(f"Duraciones calculadas correctamente: {len(duraciones)} de {len(filas_datos)} filas.")

st.subheader("Duración por partida")
datos_grafico = []
for numero_fila, fila, _, segundos, _ in duraciones:
    identificador = fila[indice_id] if indice_id is not None else None
    etiqueta = str(identificador) if identificador not in (None, "") else f"Fila {numero_fila}"
    datos_grafico.append({"Partida": etiqueta, "Duración (minutos)": segundos / 60})
st.bar_chart(pd.DataFrame(datos_grafico).set_index("Partida"), y="Duración (minutos)")

st.subheader("Datos procesados")
st.dataframe(pd.DataFrame(resultados), use_container_width=True, hide_index=True)

columnas_salida = columnas + [
    "Duración (HH:MM:SS)",
    "Duración (minutos)",
    "Duración (segundos)",
]
archivo_salida = crear_excel(columnas_salida, filas_excel_salida)
st.download_button(
    "Descargar resultados en Excel",
    data=archivo_salida,
    file_name="resultados_duracion.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
