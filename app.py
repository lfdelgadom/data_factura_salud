# -*- coding: utf-8 -*-
"""
ANALÍTICA DE FACTURACIÓN EN SALUD
Facturación, radicación, glosas, recaudo y cartera

Aplicación Streamlit de un solo archivo para cargar, validar, explorar y
analizar archivos de facturación del sector salud (Excel o CSV).

- Funciona exclusivamente con el archivo que carga el usuario (en memoria).
- No usa rutas locales, bases de datos, APIs externas ni secretos.
- Es una herramienta exploratoria: no reemplaza una auditoría financiera,
  contable, médica o contractual.
"""

import csv
import hashlib
import io
import re
import unicodedata
import warnings
from contextlib import contextmanager
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from scipy import stats

# =============================================================================
# 1. CONSTANTES Y CONFIGURACIÓN GENERAL
# =============================================================================

APP_TITULO = "Analítica de Facturación en Salud"
APP_SUBTITULO = "Facturación, radicación, glosas, recaudo y cartera"

# Paleta relacionada con salud y analítica.
AZUL_OSCURO = "#0B2545"
AZUL_MEDIO = "#1F6FB2"
VERDE_AZULADO = "#13A89E"
GRIS_CLARO = "#E9EEF2"
NARANJA = "#E07A1F"
ROJO = "#C0392B"
PALETA = [AZUL_MEDIO, VERDE_AZULADO, AZUL_OSCURO, NARANJA, "#6C8EBF", "#7BC8A4", "#8E6C8A", "#B5B5B5", "#4F9DDE", "#F2B880"]
ESCALA_SECUENCIAL = [[0, "#EAF2FA"], [0.5, AZUL_MEDIO], [1, AZUL_OSCURO]]
ESCALA_DIVERGENTE = [[0, ROJO], [0.5, "#F7F7F7"], [1, AZUL_MEDIO]]

MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio",
            8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
DIAS_ES = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes", 5: "Sábado", 6: "Domingo"}
SIMBOLOS_MONEDA = {"COP": "$", "USD": "US$", "EUR": "€", "Sin símbolo": ""}

SIN_INFO = "Sin información"
NO_DISPONIBLE = "Columna no disponible"

AVISO_EXPLORATORIO = ("Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta "
                      "facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.")
AVISO_PRIVACIDAD = ("Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, historias clínicas, "
                    "diagnósticos u otra información personal o sensible. Utilice archivos anonimizados y siga las políticas "
                    "de tratamiento de datos de su organización.")
AVISO_CORRELACION = ("La correlación no implica causalidad. Algunas relaciones pueden surgir directamente de las fórmulas "
                     "contables entre las variables.")
AVISO_CALIDAD = "El puntaje es un indicador exploratorio de calidad y no reemplaza una auditoría de datos."

# Catálogo de campos esperados: etiqueta, tipo y nombres equivalentes frecuentes.
CAMPOS = {
    "ID_Factura": {"etiqueta": "Identificador de factura", "tipo": "id",
                   "sinonimos": ["factura", "numero_factura", "no_factura", "nro_factura", "num_factura", "factura_id", "id_fact", "codigo_factura"]},
    "Numero_Radicado": {"etiqueta": "Número de radicado", "tipo": "id",
                        "sinonimos": ["radicado", "no_radicado", "nro_radicado", "num_radicado", "id_radicado"]},
    "Fecha_Servicio": {"etiqueta": "Fecha de servicio", "tipo": "fecha", "sinonimos": ["fecha_atencion", "fecha_prestacion", "f_servicio"]},
    "Fecha_Emision": {"etiqueta": "Fecha de emisión", "tipo": "fecha", "sinonimos": ["fecha_factura", "fecha_expedicion", "f_emision", "fecha_facturacion"]},
    "Fecha_Radicacion": {"etiqueta": "Fecha de radicación", "tipo": "fecha", "sinonimos": ["fecha_radicado", "f_radicacion", "fecha_de_radicacion"]},
    "Fecha_Aceptacion": {"etiqueta": "Fecha de aceptación", "tipo": "fecha", "sinonimos": ["f_aceptacion"]},
    "Fecha_Pago": {"etiqueta": "Fecha de pago", "tipo": "fecha", "sinonimos": ["f_pago", "fecha_recaudo"]},
    "Periodo_Cierre": {"etiqueta": "Periodo de cierre", "tipo": "periodo", "sinonimos": ["periodo", "mes_cierre", "periodo_contable", "cierre"]},
    "Codigo_Entidad": {"etiqueta": "Código de entidad", "tipo": "categoria", "sinonimos": ["cod_entidad", "codigo_eps", "nit_entidad"]},
    "Entidad_Pagadora": {"etiqueta": "Entidad pagadora", "tipo": "categoria",
                         "sinonimos": ["entidad", "eps", "pagador", "aseguradora", "erp", "entidad_responsable_pago", "cliente"]},
    "Regimen": {"etiqueta": "Régimen", "tipo": "categoria", "sinonimos": ["tipo_regimen", "regimen_afiliacion"]},
    "Sede": {"etiqueta": "Sede", "tipo": "categoria", "sinonimos": ["sede_atencion", "sucursal", "ips", "punto_atencion"]},
    "Codigo_Servicio": {"etiqueta": "Código de servicio", "tipo": "categoria", "sinonimos": ["cod_servicio", "codigo_cups", "cups"]},
    "Servicio": {"etiqueta": "Servicio", "tipo": "categoria", "sinonimos": ["nombre_servicio", "descripcion_servicio", "procedimiento"]},
    "Linea_Servicio": {"etiqueta": "Línea de servicio", "tipo": "categoria", "sinonimos": ["linea", "linea_negocio", "unidad_funcional"]},
    "Cantidad": {"etiqueta": "Cantidad", "tipo": "numero", "sinonimos": ["cant", "unidades", "cantidad_servicios"]},
    "Canal_Radicacion": {"etiqueta": "Canal de radicación", "tipo": "categoria", "sinonimos": ["canal", "medio_radicacion"]},
    "Estado_Factura": {"etiqueta": "Estado de factura", "tipo": "categoria", "sinonimos": ["estado", "estado_cartera", "situacion"]},
    "Dias_Emision_Radicacion": {"etiqueta": "Días de radicación", "tipo": "numero",
                                "sinonimos": ["dias_radicacion", "dias_emision_a_radicacion", "tiempo_radicacion"]},
    "Soporte_Completo": {"etiqueta": "Soporte completo", "tipo": "soporte", "sinonimos": ["soporte", "soportes", "soportes_completos", "soporte_ok"]},
    "Valor_Bruto": {"etiqueta": "Valor bruto", "tipo": "dinero", "sinonimos": ["bruto", "valor_total", "subtotal"]},
    "Descuento": {"etiqueta": "Descuento", "tipo": "dinero", "sinonimos": ["descuentos", "valor_descuento"]},
    "Copago_Cuota": {"etiqueta": "Copago o cuota moderadora", "tipo": "dinero",
                     "sinonimos": ["copago", "cuota_moderadora", "copago_cuota_moderadora", "copagos", "cuota"]},
    "Valor_Neto": {"etiqueta": "Valor neto", "tipo": "dinero", "sinonimos": ["neto", "valor_facturado", "total_neto", "valor_factura"]},
    "Glosa_Inicial": {"etiqueta": "Glosa", "tipo": "dinero", "sinonimos": ["glosa", "valor_glosa", "valor_glosado", "glosas"]},
    "Valor_Pagado": {"etiqueta": "Valor pagado", "tipo": "dinero", "sinonimos": ["pagado", "valor_pago", "recaudo", "valor_recaudado", "pago"]},
    "Saldo_Cartera": {"etiqueta": "Saldo de cartera", "tipo": "dinero", "sinonimos": ["saldo", "cartera", "saldo_pendiente"]},
}
COLUMNAS_MINIMAS = ["ID_Factura", "Fecha_Radicacion o Periodo_Cierre", "Valor_Neto"]
COLUMNAS_RECOMENDADAS = ["Entidad_Pagadora", "Estado_Factura", "Valor_Pagado", "Saldo_Cartera", "Glosa_Inicial"]
COLUMNAS_DINERO = [c for c, m in CAMPOS.items() if m["tipo"] == "dinero"]
COLUMNAS_FECHA = [c for c, m in CAMPOS.items() if m["tipo"] == "fecha"]
COLUMNAS_CATEGORIA = [c for c, m in CAMPOS.items() if m["tipo"] in ("categoria", "soporte")]
VARIABLES_CORRELACION = ["Cantidad", "Valor_Bruto", "Descuento", "Copago_Cuota", "Valor_Neto",
                         "Dias_Emision_Radicacion", "Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera"]
PALABRAS_VACIAS = {"de", "del", "la", "el", "los", "las", "y", "en", "por", "a"}

PAGINAS = [
    "1. Inicio y carga de datos", "2. Calidad de datos", "3. Resumen ejecutivo", "4. Cierre mensual",
    "5. Entidades pagadoras", "6. Servicios y líneas", "7. Sedes y regímenes", "8. Estados de facturación",
    "9. Glosas", "10. Recaudo y cartera", "11. Tiempos de radicación", "12. Análisis estadístico",
    "13. Valores atípicos", "14. Insights y alertas", "15. Explorador de datos", "16. Descargas", "17. Metodología",
]

# Umbrales demostrativos para los semáforos (editables por el usuario).
UMBRALES_DEFECTO = {
    "glosa_verde": 5.0, "glosa_amarillo": 10.0,
    "cartera_verde": 30.0, "cartera_amarillo": 50.0,
    "recaudo_verde": 70.0, "recaudo_amarillo": 50.0,
    "dias_verde": 7.0, "dias_amarillo": 15.0,
    "soporte_verde": 5.0, "soporte_amarillo": 10.0,
    "crec_verde": 0.0, "crec_amarillo": -10.0,
}


# =============================================================================
# 2. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# =============================================================================

def configurar_pagina():
    """Configura título, ícono, diseño ancho y barra lateral abierta."""
    st.set_page_config(page_title=APP_TITULO, page_icon="📊", layout="wide", initial_sidebar_state="expanded")


def aplicar_estilos():
    """Aplica estilos CSS moderados para tarjetas, títulos, alertas y tablas."""
    st.markdown(
        f"""
        <style>
        .bloque-encabezado {{
            background: linear-gradient(90deg, {AZUL_OSCURO} 0%, {AZUL_MEDIO} 70%, {VERDE_AZULADO} 100%);
            color: #FFFFFF; padding: 1.1rem 1.4rem; border-radius: 12px; margin-bottom: 0.8rem;
        }}
        .bloque-encabezado h1 {{ color: #FFFFFF; font-size: 1.75rem; margin: 0; padding: 0; letter-spacing: 0.02em; }}
        .bloque-encabezado p {{ color: #EAF4FB; margin: 0.25rem 0 0 0; font-size: 1.02rem; }}
        .aviso-suave {{
            background: {GRIS_CLARO}; border-left: 5px solid {AZUL_MEDIO}; color: #1C2B3A;
            padding: 0.65rem 0.9rem; border-radius: 6px; font-size: 0.92rem; margin-bottom: 0.6rem;
        }}
        .aviso-privacidad {{
            background: #FFF5EB; border-left: 5px solid {NARANJA}; color: #3B2410;
            padding: 0.65rem 0.9rem; border-radius: 6px; font-size: 0.92rem; margin-bottom: 0.6rem;
        }}
        div[data-testid="stMetric"] {{
            background: #FFFFFF; border: 1px solid #D5DEE7; border-top: 4px solid {AZUL_MEDIO};
            border-radius: 10px; padding: 0.7rem 0.9rem; box-shadow: 0 1px 3px rgba(11,37,69,0.06);
        }}
        div[data-testid="stMetric"] label p {{ font-weight: 600; color: #34495E; }}
        .tarjeta-insight {{
            border: 1px solid #D5DEE7; border-left: 6px solid {VERDE_AZULADO}; border-radius: 8px;
            padding: 0.8rem 1rem; margin-bottom: 0.75rem; background: #FFFFFF; color: #1C2B3A;
        }}
        .tarjeta-insight h4 {{ margin: 0 0 0.35rem 0; color: {AZUL_OSCURO}; font-size: 1.02rem; }}
        .tarjeta-insight .etq {{ font-weight: 700; color: #2C3E50; }}
        .semaforo {{ border-radius: 8px; padding: 0.6rem 0.85rem; margin-bottom: 0.5rem; border: 1px solid #D5DEE7; color: #1C2B3A; }}
        .semaforo-verde {{ background: #E8F6EF; border-left: 6px solid #1E8449; }}
        .semaforo-amarillo {{ background: #FEF5E7; border-left: 6px solid {NARANJA}; }}
        .semaforo-rojo {{ background: #FDEDEC; border-left: 6px solid {ROJO}; }}
        .semaforo-gris {{ background: {GRIS_CLARO}; border-left: 6px solid #7F8C8D; }}
        h2, h3 {{ color: {AZUL_OSCURO}; }}
        .pie-pagina {{ color: #5D6D7E; font-size: 0.82rem; margin-top: 1.5rem; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def mostrar_encabezado():
    """Muestra el encabezado profesional y el aviso de uso exploratorio."""
    st.markdown(
        f'<div class="bloque-encabezado"><h1>📊 {APP_TITULO.upper()}</h1><p>{APP_SUBTITULO}</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="aviso-suave">ℹ️ {AVISO_EXPLORATORIO}</div>', unsafe_allow_html=True)


# =============================================================================
# 3. UTILIDADES DE FORMATO, CÁLCULO SEGURO Y VISUALIZACIÓN
# =============================================================================

def simbolo_moneda():
    """Devuelve el símbolo visual de moneda elegido en la barra lateral."""
    return SIMBOLOS_MONEDA.get(st.session_state.get("moneda", "COP"), "$")


def es_invalido(valor):
    """Indica si un valor no puede mostrarse (nulo, NaN o infinito)."""
    try:
        return valor is None or pd.isna(valor) or np.isinf(float(valor))
    except (TypeError, ValueError):
        return True


def formatear_numero(valor, decimales=0):
    """Formatea un número con punto de miles y coma decimal (estilo colombiano)."""
    if es_invalido(valor):
        return "—"
    texto = f"{float(valor):,.{decimales}f}"
    return texto.replace(",", "§").replace(".", ",").replace("§", ".")


def formatear_moneda(valor):
    """Formatea un valor monetario, por ejemplo: $ 1.234.567."""
    if es_invalido(valor):
        return "No calculable"
    simbolo = simbolo_moneda()
    signo = "-" if float(valor) < 0 else ""
    cuerpo = formatear_numero(abs(float(valor)), 0)
    return f"{signo}{simbolo} {cuerpo}" if simbolo else f"{signo}{cuerpo}"


def formatear_moneda_compacta(valor):
    """Moneda abreviada en millones para tarjetas (ej.: $ 5.452,4 M); valores menores a un millón se muestran completos."""
    if es_invalido(valor):
        return "No calculable"
    if abs(float(valor)) < 1e6:
        return formatear_moneda(valor)
    simbolo = simbolo_moneda()
    signo = "-" if float(valor) < 0 else ""
    cuerpo = formatear_numero(abs(float(valor)) / 1e6, 1) + " M"
    return f"{signo}{simbolo} {cuerpo}" if simbolo else f"{signo}{cuerpo}"


def formatear_porcentaje(valor, decimales=1):
    """Formatea un porcentaje con uno o dos decimales."""
    if es_invalido(valor):
        return "No calculable"
    return f"{formatear_numero(valor, decimales)} %"


def division_segura(numerador, denominador):
    """División escalar que devuelve NaN si el denominador es cero o inválido."""
    if es_invalido(numerador) or es_invalido(denominador) or float(denominador) == 0:
        return np.nan
    return float(numerador) / float(denominador)


def division_serie(numerador, denominador):
    """División vectorizada sin infinitos: los denominadores cero producen NaN."""
    denom = pd.to_numeric(denominador, errors="coerce").replace(0, np.nan)
    resultado = pd.to_numeric(numerador, errors="coerce") / denom
    return resultado.replace([np.inf, -np.inf], np.nan)


def tiene(df, *columnas):
    """Verifica que todas las columnas existan y tengan al menos un dato."""
    return df is not None and all(c in df.columns and df[c].notna().any() for c in columnas)


def requiere(df, columnas, analisis):
    """Muestra un aviso claro si faltan columnas para un análisis y devuelve False."""
    faltantes = [c for c in columnas if not tiene(df, c)]
    if faltantes:
        st.warning(f"No se encontró la columna necesaria para este análisis ({analisis}): **{', '.join(faltantes)}**. "
                   "Puede asignarla en «Configurar correspondencia de columnas» (página de inicio).")
        return False
    return True


@contextmanager
def bloque_seguro(nombre):
    """Ejecuta un bloque y, si falla, muestra un mensaje sin traceback y permite continuar."""
    try:
        yield
    except Exception:  # noqa: BLE001 - se informa al usuario sin exponer detalles técnicos
        st.warning(f"No fue posible completar «{nombre}» con los datos actuales. Los demás módulos siguen disponibles.")


def estilo_figura(fig, titulo, eje_y_moneda=False, eje_x_moneda=False, altura=430):
    """Aplica un estilo homogéneo a los gráficos Plotly (títulos, separadores, colores)."""
    simbolo = simbolo_moneda()
    fig.update_layout(
        title=dict(text=titulo, x=0, xanchor="left", font=dict(size=16, color=AZUL_OSCURO)),
        template="plotly_white", separators=",.", height=altura, colorway=PALETA,
        margin=dict(l=10, r=10, t=70, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0, title_text=""),
        font=dict(family="Segoe UI, Roboto, Arial, sans-serif", size=12, color="#1C2B3A"),
        hoverlabel=dict(bgcolor="#FFFFFF", font_size=12),
    )
    prefijo = f"{simbolo} " if simbolo else ""
    if eje_y_moneda:
        fig.update_yaxes(tickprefix=prefijo, tickformat=",.0f")
    if eje_x_moneda:
        fig.update_xaxes(tickprefix=prefijo, tickformat=",.0f")
    return fig


def mostrar_grafico(fig):
    """Muestra un gráfico interactivo adaptado al ancho, con manejo de errores."""
    try:
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
    except Exception:  # noqa: BLE001
        st.warning("No fue posible generar el gráfico con los datos actuales.")


def tabla_formateada(df, monedas=(), porcentajes=(), enteros=(), decimales=(), fechas=()):
    """Devuelve una copia del DataFrame con valores formateados para visualización."""
    salida = df.copy()
    for col in monedas:
        if col in salida.columns:
            salida[col] = salida[col].map(formatear_moneda)
    for col in porcentajes:
        if col in salida.columns:
            salida[col] = salida[col].map(formatear_porcentaje)
    for col in enteros:
        if col in salida.columns:
            salida[col] = salida[col].map(lambda v: formatear_numero(v, 0))
    for col in decimales:
        if col in salida.columns:
            salida[col] = salida[col].map(lambda v: formatear_numero(v, 1))
    for col in fechas:
        if col in salida.columns and pd.api.types.is_datetime64_any_dtype(salida[col]):
            salida[col] = salida[col].dt.strftime("%Y-%m-%d").fillna("")
    return salida


def mostrar_tabla(df, altura=None, **formatos):
    """Muestra una tabla formateada, de ancho adaptable y altura razonable."""
    if df is None or df.empty:
        st.info("No hay registros para mostrar con los filtros actuales.")
        return
    datos = tabla_formateada(df, **formatos)
    alto = altura if altura else min(38 + 35 * len(datos), 460)
    try:
        st.dataframe(datos, width="stretch", height=alto, hide_index=True)
    except TypeError:
        st.dataframe(datos, use_container_width=True, height=alto, hide_index=True)


NOMBRES_MONEDA = {"Facturación bruta", "Descuentos", "Copagos", "Facturación neta", "Glosa", "Pagado", "Saldo", "Ticket promedio",
                  "Acumulado anual", "Promedio móvil 3 periodos", "Valor", "Saldo promedio"}
NOMBRES_ENTEROS = {"Facturas", "Cantidad", "Registros", "Facturas con glosa", "Facturas con saldo", "Soportes incompletos",
                   "Devueltas", "Número de facturas"}


def columnas_formato(df):
    """Clasifica columnas de un DataFrame por tipo de formato (moneda, porcentaje, enteros, decimales, fechas)."""
    monedas = [c for c in df.columns if c in COLUMNAS_DINERO or c in NOMBRES_MONEDA]
    porcentajes = [c for c in df.columns if str(c).startswith(("%", "Porcentaje", "Participación", "Crecimiento"))]
    enteros = [c for c in df.columns if c in NOMBRES_ENTEROS]
    decimales = [c for c in df.columns if str(c).startswith("Días")]
    fechas = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
    return dict(monedas=monedas, porcentajes=porcentajes, enteros=enteros, decimales=decimales, fechas=fechas)


def mostrar_tabla_auto(df, altura=None):
    """Muestra una tabla aplicando formatos según el nombre de cada columna."""
    if df is None or df.empty:
        st.info("No hay registros para mostrar con los filtros actuales.")
        return
    mostrar_tabla(df, altura=altura, **columnas_formato(df))


def top_categorias(df, columna, valor, n):
    """Devuelve las n categorías principales según la suma de una variable."""
    if n is None:
        return df[columna].dropna().unique().tolist()
    return df.groupby(columna, observed=True)[valor].sum().sort_values(ascending=False).head(n).index.tolist()


# =============================================================================
# 4. LECTURA DE ARCHIVOS (EN MEMORIA)
# =============================================================================

@st.cache_data(show_spinner=False)
def listar_hojas(contenido):
    """Detecta las hojas disponibles en un archivo Excel."""
    with pd.ExcelFile(io.BytesIO(contenido)) as libro:
        return list(libro.sheet_names)


@st.cache_data(show_spinner=False)
def leer_excel(contenido, hoja):
    """Lee una hoja de Excel desde memoria."""
    df = pd.read_excel(io.BytesIO(contenido), sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def detectar_codificacion(contenido):
    """Intenta detectar la codificación de un archivo CSV."""
    for codificacion in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            contenido[:200000].decode(codificacion)
            return codificacion
        except UnicodeDecodeError:
            continue
    return "latin-1"


def detectar_separador(texto):
    """Identifica el separador más probable entre coma, punto y coma, tabulación y barra vertical."""
    muestra = "\n".join(texto.splitlines()[:50])
    try:
        return csv.Sniffer().sniff(muestra, delimiters=",;\t|").delimiter
    except csv.Error:
        conteos = {sep: muestra.count(sep) for sep in [";", ",", "\t", "|"]}
        return max(conteos, key=conteos.get)


@st.cache_data(show_spinner=False)
def leer_csv(contenido, separador_elegido, codificacion_elegida):
    """Lee un CSV desde memoria detectando codificación y separador. Devuelve (df, codificación, separador)."""
    codificacion = detectar_codificacion(contenido) if codificacion_elegida == "Automática" else codificacion_elegida
    texto = contenido.decode(codificacion, errors="replace")
    if not texto.strip():
        raise ValueError("vacío")
    separadores = {"Coma": ",", "Punto y coma": ";", "Tabulación": "\t"}
    separador = separadores.get(separador_elegido) or detectar_separador(texto)
    df = pd.read_csv(io.StringIO(texto), sep=separador, dtype=str, keep_default_na=True, skipinitialspace=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df, codificacion, separador


def seleccionar_hoja(hojas):
    """Selecciona automáticamente la hoja «Facturacion» si existe; si no, permite elegir otra."""
    preferida = next((h for h in hojas if normalizar_nombre_columna(h) == "facturacion"), None)
    indice = hojas.index(preferida) if preferida else 0
    hoja = st.sidebar.selectbox("Hoja a analizar", hojas, index=indice, key="hoja_seleccionada",
                                help="Se selecciona automáticamente «Facturacion» si existe.")
    if not preferida:
        st.sidebar.caption("No se encontró una hoja llamada «Facturacion». Elija la hoja que contiene los datos.")
    return hoja


def cargar_archivo():
    """Muestra el cargador y devuelve el DataFrame original y su información, o None."""
    st.sidebar.markdown("### 📁 Archivo")
    archivo = st.sidebar.file_uploader(
        "Cargue el archivo de facturación que desea analizar", type=["xlsx", "xls", "csv"],
        key=f"cargador_{st.session_state.get('clave_cargador', 0)}",
        help="Formatos aceptados: Excel (.xlsx) y CSV. El archivo se procesa en memoria y no se guarda.",
    )
    if archivo is None:
        if st.session_state.get("firma_archivo"):
            reiniciar_estado_datos()
        return None, None
    contenido = archivo.getvalue()
    firma = hashlib.md5(contenido).hexdigest() + archivo.name
    if st.session_state.get("firma_archivo") != firma:
        reiniciar_estado_datos()
        st.session_state["firma_archivo"] = firma
    info = {"nombre": archivo.name, "tamano": len(contenido), "hoja": "No aplica (CSV)", "firma": firma}
    if len(contenido) == 0:
        st.sidebar.error("El archivo está vacío.")
        return None, info
    extension = archivo.name.lower().rsplit(".", 1)[-1]
    try:
        with st.spinner("Leyendo el archivo..."):
            if extension in ("xlsx", "xls"):
                try:
                    hojas = listar_hojas(contenido)
                except Exception:  # noqa: BLE001
                    mensaje = "No fue posible leer el archivo."
                    if extension == "xls":
                        mensaje += " Los archivos .xls antiguos no son compatibles con esta aplicación: guárdelo como .xlsx o CSV."
                    st.sidebar.error(mensaje)
                    return None, info
                hoja = seleccionar_hoja(hojas)
                info["hoja"] = hoja
                df = leer_excel(contenido, hoja)
            else:
                separador = st.sidebar.selectbox("Separador del CSV", ["Detección automática", "Coma", "Punto y coma", "Tabulación"], key="separador_csv")
                codificacion = st.sidebar.selectbox("Codificación", ["Automática", "utf-8", "cp1252", "latin-1"], key="codificacion_csv")
                df, cod, sep = leer_csv(contenido, separador, codificacion)
                nombres_sep = {",": "coma", ";": "punto y coma", "\t": "tabulación", "|": "barra vertical"}
                info["hoja"] = f"CSV · separador: {nombres_sep.get(sep, sep)} · codificación: {cod}"
                if df.shape[1] == 1:
                    st.sidebar.warning("El archivo se leyó con una sola columna. Pruebe otro separador.")
    except ValueError:
        st.sidebar.error("El archivo no contiene datos.")
        return None, info
    except Exception:  # noqa: BLE001
        st.sidebar.error("No fue posible leer el archivo. Verifique el formato, el separador o la codificación.")
        return None, info
    if df is None or df.empty or df.dropna(how="all").empty:
        st.sidebar.error("La hoja seleccionada no contiene datos.")
        return None, info
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df = df.loc[:, [not str(c).startswith("Unnamed") or df[c].notna().any() for c in df.columns]]
    st.session_state["df_original"] = df
    return df, info


# =============================================================================
# 5. NORMALIZACIÓN DE NOMBRES Y MAPEO DE COLUMNAS
# =============================================================================

def quitar_tildes(texto):
    """Elimina tildes y diacríticos para facilitar comparaciones."""
    return "".join(c for c in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(c))


def normalizar_nombre_columna(nombre):
    """Normaliza un encabezado: sin tildes, minúsculas, guiones bajos y sin espacios extra."""
    texto = quitar_tildes(nombre).strip().lower()
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r"[\s\-\./]+", "_", texto)
    texto = re.sub(r"[^a-z0-9_]", "", texto)
    return re.sub(r"_+", "_", texto).strip("_")


def clave_comparacion(nombre):
    """Clave de comparación que además ignora palabras vacías (de, la, del...)."""
    partes = [p for p in normalizar_nombre_columna(nombre).split("_") if p and p not in PALABRAS_VACIAS]
    return "_".join(partes)


@st.cache_data(show_spinner=False)
def detectar_columnas(columnas):
    """Detecta equivalencias entre encabezados del archivo y campos esperados.

    Devuelve (mapeo, equivalencias, ambiguas): mapeo campo→columna, lista de equivalencias
    aplicadas y campos con más de un candidato al mismo nivel de prioridad.
    """
    indice = {}
    for campo, meta in CAMPOS.items():
        for variante in [campo] + meta["sinonimos"]:
            indice.setdefault(clave_comparacion(variante), campo)
    candidatos = {campo: [] for campo in CAMPOS}
    for col in columnas:
        if col in CAMPOS:
            candidatos[col].append((0, col))
            continue
        clave = clave_comparacion(col)
        campo = indice.get(clave)
        if campo:
            prioridad = 1 if clave == clave_comparacion(campo) else 2
            candidatos[campo].append((prioridad, col))
    mapeo, equivalencias, ambiguas, usadas = {}, [], {}, set()
    for campo, lista in candidatos.items():
        lista = sorted(lista)
        lista = [x for x in lista if x[1] not in usadas]
        if not lista:
            continue
        if len(lista) > 1 and lista[0][0] == lista[1][0]:
            ambiguas[campo] = [c for _, c in lista]
        prioridad, col = lista[0]
        mapeo[campo] = col
        usadas.add(col)
        if col != campo:
            equivalencias.append({"Columna del archivo": col, "Campo asignado": campo,
                                  "Tipo de coincidencia": "Nombre normalizado" if prioridad == 1 else "Nombre equivalente frecuente"})
    return mapeo, equivalencias, ambiguas


def clave_mapeo(campo):
    """Clave de sesión del selector de mapeo de un campo para el archivo actual."""
    return f"map_{campo}_{st.session_state.get('firma_archivo', '')[:8]}"


def depurar_mapeo(mapeo, prioritarios=()):
    """Evita que una misma columna quede asignada a varios campos.

    Las asignaciones manuales (prioritarias) prevalecen sobre las automáticas; entre iguales se conserva la primera.
    """
    orden = [c for c in mapeo if c in prioritarios] + [c for c in mapeo if c not in prioritarios]
    vistos, depurado = set(), {}
    for campo in orden:
        col = mapeo[campo]
        if col not in vistos:
            depurado[campo] = col
            vistos.add(col)
    return {campo: depurado[campo] for campo in mapeo if campo in depurado}


def mapeo_vigente(columnas, mapeo_detectado):
    """Obtiene el mapeo vigente: selectores de la sesión, mapeo guardado o detección automática."""
    if st.session_state.get("mapeo_firma") == st.session_state.get("firma_archivo") and "mapeo" in st.session_state:
        base = st.session_state["mapeo"]
    else:
        base = dict(mapeo_detectado)
    mapeo, manuales = {}, set()
    for campo in CAMPOS:
        clave = clave_mapeo(campo)
        if clave in st.session_state:
            valor = st.session_state[clave]
            if valor != mapeo_detectado.get(campo, NO_DISPONIBLE):
                manuales.add(campo)
            if valor != NO_DISPONIBLE and valor in columnas:
                mapeo[campo] = valor
        elif base.get(campo) in columnas:
            mapeo[campo] = base[campo]
    if any(clave_mapeo(campo) in st.session_state for campo in CAMPOS):
        st.session_state["mapeo_manuales"] = sorted(manuales)
    mapeo = depurar_mapeo(mapeo, set(st.session_state.get("mapeo_manuales", [])))
    st.session_state["mapeo"] = mapeo
    st.session_state["mapeo_firma"] = st.session_state.get("firma_archivo")
    return mapeo


def mapear_columnas(columnas, mapeo, mapeo_detectado):
    """Asistente de mapeo manual. Guarda la correspondencia en st.session_state durante la sesión."""
    faltan_minimas = not columnas_minimas_ok(mapeo)
    with st.expander("⚙️ Configurar correspondencia de columnas", expanded=faltan_minimas):
        st.caption("Asocie cada campo esperado con una columna del archivo. Elija «Columna no disponible» si el campo no existe. "
                   "La correspondencia solo se conserva durante esta sesión.")
        opciones = [NO_DISPONIBLE] + list(columnas)
        campos = list(CAMPOS.keys())
        col_a, col_b, col_c = st.columns(3)
        nuevo = {}
        for i, campo in enumerate(campos):
            destino = (col_a, col_b, col_c)[i % 3]
            actual = mapeo.get(campo, NO_DISPONIBLE)
            indice = opciones.index(actual) if actual in opciones else 0
            marca = " *" if campo in ("ID_Factura", "Valor_Neto", "Fecha_Radicacion", "Periodo_Cierre") else ""
            eleccion = destino.selectbox(f"{CAMPOS[campo]['etiqueta']}{marca}", opciones, index=indice,
                                         key=clave_mapeo(campo), help=f"Campo esperado: {campo}")
            if eleccion != NO_DISPONIBLE:
                nuevo[campo] = eleccion
        repetidas = pd.Series(list(nuevo.values())).value_counts()
        repetidas = repetidas[repetidas > 1].index.tolist()
        if repetidas:
            conservados = ", ".join(f"{col} → {campo}" for campo, col in mapeo.items() if col in repetidas)
            st.warning(f"Una misma columna está asignada a varios campos: {', '.join(map(str, repetidas))}. "
                       f"Para evitar nombres duplicados se conserva la asignación manual más reciente ({conservados}); "
                       "el otro campo queda como no disponible. Ajuste los selectores si desea otra combinación.")
            nuevo = dict(mapeo)
        st.caption("* Campos usados para las columnas mínimas (ID de factura, valor neto y fecha de radicación o periodo de cierre).")
        if st.button("Restablecer detección automática", key="restablecer_mapeo"):
            st.session_state["mapeo"] = dict(mapeo_detectado)
            st.session_state["mapeo_manuales"] = []
            for campo in campos:
                st.session_state.pop(clave_mapeo(campo), None)
            st.rerun()
    return nuevo


def columnas_minimas_ok(mapeo):
    """Verifica la presencia de las columnas mínimas en el mapeo."""
    return "ID_Factura" in mapeo and "Valor_Neto" in mapeo and ("Fecha_Radicacion" in mapeo or "Periodo_Cierre" in mapeo)


def faltantes_minimas(mapeo):
    """Lista las columnas mínimas que faltan."""
    faltan = []
    if "ID_Factura" not in mapeo:
        faltan.append("ID_Factura")
    if "Fecha_Radicacion" not in mapeo and "Periodo_Cierre" not in mapeo:
        faltan.append("Fecha_Radicacion o Periodo_Cierre")
    if "Valor_Neto" not in mapeo:
        faltan.append("Valor_Neto")
    return faltan


@st.cache_data(show_spinner=False)
def aplicar_mapeo(df, mapeo_items):
    """Renombra las columnas según el mapeo, sin generar nombres duplicados."""
    mapeo = dict(mapeo_items)
    renombrar = {col: campo for campo, col in mapeo.items() if col in df.columns}
    datos = df[list(renombrar.keys())].rename(columns=renombrar)
    otras = [c for c in df.columns if c not in renombrar]
    extras = df[otras].copy()
    extras.columns = [f"Original_{c}" if (c in CAMPOS or c in datos.columns) else c for c in otras]
    return pd.concat([datos, extras], axis=1)


# =============================================================================
# 6. CONVERSIÓN DE TIPOS (FECHAS, VALORES MONETARIOS, PERIODOS, CATEGORÍAS)
# =============================================================================

def convertir_valores_monetarios(serie, robusta=True):
    """Convierte valores monetarios de forma robusta.

    Reconoce números, símbolos de moneda, separadores de miles, comas o puntos decimales
    y espacios. Devuelve (serie_float, no_convertidos, ambiguos).
    """
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").astype(float), 0, 0
    original = serie.copy()
    texto = original.where(original.notna(), "").astype(str).str.strip()
    vacio = texto.eq("") | texto.str.lower().isin(["nan", "none", "null", "na", "n/a", "-", "nat"])
    if not robusta:
        valores = pd.to_numeric(texto.where(~vacio), errors="coerce")
        return valores.astype(float), int((~vacio & valores.isna()).sum()), 0
    es_numero = original.map(lambda x: isinstance(x, (int, float, np.integer, np.floating)) and not isinstance(x, bool))
    negativo_parentesis = texto.str.match(r"^\(.*\)$", na=False)
    limpio = texto.str.replace(r"(?i)(us\$|col\$|cop|usd|eur|\$|€)", "", regex=True)
    limpio = limpio.str.replace(r"[\s '’()]", "", regex=True)
    tiene_punto = limpio.str.contains(".", regex=False)
    tiene_coma = limpio.str.contains(",", regex=False)
    ambos = tiene_punto & tiene_coma
    coma_decimal = ambos & (limpio.str.rfind(",") > limpio.str.rfind("."))
    solo_coma = tiene_coma & ~tiene_punto
    solo_punto = tiene_punto & ~tiene_coma
    # Patrones de un solo grupo (1.234 o 1,234) son ambiguos; se resuelven con la evidencia de toda la columna.
    ambiguo_punto = solo_punto & limpio.str.match(r"^-?\d{1,3}\.\d{3}$", na=False) & ~es_numero
    ambiguo_coma = solo_coma & limpio.str.match(r"^-?\d{1,3},\d{3}$", na=False) & ~es_numero
    miles_punto_claro = solo_punto & limpio.str.match(r"^-?\d{1,3}(\.\d{3}){2,}$", na=False)
    miles_coma_claro = solo_coma & limpio.str.match(r"^-?\d{1,3}(,\d{3}){2,}$", na=False)
    decimal_punto_claro = solo_punto & limpio.str.match(r"^-?\d*\.(\d{1,2}|\d{4,})$", na=False)
    decimal_coma_claro = solo_coma & limpio.str.match(r"^-?\d*,(\d{1,2}|\d{4,})$", na=False)
    evidencia_miles_punto = bool(miles_punto_claro.any() or coma_decimal.any())
    evidencia_miles_coma = bool(miles_coma_claro.any() or (ambos & ~coma_decimal).any())
    punto_es_decimal = not evidencia_miles_punto and bool(decimal_punto_claro.any())
    coma_es_decimal = not evidencia_miles_coma and bool(decimal_coma_claro.any())
    miles_punto = solo_punto & limpio.str.match(r"^-?\d{1,3}(\.\d{3})+$", na=False) & ~(ambiguo_punto & punto_es_decimal)
    miles_coma = solo_coma & limpio.str.match(r"^-?\d{1,3}(,\d{3})+$", na=False) & ~(ambiguo_coma & coma_es_decimal)
    ambiguo = (ambiguo_punto & (not evidencia_miles_punto) & (not punto_es_decimal)) | \
              (ambiguo_coma & (not evidencia_miles_coma) & (not coma_es_decimal))
    sin_punto = limpio.str.replace(".", "", regex=False)
    sin_coma = limpio.str.replace(",", "", regex=False)
    normalizado = np.select(
        [coma_decimal, ambos & ~coma_decimal, miles_coma, solo_coma, miles_punto],
        [sin_punto.str.replace(",", ".", regex=False), sin_coma, sin_coma, limpio.str.replace(",", ".", regex=False), sin_punto],
        default=limpio,
    )
    valores = pd.to_numeric(pd.Series(normalizado, index=serie.index).where(~vacio), errors="coerce")
    if es_numero.any():
        valores[es_numero] = pd.to_numeric(original[es_numero], errors="coerce")
    valores = valores.where(~negativo_parentesis, -valores.abs())
    no_convertidos = int((~vacio & valores.isna()).sum())
    return valores.astype(float), no_convertidos, int(ambiguo.sum())


def convertir_fecha_serie(serie, robusta=True):
    """Convierte una serie a fecha con errors='coerce'. Devuelve (serie_fecha, inválidas)."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie, 0
    if pd.api.types.is_numeric_dtype(serie):
        plausible = serie.between(20000, 80000)
        fechas = pd.to_datetime(serie.where(plausible), unit="D", origin="1899-12-30", errors="coerce")
        return fechas, int((serie.notna() & fechas.isna()).sum())
    texto = serie.where(serie.notna(), "").astype(str).str.strip()
    vacio = texto.eq("") | texto.str.lower().isin(["nan", "none", "nat", "null", "-"])
    fechas = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        iso = texto.str.match(r"^\d{4}-\d{1,2}-\d{1,2}", na=False) & ~vacio
        if iso.any():
            fechas[iso] = pd.to_datetime(texto[iso].str.slice(0, 10), format="%Y-%m-%d", errors="coerce")
        otros = ~iso & ~vacio
        if otros.any():
            if robusta:
                fechas[otros] = pd.to_datetime(texto[otros], errors="coerce", dayfirst=True, format="mixed")
            else:
                fechas[otros] = pd.to_datetime(texto[otros], errors="coerce")
    return fechas, int((~vacio & fechas.isna()).sum())


def convertir_fechas(df, robusta=True):
    """Convierte todas las columnas de fecha presentes e informa los resultados."""
    informe, mascaras = [], {}
    for col in COLUMNAS_FECHA:
        if col not in df.columns:
            continue
        nulos_antes = int(df[col].isna().sum())
        try:
            convertida, invalidas = convertir_fecha_serie(df[col], robusta)
        except Exception:  # noqa: BLE001
            convertida, invalidas = pd.to_datetime(df[col], errors="coerce"), int(df[col].notna().sum())
        mascaras[f"Fecha inválida: {col}"] = df[col].notna() & convertida.isna() & df[col].astype(str).str.strip().ne("")
        df[col] = convertida
        informe.append({"Columna": col, "Conversión": "Fecha", "Valores no convertidos": invalidas,
                        "Nulos antes": nulos_antes, "Nulos después": int(convertida.isna().sum()), "Advertencia": ""})
    return df, informe, mascaras


def normalizar_periodo(serie):
    """Valida y normaliza el periodo de cierre al formato AAAA-MM. Devuelve (serie, inválidos)."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        resultado = serie.dt.strftime("%Y-%m")
        return resultado, int((serie.notna() & resultado.isna()).sum())
    texto = serie.where(serie.notna(), "").astype(str).str.strip()
    vacio = texto.eq("") | texto.str.lower().isin(["nan", "none", "nat"])
    p1 = texto.str.extract(r"^(\d{4})[-/\.]?(\d{1,2})(?:\D|$)")
    p2 = texto.str.extract(r"^(\d{1,2})[-/\.](\d{4})$")
    anio = pd.to_numeric(p1[0], errors="coerce").fillna(pd.to_numeric(p2[1], errors="coerce"))
    mes = pd.to_numeric(p1[1], errors="coerce").fillna(pd.to_numeric(p2[0], errors="coerce"))
    valido = anio.between(1990, 2100) & mes.between(1, 12)
    resultado = (anio.astype("Int64").astype(str) + "-" + mes.astype("Int64").astype(str).str.zfill(2)).where(valido)
    return resultado, int((~vacio & ~valido).sum())


def normalizar_soporte(serie):
    """Normaliza la columna de soporte completo a «Sí» / «No» / «Sin información»."""
    texto = serie.where(serie.notna(), "").astype(str).map(quitar_tildes).str.strip().str.lower().str.replace(r"\.0$", "", regex=True)
    si = texto.isin(["si", "s", "yes", "y", "true", "1", "completo", "verdadero", "x", "ok"])
    no = texto.isin(["no", "n", "false", "0", "incompleto", "falso"])
    return pd.Series(np.select([si, no], ["Sí", "No"], default=SIN_INFO), index=serie.index)


@st.cache_data(show_spinner=False)
def preparar_datos(df_mapeado, robusta_fechas=True, robusta_moneda=True):
    """Convierte tipos sobre una copia e informa cada conversión.

    Devuelve (df_tipado, informe_conversiones, máscaras_de_problemas, advertencias).
    """
    df = df_mapeado.copy()
    informe, advertencias = [], []
    df, informe_fechas, mascaras = convertir_fechas(df, robusta_fechas)
    informe.extend(informe_fechas)
    for col in COLUMNAS_DINERO + ["Cantidad", "Dias_Emision_Radicacion"]:
        if col not in df.columns:
            continue
        nulos_antes = int(df[col].isna().sum())
        try:
            valores, no_conv, ambiguos = convertir_valores_monetarios(df[col], robusta_moneda)
        except Exception:  # noqa: BLE001
            valores = pd.to_numeric(df[col], errors="coerce")
            no_conv, ambiguos = int((df[col].notna() & valores.isna()).sum()), 0
        texto_adv = ""
        if ambiguos:
            texto_adv = (f"{ambiguos} valores con formato ambiguo (por ejemplo 1.234 o 1,234) se interpretaron "
                         "como separador de miles.")
            advertencias.append(f"{col}: {texto_adv}")
        mascaras[f"Valor no numérico: {col}"] = df[col].notna() & valores.isna() & df[col].astype(str).str.strip().ne("")
        df[col] = valores
        informe.append({"Columna": col, "Conversión": "Monetaria" if col in COLUMNAS_DINERO else "Numérica",
                        "Valores no convertidos": no_conv, "Nulos antes": nulos_antes,
                        "Nulos después": int(valores.isna().sum()), "Advertencia": texto_adv})
    if "Periodo_Cierre" in df.columns:
        nulos_antes = int(df["Periodo_Cierre"].isna().sum())
        periodo, invalidos = normalizar_periodo(df["Periodo_Cierre"])
        mascaras["Periodo inválido: Periodo_Cierre"] = df["Periodo_Cierre"].notna() & periodo.isna()
        df["Periodo_Cierre"] = periodo
        informe.append({"Columna": "Periodo_Cierre", "Conversión": "Periodo AAAA-MM", "Valores no convertidos": invalidos,
                        "Nulos antes": nulos_antes, "Nulos después": int(periodo.isna().sum()),
                        "Advertencia": "Formato esperado AAAA-MM" if invalidos else ""})
    if "Soporte_Completo" in df.columns:
        df["Soporte_Completo"] = normalizar_soporte(df["Soporte_Completo"])
    for col in ["ID_Factura", "Numero_Radicado"] + [c for c in COLUMNAS_CATEGORIA if c != "Soporte_Completo"]:
        if col in df.columns:
            texto = df[col].astype("string").str.strip()
            texto = texto.str.replace(r"\.0$", "", regex=True) if col in ("ID_Factura", "Numero_Radicado", "Codigo_Entidad", "Codigo_Servicio") else texto
            df[col] = texto.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA}).astype(object).where(lambda s: s.notna(), np.nan)
    mascaras_df = pd.DataFrame(mascaras, index=df.index) if mascaras else pd.DataFrame(index=df.index)
    return df, informe, mascaras_df, advertencias


# =============================================================================
# 7. VALIDACIÓN Y PUNTAJE DE CALIDAD
# =============================================================================

def marcar_inconsistencias(df, tolerancia):
    """Marca, fila por fila de forma vectorizada, las inconsistencias lógicas y financieras."""
    banderas = pd.DataFrame(index=df.index)
    if tiene(df, "Fecha_Radicacion", "Fecha_Emision"):
        banderas["Radicación anterior a emisión"] = df["Fecha_Radicacion"] < df["Fecha_Emision"]
    if tiene(df, "Fecha_Pago", "Fecha_Radicacion"):
        banderas["Pago anterior a radicación"] = df["Fecha_Pago"] < df["Fecha_Radicacion"]
    if tiene(df, "Valor_Neto"):
        neto = df["Valor_Neto"]
        if tiene(df, "Saldo_Cartera"):
            banderas["Saldo mayor que valor neto"] = df["Saldo_Cartera"] > neto + tolerancia
        if tiene(df, "Valor_Pagado"):
            banderas["Pago mayor que valor neto"] = df["Valor_Pagado"] > neto + tolerancia
        if tiene(df, "Glosa_Inicial"):
            banderas["Glosa mayor que valor neto"] = df["Glosa_Inicial"] > neto + tolerancia
        if tiene(df, "Valor_Bruto"):
            esperado = df["Valor_Bruto"] - df.get("Descuento", 0).fillna(0) if "Descuento" in df.columns else df["Valor_Bruto"]
            if "Copago_Cuota" in df.columns:
                esperado = esperado - df["Copago_Cuota"].fillna(0)
            banderas["Identidad neto ≠ bruto − descuento − copago"] = (neto - esperado).abs() > tolerancia
        if tiene(df, "Saldo_Cartera") and ("Glosa_Inicial" in df.columns or "Valor_Pagado" in df.columns):
            esperado = neto.copy()
            if "Glosa_Inicial" in df.columns:
                esperado = esperado - df["Glosa_Inicial"].fillna(0)
            if "Valor_Pagado" in df.columns:
                esperado = esperado - df["Valor_Pagado"].fillna(0)
            banderas["Identidad saldo ≠ neto − glosa − pagado"] = (df["Saldo_Cartera"] - esperado).abs() > tolerancia
    negativos = [c for c in COLUMNAS_DINERO if c in df.columns]
    if negativos:
        banderas["Valores monetarios negativos"] = (df[negativos] < 0).any(axis=1)
    return banderas.fillna(False).astype(bool)


@st.cache_data(show_spinner=False)
def validar_datos(df, mascaras, tolerancia, columnas_archivo):
    """Calcula el diagnóstico de calidad de datos y devuelve un diccionario de resultados."""
    n = len(df)
    resultado = {"filas": n, "columnas_archivo": len(columnas_archivo)}
    presentes = [c for c in CAMPOS if c in df.columns]
    resultado["encontradas"] = presentes
    resultado["ausentes"] = [c for c in CAMPOS if c not in df.columns]
    resultado["filas_duplicadas"] = int(df.duplicated().sum())
    resultado["id_duplicados"] = int(df["ID_Factura"].dropna().duplicated().sum()) if "ID_Factura" in df.columns else None
    resultado["radicados_duplicados"] = int(df["Numero_Radicado"].dropna().duplicated().sum()) if "Numero_Radicado" in df.columns else None
    nulos = df[presentes].isna().sum() if presentes else pd.Series(dtype=int)
    resultado["nulos"] = pd.DataFrame({"Columna": nulos.index, "Nulos": nulos.values,
                                       "% nulos": (nulos.values / n * 100) if n else 0})
    resultado["fechas_invalidas"] = {c.split(": ")[1]: int(mascaras[c].sum()) for c in mascaras.columns if c.startswith("Fecha inválida")}
    resultado["no_numericos"] = {c.split(": ")[1]: int(mascaras[c].sum()) for c in mascaras.columns if c.startswith("Valor no numérico")}
    dinero = [c for c in COLUMNAS_DINERO if c in df.columns]
    resultado["negativos"] = {c: int((df[c] < 0).sum()) for c in dinero}
    resultado["neto_cero"] = int((df["Valor_Neto"] == 0).sum()) if "Valor_Neto" in df.columns else None
    resultado["sin_entidad"] = int(df["Entidad_Pagadora"].isna().sum()) if "Entidad_Pagadora" in df.columns else None
    if "Periodo_Cierre" in df.columns:
        resultado["sin_periodo"] = int(df["Periodo_Cierre"].isna().sum())
    elif "Fecha_Radicacion" in df.columns:
        resultado["sin_periodo"] = int(df["Fecha_Radicacion"].isna().sum())
    else:
        resultado["sin_periodo"] = None
    banderas = marcar_inconsistencias(df, tolerancia)
    resultado["banderas"] = banderas
    resultado["inconsistencias"] = {c: int(banderas[c].sum()) for c in banderas.columns}
    resultado["filas_inconsistentes"] = int(banderas.any(axis=1).sum()) if not banderas.empty else 0
    resultado["puntaje"] = calcular_puntaje_calidad(df, mascaras, banderas, resultado)
    return resultado


def calcular_puntaje_calidad(df, mascaras, banderas, resultado):
    """Calcula un puntaje orientativo 0–100 con completitud, unicidad, validez y consistencia."""
    n = len(df)
    if n == 0:
        return {"Completitud": 0.0, "Unicidad": 0.0, "Validez": 0.0, "Consistencia": 0.0, "Total": 0.0}
    evaluables = [c for c in CAMPOS if c in df.columns and c not in ("Fecha_Pago", "Fecha_Aceptacion")]
    completitud = 1 - (df[evaluables].isna().to_numpy().mean() if evaluables else 0)
    duplicados = resultado["filas_duplicadas"] + (resultado["id_duplicados"] or 0)
    unicidad = max(0.0, 1 - duplicados / n)
    problemas_validez = mascaras.any(axis=1) if not mascaras.empty else pd.Series(False, index=df.index)
    if "Valor_Neto" in df.columns:
        problemas_validez = problemas_validez | df["Valor_Neto"].isna()
    validez = 1 - problemas_validez.mean()
    consistencia = 1 - (banderas.any(axis=1).mean() if not banderas.empty else 0)
    partes = {"Completitud": completitud * 100, "Unicidad": unicidad * 100, "Validez": validez * 100, "Consistencia": consistencia * 100}
    partes["Total"] = float(np.mean(list(partes.values())))
    return {k: round(float(v), 1) for k, v in partes.items()}


# =============================================================================
# 8. LIMPIEZA CONTROLADA
# =============================================================================

@st.cache_data(show_spinner=False)
def limpiar_datos(df, mascaras, opciones_items):
    """Aplica la limpieza sobre una copia y devuelve (df_limpio, resumen_de_exclusiones)."""
    opciones = dict(opciones_items)
    datos = df.copy()
    exclusiones = []
    if not opciones.get("aplicar", True):
        return datos, exclusiones

    def excluir(mascara, razon):
        nonlocal datos
        mascara = mascara.reindex(datos.index).fillna(False).astype(bool)
        cantidad = int(mascara.sum())
        if cantidad:
            datos = datos.loc[~mascara]
        exclusiones.append({"Razón de exclusión": razon, "Registros excluidos": cantidad})

    if opciones.get("vacias"):
        excluir(datos.isna().all(axis=1), "Filas completamente vacías")
    if opciones.get("duplicados"):
        excluir(datos.duplicated(), "Duplicados exactos")
    if opciones.get("neto_invalido") and "Valor_Neto" in datos.columns:
        excluir(datos["Valor_Neto"].isna(), "Valor_Neto inválido o vacío")
    if opciones.get("fechas_invalidas"):
        invalidas = pd.Series(False, index=datos.index)
        columnas_fecha = [c for c in mascaras.columns if c.startswith("Fecha inválida")]
        if columnas_fecha:
            invalidas = mascaras.loc[datos.index, columnas_fecha].any(axis=1)
        if "Fecha_Radicacion" in datos.columns and "Periodo_Cierre" not in datos.columns:
            invalidas = invalidas | datos["Fecha_Radicacion"].isna()
        excluir(invalidas, "Fechas inválidas")
    if opciones.get("rellenar"):
        for col in COLUMNAS_CATEGORIA:
            if col in datos.columns:
                datos[col] = datos[col].fillna(SIN_INFO)
    return datos, exclusiones


# =============================================================================
# 9. VARIABLES DERIVADAS
# =============================================================================

@st.cache_data(show_spinner=False)
def crear_variables_derivadas(df, limites_tiempo):
    """Crea periodos, partes de fecha, porcentajes, banderas y rangos de tiempo (sin infinitos)."""
    datos = df.copy()
    l1, l2, l3 = limites_tiempo
    if "Fecha_Radicacion" in datos.columns:
        desde_fecha = datos["Fecha_Radicacion"].dt.strftime("%Y-%m")
        if "Periodo_Cierre" in datos.columns:
            datos["Periodo_Cierre"] = datos["Periodo_Cierre"].fillna(desde_fecha)
        else:
            datos["Periodo_Cierre"] = desde_fecha
    base = datos["Fecha_Radicacion"] if "Fecha_Radicacion" in datos.columns else pd.Series(pd.NaT, index=datos.index)
    if "Periodo_Cierre" in datos.columns:
        desde_periodo = pd.to_datetime(datos["Periodo_Cierre"] + "-01", format="%Y-%m-%d", errors="coerce")
        anio = desde_periodo.dt.year
        mes = desde_periodo.dt.month
    else:
        anio, mes = base.dt.year, base.dt.month
    datos["Año_Radicacion"] = anio.astype("Int64")
    datos["Numero_Mes"] = mes.astype("Int64")
    datos["Mes_Radicacion"] = pd.Categorical(mes.map(MESES_ES), categories=list(MESES_ES.values()), ordered=True)
    datos["Trimestre_Radicacion"] = ("T" + ((mes - 1) // 3 + 1).astype("Int64").astype(str)).where(mes.notna())
    datos["Año_Mes"] = datos["Periodo_Cierre"] if "Periodo_Cierre" in datos.columns else np.nan
    if "Fecha_Radicacion" in datos.columns:
        datos["Día_Semana"] = pd.Categorical(base.dt.dayofweek.map(DIAS_ES), categories=list(DIAS_ES.values()), ordered=True)
    if "Fecha_Emision" in datos.columns and "Fecha_Radicacion" in datos.columns:
        calculado = (datos["Fecha_Radicacion"] - datos["Fecha_Emision"]).dt.days
        if "Dias_Emision_Radicacion" in datos.columns:
            datos["Dias_Emision_Radicacion"] = datos["Dias_Emision_Radicacion"].fillna(calculado)
        else:
            datos["Dias_Emision_Radicacion"] = calculado.astype(float)
    if "Fecha_Pago" in datos.columns and "Fecha_Radicacion" in datos.columns:
        datos["Dias_Radicacion_Pago"] = (datos["Fecha_Pago"] - datos["Fecha_Radicacion"]).dt.days.astype(float)
    if "Valor_Bruto" in datos.columns:
        if "Descuento" in datos.columns:
            datos["Porcentaje_Descuento"] = division_serie(datos["Descuento"], datos["Valor_Bruto"]) * 100
        if "Copago_Cuota" in datos.columns:
            datos["Porcentaje_Copago"] = division_serie(datos["Copago_Cuota"], datos["Valor_Bruto"]) * 100
    if "Valor_Neto" in datos.columns:
        if "Glosa_Inicial" in datos.columns:
            datos["Porcentaje_Glosa"] = division_serie(datos["Glosa_Inicial"], datos["Valor_Neto"]) * 100
        if "Valor_Pagado" in datos.columns:
            datos["Porcentaje_Recaudo"] = division_serie(datos["Valor_Pagado"], datos["Valor_Neto"]) * 100
        if "Saldo_Cartera" in datos.columns:
            datos["Porcentaje_Cartera"] = division_serie(datos["Saldo_Cartera"], datos["Valor_Neto"]) * 100
    if "Glosa_Inicial" in datos.columns:
        datos["Tiene_Glosa"] = datos["Glosa_Inicial"].fillna(0) > 0
    if "Saldo_Cartera" in datos.columns:
        datos["Tiene_Saldo"] = datos["Saldo_Cartera"].fillna(0) > 0
    if "Estado_Factura" in datos.columns:
        estado_norm = datos["Estado_Factura"].astype(str).map(normalizar_nombre_columna)
        datos["Es_Pagada"] = estado_norm.eq("pagada")
    if "Soporte_Completo" in datos.columns:
        datos["Soporte_Incompleto"] = datos["Soporte_Completo"].eq("No")
    if "Dias_Emision_Radicacion" in datos.columns:
        dias = datos["Dias_Emision_Radicacion"]
        etiquetas = [f"0 a {l1} días", f"{l1 + 1} a {l2} días", f"{l2 + 1} a {l3} días", f"Más de {l3} días"]
        rango = pd.cut(dias, bins=[-0.5, l1 + 0.5, l2 + 0.5, l3 + 0.5, np.inf], labels=etiquetas)
        rango = rango.cat.add_categories(["Negativo (revisar)"])
        rango[dias < 0] = "Negativo (revisar)"
        datos["Rango_Tiempo_Radicacion"] = rango
    return datos


# =============================================================================
# 10. FILTROS GLOBALES
# =============================================================================

FILTROS_CATEGORICOS = [
    ("Año_Radicacion", "Año"), ("Mes_Radicacion", "Mes"), ("Entidad_Pagadora", "Entidad pagadora"),
    ("Regimen", "Régimen"), ("Sede", "Sede"), ("Servicio", "Servicio"), ("Linea_Servicio", "Línea de servicio"),
    ("Estado_Factura", "Estado de factura"), ("Canal_Radicacion", "Canal de radicación"), ("Soporte_Completo", "Soporte completo"),
]


def opciones_filtro(df, columna):
    """Opciones ordenadas para un filtro (meses en orden cronológico)."""
    if columna == "Mes_Radicacion":
        presentes = set(df[columna].dropna().astype(str))
        return [m for m in MESES_ES.values() if m in presentes]
    valores = df[columna].dropna().unique().tolist()
    try:
        return sorted(valores)
    except TypeError:
        return sorted(map(str, valores))


def restablecer_filtros(df, seleccionar_todo):
    """Callback de los botones «Seleccionar todos» y «Limpiar filtros»."""
    for columna, _ in FILTROS_CATEGORICOS:
        if columna in df.columns:
            st.session_state[f"f_{columna}"] = opciones_filtro(df, columna) if seleccionar_todo else []
    for clave in ("f_glosa", "f_saldo"):
        st.session_state[clave] = "Todas"
    for clave in ("f_fechas", "f_valor"):
        st.session_state.pop(clave, None)


def sanear_estado_filtros(df):
    """Elimina valores guardados de filtros que ya no son válidos para los datos actuales."""
    for columna, _ in FILTROS_CATEGORICOS:
        clave = f"f_{columna}"
        if clave in st.session_state and columna in df.columns:
            validas = set(opciones_filtro(df, columna))
            st.session_state[clave] = [v for v in st.session_state[clave] if v in validas]
    if "f_valor" in st.session_state and tiene(df, "Valor_Neto"):
        vmin, vmax = float(df["Valor_Neto"].min()), float(df["Valor_Neto"].max())
        actual = st.session_state["f_valor"]
        if not (isinstance(actual, (tuple, list)) and len(actual) == 2 and vmin <= actual[0] <= actual[1] <= vmax):
            del st.session_state["f_valor"]
    if "f_fechas" in st.session_state and tiene(df, "Fecha_Radicacion"):
        minimo, maximo = df["Fecha_Radicacion"].min().date(), df["Fecha_Radicacion"].max().date()
        actual = st.session_state["f_fechas"]
        if not all(minimo <= d <= maximo for d in (actual if isinstance(actual, (tuple, list)) else [actual])):
            del st.session_state["f_fechas"]


def construir_filtros(df):
    """Dibuja los filtros globales en la barra lateral y devuelve su configuración."""
    sanear_estado_filtros(df)
    filtros = {"categorias": {}, "fechas": None, "valor": None, "glosa": "Todas", "saldo": "Todas"}
    with st.sidebar.expander("🔎 Filtros globales", expanded=True):
        c1, c2 = st.columns(2)
        c1.button("Seleccionar todos", on_click=restablecer_filtros, args=(df, True), key="btn_sel_todos", width="stretch")
        c2.button("Limpiar filtros", on_click=restablecer_filtros, args=(df, False), key="btn_limpiar", width="stretch")
        if tiene(df, "Fecha_Radicacion"):
            minimo, maximo = df["Fecha_Radicacion"].min().date(), df["Fecha_Radicacion"].max().date()
            if minimo < maximo:
                rango = st.date_input("Rango de fechas de radicación", value=(minimo, maximo), min_value=minimo,
                                      max_value=maximo, key="f_fechas", format="YYYY-MM-DD")
                if isinstance(rango, (tuple, list)) and len(rango) == 2 and (rango[0] != minimo or rango[1] != maximo):
                    filtros["fechas"] = (pd.Timestamp(rango[0]), pd.Timestamp(rango[1]))
        for columna, etiqueta in FILTROS_CATEGORICOS:
            if columna in df.columns and df[columna].notna().any():
                seleccion = st.multiselect(etiqueta, opciones_filtro(df, columna), key=f"f_{columna}",
                                           placeholder="Todos", help="Vacío = todos los valores.")
                if seleccion:
                    filtros["categorias"][columna] = seleccion
        if "Tiene_Glosa" in df.columns:
            filtros["glosa"] = st.radio("Facturas con glosa", ["Todas", "Con glosa", "Sin glosa"], horizontal=True, key="f_glosa")
        if "Tiene_Saldo" in df.columns:
            filtros["saldo"] = st.radio("Facturas con saldo", ["Todas", "Con saldo", "Sin saldo"], horizontal=True, key="f_saldo")
        if tiene(df, "Valor_Neto"):
            vmin, vmax = float(df["Valor_Neto"].min()), float(df["Valor_Neto"].max())
            if vmin < vmax:
                rango_valor = st.slider("Rango de valor neto", min_value=vmin, max_value=vmax, value=(vmin, vmax), key="f_valor")
                if rango_valor[0] > vmin or rango_valor[1] < vmax:
                    filtros["valor"] = rango_valor
    return filtros


def aplicar_filtros(df, filtros):
    """Aplica los filtros globales de forma vectorizada y devuelve el subconjunto."""
    mascara = pd.Series(True, index=df.index)
    if filtros.get("fechas") and "Fecha_Radicacion" in df.columns:
        inicio, fin = filtros["fechas"]
        mascara &= df["Fecha_Radicacion"].between(inicio, fin + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))
    for columna, seleccion in filtros.get("categorias", {}).items():
        if columna == "Mes_Radicacion":
            mascara &= df[columna].astype(str).isin(seleccion)
        else:
            mascara &= df[columna].isin(seleccion)
    if filtros.get("glosa") == "Con glosa":
        mascara &= df["Tiene_Glosa"]
    elif filtros.get("glosa") == "Sin glosa":
        mascara &= ~df["Tiene_Glosa"]
    if filtros.get("saldo") == "Con saldo":
        mascara &= df["Tiene_Saldo"]
    elif filtros.get("saldo") == "Sin saldo":
        mascara &= ~df["Tiene_Saldo"]
    if filtros.get("valor") and "Valor_Neto" in df.columns:
        mascara &= df["Valor_Neto"].between(*filtros["valor"])
    return df.loc[mascara]


# =============================================================================
# 11. INDICADORES (KPI) Y AGREGACIONES
# =============================================================================

def contar_facturas(df):
    """Cuenta facturas únicas (sin contar varias veces una misma factura)."""
    if "ID_Factura" in df.columns and df["ID_Factura"].notna().any():
        return int(df["ID_Factura"].nunique())
    return int(len(df))


def suma(df, columna):
    """Suma segura de una columna (NaN si no existe)."""
    return float(df[columna].sum()) if columna in df.columns and df[columna].notna().any() else np.nan


def calcular_kpis(df):
    """Calcula los indicadores principales sobre el conjunto recibido."""
    facturas = contar_facturas(df)
    neto = suma(df, "Valor_Neto")
    kpis = {
        "bruto": suma(df, "Valor_Bruto"), "neto": neto, "facturas": facturas,
        "cantidad": suma(df, "Cantidad"), "pagado": suma(df, "Valor_Pagado"), "saldo": suma(df, "Saldo_Cartera"),
        "glosa": suma(df, "Glosa_Inicial"), "descuento": suma(df, "Descuento"), "copago": suma(df, "Copago_Cuota"),
    }
    kpis["pct_recaudo"] = division_segura(kpis["pagado"], neto) * 100
    kpis["pct_glosa"] = division_segura(kpis["glosa"], neto) * 100
    kpis["pct_cartera"] = division_segura(kpis["saldo"], neto) * 100
    kpis["ticket"] = division_segura(neto, facturas)
    kpis["dias"] = float(df["Dias_Emision_Radicacion"].mean()) if tiene(df, "Dias_Emision_Radicacion") else np.nan
    kpis["pct_soporte_incompleto"] = float(df["Soporte_Incompleto"].mean() * 100) if "Soporte_Incompleto" in df.columns and len(df) else np.nan
    return kpis


def agregar_por(df, grupo):
    """Agregación financiera estándar por una o varias columnas de agrupación."""
    grupos = [grupo] if isinstance(grupo, str) else list(grupo)
    agregados = {"Registros": ("Valor_Neto", "size")}
    if "ID_Factura" in df.columns:
        agregados["Facturas"] = ("ID_Factura", "nunique")
    columnas = {"Cantidad": "Cantidad", "Valor_Bruto": "Facturación bruta", "Descuento": "Descuentos", "Copago_Cuota": "Copagos",
                "Valor_Neto": "Facturación neta", "Glosa_Inicial": "Glosa", "Valor_Pagado": "Pagado", "Saldo_Cartera": "Saldo"}
    for origen, destino in columnas.items():
        if origen in df.columns:
            agregados[destino] = (origen, "sum")
    if "Dias_Emision_Radicacion" in df.columns:
        agregados["Días promedio de radicación"] = ("Dias_Emision_Radicacion", "mean")
    if "Soporte_Incompleto" in df.columns:
        agregados["Soportes incompletos"] = ("Soporte_Incompleto", "sum")
    if "Tiene_Glosa" in df.columns:
        agregados["Facturas con glosa"] = ("Tiene_Glosa", "sum")
    if "Tiene_Saldo" in df.columns:
        agregados["Facturas con saldo"] = ("Tiene_Saldo", "sum")
    tabla = df.groupby(grupos, observed=True, dropna=False).agg(**agregados).reset_index()
    if "Facturas" not in tabla.columns:
        tabla["Facturas"] = tabla["Registros"]
    neto = tabla["Facturación neta"]
    tabla["Ticket promedio"] = division_serie(neto, tabla["Facturas"])
    if "Glosa" in tabla.columns:
        tabla["% glosa"] = division_serie(tabla["Glosa"], neto) * 100
    if "Pagado" in tabla.columns:
        tabla["% recaudo"] = division_serie(tabla["Pagado"], neto) * 100
    if "Saldo" in tabla.columns:
        tabla["% cartera"] = division_serie(tabla["Saldo"], neto) * 100
    if "Soportes incompletos" in tabla.columns:
        tabla["% soportes incompletos"] = division_serie(tabla["Soportes incompletos"], tabla["Registros"]) * 100
    total = neto.sum()
    tabla["Participación en la facturación"] = division_serie(neto, pd.Series(total, index=tabla.index)) * 100
    for g in grupos:
        if tabla[g].dtype == object:
            tabla[g] = tabla[g].fillna(SIN_INFO)
    return tabla.sort_values("Facturación neta", ascending=False)


def clave_periodo(df, periodicidad):
    """Construye la clave de periodo según la periodicidad (mensual, trimestral o anual)."""
    if periodicidad == "Trimestral":
        return df["Año_Radicacion"].astype(str) + "-" + df["Trimestre_Radicacion"].astype(str)
    if periodicidad == "Anual":
        return df["Año_Radicacion"].astype(str)
    return df["Periodo_Cierre"]


@st.cache_data(show_spinner=False)
def construir_cierre_mensual(df, periodicidad="Mensual"):
    """Construye la tabla de cierre por periodo con crecimiento, acumulado anual y promedio móvil."""
    if "Periodo_Cierre" not in df.columns or df["Periodo_Cierre"].isna().all():
        return pd.DataFrame()
    datos = df[df["Periodo_Cierre"].notna()].copy()
    datos["Periodo"] = clave_periodo(datos, periodicidad)
    datos = datos[datos["Periodo"].notna() & ~datos["Periodo"].astype(str).str.contains("<NA>|nan", regex=True)]
    if datos.empty:
        return pd.DataFrame()
    tabla = agregar_por(datos, "Periodo").drop(columns=["Participación en la facturación"], errors="ignore")
    tabla = tabla.sort_values("Periodo").reset_index(drop=True)
    tabla["Año"] = tabla["Periodo"].astype(str).str.slice(0, 4)
    tabla["Crecimiento vs periodo anterior"] = tabla["Facturación neta"].pct_change().replace([np.inf, -np.inf], np.nan) * 100
    tabla["Acumulado anual"] = tabla.groupby("Año")["Facturación neta"].cumsum()
    tabla["Promedio móvil 3 periodos"] = tabla["Facturación neta"].rolling(3, min_periods=3).mean()
    orden = ["Periodo", "Registros", "Facturas", "Cantidad", "Facturación bruta", "Descuentos", "Copagos", "Facturación neta", "Glosa", "Pagado",
             "Saldo", "Ticket promedio", "% glosa", "% recaudo", "% cartera", "Crecimiento vs periodo anterior", "Acumulado anual",
             "Promedio móvil 3 periodos"]
    return tabla[[c for c in orden if c in tabla.columns]]


def periodo_anterior_consecutivo(ultimo, anterior):
    """Comprueba que dos periodos AAAA-MM sean consecutivos."""
    try:
        u, a = pd.Period(ultimo, freq="M"), pd.Period(anterior, freq="M")
        return (u - a).n == 1
    except Exception:  # noqa: BLE001
        return False


def periodos_comparables(periodos, registros):
    """Elige el último periodo comparable y su anterior consecutivo.

    Si el último periodo tiene menos de la mitad de los registros de la mediana de los anteriores,
    se considera posiblemente incompleto y se compara el periodo previo. Devuelve (último, anterior, nota) o None.
    """
    periodos, registros = list(periodos), list(registros)
    if len(periodos) < 2:
        return None
    nota = ""
    if len(periodos) >= 4 and registros[-1] < 0.5 * float(np.median(registros[:-1])):
        nota = (f"El periodo {periodos[-1]} tiene pocos registros frente a los anteriores y podría estar incompleto; "
                f"por eso se compara {periodos[-2]} con {periodos[-3]}.")
        periodos, registros = periodos[:-1], registros[:-1]
    if not periodo_anterior_consecutivo(periodos[-1], periodos[-2]):
        return None
    return periodos[-1], periodos[-2], nota


def comparacion_periodos(df):
    """Devuelve los KPI del último periodo comparable y del periodo inmediatamente anterior, si existe."""
    if "Periodo_Cierre" not in df.columns:
        return None
    conteos = df.groupby("Periodo_Cierre").size().sort_index()
    eleccion = periodos_comparables(conteos.index, conteos.values)
    if not eleccion:
        return None
    ultimo, anterior, nota = eleccion
    return {"ultimo": ultimo, "anterior": anterior, "nota": nota,
            "k_ultimo": calcular_kpis(df[df["Periodo_Cierre"] == ultimo]),
            "k_anterior": calcular_kpis(df[df["Periodo_Cierre"] == anterior])}


def variacion_comparable(cierre):
    """Variación de la facturación neta entre los periodos mensuales comparables más recientes del cierre."""
    if cierre.empty or len(cierre) < 2:
        return None
    eleccion = periodos_comparables(cierre["Periodo"], cierre["Registros"] if "Registros" in cierre else cierre["Facturas"])
    if not eleccion:
        return None
    ultimo, anterior, nota = eleccion
    actual = cierre.loc[cierre["Periodo"] == ultimo, "Facturación neta"].iloc[0]
    previo = cierre.loc[cierre["Periodo"] == anterior, "Facturación neta"].iloc[0]
    return {"ultimo": ultimo, "anterior": anterior, "nota": nota, "variacion": division_segura(actual - previo, abs(previo)) * 100}


def texto_delta(comp, clave, puntos=False):
    """Texto de variación del último periodo frente al anterior para st.metric."""
    if not comp:
        return None
    actual, previo = comp["k_ultimo"].get(clave), comp["k_anterior"].get(clave)
    if es_invalido(actual) or es_invalido(previo):
        return None
    if puntos:
        dif = actual - previo
        return f"{'+' if dif >= 0 else ''}{formatear_numero(dif, 1)} p.p. ({comp['ultimo']} vs {comp['anterior']})"
    variacion = division_segura(actual - previo, abs(previo))
    if es_invalido(variacion):
        return None
    return f"{'+' if variacion >= 0 else ''}{formatear_numero(variacion * 100, 1)} % ({comp['ultimo']} vs {comp['anterior']})"


# =============================================================================
# 12. ANÁLISIS ESPECÍFICOS
# =============================================================================

@st.cache_data(show_spinner=False)
def analizar_entidades(df):
    """Agregación por entidad pagadora."""
    return agregar_por(df, "Entidad_Pagadora")


@st.cache_data(show_spinner=False)
def analizar_servicios(df, columna):
    """Agregación por servicio o línea de servicio."""
    return agregar_por(df, columna)


@st.cache_data(show_spinner=False)
def analizar_sedes(df, columna):
    """Agregación por sede o régimen, incluidas facturas devueltas."""
    tabla = agregar_por(df, columna)
    if "Estado_Factura" in df.columns:
        devueltas = df[df["Estado_Factura"].astype(str).map(normalizar_nombre_columna).str.contains("devuel")]
        conteo = devueltas.groupby(columna, observed=True).size().rename("Devueltas")
        tabla = tabla.merge(conteo, left_on=columna, right_index=True, how="left")
        tabla["Devueltas"] = tabla["Devueltas"].fillna(0).astype(int)
    return tabla


@st.cache_data(show_spinner=False)
def analizar_estados(df):
    """Agregación por estado de factura."""
    tabla = agregar_por(df, "Estado_Factura")
    tabla["% de facturas"] = division_serie(tabla["Facturas"], pd.Series(tabla["Facturas"].sum(), index=tabla.index)) * 100
    return tabla


def calcular_pareto(df, categoria, valor):
    """Análisis de Pareto: participación individual, acumulada, categorías hasta ~80 % e índices de concentración."""
    serie = df.groupby(categoria, observed=True)[valor].sum()
    serie = serie[serie > 0].sort_values(ascending=False)
    total = serie.sum()
    if serie.empty or total <= 0:
        return None
    tabla = pd.DataFrame({categoria: serie.index.astype(str), "Valor": serie.values})
    tabla["Participación"] = tabla["Valor"] / total * 100
    tabla["Participación acumulada"] = tabla["Participación"].cumsum()
    hasta_80 = int((tabla["Participación acumulada"] < 80).sum() + 1)
    hasta_80 = min(hasta_80, len(tabla))
    cuotas = tabla["Participación"] / 100
    return {"tabla": tabla, "n_80": hasta_80, "total_categorias": len(tabla),
            "hhi": float((cuotas ** 2).sum() * 10000),
            "cr3": float(tabla["Participación"].head(3).sum()), "cr5": float(tabla["Participación"].head(5).sum())}


def grafico_pareto(resultado, categoria, titulo):
    """Gráfico combinado: barras de valor, línea de % acumulado y referencia del 80 %."""
    tabla = resultado["tabla"].head(30)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=tabla[categoria], y=tabla["Valor"], name="Valor", marker_color=AZUL_MEDIO,
                         hovertemplate="%{x}<br>Valor: %{y:,.0f}<extra></extra>"), secondary_y=False)
    fig.add_trace(go.Scatter(x=tabla[categoria], y=tabla["Participación acumulada"], name="% acumulado", mode="lines+markers",
                             line=dict(color=NARANJA, width=3), hovertemplate="%{x}<br>Acumulado: %{y:.1f} %<extra></extra>"),
                  secondary_y=True)
    fig.add_hline(y=80, line_dash="dash", line_color=ROJO, secondary_y=True, annotation_text="Referencia 80 %",
                  annotation_position="top left")
    estilo_figura(fig, titulo, eje_y_moneda=True)
    fig.update_yaxes(title_text="% acumulado", range=[0, 105], secondary_y=True, ticksuffix=" %")
    return fig


@st.cache_data(show_spinner=False)
def calcular_estadisticos(serie):
    """Calcula estadísticos descriptivos completos de una variable numérica."""
    datos = pd.to_numeric(serie, errors="coerce").dropna()
    if len(datos) < 2:
        return None
    q1, q3 = datos.quantile(0.25), datos.quantile(0.75)
    conteo_moda = datos.value_counts()
    moda = conteo_moda.index[0] if not conteo_moda.empty and conteo_moda.iloc[0] > 1 else np.nan
    media = datos.mean()
    return {
        "Conteo": len(datos), "Media": media, "Mediana": datos.median(), "Moda": moda,
        "Frecuencia de la moda": int(conteo_moda.iloc[0]) if not conteo_moda.empty else 0,
        "Desviación estándar": datos.std(), "Varianza": datos.var(), "Mínimo": datos.min(), "Máximo": datos.max(),
        "Rango": datos.max() - datos.min(), "Cuartil 1": q1, "Cuartil 3": q3, "Rango intercuartílico": q3 - q1,
        "Percentil 90": datos.quantile(0.9), "Percentil 95": datos.quantile(0.95),
        "Coeficiente de variación (%)": division_segura(datos.std(), abs(media)) * 100,
        "Asimetría": float(stats.skew(datos, bias=False)) if datos.nunique() > 2 else np.nan,
        "Curtosis (exceso)": float(stats.kurtosis(datos, bias=False)) if datos.nunique() > 3 else np.nan,
    }


def interpretar_estadisticos(e, variable):
    """Interpretación automática basada únicamente en los valores calculados."""
    textos = [f"Se analizaron {formatear_numero(e['Conteo'])} valores válidos de **{variable}**."]
    if not es_invalido(e["Media"]) and not es_invalido(e["Mediana"]) and e["Mediana"] != 0:
        rel = (e["Media"] - e["Mediana"]) / abs(e["Mediana"]) * 100
        if rel > 10:
            textos.append(f"La media supera a la mediana en {formatear_numero(rel, 1)} %: los datos muestran valores altos que elevan el promedio.")
        elif rel < -10:
            textos.append(f"La media es inferior a la mediana en {formatear_numero(abs(rel), 1)} %: se observan valores bajos que reducen el promedio.")
        else:
            textos.append("La media y la mediana son cercanas (diferencia inferior al 10 %).")
    asim = e["Asimetría"]
    if not es_invalido(asim):
        if asim > 1:
            textos.append(f"La asimetría es {formatear_numero(asim, 2)}: distribución con cola marcada hacia valores altos.")
        elif asim < -1:
            textos.append(f"La asimetría es {formatear_numero(asim, 2)}: distribución con cola marcada hacia valores bajos.")
        elif abs(asim) <= 0.5:
            textos.append(f"La asimetría es {formatear_numero(asim, 2)}: la distribución es aproximadamente simétrica según este indicador (esto no equivale a evaluar normalidad).")
        else:
            textos.append(f"La asimetría es {formatear_numero(asim, 2)}: asimetría moderada.")
    cv = e["Coeficiente de variación (%)"]
    if not es_invalido(cv):
        nivel = "alta" if cv > 100 else ("moderada" if cv > 50 else "baja")
        textos.append(f"El coeficiente de variación es {formatear_numero(cv, 1)} %: dispersión relativa {nivel}.")
    kurt = e["Curtosis (exceso)"]
    if not es_invalido(kurt) and kurt > 3:
        textos.append(f"La curtosis en exceso es {formatear_numero(kurt, 2)}: hay más valores extremos que en una distribución de colas ligeras.")
    textos.append("Para afirmar normalidad se requiere una prueba formal (ver pestaña «Pruebas estadísticas»).")
    return textos


@st.cache_data(show_spinner=False)
def calcular_correlaciones(df, columnas, metodo):
    """Matriz de correlación (Pearson o Spearman) con las variables disponibles."""
    metodo_pd = "pearson" if metodo == "Pearson" else "spearman"
    datos = df[list(columnas)].apply(pd.to_numeric, errors="coerce")
    datos = datos.loc[:, datos.notna().sum() >= 3]
    datos = datos.loc[:, datos.std(numeric_only=True) > 0]
    if datos.shape[1] < 2:
        return pd.DataFrame()
    return datos.corr(method=metodo_pd)


def detectar_atipicos(df, variable, metodo, parametros, grupo=None):
    """Detecta valores atípicos (IQR, puntaje Z o percentiles), global o por grupo.

    Devuelve (máscara, límite_inferior, límite_superior) como series alineadas al DataFrame.
    """
    valores = pd.to_numeric(df[variable], errors="coerce")
    agrupado = valores.groupby(df[grupo], observed=True) if grupo else None

    def transformar(funcion):
        return agrupado.transform(funcion) if grupo else pd.Series(funcion(valores), index=df.index)

    if metodo == "Rango intercuartílico (IQR)":
        k = parametros["k"]
        q1 = transformar(lambda s: s.quantile(0.25))
        q3 = transformar(lambda s: s.quantile(0.75))
        inferior, superior = q1 - k * (q3 - q1), q3 + k * (q3 - q1)
    elif metodo == "Puntaje Z":
        z = parametros["z"]
        media = transformar(lambda s: s.mean())
        desv = transformar(lambda s: s.std())
        inferior, superior = media - z * desv, media + z * desv
    else:
        inferior = transformar(lambda s: s.quantile(parametros["p_inf"] / 100))
        superior = transformar(lambda s: s.quantile(parametros["p_sup"] / 100))
    mascara = (valores < inferior) | (valores > superior)
    return mascara.fillna(False), inferior, superior


def ejecutar_prueba_estadistica(tipo, datos, alfa):
    """Ejecuta la prueba estadística seleccionada y devuelve un diccionario de resultados."""
    if tipo == "t":
        a, b = datos
        estadistico, p = stats.ttest_ind(a, b, equal_var=False)
        return {"prueba": "t de Student (Welch, varianzas no asumidas iguales)", "estadistico": estadistico, "p": p,
                "h0": "Las medias de los dos grupos son iguales.", "h1": "Las medias de los dos grupos son diferentes."}
    if tipo == "mw":
        a, b = datos
        estadistico, p = stats.mannwhitneyu(a, b, alternative="two-sided")
        return {"prueba": "Mann-Whitney U", "estadistico": estadistico, "p": p,
                "h0": "Las distribuciones de los dos grupos son iguales.", "h1": "Las distribuciones de los dos grupos difieren."}
    if tipo == "anova":
        estadistico, p = stats.f_oneway(*datos)
        return {"prueba": "ANOVA de una vía", "estadistico": estadistico, "p": p,
                "h0": "Todas las medias de los grupos son iguales.", "h1": "Al menos una media de grupo es diferente."}
    if tipo == "kruskal":
        estadistico, p = stats.kruskal(*datos)
        return {"prueba": "Kruskal-Wallis", "estadistico": estadistico, "p": p,
                "h0": "Las distribuciones de los grupos son iguales.", "h1": "Al menos un grupo tiene una distribución diferente."}
    if tipo == "chi2":
        estadistico, p, gl, esperadas = stats.chi2_contingency(datos)
        n = datos.to_numpy().sum()
        k = min(datos.shape) - 1
        cramer = np.sqrt(estadistico / (n * k)) if n and k > 0 else np.nan
        return {"prueba": "Chi-cuadrado de independencia", "estadistico": estadistico, "p": p, "gl": gl,
                "pct_esperadas_bajas": float((esperadas < 5).mean() * 100), "cramer": cramer,
                "h0": "Las dos variables son independientes.", "h1": "Existe asociación entre las dos variables."}
    if tipo == "shapiro":
        estadistico, p = stats.shapiro(datos)
        return {"prueba": "Shapiro-Wilk", "estadistico": estadistico, "p": p,
                "h0": "Los datos provienen de una distribución normal.", "h1": "Los datos no provienen de una distribución normal."}
    raise ValueError("Prueba no reconocida")


# =============================================================================
# 13. INSIGHTS Y ALERTAS
# =============================================================================

def nuevo_insight(categoria, hallazgo, valor, contexto, implicacion, accion):
    """Construye un insight con hallazgo calculado, interpretación y recomendación."""
    return {"Categoría": categoria, "Hallazgo": hallazgo, "Valor observado": valor, "Contexto": contexto,
            "Posible implicación": implicacion, "Acción sugerida para revisión": accion}


def generar_insights(df, validacion, umbrales):
    """Genera insights mediante reglas programadas sobre el DataFrame filtrado (sin modelos externos)."""
    insights = []
    if df.empty:
        return insights
    kpis = calcular_kpis(df)
    cierre = construir_cierre_mensual(df, "Mensual")
    if not cierre.empty and len(cierre) >= 2:
        fila_max = cierre.loc[cierre["Facturación neta"].idxmax()]
        insights.append(nuevo_insight(
            "Tendencia", f"El periodo con mayor facturación neta es {fila_max['Periodo']}.", formatear_moneda(fila_max["Facturación neta"]),
            f"Se compararon {len(cierre)} periodos con datos.",
            "Los datos muestran un pico de facturación en ese periodo.",
            "Conviene revisar qué servicios o entidades explican ese comportamiento."))
        crec = cierre["Crecimiento vs periodo anterior"].dropna()
        if not crec.empty and crec.min() < 0:
            fila_caida = cierre.loc[crec.idxmin()]
            insights.append(nuevo_insight(
                "Tendencia", f"La mayor caída frente al periodo anterior se observa en {fila_caida['Periodo']}.",
                formatear_porcentaje(fila_caida["Crecimiento vs periodo anterior"]),
                "Variación porcentual de la facturación neta entre periodos consecutivos de la tabla.",
                "Se observa una reducción relevante; la información disponible no permite establecer la causa.",
                "Este resultado podría justificar una revisión de radicaciones pendientes o cambios de volumen."))
        variacion = variacion_comparable(cierre)
        if variacion and not es_invalido(variacion["variacion"]):
            insights.append(nuevo_insight(
                "Tendencia", f"Variación del periodo {variacion['ultimo']} frente a {variacion['anterior']}.",
                formatear_porcentaje(variacion["variacion"]),
                "Comparación entre los periodos mensuales comparables más recientes. " + (variacion["nota"] or "El último periodo podría estar incompleto."),
                "Indica la dirección reciente de la facturación.", "Conviene confirmar que los periodos comparados estén cerrados antes de concluir."))
        if len(cierre) >= 4:
            y = cierre["Facturación neta"].to_numpy(dtype=float)
            pendiente, _, r, p, _ = stats.linregress(np.arange(len(y)), y)
            if p < 0.05:
                direccion = "creciente" if pendiente > 0 else "decreciente"
                insights.append(nuevo_insight(
                    "Tendencia", f"Se observa una tendencia lineal {direccion} de la facturación neta mensual.",
                    f"{formatear_moneda(pendiente)} por periodo (p = {formatear_numero(p, 3)})",
                    f"Regresión lineal simple sobre {len(y)} periodos.", "La tendencia es una descripción estadística, no una proyección.",
                    "Conviene revisar si la tendencia se mantiene al excluir periodos atípicos o incompletos."))
    if tiene(df, "Entidad_Pagadora"):
        ent = analizar_entidades(df)
        principal = ent.iloc[0]
        insights.append(nuevo_insight(
            "Entidades", f"La entidad con mayor facturación neta es {principal['Entidad_Pagadora']}.",
            f"{formatear_moneda(principal['Facturación neta'])} ({formatear_porcentaje(principal['Participación en la facturación'])} del total)",
            f"Se compararon {len(ent)} entidades.", "Una alta participación implica dependencia de ese pagador.",
            "Conviene revisar las condiciones de radicación y recaudo con esta entidad."))
        if "Saldo" in ent.columns and ent["Saldo"].max() > 0:
            fila = ent.loc[ent["Saldo"].idxmax()]
            insights.append(nuevo_insight("Cartera", f"La entidad con mayor saldo de cartera es {fila['Entidad_Pagadora']}.",
                                          formatear_moneda(fila["Saldo"]), "Suma del saldo de cartera en el conjunto filtrado.",
                                          "Concentra la mayor exposición de cartera.", "Conviene revisar la antigüedad y el estado de esas facturas."))
        relevantes = ent[ent["Facturas"] >= max(5, ent["Facturas"].median() * 0.2)]
        if relevantes.empty:
            relevantes = ent
        if "% recaudo" in ent.columns and relevantes["% recaudo"].notna().any():
            fila = relevantes.loc[relevantes["% recaudo"].idxmin()]
            insights.append(nuevo_insight("Recaudo", f"La entidad con menor porcentaje de recaudo es {fila['Entidad_Pagadora']}.",
                                          formatear_porcentaje(fila["% recaudo"]), "Se consideran entidades con un número representativo de facturas.",
                                          "Los datos muestran un recaudo inferior al de las demás entidades.",
                                          "Este resultado podría justificar una revisión de pagos pendientes y conciliaciones."))
        if "% glosa" in ent.columns and relevantes["% glosa"].notna().any() and relevantes["% glosa"].max() > 0:
            fila = relevantes.loc[relevantes["% glosa"].idxmax()]
            insights.append(nuevo_insight("Glosas", f"La entidad con mayor porcentaje de glosa es {fila['Entidad_Pagadora']}.",
                                          formatear_porcentaje(fila["% glosa"]), "Valor glosado sobre valor neto por entidad.",
                                          "Se observa una severidad de glosa superior a la del resto.",
                                          "Conviene revisar los motivos de glosa y la calidad de los soportes enviados a esta entidad."))
        pareto = calcular_pareto(df, "Entidad_Pagadora", "Valor_Neto")
        if pareto and pareto["total_categorias"] >= 3:
            insights.append(nuevo_insight("Concentración", f"{pareto['n_80']} de {pareto['total_categorias']} entidades explican cerca del 80 % de la facturación.",
                                          f"Top 3: {formatear_porcentaje(pareto['cr3'])} · HHI: {formatear_numero(pareto['hhi'])}",
                                          "Análisis de Pareto usado como herramienta descriptiva, no como regla.",
                                          "Una alta concentración aumenta la sensibilidad a cambios de pocos pagadores.",
                                          "Conviene monitorear periódicamente a los pagadores principales."))
        if tiene(df, "Saldo_Cartera"):
            pareto_c = calcular_pareto(df, "Entidad_Pagadora", "Saldo_Cartera")
            if pareto_c and pareto_c["total_categorias"] >= 3:
                insights.append(nuevo_insight("Concentración", f"{pareto_c['n_80']} de {pareto_c['total_categorias']} entidades concentran cerca del 80 % de la cartera.",
                                              f"Top 3: {formatear_porcentaje(pareto_c['cr3'])}", "Concentración del saldo de cartera.",
                                              "La exposición de cartera se concentra en pocos pagadores.",
                                              "Conviene priorizar la gestión de cobro de esas entidades."))
    if tiene(df, "Servicio"):
        serv = analizar_servicios(df, "Servicio")
        insights.append(nuevo_insight("Servicios", f"El servicio con mayor facturación neta es {serv.iloc[0]['Servicio']}.",
                                      formatear_moneda(serv.iloc[0]["Facturación neta"]), f"Se compararon {len(serv)} servicios.",
                                      "Es el principal generador de ingresos facturados.", "Conviene revisar su recaudo y glosas."))
        if "Saldo" in serv.columns and serv["Saldo"].max() > 0:
            fila = serv.loc[serv["Saldo"].idxmax()]
            insights.append(nuevo_insight("Cartera", f"El servicio con mayor cartera es {fila['Servicio']}.", formatear_moneda(fila["Saldo"]),
                                          "Suma del saldo por servicio.", "Concentra la mayor proporción de saldo pendiente.",
                                          "Conviene revisar el estado de las facturas de este servicio."))
    if tiene(df, "Dias_Emision_Radicacion"):
        for columna, nombre in (("Sede", "sede"), ("Canal_Radicacion", "canal")):
            if tiene(df, columna):
                tiempos = df.groupby(columna, observed=True)["Dias_Emision_Radicacion"].agg(["mean", "size"])
                tiempos = tiempos[tiempos["size"] >= 5]
                if len(tiempos) >= 2:
                    nombre_max = tiempos["mean"].idxmax()
                    insights.append(nuevo_insight("Radicación", f"La {nombre} con mayor tiempo promedio de radicación es {nombre_max}.",
                                                  f"{formatear_numero(tiempos.loc[nombre_max, 'mean'], 1)} días",
                                                  f"Promedio general: {formatear_numero(df['Dias_Emision_Radicacion'].mean(), 1)} días.",
                                                  "Se observa un tiempo superior al de las demás categorías.",
                                                  "Conviene revisar el flujo de radicación de esta categoría."))
    if "Soporte_Incompleto" in df.columns and df["Soporte_Completo"].ne(SIN_INFO).any():
        pct = df["Soporte_Incompleto"].mean() * 100
        insights.append(nuevo_insight("Soportes", "Porcentaje de facturas con soporte incompleto.", formatear_porcentaje(pct),
                                      f"Calculado sobre {formatear_numero(len(df))} registros filtrados.",
                                      "Los soportes incompletos pueden asociarse con devoluciones o glosas.",
                                      "Conviene revisar los procesos de verificación de soportes antes de radicar."))
        if tiene(df, "Glosa_Inicial") and df["Soporte_Incompleto"].nunique() == 2:
            frecuencia = df.groupby("Soporte_Incompleto")["Tiene_Glosa"].mean() * 100
            insights.append(nuevo_insight("Soportes", "Frecuencia de glosa según soporte.",
                                          f"Incompleto: {formatear_porcentaje(frecuencia.get(True))} · Completo: {formatear_porcentaje(frecuencia.get(False))}",
                                          "Proporción de facturas con glosa en cada grupo.",
                                          "Es una asociación observada; no demuestra causalidad.",
                                          "Conviene revisar si las glosas de facturas con soporte incompleto comparten motivos."))
    if tiene(df, "Glosa_Inicial"):
        insights.append(nuevo_insight("Glosas", "Proporción de facturas con glosa.", formatear_porcentaje(df["Tiene_Glosa"].mean() * 100),
                                      f"Severidad global: {formatear_porcentaje(kpis['pct_glosa'])} del valor neto.",
                                      "Mide la frecuencia de glosa (no su valor).", "Conviene analizar glosas por entidad y servicio."))
    if validacion:
        n_inc = validacion.get("filas_inconsistentes", 0)
        if n_inc:
            insights.append(nuevo_insight("Calidad", "Registros con alguna inconsistencia lógica o financiera.", formatear_numero(n_inc),
                                          "Detectados en el diagnóstico de calidad sobre el archivo cargado.",
                                          "Pueden afectar la precisión de los indicadores.", "Conviene revisar el detalle en «Calidad de datos»."))
    if tiene(df, "Valor_Neto") and df["Valor_Neto"].notna().sum() >= 10:
        mascara, _, _ = detectar_atipicos(df, "Valor_Neto", "Rango intercuartílico (IQR)", {"k": 1.5})
        insights.append(nuevo_insight("Atípicos", "Facturas con valor neto atípico (método IQR, k = 1,5).", formatear_numero(int(mascara.sum())),
                                      f"Equivale al {formatear_porcentaje(mascara.mean() * 100)} de los registros.",
                                      "Un valor atípico no es necesariamente un error.",
                                      "Conviene revisar si corresponden a servicios de alta complejidad."))
    return insights


def clasificar_semaforo(valor, verde, amarillo, mayor_es_mejor=False):
    """Clasifica un valor en Verde, Amarillo o Rojo según umbrales definidos por el usuario."""
    if es_invalido(valor):
        return "Sin dato"
    if mayor_es_mejor:
        return "Verde" if valor >= verde else ("Amarillo" if valor >= amarillo else "Rojo")
    return "Verde" if valor <= verde else ("Amarillo" if valor <= amarillo else "Rojo")


def generar_alertas(df, umbrales):
    """Genera semáforos con los umbrales configurados (parámetros demostrativos y editables)."""
    kpis = calcular_kpis(df)
    cierre = construir_cierre_mensual(df, "Mensual")
    variacion = variacion_comparable(cierre)
    crecimiento = variacion["variacion"] if variacion else np.nan
    u = umbrales
    fmt = lambda v: formatear_numero(v, 1)  # noqa: E731
    filas = [
        ("Porcentaje de glosa", kpis["pct_glosa"], formatear_porcentaje(kpis["pct_glosa"]),
         clasificar_semaforo(kpis["pct_glosa"], u["glosa_verde"], u["glosa_amarillo"]), f"Verde ≤ {fmt(u['glosa_verde'])} % · Amarillo ≤ {fmt(u['glosa_amarillo'])} %"),
        ("Porcentaje de cartera", kpis["pct_cartera"], formatear_porcentaje(kpis["pct_cartera"]),
         clasificar_semaforo(kpis["pct_cartera"], u["cartera_verde"], u["cartera_amarillo"]), f"Verde ≤ {fmt(u['cartera_verde'])} % · Amarillo ≤ {fmt(u['cartera_amarillo'])} %"),
        ("Porcentaje de recaudo", kpis["pct_recaudo"], formatear_porcentaje(kpis["pct_recaudo"]),
         clasificar_semaforo(kpis["pct_recaudo"], u["recaudo_verde"], u["recaudo_amarillo"], True), f"Verde ≥ {fmt(u['recaudo_verde'])} % · Amarillo ≥ {fmt(u['recaudo_amarillo'])} %"),
        ("Días promedio de radicación", kpis["dias"], formatear_numero(kpis["dias"], 1) + " días" if not es_invalido(kpis["dias"]) else "Sin dato",
         clasificar_semaforo(kpis["dias"], u["dias_verde"], u["dias_amarillo"]), f"Verde ≤ {fmt(u['dias_verde'])} · Amarillo ≤ {fmt(u['dias_amarillo'])} días"),
        ("Soportes incompletos", kpis["pct_soporte_incompleto"], formatear_porcentaje(kpis["pct_soporte_incompleto"]),
         clasificar_semaforo(kpis["pct_soporte_incompleto"], u["soporte_verde"], u["soporte_amarillo"]), f"Verde ≤ {fmt(u['soporte_verde'])} % · Amarillo ≤ {fmt(u['soporte_amarillo'])} %"),
        ("Crecimiento del último periodo comparable" + (f" ({variacion['ultimo']} vs {variacion['anterior']})" if variacion else ""), crecimiento, formatear_porcentaje(crecimiento),
         clasificar_semaforo(crecimiento, u["crec_verde"], u["crec_amarillo"], True), f"Verde ≥ {fmt(u['crec_verde'])} % · Amarillo ≥ {fmt(u['crec_amarillo'])} %"),
    ]
    return pd.DataFrame(filas, columns=["Indicador", "Valor numérico", "Valor", "Estado", "Umbrales configurados"])


# =============================================================================
# 14. DESCARGAS (EN MEMORIA)
# =============================================================================

@st.cache_data(show_spinner=False)
def convertir_dataframe_csv(df):
    """Convierte un DataFrame a CSV (separado por punto y coma, compatible con Excel en español)."""
    salida = io.StringIO()
    df.to_csv(salida, index=False, sep=";", decimal=",", date_format="%Y-%m-%d")
    return salida.getvalue().encode("utf-8-sig")


@st.cache_data(show_spinner=False)
def convertir_dataframe_excel(hojas):
    """Convierte uno o varios DataFrames a un archivo Excel en memoria (una hoja por DataFrame)."""
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as escritor:
        for nombre, tabla in hojas.items():
            datos = tabla.copy()
            for col in datos.columns:
                if isinstance(datos[col].dtype, pd.CategoricalDtype):
                    datos[col] = datos[col].astype(str)
            datos.to_excel(escritor, sheet_name=str(nombre)[:31], index=False)
    return salida.getvalue()


def boton_descarga(etiqueta, datos, nombre, mime, clave):
    """Botón de descarga con manejo de errores."""
    try:
        st.download_button(etiqueta, data=datos, file_name=nombre, mime=mime, key=clave, width="stretch")
    except TypeError:
        st.download_button(etiqueta, data=datos, file_name=nombre, mime=mime, key=clave, use_container_width=True)
    except Exception:  # noqa: BLE001
        st.warning(f"No fue posible preparar la descarga «{nombre}».")


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# =============================================================================
# 15. ESTADO DE SESIÓN Y BARRA LATERAL
# =============================================================================

def inicializar_estado():
    """Inicializa las claves de st.session_state usadas durante la sesión."""
    valores = {"clave_cargador": 0, "moneda": "COP", "pagina": PAGINAS[0], "umbrales": dict(UMBRALES_DEFECTO),
               "limpieza_aplicar": True, "limpieza_vacias": True, "limpieza_duplicados": False, "limpieza_neto": False,
               "limpieza_fechas": False, "limpieza_rellenar": True, "limpieza_fechas_conv": True, "limpieza_moneda_conv": True}
    for clave, valor in valores.items():
        st.session_state.setdefault(clave, valor)


def reiniciar_estado_datos():
    """Reinicia el estado dependiente del archivo (mapeo, filtros, limpieza) al cambiar de archivo."""
    for clave in list(st.session_state.keys()):
        if clave.startswith(("f_", "map_", "hoja_", "separador_", "codificacion_")) or clave in (
                "mapeo", "mapeo_firma", "mapeo_manuales", "df_original", "firma_archivo"):
            del st.session_state[clave]


def reiniciar_aplicacion():
    """Callback del botón «Reiniciar aplicación»: limpia todo y vuelve a la pantalla inicial."""
    siguiente = st.session_state.get("clave_cargador", 0) + 1
    for clave in list(st.session_state.keys()):
        del st.session_state[clave]
    st.session_state["clave_cargador"] = siguiente
    st.session_state["pagina"] = PAGINAS[0]


def restaurar_datos_originales():
    """Callback del botón «Restaurar datos originales»: desactiva todas las transformaciones de limpieza."""
    st.session_state["limpieza_aplicar"] = False
    for clave in ("limpieza_vacias", "limpieza_duplicados", "limpieza_neto", "limpieza_fechas", "limpieza_rellenar"):
        st.session_state[clave] = False
    st.session_state["limpieza_fechas_conv"] = True
    st.session_state["limpieza_moneda_conv"] = True


def opciones_barra_lateral():
    """Opciones generales: navegación, moneda, limpieza, rangos, tolerancia y alertas."""
    st.sidebar.markdown("### 🧭 Navegación")
    pagina = st.sidebar.radio("Sección", PAGINAS, key="pagina", label_visibility="collapsed")
    st.sidebar.markdown("### ⚙️ Preferencias")
    st.sidebar.selectbox("Etiqueta de moneda", list(SIMBOLOS_MONEDA.keys()), key="moneda",
                         help="Cambio solamente visual: no se aplican tasas de cambio.")
    with st.sidebar.expander("🧹 Limpieza controlada", expanded=False):
        st.caption("Las transformaciones se aplican sobre una copia; los datos originales se conservan en memoria.")
        opciones = {
            "aplicar": st.checkbox("Aplicar transformaciones de limpieza", key="limpieza_aplicar"),
            "vacias": st.checkbox("Eliminar filas completamente vacías", key="limpieza_vacias"),
            "duplicados": st.checkbox("Eliminar duplicados exactos", key="limpieza_duplicados"),
            "neto_invalido": st.checkbox("Excluir registros con Valor_Neto inválido", key="limpieza_neto"),
            "fechas_invalidas": st.checkbox("Excluir registros con fechas inválidas", key="limpieza_fechas"),
            "rellenar": st.checkbox("Reemplazar categorías vacías por «Sin información»", key="limpieza_rellenar"),
        }
        conv_fechas = st.checkbox("Convertir columnas de fecha (detección robusta de formatos)", key="limpieza_fechas_conv")
        conv_moneda = st.checkbox("Convertir columnas monetarias (símbolos, miles y decimales)", key="limpieza_moneda_conv",
                                  help="Si se desactiva, solo se aceptan números simples; los textos con símbolos quedarán nulos.")
        st.button("Restaurar datos originales", on_click=restaurar_datos_originales, key="btn_restaurar", width="stretch")
    with st.sidebar.expander("⏱️ Rangos analíticos de radicación", expanded=False):
        st.caption("Parámetros analíticos configurables; no son estándares normativos.")
        l1 = st.number_input("Límite rango 1 (días)", min_value=0, max_value=365, value=3, step=1, key="rango_l1")
        l2 = st.number_input("Límite rango 2 (días)", min_value=1, max_value=366, value=7, step=1, key="rango_l2")
        l3 = st.number_input("Límite rango 3 (días)", min_value=2, max_value=367, value=15, step=1, key="rango_l3")
        if not (l1 < l2 < l3):
            st.warning("Los límites deben ser crecientes. Se ajustaron automáticamente.")
            l2 = max(int(l2), int(l1) + 1)
            l3 = max(int(l3), int(l2) + 1)
    with st.sidebar.expander("🚦 Configuración de alertas", expanded=False):
        st.caption("Valores predeterminados demostrativos y editables. No son estándares oficiales ni metas contractuales.")
        u = st.session_state["umbrales"]
        etiquetas = {"glosa": "% glosa", "cartera": "% cartera", "recaudo": "% recaudo", "dias": "Días de radicación",
                     "soporte": "% soportes incompletos", "crec": "Crecimiento (%)"}
        for clave, etiqueta in etiquetas.items():
            c1, c2 = st.columns(2)
            u[f"{clave}_verde"] = c1.number_input(f"{etiqueta}: verde", value=float(u[f"{clave}_verde"]), step=1.0, key=f"um_{clave}_v")
            u[f"{clave}_amarillo"] = c2.number_input(f"{etiqueta}: amarillo", value=float(u[f"{clave}_amarillo"]), step=1.0, key=f"um_{clave}_a")
    tolerancia = st.sidebar.number_input("Tolerancia de redondeo para validaciones", min_value=0.0, value=1.0, step=1.0, key="tolerancia",
                                         help="Diferencia máxima aceptada al validar las identidades financieras.")
    st.sidebar.button("🔄 Reiniciar aplicación", on_click=reiniciar_aplicacion, key="btn_reiniciar", width="stretch")
    return pagina, opciones, conv_fechas, conv_moneda, (int(l1), int(l2), int(l3)), tolerancia


# =============================================================================
# 16. PÁGINAS DE LA APLICACIÓN
# =============================================================================

def mostrar_bienvenida():
    """Contenido de la página de inicio antes de cargar un archivo."""
    st.markdown(f'<div class="aviso-privacidad">🔒 {AVISO_PRIVACIDAD}</div>', unsafe_allow_html=True)
    st.subheader("Bienvenido")
    st.write("Esta aplicación permite **cargar, validar, explorar y analizar** archivos de facturación del sector salud: "
             "facturación bruta y neta, radicación, glosas, recaudo, cartera, estados, entidades, sedes y servicios.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📄 Archivos aceptados")
        st.markdown("- Excel **.xlsx** (se elige la hoja; se usa «Facturacion» si existe).\n"
                    "- **CSV** (detección de separador y codificación).\n"
                    "- Los archivos **.xls** antiguos deben guardarse como .xlsx o CSV.")
        st.markdown("#### 🧩 Columnas mínimas")
        st.markdown("\n".join(f"- `{c}`" for c in COLUMNAS_MINIMAS))
        st.markdown("#### ⭐ Columnas recomendadas")
        st.markdown("\n".join(f"- `{c}`" for c in COLUMNAS_RECOMENDADAS))
    with c2:
        st.markdown("#### 🪜 Instrucciones")
        st.markdown("1. Cargue el archivo en la barra lateral.\n2. Revise la hoja y la correspondencia de columnas.\n"
                    "3. Consulte el diagnóstico de calidad y ajuste la limpieza si lo necesita.\n"
                    "4. Use los filtros globales y recorra las secciones del menú.\n5. Descargue los resultados.")
        st.markdown("#### 🔐 Privacidad")
        st.markdown("El archivo se procesa **solo en memoria** durante la sesión. No se guarda, no se envía a servicios externos y "
                    "se solicita nuevamente en cada sesión nueva.")
        st.markdown("#### 🔍 Uso exploratorio")
        st.markdown(AVISO_EXPLORATORIO)
    with st.expander("Ver todas las columnas esperadas"):
        st.dataframe(pd.DataFrame([{"Campo": c, "Descripción": m["etiqueta"],
                                    "Grupo": "Mínima" if c in ("ID_Factura", "Valor_Neto", "Fecha_Radicacion", "Periodo_Cierre")
                                    else ("Recomendada" if c in COLUMNAS_RECOMENDADAS else "Opcional")} for c, m in CAMPOS.items()]),
                     hide_index=True)


def mostrar_inicio(ctx):
    """Página 1: inicio y carga de datos."""
    st.header("1. Inicio y carga de datos")
    if ctx["df_original"] is None:
        mostrar_bienvenida()
        return
    st.markdown(f'<div class="aviso-privacidad">🔒 {AVISO_PRIVACIDAD}</div>', unsafe_allow_html=True)
    info, base = ctx["info"], ctx["df_base"]
    st.markdown(f"**Archivo:** {info['nombre']}  \n**Hoja utilizada:** {info['hoja']}")
    c1, c2, c3 = st.columns(3)
    c1.metric("Tamaño", f"{formatear_numero(info['tamano'] / 1024, 1)} KB")
    c2.metric("Filas", formatear_numero(len(ctx["df_original"])))
    c3.metric("Columnas", formatear_numero(ctx["df_original"].shape[1]))
    if ctx["equivalencias"]:
        with st.expander(f"Equivalencias de encabezados detectadas ({len(ctx['equivalencias'])})", expanded=False):
            st.caption("Antes de renombrar se muestran las equivalencias encontradas. Puede ajustarlas en la correspondencia de columnas.")
            st.dataframe(pd.DataFrame(ctx["equivalencias"]), hide_index=True)
    if ctx["ambiguas"]:
        st.warning("Se encontraron campos con más de una columna candidata: " +
                   "; ".join(f"{k}: {', '.join(v)}" for k, v in ctx["ambiguas"].items()) +
                   ". Verifique la asignación en la correspondencia de columnas.")
    mapear_columnas(list(ctx["df_original"].columns), ctx["mapeo"], ctx["mapeo_detectado"])
    if not ctx["minimas_ok"]:
        mostrar_faltantes_minimas(ctx)
        return
    if base is None or base.empty:
        st.warning("No quedaron registros después de la limpieza. Revise las opciones de limpieza.")
        return
    periodos = base["Periodo_Cierre"].dropna() if "Periodo_Cierre" in base.columns else pd.Series(dtype=str)

    def conteo(columna):
        if columna not in base.columns:
            return "No disponible"
        return formatear_numero(base[columna].dropna().loc[lambda s: s != SIN_INFO].nunique())

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Periodo inicial", periodos.min() if not periodos.empty else "No identificado")
    c2.metric("Periodo final", periodos.max() if not periodos.empty else "No identificado")
    c3.metric("Entidades", conteo("Entidad_Pagadora"))
    c4.metric("Sedes", conteo("Sede"))
    c5.metric("Servicios", conteo("Servicio"))
    faltan_rec = [c for c in COLUMNAS_RECOMENDADAS if c not in base.columns]
    if faltan_rec:
        st.warning(f"Faltan columnas recomendadas: {', '.join(faltan_rec)}. Los análisis que dependen de ellas se desactivarán.")
    else:
        st.success("Se encontraron todas las columnas mínimas y recomendadas.")
    st.subheader("Vista previa")
    modo = st.radio("Registros a visualizar", ["Primeros registros", "Muestra aleatoria"], horizontal=True, key="vista_previa_modo")
    n = st.slider("Número de registros", 5, 100, 15, key="vista_previa_n")
    vista = ctx["df_original"].head(n) if modo == "Primeros registros" else ctx["df_original"].sample(min(n, len(ctx["df_original"])), random_state=42)
    st.dataframe(vista, hide_index=True)


def mostrar_faltantes_minimas(ctx):
    """Explica qué columnas mínimas faltan y cómo corregir el archivo."""
    faltan = faltantes_minimas(ctx["mapeo"])
    st.error("No se encontraron todas las columnas mínimas. Los análisis financieros principales están desactivados.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Columnas faltantes**")
        st.markdown("\n".join(f"- `{c}`" for c in faltan))
        st.markdown("**Nombres esperados**")
        st.markdown("\n".join(f"- `{c}`" for c in COLUMNAS_MINIMAS))
    with c2:
        st.markdown("**Columnas encontradas en el archivo**")
        st.write(", ".join(map(str, ctx["df_original"].columns)))
    st.info("Cómo corregirlo: 1) asigne las columnas en «Configurar correspondencia de columnas» (página de inicio), o "
            "2) renombre los encabezados del archivo con los nombres esperados y vuelva a cargarlo.")


def mostrar_calidad(ctx):
    """Página 2: diagnóstico de calidad de datos."""
    st.header("2. Diagnóstico de calidad de datos")
    v = ctx["validacion"]
    if v is None:
        st.info("Cargue un archivo y complete las columnas mínimas para ver el diagnóstico.")
        return
    puntaje = v["puntaje"]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Puntaje de calidad", f"{formatear_numero(puntaje['Total'], 1)} / 100", help=AVISO_CALIDAD)
    c2.metric("Completitud", formatear_porcentaje(puntaje["Completitud"]))
    c3.metric("Unicidad", formatear_porcentaje(puntaje["Unicidad"]))
    c4.metric("Validez", formatear_porcentaje(puntaje["Validez"]))
    c5.metric("Consistencia", formatear_porcentaje(puntaje["Consistencia"]))
    st.caption(f"ℹ️ {AVISO_CALIDAD}")
    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Filas analizadas", formatear_numero(v["filas"]))
    c2.metric("Columnas del archivo", formatear_numero(v["columnas_archivo"]))
    c3.metric("Filas duplicadas", formatear_numero(v["filas_duplicadas"]))
    c4.metric("ID de factura duplicados", formatear_numero(v["id_duplicados"]) if v["id_duplicados"] is not None else "No disponible")
    pestanas = st.tabs(["Problemas detectados", "Nulos por columna", "Conversiones de tipo", "Columnas", "Limpieza aplicada"])
    with pestanas[0]:
        filas = [("Radicados duplicados", v["radicados_duplicados"]), ("Valores netos iguales a cero", v["neto_cero"]),
                 ("Facturas sin entidad pagadora", v["sin_entidad"]), ("Facturas sin periodo", v["sin_periodo"])]
        filas += [(f"Fechas inválidas en {c}", n) for c, n in v["fechas_invalidas"].items()]
        filas += [(f"Valores no numéricos en {c}", n) for c, n in v["no_numericos"].items()]
        filas += [(f"Valores negativos en {c}", n) for c, n in v["negativos"].items()]
        filas += list(v["inconsistencias"].items())
        tabla = pd.DataFrame([(a, b) for a, b in filas if b is not None], columns=["Verificación", "Registros"])
        tabla["Estado"] = np.where(tabla["Registros"] > 0, "⚠️ Revisar", "✔️ Sin hallazgos")
        mostrar_tabla(tabla, enteros=["Registros"], altura=min(40 + 35 * len(tabla), 640))
        st.caption(f"Tolerancia de redondeo usada en las identidades financieras: {formatear_numero(st.session_state.get('tolerancia', 1), 2)}. "
                   "Identidades: Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota; Saldo_Cartera ≈ Valor_Neto − Glosa_Inicial − Valor_Pagado.")
        banderas = v["banderas"]
        if not banderas.empty and banderas.any(axis=1).any():
            with st.expander("Ver registros con inconsistencias (máximo 500)"):
                casos = ctx["df_tipado"].loc[banderas.any(axis=1)].head(500).copy()
                casos.insert(0, "Inconsistencias", banderas.loc[casos.index].apply(lambda fila: ", ".join(fila.index[fila]), axis=1))
                mostrar_tabla(casos, monedas=[c for c in COLUMNAS_DINERO if c in casos.columns], fechas=COLUMNAS_FECHA)
    with pestanas[1]:
        nulos = v["nulos"].sort_values("% nulos", ascending=False)
        mostrar_tabla(nulos, enteros=["Nulos"], porcentajes=["% nulos"])
        with bloque_seguro("gráfico de nulos"):
            fig = px.bar(nulos, x="Columna", y="% nulos", color_discrete_sequence=[AZUL_MEDIO])
            estilo_figura(fig, "Porcentaje de valores nulos por columna")
            fig.update_yaxes(ticksuffix=" %")
            mostrar_grafico(fig)
    with pestanas[2]:
        st.caption("Ninguna conversión modifica silenciosamente los datos: aquí se informa cada columna convertida.")
        if ctx["conversiones"]:
            mostrar_tabla(pd.DataFrame(ctx["conversiones"]), enteros=["Valores no convertidos", "Nulos antes", "Nulos después"])
        for adv in ctx["advertencias"]:
            st.warning(adv)
    with pestanas[3]:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Campos encontrados**")
            st.write(", ".join(v["encontradas"]) or "Ninguno")
        with c2:
            st.markdown("**Campos ausentes**")
            ausentes = v["ausentes"]
            st.write(", ".join(ausentes) if ausentes else "Ninguno")
            rec = [c for c in COLUMNAS_RECOMENDADAS if c in ausentes]
            if rec:
                st.warning(f"Columnas recomendadas ausentes: {', '.join(rec)}.")
    with pestanas[4]:
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros iniciales", formatear_numero(len(ctx["df_tipado"])))
        c2.metric("Registros después de la limpieza", formatear_numero(len(ctx["df_base"])))
        c3.metric("Registros excluidos", formatear_numero(len(ctx["df_tipado"]) - len(ctx["df_base"])))
        if ctx["exclusiones"]:
            mostrar_tabla(pd.DataFrame(ctx["exclusiones"]), enteros=["Registros excluidos"])
        else:
            st.info("No se aplicaron exclusiones.")


def mostrar_tarjetas_kpi(df):
    """Tarjetas KPI con comparación frente al periodo anterior cuando es posible."""
    k = calcular_kpis(df)
    comp = comparacion_periodos(df)
    tarjetas = [
        ("Facturación bruta", formatear_moneda_compacta(k["bruto"]), texto_delta(comp, "bruto"), "normal", f"Suma de Valor_Bruto: {formatear_moneda(k['bruto'])}."),
        ("Facturación neta", formatear_moneda_compacta(k["neto"]), texto_delta(comp, "neto"), "normal", f"Suma de Valor_Neto: {formatear_moneda(k['neto'])}."),
        ("Número de facturas", formatear_numero(k["facturas"]), texto_delta(comp, "facturas"), "normal", "Facturas únicas por ID_Factura."),
        ("Cantidad de servicios", formatear_numero(k["cantidad"]) if not es_invalido(k["cantidad"]) else "No calculable",
         texto_delta(comp, "cantidad"), "normal", "Suma de Cantidad."),
        ("Valor pagado", formatear_moneda_compacta(k["pagado"]), texto_delta(comp, "pagado"), "normal", f"Suma de Valor_Pagado: {formatear_moneda(k['pagado'])}."),
        ("Saldo de cartera", formatear_moneda_compacta(k["saldo"]), texto_delta(comp, "saldo"), "inverse", f"Suma de Saldo_Cartera: {formatear_moneda(k['saldo'])}."),
        ("Valor glosado", formatear_moneda_compacta(k["glosa"]), texto_delta(comp, "glosa"), "inverse", f"Suma de Glosa_Inicial: {formatear_moneda(k['glosa'])}."),
        ("% de recaudo", formatear_porcentaje(k["pct_recaudo"]), texto_delta(comp, "pct_recaudo", True), "normal", "Valor_Pagado / Valor_Neto × 100."),
        ("% de glosa", formatear_porcentaje(k["pct_glosa"]), texto_delta(comp, "pct_glosa", True), "inverse", "Glosa_Inicial / Valor_Neto × 100."),
        ("Ticket promedio", formatear_moneda(k["ticket"]), texto_delta(comp, "ticket"), "normal", "Valor_Neto / número de facturas."),
        ("Tiempo promedio de radicación", f"{formatear_numero(k['dias'], 1)} días" if not es_invalido(k["dias"]) else "No calculable",
         texto_delta(comp, "dias"), "inverse", "Promedio de Dias_Emision_Radicacion."),
    ]
    for inicio in range(0, len(tarjetas), 4):
        columnas = st.columns(4)
        for col, (etiqueta, valor, delta, color, ayuda) in zip(columnas, tarjetas[inicio:inicio + 4]):
            col.metric(etiqueta, valor, delta=delta, delta_color=color, help=ayuda)
    if comp:
        st.caption(f"Las variaciones comparan el periodo {comp['ultimo']} con el periodo inmediatamente anterior ({comp['anterior']}) "
                   "dentro de los filtros aplicados. Los valores principales corresponden a todo el conjunto filtrado. "
                   "«M» indica millones; el valor completo aparece en la ayuda de cada tarjeta.")
        if comp.get("nota"):
            st.info(comp["nota"])
    else:
        st.caption("No se muestran variaciones: no existe un periodo anterior consecutivo comparable en el conjunto filtrado.")
    return k


def mostrar_resumen(ctx):
    """Página 3: resumen ejecutivo."""
    st.header("3. Resumen ejecutivo")
    df, base = ctx["df"], ctx["df_base"]
    mostrar_tarjetas_kpi(df)
    with st.expander("Comparación: datos completos frente a datos filtrados"):
        kc, kf = calcular_kpis(base), calcular_kpis(df)
        comparacion = pd.DataFrame({
            "Indicador": ["Registros", "Facturas", "Facturación neta", "Valor pagado", "Saldo de cartera"],
            "Datos completos": [formatear_numero(len(base)), formatear_numero(kc["facturas"]), formatear_moneda(kc["neto"]),
                                formatear_moneda(kc["pagado"]), formatear_moneda(kc["saldo"])],
            "Datos filtrados": [formatear_numero(len(df)), formatear_numero(kf["facturas"]), formatear_moneda(kf["neto"]),
                                formatear_moneda(kf["pagado"]), formatear_moneda(kf["saldo"])],
        })
        st.dataframe(comparacion, hide_index=True)
    st.divider()
    cierre = construir_cierre_mensual(df, "Mensual")
    c1, c2 = st.columns(2)
    with c1, bloque_seguro("facturación neta mensual"):
        if not cierre.empty:
            fig = px.bar(cierre, x="Periodo", y="Facturación neta", color_discrete_sequence=[AZUL_MEDIO])
            mostrar_grafico(estilo_figura(fig, "Facturación neta mensual", eje_y_moneda=True))
    with c2, bloque_seguro("facturación, pagos y cartera"):
        series = [c for c in ["Facturación neta", "Pagado", "Saldo"] if c in cierre.columns]
        if not cierre.empty and len(series) > 1:
            largo = cierre.melt(id_vars="Periodo", value_vars=series, var_name="Concepto", value_name="Valor")
            fig = px.line(largo, x="Periodo", y="Valor", color="Concepto", markers=True)
            mostrar_grafico(estilo_figura(fig, "Facturación, pagos y cartera por periodo", eje_y_moneda=True))
        else:
            st.info("Se requieren Valor_Pagado o Saldo_Cartera para comparar facturación, pagos y cartera.")
    c1, c2 = st.columns(2)
    with c1, bloque_seguro("distribución de estados"):
        if tiene(df, "Estado_Factura"):
            estados = df["Estado_Factura"].value_counts().reset_index()
            estados.columns = ["Estado", "Registros"]
            fig = px.pie(estados, names="Estado", values="Registros", hole=0.45)
            fig.update_traces(textinfo="percent+label")
            mostrar_grafico(estilo_figura(fig, "Distribución de estados de factura"))
        else:
            st.info("No se encontró la columna Estado_Factura.")
    with c2, bloque_seguro("facturación por entidad"):
        if tiene(df, "Entidad_Pagadora"):
            ent = analizar_entidades(df).head(15).sort_values("Facturación neta")
            fig = px.bar(ent, x="Facturación neta", y="Entidad_Pagadora", orientation="h", color_discrete_sequence=[VERDE_AZULADO])
            mostrar_grafico(estilo_figura(fig, "Facturación neta por entidad pagadora (15 principales)", eje_x_moneda=True))
        else:
            st.info("No se encontró la columna Entidad_Pagadora.")
    c1, c2 = st.columns(2)
    with c1, bloque_seguro("participación por línea"):
        if tiene(df, "Linea_Servicio"):
            lineas = analizar_servicios(df, "Linea_Servicio")
            fig = px.pie(lineas, names="Linea_Servicio", values="Facturación neta", hole=0.45)
            fig.update_traces(textinfo="percent+label")
            mostrar_grafico(estilo_figura(fig, "Participación por línea de servicio"))
        else:
            st.info("No se encontró la columna Linea_Servicio.")
    with c2, bloque_seguro("glosa y recaudo por periodo"):
        if not cierre.empty and ("% glosa" in cierre.columns or "% recaudo" in cierre.columns):
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            if "Glosa" in cierre.columns:
                fig.add_trace(go.Bar(x=cierre["Periodo"], y=cierre["Glosa"], name="Glosa", marker_color=NARANJA), secondary_y=False)
            if "% recaudo" in cierre.columns:
                fig.add_trace(go.Scatter(x=cierre["Periodo"], y=cierre["% recaudo"], name="% recaudo", mode="lines+markers",
                                         line=dict(color=AZUL_OSCURO, width=3)), secondary_y=True)
            estilo_figura(fig, "Glosa y recaudo por periodo", eje_y_moneda=True)
            fig.update_yaxes(ticksuffix=" %", secondary_y=True)
            mostrar_grafico(fig)
        else:
            st.info("Se requieren Glosa_Inicial o Valor_Pagado para este gráfico.")


def mostrar_cierre(ctx):
    """Página 4: cierre mensual, trimestral o anual."""
    st.header("4. Cierre mensual")
    df = ctx["df"]
    periodicidad = st.radio("Periodicidad", ["Mensual", "Trimestral", "Anual"], horizontal=True, key="cierre_periodicidad")
    cierre = construir_cierre_mensual(df, periodicidad)
    if cierre.empty:
        st.warning("No fue posible construir el cierre: no hay periodos válidos con los filtros actuales.")
        return
    destacados = []
    for columna, texto, funcion in [("Facturación neta", "Mayor facturación", "idxmax"),
                                    ("Facturación neta", "Menor facturación", "idxmin"),
                                    ("Pagado", "Mayor recaudo", "idxmax"), ("Saldo", "Mayor cartera", "idxmax"),
                                    ("Glosa", "Mayor glosa", "idxmax")]:
        if columna in cierre.columns and cierre[columna].notna().any():
            fila = cierre.loc[getattr(cierre[columna], funcion)()]
            destacados.append((texto, fila["Periodo"], fila[columna]))
    columnas = st.columns(3) + st.columns(3)
    for col, (texto, periodo, valor) in zip(columnas, destacados):
        col.metric(f"{texto}: {periodo}", formatear_moneda_compacta(valor), help=f"{texto} en {periodo}: {formatear_moneda_compacta(valor)}")
    columnas = columnas[:len(destacados) + 1]
    nota_variacion = ""
    if periodicidad == "Mensual":
        variacion = variacion_comparable(cierre)
        if variacion and not es_invalido(variacion["variacion"]):
            columnas[-1].metric(f"Variación {variacion['ultimo']} vs {variacion['anterior']}", formatear_porcentaje(variacion["variacion"]))
            nota_variacion = variacion["nota"]
        else:
            columnas[-1].metric("Variación del último periodo", "No comparable", help="No existe un periodo anterior consecutivo.")
    elif len(cierre) >= 2 and not es_invalido(cierre.iloc[-1]["Crecimiento vs periodo anterior"]):
        columnas[-1].metric(f"Variación {cierre.iloc[-1]['Periodo']} vs {cierre.iloc[-2]['Periodo']}",
                            formatear_porcentaje(cierre.iloc[-1]["Crecimiento vs periodo anterior"]),
                            help="El último periodo podría estar incompleto.")
    else:
        columnas[-1].metric("Variación del último periodo", "No comparable", help="Solo existe un periodo.")
    if nota_variacion:
        st.info(nota_variacion)
    st.subheader("Tabla de cierre")
    mostrar_tabla_auto(cierre)
    st.caption("El promedio móvil usa tres periodos completos de la tabla; el acumulado anual se reinicia en cada año.")
    pestanas = st.tabs(["Facturación y crecimiento", "Recaudo y cartera", "Comparación entre años", "Promedio móvil"])
    with pestanas[0], bloque_seguro("gráfico combinado"):
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=cierre["Periodo"], y=cierre["Facturación neta"], name="Facturación neta", marker_color=AZUL_MEDIO), secondary_y=False)
        fig.add_trace(go.Scatter(x=cierre["Periodo"], y=cierre["Crecimiento vs periodo anterior"], name="Crecimiento %", mode="lines+markers",
                                 line=dict(color=NARANJA, width=3)), secondary_y=True)
        estilo_figura(fig, "Facturación neta y crecimiento frente al periodo anterior", eje_y_moneda=True)
        fig.update_yaxes(ticksuffix=" %", secondary_y=True)
        mostrar_grafico(fig)
        fig = px.line(cierre, x="Periodo", y="Facturación neta", markers=True, color_discrete_sequence=[AZUL_OSCURO])
        mostrar_grafico(estilo_figura(fig, "Línea de facturación neta", eje_y_moneda=True))
    with pestanas[1], bloque_seguro("recaudo y cartera"):
        c1, c2 = st.columns(2)
        if "Pagado" in cierre.columns:
            fig = px.line(cierre, x="Periodo", y="Pagado", markers=True, color_discrete_sequence=[VERDE_AZULADO])
            c1.plotly_chart(estilo_figura(fig, "Línea de recaudo", eje_y_moneda=True), config={"displaylogo": False})
        else:
            c1.info("No se encontró Valor_Pagado.")
        if "Saldo" in cierre.columns:
            fig = px.line(cierre, x="Periodo", y="Saldo", markers=True, color_discrete_sequence=[ROJO])
            c2.plotly_chart(estilo_figura(fig, "Línea de cartera", eje_y_moneda=True), config={"displaylogo": False})
        else:
            c2.info("No se encontró Saldo_Cartera.")
    with pestanas[2], bloque_seguro("comparación entre años"):
        mensual = construir_cierre_mensual(df, "Mensual")
        mensual["Año"] = mensual["Periodo"].str.slice(0, 4)
        mensual["Mes"] = pd.Categorical(mensual["Periodo"].str.slice(5, 7).astype(int).map(MESES_ES), categories=list(MESES_ES.values()), ordered=True)
        if mensual["Año"].nunique() >= 2:
            fig = px.line(mensual.sort_values("Mes"), x="Mes", y="Facturación neta", color="Año", markers=True,
                          category_orders={"Mes": list(MESES_ES.values())})
            mostrar_grafico(estilo_figura(fig, "Facturación neta mensual por año", eje_y_moneda=True))
        else:
            st.info("Se requiere información de al menos dos años para esta comparación.")
    with pestanas[3], bloque_seguro("promedio móvil"):
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=cierre["Periodo"], y=cierre["Facturación neta"], name="Facturación neta", mode="lines+markers", line=dict(color=AZUL_MEDIO)))
        fig.add_trace(go.Scatter(x=cierre["Periodo"], y=cierre["Promedio móvil 3 periodos"], name="Promedio móvil (3)", mode="lines",
                                 line=dict(color=NARANJA, width=3, dash="dash")))
        mostrar_grafico(estilo_figura(fig, "Facturación neta y promedio móvil de tres periodos", eje_y_moneda=True))


def mostrar_pareto(df):
    """Análisis de concentración y Pareto para facturación, cartera, glosas y servicios."""
    st.caption("El 80 % se usa como referencia de análisis de Pareto, no como regla obligatoria. "
               "HHI: índice de Herfindahl-Hirschman (suma de participaciones al cuadrado × 10.000). CR3/CR5: participación de los 3/5 principales.")
    opciones = [("Facturación por entidad", "Entidad_Pagadora", "Valor_Neto"), ("Cartera por entidad", "Entidad_Pagadora", "Saldo_Cartera"),
                ("Glosas por entidad", "Entidad_Pagadora", "Glosa_Inicial"), ("Facturación por servicio", "Servicio", "Valor_Neto")]
    disponibles = [o for o in opciones if tiene(df, o[1], o[2])]
    if not disponibles:
        st.info("No hay columnas suficientes para el análisis de Pareto.")
        return
    eleccion = st.radio("Concentración a analizar", [o[0] for o in disponibles], horizontal=True, key="pareto_tipo")
    nombre, categoria, valor = next(o for o in disponibles if o[0] == eleccion)
    resultado = calcular_pareto(df, categoria, valor)
    if not resultado:
        st.info("No hay valores positivos para este análisis.")
        return
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Categorías hasta ~80 %", f"{resultado['n_80']} de {resultado['total_categorias']}")
    c2.metric("Participación top 3 (CR3)", formatear_porcentaje(resultado["cr3"]))
    c3.metric("Participación top 5 (CR5)", formatear_porcentaje(resultado["cr5"]))
    c4.metric("Índice HHI", formatear_numero(resultado["hhi"]))
    with bloque_seguro("gráfico de Pareto"):
        mostrar_grafico(grafico_pareto(resultado, categoria, f"Pareto: {nombre.lower()}"))
    tabla = resultado["tabla"].rename(columns={"Valor": nombre})
    mostrar_tabla(tabla, monedas=[nombre], porcentajes=["Participación", "Participación acumulada"])
    st.caption(f"Suma de participaciones: {formatear_porcentaje(resultado['tabla']['Participación'].sum(), 2)} (puede diferir de 100 % por redondeo).")


def mostrar_entidades(ctx):
    """Página 5: entidades pagadoras, rankings, dispersión, mapa de calor, detalle y Pareto."""
    st.header("5. Entidades pagadoras")
    df = ctx["df"]
    if not requiere(df, ["Entidad_Pagadora"], "análisis por entidad"):
        return
    tabla = analizar_entidades(df)
    pestanas = st.tabs(["Tabla comparativa", "Rankings", "Dispersión y mapa de calor", "Detalle de una entidad", "Concentración y Pareto"])
    with pestanas[0]:
        mostrar_tabla_auto(tabla)
    with pestanas[1], bloque_seguro("rankings de entidades"):
        metricas = [m for m in ["Facturación neta", "Saldo", "Glosa", "% recaudo"] if m in tabla.columns]
        nombres = {"Facturación neta": "Ranking por facturación", "Saldo": "Ranking por cartera", "Glosa": "Ranking por glosa", "% recaudo": "Ranking por recaudo"}
        metrica = st.radio("Ranking", metricas, format_func=lambda m: nombres[m], horizontal=True, key="ent_ranking")
        top = st.select_slider("Entidades a mostrar", options=[5, 10, 15, 20, "Todas"], value=10, key="ent_top")
        datos = tabla.sort_values(metrica, ascending=False)
        datos = datos if top == "Todas" else datos.head(top)
        fig = px.bar(datos.sort_values(metrica), x=metrica, y="Entidad_Pagadora", orientation="h", color_discrete_sequence=[AZUL_MEDIO],
                     hover_data={"Facturas": True})
        estilo_figura(fig, nombres[metrica], eje_x_moneda=not metrica.startswith("%"), altura=max(380, 28 * len(datos) + 120))
        if metrica.startswith("%"):
            fig.update_xaxes(ticksuffix=" %")
        mostrar_grafico(fig)
    with pestanas[2]:
        with bloque_seguro("dispersión de entidades"):
            if "% recaudo" in tabla.columns:
                datos = tabla.copy()
                datos["Tamaño"] = datos["Saldo"].clip(lower=0).fillna(0) if "Saldo" in datos.columns else datos["Facturas"]
                fig = px.scatter(datos, x="Facturación neta", y="% recaudo", size="Tamaño", size_max=55,
                                 color="% glosa" if "% glosa" in datos.columns else None, color_continuous_scale="Oranges",
                                 hover_name="Entidad_Pagadora", hover_data={"Facturas": True, "Ticket promedio": ":,.0f", "Tamaño": False})
                estilo_figura(fig, "Facturación neta vs % de recaudo (tamaño: saldo de cartera; color: % de glosa)", eje_x_moneda=True)
                fig.update_yaxes(ticksuffix=" %")
                mostrar_grafico(fig)
            else:
                st.info("Se requiere Valor_Pagado para la dispersión de recaudo.")
        with bloque_seguro("mapa de calor entidad-periodo"):
            principales = tabla.head(15)["Entidad_Pagadora"].tolist()
            pivote = df[df["Entidad_Pagadora"].isin(principales)].pivot_table(index="Entidad_Pagadora", columns="Periodo_Cierre",
                                                                               values="Valor_Neto", aggfunc="sum", fill_value=0)
            fig = px.imshow(pivote, aspect="auto", color_continuous_scale=ESCALA_SECUENCIAL, labels=dict(color="Facturación neta"))
            mostrar_grafico(estilo_figura(fig, "Facturación neta por entidad y periodo (15 principales)", altura=max(400, 30 * len(pivote) + 150)))
    with pestanas[3], bloque_seguro("detalle de entidad"):
        entidad = st.selectbox("Seleccione una entidad", tabla["Entidad_Pagadora"].tolist(), key="ent_detalle")
        sub = df[df["Entidad_Pagadora"] == entidad]
        k = calcular_kpis(sub)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Facturación neta", formatear_moneda_compacta(k["neto"]))
        c2.metric("Facturas", formatear_numero(k["facturas"]))
        c3.metric("% recaudo", formatear_porcentaje(k["pct_recaudo"]))
        c4.metric("% glosa", formatear_porcentaje(k["pct_glosa"]))
        cierre = construir_cierre_mensual(sub, "Mensual")
        if not cierre.empty:
            series = [c for c in ["Facturación neta", "Pagado", "Saldo"] if c in cierre.columns]
            largo = cierre.melt(id_vars="Periodo", value_vars=series, var_name="Concepto", value_name="Valor")
            fig = px.line(largo, x="Periodo", y="Valor", color="Concepto", markers=True)
            mostrar_grafico(estilo_figura(fig, f"Evolución mensual: {entidad}", eje_y_moneda=True))
        if tiene(sub, "Servicio"):
            mostrar_tabla_auto(analizar_servicios(sub, "Servicio").head(15))
    with pestanas[4]:
        mostrar_pareto(df)


def selector_top(clave, etiqueta="Categorías a mostrar"):
    """Selector Top 5 / 10 / 15 / Todos."""
    eleccion = st.radio(etiqueta, ["Top 5", "Top 10", "Top 15", "Todos"], index=1, horizontal=True, key=clave)
    return None if eleccion == "Todos" else int(eleccion.split()[1])


def mostrar_servicios(ctx):
    """Página 6: servicios y líneas de servicio."""
    st.header("6. Servicios y líneas")
    df = ctx["df"]
    disponibles = [c for c in ["Servicio", "Linea_Servicio"] if tiene(df, c)]
    if not disponibles:
        st.warning("No se encontraron las columnas Servicio ni Linea_Servicio. Esta sección se desactivó.")
        return
    nivel = st.radio("Nivel de análisis", disponibles, format_func=lambda c: "Servicio" if c == "Servicio" else "Línea de servicio",
                     horizontal=True, key="serv_nivel")
    top = selector_top("serv_top")
    tabla = analizar_servicios(df, nivel)
    datos = tabla if top is None else tabla.head(top)
    if top is None and len(tabla) > 30:
        st.caption("Hay muchas categorías: los gráficos muestran las 30 principales para mantener la legibilidad.")
        datos = tabla.head(30)
    pestanas = st.tabs(["Tabla", "Barras y treemap", "Distribución (boxplot)", "Burbujas", "Servicio vs entidad", "Evolución"])
    with pestanas[0]:
        mostrar_tabla_auto(tabla)
    with pestanas[1]:
        with bloque_seguro("barras de servicios"):
            fig = px.bar(datos.sort_values("Facturación neta"), x="Facturación neta", y=nivel, orientation="h", color_discrete_sequence=[AZUL_MEDIO],
                         hover_data={"Facturas": True, "Ticket promedio": ":,.0f"})
            mostrar_grafico(estilo_figura(fig, "Facturación neta", eje_x_moneda=True, altura=max(380, 26 * len(datos) + 120)))
        with bloque_seguro("treemap"):
            if tiene(df, "Linea_Servicio", "Servicio"):
                arbol = df.groupby(["Linea_Servicio", "Servicio"], observed=True)["Valor_Neto"].sum().reset_index()
                arbol = arbol[arbol["Valor_Neto"] > 0]
                fig = px.treemap(arbol, path=[px.Constant("Total"), "Linea_Servicio", "Servicio"], values="Valor_Neto",
                                 color_discrete_sequence=PALETA)
                fig.update_traces(hovertemplate="%{label}<br>Facturación neta: %{value:,.0f}<extra></extra>")
                mostrar_grafico(estilo_figura(fig, "Treemap de facturación neta por línea y servicio", altura=520))
    with pestanas[2], bloque_seguro("boxplot de valor neto"):
        categorias = datos[nivel].tolist()[:15]
        sub = df[df[nivel].isin(categorias)]
        fig = px.box(sub, x=nivel, y="Valor_Neto", color_discrete_sequence=[VERDE_AZULADO], points=False,
                     category_orders={nivel: categorias})
        mostrar_grafico(estilo_figura(fig, "Distribución del valor neto por categoría (15 principales)", eje_y_moneda=True, altura=480))
    with pestanas[3], bloque_seguro("burbujas"):
        fig = px.scatter(datos, x="Facturas", y="Ticket promedio", size=datos["Facturación neta"].clip(lower=0), size_max=55,
                         color="% glosa" if "% glosa" in datos.columns else None, color_continuous_scale="Oranges", hover_name=nivel)
        estilo_figura(fig, "Facturas vs ticket promedio (tamaño: facturación; color: % glosa)", eje_y_moneda=True)
        mostrar_grafico(fig)
    with pestanas[4], bloque_seguro("mapa de calor servicio-entidad"):
        if tiene(df, "Entidad_Pagadora"):
            filas = top_categorias(df, nivel, "Valor_Neto", 15)
            columnas = top_categorias(df, "Entidad_Pagadora", "Valor_Neto", 12)
            pivote = df[df[nivel].isin(filas) & df["Entidad_Pagadora"].isin(columnas)].pivot_table(
                index=nivel, columns="Entidad_Pagadora", values="Valor_Neto", aggfunc="sum", fill_value=0)
            fig = px.imshow(pivote, aspect="auto", color_continuous_scale=ESCALA_SECUENCIAL, labels=dict(color="Facturación neta"))
            mostrar_grafico(estilo_figura(fig, "Facturación neta: categoría vs entidad (principales)", altura=max(420, 30 * len(pivote) + 160)))
        else:
            st.info("Se requiere Entidad_Pagadora para este mapa de calor.")
    with pestanas[5], bloque_seguro("evolución del servicio"):
        elegido = st.selectbox("Seleccione una categoría", tabla[nivel].tolist(), key="serv_evolucion")
        cierre = construir_cierre_mensual(df[df[nivel] == elegido], "Mensual")
        if not cierre.empty:
            fig = px.line(cierre, x="Periodo", y="Facturación neta", markers=True, color_discrete_sequence=[AZUL_MEDIO])
            mostrar_grafico(estilo_figura(fig, f"Evolución mensual: {elegido}", eje_y_moneda=True))


def mostrar_dimension(df, columna, nombre):
    """Análisis de una dimensión (sede o régimen) con tabla, barras, 100 %, boxplot y mapa de calor."""
    tabla = analizar_sedes(df, columna)
    mostrar_tabla_auto(tabla)
    c1, c2 = st.columns(2)
    with c1, bloque_seguro(f"barras agrupadas por {nombre}"):
        series = [c for c in ["Facturación neta", "Pagado", "Saldo", "Glosa"] if c in tabla.columns]
        largo = tabla.melt(id_vars=columna, value_vars=series, var_name="Concepto", value_name="Valor")
        fig = px.bar(largo, x=columna, y="Valor", color="Concepto", barmode="group")
        mostrar_grafico(estilo_figura(fig, f"Facturación, recaudo, cartera y glosa por {nombre}", eje_y_moneda=True))
    with c2, bloque_seguro(f"barras al 100 % por {nombre}"):
        if tiene(df, "Estado_Factura"):
            cruce = pd.crosstab(df[columna], df["Estado_Factura"], normalize="index") * 100
            largo = cruce.reset_index().melt(id_vars=columna, var_name="Estado", value_name="Porcentaje")
            fig = px.bar(largo, x=columna, y="Porcentaje", color="Estado", barmode="stack")
            estilo_figura(fig, f"Composición de estados por {nombre} (100 %)")
            fig.update_yaxes(ticksuffix=" %", range=[0, 100])
            mostrar_grafico(fig)
        else:
            st.info("Se requiere Estado_Factura para la composición al 100 %.")
    c1, c2 = st.columns(2)
    with c1, bloque_seguro(f"boxplot por {nombre}"):
        fig = px.box(df, x=columna, y="Valor_Neto", points=False, color_discrete_sequence=[VERDE_AZULADO])
        mostrar_grafico(estilo_figura(fig, f"Distribución del valor neto por {nombre}", eje_y_moneda=True))
    with c2, bloque_seguro(f"mapa de calor por {nombre}"):
        pivote = df.pivot_table(index=columna, columns="Periodo_Cierre", values="Valor_Neto", aggfunc="sum", fill_value=0)
        fig = px.imshow(pivote, aspect="auto", color_continuous_scale=ESCALA_SECUENCIAL, labels=dict(color="Facturación neta"))
        mostrar_grafico(estilo_figura(fig, f"Facturación neta por {nombre} y periodo"))


def mostrar_sedes_regimenes(ctx):
    """Página 7: sedes y regímenes (solo si las columnas existen)."""
    st.header("7. Sedes y regímenes")
    df = ctx["df"]
    disponibles = [(c, n) for c, n in [("Sede", "sede"), ("Regimen", "régimen")] if tiene(df, c)]
    for columna, nombre in [("Sede", "sede"), ("Regimen", "régimen")]:
        if not tiene(df, columna):
            st.warning(f"No se encontró la columna {columna}: el análisis por {nombre} está desactivado.")
    if not disponibles:
        return
    pestanas = st.tabs([n.capitalize() for _, n in disponibles])
    for pestana, (columna, nombre) in zip(pestanas, disponibles):
        with pestana:
            mostrar_dimension(df, columna, nombre)


def mostrar_estados(ctx):
    """Página 8: estados de facturación (categorías detectadas dinámicamente)."""
    st.header("8. Estados de facturación")
    df = ctx["df"]
    if not requiere(df, ["Estado_Factura"], "estados de facturación"):
        return
    tabla = analizar_estados(df)
    indicadores = {"pagada": "Pagadas", "parcialmente_pagada": "Parcialmente pagadas", "radicada": "Radicadas",
                   "aceptada": "Aceptadas", "glosada": "Glosadas", "devuelta": "Devueltas"}
    normalizados = tabla["Estado_Factura"].astype(str).map(normalizar_nombre_columna)
    encontrados = [(nombre, tabla.loc[normalizados == clave]) for clave, nombre in indicadores.items() if (normalizados == clave).any()]
    if encontrados:
        columnas = st.columns(len(encontrados))
        for col, (nombre, fila) in zip(columnas, encontrados):
            col.metric(nombre, formatear_numero(fila["Facturas"].sum()), f"{formatear_porcentaje(fila['% de facturas'].sum())} de las facturas", delta_color="off")
    otros = sorted(set(tabla["Estado_Factura"].astype(str)) - {fila["Estado_Factura"].iloc[0] for _, fila in encontrados})
    if otros:
        st.caption("Otros estados detectados en el archivo: " + ", ".join(otros))
    mostrar_tabla_auto(tabla[[c for c in ["Estado_Factura", "Facturas", "% de facturas", "Facturación neta", "Pagado", "Saldo", "Glosa"] if c in tabla.columns]])
    c1, c2 = st.columns(2)
    with c1, bloque_seguro("barras por estado"):
        fig = px.bar(tabla, x="Estado_Factura", y="Facturas", color="Estado_Factura", text="Facturas")
        estilo_figura(fig, "Facturas por estado")
        fig.update_layout(showlegend=False)
        mostrar_grafico(fig)
    with c2, bloque_seguro("valor por estado"):
        series = [c for c in ["Facturación neta", "Pagado", "Saldo"] if c in tabla.columns]
        largo = tabla.melt(id_vars="Estado_Factura", value_vars=series, var_name="Concepto", value_name="Valor")
        fig = px.bar(largo, x="Estado_Factura", y="Valor", color="Concepto", barmode="group")
        mostrar_grafico(estilo_figura(fig, "Valor neto, pagado y saldo por estado", eje_y_moneda=True))
    c1, c2 = st.columns(2)
    with c1, bloque_seguro("estados por periodo"):
        cruce = df.groupby(["Periodo_Cierre", "Estado_Factura"], observed=True).size().reset_index(name="Facturas")
        fig = px.bar(cruce, x="Periodo_Cierre", y="Facturas", color="Estado_Factura", barmode="stack")
        mostrar_grafico(estilo_figura(fig, "Estados por periodo (apilado)"))
    with c2, bloque_seguro("estados por periodo 100 %"):
        cruce100 = (pd.crosstab(df["Periodo_Cierre"], df["Estado_Factura"], normalize="index") * 100).reset_index().melt(
            id_vars="Periodo_Cierre", var_name="Estado", value_name="Porcentaje")
        fig = px.bar(cruce100, x="Periodo_Cierre", y="Porcentaje", color="Estado", barmode="stack")
        estilo_figura(fig, "Estados por periodo (100 %)")
        fig.update_yaxes(ticksuffix=" %", range=[0, 100])
        mostrar_grafico(fig)
    dimension = st.radio("Cruzar estado con", [c for c in ["Entidad_Pagadora", "Sede"] if tiene(df, c)] or ["Periodo_Cierre"],
                         horizontal=True, key="estado_cruce")
    with bloque_seguro("tabla cruzada de estados"):
        cruce = pd.crosstab(df[dimension], df["Estado_Factura"])
        fig = px.imshow(cruce, aspect="auto", color_continuous_scale=ESCALA_SECUENCIAL, text_auto=True, labels=dict(color="Facturas"))
        mostrar_grafico(estilo_figura(fig, f"Mapa de calor: estado por {dimension}", altura=max(400, 30 * len(cruce) + 150)))
        tabla_cruzada = cruce.reset_index()
        tabla_cruzada.columns = [str(c) for c in tabla_cruzada.columns]
        st.dataframe(tabla_cruzada, hide_index=True)


def mostrar_glosas(ctx):
    """Página 9: análisis de glosas (frecuencia y severidad)."""
    st.header("9. Glosas")
    df = ctx["df"]
    if not requiere(df, ["Glosa_Inicial"], "glosas"):
        return
    glosadas = df[df["Tiene_Glosa"]]
    k = calcular_kpis(df)
    frecuencia = division_segura(len(glosadas), len(df)) * 100
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Valor total glosado", formatear_moneda_compacta(k["glosa"]))
    c2.metric("Facturas con glosa", formatear_numero(glosadas["ID_Factura"].nunique() if "ID_Factura" in glosadas else len(glosadas)))
    c3.metric("Frecuencia de glosa", formatear_porcentaje(frecuencia), help="Facturas con glosa / total de facturas.")
    c4.metric("Severidad de glosa", formatear_porcentaje(k["pct_glosa"]), help="Valor glosado / valor neto.")
    c1, c2, c3 = st.columns(3)
    c1.metric("Glosa promedio (facturas glosadas)", formatear_moneda_compacta(glosadas["Glosa_Inicial"].mean() if len(glosadas) else np.nan))
    c2.metric("Glosa mediana", formatear_moneda_compacta(glosadas["Glosa_Inicial"].median() if len(glosadas) else np.nan))
    c3.metric("Glosa máxima", formatear_moneda_compacta(glosadas["Glosa_Inicial"].max() if len(glosadas) else np.nan))
    st.caption("La frecuencia mide cuántas facturas se glosan; la severidad mide qué proporción del valor neto fue glosada.")
    if glosadas.empty:
        st.info("No hay facturas con glosa en el conjunto filtrado.")
        return
    pestanas = st.tabs(["Tendencia", "Rankings", "Distribución", "Soporte completo", "Glosas más altas"])
    with pestanas[0], bloque_seguro("tendencia de glosa"):
        tendencia = df.groupby("Periodo_Cierre").agg(Glosa=("Glosa_Inicial", "sum"), Neto=("Valor_Neto", "sum"), Frecuencia=("Tiene_Glosa", "mean")).reset_index()
        tendencia["Severidad"] = division_serie(tendencia["Glosa"], tendencia["Neto"]) * 100
        tendencia["Frecuencia"] = tendencia["Frecuencia"] * 100
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=tendencia["Periodo_Cierre"], y=tendencia["Glosa"], name="Valor glosado", marker_color=NARANJA), secondary_y=False)
        fig.add_trace(go.Scatter(x=tendencia["Periodo_Cierre"], y=tendencia["Frecuencia"], name="Frecuencia %", mode="lines+markers", line=dict(color=AZUL_OSCURO)), secondary_y=True)
        fig.add_trace(go.Scatter(x=tendencia["Periodo_Cierre"], y=tendencia["Severidad"], name="Severidad %", mode="lines+markers", line=dict(color=ROJO, dash="dot")), secondary_y=True)
        estilo_figura(fig, "Tendencia mensual de glosas", eje_y_moneda=True)
        fig.update_yaxes(ticksuffix=" %", secondary_y=True)
        mostrar_grafico(fig)
    with pestanas[1]:
        for columna, titulo in [("Entidad_Pagadora", "entidad"), ("Servicio", "servicio"), ("Sede", "sede")]:
            if tiene(df, columna):
                with bloque_seguro(f"ranking de glosa por {titulo}"):
                    tabla = agregar_por(df, columna).sort_values("Glosa", ascending=False).head(15)
                    fig = px.bar(tabla.sort_values("Glosa"), x="Glosa", y=columna, orientation="h", color="% glosa",
                                 color_continuous_scale="Oranges", hover_data={"Facturas con glosa": True})
                    mostrar_grafico(estilo_figura(fig, f"Glosa por {titulo} (15 principales; color: % de glosa)", eje_x_moneda=True,
                                                  altura=max(380, 26 * len(tabla) + 120)))
    with pestanas[2]:
        c1, c2 = st.columns(2)
        with c1, bloque_seguro("boxplot de glosas"):
            fig = px.box(glosadas, y="Glosa_Inicial", x="Entidad_Pagadora" if tiene(glosadas, "Entidad_Pagadora") else None,
                         points=False, color_discrete_sequence=[NARANJA])
            mostrar_grafico(estilo_figura(fig, "Distribución del valor glosado", eje_y_moneda=True))
        with c2, bloque_seguro("histograma de glosas"):
            fig = px.histogram(glosadas, x="Glosa_Inicial", nbins=40, color_discrete_sequence=[NARANJA])
            mostrar_grafico(estilo_figura(fig, "Histograma del valor glosado (facturas glosadas)", eje_x_moneda=True))
    with pestanas[3], bloque_seguro("glosa según soporte"):
        if tiene(df, "Soporte_Completo"):
            comp = df.groupby("Soporte_Completo").agg(Facturas=("Valor_Neto", "size"), Frecuencia=("Tiene_Glosa", "mean"),
                                                      Glosa=("Glosa_Inicial", "sum"), Neto=("Valor_Neto", "sum")).reset_index()
            comp["Frecuencia de glosa"] = comp["Frecuencia"] * 100
            comp["Severidad de glosa"] = division_serie(comp["Glosa"], comp["Neto"]) * 100
            mostrar_tabla(comp[["Soporte_Completo", "Facturas", "Frecuencia de glosa", "Severidad de glosa", "Glosa"]],
                          enteros=["Facturas"], porcentajes=["Frecuencia de glosa", "Severidad de glosa"], monedas=["Glosa"])
            largo = comp.melt(id_vars="Soporte_Completo", value_vars=["Frecuencia de glosa", "Severidad de glosa"], var_name="Indicador", value_name="Porcentaje")
            fig = px.bar(largo, x="Soporte_Completo", y="Porcentaje", color="Indicador", barmode="group")
            estilo_figura(fig, "Frecuencia y severidad de glosa según soporte completo")
            fig.update_yaxes(ticksuffix=" %")
            mostrar_grafico(fig)
            st.info("Una asociación estadística no demuestra causalidad: otros factores pueden explicar las diferencias observadas.")
        else:
            st.info("No se encontró la columna Soporte_Completo.")
    with pestanas[4]:
        columnas = [c for c in ["ID_Factura", "Periodo_Cierre", "Entidad_Pagadora", "Servicio", "Sede", "Estado_Factura", "Soporte_Completo",
                                "Valor_Neto", "Glosa_Inicial", "Porcentaje_Glosa"] if c in df.columns]
        mostrar_tabla(glosadas.sort_values("Glosa_Inicial", ascending=False)[columnas].head(50),
                      monedas=["Valor_Neto", "Glosa_Inicial"], porcentajes=["Porcentaje_Glosa"])


def mostrar_cartera(ctx):
    """Página 10: recaudo y cartera, incluida la antigüedad aproximada."""
    st.header("10. Recaudo y cartera")
    df = ctx["df"]
    if not (tiene(df, "Valor_Pagado") or tiene(df, "Saldo_Cartera")):
        st.warning("No se encontraron las columnas Valor_Pagado ni Saldo_Cartera. Esta sección se desactivó.")
        return
    k = calcular_kpis(df)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Valor neto", formatear_moneda_compacta(k["neto"]))
    c2.metric("Valor pagado", formatear_moneda_compacta(k["pagado"]))
    c3.metric("Saldo", formatear_moneda_compacta(k["saldo"]))
    c4.metric("% recaudo", formatear_porcentaje(k["pct_recaudo"]))
    c5.metric("% cartera", formatear_porcentaje(k["pct_cartera"]))
    tolerancia = st.session_state.get("tolerancia", 1.0)
    if tiene(df, "Valor_Pagado"):
        pagado, neto = df["Valor_Pagado"].fillna(0), df["Valor_Neto"]
        total = int(((pagado >= neto - tolerancia) & (neto > 0)).sum())
        parcial = int(((pagado > 0) & (pagado < neto - tolerancia)).sum())
    else:
        total = parcial = None
    con_saldo = df[df["Tiene_Saldo"]] if "Tiene_Saldo" in df.columns else df.iloc[0:0]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Totalmente pagadas", formatear_numero(total) if total is not None else "No disponible", help="Valor pagado ≥ valor neto (con tolerancia).")
    c2.metric("Parcialmente pagadas", formatear_numero(parcial) if parcial is not None else "No disponible")
    c3.metric("Facturas con saldo", formatear_numero(len(con_saldo)) if "Tiene_Saldo" in df.columns else "No disponible")
    c4.metric("Saldo promedio", formatear_moneda_compacta(con_saldo["Saldo_Cartera"].mean() if len(con_saldo) else np.nan))
    c5.metric("Saldo mediano", formatear_moneda_compacta(con_saldo["Saldo_Cartera"].median() if len(con_saldo) else np.nan))
    if tiene(df, "Saldo_Cartera") and (df["Saldo_Cartera"] < 0).any():
        st.warning(f"Registros con saldo negativo: {formatear_numero((df['Saldo_Cartera'] < 0).sum())}. Conviene revisarlos en «Calidad de datos».")
    cierre = construir_cierre_mensual(df, "Mensual")
    pestanas = st.tabs(["Evolución", "Rankings de cartera", "Facturación vs recaudo", "Antigüedad", "Mayores saldos"])
    with pestanas[0], bloque_seguro("evolución de recaudo y cartera"):
        series = [c for c in ["Facturación neta", "Pagado", "Saldo"] if c in cierre.columns]
        largo = cierre.melt(id_vars="Periodo", value_vars=series, var_name="Concepto", value_name="Valor")
        fig = px.bar(largo, x="Periodo", y="Valor", color="Concepto", barmode="group")
        mostrar_grafico(estilo_figura(fig, "Comparación entre valor neto, pago y saldo por periodo", eje_y_moneda=True))
        lineas = [c for c in ["% recaudo", "% cartera"] if c in cierre.columns]
        largo = cierre.melt(id_vars="Periodo", value_vars=lineas, var_name="Indicador", value_name="Porcentaje")
        fig = px.line(largo, x="Periodo", y="Porcentaje", color="Indicador", markers=True)
        estilo_figura(fig, "Evolución del % de recaudo y del % de cartera")
        fig.update_yaxes(ticksuffix=" %")
        mostrar_grafico(fig)
    with pestanas[1]:
        if tiene(df, "Saldo_Cartera"):
            for columna, nombre in [("Entidad_Pagadora", "entidad"), ("Servicio", "servicio"), ("Sede", "sede")]:
                if tiene(df, columna):
                    with bloque_seguro(f"cartera por {nombre}"):
                        tabla = agregar_por(df, columna).sort_values("Saldo", ascending=False).head(15)
                        fig = px.bar(tabla.sort_values("Saldo"), x="Saldo", y=columna, orientation="h", color="% cartera",
                                     color_continuous_scale="Reds")
                        mostrar_grafico(estilo_figura(fig, f"Cartera por {nombre} (15 principales; color: % de cartera)", eje_x_moneda=True,
                                                      altura=max(380, 26 * len(tabla) + 120)))
        else:
            st.info("Se requiere Saldo_Cartera.")
    with pestanas[2], bloque_seguro("relación facturación-recaudo"):
        if tiene(df, "Entidad_Pagadora", "Valor_Pagado"):
            ent = analizar_entidades(df)
            mediana_neto, mediana_recaudo = ent["Facturación neta"].median(), ent["% recaudo"].median()
            ent["Grupo"] = np.where((ent["Facturación neta"] >= mediana_neto) & (ent["% recaudo"] < mediana_recaudo),
                                    "Alto volumen y bajo recaudo", "Otras entidades")
            fig = px.scatter(ent, x="Facturación neta", y="% recaudo", color="Grupo", symbol="Grupo", hover_name="Entidad_Pagadora",
                             color_discrete_map={"Alto volumen y bajo recaudo": ROJO, "Otras entidades": AZUL_MEDIO}, size_max=20)
            fig.add_vline(x=mediana_neto, line_dash="dot", line_color="#7F8C8D")
            fig.add_hline(y=mediana_recaudo, line_dash="dot", line_color="#7F8C8D")
            estilo_figura(fig, "Relación facturación-recaudo por entidad (líneas: medianas)", eje_x_moneda=True)
            fig.update_yaxes(ticksuffix=" %")
            mostrar_grafico(fig)
            seleccion = ent[ent["Grupo"] == "Alto volumen y bajo recaudo"]
            st.markdown("**Entidades con alto volumen y bajo recaudo** (facturación ≥ mediana y % de recaudo < mediana)")
            mostrar_tabla_auto(seleccion[["Entidad_Pagadora", "Facturación neta", "Pagado", "Saldo", "% recaudo"]] if "Saldo" in seleccion else seleccion)
        else:
            st.info("Se requieren Entidad_Pagadora y Valor_Pagado.")
    with pestanas[3], bloque_seguro("antigüedad de cartera"):
        if tiene(df, "Fecha_Radicacion", "Saldo_Cartera"):
            st.caption("Antigüedad aproximada desde la fecha de radicación hasta la fecha de corte. Los rangos son analíticos y configurables; "
                       "no constituyen una política contable o contractual.")
            c1, c2 = st.columns([1, 2])
            corte = c1.date_input("Fecha de corte", value=df["Fecha_Radicacion"].max().date(), key="cartera_corte")
            limites_texto = c2.text_input("Límites de los rangos (días, separados por coma)", value="30, 60, 90, 180", key="cartera_limites")
            try:
                limites = sorted({int(x) for x in re.findall(r"\d+", limites_texto)})
            except ValueError:
                limites = [30, 60, 90, 180]
            limites = limites or [30, 60, 90, 180]
            saldo = df[df["Tiene_Saldo"]].copy()
            saldo["Días"] = (pd.Timestamp(corte) - saldo["Fecha_Radicacion"]).dt.days
            etiquetas = [f"0 a {limites[0]} días"] + [f"{limites[i] + 1} a {limites[i + 1]} días" for i in range(len(limites) - 1)] + [f"Más de {limites[-1]} días"]
            saldo["Rango de antigüedad"] = pd.cut(saldo["Días"], bins=[-np.inf] + limites + [np.inf], labels=etiquetas)
            tabla = saldo.groupby("Rango de antigüedad", observed=False).agg(Facturas=("Saldo_Cartera", "size"), Saldo=("Saldo_Cartera", "sum")).reset_index()
            tabla["Participación"] = division_serie(tabla["Saldo"], pd.Series(tabla["Saldo"].sum(), index=tabla.index)) * 100
            futuras = int((saldo["Días"] < 0).sum())
            if futuras:
                st.warning(f"{futuras} facturas tienen fecha de radicación posterior a la fecha de corte y se incluyen en el primer rango.")
            mostrar_tabla(tabla.astype({"Rango de antigüedad": str}), enteros=["Facturas"], monedas=["Saldo"], porcentajes=["Participación"])
            fig = px.bar(tabla, x="Rango de antigüedad", y="Saldo", color_discrete_sequence=[ROJO], text=tabla["Participación"].map(lambda v: formatear_porcentaje(v)))
            mostrar_grafico(estilo_figura(fig, "Saldo de cartera por antigüedad aproximada", eje_y_moneda=True))
        else:
            st.info("Se requieren Fecha_Radicacion y Saldo_Cartera para calcular la antigüedad.")
    with pestanas[4]:
        if len(con_saldo):
            columnas = [c for c in ["ID_Factura", "Fecha_Radicacion", "Periodo_Cierre", "Entidad_Pagadora", "Servicio", "Estado_Factura",
                                    "Valor_Neto", "Valor_Pagado", "Saldo_Cartera"] if c in df.columns]
            mostrar_tabla(con_saldo.sort_values("Saldo_Cartera", ascending=False)[columnas].head(50),
                          monedas=["Valor_Neto", "Valor_Pagado", "Saldo_Cartera"], fechas=["Fecha_Radicacion"])
        else:
            st.info("No hay facturas con saldo en el conjunto filtrado.")


def estadisticos_tiempo(serie):
    """Estadísticos de tiempos de radicación."""
    datos = serie.dropna()
    return {"Promedio": datos.mean(), "Mediana": datos.median(), "Desviación estándar": datos.std(), "Mínimo": datos.min(),
            "Máximo": datos.max(), "Percentil 75": datos.quantile(0.75), "Percentil 90": datos.quantile(0.9), "Percentil 95": datos.quantile(0.95)}


def mostrar_tiempos(ctx, limites):
    """Página 11: tiempos entre emisión y radicación."""
    st.header("11. Tiempos de radicación")
    df = ctx["df"]
    if not tiene(df, "Dias_Emision_Radicacion"):
        st.warning("No se encontró Dias_Emision_Radicacion y no fue posible calcularlo (se requieren Fecha_Emision y Fecha_Radicacion).")
        return
    dias = df["Dias_Emision_Radicacion"]
    e = estadisticos_tiempo(dias)
    columnas = st.columns(4)
    for i, (nombre, valor) in enumerate(e.items()):
        columnas[i % 4].metric(nombre, f"{formatear_numero(valor, 1)} días")
    st.caption(f"Rangos analíticos configurables en la barra lateral: 0–{limites[0]}, {limites[0] + 1}–{limites[1]}, "
               f"{limites[1] + 1}–{limites[2]} y más de {limites[2]} días. No son estándares normativos.")
    if (dias < 0).any():
        st.warning(f"{formatear_numero((dias < 0).sum())} registros tienen días negativos (radicación antes de la emisión).")
    pestanas = st.tabs(["Distribución", "Por dimensión", "Rangos", "Tendencia", "Casos extremos"])
    with pestanas[0]:
        c1, c2 = st.columns(2)
        with c1, bloque_seguro("histograma de tiempos"):
            fig = px.histogram(df, x="Dias_Emision_Radicacion", nbins=40, color_discrete_sequence=[AZUL_MEDIO])
            fig.add_vline(x=e["Promedio"], line_dash="dash", line_color=NARANJA, annotation_text="Promedio")
            fig.add_vline(x=e["Mediana"], line_dash="dot", line_color=AZUL_OSCURO, annotation_text="Mediana", annotation_position="bottom right")
            mostrar_grafico(estilo_figura(fig, "Histograma de días entre emisión y radicación"))
        with c2, bloque_seguro("boxplot de tiempos"):
            fig = px.box(df, y="Dias_Emision_Radicacion", points="outliers", color_discrete_sequence=[VERDE_AZULADO])
            mostrar_grafico(estilo_figura(fig, "Boxplot de días de radicación"))
    with pestanas[1], bloque_seguro("tiempos por dimensión"):
        dimensiones = [c for c in ["Entidad_Pagadora", "Sede", "Canal_Radicacion", "Servicio", "Estado_Factura", "Soporte_Completo", "Periodo_Cierre"] if tiene(df, c)]
        dimension = st.selectbox("Analizar por", dimensiones, key="tiempo_dim")
        tabla = df.groupby(dimension, observed=True)["Dias_Emision_Radicacion"].agg(
            Registros="size", Promedio="mean", Mediana="median", Percentil_90=lambda s: s.quantile(0.9)).reset_index().sort_values("Promedio", ascending=False)
        tabla = tabla.rename(columns={"Percentil_90": "Percentil 90"})
        mostrar_tabla(tabla, enteros=["Registros"], decimales=["Promedio", "Mediana", "Percentil 90"])
        datos = tabla.head(25)
        fig = px.bar(datos, x=dimension, y=["Promedio", "Mediana"], barmode="group")
        mostrar_grafico(estilo_figura(fig, f"Días promedio y mediana por {dimension}"))
        categorias = datos[dimension].tolist()
        fig = px.box(df[df[dimension].isin(categorias)], x=dimension, y="Dias_Emision_Radicacion", points=False, color_discrete_sequence=[AZUL_MEDIO])
        mostrar_grafico(estilo_figura(fig, f"Distribución de días por {dimension}"))
    with pestanas[2], bloque_seguro("rangos de tiempo"):
        rangos = df["Rango_Tiempo_Radicacion"].value_counts(sort=False).reset_index()
        rangos.columns = ["Rango", "Registros"]
        rangos["Porcentaje"] = division_serie(rangos["Registros"], pd.Series(rangos["Registros"].sum(), index=rangos.index)) * 100
        fig = px.bar(rangos, x="Rango", y="Registros", text=rangos["Porcentaje"].map(lambda v: formatear_porcentaje(v)), color_discrete_sequence=[AZUL_MEDIO])
        mostrar_grafico(estilo_figura(fig, "Registros por rango de tiempo de radicación"))
    with pestanas[3], bloque_seguro("tendencia de tiempos"):
        tendencia = df.groupby("Periodo_Cierre")["Dias_Emision_Radicacion"].agg(["mean", "median"]).reset_index()
        tendencia.columns = ["Periodo", "Promedio", "Mediana"]
        fig = px.line(tendencia, x="Periodo", y=["Promedio", "Mediana"], markers=True)
        mostrar_grafico(estilo_figura(fig, "Tendencia mensual de días de radicación"))
    with pestanas[4]:
        columnas = [c for c in ["ID_Factura", "Fecha_Emision", "Fecha_Radicacion", "Dias_Emision_Radicacion", "Entidad_Pagadora", "Sede",
                                "Canal_Radicacion", "Servicio", "Estado_Factura", "Valor_Neto"] if c in df.columns]
        mostrar_tabla(df.sort_values("Dias_Emision_Radicacion", ascending=False)[columnas].head(30), monedas=["Valor_Neto"],
                      fechas=["Fecha_Emision", "Fecha_Radicacion"])
    if tiene(df, "Dias_Radicacion_Pago"):
        with st.expander("Días entre radicación y pago"):
            e2 = estadisticos_tiempo(df["Dias_Radicacion_Pago"])
            st.dataframe(pd.DataFrame({"Estadístico": list(e2.keys()), "Días": [formatear_numero(v, 1) for v in e2.values()]}), hide_index=True)


def variables_numericas(df):
    """Lista de variables numéricas disponibles para los análisis estadísticos."""
    preferidas = VARIABLES_CORRELACION + ["Dias_Radicacion_Pago", "Porcentaje_Glosa", "Porcentaje_Recaudo", "Porcentaje_Cartera",
                                          "Porcentaje_Descuento", "Porcentaje_Copago"]
    return [c for c in preferidas if c in df.columns and pd.api.types.is_numeric_dtype(df[c]) and df[c].notna().sum() >= 3]


def variables_categoricas(df):
    """Variables categóricas con un número manejable de categorías."""
    candidatas = COLUMNAS_CATEGORIA + ["Mes_Radicacion", "Trimestre_Radicacion", "Rango_Tiempo_Radicacion", "Año_Radicacion"]
    return [c for c in candidatas if c in df.columns and 2 <= df[c].nunique() <= 60]


def mostrar_estadistica(ctx):
    """Página 12: estadística descriptiva, correlaciones y pruebas estadísticas."""
    st.header("12. Análisis estadístico")
    df = ctx["df"]
    numericas = variables_numericas(df)
    if not numericas:
        st.warning("No hay variables numéricas con datos suficientes.")
        return
    pestanas = st.tabs(["Descriptiva", "Correlaciones", "Pruebas estadísticas"])
    with pestanas[0]:
        mostrar_descriptiva(df, numericas)
    with pestanas[1]:
        mostrar_correlaciones(df)
    with pestanas[2]:
        mostrar_pruebas(df, numericas)


def mostrar_descriptiva(df, numericas):
    """Estadística descriptiva de una variable con histograma, boxplot y curva observada."""
    variable = st.selectbox("Variable numérica", numericas, index=numericas.index("Valor_Neto") if "Valor_Neto" in numericas else 0, key="desc_var")
    e = calcular_estadisticos(df[variable])
    if e is None:
        st.warning("La variable seleccionada no contiene suficientes datos numéricos.")
        return
    es_moneda = variable in COLUMNAS_DINERO
    filas = []
    for nombre, valor in e.items():
        if nombre in ("Conteo", "Frecuencia de la moda"):
            texto = formatear_numero(valor)
        elif nombre in ("Asimetría", "Curtosis (exceso)"):
            texto = formatear_numero(valor, 3)
        elif nombre == "Coeficiente de variación (%)":
            texto = formatear_porcentaje(valor)
        elif nombre == "Moda" and es_invalido(valor):
            texto = "No aplica (ningún valor se repite)"
        elif nombre == "Varianza":
            texto = formatear_numero(valor, 2)
        else:
            texto = formatear_moneda(valor) if es_moneda else formatear_numero(valor, 2)
        filas.append({"Estadístico": nombre, "Valor": texto})
    c1, c2 = st.columns([1, 2])
    with c1:
        st.dataframe(pd.DataFrame(filas), hide_index=True, height=665)
    with c2:
        with bloque_seguro("histograma y curva observada"):
            datos = df[variable].dropna()
            fig = go.Figure()
            fig.add_trace(go.Histogram(x=datos, histnorm="probability density", name="Histograma", marker_color=AZUL_MEDIO, opacity=0.75, nbinsx=40))
            if datos.nunique() > 5 and datos.std() > 0:
                muestra = datos.sample(min(len(datos), 5000), random_state=42)
                kde = stats.gaussian_kde(muestra)
                eje = np.linspace(datos.min(), datos.max(), 300)
                fig.add_trace(go.Scatter(x=eje, y=kde(eje), name="Curva de densidad observada (KDE)", line=dict(color=NARANJA, width=3)))
            estilo_figura(fig, f"Distribución de {variable}", eje_x_moneda=es_moneda)
            mostrar_grafico(fig)
        with bloque_seguro("boxplot"):
            fig = px.box(df, x=variable, points="outliers", color_discrete_sequence=[VERDE_AZULADO])
            mostrar_grafico(estilo_figura(fig, f"Boxplot de {variable}", eje_x_moneda=es_moneda, altura=260))
    st.markdown("#### Interpretación automática")
    for texto in interpretar_estadisticos(e, variable):
        st.markdown(f"- {texto}")


def mostrar_correlaciones(df):
    """Matriz de correlación, pares destacados y dispersión entre dos variables."""
    disponibles = [c for c in VARIABLES_CORRELACION if c in df.columns and pd.api.types.is_numeric_dtype(df[c])]
    c1, c2, c3 = st.columns(3)
    metodo = c1.radio("Método", ["Pearson", "Spearman"], horizontal=True, key="corr_metodo")
    fuerte = c2.slider("Umbral de correlación fuerte", 0.5, 0.95, 0.7, 0.05, key="corr_fuerte")
    debil = c3.slider("Umbral de correlación débil", 0.05, 0.5, 0.3, 0.05, key="corr_debil")
    matriz = calcular_correlaciones(df, tuple(disponibles), metodo)
    if matriz.empty:
        st.warning("No hay suficientes variables numéricas con variación para calcular correlaciones.")
        return
    st.warning(AVISO_CORRELACION)
    with bloque_seguro("mapa de calor de correlaciones"):
        fig = px.imshow(matriz, text_auto=".2f", color_continuous_scale=ESCALA_DIVERGENTE, zmin=-1, zmax=1, aspect="auto")
        mostrar_grafico(estilo_figura(fig, f"Matriz de correlación ({metodo})", altura=520))
    pares = matriz.where(np.triu(np.ones(matriz.shape, dtype=bool), k=1)).stack().reset_index()
    pares.columns = ["Variable 1", "Variable 2", "Coeficiente"]
    pares["Clasificación"] = np.select(
        [pares["Coeficiente"] >= fuerte, pares["Coeficiente"] <= -fuerte, pares["Coeficiente"].abs() < debil],
        ["Asociación positiva fuerte", "Asociación negativa fuerte", "Asociación débil"], default="Asociación moderada")
    pares = pares.reindex(pares["Coeficiente"].abs().sort_values(ascending=False).index)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Tabla de correlaciones**")
        st.dataframe(pares.assign(Coeficiente=pares["Coeficiente"].map(lambda v: formatear_numero(v, 3))), hide_index=True, height=360)
    with c2:
        for clase in ["Asociación positiva fuerte", "Asociación negativa fuerte"]:
            sub = pares[pares["Clasificación"] == clase]
            st.markdown(f"**{clase}:** " + ("; ".join(f"{a} – {b} ({formatear_numero(r, 2)})" for a, b, r in sub[["Variable 1", "Variable 2", "Coeficiente"]].itertuples(index=False)) or "ninguna"))
        st.markdown(f"**Asociaciones débiles (|r| < {formatear_numero(debil, 2)}):** {int((pares['Clasificación'] == 'Asociación débil').sum())} pares.")
    st.markdown("#### Dispersión entre dos variables")
    c1, c2 = st.columns(2)
    x = c1.selectbox("Variable X", list(matriz.columns), key="corr_x")
    y = c2.selectbox("Variable Y", list(matriz.columns), index=min(1, len(matriz.columns) - 1), key="corr_y")
    with bloque_seguro("dispersión"):
        datos = df[[x, y]].dropna()
        if len(datos) > 5000:
            datos = datos.sample(5000, random_state=42)
            st.caption("Se muestra una muestra reproducible de 5.000 puntos.")
        fig = px.scatter(datos, x=x, y=y, opacity=0.55, color_discrete_sequence=[AZUL_MEDIO])
        if x != y and len(datos) >= 3 and datos[x].std() > 0:
            pendiente, intercepto = np.polyfit(datos[x], datos[y], 1)
            eje = np.linspace(datos[x].min(), datos[x].max(), 50)
            fig.add_trace(go.Scatter(x=eje, y=pendiente * eje + intercepto, mode="lines", name="Tendencia lineal", line=dict(color=NARANJA, width=3)))
        mostrar_grafico(estilo_figura(fig, f"{y} frente a {x}"))


def mostrar_resultado_prueba(r, alfa, tamanos):
    """Muestra hipótesis, estadístico, p-valor, decisión e interpretación de una prueba."""
    rechaza = r["p"] < alfa
    c1, c2, c3 = st.columns(3)
    c1.metric("Estadístico", formatear_numero(r["estadistico"], 4))
    c2.metric("p-valor", formatear_numero(r["p"], 4) if r["p"] >= 0.0001 else "< 0,0001")
    c3.metric("Nivel de significancia (α)", formatear_numero(alfa, 2))
    st.markdown(f"**Prueba:** {r['prueba']}  \n**Hipótesis nula (H₀):** {r['h0']}  \n**Hipótesis alternativa (H₁):** {r['h1']}")
    st.markdown("**Tamaño de las muestras:** " + ", ".join(f"{k}: {formatear_numero(v)}" for k, v in tamanos.items()))
    if rechaza:
        st.success(f"Decisión estadística: se rechaza H₀ (p < α). Los datos muestran evidencia de diferencias o asociación al nivel α = {formatear_numero(alfa, 2)}.")
    else:
        st.info("Decisión estadística: no se rechaza H₀ (p ≥ α). No hay evidencia suficiente para afirmar diferencias o asociación; "
                "esto no demuestra que H₀ sea verdadera.")
    st.caption("Las pruebas describen asociaciones o diferencias estadísticas y no establecen causalidad.")


def mostrar_pruebas(df, numericas):
    """Pruebas estadísticas: dos grupos, más de dos grupos, chi-cuadrado y normalidad."""
    tipo = st.radio("Tipo de análisis", ["Variable numérica vs dos grupos", "Variable numérica vs más de dos grupos",
                                         "Dos variables categóricas", "Normalidad"], key="prueba_tipo")
    alfa = st.select_slider("Nivel de significancia (α)", options=[0.01, 0.05, 0.10], value=0.05, key="prueba_alfa")
    categoricas = variables_categoricas(df)
    try:
        if tipo == "Normalidad":
            variable = st.selectbox("Variable", numericas, key="norm_var")
            datos = df[variable].dropna()
            if len(datos) < 3:
                st.warning("La variable seleccionada no contiene suficientes datos numéricos.")
                return
            nota = ""
            if len(datos) > 5000:
                datos = datos.sample(5000, random_state=42)
                nota = "El conjunto es grande: se usó una muestra reproducible de 5.000 valores (semilla 42)."
            r = ejecutar_prueba_estadistica("shapiro", datos, alfa)
            if nota:
                st.caption(nota)
            mostrar_resultado_prueba(r, alfa, {variable: len(datos)})
            st.caption("Con muestras grandes, desviaciones pequeñas de la normalidad suelen resultar significativas.")
            return
        if not categoricas:
            st.warning("No existen variables categóricas con un número adecuado de grupos.")
            return
        if tipo == "Dos variables categóricas":
            c1, c2 = st.columns(2)
            a = c1.selectbox("Variable categórica 1", categoricas, key="chi_a")
            b = c2.selectbox("Variable categórica 2", [c for c in categoricas if c != a], key="chi_b")
            tabla = pd.crosstab(df[a], df[b])
            if tabla.shape[0] < 2 or tabla.shape[1] < 2:
                st.warning("No existen suficientes grupos para realizar la prueba.")
                return
            r = ejecutar_prueba_estadistica("chi2", tabla, alfa)
            st.dataframe(tabla.reset_index().rename(columns=str), hide_index=True)
            mostrar_resultado_prueba(r, alfa, {"Registros": int(tabla.to_numpy().sum())})
            st.markdown(f"Grados de libertad: {r['gl']} · V de Cramér (fuerza de la asociación, 0 a 1): {formatear_numero(r['cramer'], 3)}")
            if r["pct_esperadas_bajas"] > 20:
                st.warning(f"El {formatear_porcentaje(r['pct_esperadas_bajas'])} de las frecuencias esperadas es menor que 5: el resultado debe interpretarse con cautela.")
            return
        c1, c2 = st.columns(2)
        variable = c1.selectbox("Variable numérica", numericas, key="prueba_var")
        grupo = c2.selectbox("Agrupación", categoricas, key="prueba_grupo")
        conteos = df.dropna(subset=[variable]).groupby(grupo, observed=True).size()
        validos = conteos[conteos >= 3].sort_values(ascending=False)
        if tipo == "Variable numérica vs dos grupos":
            if len(validos) < 2:
                st.warning("No existen suficientes grupos para realizar la prueba (se requieren al menos 2 con 3 o más registros).")
                return
            c1, c2 = st.columns(2)
            g1 = c1.selectbox("Grupo 1", validos.index.tolist(), key="prueba_g1")
            g2 = c2.selectbox("Grupo 2", [g for g in validos.index if g != g1], key="prueba_g2")
            a = df.loc[df[grupo] == g1, variable].dropna()
            b = df.loc[df[grupo] == g2, variable].dropna()
            normales = all(stats.shapiro(x.sample(min(len(x), 5000), random_state=42))[1] >= 0.05 for x in (a, b))
            sugerida = "t de Student" if normales or min(len(a), len(b)) >= 30 else "Mann-Whitney U"
            st.caption(f"Revisión de supuestos: normalidad de ambos grupos (Shapiro-Wilk) {'razonable' if normales else 'no razonable'}; "
                       f"tamaños {len(a)} y {len(b)}. Sugerencia: **{sugerida}** (con muestras grandes la prueba t es robusta; Mann-Whitney es la alternativa no paramétrica).")
            prueba = st.radio("Prueba", ["t de Student", "Mann-Whitney U"], index=0 if sugerida == "t de Student" else 1, horizontal=True, key="prueba_2g")
            r = ejecutar_prueba_estadistica("t" if prueba == "t de Student" else "mw", (a, b), alfa)
            mostrar_resultado_prueba(r, alfa, {str(g1): len(a), str(g2): len(b)})
        else:
            if len(validos) < 3:
                st.warning("No existen suficientes grupos para realizar la prueba (se requieren al menos 3 con 3 o más registros).")
                return
            elegidos = validos.index.tolist()[:10]
            if len(validos) > 10:
                st.caption("Se usan los 10 grupos con más registros.")
            muestras = [df.loc[df[grupo] == g, variable].dropna() for g in elegidos]
            homogeneas = stats.levene(*muestras)[1] >= 0.05
            sugerida = "ANOVA de una vía" if homogeneas else "Kruskal-Wallis"
            st.caption(f"Revisión de supuestos: homogeneidad de varianzas (Levene) {'razonable' if homogeneas else 'no razonable'}. Sugerencia: **{sugerida}**.")
            prueba = st.radio("Prueba", ["ANOVA de una vía", "Kruskal-Wallis"], index=0 if homogeneas else 1, horizontal=True, key="prueba_kg")
            r = ejecutar_prueba_estadistica("anova" if prueba.startswith("ANOVA") else "kruskal", muestras, alfa)
            mostrar_resultado_prueba(r, alfa, {str(g): len(m) for g, m in zip(elegidos, muestras)})
    except Exception:  # noqa: BLE001
        st.warning("No fue posible ejecutar la prueba con los datos seleccionados.")


def mostrar_atipicos(ctx):
    """Página 13: detección de valores atípicos."""
    st.header("13. Valores atípicos")
    df = ctx["df"]
    numericas = variables_numericas(df)
    if not numericas:
        st.warning("No hay variables numéricas con datos suficientes.")
        return
    st.info("Un valor atípico no es necesariamente un error: puede representar una factura legítima de mayor complejidad o valor.")
    c1, c2, c3 = st.columns(3)
    variable = c1.selectbox("Variable", numericas, key="atip_var")
    metodo = c2.selectbox("Método", ["Rango intercuartílico (IQR)", "Puntaje Z", "Percentiles"], key="atip_metodo")
    ambitos = {"Global": None}
    for col, nombre in [("Servicio", "Por servicio"), ("Linea_Servicio", "Por línea de servicio"), ("Entidad_Pagadora", "Por entidad")]:
        if tiene(df, col):
            ambitos[nombre] = col
    ambito = c3.selectbox("Ámbito", list(ambitos.keys()), key="atip_ambito")
    if metodo == "Rango intercuartílico (IQR)":
        parametros = {"k": st.slider("Multiplicador del IQR", 0.5, 5.0, 1.5, 0.1, key="atip_k")}
        st.caption("Límite inferior = Q1 − k × IQR · Límite superior = Q3 + k × IQR.")
    elif metodo == "Puntaje Z":
        parametros = {"z": st.slider("Umbral de puntaje Z", 1.5, 5.0, 3.0, 0.1, key="atip_z")}
        st.caption("El puntaje Z es más apropiado para distribuciones aproximadamente simétricas; con asimetría marcada, prefiera IQR o percentiles.")
    else:
        p1, p2 = st.slider("Percentiles inferior y superior", 0.0, 100.0, (1.0, 99.0), 0.5, key="atip_p")
        parametros = {"p_inf": p1, "p_sup": p2}
    with bloque_seguro("detección de atípicos"):
        datos = df[df[variable].notna()]
        mascara, inferior, superior = detectar_atipicos(datos, variable, metodo, parametros, ambitos[ambito])
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Valores atípicos", formatear_numero(int(mascara.sum())))
        c2.metric("Porcentaje", formatear_porcentaje(mascara.mean() * 100 if len(mascara) else np.nan))
        es_moneda = variable in COLUMNAS_DINERO
        fmt = formatear_moneda if es_moneda else (lambda v: formatear_numero(v, 2))
        if ambitos[ambito] is None:
            c3.metric("Límite inferior", fmt(inferior.iloc[0]))
            c4.metric("Límite superior", fmt(superior.iloc[0]))
        else:
            c3.metric("Límite inferior", "Variable por grupo")
            c4.metric("Límite superior", "Variable por grupo")
        etiquetado = datos.assign(Clasificación=np.where(mascara, "Atípico", "Dentro de límites"))
        c1, c2 = st.columns(2)
        with c1:
            fig = px.box(etiquetado, y=variable, x=ambitos[ambito] if ambitos[ambito] else None, points="outliers", color_discrete_sequence=[AZUL_MEDIO])
            mostrar_grafico(estilo_figura(fig, f"Boxplot de {variable}", eje_y_moneda=es_moneda))
        with c2:
            fig = px.histogram(etiquetado, x=variable, color="Clasificación", nbins=50,
                               color_discrete_map={"Atípico": ROJO, "Dentro de límites": AZUL_MEDIO})
            if ambitos[ambito] is None:
                fig.add_vline(x=inferior.iloc[0], line_dash="dash", line_color=NARANJA)
                fig.add_vline(x=superior.iloc[0], line_dash="dash", line_color=NARANJA)
            mostrar_grafico(estilo_figura(fig, f"Histograma de {variable}", eje_x_moneda=es_moneda))
        eje_x = "Fecha_Radicacion" if tiene(etiquetado, "Fecha_Radicacion") else "Periodo_Cierre"
        muestra = etiquetado if len(etiquetado) <= 8000 else pd.concat([etiquetado[mascara], etiquetado[~mascara].sample(6000, random_state=42)])
        fig = px.scatter(muestra, x=eje_x, y=variable, color="Clasificación", symbol="Clasificación", opacity=0.6,
                         color_discrete_map={"Atípico": ROJO, "Dentro de límites": AZUL_MEDIO},
                         hover_data=[c for c in ["ID_Factura", "Entidad_Pagadora", "Servicio"] if c in etiquetado.columns])
        mostrar_grafico(estilo_figura(fig, f"Dispersión de {variable} por periodo", eje_y_moneda=es_moneda))
        columnas = [c for c in ["ID_Factura", "Periodo_Cierre", "Entidad_Pagadora", "Servicio", "Linea_Servicio", "Sede", variable] if c in datos.columns]
        registros = datos.loc[mascara, columnas].sort_values(variable, ascending=False)
        st.markdown(f"**Registros atípicos ({formatear_numero(len(registros))}; se muestran hasta 200)**")
        mostrar_tabla(registros.head(200), monedas=[variable] if es_moneda else [])


def mostrar_insights(ctx):
    """Página 14: insights automáticos y semáforos configurables."""
    st.header("14. Insights y alertas")
    df = ctx["df"]
    st.caption("Insights generados con reglas programadas sobre los datos filtrados (sin APIs ni modelos externos). "
               "Cada uno separa el hallazgo calculado, la interpretación y la recomendación para investigar.")
    st.subheader("🚦 Semáforos")
    st.caption("Umbrales demostrativos y editables en la barra lateral («Configuración de alertas»). No son normas, contratos ni políticas institucionales.")
    with bloque_seguro("semáforos"):
        alertas = generar_alertas(df, st.session_state["umbrales"])
        columnas = st.columns(3)
        iconos = {"Verde": "🟢", "Amarillo": "🟡", "Rojo": "🔴", "Sin dato": "⚪"}
        clases = {"Verde": "verde", "Amarillo": "amarillo", "Rojo": "rojo", "Sin dato": "gris"}
        for i, fila in alertas.iterrows():
            columnas[i % 3].markdown(
                f'<div class="semaforo semaforo-{clases[fila["Estado"]]}"><b>{iconos[fila["Estado"]]} {fila["Indicador"]}</b><br>'
                f'Valor: {fila["Valor"]} · Estado: <b>{fila["Estado"]}</b><br><small>{fila["Umbrales configurados"]}</small></div>',
                unsafe_allow_html=True)
    st.subheader("💡 Insights")
    with st.spinner("Generando insights..."):
        insights = generar_insights(df, ctx["validacion"], st.session_state["umbrales"])
    if not insights:
        st.info("Los datos disponibles no son suficientes para generar insights.")
        return
    categorias = sorted({i["Categoría"] for i in insights})
    elegidas = st.multiselect("Categorías", categorias, default=categorias, key="insights_cat")
    for ins in [i for i in insights if i["Categoría"] in elegidas]:
        st.markdown(
            f'<div class="tarjeta-insight"><h4>{ins["Categoría"]} · {ins["Hallazgo"]}</h4>'
            f'<span class="etq">Hallazgo calculado:</span> {ins["Valor observado"]} — {ins["Contexto"]}<br>'
            f'<span class="etq">Interpretación:</span> {ins["Posible implicación"]}<br>'
            f'<span class="etq">Recomendación para investigar:</span> {ins["Acción sugerida para revisión"]}</div>',
            unsafe_allow_html=True)


def mostrar_explorador(ctx):
    """Página 15: explorador de registros con búsqueda, columnas, orden y filtros rápidos."""
    st.header("15. Explorador de datos")
    df = ctx["df"]
    if df.empty:
        st.warning("Los filtros actuales no tienen registros.")
        return
    c1, c2 = st.columns([2, 1])
    busqueda = c1.text_input("Buscar texto", key="exp_buscar", placeholder="Ejemplo: número de factura, entidad o servicio")
    n_filas = c2.selectbox("Filas a mostrar", [50, 100, 250, 500, 1000], index=1, key="exp_filas")
    c1, c2, c3, c4, c5 = st.columns(5)
    solo_saldo = c1.checkbox("Solo con saldo", key="exp_saldo", disabled="Tiene_Saldo" not in df.columns)
    solo_glosa = c2.checkbox("Solo con glosa", key="exp_glosa", disabled="Tiene_Glosa" not in df.columns)
    solo_soporte = c3.checkbox("Soporte incompleto", key="exp_soporte", disabled="Soporte_Incompleto" not in df.columns)
    solo_incons = c4.checkbox("Posibles inconsistencias", key="exp_incons")
    solo_atip = c5.checkbox("Valores atípicos (IQR en Valor_Neto)", key="exp_atip")
    datos = df
    if solo_saldo:
        datos = datos[datos["Tiene_Saldo"]]
    if solo_glosa:
        datos = datos[datos["Tiene_Glosa"]]
    if solo_soporte:
        datos = datos[datos["Soporte_Incompleto"]]
    if solo_incons and ctx["validacion"] is not None:
        banderas = ctx["validacion"]["banderas"]
        indices = banderas.index[banderas.any(axis=1)] if not banderas.empty else []
        datos = datos[datos.index.isin(indices)]
    if solo_atip and tiene(datos, "Valor_Neto"):
        mascara, _, _ = detectar_atipicos(df, "Valor_Neto", "Rango intercuartílico (IQR)", {"k": 1.5})
        datos = datos[datos.index.isin(mascara[mascara].index)]
    if busqueda:
        texto = busqueda.strip().lower()
        columnas_texto = [c for c in datos.columns if datos[c].dtype == object]
        coincide = pd.Series(False, index=datos.index)
        for col in columnas_texto:
            coincide |= datos[col].astype(str).str.lower().str.contains(texto, regex=False, na=False)
        datos = datos[coincide]
    todas = list(df.columns)
    predeterminadas = [c for c in ["ID_Factura", "Fecha_Radicacion", "Periodo_Cierre", "Entidad_Pagadora", "Sede", "Servicio", "Estado_Factura",
                                   "Valor_Neto", "Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera"] if c in todas]
    completos = st.checkbox("Mostrar registros completos (todas las columnas)", key="exp_completos")
    columnas = todas if completos else st.multiselect("Columnas", todas, default=predeterminadas or todas[:10], key="exp_columnas")
    c1, c2 = st.columns(2)
    orden = c1.selectbox("Ordenar por", ["(sin orden)"] + columnas, key="exp_orden")
    ascendente = c2.radio("Sentido", ["Descendente", "Ascendente"], horizontal=True, key="exp_sentido") == "Ascendente"
    if orden != "(sin orden)":
        datos = datos.sort_values(orden, ascending=ascendente)
    st.caption(f"Registros que cumplen las condiciones: {formatear_numero(len(datos))}. Se muestran como máximo {n_filas}.")
    if not columnas:
        st.info("Seleccione al menos una columna.")
        return
    vista = datos[columnas].head(n_filas)
    mostrar_tabla(vista, altura=520, monedas=[c for c in columnas if c in COLUMNAS_DINERO],
                  porcentajes=[c for c in columnas if c.startswith("Porcentaje_")],
                  fechas=[c for c in columnas if pd.api.types.is_datetime64_any_dtype(vista[c])])


def mostrar_descargas(ctx):
    """Página 16: descargas generadas en memoria."""
    st.header("16. Descargas")
    df = ctx["df"]
    st.caption("Los archivos se generan en memoria a partir de los datos filtrados. Los CSV usan punto y coma y coma decimal (compatibles con Excel en español).")
    with st.spinner("Preparando archivos de descarga..."):
        try:
            exportable = df.copy()
            for col in exportable.columns:
                if isinstance(exportable[col].dtype, pd.CategoricalDtype):
                    exportable[col] = exportable[col].astype(str)
            cierre = construir_cierre_mensual(df, "Mensual")
            secciones = st.columns(3)
            with secciones[0]:
                st.markdown("**Datos y cierre**")
                boton_descarga("Datos filtrados (CSV)", convertir_dataframe_csv(exportable), "facturacion_filtrada.csv", "text/csv", "d1")
                boton_descarga("Datos filtrados (Excel)", convertir_dataframe_excel({"Facturacion_filtrada": exportable}), "facturacion_filtrada.xlsx", MIME_XLSX, "d2")
                if not cierre.empty:
                    boton_descarga("Cierre mensual (CSV)", convertir_dataframe_csv(cierre), "cierre_mensual.csv", "text/csv", "d3")
                    boton_descarga("Cierre mensual (Excel)", convertir_dataframe_excel({"Cierre_mensual": cierre}), "cierre_mensual.xlsx", MIME_XLSX, "d4")
            with secciones[1]:
                st.markdown("**Análisis por dimensión**")
                if tiene(df, "Entidad_Pagadora"):
                    boton_descarga("Análisis por entidades (Excel)", convertir_dataframe_excel({"Entidades": analizar_entidades(df)}), "analisis_entidades.xlsx", MIME_XLSX, "d5")
                if tiene(df, "Servicio") or tiene(df, "Linea_Servicio"):
                    hojas = {}
                    if tiene(df, "Servicio"):
                        hojas["Servicios"] = analizar_servicios(df, "Servicio")
                    if tiene(df, "Linea_Servicio"):
                        hojas["Lineas"] = analizar_servicios(df, "Linea_Servicio")
                    boton_descarga("Análisis por servicios (Excel)", convertir_dataframe_excel(hojas), "analisis_servicios.xlsx", MIME_XLSX, "d6")
                if tiene(df, "Sede") or tiene(df, "Regimen"):
                    hojas = {}
                    if tiene(df, "Sede"):
                        hojas["Sedes"] = analizar_sedes(df, "Sede")
                    if tiene(df, "Regimen"):
                        hojas["Regimenes"] = analizar_sedes(df, "Regimen")
                    boton_descarga("Análisis por sedes y regímenes (Excel)", convertir_dataframe_excel(hojas), "analisis_sedes.xlsx", MIME_XLSX, "d7")
            with secciones[2]:
                st.markdown("**Registros específicos e informes**")
                if "Tiene_Glosa" in exportable.columns:
                    boton_descarga("Registros con glosa (CSV)", convertir_dataframe_csv(exportable[exportable["Tiene_Glosa"]]), "registros_glosados.csv", "text/csv", "d8")
                if "Tiene_Saldo" in exportable.columns:
                    boton_descarga("Registros con cartera (CSV)", convertir_dataframe_csv(exportable[exportable["Tiene_Saldo"]]), "registros_cartera.csv", "text/csv", "d9")
                if tiene(df, "Valor_Neto"):
                    mascara, _, _ = detectar_atipicos(df, "Valor_Neto", "Rango intercuartílico (IQR)", {"k": 1.5})
                    boton_descarga("Registros atípicos en Valor_Neto (CSV)", convertir_dataframe_csv(exportable[mascara.reindex(exportable.index).fillna(False)]),
                                   "registros_atipicos.csv", "text/csv", "d10")
                if ctx["validacion"] is not None:
                    boton_descarga("Informe de calidad (Excel)", convertir_dataframe_excel(informe_calidad_hojas(ctx)), "informe_calidad.xlsx", MIME_XLSX, "d11")
                insights = generar_insights(df, ctx["validacion"], st.session_state["umbrales"])
                boton_descarga("Resumen de insights (TXT)", texto_insights(insights, df).encode("utf-8"), "insights_facturacion.txt", "text/plain", "d12")
        except Exception:  # noqa: BLE001
            st.warning("No fue posible preparar algunas descargas con los datos actuales.")


def informe_calidad_hojas(ctx):
    """Hojas del informe de calidad para exportar."""
    v = ctx["validacion"]
    resumen = pd.DataFrame([
        ("Filas", v["filas"]), ("Columnas del archivo", v["columnas_archivo"]), ("Filas duplicadas", v["filas_duplicadas"]),
        ("ID duplicados", v["id_duplicados"]), ("Radicados duplicados", v["radicados_duplicados"]), ("Valores netos en cero", v["neto_cero"]),
        ("Facturas sin entidad", v["sin_entidad"]), ("Facturas sin periodo", v["sin_periodo"]), ("Filas con inconsistencias", v["filas_inconsistentes"]),
    ] + [(f"Puntaje: {k}", val) for k, val in v["puntaje"].items()], columns=["Indicador", "Valor"])
    hojas = {"Resumen": resumen, "Nulos": v["nulos"], "Inconsistencias": pd.DataFrame(list(v["inconsistencias"].items()), columns=["Verificación", "Registros"])}
    if ctx["conversiones"]:
        hojas["Conversiones"] = pd.DataFrame(ctx["conversiones"])
    if ctx["exclusiones"]:
        hojas["Limpieza"] = pd.DataFrame(ctx["exclusiones"])
    hojas["Aviso"] = pd.DataFrame({"Aviso": [AVISO_CALIDAD, AVISO_EXPLORATORIO]})
    return hojas


def texto_insights(insights, df):
    """Texto plano con el resumen de insights."""
    lineas = [APP_TITULO.upper(), APP_SUBTITULO, f"Fecha de generación: {date.today().isoformat()}",
              f"Registros analizados (filtrados): {formatear_numero(len(df))}", "", AVISO_EXPLORATORIO, ""]
    if not insights:
        lineas.append("Los datos disponibles no son suficientes para generar insights.")
    for i, ins in enumerate(insights, 1):
        lineas += [f"{i}. [{ins['Categoría']}] {ins['Hallazgo']}", f"   Hallazgo calculado: {ins['Valor observado']}",
                   f"   Contexto: {ins['Contexto']}", f"   Interpretación: {ins['Posible implicación']}",
                   f"   Recomendación para investigar: {ins['Acción sugerida para revisión']}", ""]
    return "\n".join(lineas)


def mostrar_metodologia():
    """Página 17: metodología y transparencia."""
    st.header("17. Metodología")
    st.markdown(f'<div class="aviso-suave">{AVISO_EXPLORATORIO}</div>', unsafe_allow_html=True)
    st.markdown("""
#### Indicadores principales
| Indicador | Cálculo |
|---|---|
| Facturación bruta | Suma de `Valor_Bruto` |
| Facturación neta | Suma de `Valor_Neto` (en el archivo se espera `Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota`) |
| Número de facturas | Conteo de `ID_Factura` únicos (no se cuenta varias veces la misma factura) |
| Porcentaje de glosa | `Glosa_Inicial / Valor_Neto × 100` |
| Porcentaje de recaudo | `Valor_Pagado / Valor_Neto × 100` |
| Porcentaje de cartera | `Saldo_Cartera / Valor_Neto × 100` |
| Ticket promedio | `Valor_Neto / número de facturas` |
| Frecuencia de glosa | Facturas con glosa / total de facturas |
| Severidad de glosa | Valor glosado / valor neto |
| Crecimiento | (Periodo actual − periodo anterior) / periodo anterior × 100, solo entre periodos consecutivos del mismo nivel de agrupación |

Cuando un denominador es cero o no existe, el indicador se muestra como **«No calculable»**: no se generan infinitos.
""")
    st.markdown("#### Fórmulas")
    st.latex(r"\%\,\text{Glosa} = \frac{\text{Glosa\_Inicial}}{\text{Valor\_Neto}} \times 100")
    st.latex(r"\%\,\text{Recaudo} = \frac{\text{Valor\_Pagado}}{\text{Valor\_Neto}} \times 100")
    st.latex(r"\%\,\text{Cartera} = \frac{\text{Saldo\_Cartera}}{\text{Valor\_Neto}} \times 100")
    st.latex(r"\text{Ticket promedio} = \frac{\sum \text{Valor\_Neto}}{\text{número de facturas}}")
    st.markdown("""
#### Conceptos estadísticos
- **Correlación:** mide la intensidad de la asociación entre dos variables (−1 a 1). Pearson evalúa relaciones lineales; Spearman, relaciones monótonas basadas en rangos. **La correlación no implica causalidad**, y algunas relaciones provienen de las fórmulas contables (por ejemplo, el valor neto se deriva del bruto).
- **Valor atípico:** observación muy alejada del resto. No es necesariamente un error.
- **Método IQR:** IQR = Q3 − Q1. Límite inferior = Q1 − k × IQR; límite superior = Q3 + k × IQR (k = 1,5 por defecto y configurable).
- **Puntaje Z:** distancia a la media en desviaciones estándar; es más apropiado con distribuciones aproximadamente simétricas.
- **p-valor:** probabilidad de observar un resultado igual o más extremo si la hipótesis nula fuera cierta. Un p-valor menor que α lleva a rechazar H₀; un p-valor mayor **no demuestra** que H₀ sea verdadera.
- **Pruebas:** t de Student (Welch) o Mann-Whitney U para dos grupos; ANOVA o Kruskal-Wallis para más de dos grupos; chi-cuadrado de independencia para dos variables categóricas; Shapiro-Wilk para normalidad (con una muestra reproducible de 5.000 valores si el conjunto es grande).
- **Pareto y concentración:** participación individual y acumulada; el 80 % es una referencia de análisis, no una regla. HHI = Σ(participación²) × 10.000.

#### Calidad de datos
- El **puntaje de calidad** promedia completitud, unicidad, validez y consistencia (0–100). Es exploratorio y no reemplaza una auditoría de datos.
- Las conversiones de tipo se informan columna por columna; los datos no se modifican silenciosamente.
- Los valores monetarios con formato ambiguo (por ejemplo, 1.234 o 1,234) se interpretan como separador de miles y se advierte en el diagnóstico.

#### Presentación de cifras
- Los valores monetarios usan punto como separador de miles y coma decimal (por ejemplo, $ 1.234.567).
- En las tarjetas, los montos de un millón o más se abrevian con **«M» (millones)**; las tablas y descargas conservan el valor completo.
- La etiqueta de moneda es solo visual: no se aplican tasas de cambio.

#### Rangos y umbrales
- Los rangos de tiempo de radicación y de antigüedad de cartera son **parámetros analíticos configurables**, no estándares normativos, contables ni contractuales.
- Los umbrales de los semáforos son **demostrativos y editables**. La aplicación no afirma incumplimientos de normas, contratos o políticas.

#### Limitaciones
- Los resultados dependen de la calidad, la completitud y la estructura del archivo cargado.
- El último periodo puede estar incompleto si el archivo se generó a mitad de mes.
- Las asociaciones estadísticas no establecen causalidad.
- La herramienta no es un sistema oficial de auditoría, contabilidad ni historia clínica.
""")
    st.markdown(f'<div class="aviso-privacidad">🔒 {AVISO_PRIVACIDAD}</div>', unsafe_allow_html=True)


# =============================================================================
# 17. FLUJO PRINCIPAL
# =============================================================================

def construir_contexto(df_original, info, opciones, conv_fechas, conv_moneda, limites, tolerancia):
    """Ejecuta el flujo de preparación (mapeo, tipos, calidad, limpieza y variables derivadas)."""
    ctx = {"df_original": df_original, "info": info, "df_base": None, "df": None, "validacion": None,
           "conversiones": [], "advertencias": [], "exclusiones": [], "equivalencias": [], "ambiguas": {},
           "mapeo": {}, "mapeo_detectado": {}, "minimas_ok": False, "df_tipado": None}
    if df_original is None:
        return ctx
    mapeo_detectado, equivalencias, ambiguas = detectar_columnas(tuple(df_original.columns))
    ctx.update(mapeo_detectado=mapeo_detectado, equivalencias=equivalencias, ambiguas=ambiguas)
    mapeo = mapeo_vigente(list(df_original.columns), mapeo_detectado)
    ctx["mapeo"] = mapeo
    ctx["minimas_ok"] = columnas_minimas_ok(mapeo)
    if not ctx["minimas_ok"]:
        return ctx
    with st.spinner("Preparando y validando los datos..."):
        mapeado = aplicar_mapeo(df_original, tuple(sorted(mapeo.items())))
        tipado, conversiones, mascaras, advertencias = preparar_datos(mapeado, conv_fechas, conv_moneda)
        validacion = validar_datos(tipado, mascaras, float(tolerancia), tuple(df_original.columns))
        limpio, exclusiones = limpiar_datos(tipado, mascaras, tuple(sorted(opciones.items())))
        base = crear_variables_derivadas(limpio, limites) if not limpio.empty else limpio
    ctx.update(df_tipado=tipado, conversiones=conversiones, advertencias=advertencias, validacion=validacion,
               exclusiones=exclusiones, df_base=base)
    return ctx


def indicador_filtrado(ctx):
    """Indicador del número y porcentaje de registros visibles tras los filtros."""
    total, visibles = len(ctx["df_base"]), len(ctx["df"])
    porcentaje = division_segura(visibles, total) * 100
    st.sidebar.markdown(f"**Registros filtrados:** {formatear_numero(visibles)} de {formatear_numero(total)} "
                        f"({formatear_porcentaje(porcentaje)} visibles)")
    st.sidebar.progress(min(1.0, visibles / total) if total else 0.0)


def main():
    """Punto de entrada de la aplicación."""
    configurar_pagina()
    aplicar_estilos()
    inicializar_estado()
    df_original, info = cargar_archivo()
    pagina, opciones, conv_fechas, conv_moneda, limites, tolerancia = opciones_barra_lateral()
    mostrar_encabezado()
    try:
        ctx = construir_contexto(df_original, info, opciones, conv_fechas, conv_moneda, limites, tolerancia)
    except Exception:  # noqa: BLE001
        st.error("No fue posible procesar el archivo. Revise su estructura o la correspondencia de columnas.")
        ctx = construir_contexto(None, info, opciones, conv_fechas, conv_moneda, limites, tolerancia)

    if ctx["df_base"] is not None and not ctx["df_base"].empty:
        filtros = construir_filtros(ctx["df_base"])
        ctx["df"] = aplicar_filtros(ctx["df_base"], filtros)
        indicador_filtrado(ctx)

    paginas_sin_datos = {PAGINAS[0]: lambda: mostrar_inicio(ctx), PAGINAS[16]: mostrar_metodologia}
    if pagina in paginas_sin_datos:
        try:
            paginas_sin_datos[pagina]()
        except Exception:  # noqa: BLE001
            st.error("Ocurrió un problema al mostrar esta sección. Las demás secciones siguen disponibles.")
        return
    if ctx["df_original"] is None:
        st.info("📁 Cargue el archivo de facturación que desea analizar desde la barra lateral para habilitar esta sección.")
        return
    if not ctx["minimas_ok"]:
        mostrar_faltantes_minimas(ctx)
        return
    if pagina == PAGINAS[1]:
        mostrar_calidad(ctx)
        return
    if ctx["df_base"] is None or ctx["df_base"].empty:
        st.warning("No quedaron registros después de la limpieza. Revise las opciones de limpieza o use «Restaurar datos originales».")
        return
    if ctx["df"].empty:
        st.warning("Los filtros actuales no tienen registros. Use «Limpiar filtros» en la barra lateral para restablecerlos.")
        return
    secciones = {
        PAGINAS[2]: mostrar_resumen, PAGINAS[3]: mostrar_cierre, PAGINAS[4]: mostrar_entidades, PAGINAS[5]: mostrar_servicios,
        PAGINAS[6]: mostrar_sedes_regimenes, PAGINAS[7]: mostrar_estados, PAGINAS[8]: mostrar_glosas, PAGINAS[9]: mostrar_cartera,
        PAGINAS[10]: lambda c: mostrar_tiempos(c, limites), PAGINAS[11]: mostrar_estadistica, PAGINAS[12]: mostrar_atipicos,
        PAGINAS[13]: mostrar_insights, PAGINAS[14]: mostrar_explorador, PAGINAS[15]: mostrar_descargas,
    }
    try:
        secciones[pagina](ctx)
    except Exception:  # noqa: BLE001
        st.error("Ocurrió un problema al generar esta sección con los datos actuales. Las demás secciones siguen disponibles.")
    st.markdown(f'<p class="pie-pagina">{APP_TITULO} · Herramienta de análisis exploratorio. {AVISO_EXPLORATORIO}</p>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
