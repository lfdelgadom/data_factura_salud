# Analítica de Facturación en Salud

Aplicación web desarrollada en **Python y Streamlit** para explorar archivos de facturación del sector salud. Permite analizar facturación, radicación, glosas, recaudo y cartera mediante indicadores, filtros, gráficos interactivos y análisis estadísticos.

## Funcionalidades principales

- Carga de archivos **XLSX** y **CSV**.
- Selección de hojas de Excel y detección de la hoja `Facturacion`.
- Normalización y mapeo manual de columnas.
- Diagnóstico de calidad y consistencia de datos.
- Indicadores financieros y cierre mensual.
- Análisis por entidades pagadoras, servicios, sedes, regímenes y estados.
- Evaluación de glosas, recaudo, cartera y tiempos de radicación.
- Estadística descriptiva, correlaciones, pruebas estadísticas y valores atípicos.
- Insights automáticos basados en reglas.
- Descarga de resultados en CSV y Excel.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue

El proyecto es compatible con **Streamlit Community Cloud**. Los archivos `app.py`, `requirements.txt` y `README.md` deben ubicarse en la raíz del repositorio. Al desplegar, seleccione `app.py` como archivo principal.

## Privacidad y alcance

La aplicación procesa los archivos en memoria y no utiliza APIs externas ni bases de datos. Se recomienda cargar únicamente información anonimizada y evitar nombres de pacientes, documentos de identidad, historias clínicas, diagnósticos u otros datos personales o sensibles.

Los resultados son de carácter exploratorio y no reemplazan una auditoría financiera, contable, médica o contractual.
