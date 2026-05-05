"""
LP Training - Generador de Informes Mensuales
App web para procesar CSV de TrainingPeaks y rellenar la plantilla.
"""
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
import shutil
import io
import os
import tempfile

# Tipos que NO cuentan como entreno programado
EXCLUDE_FROM_ADHERENCE = {'Day Off', 'Note'}
RUN_TYPE = 'Run'

PLANTILLA_PATH = 'Plantilla_Informe_LP_Training.xlsx'

# ==========================================================
# LÓGICA DE PROCESAMIENTO (igual que el script CLI)
# ==========================================================

def cargar_csv(file_obj):
    df = pd.read_csv(file_obj)
    df['WorkoutDay'] = pd.to_datetime(df['WorkoutDay'])
    for col in ['DistanceInMeters', 'TimeTotalInHours', 'PlannedDuration',
                'PlannedDistanceInMeters', 'VelocityAverage']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


def detectar_mes(df):
    primer_dia = df['WorkoutDay'].min()
    return primer_dia.year, primer_dia.month


def calcular_adherencia(df):
    programados_mask = ~df['WorkoutType'].isin(EXCLUDE_FROM_ADHERENCE)
    programados = programados_mask.sum()
    completados = (programados_mask & (df['TimeTotalInHours'] > 0)).sum()
    return int(programados), int(completados)


def calcular_metricas_running(df):
    runs = df[df['WorkoutType'] == RUN_TYPE]
    km_totales = runs['DistanceInMeters'].sum() / 1000
    rodaje_max = runs['DistanceInMeters'].max() / 1000 if len(runs) else 0

    runs_con_datos = runs[runs['DistanceInMeters'] > 0]
    if len(runs_con_datos) > 0:
        tiempo_total_seg = runs_con_datos['TimeTotalInHours'].sum() * 3600
        dist_total_km = runs_con_datos['DistanceInMeters'].sum() / 1000
        if dist_total_km > 0:
            seg_por_km = tiempo_total_seg / dist_total_km
            mins = int(seg_por_km // 60)
            segs = int(seg_por_km % 60)
            ritmo_str = f"{mins:02d}:{segs:02d}"
        else:
            ritmo_str = "00:00"
    else:
        ritmo_str = "00:00"

    return round(km_totales, 1), round(rodaje_max, 1), ritmo_str


def calcular_horas_totales(df):
    total_horas = df['TimeTotalInHours'].sum()
    h = int(total_horas)
    m = int((total_horas - h) * 60)
    s = int(((total_horas - h) * 60 - m) * 60)
    return f"{h}:{m:02d}:{s:02d}"


def calcular_por_semana(df, year, month):
    primer_dia = datetime(year, month, 1)
    if month == 12:
        ultimo_dia = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        ultimo_dia = datetime(year, month + 1, 1) - timedelta(days=1)

    semanas = []
    inicio = primer_dia
    n = 1
    while inicio <= ultimo_dia and n <= 5:
        fin = min(inicio + timedelta(days=6), ultimo_dia)
        mask = (df['WorkoutDay'] >= inicio) & (df['WorkoutDay'] <= fin)
        df_sem = df[mask]

        km_sem = df_sem[df_sem['WorkoutType'] == RUN_TYPE]['DistanceInMeters'].sum() / 1000
        horas_sem = df_sem['TimeTotalInHours'].sum()
        h = int(horas_sem)
        m = int((horas_sem - h) * 60)
        s = int(((horas_sem - h) * 60 - m) * 60)
        horas_str = f"{h}:{m:02d}:{s:02d}"

        label = f"S {n} ({inicio.day} a {fin.day})"
        semanas.append({'label': label, 'km': round(km_sem, 1), 'horas': horas_str})

        inicio = fin + timedelta(days=1)
        n += 1

    while len(semanas) < 5:
        semanas.append({'label': f'S {len(semanas)+1} (-)', 'km': 0, 'horas': '0:00:00'})

    return semanas


def nombre_mes(month):
    nombres = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
               'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    return nombres[month - 1]


def rellenar_plantilla(plantilla_path, datos):
    """Carga plantilla, rellena, devuelve los bytes del Excel resultante."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        shutil.copy(plantilla_path, tmp.name)
        tmp_path = tmp.name

    wb = load_workbook(tmp_path)
    ws = wb['INPUTS']

    ws['C6'] = datos['nombre_cliente']
    ws['F6'] = datos['mes_nombre']
    ws['F7'] = datos['año']
    ws['F8'] = datos['mes_numero']
    ws['C11'] = datos['programados']
    ws['F11'] = datos['completados']
    ws['C12'] = datos['km_totales']
    ws['C13'] = datos['horas_totales']
    ws['F13'] = datos['rodaje_max']
    ws['C19'] = datos['ritmo_medio']

    for i, sem in enumerate(datos['semanas']):
        r = 23 + i
        ws[f'B{r}'] = sem['label']
        ws[f'C{r}'] = sem['km']
        ws[f'E{r}'] = sem['label']
        ws[f'F{r}'] = sem['horas']

    # Guardar a buffer para descarga
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    os.unlink(tmp_path)
    return buffer.getvalue()


# ==========================================================
# INTERFAZ STREAMLIT
# ==========================================================

st.set_page_config(
    page_title="LP Training - Generador de Informes",
    page_icon="🏃",
    layout="centered"
)

# Estilo personalizado con tu paleta amarillo/negro
st.markdown("""
<style>
    .main-header {
        background-color: #1A1A1A;
        color: white;
        padding: 1.5rem;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 2rem;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2rem;
        letter-spacing: 1px;
    }
    .main-header p {
        margin: 0.3rem 0 0 0;
        font-size: 0.9rem;
        opacity: 0.85;
        font-style: italic;
    }
    .stButton > button {
        background-color: #FFD600;
        color: #1A1A1A;
        font-weight: bold;
        border: 2px solid #F9A825;
        padding: 0.5rem 2rem;
    }
    .stButton > button:hover {
        background-color: #F9A825;
        color: #1A1A1A;
        border-color: #1A1A1A;
    }
    .metric-card {
        background-color: #FFF9C4;
        padding: 1rem;
        border-radius: 6px;
        border-left: 4px solid #FFD600;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🏃 MÉTODO LP TRAINING</h1>
    <p>Generador de Informes Mensuales · @LP.RTRAINING_</p>
</div>
""", unsafe_allow_html=True)

st.markdown("### Paso 1 · Datos del cliente")

nombre_cliente = st.text_input(
    "Nombre del cliente",
    placeholder="Ej: SERGIO ALONSO",
    help="Se mostrará tal cual en el informe"
).strip().upper()

st.markdown("### Paso 2 · CSV de TrainingPeaks")

archivo_csv = st.file_uploader(
    "Sube el archivo workouts.csv que descargaste de TrainingPeaks",
    type=['csv'],
    help="Vista de calendario en TrainingPeaks → Export → CSV"
)

st.markdown("### Paso 3 · Generar")

generar = st.button("🚀 Generar Excel del informe", disabled=(not nombre_cliente or not archivo_csv))

if generar:
    try:
        with st.spinner("Procesando datos..."):
            df = cargar_csv(archivo_csv)
            year, month = detectar_mes(df)
            mes_str = nombre_mes(month)

            programados, completados = calcular_adherencia(df)
            km_totales, rodaje_max, ritmo_medio = calcular_metricas_running(df)
            horas_totales = calcular_horas_totales(df)
            semanas = calcular_por_semana(df, year, month)

            datos = {
                'nombre_cliente': nombre_cliente,
                'mes_nombre': mes_str,
                'año': year,
                'mes_numero': month,
                'programados': programados,
                'completados': completados,
                'km_totales': km_totales,
                'horas_totales': horas_totales,
                'rodaje_max': rodaje_max,
                'ritmo_medio': ritmo_medio,
                'semanas': semanas,
            }

            excel_bytes = rellenar_plantilla(PLANTILLA_PATH, datos)

        # Mostrar resumen
        st.success(f"✅ Informe de **{nombre_cliente}** generado para **{mes_str} {year}**")

        st.markdown("#### Resumen calculado")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Adherencia", f"{int(completados/programados*100)}%",
                     f"{completados}/{programados}")
        with col2:
            st.metric("KM totales", f"{km_totales}")
        with col3:
            st.metric("Horas", horas_totales)
        with col4:
            st.metric("Rodaje máximo", f"{rodaje_max} km")

        st.markdown("**Por semana:**")
        df_semanas = pd.DataFrame([{
            'Semana': s['label'],
            'KM': s['km'],
            'Horas': s['horas']
        } for s in semanas])
        st.dataframe(df_semanas, hide_index=True, use_container_width=True)

        # Botón de descarga
        nombre_limpio = nombre_cliente.replace(' ', '_')
        nombre_fichero = f"Informe_{nombre_limpio}_{mes_str}_{year}.xlsx"

        st.download_button(
            label="📥 Descargar Excel",
            data=excel_bytes,
            file_name=nombre_fichero,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        with st.expander("¿Qué tienes que rellenar manualmente al abrir el Excel?"):
            st.markdown("""
            La hoja **INPUTS** del Excel ya tiene rellenado todo lo que se puede sacar del CSV.
            Antes de exportar a PDF, completa también:

            - **Bloque 1**: Distancia objetivo del cliente y fecha de inicio del plan
            - **Bloque 2**: Desnivel acumulado del mes (no viene en el CSV)
            - **Bloque 3**: Datos del mes anterior (para la comparación)
            - **Bloques 5 y 6**: Zonas de velocidad y FC (cópialas desde el Panel de Control de TrainingPeaks)
            - **Bloque 7**: Comentarios del entrenador y próximos pasos
            - **Bloque 8**: Histórico anual (actualiza con el km del mes actual)

            Después, exporta a PDF seleccionando solo las hojas «INFORME MENSUAL» y «RESUMEN ANUAL».
            """)

    except Exception as e:
        st.error(f"❌ Hubo un error al procesar el archivo:\n\n```\n{e}\n```")
        st.info("Comprueba que el CSV es el original de TrainingPeaks (sin haberlo abierto y guardado en Excel) y vuelve a intentarlo.")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#888; font-size:0.85rem;'>"
    "Lorenzo Pla · Entrenador de corredores online"
    "</p>",
    unsafe_allow_html=True
)
