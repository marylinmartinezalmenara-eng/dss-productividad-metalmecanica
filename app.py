import io
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='DSS | Productividad y mantenimiento', page_icon='🏭', layout='wide')
st.title('DSS | Productividad y mantenimiento')
st.caption('Prototipo experimental · Datos modelados · Escenarios hipotéticos · IoT simulado')

REQUIRED = {
    'INDICADORES': ['Indicador','Valor','Unidad'],
    'KPIS_MENSUALES': ['Periodo','Tanques equivalentes','HH totales','Productividad TE/HH'],
    'PARETO_CAUSAS': ['Causa','HH_perdidas'],
    'PERDIDAS_GRUPOS': ['Grupo DSS','HH_perdidas'],
    'DECISIONES_DSS': ['Prioridad','Problema','Herramienta','HH perdidas','Puntuación DSS','Acción recomendada','Responsable','Indicador de seguimiento'],
    'ESCENARIOS': ['Escenario','Productividad actual TE/HH','Productividad A TE/HH','Mejora A (%)','Productividad B potencial TE/HH','Mejora B potencial (%)','HH liberadas hipotéticas','Estado B'],
    'MONITOREO_EQUIPOS': ['ID equipo','Alertas_advertencia','Alertas_criticas','Prioridad de revisión'],
    'ALERTAS_IOT': ['ID equipo','Fecha y hora','Temperatura (°C)','Vibración (mm/s)','Estado DSS'],
    'VALIDACIONES': ['Prueba','Estado'],
    'LIMITACIONES': ['Aspecto','Limitación'],
}

@st.cache_data(show_spinner=False)
def load_excel(data: bytes):
    book = pd.ExcelFile(io.BytesIO(data))
    missing_sheets = set(REQUIRED) - set(book.sheet_names)
    if missing_sheets:
        raise ValueError('Faltan hojas: ' + ', '.join(sorted(missing_sheets)))
    result = {sheet: pd.read_excel(book, sheet_name=sheet) for sheet in REQUIRED}
    for sheet, cols in REQUIRED.items():
        missing = set(cols) - set(result[sheet].columns)
        if missing:
            raise ValueError(f'{sheet}: faltan columnas {sorted(missing)}')
    if not result['VALIDACIONES']['Estado'].astype(str).eq('APROBADA').all():
        raise ValueError('El checkpoint integrado contiene validaciones no aprobadas.')
    return result

def download_df(df, name):
    st.download_button('Descargar tabla CSV', df.to_csv(index=False).encode('utf-8-sig'), name, 'text/csv')

uploaded = st.file_uploader('Cargar DSS_CHECKPOINT_08_INTEGRADO.xlsx', type=['xlsx'])
if uploaded is None:
    st.info('Carga el Excel del Checkpoint 08 para visualizar el tablero.')
    st.stop()
try:
    data = load_excel(uploaded.getvalue())
except Exception as exc:
    st.error(f'No se pudo validar el archivo: {exc}')
    st.stop()

ind = data['INDICADORES']
ind_map = dict(zip(ind['Indicador'], ind['Valor']))

menu = st.sidebar.radio('Módulo', [
    'Resumen ejecutivo', 'Diagnóstico', 'Recomendaciones',
    'Escenarios', 'Monitoreo IoT', 'Trazabilidad y límites'
])
st.sidebar.warning('Uso académico. No usar las alertas simuladas para operar equipos reales.')

if menu == 'Resumen ejecutivo':
    st.subheader('Situación actual y decisiones prioritarias')
    c1,c2,c3,c4 = st.columns(4)
    c1.metric('Productividad actual', f"{float(ind_map['Productividad actual']):.6f} TE/HH")
    c2.metric('Horas-hombre', f"{float(ind_map['Horas-hombre trabajadas']):,.2f} HH")
    c3.metric('Horas extra', f"{float(ind_map['Horas extra']):,.2f} HH")
    c4.metric('Pérdidas registradas', f"{float(ind_map['HH perdidas registradas']):,.2f} HH")
    top = data['DECISIONES_DSS'].sort_values('Prioridad').iloc[0]
    st.markdown(f"**Primera prioridad:** {top['Problema']} · **Herramienta:** {top['Herramienta']}")
    st.write(top['Acción recomendada'])
    scenario = data['ESCENARIOS'].loc[lambda d: d['Escenario'].eq('Integrado 5S + SLP + TPM')].iloc[0]
    a,b = st.columns(2)
    a.metric('Mejora hipotética A', f"{scenario['Mejora A (%)']:.2f}%")
    b.metric('HH liberadas hipotéticas', f"{scenario['HH liberadas hipotéticas']:.2f} HH")
    st.caption('Escenario B: potencial de producción no validado con demanda, capacidad y mezcla.')

elif menu == 'Diagnóstico':
    st.subheader('Pérdidas de mano de obra')
    groups = data['PERDIDAS_GRUPOS'].sort_values('HH_perdidas', ascending=False)
    st.plotly_chart(px.bar(groups, x='HH_perdidas', y='Grupo DSS', orientation='h', title='HH perdidas por herramienta'), use_container_width=True)
    pareto = data['PARETO_CAUSAS'].sort_values('HH_perdidas', ascending=False)
    st.plotly_chart(px.bar(pareto, x='Causa', y='HH_perdidas', title='Pérdidas por causa'), use_container_width=True)
    st.dataframe(pareto, hide_index=True, use_container_width=True)
    download_df(pareto, 'diagnostico_dss.csv')
    st.subheader('Evolución mensual')
    monthly = data['KPIS_MENSUALES'].copy()
    monthly['Periodo'] = monthly['Periodo'].astype(str)
    st.plotly_chart(px.line(monthly, x='Periodo', y='Productividad TE/HH', markers=True), use_container_width=True)

elif menu == 'Recomendaciones':
    st.subheader('Motor de decisiones 5S · SLP · TPM')
    rec = data['DECISIONES_DSS'].sort_values('Prioridad')
    choices = ['Todas'] + sorted(rec['Herramienta'].dropna().unique().tolist())
    tool = st.selectbox('Filtrar por herramienta', choices)
    view = rec if tool == 'Todas' else rec[rec['Herramienta'].eq(tool)]
    st.dataframe(view, hide_index=True, use_container_width=True)
    if not view.empty:
        choice = st.selectbox('Consultar una causa', view['Problema'].tolist())
        selected = view.loc[view['Problema'].eq(choice)].iloc[0]
        st.markdown(f"**Acción propuesta:** {selected['Acción recomendada']}")
        st.write('Responsable:', selected['Responsable'])
        st.write('KPI:', selected['Indicador de seguimiento'])
        st.caption('Prioridad provisional; factibilidad y urgencia pendientes de juicio experto.')
    download_df(view, 'recomendaciones_dss.csv')

elif menu == 'Escenarios':
    st.subheader('Comparación de productividad')
    scenarios = data['ESCENARIOS'].copy()
    selected_names = st.multiselect('Escenarios', scenarios['Escenario'].tolist(), default=scenarios['Escenario'].tolist())
    view = scenarios[scenarios['Escenario'].isin(selected_names)]
    metric = st.radio('Enfoque', ['A: misma producción, menos HH', 'B: mismas HH, producción potencial'], horizontal=True)
    col = 'Mejora A (%)' if metric.startswith('A:') else 'Mejora B potencial (%)'
    if not view.empty:
        st.plotly_chart(px.bar(view, x='Escenario', y=col, title='Mejora porcentual hipotética'), use_container_width=True)
        st.dataframe(view[['Escenario','HH liberadas hipotéticas','Productividad actual TE/HH','Productividad A TE/HH','Mejora A (%)','Productividad B potencial TE/HH','Mejora B potencial (%)','Estado B']], hide_index=True, use_container_width=True)
        download_df(view, 'escenarios_dss.csv')
    st.warning('El escenario B es potencial teórico; no constituye producción adicional factible confirmada.')

elif menu == 'Monitoreo IoT':
    st.subheader('Cuatro máquinas de corte y doblado')
    machines = data['MONITOREO_EQUIPOS'].copy()
    st.dataframe(machines, hide_index=True, use_container_width=True)
    machine = st.selectbox('Equipo', sorted(machines['ID equipo'].astype(str).unique()))
    alerts = data['ALERTAS_IOT'].copy()
    alerts['Fecha y hora'] = pd.to_datetime(alerts['Fecha y hora'], errors='coerce')
    filtered = alerts.loc[alerts['ID equipo'].astype(str).eq(machine)].sort_values('Fecha y hora')
    st.metric('Lecturas con alerta', len(filtered))
    if not filtered.empty:
        fig = px.line(filtered, x='Fecha y hora', y=['Temperatura (°C)','Vibración (mm/s)'], title='Valores en lecturas con alerta (unidades diferentes)')
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(filtered, hide_index=True, use_container_width=True)
        download_df(filtered, f'alertas_{machine}.csv')
    else:
        st.info('No se registraron alertas simuladas para este equipo.')
    st.caption('Se muestran solo lecturas con alerta, no la serie completa. Una alerta no equivale a una falla.')

else:
    st.subheader('Validación y trazabilidad')
    st.success('Validaciones de integración aprobadas según el archivo cargado.')
    st.dataframe(data['VALIDACIONES'], hide_index=True, use_container_width=True)
    st.subheader('Limitaciones metodológicas')
    st.dataframe(data['LIMITACIONES'], hide_index=True, use_container_width=True)
    st.info('Las tasas de mejora, umbrales IoT y puntuaciones expertas requieren validación antes de uso operativo.')
