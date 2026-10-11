import io
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='DSS | AS IS vs TO BE', page_icon='🏭', layout='wide')
st.markdown('''<style>
.block-container {padding-top:1.6rem; max-width:1600px}
[data-testid="stMetric"] {border:1px solid #d9dfe5; border-radius:10px; padding:12px; background:#fff}
[data-testid="stMetricLabel"] {font-weight:650}
h1,h2,h3 {letter-spacing:-0.025em}
</style>''', unsafe_allow_html=True)
st.title('DSS | Productividad de mano de obra')
st.caption('Caso metalmecánico · 5S + SLP + TPM · IoT simulado · Datos modelados para fines académicos')

REQUIRED = {
 'INDICADORES': ['Indicador','Valor','Unidad'],
 'KPIS_MENSUALES': ['Periodo','Unidades','Tanques equivalentes','HH totales','HH extras asignadas'],
 'PARETO_CAUSAS': ['Causa','HH_perdidas'],
 'PERDIDAS_GRUPOS': ['Grupo DSS','HH_perdidas'],
 'DECISIONES_DSS': ['Prioridad','Problema','Herramienta','HH perdidas','Puntuación DSS','Acción recomendada','Responsable','Indicador de seguimiento'],
 'ESCENARIOS': ['Escenario','HH liberadas hipotéticas','Mejora A (%)','Mejora B potencial (%)','Estado B'],
 'MONITOREO_EQUIPOS': ['ID equipo','Alertas_advertencia','Alertas_criticas','Prioridad de revisión'],
 'ALERTAS_IOT': ['ID equipo','Fecha y hora','Temperatura (°C)','Vibración (mm/s)','Estado DSS'],
 'VALIDACIONES': ['Prueba','Estado'],
 'LIMITACIONES': ['Aspecto','Limitación']
}

@st.cache_data(show_spinner=False)
def read_data(file_bytes):
    book = pd.ExcelFile(io.BytesIO(file_bytes))
    missing = set(REQUIRED) - set(book.sheet_names)
    if missing:
        raise ValueError('Faltan hojas: '+', '.join(sorted(missing)))
    data = {name: pd.read_excel(book, sheet_name=name) for name in REQUIRED}
    for name, columns in REQUIRED.items():
        missing_cols = set(columns) - set(data[name].columns)
        if missing_cols:
            raise ValueError(f'{name}: faltan columnas {sorted(missing_cols)}')
    if not data['VALIDACIONES']['Estado'].astype(str).eq('APROBADA').all():
        raise ValueError('El archivo registra validaciones no aprobadas.')
    return data

def download_csv(df, filename, label='Descargar CSV'):
    st.download_button(label, data=df.to_csv(index=False).encode('utf-8-sig'), file_name=filename, mime='text/csv')

uploaded = st.sidebar.file_uploader('Cargar Checkpoint 08 (.xlsx)', type='xlsx')
if uploaded is None:
    st.info('Carga el archivo del Checkpoint 08 para activar el simulador y los módulos del DSS.')
    st.stop()
try:
    data = read_data(uploaded.getvalue())
except Exception as exc:
    st.error(f'Error al leer el Excel: {exc}')
    st.stop()

monthly = data['KPIS_MENSUALES'].copy()
for col in ['Unidades','Tanques equivalentes','HH totales','HH extras asignadas']:
    monthly[col] = pd.to_numeric(monthly[col], errors='raise')
physical = float(monthly['Unidades'].sum())
te = float(monthly['Tanques equivalentes'].sum())
hh = float(monthly['HH totales'].sum())
extra = float(monthly['HH extras asignadas'].sum())
if min(physical, te, hh) <= 0:
    st.error('La producción y las HH deben ser positivas.')
    st.stop()

losses = data['PERDIDAS_GRUPOS'].copy()
losses['Grupo DSS'] = losses['Grupo DSS'].astype(str).str.strip().str.upper()
losses['HH_perdidas'] = pd.to_numeric(losses['HH_perdidas'], errors='raise')
if losses['HH_perdidas'].lt(0).any():
    st.error('Se encontraron pérdidas negativas.')
    st.stop()
loss_map = losses.groupby('Grupo DSS')['HH_perdidas'].sum().to_dict()
if sum(loss_map.values()) > hh + 0.02:
    st.error('Las pérdidas registradas exceden las HH totales.')
    st.stop()

st.sidebar.header('Navegación')
menu = st.sidebar.radio('Módulo', [
    'Dashboard integrado', 'AS IS vs TO BE', 'Resumen ejecutivo', 'Diagnóstico',
    'Recomendaciones', 'Escenarios originales', 'Monitoreo IoT',
    'Trazabilidad y límites'
])
st.sidebar.divider()
st.sidebar.caption('Fuente: Checkpoint 08. Los datos empresariales son modelados y las mejoras son hipótesis, no resultados observados.')

if menu == 'Dashboard integrado':
    st.header('Dashboard integrado · diagnóstico general')
    st.caption('Indicadores de la situación AS IS. El TO BE se calcula únicamente en el simulador interactivo.')
    total_lost = float(losses['HH_perdidas'].sum())
    overtime_rate = extra/hh*100
    a,b,c,d = st.columns(4)
    a.metric('Productividad física AS IS',f'{physical/hh:.5f} tanques/HH')
    b.metric('Productividad equivalente AS IS',f'{te/hh:.5f} TE/HH')
    c.metric('Horas extra / HH',f'{overtime_rate:.2f}%')
    d.metric('HH perdidas clasificadas',f'{total_lost:,.2f} HH')

    left,right = st.columns([1.35,1],gap='large')
    with left:
        st.subheader('Pérdidas cuantificadas por herramienta')
        g = losses.groupby('Grupo DSS',as_index=False)['HH_perdidas'].sum().sort_values('HH_perdidas')
        fig = px.bar(g,x='HH_perdidas',y='Grupo DSS',orientation='h',text='HH_perdidas',
                     labels={'HH_perdidas':'Horas-hombre perdidas','Grupo DSS':'Grupo'})
        fig.update_traces(marker_color='#37678a',texttemplate='%{text:.2f}',textposition='outside')
        fig.update_layout(height=320,margin=dict(l=10,r=45,t=10,b=20),showlegend=False)
        st.plotly_chart(fig,use_container_width=True)
    with right:
        st.subheader('Recomendaciones priorizadas')
        rec = data['DECISIONES_DSS'].sort_values('Prioridad').copy()
        st.dataframe(rec[['Prioridad','Problema','Herramienta','HH perdidas','Acción recomendada']].head(5),
                     hide_index=True,use_container_width=True,height=295)

    colplan,coltrend = st.columns([1.35,1],gap='large')
    with colplan:
        st.subheader('Distribución actual de planta · SLP')
        from pathlib import Path
        path = Path(__file__).with_name('layout_actual.png')
        if path.exists():
            st.image(str(path),caption='Plano proporcionado: distribución AS IS (áreas A–I). No representa rutas ni distancias medidas.',use_container_width=True)
        else:
            st.warning('Para visualizar el plano, sube layout_actual.png a la raíz del repositorio.')
        with st.expander('Identificación de áreas del plano'):
            areas = {'A':'Vestuario','B':'Área administrativa','C':'Área de corte',
                     'D':'Pintura / almacén de pinturas','E':'Baños y lavado',
                     'F':'Área de arenado','G':'Almacén de planchas',
                     'H':'Almacén de repuestos y máquinas','I':'Patio de maniobras'}
            st.dataframe(pd.DataFrame({'Área':list(areas),'Descripción':list(areas.values())}),hide_index=True,use_container_width=True)
        st.info('Para dibujar flujos o calcular distancias SLP se necesita una matriz From–To y distancias por recorrido. El Checkpoint 08 no incluye geometría de rutas verificadas.')
    with coltrend:
        st.subheader('Productividad equivalente mensual · AS IS')
        monthly_view = monthly.copy()
        monthly_view['Periodo'] = monthly_view['Periodo'].astype(str)
        monthly_view['TE/HH'] = monthly_view['Tanques equivalentes']/monthly_view['HH totales']
        fig = px.line(monthly_view,x='Periodo',y='TE/HH',markers=True,
                      labels={'Periodo':'Período','TE/HH':'Tanques equivalentes / HH'})
        fig.update_traces(line_color='#37678a')
        fig.update_layout(height=280,margin=dict(l=10,r=10,t=10,b=30))
        st.plotly_chart(fig,use_container_width=True)
        st.subheader('Alertas de condición · IoT simulado')
        eq = data['MONITOREO_EQUIPOS'].copy()
        st.dataframe(eq[['ID equipo','Alertas_advertencia','Alertas_criticas','Prioridad de revisión']],
                     hide_index=True,use_container_width=True)
        st.caption('Las alertas son simulaciones por umbrales; no son fallas reales ni datos de sensores instalados.')

elif menu == 'AS IS vs TO BE':
    st.subheader('Simulador de productividad AS IS vs TO BE')
    st.write('Ajusta la reducción **hipotética** de las pérdidas de cada herramienta. Las HH liberadas se calculan únicamente sobre las pérdidas atribuidas a cada grupo.')
    col1, col2, col3 = st.columns(3)
    with col1:
        rate_5s = st.slider('5S · reducción de búsquedas y desorden', 0, 60, 20, 5, help=f'Pérdidas 5S: {loss_map.get("5S",0):.2f} HH')
    with col2:
        rate_slp = st.slider('SLP · reducción de recorridos', 0, 60, 25, 5, help=f'Pérdidas SLP: {loss_map.get("SLP",0):.2f} HH')
    with col3:
        rate_tpm = st.slider('TPM · reducción de pérdidas por paradas', 0, 60, 30, 5, help=f'Pérdidas TPM: {loss_map.get("TPM",0):.2f} HH')

    rates = {'5S':rate_5s/100, 'SLP':rate_slp/100, 'TPM':rate_tpm/100}
    detail = pd.DataFrame([
        {'Herramienta':key, 'Pérdidas AS IS (HH)':loss_map.get(key,0),
         'Reducción hipotética (%)':rates[key]*100,
         'HH liberadas hipotéticas':loss_map.get(key,0)*rates[key],
         'Pérdidas TO BE (HH)':loss_map.get(key,0)*(1-rates[key])}
        for key in ['5S','SLP','TPM']
    ])
    freed = float(detail['HH liberadas hipotéticas'].sum())
    hh_to_be = hh - freed
    if hh_to_be <= 0:
        st.error('Las HH propuestas no pueden ser cero o negativas.')
        st.stop()
    p_as_is = physical/hh
    p_to_be = physical/hh_to_be
    te_as_is = te/hh
    te_to_be = te/hh_to_be
    improvement = (p_to_be/p_as_is-1)*100

    st.markdown('#### Escenario A: misma producción, menos HH')
    a,b,c,d = st.columns(4)
    a.metric('AS IS · Tanques/HH', f'{p_as_is:.5f}')
    b.metric('TO BE · Tanques/HH', f'{p_to_be:.5f}', delta=f'+{improvement:.2f}% hipotético')
    c.metric('AS IS · TE/HH', f'{te_as_is:.5f}')
    d.metric('TO BE · TE/HH', f'{te_to_be:.5f}')
    e,f,g = st.columns(3)
    e.metric('HH AS IS', f'{hh:,.2f}')
    f.metric('HH TO BE', f'{hh_to_be:,.2f}')
    g.metric('HH liberadas (hipótesis)', f'{freed:,.2f}')
    comparison = pd.DataFrame([
        {'Estado':'AS IS','Tanques/HH':p_as_is,'TE/HH':te_as_is},
        {'Estado':'TO BE hipotético','Tanques/HH':p_to_be,'TE/HH':te_to_be}
    ])
    c1,c2 = st.columns(2)
    with c1:
        fig = px.bar(comparison, x='Estado', y='Tanques/HH', text_auto='.5f', title='Productividad física (tanques/HH)')
        fig.update_yaxes(range=[0, max(p_as_is,p_to_be)*1.18])
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(detail, x='Herramienta', y='HH liberadas hipotéticas', text_auto='.2f', title='HH potencialmente liberadas por herramienta')
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(detail.round(3), hide_index=True, use_container_width=True)
    st.caption('No se suman las horas de parada de máquina a las HH perdidas. Las tasas se aplican una vez a cada grupo de pérdidas; no se supone recuperación completa.')

    st.markdown('#### Escenario B: mismas HH, producción adicional potencial')
    st.write('Conversión teórica de las HH liberadas a producción adicional usando la productividad actual. No demuestra capacidad real ni pedidos adicionales.')
    utilization = st.slider('Aprovechamiento hipotético de HH liberadas (%)', 0, 100, 100, 5)
    extra_te = freed * utilization/100 * te_as_is
    extra_units = freed * utilization/100 * p_as_is
    q1,q2,q3 = st.columns(3)
    q1.metric('TE adicionales potenciales', f'{extra_te:.2f}')
    q2.metric('Tanques físicos equivalentes aproximados', f'{extra_units:.2f}')
    q3.metric('Mejora teórica B', f'{freed/hh*utilization:.2f}%')
    st.warning('Escenario B NO VALIDADO: faltan demanda, capacidad de los equipos, cuello de botella y mezcla de tanques. El número de tanques es una equivalencia matemática, no un plan de producción de unidades enteras.')
    st.markdown('#### Trazabilidad de la mejora')
    st.write(f'AS IS: **{physical:.0f} tanques / {hh:,.2f} HH = {p_as_is:.5f} tanques/HH**')
    st.write(f'TO BE A: **{physical:.0f} tanques / ({hh:,.2f} − {freed:,.2f}) HH = {p_to_be:.5f} tanques/HH**')
    st.write(f'Incremento proyectado: **{improvement:.2f}%**. Esta variación es un resultado del modelo, no una mejora implementada y observada.')
    export = detail.copy()
    export['Tanques AS IS'] = physical
    export['HH totales AS IS'] = hh
    export['Productividad AS IS (tanques/HH)'] = p_as_is
    export['Productividad TO BE (tanques/HH)'] = p_to_be
    export['Mejora hipotética total (%)'] = improvement
    download_csv(export, 'simulacion_as_is_to_be.csv', 'Descargar simulación AS IS vs TO BE')

elif menu == 'Resumen ejecutivo':
    st.subheader('Situación actual y decisión principal')
    a,b,c,d = st.columns(4)
    a.metric('Productividad física AS IS',f'{physical/hh:.5f} tanques/HH')
    b.metric('Productividad equivalente AS IS',f'{te/hh:.5f} TE/HH')
    c.metric('Horas extra',f'{extra:,.2f} HH')
    d.metric('HH perdidas',f'{sum(loss_map.values()):,.2f}')
    top = data['DECISIONES_DSS'].sort_values('Prioridad').iloc[0]
    st.info(f"**Prioridad 1:** {top['Problema']} · **Herramienta:** {top['Herramienta']}\n\n**Acción:** {top['Acción recomendada']}")
    st.write('El TO BE se calcula en el módulo **AS IS vs TO BE**, con parámetros modificables y explícitamente hipotéticos.')

elif menu == 'Diagnóstico':
    st.subheader('Diagnóstico de pérdidas y productividad')
    groups = losses.sort_values('HH_perdidas',ascending=False)
    st.plotly_chart(px.bar(groups,x='HH_perdidas',y='Grupo DSS',orientation='h',title='HH perdidas por grupo'),use_container_width=True)
    pareto = data['PARETO_CAUSAS'].sort_values('HH_perdidas',ascending=False)
    st.plotly_chart(px.bar(pareto,x='Causa',y='HH_perdidas',title='HH perdidas por causa'),use_container_width=True)
    download_csv(pareto,'diagnostico.csv')
    monthly['Periodo'] = monthly['Periodo'].astype(str)
    monthly['Productividad física (tanques/HH)'] = monthly['Unidades']/monthly['HH totales']
    monthly['Productividad equivalente (TE/HH)'] = monthly['Tanques equivalentes']/monthly['HH totales']
    st.plotly_chart(px.line(monthly,x='Periodo',y='Productividad física (tanques/HH)',markers=True,title='Productividad física mensual AS IS'),use_container_width=True)
    st.dataframe(monthly,hide_index=True,use_container_width=True)

elif menu == 'Recomendaciones':
    st.subheader('Reglas de decisión 5S · SLP · TPM')
    rec = data['DECISIONES_DSS'].sort_values('Prioridad')
    choice = st.selectbox('Herramienta',['Todas']+sorted(rec['Herramienta'].dropna().unique().tolist()))
    view = rec if choice=='Todas' else rec[rec['Herramienta'].eq(choice)]
    st.dataframe(view,hide_index=True,use_container_width=True)
    if not view.empty:
        cause = st.selectbox('Causa a consultar',view['Problema'].tolist())
        row = view.loc[view['Problema'].eq(cause)].iloc[0]
        st.markdown(f"**Acción:** {row['Acción recomendada']}")
        st.write('Responsable:',row['Responsable'])
        st.write('KPI:',row['Indicador de seguimiento'])
    st.caption('La factibilidad y urgencia del ranking todavía requieren evaluación experta.')
    download_csv(view,'recomendaciones.csv')

elif menu == 'Escenarios originales':
    st.subheader('Ocho escenarios del Checkpoint 06')
    scenarios = data['ESCENARIOS'].copy()
    selected = st.multiselect('Escenarios',scenarios['Escenario'].tolist(),default=scenarios['Escenario'].tolist())
    view = scenarios[scenarios['Escenario'].isin(selected)]
    mode = st.radio('Enfoque',['A · mismas unidades, menos HH','B · mismas HH, producción potencial'],horizontal=True)
    col = 'Mejora A (%)' if mode.startswith('A') else 'Mejora B potencial (%)'
    if not view.empty:
        st.plotly_chart(px.bar(view,x='Escenario',y=col,title='Mejora hipotética (%)'),use_container_width=True)
        st.dataframe(view,hide_index=True,use_container_width=True)
        download_csv(view,'escenarios_originales.csv')
    st.warning('Las mejoras corresponden a hipótesis de CP06; la producción adicional B no está validada.')

elif menu == 'Monitoreo IoT':
    st.subheader('Monitoreo simulado de cuatro máquinas')
    machines = data['MONITOREO_EQUIPOS'].copy()
    st.dataframe(machines,hide_index=True,use_container_width=True)
    machine = st.selectbox('Máquina',sorted(machines['ID equipo'].astype(str).unique()))
    alerts = data['ALERTAS_IOT'].copy()
    alerts['Fecha y hora'] = pd.to_datetime(alerts['Fecha y hora'],errors='coerce')
    view = alerts.loc[alerts['ID equipo'].astype(str).eq(machine)].sort_values('Fecha y hora')
    st.metric('Lecturas con alerta',len(view))
    if not view.empty:
        for measure in ['Temperatura (°C)','Vibración (mm/s)']:
            st.plotly_chart(px.line(view,x='Fecha y hora',y=measure,markers=True,title=f'{measure} · solo lecturas con alerta'),use_container_width=True)
        st.dataframe(view,hide_index=True,use_container_width=True)
        download_csv(view,f'alertas_{machine}.csv')
    else:
        st.info('Sin alertas simuladas para esta máquina.')
    st.caption('Alertas simuladas ≠ fallas confirmadas. No utilizar para operación real ni calcular MTBF/MTTR.')

else:
    st.subheader('Validaciones, trazabilidad y limitaciones')
    st.dataframe(data['VALIDACIONES'],hide_index=True,use_container_width=True)
    st.dataframe(data['LIMITACIONES'],hide_index=True,use_container_width=True)
    st.warning('La validación automática comprueba coherencia de datos y reglas; no acredita mejoras reales, viabilidad operativa ni precisión de sensores.')
