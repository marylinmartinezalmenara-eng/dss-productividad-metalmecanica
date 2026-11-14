import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="DSS | Productividad de mano de obra", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
:root {--navy:#14263f;--blue:#245783;--ink:#18324c;--paper:#f5f7fb;}
.stApp {background:#f5f7fb;}
.block-container {max-width:1490px;padding-top:1.4rem;padding-bottom:3rem;}
[data-testid="stSidebar"] {background:#14263f;}
[data-testid="stSidebar"] * {color:#edf4fc;}
[data-testid="stSidebar"] [data-testid="stFileUploader"] {background:#f1f5fa;border-radius:11px;padding:9px;}
[data-testid="stSidebar"] [data-testid="stFileUploader"] * {color:#19354f!important;}
[data-testid="stSidebar"] [data-testid="stRadio"] label {font-size:.91rem;}
[data-testid="stSidebar"] [data-testid="stSlider"] label {font-size:.85rem;}
[data-testid="stSidebar"] hr {border-color:#53657b;}
h1,h2,h3 {color:#193b56;letter-spacing:-.02em;}
div[data-testid="stMetric"] {background:white;padding:17px 19px;border:1px solid #e0e7ef;border-radius:14px;box-shadow:0 3px 12px rgba(24,50,76,.035);}
[data-testid="stMetricLabel"] {color:#49627c;}
[data-testid="stMetricValue"] {color:#163a5e;}
[data-testid="stDataFrame"], [data-testid="stPlotlyChart"] {background:white;border-radius:12px;}
.dss-hero {background:linear-gradient(110deg,#162b45,#255985);border-radius:0 0 17px 17px;padding:28px 27px 22px;color:white;margin:-18px 0 20px;}
.dss-hero h1 {color:white!important;margin:0 0 9px;font-size:2rem;}
.dss-hero p {color:#dce9f7;margin:0;font-size:.94rem;}
.dss-caption {color:#63758b;font-size:.79rem;margin-top:8px;}
</style>
""",unsafe_allow_html=True)
REQUIRED={
 '00_PARAMETROS':['Variable','Valor'],
 '01_AS_IS_OP':['OP','Mes','Capacidad gal','Tanques','Máquina corte','HH fabricación','HH búsqueda 5S','HH recorridos SLP','HH otras NVA'],
 '02_PARADAS_TPM':['ID','Fecha','Mes','Máquina','Motivo','Tipo','Horas parada']}

def parse_file(file):
    x=pd.ExcelFile(file)
    for name,cols in REQUIRED.items():
        if name not in x.sheet_names: raise ValueError(f'Falta hoja obligatoria: {name}')
        have=pd.read_excel(x,sheet_name=name,nrows=0).columns.tolist()
        missing=[c for c in cols if c not in have]
        if missing:raise ValueError(f'En {name} faltan columnas: {", ".join(missing)}')
    p=pd.read_excel(x,'00_PARAMETROS'); a=pd.read_excel(x,'01_AS_IS_OP'); f=pd.read_excel(x,'02_PARADAS_TPM')
    a=a.dropna(subset=['OP']).copy();f=f.dropna(subset=['ID']).copy()
    for col in ['Capacidad gal','Tanques','HH fabricación','HH búsqueda 5S','HH recorridos SLP','HH otras NVA']:
        a[col]=pd.to_numeric(a[col],errors='coerce')
    f['Horas parada']=pd.to_numeric(f['Horas parada'],errors='coerce')
    if a.empty or f.empty:raise ValueError('La base no contiene OP o paradas.')
    if a['OP'].duplicated().any():raise ValueError('Hay códigos de OP duplicados.')
    if a[['Capacidad gal','Tanques','HH fabricación','HH búsqueda 5S','HH recorridos SLP','HH otras NVA']].isna().any().any():raise ValueError('Existen valores numéricos vacíos o inválidos en las OP.')
    if (a[['Tanques','HH fabricación','HH búsqueda 5S','HH recorridos SLP','HH otras NVA']]<0).any().any() or (a['HH fabricación']<=0).any():raise ValueError('Existen HH negativas o HH de fabricación iguales a cero.')
    if f['Horas parada'].isna().any() or (f['Horas parada']<0).any():raise ValueError('Hay horas de parada inválidas.')
    if not a['Mes'].astype(str).str.fullmatch(r'\d{4}-\d{2}').all():raise ValueError('El campo Mes debe ser AAAA-MM.')
    params=dict(zip(p['Variable'].astype(str),pd.to_numeric(p['Valor'],errors='coerce')))
    return a,f,params

@st.cache_data(show_spinner=False)
def load_bytes(data):return parse_file(io.BytesIO(data))

def number(x,d=2):return f'{x:,.{d}f}'

def scenario(a,f,ratios,tef,workers,overlap,mode):
    z=a.copy();z['TE']=z['Tanques']*z['Capacidad gal'].map(tef)
    z['HH 5S AS IS']=z['HH búsqueda 5S'];z['HH SLP AS IS']=z['HH recorridos SLP']
    z['HH otras AS IS']=z['HH otras NVA'];z['HH fabricación base']=z['HH fabricación']
    z['HH AS IS']=z['HH fabricación base']+z['HH 5S AS IS']+z['HH SLP AS IS']+z['HH otras AS IS']
    z['HH recuperadas 5S']=z['HH 5S AS IS']*ratios['5S'];z['HH recuperadas SLP']=z['HH SLP AS IS']*ratios['SLP']
    # Allocate TPM loss by machine/month to associated OPs, not by arbitrary equal shares across the whole plant.
    ff=f.groupby(['Mes','Máquina'],as_index=False)['Horas parada'].sum().rename(columns={'Mes':'Mes','Máquina':'Máquina corte','Horas parada':'Horas parada máquina'})
    z=z.merge(ff,on=['Mes','Máquina corte'],how='left');z['Horas parada máquina']=z['Horas parada máquina'].fillna(0)
    n=z.groupby(['Mes','Máquina corte'])['OP'].transform('count').clip(lower=1)
    z['HH TPM atribuibles']=z['Horas parada máquina']*workers/n
    z['HH recuperadas TPM']=z['HH TPM atribuibles']*ratios['TPM']*overlap
    z['HH recuperadas']=z[['HH recuperadas 5S','HH recuperadas SLP','HH recuperadas TPM']].sum(axis=1)
    z['HH TO BE']=z['HH AS IS']-z['HH recuperadas']
    if mode=='Conservador (no reducir HH fabricación)':
        z['HH TO BE']=np.maximum(z['HH TO BE'],z['HH fabricación base']+z['HH otras AS IS'])
        z['HH recuperadas']=z['HH AS IS']-z['HH TO BE']
    z['Tanques/HH AS IS']=z['Tanques']/z['HH AS IS']
    z['Tanques/HH TO BE']=z['Tanques']/z['HH TO BE']
    z['TE/HH AS IS']=z['TE']/z['HH AS IS']
    z['TE/HH TO BE']=z['TE']/z['HH TO BE']
    z['Mejora %']=(z['TE/HH TO BE']/z['TE/HH AS IS']-1)*100
    return z

def aggregate(z):
    cols=['Tanques','TE','HH AS IS','HH TO BE','HH 5S AS IS','HH SLP AS IS','HH otras AS IS','HH recuperadas 5S','HH recuperadas SLP','HH recuperadas TPM','HH recuperadas']
    m=z.groupby('Mes',as_index=False)[cols].sum().sort_values('Mes')
    for x in [m]:
        x['Tanques/HH AS IS']=x['Tanques']/x['HH AS IS'];x['Tanques/HH TO BE']=x['Tanques']/x['HH TO BE']
        x['TE/HH AS IS']=x['TE']/x['HH AS IS'];x['TE/HH TO BE']=x['TE']/x['HH TO BE']
        x['Mejora %']=(x['TE/HH TO BE']/x['TE/HH AS IS']-1)*100
    return m

def chart_bars(labels,values,title):
    fig=px.bar(x=values,y=labels,orientation='h',labels={'x':'Horas-hombre (HH)','y':''},title=title,color_discrete_sequence=['#246c91'])
    fig.update_layout(height=330,margin=dict(l=10,r=10,t=55,b=10));return fig

st.markdown("""<div class="dss-hero"><h1>DSS · Productividad de mano de obra</h1><p>Fabricación de tanques de acero · 5S + SLP + TPM · Mejora continua · Agosto 2025 – julio 2026</p></div>""",unsafe_allow_html=True)
with st.sidebar:
    st.markdown('### 🏭 DSS INDUSTRIAL')
    st.caption('DECISION SUPPORT SYSTEM · PRODUCTIVIDAD')
    st.divider()
    st.markdown('#### 📁 Base de datos')
    up=st.file_uploader('Subir Excel DSS',type=['xlsx'])
    st.caption('Formato: DSS_CASAL_INPUT_12_MESES.xlsx')
    if up is None:
        st.info('Sube el Excel para iniciar el análisis.');st.stop()
    try:a,f,p=load_bytes(up.getvalue())
    except Exception as e:st.error(f'No se puede cargar el Excel: {e}');st.stop()
    st.success(f'{len(a)} OP · {len(f)} paradas · {a.Mes.nunique()} meses')
    st.header('⚙️ Hipótesis TO BE')
    r5=st.slider('Reducción de búsquedas · 5S (%)',0,100,int(float(p.get('Reducción 5S',0.5))*100),5)/100
    rs=st.slider('Reducción de recorridos · SLP (%)',0,100,int(float(p.get('Reducción SLP',0.5))*100),5)/100
    rt=st.slider('Reducción de paradas · TPM (%)',0,100,int(float(p.get('Reducción TPM',0.4))*100),5)/100
    workers=st.number_input('Operarios afectados por parada (supuesto)',min_value=1,max_value=45,value=int(p.get('Operarios afectados por parada',2)))
    overlap=st.slider('Fracción TPM no contabilizada en otras pérdidas (%)',0,100,int(float(p.get('Fracción TPM sin solapamiento',0))*100),5)/100
    st.caption('Mantener TPM en 0% de fracción no solapada hasta verificar que las esperas por paradas no estén incluidas en otras NVA.')
    goal=st.number_input('Meta de productividad (%)',min_value=0.0,max_value=100.0,value=float(p.get('Meta mejora',0.1))*100,step=1.0)
    mode=st.selectbox('Criterio de simulación',['Conservador (no reducir HH fabricación)'])
    with st.expander('Factores de tanque equivalente (TE)'):
        t250=st.number_input('250 gal',value=float(p.get('TE 250',1)),min_value=.01)
        t1000=st.number_input('1,000 gal',value=float(p.get('TE 1000',2)),min_value=.01)
        t5000=st.number_input('5,000 gal',value=float(p.get('TE 5000',5)),min_value=.01)
    st.divider()
    page=st.radio('Navegación',['Dashboard integrado','AS IS por OP','Pérdidas y Pareto','Simulador TO BE','TPM e IoT simulado','Recomendaciones trazables','Calidad y exportación'])

z=scenario(a,f,{'5S':r5,'SLP':rs,'TPM':rt},{250:t250,1000:t1000,5000:t5000},workers,overlap,mode)
m=aggregate(z)
phys=z['Tanques'].sum();te=z['TE'].sum();h0=z['HH AS IS'].sum();h1=z['HH TO BE'].sum();rec=h0-h1
p0=te/h0;p1=te/h1;delta=(p1/p0-1)*100

def summary_table():
    return pd.DataFrame({'Indicador':['Tanques físicos','Tanques equivalentes (TE)','HH utilizadas / requeridas','Tanques/HH','TE/HH','Mejora de productividad TE/HH'],
                         'AS IS':[phys,te,h0,phys/h0,p0,0], 'TO BE simulado':[phys,te,h1,phys/h1,p1,delta],
                         'Meta / criterio':['Misma producción','Misma mezcla','Reducción de HH requerida','Mejorar',f'+{goal:.1f}%',f'+{goal:.1f}%']})

def show_status():
    if delta>=goal:st.success(f'La proyección alcanza la meta (+{delta:.2f}% frente a +{goal:.1f}%).')
    else:st.warning(f'La simulación NO alcanza la meta: +{delta:.2f}% frente a +{goal:.1f}%. Brecha: {goal-delta:.2f} puntos porcentuales.')

if page=='Dashboard integrado':
    st.subheader('Diagnóstico integral · AS IS vs. TO BE vs. meta')
    c1,c2,c3,c4=st.columns(4)
    c1.metric('Tanques/HH · AS IS',f'{phys/h0:.5f}')
    c2.metric('TE/HH · AS IS',f'{p0:.5f}')
    c3.metric('HH recuperables',number(rec))
    c4.metric('Mejora TO BE',f'{delta:.2f}%',delta=f'{delta-goal:+.2f} pp vs meta',delta_color='normal')
    st.dataframe(summary_table(),use_container_width=True,hide_index=True)
    show_status()
    c1,c2=st.columns(2)
    with c1:
        losses=pd.DataFrame({'Causa':['5S · búsquedas','SLP · recorridos','Otras NVA','TPM · horas parada máquina'],
            'HH':[z['HH 5S AS IS'].sum(),z['HH SLP AS IS'].sum(),z['HH otras AS IS'].sum(),f['Horas parada'].sum()]})
        st.plotly_chart(chart_bars(losses['Causa'],losses['HH'],'Pérdidas AS IS (TPM en horas máquina)'),use_container_width=True)
    with c2:
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=m['Mes'],y=m['TE/HH AS IS'],name='AS IS',mode='lines+markers'))
        fig.add_trace(go.Scatter(x=m['Mes'],y=m['TE/HH TO BE'],name='TO BE simulado',mode='lines+markers'))
        fig.add_trace(go.Scatter(x=m['Mes'],y=m['TE/HH AS IS']*(1+goal/100),name='Meta mensual (+%)',line=dict(dash='dash')))
        fig.update_layout(title='Productividad TE/HH por mes',height=330,margin=dict(l=10,r=10,t=55,b=10))
        st.plotly_chart(fig,use_container_width=True)
    st.subheader('Comparación mensual')
    st.dataframe(m[['Mes','Tanques','TE','HH AS IS','HH recuperadas','HH TO BE','TE/HH AS IS','TE/HH TO BE','Mejora %']],hide_index=True,use_container_width=True)

elif page=='AS IS por OP':
    st.subheader('Órdenes de producción · escenario actual')
    months=st.multiselect('Meses',sorted(z['Mes'].unique()),default=sorted(z['Mes'].unique()))
    sizes=st.multiselect('Tamaños (gal)',sorted(z['Capacidad gal'].unique()),default=sorted(z['Capacidad gal'].unique()))
    v=z[z['Mes'].isin(months)&z['Capacidad gal'].isin(sizes)]
    st.dataframe(v[['OP','Mes','Capacidad gal','Tanques','Máquina corte','HH fabricación base','HH 5S AS IS','HH SLP AS IS','HH otras AS IS','HH AS IS','TE','TE/HH AS IS']],hide_index=True,use_container_width=True)
    st.caption('HH AS IS = HH fabricación + HH búsqueda + HH recorridos + HH otras NVA. Revisar si la fabricación ya incorpora esperas para evitar doble conteo.')

elif page=='Pérdidas y Pareto':
    st.subheader('Análisis de pérdidas · AS IS')
    causes=pd.DataFrame({'Causa':['Búsqueda de herramientas','Recorridos innecesarios','Otras NVA'],
        'Herramienta':['5S','SLP','No atribuida'],
        'HH':[z['HH 5S AS IS'].sum(),z['HH SLP AS IS'].sum(),z['HH otras AS IS'].sum()]})
    causes=causes.sort_values('HH',ascending=False)
    causes['% acumulado']=causes['HH'].cumsum()/causes['HH'].sum()*100
    st.dataframe(causes,hide_index=True,use_container_width=True)
    st.plotly_chart(chart_bars(causes['Causa'],causes['HH'],'Pareto de HH NVA clasificadas'),use_container_width=True)
    st.info('Las horas de parada de máquinas se presentan aparte: no son automáticamente HH perdidas de operarios.')
    st.dataframe(f.groupby(['Máquina','Motivo'],as_index=False).agg(Eventos=('ID','count'),**{'Horas parada':('Horas parada','sum')}).sort_values('Horas parada',ascending=False),hide_index=True,use_container_width=True)

elif page=='Simulador TO BE':
    st.subheader('Aplicación de herramientas → HH recuperadas → productividad')
    tools=pd.DataFrame({'Herramienta':['5S','SLP','TPM'],
        'Pérdida de referencia (HH)':[z['HH 5S AS IS'].sum(),z['HH SLP AS IS'].sum(),z['HH TPM atribuibles'].sum()],
        'Reducción supuesta':[r5,rs,rt*overlap],
        'HH recuperadas':[z['HH recuperadas 5S'].sum(),z['HH recuperadas SLP'].sum(),z['HH recuperadas TPM'].sum()]})
    st.dataframe(tools,hide_index=True,use_container_width=True,column_config={'Reducción supuesta':st.column_config.NumberColumn(format='%.1f')})
    st.metric('HH recuperadas calculadas',number(rec))
    st.dataframe(summary_table(),hide_index=True,use_container_width=True)
    show_status()
    st.caption('En este modelo TO BE se mantiene la misma producción y se reduce el tiempo requerido. Para afirmar que se fabricarán tanques adicionales es necesario verificar demanda, tiempos estándar y recursos disponibles.')
    st.subheader('Trazabilidad por OP')
    st.dataframe(z[['OP','Mes','Capacidad gal','HH AS IS','HH recuperadas 5S','HH recuperadas SLP','HH recuperadas TPM','HH TO BE','TE/HH AS IS','TE/HH TO BE','Mejora %']],hide_index=True,use_container_width=True)

elif page=='TPM e IoT simulado':
    st.subheader('Paradas de máquinas de corte · registro histórico del dataset')
    g=f.groupby('Máquina',as_index=False).agg(Eventos=('ID','count'),Horas_parada=('Horas parada','sum'),MTTR_h=('Horas parada','mean'))
    st.dataframe(g.rename(columns={'Horas_parada':'Horas parada','MTTR_h':'MTTR (h/falla)'}),hide_index=True,use_container_width=True)
    st.plotly_chart(px.bar(g,x='Máquina',y='Horas_parada',title='Paradas por equipo',color_discrete_sequence=['#246c91']),use_container_width=True)
    st.info('MTBF requiere horas reales de funcionamiento por máquina; el Excel actual no las contiene. No se calcula un MTBF ficticio.')
    st.subheader('Alertas IoT · reglas simuladas, NO sensores conectados')
    threshold=st.number_input('Umbral de alerta por evento (horas)',min_value=0.1,value=2.0,step=0.5)
    alerts=f[f['Horas parada']>=threshold].copy()
    alerts['Alerta simulada']=np.where(alerts['Horas parada']>=threshold*2,'CRÍTICA','REVISAR')
    st.dataframe(alerts[['ID','Mes','Máquina','Motivo','Horas parada','Alerta simulada']],hide_index=True,use_container_width=True)

elif page=='Recomendaciones trazables':
    st.subheader('Recomendaciones según causas cuantificadas')
    catalog=pd.DataFrame([
        ['5S','Búsqueda de herramientas',z['HH 5S AS IS'].sum(),'Clasificar, señalizar y ubicar herramientas por punto de uso','Jefe de producción','Minutos de búsqueda/OP'],
        ['SLP','Recorridos innecesarios',z['HH SLP AS IS'].sum(),'Medir recorridos, diseñar FROM–TO y rediseñar ubicaciones','Ingeniería de procesos','Metros recorridos/OP'],
        ['TPM','Paradas correctivas',z['HH TPM atribuibles'].sum(),'Mantenimiento autónomo, inspecciones y análisis de fallas','Mantenimiento','HH parada y MTTR'] ],columns=['Herramienta','Problema','HH impacto potencial','Acción sugerida','Responsable sugerido','KPI'])
    catalog['Score heurístico']=np.where(catalog['HH impacto potencial'].max()>0,100*catalog['HH impacto potencial']/catalog['HH impacto potencial'].max(),0)
    catalog=catalog.sort_values('Score heurístico',ascending=False)
    tool=st.selectbox('Filtrar herramienta',['Todas','5S','SLP','TPM'])
    view=catalog if tool=='Todas' else catalog[catalog['Herramienta']==tool]
    st.dataframe(view,hide_index=True,use_container_width=True)
    st.caption('El puntaje ordena por HH potenciales y debe contrastarse con criterios de criticidad y factibilidad.')

elif page=='Calidad y exportación':
    st.subheader('Auditoría de datos y descarga')
    checks=pd.DataFrame([
        ['OP únicas',len(z),len(z['OP'].unique())==len(z)],
        ['Meses',z['Mes'].nunique(),z['Mes'].nunique()==12],
        ['Paradas',len(f),True],
        ['Máquinas de corte',f['Máquina'].nunique(),f['Máquina'].nunique()==4],
        ['HH TO BE positivas',int((z['HH TO BE']>0).sum()),bool((z['HH TO BE']>0).all())],
        ['TPM sin solapamiento validado',f'{overlap*100:.0f}%','PENDIENTE DE EVIDENCIA' if overlap==0 else 'HIPÓTESIS EDITADA'] ],columns=['Control','Resultado','Cumple / estado'])
    st.dataframe(checks,hide_index=True,use_container_width=True)
    with st.expander('Alcance metodológico y datos pendientes de validar'):
        st.caption('Los escenarios TO BE son proyecciones basadas en supuestos editables; no equivalen a resultados observados en planta. Verificar HH extra efectivas, si HH fabricación incluye NVA, horas operativas para MTBF, fracción TPM no solapada y factores TE.')
    output=io.BytesIO()
    with pd.ExcelWriter(output,engine='openpyxl') as writer:
        z.to_excel(writer,sheet_name='DSS_OP_CALCULADO',index=False)
        m.to_excel(writer,sheet_name='DSS_COMPARATIVO_12M',index=False)
        f.to_excel(writer,sheet_name='PARADAS_ORIGEN',index=False)
        summary_table().to_excel(writer,sheet_name='DSS_RESUMEN',index=False)
    st.download_button('⬇️ Exportar cálculos DSS en Excel',data=output.getvalue(),file_name='DSS_RESULTADOS_ESCENARIOS.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
