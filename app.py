import io, json, time, platform
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='DSS | Productividad industrial', page_icon='🏭', layout='wide')
st.markdown('''<style>.block-container{padding-top:1.3rem;max-width:1500px}div[data-testid="stMetric"]{background:#f2f5fa;border:1px solid #dce5ef;padding:14px;border-radius:12px}h1,h2,h3{color:#17324d}.stAlert{border-radius:10px}</style>''',unsafe_allow_html=True)
st.title('DSS | Productividad de mano de obra')
st.caption('Empresa A · Fabricación de tanques de acero · Versión V7 · Integración histórica y validación V5 · Plan pendiente de aprobación del asesor')

CAPS=[250,1000,5000]
@st.cache_data(show_spinner=False)
def load_book(content):
    bio=io.BytesIO(content)
    names=pd.ExcelFile(bio).sheet_names
    bio.seek(0)
    if {'OP_REALES','NVA_REALES','PARADAS_REALES'}.issubset(names):
        mapping={'Ordenes_OP':'OP_REALES','Actividades_NVA':'NVA_REALES','Paradas_Corte':'PARADAS_REALES','Personal':'PERSONAL_SUPUESTO'}
        book={k:pd.read_excel(bio,sheet_name=v) for k,v in mapping.items()}
        book['V5']=pd.read_excel(bio,sheet_name='RESULTADOS_V5') if 'RESULTADOS_V5' in names else pd.DataFrame()
        book['semillas']=pd.read_excel(bio,sheet_name='SEMILLAS_Y_TIEMPOS') if 'SEMILLAS_Y_TIEMPOS' in names else pd.DataFrame()
        return book
    book={n:pd.read_excel(bio,sheet_name=n) for n in ['Ordenes_OP','Actividades_NVA','Paradas_Corte','Personal']}
    book['V5']=pd.DataFrame()
    book['semillas']=pd.DataFrame()
    return book

def normalize(book):
    op=book['Ordenes_OP'].copy(); nv=book['Actividades_NVA'].copy(); pa=book['Paradas_Corte'].copy()
    for df in [op,nv,pa]: df.columns=df.columns.astype(str).str.strip()
    for df in [op,nv,pa]:
        if 'Mes' in df: df['Mes']=df['Mes'].astype(str).str.strip()
    for c in ['Capacidad gal','Cantidad','HH fabricación','HH NVA','HH total OP','Distancia NVA m']:
        if c not in op: raise ValueError(f'Falta columna: {c}')
        op[c]=pd.to_numeric(op[c],errors='coerce')
    for c in ['HH perdidas','Distancia recorrida m']: nv[c]=pd.to_numeric(nv[c],errors='coerce')
    pa['Horas parada / reparación']=pd.to_numeric(pa['Horas parada / reparación'],errors='coerce')
    required=['OP','Mes','Capacidad gal','Cantidad','HH fabricación','HH NVA','HH total OP']
    if any(c not in op for c in required): raise ValueError('Faltan columnas esenciales de OP')
    if op['OP'].duplicated().any(): raise ValueError('Códigos OP duplicados')
    op['Capacidad gal']=op['Capacidad gal'].astype('Int64')
    nvagg=nv.groupby('OP',as_index=False).agg(nva_reg=('HH perdidas','sum'),dist_reg=('Distancia recorrida m','sum'))
    op=op.merge(nvagg,on='OP',how='left',validate='one_to_one')
    return op,nv,pa

def productivity(df):
    t=df.groupby('Capacidad gal',dropna=False).agg(tanques=('Cantidad','sum'),HH=('HH total OP','sum'),NVA=('HH NVA','sum'),distancia=('dist_reg','sum'),OP=('OP','count')).reset_index()
    t['tanques_HH']=np.where(t['HH']>0,t['tanques']/t['HH'],np.nan)
    return t

def fmt(n,dec=3): return f'{n:,.{dec}f}'

st.sidebar.header('Datos y configuración')
upload=st.sidebar.file_uploader('Cargar Excel integrado V5 (.xlsx)',type=['xlsx'])
default=Path('/mnt/data/Validacion_DSS_Integrada_Rubrica(1).xlsx')
if upload is None and default.exists():
    content=default.read_bytes(); st.sidebar.info('Se usa la base de demostración local. En Streamlit Cloud carga el Excel.')
elif upload is not None: content=upload.getvalue()
else:
    st.info('Carga Validacion_DSS_Integrada_Rubrica.xlsx desde la barra lateral para comenzar.'); st.stop()
try:
    book=load_book(content)
    op,nva,par=normalize(book)
    v5=book.get('V5',pd.DataFrame())
    v5_seeds=book.get('semillas',pd.DataFrame())
except Exception as e:
    st.error(f'No se pudo leer la estructura: {e}'); st.stop()
months=sorted(op['Mes'].dropna().unique())
selected=st.sidebar.multiselect('Meses del análisis histórico',months,default=months)
capacities=st.sidebar.multiselect('Capacidades (gal)',CAPS,default=CAPS)
if not selected or not capacities: st.warning('Selecciona al menos un mes y una capacidad.');st.stop()
base=op[op['Mes'].isin(selected)&op['Capacidad gal'].isin(capacities)].copy()
pa=par[par['Mes'].isin(selected)].copy()
nv=nva[nva['Mes'].isin(selected)&nva['Capacidad gal'].isin(capacities)].copy()
rec=st.sidebar.slider('Recuperación hipotética de HH NVA (%)',0,40,40,5)/100
nrep=st.sidebar.number_input('Réplicas por escenario',min_value=30,max_value=500,value=50,step=10)
period=st.sidebar.number_input('Meses por réplica',min_value=1,max_value=12,value=6)
seed_master=st.sidebar.number_input('Semilla maestra',min_value=1,value=20261114)
limit=st.sidebar.number_input('Umbral provisional por réplica (segundos)',min_value=1,value=60)
st.sidebar.warning('Recuperación 40 % = meta/escenario, no resultado observado. Personal = SUPUESTO.')

pages=st.tabs(['Resumen ejecutivo','Diagnóstico y priorización','AS-IS / TO-BE','Escenarios y réplicas','Validación y evidencias','Resultados V5 integrados'])
with pages[0]:
    st.subheader('Situación histórica AS-IS')
    st.info('AS-IS histórico: tanques terminados / HH total OP, por capacidad. TO-BE: mejora PROYECTADA, no observada.')
    a,b,c,d=st.columns(4)
    a.metric('OP',f'{len(base):,}');b.metric('Tanques',fmt(base['Cantidad'].sum(),0));c.metric('HH totales',fmt(base['HH total OP'].sum(),1));d.metric('HH NVA',fmt(base['HH NVA'].sum(),1))
    tab=productivity(base);st.dataframe(tab.rename(columns={'tanques_HH':'Tanques/HH','distancia':'Recorridos (m)'}),use_container_width=True,hide_index=True)
    fig=px.bar(tab,x='Capacidad gal',y='tanques_HH',text_auto='.4f',title='Productividad histórica por capacidad',labels={'tanques_HH':'Tanques / HH'})
    st.plotly_chart(fig,use_container_width=True)
    monthly=base.groupby(['Mes','Capacidad gal'],as_index=False).agg(tanques=('Cantidad','sum'),hh=('HH total OP','sum'))
    monthly['productividad']=monthly['tanques']/monthly['hh']
    st.plotly_chart(px.line(monthly,x='Mes',y='productividad',color='Capacidad gal',markers=True,title='Evolución mensual (sin mezclar capacidades)'),use_container_width=True)
with pages[1]:
    st.subheader('Pérdidas observadas y priorización orientativa')
    n=nv.groupby('Actividad sin valor',as_index=False).agg(HH=('HH perdidas','sum'),distancia_m=('Distancia recorrida m','sum')).sort_values('HH',ascending=False)
    st.plotly_chart(px.bar(n,x='Actividad sin valor',y='HH',title='Pareto de HH NVA por actividad'),use_container_width=True)
    st.dataframe(n,use_container_width=True,hide_index=True)
    st.caption('Regla de línea base ingenua: priorizar únicamente por HH NVA acumuladas. No equivale a una comparación experimental ya ejecutada.')
    p=pa.groupby(['Máquina','Motivo'],as_index=False).agg(horas=('Horas parada / reparación','sum'),eventos=('Motivo','size')).sort_values('horas',ascending=False)
    st.plotly_chart(px.bar(p,x='Máquina',y='horas',color='Motivo',title='Horas de parada por máquina y motivo'),use_container_width=True)
    st.dataframe(p,use_container_width=True,hide_index=True)
    st.markdown('**Reglas de recomendación (heurísticas, no optimizadas):** búsquedas → 5S; recorridos → SLP; paradas → TPM; eventos digitales → lógica IoT simulada. Se debe comprobar su pertinencia con el responsable.')
    st.warning('MTTR = duración media de reparación solo si cada evento corresponde a una reparación completa. MTBF requiere horas efectivas de operación por máquina, no disponibles de forma verificada.')
    if len(pa): st.metric('Duración media de parada (proxy; no MTTR validado)',f"{pa['Horas parada / reparación'].mean():.2f} h")
with pages[2]:
    st.subheader('Comparación AS-IS / TO-BE hipotética')
    st.caption('Las HH recuperadas se calculan como recuperación × HH NVA; no se aplica sobre todas las HH ni sobre el 11 % estimado.')
    alternative=st.radio('Interpretación TO-BE',['Producción constante: menos HH','HH constantes: capacidad equivalente potencial'],horizontal=True)
    rows=[]
    for _,r in productivity(base).iterrows():
        hrec=rec*r['NVA'];ht=r['HH']-hrec
        if ht<=0: continue
        if alternative.startswith('Producción'):
            qt=r['tanques'];hh=ht;note='Proyección de HH liberadas, no ahorro observado'
        else:
            qt=r['tanques']*r['HH']/ht;hh=r['HH'];note='Equivalentes continuos, NO tanques terminados'
        rows.append({'Capacidad gal':int(r['Capacidad gal']),'Tanques AS-IS':r['tanques'],'HH AS-IS':r['HH'],'Tanques TO-BE hipotéticos':qt,'HH TO-BE hipotéticas':hh,'Productividad AS-IS':r['tanques_HH'],'Productividad TO-BE hipotética':qt/hh,'Variación proyectada (%)':100*((qt/hh)/r['tanques_HH']-1),'HH NVA recuperables hipotéticas':hrec,'Interpretación':note})
    tobe=pd.DataFrame(rows)
    st.dataframe(tobe,use_container_width=True,hide_index=True)
    chart=tobe.melt(id_vars=['Capacidad gal'],value_vars=['Productividad AS-IS','Productividad TO-BE hipotética'],var_name='Estado',value_name='Tanques/HH')
    st.plotly_chart(px.bar(chart,x='Capacidad gal',y='Tanques/HH',color='Estado',barmode='group',title='Comparación por capacidad: observado vs hipotético'),use_container_width=True)
    st.warning('La variación porcentual es una MEJORA PROYECTADA por recuperación hipotética de pérdidas; NO es mejora real medida. El 40 % corresponde a recuperación de pérdidas, no a aumento directo de tanques/HH.')
    st.markdown('#### Sensibilidad de mejora proyectada por capacidad')
    sens=[]
    for recovery in [0,.10,.20,.30,.40]:
        for _,r in productivity(base).iterrows():
            hh_new=r['HH']-recovery*r['NVA']
            if hh_new>0 and r['HH']>0:
                sens.append({'Capacidad gal':int(r['Capacidad gal']),'Recuperación de NVA':f'{recovery:.0%}',
                    'Productividad AS-IS':r['tanques']/r['HH'],
                    'Productividad TO-BE proyectada':r['tanques']/hh_new,
                    'Mejora de productividad proyectada (%)':100*(r['HH']/hh_new-1)})
    sensitivity=pd.DataFrame(sens)
    st.dataframe(sensitivity,use_container_width=True,hide_index=True)
    st.plotly_chart(px.line(sensitivity,x='Recuperación de NVA',y='Mejora de productividad proyectada (%)',
        color='Capacidad gal',markers=True,title='Mejora de productividad proyectada, NO observada'),use_container_width=True)
    st.download_button('Exportar sensibilidad TO-BE',tobe.to_csv(index=False).encode('utf-8-sig'),'tobe_hipotetico.csv','text/csv')

# Una réplica = una ventana contigua de meses, con bootstrap conjunto por capacidad.
# No genera OP nuevas a partir de datos patológicos sin identificarlas.
def simulate(source,months,replicas,period,master,scenario,threshold):
    if len(months)<period: raise ValueError('Periodo solicitado mayor que el historial disponible')
    windows=[months[i:i+period] for i in range(len(months)-period+1)]
    children=np.random.SeedSequence(master).spawn(replicas)
    outputs=[]; samples=[]; seeds=[]
    for idx,child in enumerate(children,1):
        start=time.perf_counter();seed=int(child.generate_state(1)[0]);rng=np.random.default_rng(seed)
        win=windows[int(rng.integers(len(windows)))];ref=source[source['Mes'].isin(win)].copy()
        generated=[]
        for cap,g in ref.groupby('Capacidad gal'):
            pool=source[source['Capacidad gal']==cap]
            take=pool.iloc[rng.integers(0,len(pool),size=len(g))].reset_index(drop=True)
            v=g.reset_index(drop=True).copy()
            # Conserva el conjunto correlacionado de variables por OP.
            for col in ['HH fabricación','HH NVA','HH total OP','nva_reg','dist_reg','Cantidad']:
                v[col]=take[col].to_numpy()
            # Factores exploratorios, no estimaciones empresariales.
            if scenario=='Borde': factor=rng.choice([0.9,1.1])
            elif scenario=='Extremo': factor=rng.choice([0.7,1.3])
            else: factor=1.0
            if scenario in ['Borde','Extremo']:
                v['HH fabricación']*=factor;v['HH NVA']*=factor;v['HH total OP']*=factor
                v['nva_reg']*=factor;v['dist_reg']*=factor
            if scenario=='Patológico':
                # Prueba adversarial, excluir del cálculo hasta resolver inválidos.
                n=max(1,int(np.ceil(len(v)*0.02)))
                bad=rng.choice(len(v),size=min(n,len(v)),replace=False)
                v.loc[bad,'HH total OP']=np.nan
                if len(v)>2:v.loc[v.index[-1],'HH total OP']=0
            v['OP']=[f'SYN-{scenario[:3].upper()}-{idx:03d}-{int(cap)}-{j:04d}' for j in range(len(v))]
            v['replica']=idx;v['semilla']=seed;v['escenario']=scenario;v['origen']='SINTETICO';v['ventana']=';'.join(win)
            generated.append(v)
        z=pd.concat(generated,ignore_index=True)
        invalid=z['HH total OP'].isna()|(z['HH total OP']<=0)
        z['registro_invalido']=invalid
        valid=z.loc[~invalid]
        for cap,g in valid.groupby('Capacidad gal'):
            qty=g['Cantidad'].sum();hh=g['HH total OP'].sum();nva=g['HH NVA'].sum()
            outputs.append({'escenario':scenario,'replica':idx,'semilla':seed,'ventana':'|'.join(win),'capacidad_gal':int(cap),'OP_validas':len(g),'OP_invalidas':int(invalid.sum()),'tanques':qty,'HH':hh,'HH_NVA':nva,'productividad_ASIS_sintetica':qty/hh if hh>0 else np.nan,'origen':'SINTETICO','mejora_real_demostrada':False})
        samples.append(z)
        elapsed=time.perf_counter()-start
        seeds.append({'escenario':scenario,'replica':idx,'semilla':seed,'ventana':'|'.join(win),'tiempo_s':elapsed,'excede_umbral_provisional':elapsed>threshold})
    return pd.DataFrame(outputs),pd.concat(samples,ignore_index=True),pd.DataFrame(seeds)

with pages[3]:
    st.subheader('Generación de escenarios experimentales')
    st.info('Cada réplica abarca una ventana contigua de seis meses (configurable). El bootstrap preserva asociaciones entre variables de cada OP, pero no crea evidencia empresarial nueva.')
    scenarios=st.multiselect('Escenarios a generar',['Típico','Borde','Extremo','Patológico'],default=['Típico'])
    st.caption('Borde ±10 % y extremo ±30 % son perturbaciones exploratorias, no rangos validados. Patológico introduce HH faltantes/cero para verificar rechazo de entradas.')
    if st.button('Ejecutar réplicas y guardar evidencia',type='primary'):
        if not scenarios:st.warning('Selecciona al menos un escenario.')
        else:
            t0=time.perf_counter();outs=[];details=[];seedlist=[]
            try:
                for j,s in enumerate(scenarios):
                    o,d,sd=simulate(op,sorted(op['Mes'].unique()),int(nrep),int(period),int(seed_master)+100000*j,s,int(limit))
                    outs.append(o);details.append(d);seedlist.append(sd)
                st.session_state['exp_results']=pd.concat(outs,ignore_index=True)
                st.session_state['exp_details']=pd.concat(details,ignore_index=True)
                st.session_state['exp_seeds']=pd.concat(seedlist,ignore_index=True)
                st.session_state['exp_elapsed']=time.perf_counter()-t0
                st.success('Generación completada; resultados etiquetados SINTÉTICO.')
            except Exception as e:st.error(f'Generación interrumpida: {e}')
    if 'exp_results' in st.session_state:
        rr=st.session_state['exp_results'];ss=st.session_state['exp_seeds']
        st.metric('Duración total',f"{st.session_state['exp_elapsed']:.2f} s")
        if ss['excede_umbral_provisional'].any():st.error('Contingencia: una o más réplicas excedieron el umbral provisional.')
        st.dataframe(rr.head(100),use_container_width=True,hide_index=True)
        st.plotly_chart(px.box(rr,x='capacidad_gal',y='productividad_ASIS_sintetica',color='escenario',title='Variabilidad sintética por escenario y capacidad'),use_container_width=True)
        z=io.BytesIO()
        with pd.ExcelWriter(z,engine='xlsxwriter') as writer:
            rr.to_excel(writer,sheet_name='Resultados_replicas',index=False)
            st.session_state['exp_details'].to_excel(writer,sheet_name='OP_sinteticas',index=False)
            ss.to_excel(writer,sheet_name='Semillas_y_tiempos',index=False)
        st.download_button('Descargar evidencia experimental (.xlsx)',z.getvalue(),'DSS_V6_escenarios.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
with pages[4]:
    st.subheader('Auditoría y preparación de validación')
    mismatch=(base['HH fabricación']+base['HH NVA']-base['HH total OP']).abs()
    mismatch_nva=(base['HH NVA']-base['nva_reg']).abs()
    checks=pd.DataFrame([
        ['Unicidad de OP',bool(not op['OP'].duplicated().any()),'Registro histórico'],
        ['HH total OP = HH fabricación + HH NVA',bool((mismatch<0.02).all()),'Tolerancia 0,02 HH; [VERIFICAR] definición'],
        ['NVA por OP concilia',bool((mismatch_nva<0.02).all()),'Tolerancia 0,02 HH'],
        ['HH total OP positivas',bool((base['HH total OP']>0).all()),'Necesario para tanques/HH'],
        ['Datos Personal confirmados',False,'Hoja Personal marcada SUPUESTO'],
        ['IoT físico validado',False,'No hay dispositivo operativo'],
        ['Benchmarking DSS vs literatura ejecutado',False,'No hay método reproducible implementado'],
        ['Mejora TO-BE observada',False,'Solo escenarios hipotéticos'],
    ],columns=['Control','Cumple','Interpretación'])
    st.dataframe(checks,use_container_width=True,hide_index=True)
    st.markdown('**Comparación pendiente:** para afirmar competitividad, ejecutar el DSS y la línea base ingenua con las mismas OP, semillas, hardware y presupuesto de tiempo. El Pareto descriptivo no demuestra superioridad.')
    st.markdown('**Revisión del responsable:** documentar pertinencia de pérdidas, factibilidad de acciones, observaciones y firma/código del revisor. No sustituye mediciones posteriores.')
    st.markdown('**Protocolo:** 50 réplicas, 6 meses, 60 s provisional; calibración separada de evaluación si se ajustan parámetros. Umbrales y V2/V5 pendientes de aprobación del asesor.')
    st.download_button('Exportar auditoría',checks.to_csv(index=False).encode('utf-8-sig'),'auditoria_DSS_V6.csv','text/csv')
    st.caption(f'Entorno: Python {platform.python_version()} · pandas {pd.__version__} · numpy {np.__version__}. Registrar versiones del equipo donde se ejecuten las comparaciones.')

with pages[5]:
    st.subheader('Evidencia importada desde V5: 50 réplicas de seis meses')
    if v5.empty:
        st.warning('El archivo cargado no contiene RESULTADOS_V5. Carga el Excel integrado.')
    else:
        v5.columns=v5.columns.astype(str).str.strip()
        for c in ['productividad','recuperacion_supuesta','capacidad_gal','HH','tanques']:
            v5[c]=pd.to_numeric(v5[c],errors='coerce')
        st.info('Estas son proyecciones ya generadas en V5, NO corridas del DSS contra una línea base.')
        x=v5[(v5['alternativa']=='A_produccion_constante_HH_menores')].copy()
        caps_v5=sorted(x['capacidad_gal'].dropna().unique())
        cap_v5=st.selectbox('Capacidad a analizar (gal)',caps_v5,key='cap_v5')
        y=x[x['capacidad_gal']==cap_v5].copy()
        summary=y.groupby('recuperacion_supuesta',as_index=False).agg(
            media_productividad=('productividad','mean'),
            minimo=('productividad','min'),maximo=('productividad','max'),
            replicas=('replica','nunique'))
        p0=summary.loc[summary['recuperacion_supuesta']==0,'media_productividad']
        if len(p0) and p0.iloc[0]>0:
            summary['mejora_proyectada_pct']=(summary['media_productividad']/p0.iloc[0]-1)*100
        st.dataframe(summary,use_container_width=True,hide_index=True)
        st.plotly_chart(px.line(summary,x='recuperacion_supuesta',y='mejora_proyectada_pct',
            markers=True,title=f'Sensibilidad V5 de tanques/HH: {int(cap_v5):,} gal (hipotética)',
            labels={'recuperacion_supuesta':'Fracción recuperada de HH NVA',
                    'mejora_proyectada_pct':'Mejora proyectada (%)'}),use_container_width=True)
        if not v5_seeds.empty:
            st.metric('Semillas de evaluación documentadas',v5_seeds['semilla_evaluacion'].nunique()
                      if 'semilla_evaluacion' in v5_seeds else len(v5_seeds))
        st.download_button('Descargar resultados V5 originales',v5.to_csv(index=False).encode('utf-8-sig'),
            'V5_resultados_integrados.csv','text/csv')
        st.warning('No afirmar superioridad del DSS: no hay ejecución comparable de la línea base ni resultados empresariales posteriores a implementación.')
