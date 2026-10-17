"""DSS Productividad Industrial · V4 | Streamlit.
Archivo esperado: DSS_FINAL_PARA_STREAMLIT_V4.xlsx
"""
from pathlib import Path
import io
import math
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="DSS | Productividad industrial", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

NAVY="#13243B"; BLUE="#2363AC"; TEAL="#087F8C"; ORANGE="#E8A44A"; RED="#C94B50"
PALETTE=[BLUE,TEAL,"#6B78B6",ORANGE,"#7D91A6"]
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"],.stApp{font-family:Inter,Arial,sans-serif}
.stApp{background:#F4F7FB;color:#13243B}
.block-container{max-width:1480px;padding-top:1.3rem;padding-bottom:3rem}
[data-testid="stSidebar"]{background:#13243B}
[data-testid="stSidebar"] *{color:#EDF4FD!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"]{border-radius:12px}
[data-testid="stMetric"]{background:white;border:1px solid #E0E7F1;border-radius:16px;padding:17px 18px;box-shadow:0 4px 14px #1A33500A}
[data-testid="stMetricLabel"]{font-weight:650;color:#50657D}
[data-testid="stMetricValue"]{font-weight:800;color:#13243B}
[data-testid="stMetricDelta"]{font-weight:650}
h1,h2,h3{letter-spacing:-.5px}
.hero{padding:24px 28px;background:linear-gradient(105deg,#13243B,#245381);color:white;border-radius:18px;margin-bottom:22px}
.hero h1{color:white!important;font-size:2rem;margin:0 0 5px 0}
.hero p{color:#D6E5F4;margin:0;font-size:.95rem}
.section-hint{color:#657A90;font-size:.88rem;margin:-4px 0 16px}
.tag{display:inline-block;padding:4px 10px;background:#EAF1FB;color:#205D9A;border-radius:12px;font-size:.78rem;font-weight:700}
[data-testid="stAlert"]{border-radius:12px}
</style>""", unsafe_allow_html=True)

EXPECTED={
 "00_PARAMETROS":{"Parámetro","Valor"},
 "01_CALENDARIO":{"Periodo","Operarios","HH capacidad nominal"},
 "02_ASISTENCIA":{"Periodo","Operario","HH trabajadas","HH extras"},
 "03_ORDENES":{"Orden","Periodo","Galones","Unidades","Peso TE","Producción TE"},
 "04_PROCESOS":{"ID","Orden","Periodo","Proceso","HH estándar","HH observadas"},
 "05_PERDIDAS":{"Evento","Periodo","Herramienta","Causa","Proceso","HH perdidas"},
 "06_EQUIPOS":{"Periodo","Equipo","Proceso","Horas planificadas","Horas parada"},
 "07_RECORRIDOS":{"Área origen","Área destino","Viajes","Metros por viaje"},
 "10_MENSUAL":{"Periodo","TE","HH","HH pérdidas"},
 "12_SEGUIMIENTO":{"Fase","Intervención","TE observados","HH observadas"},
 "13_CAPACIDAD_MODELO":{"Proceso","Capacidad base TE/6m","Ganancia conservadora","Ganancia moderada","Ganancia optimista"},
 "15_FROM_TO_SIMULADO":{"Origen","Destino","Flujo","Viajes / mes","Distancia AS IS (m)","Distancia TO BE (m)"},
}
REQUIRED=["00_PARAMETROS","01_CALENDARIO","02_ASISTENCIA","03_ORDENES","05_PERDIDAS","13_CAPACIDAD_MODELO"]
NUMERIC={
 "01_CALENDARIO":["Operarios","HH capacidad nominal"],
 "02_ASISTENCIA":["HH teóricas","Factor asistencia","HH extras","HH trabajadas"],
 "03_ORDENES":["Galones","Unidades","Peso TE","Producción TE"],
 "04_PROCESOS":["HH estándar","HH observadas"],
 "05_PERDIDAS":["Duración h","Personas afectadas","HH perdidas"],
 "06_EQUIPOS":["Horas planificadas","Horas parada","MTBF","MTTR"],
 "13_CAPACIDAD_MODELO":["Capacidad base TE/6m","Ganancia conservadora","Ganancia moderada","Ganancia optimista"],
 "15_FROM_TO_SIMULADO":["Viajes / mes","Distancia AS IS (m)","Distancia TO BE (m)","Operarios involucrados","Minutos / viaje AS IS"],
}

def n(x, default=np.nan):
    try:
        if pd.isna(x) or str(x).strip()=="":return default
        return float(x)
    except (TypeError,ValueError):return default

def nice(x,d=2):return f"{x:,.{d}f}"

def param(tables,label,default=np.nan):
    df=tables.get("00_PARAMETROS",pd.DataFrame())
    if df.empty:return default
    hit=df.loc[df["Parámetro"].astype(str).str.strip().str.casefold()==label.casefold(),"Valor"]
    return n(hit.iloc[0],default) if len(hit) else default

def load_xlsx(blob):
    try:
        xl=pd.ExcelFile(io.BytesIO(blob),engine="openpyxl")
        if len(xl.sheet_names)>45:raise ValueError("Demasiadas hojas para el formato V4")
        tables={name:pd.read_excel(xl,sheet_name=name,engine="openpyxl") for name in xl.sheet_names}
        for name,df in tables.items():
            df.columns=[str(c).strip() for c in df.columns]
            if name in NUMERIC:
                for col in NUMERIC[name]:
                    if col in df:df[col]=pd.to_numeric(df[col],errors="coerce")
        return tables,None
    except Exception as exc:return None,str(exc)

def check_data(t):
    problems=[]
    def add(code,level,message,action):problems.append({"Código":code,"Nivel":level,"Detalle":message,"Acción sugerida":action})
    for sheet,cols in EXPECTED.items():
        if sheet not in t:
            add("E01" if sheet in REQUIRED else "A01","ERROR" if sheet in REQUIRED else "AVISO",f"Falta la hoja {sheet}","Utilizar la base V4 y no eliminar hojas")
        else:
            missing=cols-set(t[sheet].columns)
            if missing:add("E02","ERROR",f"{sheet}: faltan columnas {', '.join(sorted(missing))}","Restaurar los encabezados originales")
    if any(x["Nivel"]=="ERROR" for x in problems):return pd.DataFrame(problems)
    orders=t["03_ORDENES"].copy(); staff=t["02_ASISTENCIA"].copy(); events=t["05_PERDIDAS"].copy()
    if orders["Orden"].isna().any() or orders["Orden"].duplicated().any():add("E03","ERROR","Órdenes duplicadas o sin ID","Cada orden debe tener un identificador único")
    if events["Evento"].isna().any() or events["Evento"].duplicated().any():add("E04","ERROR","Eventos duplicados o sin ID","Revisar ID de eventos de pérdidas")
    if staff[["Periodo","Operario"]].duplicated().any():add("E05","ERROR","Asistencia duplicada por período y operario","Una fila por trabajador y mes")
    for sheet,cols in NUMERIC.items():
        if sheet not in t:continue
        for col in cols:
            if col not in t[sheet]:continue
            vals=t[sheet][col].dropna()
            if not np.isfinite(vals.to_numpy(dtype=float)).all():add("E06","ERROR",f"{sheet}/{col}: números no finitos","Eliminar valores infinitos")
            if (vals<0).any():add("E07","ERROR",f"{sheet}/{col}: valores negativos","Corregir cantidades, horas o tasas")
    if (orders["Unidades"].isna() | (orders["Unidades"]<=0)).any():add("E08","ERROR","Órdenes con unidades vacías o no positivas","Completar unidades fabricadas")
    if (staff["HH trabajadas"].isna() | (staff["HH trabajadas"]<=0)).any():add("E09","ERROR","Asistencia con HH faltantes o nulas","Revisar jornada y asistencia")
    if (staff["HH extras"].fillna(0)>staff["HH trabajadas"]).any():add("E10","ERROR","HH extras mayores que HH trabajadas","Corregir asistencia")
    if (staff["Factor asistencia"].dropna()>1).any():add("E11","ERROR","Factor de asistencia mayor que 100%","Usar valores de 0 a 1")
    if (events["HH perdidas"].isna()).any():add("E12","ERROR","Eventos con HH perdidas vacías","Completar duración y personal afectado")
    if "Duración h" in events and "Personas afectadas" in events:
        calculadas=events["Duración h"]*events["Personas afectadas"]
        diff=(calculadas-events["HH perdidas"]).abs()
        if (diff.dropna()>0.1).any():add("E13","ERROR",f"{(diff>0.1).sum()} eventos: duración × personas ≠ HH perdidas","Revisar cálculo de cada evento")
    periods_orders=set(orders["Periodo"].dropna().astype(str))
    periods_staff=set(staff["Periodo"].dropna().astype(str))
    if periods_orders!=periods_staff:add("E14","ERROR","Los períodos de órdenes y asistencia no coinciden","Usar los mismos meses en ambas hojas")
    missing_orders=orders["Orden"].dropna().astype(str)
    if "04_PROCESOS" in t and {"Orden","HH observadas"}.issubset(t["04_PROCESOS"].columns):
        pr=t["04_PROCESOS"]
        orphan=~pr["Orden"].astype(str).isin(set(missing_orders))
        if orphan.any():add("E15","ERROR",f"{int(orphan.sum())} procesos sin orden válida","Corregir referencias a órdenes")
    te_calc=(orders["Unidades"]*orders["Peso TE"]).sum()
    te_saved=orders["Producción TE"].sum()
    if abs(te_calc-te_saved)>max(.01,te_calc*.001):add("E16","ERROR","La producción TE no concilia con unidades × peso TE","Recalcular producción equivalente")
    hh=staff["HH trabajadas"].sum(); loss=events["HH perdidas"].sum()
    if loss>hh:add("E17","ERROR","Pérdidas superan las HH trabajadas","Revisar duplicidades y período")
    if "01_CALENDARIO" in t:
        cal=t["01_CALENDARIO"]
        nominal=cal["HH capacidad nominal"].sum()
        extra=staff["HH extras"].sum()
        if hh>nominal+extra+1:add("A02","AVISO","HH trabajadas superan capacidad nominal + extras","Revisar calendario, dotación y extras")
        if len(staff["Operario"].dropna().unique())!=45:add("A03","AVISO","La dotación no coincide con 45 operarios exclusivos","Confirmar cambios de personal")
    if events["Estado"].astype(str).str.contains("NO VERIFICADO|MODELO",case=False,na=False).any():add("A04","AVISO","Las pérdidas son modeladas/no verificadas","No presentarlas como pérdidas medidas en planta")
    if "06_EQUIPOS" in t and t["06_EQUIPOS"]["Horas parada"].notna().sum()==0:add("A05","AVISO","Faltan horas de parada de equipos para TPM","Registrar paradas reales antes de validar MTBF y MTTR")
    if "07_RECORRIDOS" in t and t["07_RECORRIDOS"]["Metros por viaje"].notna().sum()==0:add("A06","AVISO","No hay recorridos reales medidos","Usar From-To modelado solo como ilustración")
    if "13_CAPACIDAD_MODELO" in t and t["13_CAPACIDAD_MODELO"]["Estado"].astype(str).str.contains("MODELO",case=False,na=False).any():add("A07","AVISO","Capacidades de proceso no verificadas","Validar con cronometraje y disponibilidad")
    if "12_SEGUIMIENTO" in t:
        f=t["12_SEGUIMIENTO"]
        if f["TE observados"].notna().sum()==0:add("A08","AVISO","Sin resultados posteriores observados","Registrar mediciones antes/después para demostrar mejora continua")
    if "04_PROCESOS" in t:
        p=t["04_PROCESOS"]
        if (p["HH observadas"].sum()>hh*1.05):add("A09","AVISO","HH de procesos superan HH trabajadas; posible diferencia de alcance","No sumar procesos a HH de asistencia; conciliar por orden")
    return pd.DataFrame(problems,columns=["Código","Nivel","Detalle","Acción sugerida"])

def build_kpis(t):
    o=t["03_ORDENES"]; a=t["02_ASISTENCIA"]; e=t["05_PERDIDAS"]
    hh=float(a["HH trabajadas"].sum()); te=float(o["Producción TE"].sum())
    physical=float(o["Unidades"].sum()); losses=e.groupby(e["Herramienta"].astype(str).str.upper().str.strip())["HH perdidas"].sum().to_dict()
    return {"hh":hh,"te":te,"physical":physical,"p":te/hh,"losses":losses,"lost":sum(losses.values()),"ot":float(a["HH extras"].sum())}

def capacity(t,mode):
    c=t["13_CAPACIDAD_MODELO"].copy()
    column={"Conservador":"Ganancia conservadora","Moderado":"Ganancia moderada","Optimista":"Ganancia optimista"}.get(mode)
    if column:c["Mejora proceso"]=c[column].fillna(0)
    else:c["Mejora proceso"]=0.0
    c["Capacidad TO BE (TE)"]=c["Capacidad base TE/6m"]*(1+c["Mejora proceso"])
    return c

def simulate(k,t,reductions,conversion,mode,extra_demand=None):
    freed=sum(k["losses"].get(key,0)*float(reductions.get(key,0)) for key in k["losses"])
    usable=freed*conversion
    # Base conversion TE/HH is used only as an initial approximation.
    te_hh=usable*k["p"]
    cap=capacity(t,mode)
    bottleneck=float(cap["Capacidad TO BE (TE)"].min())
    slack=max(0,bottleneck-k["te"])
    extra=min(te_hh,slack)
    if extra_demand is not None:extra=min(extra,max(0,extra_demand))
    new_te=k["te"]+extra
    return {"freed":freed,"usable":usable,"extra":extra,"new_te":new_te,"gain":new_te/k["te"]-1,"cap":cap,"bottleneck":bottleneck,"hh_cap":te_hh,"slack":slack}

with st.sidebar:
    st.markdown("## 🏭 DSS INDUSTRIAL")
    st.caption("DECISION SUPPORT SYSTEM · V4")
    page=st.radio("Navegación",["Resumen ejecutivo","Productividad y pérdidas","Capacidad y procesos","Simulador de mejora","SLP y recorridos","TPM e IoT","Recomendaciones","Calidad de datos","Seguimiento"],label_visibility="collapsed")
    st.divider()
    uploaded=st.file_uploader("Cargar Excel final V4",type=["xlsx"],help="Solo DSS_FINAL_PARA_STREAMLIT_V4.xlsx. Otros archivos se validan y rechazan si no tienen la estructura requerida.")
    st.caption("Los datos cargados se procesan en la sesión. Evita subir información confidencial sin autorización.")
    st.markdown("<span class='tag'>MODO ACADÉMICO · SIMULACIÓN</span>",unsafe_allow_html=True)

st.markdown("<div class='hero'><h1>DSS · Productividad de mano de obra</h1><p>Fabricación de tanques de acero · 5S + SLP + TPM · Capacidad y mejora continua</p></div>",unsafe_allow_html=True)
if uploaded is None:
    st.info("**Carga el Excel final V4** desde el menú lateral. El sistema no utiliza cifras antiguas del Checkpoint 08 ni inventa resultados al faltar un archivo.")
    st.stop()
if uploaded.size>18*1024*1024:
    st.error("E00 · El archivo supera 18 MB. Revisa si contiene imágenes u hojas ajenas a la plantilla V4.");st.stop()
t,error=load_xlsx(uploaded.getvalue())
if error:
    st.error(f"E00 · No se pudo leer el Excel. Verifica que no esté dañado o protegido. Detalle: {error}");st.stop()
checks=check_data(t)
errors=checks[checks["Nivel"]=="ERROR"]
warns=checks[checks["Nivel"]=="AVISO"]
if len(errors):
    st.error(f"**Carga rechazada: {len(errors)} error(es) de estructura o consistencia.** Corrige el Excel y vuelve a subirlo.")
    st.dataframe(errors,use_container_width=True,hide_index=True)
    st.download_button("Descargar errores CSV",errors.to_csv(index=False).encode("utf-8-sig"),file_name="errores_carga_dss.csv",mime="text/csv")
    st.stop()
k=build_kpis(t)
meta=param(t,"Meta productividad",.10)
if not 0<=meta<=1:
    st.error("E18 · La meta de productividad en 00_PARAMETROS debe estar entre 0 y 1.");st.stop()
if len(warns):st.warning(f"**{len(warns)} advertencias de calidad de datos.** Consulta «Calidad de datos» para ver cómo corregirlas.")
st.caption(f"Base cargada: {uploaded.name} · {len(t)} hojas · {len(t['03_ORDENES'])} órdenes · {len(t['05_PERDIDAS'])} eventos · **datos modelados, no observaciones verificadas**")

if page=="Resumen ejecutivo":
    st.subheader("Visión ejecutiva · Línea base AS IS")
    a,b,c,d=st.columns(4)
    a.metric("Productividad equivalente",f"{k['p']:.5f} TE/HH")
    b.metric("Producción",f"{k['physical']:,.0f} tanques",f"{k['te']:.1f} TE",delta_color="off")
    c.metric("Horas-hombre trabajadas",nice(k["hh"],0))
    d.metric("HH de eventos",nice(k["lost"],1),f"{k['lost']/k['hh']:.1%} de HH",delta_color="off")
    st.markdown("### Productividad y meta")
    l,r=st.columns([1.4,1])
    with l:
        month=t["02_ASISTENCIA"].groupby("Periodo")["HH trabajadas"].sum().rename("HH")
        te=t["03_ORDENES"].groupby("Periodo")["Producción TE"].sum().rename("TE")
        monthly=pd.concat([te,month],axis=1).fillna(0).reset_index().sort_values("Periodo")
        monthly["Productividad TE/HH"]=monthly["TE"]/monthly["HH"].replace(0,np.nan)
        fig=px.line(monthly,x="Periodo",y="Productividad TE/HH",markers=True,title="Evolución mensual · AS IS modelado")
        fig.update_traces(line_color=TEAL,line_width=3,marker_size=9)
        fig.update_layout(yaxis_tickformat=".4f",margin=dict(l=15,r=15,t=50,b=10),height=330,plot_bgcolor="white")
        st.plotly_chart(fig,use_container_width=True)
    with r:
        target=k["p"]*(1+meta)
        st.metric("Meta +10% (o meta editable)",f"{target:.5f} TE/HH")
        st.metric("TE adicionales requeridos",nice(k["te"]*meta,2))
        st.metric("HH extras registradas",nice(k["ot"],1))
        st.info("La meta se calcula sobre la producción equivalente y las HH del archivo. No es una mejora ya observada.")
    l,r=st.columns([1,1])
    with l:
        loss_df=pd.DataFrame([{"Herramienta":x,"HH perdidas":y} for x,y in k["losses"].items()]).sort_values("HH perdidas")
        fig=px.bar(loss_df,x="HH perdidas",y="Herramienta",orientation="h",color="Herramienta",color_discrete_sequence=PALETTE,title="Pérdidas por herramienta")
        fig.update_layout(showlegend=False,margin=dict(l=10,r=10,t=45,b=10),height=320)
        st.plotly_chart(fig,use_container_width=True)
    with r:
        cap=capacity(t,"Conservador")
        fig=px.bar(cap,x="Proceso",y="Capacidad base TE/6m",color="Proceso",color_discrete_sequence=PALETTE,title="Capacidad modelada por proceso (TE / 6 meses)")
        fig.add_hline(y=k["te"],line_dash="dash",annotation_text="Producción AS IS")
        fig.update_layout(showlegend=False,margin=dict(l=10,r=10,t=45,b=10),height=320)
        st.plotly_chart(fig,use_container_width=True)
    layout=Path(__file__).with_name("layout_actual.png")
    if layout.exists():
        with st.expander("Ver plano actual de la planta"):
            st.image(str(layout),caption="Plano de referencia. Las rutas simuladas no equivalen a recorridos medidos.",use_container_width=True)

elif page=="Productividad y pérdidas":
    st.subheader("Diagnóstico de pérdidas operativas")
    e=t["05_PERDIDAS"].copy()
    a,b,c=st.columns(3)
    a.metric("Eventos registrados",len(e));b.metric("HH improductivas modeladas",nice(k["lost"],1));c.metric("Peso sobre HH",f"{k['lost']/k['hh']:.1%}")
    grouped=e.groupby(["Herramienta","Causa"],as_index=False)["HH perdidas"].sum().sort_values("HH perdidas",ascending=False)
    fig=px.bar(grouped,x="HH perdidas",y="Causa",color="Herramienta",orientation="h",title="Pareto de causas (HH)",color_discrete_sequence=PALETTE)
    fig.update_layout(yaxis={"categoryorder":"total ascending"},height=max(340,45*len(grouped)),margin=dict(l=10,r=10,t=50,b=15))
    st.plotly_chart(fig,use_container_width=True)
    st.dataframe(grouped,use_container_width=True,hide_index=True)
    st.warning("Las HH de eventos no deben sumarse a las HH trabajadas. Son una clasificación del tiempo dentro de la jornada y requieren control de solapamiento.")

elif page=="Capacidad y procesos":
    st.subheader("Capacidad productiva y restricciones")
    c=capacity(t,"Moderado")
    bottleneck=c.loc[c["Capacidad base TE/6m"].idxmin(),"Proceso"]
    a,b,c3=st.columns(3)
    a.metric("Cuello de botella modelado",str(bottleneck))
    b.metric("Capacidad base mínima",f"{c['Capacidad base TE/6m'].min():.1f} TE")
    c3.metric("Holgura teórica AS IS",f"{max(0,c['Capacidad base TE/6m'].min()-k['te']):.1f} TE")
    long=c.melt(id_vars="Proceso",value_vars=["Capacidad base TE/6m","Capacidad TO BE (TE)"],var_name="Escenario",value_name="TE")
    fig=px.bar(long,x="Proceso",y="TE",color="Escenario",barmode="group",color_discrete_sequence=[BLUE,TEAL],title="Capacidad AS IS vs. TO BE moderado · modelo")
    fig.add_hline(y=k["te"],line_dash="dash",annotation_text="Producción base")
    fig.update_layout(height=400,margin=dict(l=10,r=10,t=50,b=10))
    st.plotly_chart(fig,use_container_width=True)
    st.dataframe(c,use_container_width=True,hide_index=True)
    if "04_PROCESOS" in t:
        p=t["04_PROCESOS"].groupby("Proceso",as_index=False)[["HH estándar","HH observadas"]].sum()
        p["Brecha HH"]=p["HH observadas"]-p["HH estándar"]
        st.markdown("#### Horas de proceso y estándares provisionales")
        st.dataframe(p,use_container_width=True,hide_index=True)
        st.caption("Estas HH de proceso no se suman a las HH de asistencia. Su alcance debe conciliarse antes de atribuir diferencias a pérdidas.")
    st.warning("La capacidad de proceso está modelada. Para validarla se necesitan tiempos de ciclo, disponibilidad, mezcla de productos y restricciones reales.")

elif page=="Simulador de mejora":
    st.subheader("Simulador AS IS → TO BE")
    st.markdown("<p class='section-hint'>Escenarios de mejora de 5S, SLP y TPM, limitados por HH aprovechables y por el proceso más restrictivo.</p>",unsafe_allow_html=True)
    presets={"Conservador":{"5S":.20,"SLP":.20,"TPM":.20,"OTROS":0},"Moderado":{"5S":.35,"SLP":.35,"TPM":.40,"OTROS":.10},"Optimista":{"5S":.55,"SLP":.55,"TPM":.65,"OTROS":.20},"Personalizado":{"5S":.35,"SLP":.35,"TPM":.40,"OTROS":.10}}
    mode=st.selectbox("Escenario",list(presets),index=1)
    reductions={}
    with st.container(border=True):
        cols=st.columns(4)
        for i,key in enumerate(["5S","SLP","TPM","OTROS"]):
            reductions[key]=cols[i].slider(f"Reducción {key}",0,100,int(presets[mode][key]*100),5,key=f"reduction_{mode}_{key}")/100
        conversion=st.slider("Porcentaje de HH liberadas que pueden convertirse en producción",0,100,75,5)/100
        st.caption("El 75% inicial es un supuesto académico; no una tasa medida.")
        demand_toggle=st.checkbox("Limitar producción a demanda adicional documentada",value=False)
        demand=st.number_input("Demanda adicional (TE)",min_value=0.0,value=0.0,step=1.0,disabled=not demand_toggle) if demand_toggle else None
        gain_mode=mode if mode!="Personalizado" else st.selectbox("Supuesto de capacidad por proceso",["Conservador","Moderado","Optimista"],index=1)
    sim=simulate(k,t,reductions,conversion,gain_mode,demand)
    a,b,c,d=st.columns(4)
    a.metric("HH liberadas potenciales",nice(sim["freed"],1))
    b.metric("TE adicionales factibles en modelo",nice(sim["extra"],2))
    c.metric("Productividad TO BE",f"{sim['new_te']/k['hh']:.5f} TE/HH")
    d.metric("Mejora proyectada",f"{sim['gain']:.2%}")
    progress=min(1,sim["gain"]/meta) if meta else 1
    st.progress(progress,text=f"Avance de la simulación hacia meta {meta:.0%}: {progress:.0%}")
    if sim["gain"]+1e-9<meta:
        st.warning(f"A10 · La meta de {meta:.0%} NO se alcanza: escenario {sim['gain']:.2%}. Faltan {k['te']*meta-sim['extra']:.2f} TE para la meta.")
    else:st.success("La meta se alcanza matemáticamente en el escenario; aún requiere validación de capacidad, demanda y tiempos.")
    if sim["hh_cap"]<sim["slack"]:st.info("Restricción activa: HH liberadas convertibles en producción.")
    else:st.info("Restricción activa: capacidad del cuello de botella modelado.")
    if demand_toggle and demand==0:st.warning("A11 · Demanda adicional igual a cero: no se autoriza producción extra en el escenario.")
    if any(x>=.90 for x in reductions.values()):st.warning("A12 · Reducción de pérdidas ≥90%: hipótesis extrema que requiere justificación adicional.")
    l,r=st.columns(2)
    with l:
        fig=go.Figure(data=[go.Bar(x=["AS IS","TO BE"],y=[k["te"],sim["new_te"]],marker_color=[BLUE,TEAL],text=[nice(k["te"],1),nice(sim["new_te"],1)],textposition="outside")])
        fig.update_layout(title="Producción equivalente (TE)",height=320,margin=dict(l=10,r=10,t=50,b=10),yaxis_range=[0,max(k["te"],sim["new_te"])*1.18])
        st.plotly_chart(fig,use_container_width=True)
    with r:
        fig=px.bar(sim["cap"],x="Proceso",y="Capacidad TO BE (TE)",title="Restricción de capacidad TO BE",color="Proceso",color_discrete_sequence=PALETTE)
        fig.add_hline(y=sim["new_te"],line_dash="dash",annotation_text="Producción simulada")
        fig.update_layout(showlegend=False,height=320,margin=dict(l=10,r=10,t=50,b=10))
        st.plotly_chart(fig,use_container_width=True)
    st.caption("Modelo: TE extra = mínimo(HH recuperadas × conversión × productividad AS IS, holgura del cuello de botella y demanda cuando se activa). No es una medición ni una predicción validada.")
    export=pd.DataFrame([{"Escenario":mode,"TE_AS_IS":k["te"],"HH_AS_IS":k["hh"],"HH_liberadas_hipoteticas":sim["freed"],"TE_extra_simuladas":sim["extra"],"TE_TO_BE":sim["new_te"],"Mejora_proyectada":sim["gain"],"Meta":meta,"Estado":"SIMULADO_NO_VALIDADO"}])
    st.download_button("Descargar resultado del escenario",export.to_csv(index=False).encode("utf-8-sig"),file_name="escenario_dss_v4.csv",mime="text/csv")

elif page=="SLP y recorridos":
    st.subheader("SLP · Distribución y flujo de materiales")
    img=Path(__file__).with_name("layout_actual.png")
    if img.exists():st.image(str(img),caption="Plano actual de referencia, sin rutas georreferenciadas",use_container_width=True)
    else:st.info("A13 · No se encontró layout_actual.png junto a app.py. El resto del módulo funciona sin la imagen.")
    f=t.get("15_FROM_TO_SIMULADO")
    if f is not None:
        view=f[["Origen","Destino","Flujo","Viajes / mes","Distancia AS IS (m)","Distancia TO BE (m)"]].copy()
        view["Ahorro metros/mes"]=(view["Distancia AS IS (m)"]-view["Distancia TO BE (m)"])*view["Viajes / mes"]
        st.dataframe(view,use_container_width=True,hide_index=True)
        st.metric("Distancia ahorrada/mes (modelada)",nice(view["Ahorro metros/mes"].sum(),0)+" m")
    st.warning("Los recorridos From-To son hipotéticos. No deben dibujarse como rutas verificadas ni sumarse a las HH de eventos SLP sin conciliación.")

elif page=="TPM e IoT":
    st.subheader("TPM · Mantenimiento y monitoreo")
    eq=t.get("06_EQUIPOS",pd.DataFrame())
    if len(eq):
        n_measured=int(eq["Horas parada"].notna().sum())
        a,b=st.columns(2)
        a.metric("Registros de equipos",len(eq));b.metric("Registros con parada medida",n_measured)
        st.dataframe(eq,use_container_width=True,hide_index=True)
        if not n_measured:st.warning("A14 · No se pueden calcular MTBF ni MTTR reales: faltan fallas, horas operativas y tiempos de reparación medidos.")
    st.markdown("#### Panel de alertas IoT · Demostración")
    st.info("La planta no tiene sensores IoT instalados. Este panel permite ensayar reglas con lecturas introducidas manualmente; no representa monitoreo en tiempo real.")
    with st.container(border=True):
        col1,col2,col3=st.columns(3)
        equip=col1.selectbox("Equipo",sorted(eq["Equipo"].dropna().astype(str).unique()) if len(eq) else ["MAQ-01"])
        temp=col2.number_input("Temperatura de prueba (°C)",min_value=-20.0,max_value=200.0,value=55.0)
        vibration=col3.number_input("Vibración de prueba (mm/s)",min_value=0.0,max_value=50.0,value=2.5,step=.1)
        if temp>=85 or vibration>=7.1:st.error(f"ALERTA CRÍTICA SIMULADA · {equip}: detener evaluación y solicitar inspección técnica según procedimiento interno.")
        elif temp>=70 or vibration>=4.5:st.warning(f"ADVERTENCIA SIMULADA · {equip}: programar inspección y revisar condición del equipo.")
        else:st.success(f"SIN ALERTA SIMULADA · {equip}: lectura dentro de umbrales de demostración.")
        st.caption("Umbrales ilustrativos; no sustituyen los límites del fabricante, normas técnicas ni una evaluación de seguridad.")

elif page=="Recomendaciones":
    st.subheader("Motor de recomendaciones DSS")
    e=t["05_PERDIDAS"]
    priorities=e.groupby(["Herramienta","Causa"],as_index=False)["HH perdidas"].sum().sort_values("HH perdidas",ascending=False)
    actions={"5S":"Organizar herramientas, aplicar control visual y auditar tiempos de búsqueda.","SLP":"Medir From-To y rediseñar recorridos según relaciones entre áreas.","TPM":"Registrar fallas, priorizar equipos críticos y programar mantenimiento.","OTROS":"Estudiar esperas, preparación y retrabajos antes de asignar una herramienta."}
    priorities["Acción recomendada"]=priorities["Herramienta"].astype(str).str.upper().map(actions).fillna("Investigar causa y proponer acción validable")
    priorities.insert(0,"Prioridad",range(1,len(priorities)+1))
    priorities["Participación"]=(priorities["HH perdidas"]/k["lost"]*100).round(1)
    st.dataframe(priorities,use_container_width=True,hide_index=True)
    st.caption("Prioridad preliminar por HH modeladas; no sustituye factibilidad, costo ni juicio experto.")

elif page=="Calidad de datos":
    st.subheader("Control de calidad y trazabilidad")
    a,b,c=st.columns(3)
    a.metric("Errores bloqueantes",len(errors));b.metric("Advertencias",len(warns));c.metric("Hojas detectadas",len(t))
    if len(checks):st.dataframe(checks,use_container_width=True,hide_index=True)
    else:st.success("No se detectaron problemas automáticos. Esto no acredita verificación en planta.")
    st.download_button("Exportar informe de calidad",checks.to_csv(index=False).encode("utf-8-sig"),file_name="validaciones_dss_v4.csv",mime="text/csv")
    st.markdown("#### Inventario de hojas")
    st.dataframe(pd.DataFrame([{"Hoja":name,"Filas":len(df),"Columnas":len(df.columns)} for name,df in t.items()]),hide_index=True,use_container_width=True)
    st.warning("Una validación informática comprueba estructura y consistencia, no certifica que los datos modelados sean reales.")

elif page=="Seguimiento":
    st.subheader("Seguimiento de mejora continua · Resultados observados")
    f=t.get("12_SEGUIMIENTO",pd.DataFrame()).copy()
    if len(f):
        f["TE observados"]=pd.to_numeric(f["TE observados"],errors="coerce")
        f["HH observadas"]=pd.to_numeric(f["HH observadas"],errors="coerce")
        f["Productividad calculada"]=np.where(f["HH observadas"]>0,f["TE observados"]/f["HH observadas"],np.nan)
        f["Cambio vs AS IS"]=f["Productividad calculada"]/k["p"]-1
        st.dataframe(f,use_container_width=True,hide_index=True)
        measured=f.dropna(subset=["Productividad calculada"])
        if measured.empty:st.info("A15 · Todavía no hay resultados posteriores medidos. Completa 12_SEGUIMIENTO con TE, HH, fecha y evidencia de cada intervención.")
        else:
            fig=px.line(measured,x="Fase",y="Cambio vs AS IS",markers=True,title="Mejora observada vs. línea base")
            fig.add_hline(y=meta,line_dash="dash",annotation_text=f"Meta {meta:.0%}")
            fig.update_layout(yaxis_tickformat=".1%",height=350)
            st.plotly_chart(fig,use_container_width=True)
    st.caption("Una mejora continua demostrada exige datos de seguimiento reales, comparables y documentados.")
