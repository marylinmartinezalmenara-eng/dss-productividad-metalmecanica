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



def monthly_kpis(t):
    hh=t["02_ASISTENCIA"].groupby("Periodo")["HH trabajadas"].sum().rename("HH")
    te=t["03_ORDENES"].groupby("Periodo")["Producción TE"].sum().rename("TE")
    ev=t["05_PERDIDAS"].groupby("Periodo")["HH perdidas"].sum().rename("HH pérdidas")
    d=pd.concat([te,hh,ev],axis=1).fillna(0).reset_index().sort_values("Periodo")
    d["Productividad"]=d["TE"]/d["HH"].replace(0,np.nan)
    d["% pérdidas"]=d["HH pérdidas"]/d["HH"].replace(0,np.nan)
    return d

def demo_iot(t):
    # Datos deterministas de DEMOSTRACIÓN: nunca se atribuyen a sensores reales.
    eq=t.get("06_EQUIPOS",pd.DataFrame())
    names=sorted(eq["Equipo"].dropna().astype(str).unique()) if "Equipo" in eq else []
    if not names:names=[f"MAQ-{i:02d}" for i in range(1,5)]
    periods=sorted(t["03_ORDENES"]["Periodo"].dropna().astype(str).unique())
    if not periods:periods=["Muestra"]
    data=[]
    for i,name in enumerate(names):
        for j,period in enumerate(periods):
            # Curvas repetibles para comprobar alertas y filtros del prototipo.
            temp=round(51+3*i+2*j+ (13 if i==2 and j>=3 else 0),1)
            vib=round(2.1+.45*i+.24*j+(2.5 if i==2 and j>=3 else 0),2)
            level="CRÍTICA" if temp>=85 or vib>=7.1 else "ADVERTENCIA" if temp>=70 or vib>=4.5 else "NORMAL"
            data.append({"Periodo":period,"Equipo":name,"Temperatura °C":temp,"Vibración mm/s":vib,"Nivel":level,"Fuente":"SIMULADO"})
    return pd.DataFrame(data)

def get_iot(t):
    if "18_IOT_LECTURAS" not in t:return demo_iot(t),"SIMULADO"
    f=t["18_IOT_LECTURAS"].copy()
    required={"Periodo","Equipo","Temperatura °C","Vibración mm/s"}
    if not required.issubset(f.columns):
        st.warning("A-IOT01 · La hoja 18_IOT_LECTURAS no tiene columnas correctas; se muestra demostración simulada.")
        return demo_iot(t),"SIMULADO"
    for c in ["Temperatura °C","Vibración mm/s"]:f[c]=pd.to_numeric(f[c],errors="coerce")
    if f[["Temperatura °C","Vibración mm/s"]].isna().any().any():
        st.error("E-IOT02 · Existen lecturas IoT vacías o no numéricas. Se muestran datos de demostración para no ocultar el problema.")
        return demo_iot(t),"SIMULADO"
    if ((f["Temperatura °C"]< -20)|(f["Temperatura °C"]>200)|(f["Vibración mm/s"]<0)|(f["Vibración mm/s"]>50)).any():
        st.error("E-IOT03 · Lecturas fuera del rango permitido. Revisa temperatura y vibración.")
        return demo_iot(t),"SIMULADO"
    f["Nivel"]=np.select([(f["Temperatura °C"]>=85)|(f["Vibración mm/s"]>=7.1),(f["Temperatura °C"]>=70)|(f["Vibración mm/s"]>=4.5)],["CRÍTICA","ADVERTENCIA"],default="NORMAL")
    f["Fuente"]="ARCHIVO CARGADO · VERIFICAR ORIGEN"
    return f,"ARCHIVO"

def demo_followup(k,t,meta):
    # Proyección progresiva de adopción, NO resultados observados.
    presets=[("Línea base",0,0,0), ("5S",.18,0,0), ("5S + SLP",.28,.24,0), ("5S + SLP + TPM",.32,.30,.28),("Estabilización",.42,.40,.38)]
    rows=[]
    for phase,a,b,c in presets:
        freed=k["losses"].get("5S",0)*a+k["losses"].get("SLP",0)*b+k["losses"].get("TPM",0)*c
        usable=freed*.75
        # El escenario demuestra el efecto de HH liberadas en eficiencia,
        # sin asumir producción adicional sin capacidad/demanda demostradas.
        hh_effective=max(1,k["hh"]-usable)
        prod=k["te"]/hh_effective
        rows.append({"Etapa":phase,"HH recuperadas hipotéticas":round(usable,2),"TE":k["te"],"HH equivalentes":round(hh_effective,2),"Productividad TE/HH":prod,"Mejora %":prod/k["p"]-1,"Tipo":"PROYECCIÓN MODELADA"})
    return pd.DataFrame(rows)

with st.sidebar:
    st.markdown("## 🏭 DSS INDUSTRIAL")
    st.caption("DECISION SUPPORT SYSTEM · V4")
    page=st.radio("Navegación",["Indicadores integrales","Resumen ejecutivo","Productividad y pérdidas","Capacidad y procesos","Simulador de mejora","SLP y recorridos","TPM e IoT","Recomendaciones","Calidad de datos","Seguimiento"],label_visibility="collapsed")
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

if page=="Indicadores integrales":
    st.subheader("Centro de control · Indicadores integrales")
    st.caption("Vista consolidada del DSS | Línea base modelada, capacidad, pérdidas, alertas y oportunidades")
    monthly=monthly_kpis(t)
    iot,source=get_iot(t)
    cap=capacity(t,"Moderado")
    bottleneck=cap.loc[cap["Capacidad base TE/6m"].idxmin(),"Proceso"]
    cols=st.columns(6)
    for col,(label,value) in zip(cols,[("Productividad física",f"{k['physical']/k['hh']:.5f} tanques/HH"),("Productividad",f"{k['p']:.5f} TE/HH"),("Producción",f"{k['te']:.1f} TE"),("HH trabajadas",nice(k['hh'],0)),("Pérdidas",f"{k['lost']/k['hh']:.1%}"),("Meta",f"+{meta:.0%}")]):col.metric(label,value)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Órdenes",len(t["03_ORDENES"]))
    c2.metric("Eventos de pérdida",len(t["05_PERDIDAS"]))
    c3.metric("Cuello de botella",str(bottleneck))
    c4.metric("Equipos en alerta (última lectura)",int(iot.sort_values("Periodo").groupby("Equipo").tail(1)["Nivel"].isin(["CRÍTICA","ADVERTENCIA"]).sum()))
    tab0,tab1,tab2,tab3=st.tabs(["AS IS vs. TO BE","Producción y productividad","Pérdidas y capacidad","Alertas y decisiones"])
    with tab0:
        st.markdown("### Comparativo de productividad y capacidad · AS IS → TO BE")
        st.caption("Escenario moderado de referencia: 5S 35%, SLP 35%, TPM 40% y OTROS 10%. Las reducciones son supuestos modelados, no resultados observados.")
        integrated_reductions={"5S":.35,"SLP":.35,"TPM":.40,"OTROS":.10}
        integrated_sim=simulate(k,t,integrated_reductions,.75,"Moderado")
        x1,x2,x3,x4=st.columns(4)
        x1.metric("Productividad AS IS",f"{k['p']:.5f} TE/HH")
        x2.metric("Productividad TO BE",f"{integrated_sim['new_te']/k['hh']:.5f} TE/HH",f"+{integrated_sim['gain']:.2%}")
        x3.metric("Producción AS IS",f"{k['te']:.1f} TE")
        x4.metric("Producción TO BE",f"{integrated_sim['new_te']:.1f} TE",f"+{integrated_sim['extra']:.1f} TE")
        st.markdown("#### Productividad física · tanques/HH")
        fis1,fis2,fis3=st.columns(3)
        fis1.metric("AS IS · tanques/HH",f"{k['physical']/k['hh']:.5f}")
        # Para estimar unidades físicas adicionales se mantiene la mezcla histórica de tamaños.
        physical_to_be=k['physical']*(integrated_sim['new_te']/k['te'])
        fis2.metric("TO BE · tanques/HH",f"{physical_to_be/k['hh']:.5f}",f"+{integrated_sim['gain']:.2%}")
        fis3.metric("Tanques adicionales equivalentes en mezcla",f"+{physical_to_be-k['physical']:.2f}")
        st.caption("Tanques TO BE estimados manteniendo la mezcla de productos actual. Es una equivalencia modelada, no unidades enteras comprometidas de fabricación.")
        y1,y2,y3=st.columns(3)
        y1.metric("HH trabajadas (mismo recurso)",f"{k['hh']:,.1f} HH")
        y2.metric("HH potencialmente recuperadas",f"{integrated_sim['freed']:,.1f} HH")
        y3.metric("Meta de productividad",f"+{meta:.0%}")
        left_compare,right_compare=st.columns(2)
        with left_compare:
            comparison=pd.DataFrame({"Escenario":["AS IS","TO BE modelado","Meta 10%"],"Producción TE":[k['te'],integrated_sim['new_te'],k['te']*(1+meta)]})
            fig=px.bar(comparison,x="Escenario",y="Producción TE",color="Escenario",text="Producción TE",title="Producción equivalente: actual, proyectada y objetivo",color_discrete_sequence=[BLUE,TEAL,ORANGE])
            fig.update_traces(texttemplate="%{text:.1f}",textposition="outside")
            fig.update_layout(showlegend=False,height=350,margin=dict(l=10,r=10,t=55,b=10),yaxis_range=[0,max(comparison['Producción TE'])*1.18])
            st.plotly_chart(fig,use_container_width=True)
        with right_compare:
            prod_compare=pd.DataFrame({"Escenario":["AS IS","TO BE modelado","Meta 10%"],"Productividad TE/HH":[k['p'],integrated_sim['new_te']/k['hh'],k['p']*(1+meta)]})
            fig=px.bar(prod_compare,x="Escenario",y="Productividad TE/HH",color="Escenario",text="Productividad TE/HH",title="Productividad: AS IS frente a TO BE",color_discrete_sequence=[BLUE,TEAL,ORANGE])
            fig.update_traces(texttemplate="%{text:.5f}",textposition="outside")
            fig.update_layout(showlegend=False,height=350,margin=dict(l=10,r=10,t=55,b=10),yaxis_range=[0,max(prod_compare['Productividad TE/HH'])*1.18])
            st.plotly_chart(fig,use_container_width=True)
        physical_compare=pd.DataFrame({"Escenario":["AS IS","TO BE modelado","Meta"],"Tanques/HH":[k['physical']/k['hh'],physical_to_be/k['hh'],k['physical']/k['hh']*(1+meta)]})
        fig=px.bar(physical_compare,x="Escenario",y="Tanques/HH",color="Escenario",text="Tanques/HH",title="Productividad física · tanques/HH",color_discrete_sequence=[BLUE,TEAL,ORANGE])
        fig.update_traces(texttemplate="%{text:.5f}",textposition="outside")
        fig.update_layout(showlegend=False,height=300,margin=dict(l=10,r=10,t=55,b=10),yaxis_range=[0,physical_compare['Tanques/HH'].max()*1.2])
        st.plotly_chart(fig,use_container_width=True)
        if integrated_sim['gain'] < meta:
            st.warning(f"A10 · Escenario moderado: mejora proyectada {integrated_sim['gain']:.2%}; aún faltan {max(0,k['te']*meta-integrated_sim['extra']):.2f} TE para alcanzar la meta del {meta:.0%}. Revisa el simulador para cambiar los supuestos.")
        else:
            st.success("La meta se alcanza matemáticamente en este escenario modelado; requiere validación de demanda, capacidad y tiempos en planta.")
        st.caption("Criterio: se mantienen las HH trabajadas y se convierte una parte de las HH liberadas en TE adicionales, limitada por la capacidad del cuello de botella. No sumar además una reducción de HH al denominador: duplicaría el beneficio.")
    with tab1:
        left,right=st.columns([1.6,1])
        with left:
            fig=px.line(monthly,x="Periodo",y="Productividad",markers=True,title="Productividad mensual · TE/HH",color_discrete_sequence=[TEAL])
            fig.add_hline(y=k["p"]*(1+meta),line_dash="dash",annotation_text="Meta",line_color=ORANGE)
            fig.update_layout(height=330,margin=dict(l=10,r=10,t=50,b=10),yaxis_tickformat=".5f")
            st.plotly_chart(fig,use_container_width=True)
        with right:
            fig=px.bar(monthly,x="Periodo",y="TE",title="Producción equivalente por mes",color_discrete_sequence=[BLUE])
            fig.update_layout(height=330,margin=dict(l=10,r=10,t=50,b=10))
            st.plotly_chart(fig,use_container_width=True)
        st.dataframe(monthly.style.format({"HH":"{:,.1f}","TE":"{:,.1f}","Productividad":"{:.5f}","HH pérdidas":"{:,.1f}","% pérdidas":"{:.1%}"}),use_container_width=True,hide_index=True)
    with tab2:
        l,r=st.columns(2)
        with l:
            e=t["05_PERDIDAS"].groupby("Herramienta",as_index=False)["HH perdidas"].sum().sort_values("HH perdidas")
            fig=px.bar(e,x="HH perdidas",y="Herramienta",orientation="h",title="Pérdidas por herramienta (HH)",color="Herramienta",color_discrete_sequence=PALETTE)
            fig.update_layout(height=350,showlegend=False,margin=dict(l=10,r=10,t=50,b=10))
            st.plotly_chart(fig,use_container_width=True)
        with r:
            fig=px.bar(cap,x="Proceso",y="Capacidad base TE/6m",title="Capacidad modelada por proceso",color="Proceso",color_discrete_sequence=PALETTE)
            fig.add_hline(y=k["te"],line_dash="dash",annotation_text="Producción base")
            fig.update_layout(height=350,showlegend=False,margin=dict(l=10,r=10,t=50,b=10))
            st.plotly_chart(fig,use_container_width=True)
        st.info(f"La capacidad mínima modelada corresponde a {bottleneck}. No representa capacidad verificada en planta.")
    with tab3:
        l,r=st.columns([1.1,1])
        with l:
            last=iot.sort_values("Periodo").groupby("Equipo").tail(1).sort_values("Equipo")
            st.markdown("**Estado por máquina · Última lectura**")
            st.dataframe(last[["Equipo","Periodo","Temperatura °C","Vibración mm/s","Nivel"]],hide_index=True,use_container_width=True)
            st.caption(f"Origen IoT: {source}. Umbrales ilustrativos, no límites validados por fabricante.")
        with r:
            priority=t["05_PERDIDAS"].groupby(["Herramienta","Causa"],as_index=False)["HH perdidas"].sum().nlargest(4,"HH perdidas")
            st.markdown("**Principales causas para actuar**")
            st.dataframe(priority,hide_index=True,use_container_width=True)
            if len(warns):st.warning(f"Hay {len(warns)} advertencias de calidad de datos. Consulta Calidad de datos.")
    with st.expander("Plano de planta · Distribución actual"):
        img=Path(__file__).with_name("layout_actual.png")
        if img.exists():st.image(str(img),use_container_width=True)
        else:st.warning("A13 · Falta layout_actual.png en el repositorio.")
    st.caption("Todas las cifras provienen del Excel cargado, salvo el panel IoT cuando se identifica como SIMULADO.")

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
    physical_to_be=k['physical']*sim['new_te']/k['te']
    ph1,ph2,ph3=st.columns(3)
    ph1.metric("Productividad física AS IS",f"{k['physical']/k['hh']:.5f} tanques/HH")
    ph2.metric("Productividad física TO BE",f"{physical_to_be/k['hh']:.5f} tanques/HH",f"+{sim['gain']:.2%}")
    ph3.metric("Producción física estimada TO BE",f"{physical_to_be:.2f} tanques")
    st.caption("La estimación física conserva la mezcla de tanques del período. La producción real debe expresarse en unidades enteras y verificarse por tipo de tanque.")
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
    export=pd.DataFrame([{"Escenario":mode,"TE_AS_IS":k["te"],"HH_AS_IS":k["hh"],"HH_liberadas_hipoteticas":sim["freed"],"TE_extra_simuladas":sim["extra"],"TE_TO_BE":sim["new_te"],"Tanques_AS_IS":k["physical"],"Tanques_TO_BE_modelados":physical_to_be,"Tanques_por_HH_AS_IS":k["physical"]/k["hh"],"Tanques_por_HH_TO_BE":physical_to_be/k["hh"],"Mejora_proyectada":sim["gain"],"Meta":meta,"Estado":"SIMULADO_NO_VALIDADO"}])
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
    st.subheader("Monitoreo TPM · Alertas por máquina")
    iot,source=get_iot(t)
    st.info(f"Origen de las lecturas: **{source}**. Si dice SIMULADO, son ejemplos reproducibles para probar el DSS; la planta no cuenta con sensores conectados.")
    last=iot.sort_values("Periodo").groupby("Equipo").tail(1).sort_values("Equipo")
    a,b,c,d=st.columns(4)
    a.metric("Máquinas",len(last))
    b.metric("Críticas",int((last["Nivel"]=="CRÍTICA").sum()))
    c.metric("En advertencia",int((last["Nivel"]=="ADVERTENCIA").sum()))
    d.metric("Sin alerta",int((last["Nivel"]=="NORMAL").sum()))
    for _,r in last.iterrows():
        msg=f"{r['Equipo']} · {r['Temperatura °C']:.1f} °C · {r['Vibración mm/s']:.2f} mm/s · {r['Periodo']}"
        if r["Nivel"]=="CRÍTICA":st.error("🔴 CRÍTICA · "+msg+" · Revisar condición con mantenimiento.")
        elif r["Nivel"]=="ADVERTENCIA":st.warning("🟠 ADVERTENCIA · "+msg+" · Programar inspección.")
        else:st.success("🟢 NORMAL · "+msg)
    equipment=st.selectbox("Consultar historial de una máquina",last["Equipo"].tolist())
    history=iot.loc[iot["Equipo"]==equipment].sort_values("Periodo")
    left,right=st.columns(2)
    with left:
        fig=px.line(history,x="Periodo",y="Temperatura °C",markers=True,title=f"{equipment} · Temperatura",color_discrete_sequence=[ORANGE])
        fig.add_hline(y=70,line_dash="dash",line_color=ORANGE,annotation_text="Advertencia")
        fig.add_hline(y=85,line_dash="dash",line_color=RED,annotation_text="Crítica")
        fig.update_layout(height=310,margin=dict(l=10,r=10,t=50,b=10));st.plotly_chart(fig,use_container_width=True)
    with right:
        fig=px.line(history,x="Periodo",y="Vibración mm/s",markers=True,title=f"{equipment} · Vibración",color_discrete_sequence=[BLUE])
        fig.add_hline(y=4.5,line_dash="dash",line_color=ORANGE,annotation_text="Advertencia")
        fig.add_hline(y=7.1,line_dash="dash",line_color=RED,annotation_text="Crítica")
        fig.update_layout(height=310,margin=dict(l=10,r=10,t=50,b=10));st.plotly_chart(fig,use_container_width=True)
    st.dataframe(iot,use_container_width=True,hide_index=True)
    st.download_button("Descargar lecturas y alertas",iot.to_csv(index=False).encode("utf-8-sig"),"alertas_iot_dss.csv","text/csv")
    st.markdown("#### Prueba manual de alertas")
    with st.container(border=True):
        col1,col2,col3=st.columns(3)
        machine=col1.selectbox("Máquina de prueba",last["Equipo"].tolist(),key="iot_test_machine")
        temp=col2.number_input("Temperatura °C",min_value=-20.,max_value=200.,value=55.,step=1.)
        vibration=col3.number_input("Vibración mm/s",min_value=0.,max_value=50.,value=2.5,step=.1)
        if temp>=85 or vibration>=7.1:st.error(f"CRÍTICA SIMULADA · {machine}: inspección prioritaria.")
        elif temp>=70 or vibration>=4.5:st.warning(f"ADVERTENCIA SIMULADA · {machine}: inspección programada.")
        else:st.success(f"NORMAL SIMULADO · {machine}")
    eq=t.get("06_EQUIPOS",pd.DataFrame())
    if len(eq) and eq["Horas parada"].notna().sum()==0:st.warning("A14 · MTBF y MTTR reales no calculables: faltan paradas y reparaciones medidas.")
    st.caption("Los umbrales de 70/85 °C y 4.5/7.1 mm/s son demostrativos; requieren especificaciones técnicas y validación por equipo.")

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
    st.subheader("Seguimiento · Mejora continua")
    st.caption("Dos vistas separadas: proyección académica ilustrativa y resultados observados cargados en Excel.")
    view=st.radio("Tipo de seguimiento",["Proyección de demostración","Resultados observados"],horizontal=True)
    if view=="Proyección de demostración":
        st.warning("**SERIE SIMULADA.** Los valores representan hipótesis progresivas de adopción 5S, SLP y TPM; no son resultados de implementación real.")
        f=demo_followup(k,t,meta)
        a,b,c=st.columns(3)
        a.metric("Productividad base",f"{k['p']:.5f} TE/HH")
        b.metric("Mejora final modelada",f"{f['Mejora %'].iloc[-1]:.2%}")
        c.metric("Meta de investigación",f"{meta:.0%}")
        fig=px.line(f,x="Etapa",y="Mejora %",markers=True,title="Trayectoria hipotética de mejora",color_discrete_sequence=[TEAL])
        fig.add_hline(y=meta,line_dash="dash",line_color=ORANGE,annotation_text=f"Meta {meta:.0%}")
        fig.update_layout(yaxis_tickformat=".1%",height=390,margin=dict(l=10,r=10,t=50,b=10))
        st.plotly_chart(fig,use_container_width=True)
        st.dataframe(f.style.format({"HH recuperadas hipotéticas":"{:,.1f}","TE":"{:,.1f}","HH equivalentes":"{:,.1f}","Productividad TE/HH":"{:.5f}","Mejora %":"{:.2%}"}),use_container_width=True,hide_index=True)
        st.download_button("Descargar proyección (SIMULADA)",f.to_csv(index=False).encode("utf-8-sig"),"seguimiento_simulado.csv","text/csv")
        st.info("Este modelo muestra productividad equivalente con HH netas tras mejoras. No debe sumarse a producción adicional en otro escenario: se debe elegir un criterio de productividad y conciliar la capacidad.")
    else:
        f=t.get("12_SEGUIMIENTO",pd.DataFrame()).copy()
        if len(f):
            for col in ["TE observados","HH observadas"]:f[col]=pd.to_numeric(f[col],errors="coerce")
            f["Productividad calculada"]=np.where(f["HH observadas"]>0,f["TE observados"]/f["HH observadas"],np.nan)
            f["Cambio vs AS IS"]=f["Productividad calculada"]/k["p"]-1
            st.dataframe(f,use_container_width=True,hide_index=True)
            measured=f.dropna(subset=["Productividad calculada"])
            if measured.empty:st.info("A15 · No hay mediciones posteriores registradas. Completa 12_SEGUIMIENTO para mostrar resultados reales.")
            else:
                fig=px.line(measured,x="Fase",y="Cambio vs AS IS",markers=True,title="Mejora observada")
                fig.add_hline(y=meta,line_dash="dash",annotation_text=f"Meta {meta:.0%}")
                fig.update_layout(yaxis_tickformat=".1%",height=350)
                st.plotly_chart(fig,use_container_width=True)
        else:st.warning("A16 · No existe la hoja 12_SEGUIMIENTO.")
