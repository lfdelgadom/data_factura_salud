import io, re, csv, unicodedata
from datetime import date
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats
import streamlit as st

st.set_page_config(page_title='Analítica de Facturación en Salud', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
CAMPOS=['ID_Factura','Numero_Radicado','Fecha_Servicio','Fecha_Emision','Fecha_Radicacion','Fecha_Aceptacion','Fecha_Pago','Periodo_Cierre','Codigo_Entidad','Entidad_Pagadora','Regimen','Sede','Codigo_Servicio','Servicio','Linea_Servicio','Cantidad','Canal_Radicacion','Estado_Factura','Dias_Emision_Radicacion','Soporte_Completo','Valor_Bruto','Descuento','Copago_Cuota','Valor_Neto','Glosa_Inicial','Valor_Pagado','Saldo_Cartera']
FECHAS=['Fecha_Servicio','Fecha_Emision','Fecha_Radicacion','Fecha_Aceptacion','Fecha_Pago']
DINERO=['Valor_Bruto','Descuento','Copago_Cuota','Valor_Neto','Glosa_Inicial','Valor_Pagado','Saldo_Cartera']
MINIMAS={'ID_Factura','Valor_Neto'}; RECOMENDADAS={'Entidad_Pagadora','Estado_Factura','Valor_Pagado','Saldo_Cartera','Glosa_Inicial'}
MESES={1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
PAGINAS=['Inicio y carga de datos','Calidad de datos','Resumen ejecutivo','Cierre mensual','Entidades pagadoras','Servicios y líneas','Sedes y regímenes','Estados de facturación','Glosas','Recaudo y cartera','Tiempos de radicación','Análisis estadístico','Valores atípicos','Insights y alertas','Explorador de datos','Descargas','Metodología']

def estilo():
 st.markdown("<style>.block-container{padding-top:1.2rem;max-width:1500px}.hero{background:linear-gradient(110deg,#083b66,#087e8b);color:white;padding:1.2rem 1.5rem;border-radius:14px}.notice{background:#eef6fb;border-left:5px solid #1677a6;padding:.8rem;border-radius:8px}.privacy{background:#fff5e8;border-left:5px solid #e67e22;padding:.8rem;border-radius:8px}</style>",unsafe_allow_html=True)
 st.markdown("<div class='hero'><h1>ANALÍTICA DE FACTURACIÓN EN SALUD</h1><p>Facturación, radicación, glosas, recaudo y cartera</p></div>",unsafe_allow_html=True)
 st.markdown("<div class='notice'>Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.</div>",unsafe_allow_html=True)

def norm(x):
 s=unicodedata.normalize('NFKD',str(x).strip()).encode('ascii','ignore').decode().lower(); s=re.sub(r'[-\s]+','_',s); return re.sub(r'[^a-z0-9_]','',s).strip('_')
ALIAS={norm(c):c for c in CAMPOS}; ALIAS.update({'numero_de_radicado':'Numero_Radicado','numero_factura':'ID_Factura','factura':'ID_Factura','linea_de_servicio':'Linea_Servicio','copago':'Copago_Cuota','valor_glosa':'Glosa_Inicial','saldo':'Saldo_Cartera','estado':'Estado_Factura'})

def moneda_serie(s):
 if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s,errors='coerce'),0,False
 ambiguo=False
 def f(v):
  nonlocal ambiguo
  if pd.isna(v) or str(v).strip()=='': return np.nan
  x=re.sub(r'[^0-9,.\-]','',str(v))
  if ',' in x and '.' in x:
   x=x.replace('.','').replace(',','.') if x.rfind(',')>x.rfind('.') else x.replace(',','')
  elif ',' in x:
   p=x.split(','); x=''.join(p[:-1])+'.'+p[-1] if len(p[-1]) in (1,2) else ''.join(p)
  elif x.count('.')>1: x=x.replace('.','')
  elif '.' in x and len(x.split('.')[-1])==3: ambiguo=True; x=x.replace('.','')
  try:return float(x)
  except:return np.nan
 out=s.map(f); return out,int(s.notna().sum()-out.notna().sum()),ambiguo

@st.cache_data(show_spinner=False)
def excel_hojas(b): return pd.ExcelFile(io.BytesIO(b)).sheet_names
@st.cache_data(show_spinner=False)
def leer_excel(b,h): return pd.read_excel(io.BytesIO(b),sheet_name=h)
@st.cache_data(show_spinner=False)
def leer_csv(b,modo):
 seps={'Automático':None,'Coma':',','Punto y coma':';','Tabulación':'\t'}
 for enc in ['utf-8-sig','utf-8','latin-1','cp1252']:
  try:
   sep=seps[modo]
   if sep is None: sep=csv.Sniffer().sniff(b[:10000].decode(enc),delimiters=',;\t|').delimiter
   return pd.read_csv(io.BytesIO(b),sep=sep,encoding=enc),enc,sep
  except: pass
 raise ValueError('No se detectó codificación o separador válido.')

def sugerencias(cols):
 d={}; eq=[]
 for c in cols:
  destino=ALIAS.get(norm(c))
  if destino and destino not in d: d[destino]=c; eq.append((c,destino)) if c!=destino else None
 return d,eq

def mapeo_ui(df,sug):
 labels={'ID_Factura':'Identificador de factura','Fecha_Servicio':'Fecha de servicio','Fecha_Emision':'Fecha de emisión','Fecha_Radicacion':'Fecha de radicación','Periodo_Cierre':'Periodo de cierre','Entidad_Pagadora':'Entidad pagadora','Regimen':'Régimen','Sede':'Sede','Servicio':'Servicio','Linea_Servicio':'Línea de servicio','Cantidad':'Cantidad','Valor_Bruto':'Valor bruto','Descuento':'Descuento','Copago_Cuota':'Copago','Valor_Neto':'Valor neto','Estado_Factura':'Estado','Dias_Emision_Radicacion':'Días de radicación','Soporte_Completo':'Soporte completo','Glosa_Inicial':'Glosa','Fecha_Pago':'Fecha de pago','Valor_Pagado':'Valor pagado','Saldo_Cartera':'Saldo de cartera'}
 with st.expander('Configurar correspondencia de columnas',expanded=False):
  opts=['Columna no disponible']+list(df.columns); m={}; cols=st.columns(2)
  for i,(dest,lab) in enumerate(labels.items()):
   old=st.session_state.get('mapa',{}).get(dest,sug.get(dest,'Columna no disponible')); val=cols[i%2].selectbox(lab,opts,index=opts.index(old) if old in opts else 0,key='m_'+dest)
   if val!='Columna no disponible':m[dest]=val
  dup=pd.Series(list(m.values())).duplicated().any()
  if dup: st.error('Una columna de origen está asociada a más de un campo.')
  if st.button('Aplicar correspondencia',type='primary',disabled=dup): st.session_state.mapa=m; st.rerun()
 return st.session_state.get('mapa',sug)

def convertir(df):
 d=df.copy(); rep=[]
 for c in FECHAS:
  if c in d:
   a=d[c].notna().sum(); d[c]=pd.to_datetime(d[c],errors='coerce',dayfirst=True); rep.append([c,'Fecha',int(a-d[c].notna().sum()),int(d[c].isna().sum())])
 for c in DINERO+['Cantidad','Dias_Emision_Radicacion']:
  if c in d:
   d[c],fallos,amb=moneda_serie(d[c]); rep.append([c,'Numérico',fallos,int(d[c].isna().sum())]); st.warning(f'Formato ambiguo detectado en {c}; revise una muestra.') if amb else None
 return d,pd.DataFrame(rep,columns=['Columna','Conversión','No convertidos','Nulos finales'])

def derivadas(d):
 d=d.copy()
 if 'Fecha_Radicacion' in d:
  f=d.Fecha_Radicacion; d['Año_Radicacion']=f.dt.year; d['Numero_Mes']=f.dt.month; d['Mes_Radicacion']=f.dt.month.map(MESES); d['Trimestre_Radicacion']='T'+f.dt.quarter.astype('Int64').astype(str); d['Año_Mes']=f.dt.to_period('M').astype(str); d['Dia_Semana']=f.dt.dayofweek.map({0:'Lunes',1:'Martes',2:'Miércoles',3:'Jueves',4:'Viernes',5:'Sábado',6:'Domingo'}); d['Periodo_Cierre']=d.get('Periodo_Cierre',d.Año_Mes)
 if {'Fecha_Emision','Fecha_Radicacion'}<=set(d): d['Dias_Emision_Radicacion']=(d.Fecha_Radicacion-d.Fecha_Emision).dt.days
 if {'Fecha_Radicacion','Fecha_Pago'}<=set(d): d['Dias_Radicacion_Pago']=(d.Fecha_Pago-d.Fecha_Radicacion).dt.days
 for out,num,den in [('Porcentaje_Descuento','Descuento','Valor_Bruto'),('Porcentaje_Copago','Copago_Cuota','Valor_Bruto'),('Porcentaje_Glosa','Glosa_Inicial','Valor_Neto'),('Porcentaje_Recaudo','Valor_Pagado','Valor_Neto'),('Porcentaje_Cartera','Saldo_Cartera','Valor_Neto')]:
  if {num,den}<=set(d): d[out]=np.where(d[den].ne(0),d[num]/d[den]*100,np.nan)
 if 'Glosa_Inicial' in d:d['Tiene_Glosa']=d.Glosa_Inicial.fillna(0)>0
 if 'Saldo_Cartera' in d:d['Tiene_Saldo']=d.Saldo_Cartera.fillna(0)>0
 if 'Estado_Factura' in d:d['Es_Pagada']=d.Estado_Factura.astype(str).str.casefold().eq('pagada')
 if 'Dias_Emision_Radicacion' in d:
  a,b,c=st.session_state.rangos; d['Rango_Tiempo_Radicacion']=pd.cut(d.Dias_Emision_Radicacion,[-np.inf,a,b,c,np.inf],labels=[f'0 a {a}',f'{a+1} a {b}',f'{b+1} a {c}',f'Más de {c}'])
 return d.replace([np.inf,-np.inf],np.nan)

def diagnostico(d,tol):
 q={'Filas':len(d),'Columnas':len(d.columns),'Duplicados exactos':int(d.duplicated().sum())}; inc=pd.Series(False,index=d.index)
 if 'ID_Factura'in d:q['ID duplicados']=int(d.ID_Factura.dropna().duplicated().sum())
 if 'Numero_Radicado'in d:q['Radicados duplicados']=int(d.Numero_Radicado.dropna().duplicated().sum())
 for c in FECHAS:
  if c in d:q['Fechas inválidas '+c]=int(d[c].isna().sum())
 for c in DINERO:
  if c in d:q['Negativos '+c]=int((d[c]<0).sum())
 if {'Fecha_Emision','Fecha_Radicacion'}<=set(d):q['Radicación anterior a emisión']=int((d.Fecha_Radicacion<d.Fecha_Emision).sum())
 if {'Fecha_Radicacion','Fecha_Pago'}<=set(d):q['Pago anterior a radicación']=int((d.Fecha_Pago<d.Fecha_Radicacion).sum())
 for c,n in [('Saldo_Cartera','Saldo superior al neto'),('Valor_Pagado','Pago superior al neto'),('Glosa_Inicial','Glosa superior al neto')]:
  if {c,'Valor_Neto'}<=set(d):q[n]=int((d[c]>d.Valor_Neto+tol).sum())
 if {'Valor_Neto','Valor_Bruto','Descuento','Copago_Cuota'}<=set(d):inc|=(d.Valor_Neto-(d.Valor_Bruto-d.Descuento.fillna(0)-d.Copago_Cuota.fillna(0))).abs()>tol
 if {'Saldo_Cartera','Valor_Neto','Glosa_Inicial','Valor_Pagado'}<=set(d):inc|=(d.Saldo_Cartera-(d.Valor_Neto-d.Glosa_Inicial.fillna(0)-d.Valor_Pagado.fillna(0))).abs()>tol
 q['Inconsistencias financieras']=int(inc.sum()); nulos=pd.DataFrame({'Columna':d.columns,'Nulos':[d[c].isna().sum() for c in d],'% nulos':[d[c].isna().mean()*100 for c in d]}); score=np.mean([100*(1-d.isna().mean().mean()),100*(1-d.duplicated().mean()),max(0,100-100*inc.mean())])
 return q,nulos,inc,float(score)

def fm(v):
 if pd.isna(v):return 'N/D'
 s={'COP':'$','USD':'US$','EUR':'€','Sin símbolo':''}[st.session_state.moneda]; return (f'{s} {v:,.0f}').replace(',','.').strip()
def pp(a,b):return np.nan if pd.isna(a) or not b else a/b*100
def plot(fig):fig.update_layout(template='plotly_white',margin=dict(l=20,r=20,t=55,b=20));st.plotly_chart(fig,use_container_width=True)

def agrupar(d,g):
 if g not in d:return pd.DataFrame()
 agg={c:'sum' for c in DINERO+['Cantidad'] if c in d}; agg.update({'ID_Factura':'nunique'} if 'ID_Factura'in d else {})
 x=d.groupby(g,dropna=False).agg(agg).reset_index().rename(columns={'ID_Factura':'Facturas'})
 if 'Valor_Neto'in x:
  x['Ticket_Promedio']=x.Valor_Neto/x.get('Facturas',1); x['Participacion']=x.Valor_Neto/x.Valor_Neto.sum()*100 if x.Valor_Neto.sum() else np.nan
  for o,c in [('Porcentaje_Recaudo','Valor_Pagado'),('Porcentaje_Cartera','Saldo_Cartera'),('Porcentaje_Glosa','Glosa_Inicial')]:
   if c in x:x[o]=np.where(x.Valor_Neto.ne(0),x[c]/x.Valor_Neto*100,np.nan)
 return x

def cierre(d,f='Mensual'):
 if 'Fecha_Radicacion'in d:d=d.assign(Periodo=d.Fecha_Radicacion.dt.to_period({'Mensual':'M','Trimestral':'Q','Anual':'Y'}[f]).astype(str))
 elif 'Periodo_Cierre'in d:d=d.assign(Periodo=d.Periodo_Cierre.astype(str))
 else:return pd.DataFrame()
 x=agrupar(d,'Periodo').sort_values('Periodo')
 if 'Valor_Neto'in x:x['Crecimiento']=x.Valor_Neto.pct_change()*100;x['Acumulado']=x.Valor_Neto.cumsum();x['Promedio_Movil_3']=x.Valor_Neto.rolling(3,min_periods=1).mean()
 return x

def filtros(d):
 x=d.copy(); st.sidebar.subheader('Filtros globales')
 if st.sidebar.button('Limpiar filtros'):st.session_state.fv=st.session_state.get('fv',0)+1;st.rerun()
 k=st.session_state.get('fv',0)
 if 'Fecha_Radicacion'in x and x.Fecha_Radicacion.notna().any():
  r=st.sidebar.date_input('Rango de fechas',(x.Fecha_Radicacion.min().date(),x.Fecha_Radicacion.max().date()),key=f'f{k}')
  if len(r)==2:x=x[x.Fecha_Radicacion.dt.date.between(*r)]
 for c in ['Año_Radicacion','Mes_Radicacion','Entidad_Pagadora','Regimen','Sede','Servicio','Linea_Servicio','Estado_Factura','Canal_Radicacion','Soporte_Completo']:
  if c in x:
   o=sorted(x[c].dropna().astype(str).unique());s=st.sidebar.multiselect(c.replace('_',' '),o,o,key=c+str(k));x=x[x[c].astype(str).isin(s)]
 if 'Tiene_Glosa'in x:
  v=st.sidebar.selectbox('Facturas con glosa',['Todas','Sí','No'],key='g'+str(k));x=x if v=='Todas' else x[x.Tiene_Glosa.eq(v=='Sí')]
 if 'Tiene_Saldo'in x:
  v=st.sidebar.selectbox('Facturas con saldo',['Todas','Sí','No'],key='s'+str(k));x=x if v=='Todas' else x[x.Tiene_Saldo.eq(v=='Sí')]
 if 'Valor_Neto'in x and x.Valor_Neto.notna().any():
  a,b=float(x.Valor_Neto.min()),float(x.Valor_Neto.max());r=st.sidebar.slider('Rango valor neto',a,b,(a,b),key='v'+str(k));x=x[x.Valor_Neto.between(*r)]
 st.sidebar.metric('Registros filtrados',len(x),f'{len(x)/max(len(d),1)*100:.1f}% visible');return x

def grupo(d,g,t):
 st.header(t)
 if g not in d:st.info(f'No se encontró {g}.');return
 x=agrupar(d,g);st.dataframe(x,use_container_width=True)
 if 'Valor_Neto'in x:
  plot(px.bar(x.nlargest(15,'Valor_Neto').sort_values('Valor_Neto'),x='Valor_Neto',y=g,orientation='h',title='Ranking por facturación'))
  z=x.sort_values('Valor_Neto',ascending=False).copy();z['Acumulado']=z.Valor_Neto.cumsum()/z.Valor_Neto.sum()*100;fig=make_subplots(specs=[[{'secondary_y':True}]]);fig.add_bar(x=z[g],y=z.Valor_Neto,name='Valor');fig.add_scatter(x=z[g],y=z.Acumulado,name='% acumulado',secondary_y=True);fig.add_hline(y=80,line_dash='dash',secondary_y=True);plot(fig)
 if {'Valor_Neto','Saldo_Cartera','Glosa_Inicial'}<=set(x):plot(px.scatter(x,x='Valor_Neto',y='Porcentaje_Recaudo',size='Saldo_Cartera',color='Porcentaje_Glosa',hover_name=g,title='Facturación, recaudo, cartera y glosa'))

def resumen(d):
 st.header('Resumen ejecutivo')
 if d.empty:st.warning('Los filtros actuales no tienen registros.');return
 s=lambda c:d[c].sum() if c in d else np.nan;n=s('Valor_Neto');p=s('Valor_Pagado');g=s('Glosa_Inicial');vals=[('Bruta',fm(s('Valor_Bruto'))),('Neta',fm(n)),('Facturas',d.ID_Factura.nunique()),('Pagado',fm(p)),('Cartera',fm(s('Saldo_Cartera'))),('Glosa',fm(g)),('% recaudo',f'{pp(p,n):.1f}%' if pd.notna(pp(p,n)) else 'N/D'),('% glosa',f'{pp(g,n):.1f}%' if pd.notna(pp(g,n)) else 'N/D'),('Ticket',fm(n/max(d.ID_Factura.nunique(),1))),('Días radicación',f'{d.Dias_Emision_Radicacion.mean():.1f}' if 'Dias_Emision_Radicacion'in d else 'N/D')]
 c=st.columns(5)
 for i,(a,b) in enumerate(vals):c[i%5].metric(a,b)
 x=cierre(d)
 if not x.empty:plot(px.line(x,x='Periodo',y=[c for c in ['Valor_Neto','Valor_Pagado','Saldo_Cartera'] if c in x],markers=True,title='Facturación, pagos y cartera'))
 for g in ['Estado_Factura','Entidad_Pagadora','Linea_Servicio']:
  x=agrupar(d,g)
  if not x.empty:plot(px.bar(x.nlargest(15,'Valor_Neto'),x=g,y='Valor_Neto',title=g.replace('_',' ')))

def estadistica(d):
 st.header('Análisis estadístico');nums=d.select_dtypes(include=np.number).columns.tolist()
 if not nums:st.info('No hay variables numéricas.');return
 a,b,c=st.tabs(['Descriptivo','Correlaciones','Pruebas'])
 with a:
  v=st.selectbox('Variable',nums);s=d[v].dropna();q1,q3=s.quantile([.25,.75]);tab={'Conteo':len(s),'Media':s.mean(),'Mediana':s.median(),'Moda':s.mode().iloc[0] if not s.mode().empty else np.nan,'Desviación':s.std(),'Varianza':s.var(),'Mínimo':s.min(),'Máximo':s.max(),'Rango':s.max()-s.min(),'Q1':q1,'Q3':q3,'IQR':q3-q1,'P90':s.quantile(.9),'P95':s.quantile(.95),'Coef. variación %':s.std()/s.mean()*100 if s.mean() else np.nan,'Asimetría':s.skew(),'Curtosis':s.kurt()};st.dataframe(pd.DataFrame(tab.items(),columns=['Estadístico','Valor']));plot(px.histogram(d,x=v,marginal='box',title=v));st.write(f'Media {s.mean():,.2f}, mediana {s.median():,.2f}, asimetría {s.skew():,.2f}. No se afirma normalidad sin evaluarla.')
 with b:
  m=st.selectbox('Método',['pearson','spearman']);co=d[nums].corr(method=m);plot(px.imshow(co,text_auto='.2f',zmin=-1,zmax=1,color_continuous_scale='RdBu_r'));x=st.selectbox('X',nums);y=st.selectbox('Y',nums,index=min(1,len(nums)-1));plot(px.scatter(d,x=x,y=y));st.warning('La correlación no implica causalidad. Algunas relaciones pueden surgir directamente de las fórmulas contables entre las variables.')
 with c:
  tipo=st.selectbox('Prueba',['Shapiro-Wilk','Mann-Whitney U','t de Student','Kruskal-Wallis','ANOVA de una vía','Chi-cuadrado']);alpha=st.number_input('Alfa',.001,.2,.05,.01);cats=[z for z in d if d[z].dtype=='object']
  try:
   if tipo=='Shapiro-Wilk':
    v=st.selectbox('Variable numérica',nums,key='pv');s=d[v].dropna();s=s.sample(min(5000,len(s)),random_state=42);e,p=stats.shapiro(s);n=len(s)
   elif tipo=='Chi-cuadrado':
    x=st.selectbox('Categoría 1',cats);y=st.selectbox('Categoría 2',cats,index=min(1,len(cats)-1));tab=pd.crosstab(d[x],d[y]);e,p,gl,_=stats.chi2_contingency(tab);n=tab.values.sum()
   else:
    v=st.selectbox('Variable numérica',nums,key='pgv');g=st.selectbox('Agrupación',cats);gr=[x[v].dropna() for _,x in d.groupby(g) if x[v].notna().sum()>=2];n=[len(x) for x in gr]
    if tipo=='Mann-Whitney U':e,p=stats.mannwhitneyu(*gr[:2])
    elif tipo=='t de Student':e,p=stats.ttest_ind(*gr[:2],equal_var=False)
    elif tipo=='Kruskal-Wallis':e,p=stats.kruskal(*gr)
    else:e,p=stats.f_oneway(*gr)
   st.write(f'H0: no hay diferencia, dependencia o desviación según la prueba. Estadístico={e:.4f}; p={p:.4g}; tamaño(s)={n}; decisión: {"rechazar H0" if p<alpha else "no rechazar H0"}. No rechazar H0 no demuestra que sea verdadera.')
  except:st.error('No existen datos o grupos suficientes para la prueba.')

def atipicos(d):
 st.header('Valores atípicos');nums=d.select_dtypes(include=np.number).columns
 if len(nums)==0:return pd.DataFrame()
 v=st.selectbox('Variable',nums);met=st.selectbox('Método',['IQR','Puntaje Z','Percentiles']);s=d[v]
 if met=='IQR':k=st.number_input('Multiplicador',.5,5.,1.5,.1);q1,q3=s.quantile([.25,.75]);li,ls=q1-k*(q3-q1),q3+k*(q3-q1);m=(s<li)|(s>ls)
 elif met=='Puntaje Z':u=st.number_input('Umbral Z',1.,6.,3.,.1);z=np.abs((s-s.mean())/s.std());li=ls=np.nan;m=z>u
 else:p=st.slider('Percentiles',0.,100.,(1.,99.));li,ls=s.quantile(p[0]/100),s.quantile(p[1]/100);m=(s<li)|(s>ls)
 o=d[m.fillna(False)];st.metric('Atípicos',len(o),f'{len(o)/max(len(d),1)*100:.2f}%');plot(px.box(d,y=v,points='outliers'));st.dataframe(o.head(500));st.caption('Un valor atípico no es necesariamente un error.');return o

def main():
 estilo()
 if st.sidebar.button('Reiniciar aplicación'):
  for k in list(st.session_state):del st.session_state[k]
  st.rerun()
 st.sidebar.selectbox('Moneda visual',['COP','USD','EUR','Sin símbolo'],key='moneda');st.sidebar.caption('No realiza conversión cambiaria.');st.session_state.rangos=(st.sidebar.number_input('Límite corto',1,30,3),st.sidebar.number_input('Límite medio',2,60,7),st.sidebar.number_input('Límite largo',3,180,15));tol=st.sidebar.number_input('Tolerancia de redondeo',0.,1e6,1.)
 f=st.sidebar.file_uploader('Cargue el archivo de facturación que desea analizar',type=['xlsx','csv'])
 if f is None:
  st.header('Inicio y carga de datos');st.markdown("<div class='privacy'><b>Privacidad:</b> Evite cargar nombres de pacientes, documentos, historias clínicas, diagnósticos u otra información sensible. Use archivos anonimizados y siga las políticas de su organización.</div>",unsafe_allow_html=True);st.info('Formatos aceptados: XLSX y CSV. Columnas mínimas: ID_Factura, Valor_Neto y Fecha_Radicacion o Periodo_Cierre.');return
 b=f.getvalue();fid=(f.name,len(b),hash(b[:2048]))
 if st.session_state.get('fid')!=fid:st.session_state.fid=fid;st.session_state.pop('mapa',None)
 try:
  if f.name.lower().endswith('.xlsx'):
   hs=excel_hojas(b);h=st.sidebar.selectbox('Hoja',hs,index=next((i for i,x in enumerate(hs) if norm(x)=='facturacion'),0));d=leer_excel(b,h)
  else:modo=st.sidebar.selectbox('Separador CSV',['Automático','Coma','Punto y coma','Tabulación']);d,enc,sep=leer_csv(b,modo);h=f'CSV {enc}'
 except:st.error('No fue posible leer el archivo.');return
 if d.empty:st.error('La hoja seleccionada no contiene datos.');return
 sug,eq=sugerencias(d.columns)
 if eq:
  with st.expander('Equivalencias encontradas'):st.dataframe(pd.DataFrame(eq,columns=['Original','Esperado']))
 mapa=mapeo_ui(d,sug);d=d.rename(columns={v:k for k,v in mapa.items()}).loc[:,lambda x:~x.columns.duplicated()]
 falt=MINIMAS-set(d);ft=not ({'Fecha_Radicacion','Periodo_Cierre'}&set(d))
 if falt or ft:st.error('Faltan columnas mínimas.');st.write('Encontradas:',list(d.columns));st.write('Faltantes:',list(falt)+(['Fecha_Radicacion o Periodo_Cierre'] if ft else []));return
 if RECOMENDADAS-set(d):st.warning('Recomendadas ausentes: '+', '.join(RECOMENDADAS-set(d)))
 st.sidebar.subheader('Limpieza controlada');vac=st.sidebar.checkbox('Eliminar filas vacías',True);dup=st.sidebar.checkbox('Eliminar duplicados exactos');conv=st.sidebar.checkbox('Convertir fechas y monetarios',True);cats=st.sidebar.checkbox("Vacíos categóricos como 'Sin información'");inv=st.sidebar.checkbox('Excluir Valor_Neto inválido');ap=st.sidebar.checkbox('Aplicar transformaciones',True)
 original=d.copy();rep=pd.DataFrame()
 if ap:
  if vac:d=d.dropna(how='all')
  if dup:d=d.drop_duplicates()
  if conv:d,rep=convertir(d)
  if cats:
   c=d.select_dtypes(include='object').columns;d[c]=d[c].fillna('Sin información')
  if inv:d=d[d.Valor_Neto.notna()]
 d=derivadas(d);st.sidebar.caption(f'Iniciales {len(original):,} · posteriores {len(d):,} · excluidos {len(original)-len(d):,}');df=filtros(d);q,nul,inc,score=diagnostico(d,tol);pag=st.sidebar.radio('Navegación',PAGINAS)
 if pag==PAGINAS[0]:st.header('Inicio');st.metric('Archivo',f.name);st.write(f'Hoja: {h} · filas: {len(d):,} · columnas: {len(d.columns):,}');st.dataframe(d.sample(min(10,len(d)),random_state=42) if st.checkbox('Muestra aleatoria') else d.head(10),use_container_width=True)
 elif pag==PAGINAS[1]:st.header('Diagnóstico de calidad de datos');st.metric('Puntaje orientativo',f'{score:.1f}/100');st.caption('Indicador exploratorio; no reemplaza una auditoría.');st.dataframe(pd.DataFrame(q.items(),columns=['Validación','Resultado']));st.dataframe(nul);st.dataframe(rep) if not rep.empty else None
 elif pag==PAGINAS[2]:resumen(df)
 elif pag==PAGINAS[3]:st.header('Cierre mensual');fr=st.selectbox('Periodicidad',['Mensual','Trimestral','Anual']);x=cierre(df,fr);st.dataframe(x);plot(px.line(x,x='Periodo',y=[c for c in ['Valor_Neto','Valor_Pagado','Saldo_Cartera','Promedio_Movil_3'] if c in x],markers=True)) if not x.empty else st.info('Sin periodos válidos.')
 elif pag==PAGINAS[4]:grupo(df,'Entidad_Pagadora','Entidades pagadoras')
 elif pag==PAGINAS[5]:grupo(df,st.selectbox('Dimensión',[c for c in ['Servicio','Linea_Servicio'] if c in df] or ['No disponible']),'Servicios y líneas')
 elif pag==PAGINAS[6]:grupo(df,st.selectbox('Dimensión',[c for c in ['Sede','Regimen'] if c in df] or ['No disponible']),'Sedes y regímenes')
 elif pag==PAGINAS[7]:grupo(df,'Estado_Factura','Estados de facturación')
 elif pag==PAGINAS[8]:
  st.header('Glosas');g=df[df.Glosa_Inicial.fillna(0)>0] if 'Glosa_Inicial'in df else pd.DataFrame();st.metric('Valor glosado',fm(g.Glosa_Inicial.sum()) if not g.empty else 'N/D');plot(px.histogram(g,x='Glosa_Inicial')) if not g.empty else st.info('No hay glosas.')
 elif pag==PAGINAS[9]:grupo(df,'Entidad_Pagadora','Recaudo y cartera')
 elif pag==PAGINAS[10]:
  st.header('Tiempos de radicación');s=df.Dias_Emision_Radicacion.dropna() if 'Dias_Emision_Radicacion'in df else pd.Series(dtype=float);st.dataframe(s.describe(percentiles=[.75,.9,.95]).to_frame()) if not s.empty else st.info('No disponible.');plot(px.histogram(df,x='Dias_Emision_Radicacion')) if not s.empty else None
 elif pag==PAGINAS[11]:estadistica(df)
 elif pag==PAGINAS[12]:st.session_state.atip=atipicos(df)
 elif pag==PAGINAS[13]:
  st.header('Insights y alertas');x=cierre(df);hall=[]
  if not x.empty:top=x.loc[x.Valor_Neto.idxmax()];hall.append(('Periodo con mayor facturación',f'{top.Periodo}: {fm(top.Valor_Neto)}'))
  for g,c,t in [('Entidad_Pagadora','Valor_Neto','Entidad con mayor facturación'),('Entidad_Pagadora','Saldo_Cartera','Entidad con mayor saldo'),('Servicio','Valor_Neto','Servicio con mayor facturación')]:
   if {g,c}<=set(df):z=df.groupby(g)[c].sum();hall.append((t,f'{z.idxmax()}: {fm(z.max())}')) if len(z) else None
  hall.append(('Inconsistencias financieras',str(q['Inconsistencias financieras'])));st.session_state.ins=pd.DataFrame(hall,columns=['Hallazgo','Valor'])
  for a,v in hall:st.info(f'**Hallazgo calculado:** {a}. **Valor observado:** {v}. Se observa este resultado en los datos filtrados. Conviene revisar su composición. La información disponible no permite establecer la causa.')
 elif pag==PAGINAS[14]:
  st.header('Explorador');qtxt=st.text_input('Buscar texto');z=df[df.astype(str).apply(lambda c:c.str.contains(qtxt,case=False,na=False)).any(axis=1)] if qtxt else df;cols=st.multiselect('Columnas',z.columns.tolist(),z.columns.tolist()[:15]);st.dataframe(z[cols].head(st.selectbox('Filas',[25,50,100,250,500])),use_container_width=True,height=550)
 elif pag==PAGINAS[15]:
  st.header('Descargas');items=[('facturacion_filtrada',df),('cierre_mensual',cierre(df)),('analisis_entidades',agrupar(df,'Entidad_Pagadora')),('analisis_servicios',agrupar(df,'Servicio')),('analisis_sedes',agrupar(df,'Sede')),('registros_glosados',df[df.Glosa_Inicial.fillna(0)>0] if 'Glosa_Inicial'in df else pd.DataFrame()),('registros_cartera',df[df.Saldo_Cartera.fillna(0)>0] if 'Saldo_Cartera'in df else pd.DataFrame()),('registros_atipicos',st.session_state.get('atip',pd.DataFrame())),('informe_calidad',pd.DataFrame(q.items(),columns=['Validación','Resultado']))]
  for n,x in items:
   if not x.empty:
    st.download_button(n+' CSV',x.to_csv(index=False).encode('utf-8-sig'),n+'.csv');buf=io.BytesIO();
    with pd.ExcelWriter(buf,engine='openpyxl') as w:x.to_excel(w,index=False)
    st.download_button(n+' Excel',buf.getvalue(),n+'.xlsx')
  if 'ins'in st.session_state:st.download_button('Insights',st.session_state.ins.to_csv(index=False),'insights_facturacion.txt')
 else:st.header('Metodología');st.markdown('**% glosa:** `Glosa_Inicial / Valor_Neto × 100`  \n**% recaudo:** `Valor_Pagado / Valor_Neto × 100`  \n**% cartera:** `Saldo_Cartera / Valor_Neto × 100`  \n**Ticket promedio:** `Valor_Neto / facturas únicas`. Los ceros producen nulos. Correlación no implica causalidad. IQR usa Q1 y Q3. El p-valor evalúa compatibilidad con H0. Todo resultado depende de la calidad de datos y no reemplaza auditorías.')
if __name__=='__main__':main()
