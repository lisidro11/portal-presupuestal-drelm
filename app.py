import io, os, sqlite3, secrets, hashlib, hmac, string, time
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4

ROOT=Path(__file__).resolve().parent
DATA=Path(os.environ.get('DRELM_DATA_DIR', str(ROOT/'datos'))); DATA.mkdir(parents=True,exist_ok=True)
DB=DATA/'usuarios.sqlite'
INITIAL=Path(os.environ.get('DRELM_INITIAL_XLSX', str(ROOT/'fuentes'/'BASE_INICIAL.xlsx')))
st.set_page_config(page_title='DRELM | Inteligencia presupuestal',page_icon='🏛️',layout='wide',initial_sidebar_state='expanded')
st.markdown('''<style>
.block-container{padding-top:.75rem;max-width:1800px;padding-bottom:2rem}
[data-testid="stMetric"]{background:#fff;border:1px solid #e0e9f2;border-radius:12px;padding:13px 16px;box-shadow:0 2px 9px #092e5208}
[data-testid="stMetricLabel"]{color:#526b84} [data-testid="stMetricValue"]{color:#103e6d;font-size:1.55rem}
h1,h2,h3{color:#103e6d} h1{font-size:1.9rem!important}
[data-testid="stSidebar"]{background:#f2f7fc} .stButton>button[kind="primary"]{background:#0d579d}
</style>''',unsafe_allow_html=True)

def conn():
 c=sqlite3.connect(DB);c.row_factory=sqlite3.Row;return c
with conn() as c:
 c.execute('CREATE TABLE IF NOT EXISTS accounts(username TEXT PRIMARY KEY, salt TEXT NOT NULL, hash TEXT NOT NULL, role TEXT NOT NULL, institute TEXT, enabled INTEGER DEFAULT 1)')
 c.execute('CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)')
 c.execute('CREATE TABLE IF NOT EXISTS audits(at TEXT,username TEXT,action TEXT)')
def log(action):
 with conn() as c:c.execute('INSERT INTO audits VALUES(?,?,?)',(datetime.now().isoformat(timespec='seconds'),st.session_state.get('login',''),action))
def hash_password(p,s):return hashlib.pbkdf2_hmac('sha256',p.encode(),bytes.fromhex(s),450000).hex()
def add_account(u,p,role,inst=None):
 salt=secrets.token_hex(16)
 with conn() as c:c.execute('INSERT INTO accounts(username,salt,hash,role,institute) VALUES (?,?,?,?,?)',(u,salt,hash_password(p,salt),role,inst))
def short_password():
 return "".join(secrets.choice(string.ascii_lowercase) for _ in range(6))

def validate(u,p):
 with conn() as c:r=c.execute('SELECT * FROM accounts WHERE username=? AND enabled=1',(u,)).fetchone()
 return r if r and hmac.compare_digest(hash_password(p,r['salt']),r['hash']) else None

def source_path():
 with conn() as c:r=c.execute("SELECT value FROM settings WHERE key='current_file'").fetchone()
 return Path(r['value']) if r and Path(r['value']).exists() else INITIAL

def money(x):return f'S/ {x:,.2f}'
def metric_data(df):
 cols=['mto_pia','mto_modificaciones','mto_pim','mto_certificado','mto_compro_anual','devengado','girado','pagado']
 m={x:float(df[x].sum()) for x in cols}
 m['avance']=m['devengado']/m['mto_pim']*100 if m['mto_pim'] else 0
 m['saldo']=m['mto_pim']-m['devengado'];m['sin_certificar']=m['mto_pim']-m['mto_certificado']
 m['sin_devengar']=m['mto_compro_anual']-m['devengado']
 return m
@st.cache_data(show_spinner='Procesando base presupuestal…',max_entries=3)
def read_data(path,mtime):
 x=pd.ExcelFile(path)
 candidates=[]
 for name in x.sheet_names:
  sample=pd.read_excel(x,sheet_name=name,nrows=0)
  columns=set(map(str,sample.columns))
  if {'mto_pim','mto_pia','mto_certificado'}.issubset(columns):candidates.append((name,len(columns),sum(z in columns for z in ['INSTITUTO','NOMBRE','RESPONSABLES','SIGLAS'])))
 if not candidates:raise ValueError('No se encontró una hoja de datos con mto_pia, mto_pim y mto_certificado.')
 sheet=sorted(candidates,key=lambda t:(t[2],t[1]),reverse=True)[0][0]
 df=pd.read_excel(x,sheet_name=sheet,dtype={'sec_func':str})
 df.columns=[str(c).strip() for c in df.columns]
 for c in ['INSTITUTO','NOMBRE','RESPONSABLES','SIGLAS','FAMILIA','generica','programa_pptal','producto_proyecto','activ_obra_accinv','sec_func','fuente_financ','rubro','categoria_gasto','subgenerica','especifica','CLASIFICADOR','Denominacion','funcion','division_fn','grupo_fn','meta','finalidad']:
  if c not in df:df[c]=''
  df[c]=df[c].fillna('').astype(str).str.strip()
 for c in ['mto_pia','mto_modificaciones','mto_pim','mto_certificado','mto_compro_anual']+[f'mto_{p}_{i:02d}' for p in ['devenga','girado','pagado','at_comp'] for i in range(1,13)]:
  if c not in df:df[c]=0
  df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0)
 for key,p in [('devengado','devenga'),('girado','girado'),('pagado','pagado'),('compromiso_mensual','at_comp')]:df[key]=df[[f'mto_{p}_{i:02d}' for i in range(1,13)]].sum(axis=1)
 return df,sheet

def excel_bytes(df):
 out=io.BytesIO()
 with pd.ExcelWriter(out,engine='xlsxwriter') as w:df.to_excel(w,index=False,sheet_name='Datos')
 return out.getvalue()
def pdf_bytes(df,title,filename):
 m=metric_data(df);o=io.BytesIO();doc=SimpleDocTemplate(o,pagesize=A4,leftMargin=39,rightMargin=39,topMargin=45,bottomMargin=40)
 ss=getSampleStyleSheet();ss.add(ParagraphStyle(name='ReportTitleD',parent=ss['Heading1'],fontSize=16,textColor=colors.HexColor('#123d6c'),leading=20,spaceAfter=10))
 ss.add(ParagraphStyle(name='ReportSmallD',parent=ss['Normal'],fontSize=8,leading=12,textColor=colors.HexColor('#5a6c7d')))
 story=[Paragraph('DRELM | INFORME EJECUTIVO PRESUPUESTAL',ss['ReportTitleD']),Paragraph(title,ss['Heading2']),Paragraph('Fuente: '+filename+' | Generado: '+datetime.now().strftime('%d/%m/%Y %H:%M'),ss['ReportSmallD']),Spacer(1,18)]
 rows=[['Indicador','Monto (S/)']]+[[lab,f'{m[key]:,.2f}'] for lab,key in [('PIA','mto_pia'),('Modificaciones','mto_modificaciones'),('PIM','mto_pim'),('Certificación','mto_certificado'),('Compromiso anual','mto_compro_anual'),('Devengado','devengado'),('Girado','girado'),('Pagado','pagado'),('PIM menos devengado','saldo')]]
 t=Table(rows,colWidths=[290,185]);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#123d6c')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#eff5fb')]),('ALIGN',(1,1),(-1,-1),'RIGHT'),('BOTTOMPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),8)]));story.extend([t,Spacer(1,16),Paragraph(f'<b>Ejecución (devengado/PIM): {m["avance"]:.2f}%</b>',ss['Normal']),Spacer(1,13)])
 story.extend([Paragraph('Síntesis ejecutiva',ss['Heading2']),Paragraph(f'El ámbito seleccionado presenta un PIM de {money(m["mto_pim"])} y devengado de {money(m["devengado"])}. La diferencia PIM–devengado asciende a {money(m["saldo"])}; no equivale necesariamente a saldo de libre disponibilidad.',ss['Normal']),Spacer(1,13)])
 for field,lab in [('NOMBRE','Unidades / institutos'),('generica','Genérica de gasto'),('programa_pptal','Programa presupuestal'),('CLASIFICADOR','Clasificador')]:
  if field=='NOMBRE' and df[field].nunique()>30:continue
  g=df.groupby(field)[['mto_pim','devengado']].sum().sort_values('mto_pim',ascending=False).head(8)
  if g.empty:continue
  rows=[['Concepto','PIM','Devengado']]+[[str(k)[:42],f'{r.mto_pim:,.0f}',f'{r.devengado:,.0f}'] for k,r in g.iterrows()]
  t=Table(rows,colWidths=[255,110,110],repeatRows=1);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e1ebf5')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f5f8fc')]),('FONTSIZE',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),6)]));story.append(KeepTogether([Paragraph(lab,ss['Heading2']),t,Spacer(1,13)]))
 story.extend([Paragraph('Metodología y alcance',ss['Heading2']),Paragraph('Los importes se agregan sobre los registros del archivo fuente, sin sumar etapas presupuestales entre sí. Devengado, girado y pagado se calculan sumando sus doce columnas mensuales. Porcentaje de ejecución = devengado / PIM × 100, cuando PIM es distinto de cero. La identificación institucional depende de los campos de la fuente y requiere validación para habilitar cuentas individuales.',ss['ReportSmallD'])])
 doc.build(story);return o.getvalue()

# Recuperación administrativa controlada desde Streamlit Cloud Secrets.
# ADMIN_NUEVA_CLAVE nunca debe publicarse en GitHub.
# Al terminar, eliminar el secreto de Streamlit Cloud.
admin_new_password = st.secrets.get('ADMIN_NUEVA_CLAVE', '')
if admin_new_password:
 if len(str(admin_new_password)) < 12:
  st.error('ADMIN_NUEVA_CLAVE debe tener al menos 12 caracteres.'); st.stop()
 with conn() as c:
  target = c.execute("SELECT username FROM accounts WHERE username=?", ('Nilton',)).fetchone()
  if target:
   existing = c.execute('SELECT salt,hash,role,enabled FROM accounts WHERE username=?', ('Nilton',)).fetchone()
   if existing['role'] != 'admin':
    st.error('El nombre Nilton pertenece a una cuenta que no es administradora.'); st.stop()
   if (existing['enabled'] != 1 or not hmac.compare_digest(hash_password(str(admin_new_password),existing['salt']),existing['hash'])):
    salt = secrets.token_hex(16)
    c.execute('UPDATE accounts SET salt=?,hash=?,enabled=1 WHERE username=?',
              (salt,hash_password(str(admin_new_password),salt),'Nilton'))
    c.execute('INSERT INTO audits VALUES(?,?,?)',
              (datetime.now().isoformat(timespec='seconds'),'Nilton','RESET_ADMIN_VIA_SECRET'))
  else:
   salt = secrets.token_hex(16)
   c.execute('INSERT INTO accounts(username,salt,hash,role,institute,enabled) VALUES(?,?,?,?,?,?)',
             ('Nilton',salt,hash_password(str(admin_new_password),salt),'admin',None,1))
   c.execute('INSERT INTO audits VALUES(?,?,?)',
             (datetime.now().isoformat(timespec='seconds'),'Nilton','CREAR_ADMIN_VIA_SECRET'))

with conn() as c:accounts=c.execute("SELECT COUNT(*) FROM accounts WHERE role='admin'").fetchone()[0]
if not accounts:
 st.title('🏛️ Configuración inicial · DRELM')
 st.info('Crea la cuenta de administración. El corte inicial ya está incluido y se mostrará inmediatamente después de ingresar.')
 with st.form('setup'):
  username=st.text_input('Usuario administrador',value='admin');p=st.text_input('Contraseña (12 caracteres como mínimo)',type='password');q=st.text_input('Confirmar contraseña',type='password');submit=st.form_submit_button('Crear administrador',type='primary')
 if submit:
  if len(p)<12 or p!=q or not username.strip():st.error('Revisa el usuario y las contraseñas.')
  elif conn().execute('SELECT 1 FROM accounts WHERE username=?',(username.strip(),)).fetchone():st.error('Ese usuario ya está reservado. Elige otro.')
  else:add_account(username.strip(),p,'admin');st.success('Administrador creado. Inicia sesión.');st.rerun()
 st.stop()
if 'login' not in st.session_state:
 # Recuperación temporal: código definido exclusivamente en Streamlit Cloud Secrets.
 # Eliminar RECOVERY_TOKEN de Secrets después de recuperar el acceso.
 recovery_token=st.secrets.get('RECOVERY_TOKEN','')
 if recovery_token and len(str(recovery_token))>=24:
  with st.expander('🔐 Recuperar acceso del administrador'):
   with st.form('admin_recovery'):
    recovery_user=st.text_input('Usuario administrador a recuperar',value='Nilton')
    recovery_code=st.text_input('Código privado de recuperación',type='password')
    recovery_pw=st.text_input('Nueva contraseña (mínimo 12 caracteres)',type='password')
    recovery_confirm=st.text_input('Confirmar nueva contraseña',type='password')
    recovery_submit=st.form_submit_button('Restablecer acceso')
   if recovery_submit:
    with conn() as c:
     recovery_account=c.execute('SELECT role FROM accounts WHERE username=?',(recovery_user.strip(),)).fetchone()
    if (not recovery_account or recovery_account['role']!='admin'
        or not hmac.compare_digest(str(recovery_code),str(recovery_token))):
     st.error('Usuario administrador o código de recuperación incorrecto.')
    elif len(recovery_pw)<12 or recovery_pw!=recovery_confirm:
     st.error('La nueva contraseña debe tener al menos 12 caracteres y coincidir.')
    else:
     new_salt=secrets.token_hex(16)
     with conn() as c:
      c.execute('UPDATE accounts SET salt=?,hash=?,enabled=1 WHERE username=?',
                (new_salt,hash_password(recovery_pw,new_salt),recovery_user.strip()))
      c.execute('INSERT INTO audits VALUES(?,?,?)',
                (datetime.now().isoformat(timespec='seconds'),recovery_user.strip(),'RECUPERACION_ADMIN'))
     st.success('Acceso restablecido. Inicia sesión con tu nueva contraseña. Elimina RECOVERY_TOKEN de Streamlit Secrets inmediatamente.')
 st.title('🏛️ DRELM | Portal de Inteligencia Presupuestal')
 st.caption('Acceso institucional a la información presupuestal autorizada')
 with st.form('signin'):
  username=st.text_input('Usuario / código modular');p=st.text_input('Contraseña',type='password');go=st.form_submit_button('Ingresar',type='primary',use_container_width=True)
 if go:
  attempts=st.session_state.get('failed_logins',0)
  last=st.session_state.get('last_failed_at',0)
  if attempts>=5 and time.time()-last<300:
   st.error('Demasiados intentos. Espera cinco minutos antes de volver a intentar.')
  else:
   if time.time()-last>=300:st.session_state.failed_logins=0
   r=validate(username.strip(),p)
   if r:st.session_state.failed_logins=0;st.session_state.login=username.strip();log('LOGIN');st.rerun()
   else:
    st.session_state.failed_logins=st.session_state.get('failed_logins',0)+1
    st.session_state.last_failed_at=time.time()
    st.error('Credenciales incorrectas o usuario deshabilitado.')
 st.stop()
with conn() as c:account=c.execute('SELECT * FROM accounts WHERE username=? AND enabled=1',(st.session_state.login,)).fetchone()
if not account:st.session_state.clear();st.rerun()
admin=account['role']=='admin'
path=source_path()
if not path.exists():
 if admin:
  st.title('Carga inicial del presupuesto')
  st.info('No hay un Excel presupuestal en este servidor. Carga la fuente enriquecida con la columna NOMBRE para activar el panel.')
  initial_upload=st.file_uploader('Excel de gastos (.xlsx)',type=['xlsx'],key='initial_source')
  if initial_upload and st.button('Validar y activar datos',type='primary'):
   candidate=DATA/'validacion_inicial.xlsx';candidate.write_bytes(initial_upload.getvalue())
   try:
    trial,_=read_data(str(candidate),candidate.stat().st_mtime)
    if not trial['NOMBRE'].str.strip().any():st.error('La fuente no identifica instituciones en la columna NOMBRE.')
    else:
     target=DATA/'corte_inicial.xlsx';target.write_bytes(initial_upload.getvalue())
     with conn() as c:c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('current_file',?)",(str(target),))
     st.cache_data.clear();st.rerun()
   except Exception as e:st.error('Archivo inválido: '+str(e))
 else:st.error('El administrador todavía no ha publicado el corte presupuestal.')
 st.stop()
try:df,sheet=read_data(str(path),path.stat().st_mtime)
except Exception as ex:
 st.error(f'No se pudo leer el Excel: {ex}');st.stop()
if not admin:
 df=df[(df['NOMBRE'].eq(account['institute'])) & (df['INSTITUTO'].isin(['ISTP','EESP']))].copy()
 if df.empty:st.warning('Tu institución aún no tiene registros asociados al corte publicado. Contacta al administrador.');st.stop()

with st.sidebar:
 st.markdown('### 🏛️ DRELM | Presupuesto')
 st.caption('Usuario: '+account['username']+' · '+('Administrador' if admin else 'Instituto'))
 if not admin:st.caption('Acceso institucional mediante código modular.')
 nav=st.radio('Módulos',['Panel ejecutivo','Clasificaciones presupuestales','Análisis integral','Ejecución mensual','Detalle y descarga','Informe ejecutivo']+(['Administración'] if admin else []))
 if admin and nav=='Administración':st.info('Administración utiliza la base completa, sin filtros de consulta.')
 st.divider();st.markdown('**Filtros de consulta**')
 if admin and nav!='Administración':
  for field,label in [('INSTITUTO','Tipo de entidad'),('NOMBRE','Instituto / dependencia'),('RESPONSABLES','Responsable'),('SIGLAS','Siglas')]:
   opts=sorted(z for z in df[field].unique().tolist() if z)
   choice=st.multiselect(label,opts,placeholder='Todos',key='f_'+field)
   if choice:df=df[df[field].isin(choice)]
 if nav!='Administración':
  for field,label in [('programa_pptal','Programa presupuestal'),('generica','Genérica de gasto'),('fuente_financ','Fuente de financiamiento'),('rubro','Rubro'),('sec_func','Secuencia funcional'),('FAMILIA','Familia')]:
   opts=sorted(z for z in df[field].unique().tolist() if z)
   choice=st.selectbox(label,['Todas']+opts,key='f_'+field)
   if choice!='Todas':df=df[df[field].eq(choice)]
 st.divider()
 if st.button('Cerrar sesión',use_container_width=True):log('LOGOUT');st.session_state.clear();st.rerun()

scope='Consolidado DRELM' if admin else account['institute']
if admin and not df.empty and df['NOMBRE'].nunique()==1:scope=df['NOMBRE'].iloc[0]
st.title('Ejecución presupuestal | '+scope)
st.caption(f'Fuente: {path.name} · Hoja: {sheet} · {len(df):,} registros filtrados · Actualizado según archivo de origen')
if df.empty and nav!='Administración':st.warning('No hay registros con los filtros seleccionados.');st.stop()
m=metric_data(df)
if nav=='Panel ejecutivo':
 a,b,c,d=st.columns(4)
 for col,label,key in [(a,'PIM','mto_pim'),(b,'Devengado','devengado'),(c,'Saldo PIM - devengado','saldo'),(d,'Certificado','mto_certificado')]:col.metric(label,money(m[key]))
 a,b,c,d=st.columns(4)
 for col,label,value in [(a,'PIA',money(m['mto_pia'])),(b,'Compromiso anual',money(m['mto_compro_anual'])),(c,'Girado',money(m['girado'])),(d,'Ejecución',f'{m["avance"]:.2f}%')]:col.metric(label,value)
 st.progress(min(1,max(0,m['avance']/100)),text=f'Avance devengado / PIM: {m["avance"]:.2f}%')
 a,b=st.columns([1.3,1])
 with a:
  months=pd.DataFrame({'Mes':['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'],'Devengado':[df[f'mto_devenga_{i:02d}'].sum() for i in range(1,13)],'Girado':[df[f'mto_girado_{i:02d}'].sum() for i in range(1,13)]})
  st.plotly_chart(px.bar(months,x='Mes',y=['Devengado','Girado'],barmode='group',title='Ejecución mensual',color_discrete_sequence=['#1768a9','#13a7a0']),use_container_width=True)
 with b:
  g=df.groupby('generica',as_index=False)[['mto_pim','devengado']].sum().sort_values('mto_pim',ascending=False).head(12)
  st.plotly_chart(px.bar(g,x='mto_pim',y='generica',orientation='h',title='PIM por genérica de gasto',color_discrete_sequence=['#1768a9']),use_container_width=True)
 a,b=st.columns(2)
 with a:
  g=df.groupby('NOMBRE',as_index=False)[['mto_pim','devengado']].sum().sort_values('mto_pim',ascending=False).head(12)
  st.plotly_chart(px.bar(g,x='mto_pim',y='NOMBRE',orientation='h',title='Principales institutos / dependencias',color_discrete_sequence=['#1985c7']),use_container_width=True)
 with b:
  g=df.groupby('programa_pptal',as_index=False)[['mto_pim','devengado']].sum().sort_values('mto_pim',ascending=False).head(12)
  st.plotly_chart(px.bar(g,x='mto_pim',y='programa_pptal',orientation='h',title='PIM por programa presupuestal',color_discrete_sequence=['#15a58c']),use_container_width=True)
 st.info(f'Lectura ejecutiva: {m["avance"]:.2f}% de ejecución devengada del PIM. Diferencia PIM–devengado: {money(m["saldo"])}. Esta diferencia no implica necesariamente disponibilidad presupuestal.')

 st.divider()
 st.subheader('📊 Cuadros presupuestales por clasificación')
 st.caption('Análisis real del Excel cargado. Cada gráfico, cuadro y descarga respeta los filtros y los permisos de la cuenta.')
 panel_dims=[('generica','Genérica de gasto'),('fuente_financ','Fuente de financiamiento'),('rubro','Rubro'),('FAMILIA','Familia'),('programa_pptal','Programa presupuestal'),('producto_proyecto','Producto / proyecto'),('activ_obra_accinv','Actividad / obra'),('sec_func','Secuencia funcional'),('CLASIFICADOR','Clasificador de gasto'),('categoria_gasto','Categoría de gasto'),('subgenerica','Subgenérica'),('RESPONSABLES','Responsable')]
 panel_amounts=['mto_pia','mto_modificaciones','mto_pim','mto_certificado','mto_compro_anual','compromiso_mensual','devengado','girado','pagado']
 def panel_group(field):
  tab=df.groupby(field,dropna=False)[panel_amounts].sum().reset_index()
  tab[field]=tab[field].replace('', 'Sin información')
  tab['Ejecución %']=np.where(tab['mto_pim']!=0,tab['devengado']/tab['mto_pim']*100,np.nan)
  tab['Saldo PIM - devengado']=tab['mto_pim']-tab['devengado']
  tab['Saldo por certificar']=tab['mto_pim']-tab['mto_certificado']
  return tab.sort_values('mto_pim',ascending=False)
 panel_tabs=st.tabs(['Genérica','Fuente','Rubro','Familia','Más clasificaciones'])
 for j,(field,label) in enumerate(panel_dims[:4]):
  with panel_tabs[j]:
   g=panel_group(field)
   st.markdown('#### '+label)
   if len(g)==1:st.info('La fuente filtrada contiene una sola categoría en esta clasificación.')
   left,right=st.columns([1.45,1])
   with left:
    plot=g.head(18).sort_values('mto_pim')
    fig=px.bar(plot,x=['mto_pim','devengado'],y=field,orientation='h',barmode='group',title='PIM frente a devengado',labels={'value':'S/','variable':'Indicador'},color_discrete_sequence=['#1663a3','#0b9c87'])
    fig.update_layout(height=max(340,min(750,len(plot)*35+130)))
    st.plotly_chart(fig,use_container_width=True,key='panel_chart_'+field)
   with right:
    st.plotly_chart(px.pie(g.head(15),names=field,values='mto_pim',hole=.55,title='Distribución del PIM'),use_container_width=True,key='panel_pie_'+field)
   st.dataframe(g,use_container_width=True,hide_index=True,height=340,column_config={'Ejecución %':st.column_config.NumberColumn(format='%.2f %%')})
   st.download_button('⬇ Descargar '+label+' en Excel',excel_bytes(g),'DRELM_'+field+'_filtrado.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='panel_excel_'+field)
 with panel_tabs[4]:
  extra=st.selectbox('Clasificación',panel_dims[4:],format_func=lambda item:item[1],key='panel_extra_dim')
  field,label=extra;g=panel_group(field)
  fig=px.bar(g.head(20).sort_values('mto_pim'),x=['mto_pim','devengado'],y=field,orientation='h',barmode='group',title=label)
  st.plotly_chart(fig,use_container_width=True)
  st.dataframe(g,use_container_width=True,hide_index=True,height=340)
  st.download_button('⬇ Descargar cuadro en Excel',excel_bytes(g),'DRELM_'+field+'_filtrado.xlsx',key='panel_extra_export')
 st.markdown('#### Descargar todos los cuadros y el detalle filtrado')
 def panel_excel_integral():
  buffer=io.BytesIO()
  with pd.ExcelWriter(buffer,engine='xlsxwriter') as writer:
   for field,label in panel_dims:
    summary=panel_group(field)
    sheet=label.replace('/','-')[:31]
    summary.to_excel(writer,sheet_name=sheet,index=False)
    worksheet=writer.sheets[sheet]
    worksheet.freeze_panes(1,1)
    worksheet.autofilter(0,0,len(summary),len(summary.columns)-1)
    worksheet.set_column(0,0,48)
    worksheet.set_column(1,len(summary.columns)-1,20)
   df.to_excel(writer,sheet_name='Detalle filtrado',index=False)
   writer.sheets['Detalle filtrado'].freeze_panes(1,0)
  return buffer.getvalue()
 st.download_button('⬇ DESCARGAR EXCEL INTEGRAL (12 CUADROS + DETALLE)',panel_excel_integral(),'DRELM_Presupuesto_Cuadros_y_Detalle.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',key='panel_integral_export')

elif nav=='Clasificaciones presupuestales':
 st.subheader('Clasificaciones presupuestales | Cuadros y gráficos')
 st.caption('Cada clasificación se calcula con los registros autorizados y los filtros de la barra lateral. No se suman etapas presupuestales entre sí.')
 dimensions=[('generica','Genérica de gasto'),('fuente_financ','Fuente de financiamiento'),('rubro','Rubro'),('FAMILIA','Familia'),('categoria_gasto','Categoría de gasto'),('subgenerica','Subgenérica'),('programa_pptal','Programa presupuestal'),('producto_proyecto','Producto / proyecto'),('activ_obra_accinv','Actividad / obra / acción de inversión'),('sec_func','Secuencia funcional'),('CLASIFICADOR','Clasificador de gasto'),('RESPONSABLES','Responsable'),('SIGLAS','Siglas')]
 monetary=['mto_pia','mto_modificaciones','mto_pim','mto_certificado','mto_compro_anual','compromiso_mensual','devengado','girado','pagado']
 def summary_by(field):
  out=df.groupby(field,dropna=False)[monetary].sum().reset_index()
  out[field]=out[field].replace('', 'Sin información')
  out['Ejecución %']=np.where(out['mto_pim']!=0,out['devengado']/out['mto_pim']*100,np.nan)
  out['Saldo PIM - devengado']=out['mto_pim']-out['devengado']
  out['Saldo por certificar']=out['mto_pim']-out['mto_certificado']
  return out.sort_values('mto_pim',ascending=False)
 tabs=st.tabs(['Genérica','Fuente','Rubro','Familia','Otros análisis'])
 for idx,(field,label) in enumerate(dimensions[:4]):
  with tabs[idx]:
   g=summary_by(field)
   st.markdown('#### '+label)
   a,b,c=st.columns(3)
   a.metric('Categorías',len(g));b.metric('PIM',money(g['mto_pim'].sum()));c.metric('Devengado',money(g['devengado'].sum()))
   left,right=st.columns([1.35,1])
   with left:
    top=g.head(18).sort_values('mto_pim')
    fig=px.bar(top,x=['mto_pim','devengado'],y=field,orientation='h',barmode='group',title='PIM y devengado por '+label.lower(),labels={'value':'S/','variable':'Etapa'},color_discrete_sequence=['#175a9c','#0da88b'])
    fig.update_layout(height=max(360,min(760,len(top)*34+140)),margin=dict(l=10,r=10,t=50,b=20))
    st.plotly_chart(fig,use_container_width=True)
   with right:
    top=g.head(12).copy();top['Categoría']=top[field].astype(str).str.slice(0,60)
    st.plotly_chart(px.pie(top,names='Categoría',values='mto_pim',hole=.53,title='Distribución del PIM'),use_container_width=True)
   st.dataframe(g,use_container_width=True,hide_index=True,height=430,column_config={'Ejecución %':st.column_config.NumberColumn(format='%.2f %%')})
   st.download_button('⬇ Descargar cuadro de '+label,excel_bytes(g),'DRELM_'+field+'.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='export_'+field)
 with tabs[4]:
  field=st.selectbox('Selecciona otra clasificación',options=[x for x,_ in dimensions[4:]],format_func=lambda x:dict(dimensions)[x],key='other_class')
  g=summary_by(field)
  st.plotly_chart(px.bar(g.head(20).sort_values('mto_pim'),x=['mto_pim','devengado'],y=field,orientation='h',barmode='group',title=dict(dimensions)[field]),use_container_width=True)
  st.dataframe(g,use_container_width=True,hide_index=True)
  st.download_button('Descargar clasificación seleccionada',excel_bytes(g),'DRELM_'+field+'.xlsx',key='export_other')
 st.divider()
 st.markdown('#### Exportación consolidada de todos los cuadros')
 def classification_workbook():
  buf=io.BytesIO()
  with pd.ExcelWriter(buf,engine='xlsxwriter') as writer:
   for field,label in dimensions:
    tab=summary_by(field)
    sheet=label.replace('/','-')[:31]
    tab.to_excel(writer,sheet_name=sheet,index=False)
    ws=writer.sheets[sheet];ws.freeze_panes(1,1);ws.autofilter(0,0,len(tab),len(tab.columns)-1)
    ws.set_column(0,0,42);ws.set_column(1,len(tab.columns)-1,19)
   df.to_excel(writer,sheet_name='Datos filtrados',index=False)
   ws=writer.sheets['Datos filtrados'];ws.freeze_panes(1,0)
  return buf.getvalue()
 st.download_button('⬇ Descargar Excel integral: 13 cuadros + detalle filtrado',classification_workbook(),'DRELM_Cuadros_Presupuestales_Integral.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
elif nav=='Análisis integral':
 field=st.selectbox('Dimensión de análisis',['NOMBRE','INSTITUTO','RESPONSABLES','SIGLAS','FAMILIA','programa_pptal','producto_proyecto','activ_obra_accinv','sec_func','fuente_financ','rubro','categoria_gasto','generica','subgenerica','especifica','CLASIFICADOR','Denominacion','funcion','division_fn','grupo_fn','meta','finalidad'])
 g=df.groupby(field,dropna=False)[['mto_pia','mto_modificaciones','mto_pim','mto_certificado','mto_compro_anual','devengado','girado','pagado']].sum().reset_index()
 g['Ejecución %']=np.where(g['mto_pim'].ne(0),g['devengado']/g['mto_pim']*100,np.nan);g['Saldo PIM - devengado']=g['mto_pim']-g['devengado'];g=g.sort_values('mto_pim',ascending=False)
 a,b=st.columns([2,1]);a.plotly_chart(px.bar(g.head(20),x='mto_pim',y=field,orientation='h',title='PIM por '+field,color_discrete_sequence=['#1669ad']),use_container_width=True)
 b.plotly_chart(px.scatter(g.head(20),x='mto_pim',y='devengado',size='mto_certificado',hover_name=field,title='PIM vs devengado'),use_container_width=True)
 st.dataframe(g,use_container_width=True,hide_index=True,height=440)
 st.download_button('Descargar análisis Excel',excel_bytes(g),'analisis_presupuestal.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
elif nav=='Ejecución mensual':
 monthnames=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
 rows=[]
 for i,month in enumerate(monthnames,1):
  rows.append({'Mes':month,'Compromiso mensual':df[f'mto_at_comp_{i:02d}'].sum(),'Devengado':df[f'mto_devenga_{i:02d}'].sum(),'Girado':df[f'mto_girado_{i:02d}'].sum(),'Pagado':df[f'mto_pagado_{i:02d}'].sum()})
 months=pd.DataFrame(rows);st.plotly_chart(px.line(months,x='Mes',y=['Compromiso mensual','Devengado','Girado','Pagado'],markers=True,title='Evolución mensual por etapa presupuestal'),use_container_width=True)
 st.dataframe(months,use_container_width=True,hide_index=True);st.download_button('Descargar serie mensual',excel_bytes(months),'ejecucion_mensual.xlsx')
elif nav=='Detalle y descarga':
 st.write(f'**{len(df):,} registros** y **{len(df.columns)} columnas** disponibles del corte, incluyendo campos originales y totales mensuales calculados.')
 st.info('Descarga filtrada: respeta los filtros laterales y la búsqueda. Descarga completa: incluye todos los registros de tu ámbito autorizado sin filtros de consulta.')
 term=st.text_input('Buscar en nombre, secuencia funcional o clasificador')
 view=df
 if term:
  mask=pd.Series(False,index=df.index)
  for field in ['NOMBRE','sec_func','CLASIFICADOR','Denominacion']:mask|=df[field].astype(str).str.contains(term,case=False,regex=False,na=False)
  view=df[mask]
 st.dataframe(view,use_container_width=True,height=520,hide_index=True)
 st.download_button('Descargar Excel filtrado (todos los campos)',excel_bytes(view),'Gastos_Filtrados_DRELM.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 if admin:
  full,_=read_data(str(path),path.stat().st_mtime)
 else:
  full,_=read_data(str(path),path.stat().st_mtime);full=full[full['NOMBRE'].eq(account['institute'])].copy()
 st.download_button('Descargar Excel completo (sin filtros)',excel_bytes(full),'Gastos_Completos_Autorizados.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
elif nav=='Informe ejecutivo':
 st.subheader('Informe ejecutivo en PDF')
 st.write('El PDF incluye indicadores, interpretación, principales institutos o dependencias, genéricas, programas, clasificadores y metodología, de acuerdo con los filtros actuales.')
 if st.button('Generar PDF',type='primary'):
  try:st.session_state['report_pdf']=pdf_bytes(df,scope,path.name);st.session_state['report_key']=(path.name,tuple(df.index.tolist()));log('PDF')
  except Exception as e:st.error('No se pudo generar el PDF: '+str(e))
 if st.session_state.get('report_pdf') and st.session_state.get('report_key')==(path.name,tuple(df.index.tolist())):st.download_button('Descargar informe ejecutivo PDF',st.session_state['report_pdf'],'Informe_Ejecutivo_DRELM.pdf',mime='application/pdf')
elif nav=='Administración':
 tabs=st.tabs(['Actualizar corte','Institutos y cuentas','Calidad y estructura','Auditoría','Historial'])
 with tabs[0]:
  st.subheader('Actualizar información con un nuevo Excel')
  st.caption('Se conserva el archivo previo en el historial. La nueva carga solo se activa después de validar y confirmar.')
  upload=st.file_uploader('Nuevo archivo Excel (.xlsx)',type=['xlsx'])
  if upload:
   candidate=DATA/'revision_temporal.xlsx';candidate.write_bytes(upload.getvalue())
   try:
    new,sheet_new=read_data(str(candidate),candidate.stat().st_mtime)
    st.success(f'Archivo válido: {len(new):,} registros; {len(new.columns)} columnas; hoja {sheet_new}')
    st.metric('PIM propuesto',money(new['mto_pim'].sum()));st.metric('Devengado propuesto',money(new['devengado'].sum()))
    if not new['NOMBRE'].str.strip().any():st.error('La fuente no identifica institutos. Carga la versión enriquecida o completa primero una matriz de asignación validada. No se publicará esta fuente.')
    elif st.button('Confirmar y publicar corte',type='primary'):
     old=source_path();folder=DATA/'historial';folder.mkdir(exist_ok=True)
     if old.exists() and old.parent==DATA:(folder/(datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+old.name)).write_bytes(old.read_bytes())
     target=DATA/('corte_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.xlsx');target.write_bytes(upload.getvalue())
     with conn() as c:c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('current_file',?)",(str(target),))
     log('PUBLICAR_CORTE');st.cache_data.clear();st.success('Nuevo corte publicado');st.rerun()
   except Exception as e:st.error('Archivo no validado: '+str(e))
 with tabs[1]:
  st.subheader('Institutos y dependencias identificados')
  institutes=df.groupby(['INSTITUTO','NOMBRE'],dropna=False).agg(registros=('mto_pim','size'),PIM=('mto_pim','sum'),Devengado=('devengado','sum')).reset_index().sort_values('PIM',ascending=False)
  st.dataframe(institutes,use_container_width=True,hide_index=True)
  st.download_button('Descargar matriz de identificación',excel_bytes(institutes),'matriz_instituciones.xlsx')
  st.info('Las cuentas institucionales iniciales se vincularon con los códigos modulares del directorio OGESUP de agosto de 2026. Verificar altas y cambios posteriores.')
  mapping=st.file_uploader('Cargar matriz validada (columnas: NOMBRE, CODIGO_MODULAR)',type=['xlsx'],key='map_upload')
  if mapping:
   try:
    mp=pd.read_excel(mapping,dtype=str).fillna('');mp.columns=mp.columns.str.strip().str.upper()
    if not {'NOMBRE','CODIGO_MODULAR'}.issubset(mp.columns):st.error('Faltan NOMBRE y CODIGO_MODULAR.')
    else:
     mp=mp[['NOMBRE','CODIGO_MODULAR']].drop_duplicates();mp['NOMBRE']=mp['NOMBRE'].str.strip();mp['CODIGO_MODULAR']=mp['CODIGO_MODULAR'].str.strip()
     mp=mp[(mp['NOMBRE']!='')&(mp['CODIGO_MODULAR']!='')]
     valid=set(df['NOMBRE']);mp=mp[mp['NOMBRE'].isin(valid)]
     duplicate=mp['CODIGO_MODULAR'].duplicated().any() or mp['NOMBRE'].duplicated().any()
     st.dataframe(mp,use_container_width=True,hide_index=True)
     if duplicate:st.error('Existen códigos o nombres repetidos. Corrige la matriz antes de crear usuarios.')
     elif st.button('Generar cuentas y contraseñas aleatorias'):
      rows=[]
      for _,r in mp.iterrows():
       with conn() as c:existing=c.execute('SELECT 1 FROM accounts WHERE username=?',(r['CODIGO_MODULAR'],)).fetchone()
       if existing:continue
       pw=short_password();add_account(r['CODIGO_MODULAR'],pw,'instituto',r['NOMBRE'])
       rows.append({'CODIGO_MODULAR':r['CODIGO_MODULAR'],'INSTITUTO':r['NOMBRE'],'CLAVE_TEMPORAL':pw})
      if rows:st.session_state['keys_excel']=excel_bytes(pd.DataFrame(rows));log('CREAR_CUENTAS');st.success(f'{len(rows)} cuentas creadas. Descarga y resguarda las credenciales.')
      else:st.info('No hay cuentas nuevas para crear.')
   except Exception as e:st.error(str(e))
  if st.session_state.get('keys_excel'):st.download_button('Descargar credenciales (CONFIDENCIAL)',st.session_state['keys_excel'],'CREDENCIALES_CONFIDENCIAL.xlsx')
  with conn() as c:users=pd.read_sql_query('SELECT username,role,institute,enabled FROM accounts ORDER BY role,username',c)
  st.dataframe(users,use_container_width=True,hide_index=True)
  st.caption('Claves iniciales de institutos: seis letras minúsculas. Las claves del administrador mantienen mayor seguridad.')
  st.markdown('#### Restablecer clave o activar/desactivar cuenta')
  selected=st.selectbox('Cuenta a gestionar',users['username'].tolist(),key='manage_user')
  action=st.radio('Acción',['Restablecer contraseña','Activar cuenta','Desactivar cuenta'],horizontal=True)
  selected_role=users.loc[users['username'].eq(selected),'role'].iloc[0]
  custom_pw=None
  if action=='Restablecer contraseña' and selected_role=='admin':
   st.info('Para el administrador, escribe una nueva contraseña de al menos 12 caracteres. No se generará una clave aleatoria.')
   custom_pw=st.text_input('Nueva contraseña del administrador',type='password',key='custom_admin_password')
   confirm_pw=st.text_input('Confirmar nueva contraseña',type='password',key='confirm_admin_password')
  if st.button('Aplicar cambio',type='primary'):
   if action=='Desactivar cuenta' and selected==account['username']:
    st.error('No puedes desactivar tu propia cuenta.')
   elif action=='Restablecer contraseña':
    if selected_role=='admin' and (not custom_pw or len(custom_pw)<12 or custom_pw!=confirm_pw):
     st.error('La contraseña debe tener al menos 12 caracteres y coincidir con la confirmación.')
    else:
     new_pw=custom_pw if selected_role=='admin' else short_password()
     salt=secrets.token_hex(16)
     with conn() as c:
      c.execute('UPDATE accounts SET salt=?,hash=? WHERE username=?',(salt,hash_password(new_pw,salt),selected))
     st.session_state.pop('reset_credential',None)
     if selected_role=='instituto':
      st.session_state['reset_credential']=excel_bytes(pd.DataFrame([{'USUARIO':selected,'NUEVA_CLAVE_TEMPORAL':new_pw}]))
      st.success('Contraseña del instituto restablecida. Descarga la clave temporal.')
     else:
      st.success('Contraseña del administrador actualizada. Utiliza la nueva clave en tu próximo ingreso.')
     log('RESTABLECER_CLAVE:'+selected)
   else:
    with conn() as c:
     c.execute('UPDATE accounts SET enabled=? WHERE username=?',(1 if action=='Activar cuenta' else 0,selected))
    log(action+':'+selected)
    st.success('Estado actualizado. Actualiza la página para ver el cambio.')
  if st.session_state.get('reset_credential'):st.download_button('Descargar nueva clave (CONFIDENCIAL)',st.session_state['reset_credential'],'Clave_Temporal_Confidencial.xlsx')
 with tabs[2]:
  st.subheader('Diagnóstico del archivo')
  st.metric('Registros',f'{len(df):,}');st.metric('Columnas disponibles',len(df.columns))
  st.write('Registros sin instituto / dependencia:',int(df['NOMBRE'].eq('').sum()))
  st.write('Campos disponibles:',', '.join(df.columns))
  st.download_button('Descargar detalle de calidad',excel_bytes(df[df['NOMBRE'].eq('')]),'registros_sin_instituto.xlsx')
 with tabs[3]:
  with conn() as c:logs=pd.read_sql_query('SELECT * FROM audits ORDER BY at DESC LIMIT 300',c)
  st.dataframe(logs,use_container_width=True,hide_index=True)

 with tabs[4]:
  st.subheader('Historial de archivos publicados')
  folder=DATA/'historial'
  history=sorted(folder.glob('*.xlsx'),reverse=True) if folder.exists() else []
  st.caption('El archivo inicial y el corte vigente se conservan por separado. Los cortes anteriores cargados por el administrador se guardan aquí.')
  st.write('Corte vigente: **'+path.name+'**')
  if history:
   for old in history:
    st.download_button('Descargar '+old.name,old.read_bytes(),old.name,key='history_'+old.name)
  else:st.info('Aún no se han publicado cortes posteriores al archivo inicial.')
