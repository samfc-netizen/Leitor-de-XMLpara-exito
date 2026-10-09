import io
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Painel de Notas Fiscais', page_icon='🧾', layout='wide')
st.markdown('''<style>
.stApp {background:#f4f7fb;color:#14233d} .block-container{padding-top:1.8rem;max-width:1500px}
[data-testid="stMetric"]{background:white;border:1px solid #e4eaf2;padding:18px;border-radius:16px;box-shadow:0 5px 18px #18314d0b}
[data-testid="stMetricLabel"]{color:#53647b} h1,h2,h3{color:#152a46} div[data-testid="stFileUploader"]{background:white;border-radius:14px;padding:12px}
</style>''', unsafe_allow_html=True)

def child(node, *names):
    if node is None:return None
    for name in names:
        for c in node:
            if c.tag.rsplit('}',1)[-1]==name:return c
    return None

def txt(node,*names):
    x=child(node,*names)
    return (x.text or '').strip() if x is not None else ''

def descend(node,*names):
    if node is None:return None
    for name in names:
        for x in node.iter():
            if x.tag.rsplit('}',1)[-1]==name:return x
    return None

def money(value):
    try:return float(str(value).replace(',','.'))
    except:return 0.0

def parse_xml(data, source):
    root=ET.fromstring(data)
    nfse=descend(root,'infNFSe')
    if nfse is not None:
        emit=child(nfse,'emit'); dps=descend(nfse,'infDPS'); prest=child(dps,'prest'); toma=child(dps,'toma')
        vals=child(nfse,'valores'); serv=descend(dps,'vServPrest')
        valor=txt(vals,'vLiq') or txt(serv,'vServ')
        numero=txt(nfse,'nNFSe') or txt(nfse,'nDFSe')
        return {'Tipo':'NFS-e','Número':numero,'Emissor':txt(emit,'xNome') or txt(prest,'xNome') or txt(emit,'CNPJ'),
                'CNPJ Emissor':txt(emit,'CNPJ','CPF') or txt(prest,'CNPJ','CPF'),
                'Destinatário':txt(toma,'xNome') or txt(toma,'CNPJ','CPF') or 'Não identificado',
                'CNPJ Destinatário':txt(toma,'CNPJ','CPF'),
                'Emissão':txt(dps,'dhEmi') or txt(nfse,'dhProc'),'Valor':money(valor),'Arquivo':source,
                'Chave':nfse.attrib.get('Id','')}
    inf=descend(root,'infNFe')
    if inf is not None:
        ide=child(inf,'ide'); emit=child(inf,'emit'); dest=child(inf,'dest'); total=descend(inf,'ICMSTot')
        return {'Tipo':'NF-e','Número':txt(ide,'nNF'),'Emissor':txt(emit,'xNome') or txt(emit,'CNPJ'),
                'CNPJ Emissor':txt(emit,'CNPJ','CPF'),'Destinatário':txt(dest,'xNome') or txt(dest,'CNPJ','CPF') or 'Não identificado',
                'CNPJ Destinatário':txt(dest,'CNPJ','CPF'),'Emissão':txt(ide,'dhEmi','dEmi'),
                'Valor':money(txt(total,'vNF')),'Arquivo':source,'Chave':inf.attrib.get('Id','')}
    raise ValueError('Formato XML não reconhecido como NF-e ou NFS-e')

def read_files(uploaded):
    rows=[]; errors=[]
    for f in uploaded:
        data=f.getvalue()
        if zipfile.is_zipfile(io.BytesIO(data)):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for member in z.infolist():
                    if member.is_dir() or not member.filename.lower().endswith('.xml'):continue
                    if member.file_size>15_000_000:
                        errors.append((member.filename,'Arquivo excede 15 MB'));continue
                    try:rows.append(parse_xml(z.read(member),member.filename))
                    except Exception as e:errors.append((member.filename,str(e)))
        elif f.name.lower().endswith('.xml'):
            try:rows.append(parse_xml(data,f.name))
            except Exception as e:errors.append((f.name,str(e)))
    return rows,errors

def brl(v):return 'R$ '+f'{v:,.2f}'.replace(',','X').replace('.',',').replace('X','.')

def excel_bytes(df,by_client):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,sheet_name='Notas')
        by_client.to_excel(writer,index=False,sheet_name='Destinatarios')
    return out.getvalue()

st.title('🧾 Painel de Notas Fiscais')
st.caption('Leitura de NF-e e NFS-e • Importação de XML ou ZIP • Consolidação automática')
files=st.file_uploader('Selecione um ZIP com a pasta de XMLs ou múltiplos arquivos XML',type=['zip','xml'],accept_multiple_files=True)
if not files:
    st.info('Envie seus arquivos acima para visualizar o painel. O ZIP pode conter subpastas.')
    st.stop()
rows,errors=read_files(files)
if not rows:
    st.error('Nenhuma nota válida foi identificada.')
    if errors:st.dataframe(pd.DataFrame(errors,columns=['Arquivo','Erro']),hide_index=True)
    st.stop()
df=pd.DataFrame(rows).drop_duplicates(subset=['Tipo','Chave','CNPJ Emissor','Número'],keep='first')
df['Data']=pd.to_datetime(df['Emissão'],errors='coerce',utc=True).dt.tz_convert('America/Sao_Paulo').dt.date
with st.sidebar:
    st.header('Filtros')
    emis=st.multiselect('Emissor',sorted(df['Emissor'].unique()),default=sorted(df['Emissor'].unique()))
    dest=st.multiselect('Destinatário',sorted(df['Destinatário'].unique()),default=sorted(df['Destinatário'].unique()))
    tipos=st.multiselect('Tipo de nota',sorted(df['Tipo'].unique()),default=sorted(df['Tipo'].unique()))
    busca=st.text_input('Buscar número da nota')
    dates=df['Data'].dropna()
    periodo=st.date_input('Período de emissão',value=(dates.min(),dates.max())) if len(dates) else None
mask=df['Emissor'].isin(emis)&df['Destinatário'].isin(dest)&df['Tipo'].isin(tipos)
if busca:mask &= df['Número'].astype(str).str.contains(busca,case=False,regex=False)
if periodo is not None and isinstance(periodo,(tuple,list)) and len(periodo)==2:
    mask &= df['Data'].between(periodo[0],periodo[1])
f=df[mask].copy()
clients=f.groupby(['CNPJ Destinatário','Destinatário'],dropna=False).agg(Notas=('Número','size'),Faturamento=('Valor','sum')).reset_index().sort_values(['Faturamento','Notas'],ascending=False)
clients.insert(0,'Participação %', (clients['Faturamento']/f['Valor'].sum()*100).round(2) if f['Valor'].sum() else 0)
a,b,c,d=st.columns(4)
a.metric('Notas emitidas',f'{len(f):,}'.replace(',','.'))
b.metric('Valor total das notas',brl(f['Valor'].sum()))
c.metric('Destinatários distintos',len(clients))
d.metric('Ticket médio por nota',brl(f['Valor'].mean() if len(f) else 0))
st.divider()
left,right=st.columns([1.4,1])
with left:
    st.subheader('Faturamento por destinatário')
    top=clients.head(12).sort_values('Faturamento')
    if not top.empty:
        st.bar_chart(top.set_index('Destinatário')['Faturamento'], horizontal=True, use_container_width=True)
with right:
    st.subheader('Quantidade de notas por destinatário')
    topq=clients.nlargest(10,'Notas')
    if not topq.empty:
        st.bar_chart(topq.set_index('Destinatário')['Notas'], horizontal=True, use_container_width=True)
st.subheader('Evolução do faturamento')
valid=f.dropna(subset=['Data']).copy()
if not valid.empty:
    valid['Mês']=pd.to_datetime(valid['Data']).dt.to_period('M').astype(str)
    evolution=valid.groupby('Mês',as_index=False).agg(Faturamento=('Valor','sum'),Notas=('Número','size'))
    st.bar_chart(evolution.set_index('Mês')['Faturamento'], use_container_width=True)
st.subheader('Resumo por destinatário')
show=clients[['Destinatário','CNPJ Destinatário','Notas','Faturamento','Participação %']]
st.dataframe(show,use_container_width=True,hide_index=True,column_config={'Faturamento':st.column_config.NumberColumn(format='R$ %.2f'),'Participação %':st.column_config.NumberColumn(format='%.2f%%')})
st.subheader('Todas as notas fiscais')
detail=f[['Tipo','Número','Data','Emissor','CNPJ Emissor','Destinatário','CNPJ Destinatário','Valor','Arquivo']].sort_values('Data',ascending=False)
st.dataframe(detail,use_container_width=True,hide_index=True,column_config={'Valor':st.column_config.NumberColumn(format='R$ %.2f')})
col1,col2=st.columns(2)
with col1:st.download_button('⬇️ Exportar relatório Excel',excel_bytes(detail,show),'relatorio_notas.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
with col2:st.download_button('⬇️ Exportar notas CSV',detail.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig'),'notas_fiscais.csv','text/csv',use_container_width=True)
if errors:
    with st.expander(f'⚠️ {len(errors)} arquivos não processados'):
        st.dataframe(pd.DataFrame(errors,columns=['Arquivo','Motivo']),hide_index=True,use_container_width=True)
st.caption(f'{len(rows)} XML(s) lidos • {len(df)} nota(s) únicas • {len(errors)} erro(s). Dados processados na sessão atual.')
