#coding=utf_8
from cProfile import label
from email.mime import image
from hashlib import algorithms_available
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter.filedialog import askopenfilename
from turtle import left
import pandas as pd
import geocoder
import geopy
import tkintermapview


root = Tk()
x ='SEM ARQUIVO'

#FUNÇÕES
def subir_arquivo():
    global df1
    global x
    x = filedialog.askopenfilename()
    Lb_arq['text'] = x
    
def limpar():
    tv1.delete(*tv1.get_children())

#TIM CHAMADAS
def tim_chamadas():
    limpar()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    global dff
    global x
    
    #importação conexões
    df1 = pd.read_excel(f'{x}',sheet_name=1)
    df1 = df1.drop(columns=['REGISTRO','CENTRAL','REDIRECIONADO PARA','ÚLTIMA CGI/ERB','HORA BRASÍLIA','TIPO'])
    df1['Nº ORIGEM'] = df1['Nº ORIGEM'].astype(object)
    df1['Nº DESTINO'] = df1['Nº DESTINO'].astype(object)
    df1.rename(columns={'PRIMEIRA CGI/ERB':'CGI/ERB'}, inplace=True)

    #importação dados das antenas
    df2 = pd.read_excel(f'{x}',sheet_name=4)
    df2['ENDEREÇO ALVO'] = df2['ENDEREÇO']+', '+df2['BAIRRO']+', '+df2['CIDADE']+', '+df2['ESTADO']+' - Azimute '+df2['AZIMUTE']
    df2 = df2.drop(columns=['LONGITUDE','LATITUDE','DATA INÍCIO','NOME','AZIMUTE','ENDEREÇO','BAIRRO','CIDADE','CEP','ESTADO'])

    #planilha final
    dff = df1.merge(df2[['CGI/ERB','ENDEREÇO ALVO']], on='CGI/ERB', how='left')
    dff = dff.drop(columns=['CGI/ERB','IMEI DESTINO','IMEI ORIGEM'])

def tim_ch_imei():
    global dff
    global x
    limpar()

    #importação conexões
    df1 = pd.read_excel(f'{x}',sheet_name=1)
    df1 = df1.drop(columns=['REGISTRO','CENTRAL','REDIRECIONADO PARA','ÚLTIMA CGI/ERB','HORA BRASÍLIA','TIPO'])
    df1['Nº ORIGEM'] = df1['Nº ORIGEM'].astype(object)
    df1['Nº DESTINO'] = df1['Nº DESTINO'].astype(object)
    df1.rename(columns={'PRIMEIRA CGI/ERB':'CGI/ERB'}, inplace=True)

    #importação dados das antenas
    df2 = pd.read_excel(f'{x}',sheet_name=4)
    df2['ENDEREÇO ALVO'] = df2['ENDEREÇO']+', '+df2['BAIRRO']+', '+df2['CIDADE']+', '+df2['ESTADO']+' - Azimute '+df2['AZIMUTE']
    df2 = df2.drop(columns=['LONGITUDE','LATITUDE','DATA INÍCIO','NOME','AZIMUTE','ENDEREÇO','BAIRRO','CIDADE','CEP','ESTADO'])

    #planilha final
    dff = df1.merge(df2[['CGI/ERB','ENDEREÇO ALVO']], on='CGI/ERB', how='left')
    dff = dff.drop(columns=['CGI/ERB'])
    
    #identificando alvo
    k = dff['Nº ORIGEM']
    y = dff['Nº DESTINO']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    df5 = dff[dff['Nº ORIGEM'] == alvo]
    df6 = dff[dff['Nº DESTINO'] == alvo]
    dff = pd.DataFrame((pd.concat([df5['IMEI ORIGEM'],df6['IMEI DESTINO']])).unique())

def tim_ch_loc():
    global dff
    limpar()
    #importação conexões
    df1 = pd.read_excel(f'{x}',sheet_name=1)
    df1 = df1.drop(columns=['REGISTRO','CENTRAL','REDIRECIONADO PARA','ÚLTIMA CGI/ERB','HORA BRASÍLIA','TIPO'])
    df1['Nº ORIGEM'] = df1['Nº ORIGEM'].astype(object)
    df1['Nº DESTINO'] = df1['Nº DESTINO'].astype(object)
    df1.rename(columns={'PRIMEIRA CGI/ERB':'CGI/ERB'}, inplace=True)

    #importação dados das antenas
    df2 = pd.read_excel(f'{x}',sheet_name=4)
    df2['ENDEREÇO ALVO'] = df2['ENDEREÇO']+', '+df2['BAIRRO']+', '+df2['CIDADE']+', '+df2['ESTADO']+' - Azimute '+df2['AZIMUTE']
    df2 = df2.drop(columns=['LONGITUDE','LATITUDE','DATA INÍCIO','NOME','AZIMUTE','ENDEREÇO','BAIRRO','CIDADE','CEP','ESTADO'])

    #planilha final
    dff = df1.merge(df2[['CGI/ERB','ENDEREÇO ALVO']], on='CGI/ERB', how='left')
    dff = dff.drop(columns=['CGI/ERB','IMEI DESTINO','IMEI ORIGEM'])
    dff = pd.DataFrame(dff['ENDEREÇO ALVO'].value_counts())
    dff = dff.reset_index()

def tim_mapa():
    global dff
    global x

    #importação conexões
    df1 = pd.read_excel(f'{x}',sheet_name=1)
    df1 = df1.drop(columns=['REGISTRO','CENTRAL','REDIRECIONADO PARA','ÚLTIMA CGI/ERB','HORA BRASÍLIA','TIPO'])
    df1['Nº ORIGEM'] = df1['Nº ORIGEM'].astype(object)
    df1['Nº DESTINO'] = df1['Nº DESTINO'].astype(object)
    df1.rename(columns={'PRIMEIRA CGI/ERB':'CGI/ERB'}, inplace=True)
    
    #importações planilha
    antenas = pd.read_excel(f'{x}',sheet_name=4)
    antenas = antenas.drop(columns=['DATA INÍCIO','NOME','AZIMUTE','ENDEREÇO','BAIRRO','CIDADE','CEP','ESTADO'])
    antenas.rename(columns={'PRIMEIRA CGI/ERB':'CGI/ERB'},inplace=True)
    planilha_1 = df1.drop(columns = ['Nº ORIGEM','Nº DESTINO','IMEI ORIGEM','IMEI DESTINO','HORA LOCAL','DURAÇÃO'])

    #remover dados com '.'
    a_remover = antenas.loc[(antenas['LATITUDE'] == '.') | (antenas['LONGITUDE'] == '.')]
    antenas = antenas.drop(a_remover.index).dropna()

    #planilha final
    antenas_con_lat = planilha_1.merge(antenas[['CGI/ERB','LATITUDE']], on='CGI/ERB', how='left')
    antenas_con_lat_long = antenas_con_lat.merge(antenas[['CGI/ERB','LONGITUDE']], on='CGI/ERB', how='left').dropna()

    antenas_con_lat_long['LATITUDE'] = antenas_con_lat_long['LATITUDE'].astype(str)
    antenas_con_lat_long['LONGITUDE'] = antenas_con_lat_long['LONGITUDE'].astype(str)
    antenas_con_lat_long['LATITUDE'] = antenas_con_lat_long['LATITUDE'].str.replace(',', '.')
    antenas_con_lat_long['LONGITUDE'] = antenas_con_lat_long['LONGITUDE'].str.replace(',', '.')
    antenas_con_lat_long = antenas_con_lat_long.drop(columns=['CGI/ERB']) 
    df_dados = pd.DataFrame(antenas_con_lat_long.value_counts())
    df_dados.rename(columns={0:'reincidencia'}, inplace=True)
    df_dados = df_dados.reset_index()

    #endereço inicial do mapa
    end_ini = pd.DataFrame(antenas_con_lat_long.value_counts())
    end_ini = end_ini.reset_index()
    lat_ini = float(end_ini.iloc[0,0])
    long_ini = float(end_ini.iloc[0,1])
    
    #plotar mapa
    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000, height=700, corner_radius=0)
    map.pack()
    map.set_position(lat_ini,long_ini) #posição inicial
    map.set_zoom(100)
    for i in df_dados.itertuples():
        map.set_marker(float(i.LATITUDE), float(i.LONGITUDE), marker_color_circle='black', marker_color_outside='darkblue')
    map.set_marker(lat_ini, long_ini, text='MAIOR REINCIDÊNCIA')
        
#TIM IP
def tim_ip():
    limpar()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    global x
    global dff

    #importação conexões
    df1 = pd.read_excel(f'{x}',sheet_name=2)
    alvo = df1.iloc[1,0]
    df1 = df1.drop(columns=['MSISDN','SERVICE_ID','IPV4','PORTA INÍCIO','PORTA FIM','IPV6','UPLINK','DOWNLINK', 'HORA BRASÍLIA (INÍCIO)','HORA BRASÍLIA (FIM)','HORA LOCAL (FIM)','TIPO'])
    df1.rename(columns={'CÉLULA (CGI)':'CGI/ERB'}, inplace=True)
    
    #importação dados das antenas
    df2 = pd.read_excel(f'{x}',sheet_name=3)
    df2['dados'] = df2['ENDEREÇO']+', '+df2['BAIRRO']+', '+df2['CIDADE']+', '+df2['ESTADO']+' - Azimute '+df2['AZIMUTE']
    df2 = df2.drop(columns=['LONGITUDE','LATITUDE','DATA INÍCIO','NOME','AZIMUTE','ENDEREÇO','BAIRRO','CIDADE','CEP','ESTADO'])

    #planilha final
    dff = df1.merge(df2[['CGI/ERB','dados']], on='CGI/ERB', how='left')
    dff = dff.drop(columns=['CGI/ERB'])
    
def tim_ip_loc():
    limpar()
    global dff
    tim_ip()
    dff = pd.DataFrame(dff['dados'].value_counts())
    dff = dff.reset_index()

def tim_ip_mapa():
    #importações
    antenas = pd.read_excel(f'{x}',sheet_name=3)
    antenas
    #remover dados com '.'
    a_remover = antenas.loc[(antenas['LATITUDE'] == '.') | (antenas['LONGITUDE'] == '.')]
    antenas = antenas.drop(a_remover.index).dropna()
    antenas

    # antenas['LATITUDE'] = antenas['LATITUDE'].astype(str)
    # antenas['LONGITUDE'] = antenas['LONGITUDE'].astype(str)
    antenas['LATITUDE'] = antenas['LATITUDE'].str.replace(',', '.')
    antenas['LONGITUDE'] = antenas['LONGITUDE'].str.replace(',', '.')
    df_dados = pd.DataFrame(antenas.value_counts())
    df_dados = df_dados.reset_index()

    #endereço inicial do mapa
    end_ini = pd.DataFrame(antenas.value_counts())
    end_ini = end_ini.reset_index()
    lat_ini = float(end_ini.iloc[0,1])
    long_ini = float(end_ini.iloc[0,2])

    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000,height=700, corner_radius=0)
    map.pack()
    map.set_position(lat_ini, long_ini) #posição inicial
    map.set_zoom(100)
    for i in df_dados.itertuples():
        map.set_marker(float(i.LATITUDE), float(i.LONGITUDE), marker_color_circle='black', marker_color_outside='darkblue')
    map.set_marker(lat_ini, long_ini, text='MAIOR REINCIDÊNCIA')

#VIVO CHAMADAS
def vivo_chamadas():
    limpar()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    
    global dff
    global x
    #importação chamadas
    df1 = pd.read_excel(f'{x}', header=5)
    df1 = df1.drop(columns=['Desc','IMEI Chamado','IMEI Chamador','Hora desc','Status','Tra','Durac','GMT Origem', 'GMT Destino','Tec Origem','Tec Destino'])
    df1.rename(columns={'Local Destino':'CGI'}, inplace=True)

    # planilha das antenas
    df2 = pd.read_excel(f'{x}',sheet_name=2, header=5)
    df2 = df2.drop(columns=['CCC','Sigla','ERB','Sigla.1','Set','Tecnologia'])
    df2['dados'] = df2['Endereço']+', '+df2['Bairro']+', '+df2['Cidade']+', '+df2['UF']#+' - Azimute '+df2['Azi']

    # planilha mesclada chamado
    df3 = df1.merge(df2[['CGI','dados']], on='CGI', how='left') 

    # planilha dados chamador
    df4 = pd.read_excel(f'{x}', header=5)
    df4 = df4.drop(columns=['IMEI Chamado','Desc','IMEI Chamador','Hora desc','Status','Tra','Durac','GMT Origem', 'GMT Destino','Tec Origem','Tec Destino'])
    df4.rename(columns={'Local Origem':'CGI'}, inplace=True)

    # planilha mesclada Chamador
    df5 = df4.merge(df2[['CGI','dados']], on='CGI', how='left') 

    #planilha final
    df6 = pd.DataFrame(df1[['Data','Hora','Chamador','Chamado','Local Origem']])
    df6.insert(5,'Local Destino',df4['Local Destino'])
    df6.insert(6, 'Erb Origem', df5['dados'])
    df6.insert(7, 'Erb Destino', df3['dados'])
    dff = df6.drop(columns=['Local Origem','Local Destino'])

def vivo_ch_imei():
    limpar()
    global dff
    global x
    vivo_chamadas()
    #identificando alvo
    k = dff['Chamado']
    y = dff['Chamador']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]
    
    #identificando imeis
    df10 = pd.read_excel(f'{x}',header=5)
    df10['IMEI Chamador'] = df10['IMEI Chamador'].astype(str)
    df10['IMEI Chamado'] = df10['IMEI Chamado'].astype(str)
    df8 = df10[df10['Chamador'] == alvo]
    df9 = df10[df10['Chamado'] == alvo]
    dff = pd.DataFrame((pd.concat([df8['IMEI Chamador'],df9['IMEI Chamado']])).unique())

def vivo_ch_loc():
    limpar()
    global dff
    global x
    vivo_chamadas()
    #pegando numero com maior reincidência (alvo)
    k = dff['Chamado']
    y = dff['Chamador']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #calculando
    df1= dff[dff['Chamado'] == alvo]
    df2 = dff[dff['Chamador'] == alvo]
    a = df1['Erb Destino']
    b = df2['Erb Origem']
    c = pd.DataFrame(pd.concat([a,b],axis=0).value_counts())
    dff = c.reset_index()

def vivo_mapa():
    global dff
    global x
    limpar()        
    vivo_chamadas()
    
    #pegando alvo
    k = dff['Chamado']
    y = dff['Chamador']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #separando as coordenadas
    q = dff[dff['Chamador'] == alvo]
    p = q['Erb Origem']
    r = dff[dff['Chamado'] == alvo]
    s = r['Erb Destino']
    EndCon = pd.DataFrame(pd.concat([p,s],axis=0)).dropna()
    EndCon_uni = pd.DataFrame(EndCon[0].unique())

    #pegando a coodenada referência para a abertura do mapa
    x = pd.DataFrame(EndCon.value_counts())
    x = x.rename(columns={0:'Reincidência'})
    x = x.reset_index()
    MaiRei = x.iloc[0,0]
        
    #transformando código end em coo
    plan_con = []
    for i in EndCon_uni.itertuples():
        codigo = geocoder.osm(i)
        codigo_fim = [codigo.lat, codigo.lng]
        plan_con.append(codigo_fim)
            
    plan_coo = pd.DataFrame(plan_con).dropna()
    plan_coo = plan_coo.rename(columns={0:'LAT',1:"LON"})
    
    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000,height=700, corner_radius=0)
    map.pack()
    map.set_address(MaiRei, marker=True) #posição inicial
    map.set_zoom(100)
    for i in plan_coo.itertuples():
        map.set_marker(i.LAT, i.LON, marker_color_circle='black', marker_color_outside='darkblue')
    map.set_marker(encod.lat, encod.lng, text='MAIOR REINCIDÊNCIA')
    print(MaiRei)
#CLARO CHAMADAS
def claro_chamadas():
    limpar()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    global x
    global dff

    #importação planilha dados de loc
    df1 = pd.read_excel(f'{x}')
    df1 = df1.drop(columns=['SiteID RE','Raio RE','Rota Entrada','SiteID RS','SetorID RS', 'Raio RS','Rota Saída','Central','Tipo Registro','Tipo Rota'])
    df1['Azimute RE'] = df1['Azimute RE'].astype(str)
    df1['Azimute RS'] = df1['Azimute RS'].astype(str)
    df1['Número A'] = df1['Número A'].astype(str)
    df1['Número B'] = df1['Número B'].astype(str)
    df1['Endereço Chamador'] = df1['Endereço ERB RE'] +', ' + df1['Município ERB RE'] + ' - Azimute: ' + df1['Azimute RE']
    df1['Endereço Chamado'] = df1['Endereço ERB RS'] +', ' + df1['Município ERB RS'] + ' - Azimute: ' + df1['Azimute RS']
    dff = df1.drop(columns=['Endereço ERB RE','Município ERB RE','Azimute RE','Endereço ERB RS','Município ERB RS','Azimute RS','IMEI A',
    'IMEI B','Nome A','CNPJ CPF A','Identidade A','Endereço A','Cidade A','UF A','CEP A','Data Ativação A',
    'Data Cancel A','Identidade B','Nome B','CNPJ CPF B','Endereço B','Status','Cidade B','UF B','CEP B','Data Ativação B',
    'Data Cancel B','Latitude RE','Longitude RE','Latitude RS','Longitude RS','SetorID RE'])

def claro_ch_imei():
    limpar()
    global dff
    df1 = pd.read_excel(f'{x}')
    k = df1['Número A']
    y = df1['Número B']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    df1['IMEI A'] = df1['IMEI A'].astype(str)
    df1['IMEI B'] = df1['IMEI B'].astype(str)

    df2 = df1[df1['Número A'] == alvo] 
    df3 = df1[df1['Número B'] == alvo] 
    dff = pd.DataFrame((pd.concat([df2['IMEI A'],df3['IMEI B']])).unique())

def claro_ch_loc():
    limpar()
    global dff
    global x
    claro_chamadas()
    k = dff['Número A']
    y = dff['Número B']
    z = pd.DataFrame(pd.concat([k,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #calculando
    df2 = dff[dff['Número A'] == alvo]
    df3 = dff[dff['Número B'] == alvo]
    a = df2['Endereço Chamador']
    b = df3['Endereço Chamado']
    c = pd.DataFrame(pd.concat([a,b],axis=0).value_counts())
    dff = c.reset_index()

def claro_ch_map():
    global x
    global dff
    df1 = pd.read_excel(f'{x}')
    df1 = df1.drop(columns=['SiteID RE','Raio RE','Rota Entrada','SiteID RS','SetorID RS', 'Raio RS','Rota Saída','Central','Tipo Registro','Tipo Rota'])
    df1['Azimute RE'] = df1['Azimute RE'].astype(str)
    df1['Azimute RS'] = df1['Azimute RS'].astype(str)
    df1['Número A'] = df1['Número A'].astype(str)
    df1['Número B'] = df1['Número B'].astype(str)
    df1['Endereço Chamador'] = df1['Endereço ERB RE'] +', ' + df1['Município ERB RE'] + ' - Azimute: ' + df1['Azimute RE']
    df1['Endereço Chamado'] = df1['Endereço ERB RS'] +', ' + df1['Município ERB RS'] + ' - Azimute: ' + df1['Azimute RS']
    dff = df1.drop(columns=['Endereço ERB RE','Município ERB RE','Azimute RE','Endereço ERB RS','Município ERB RS','Azimute RS','IMEI A',
    'IMEI B','Nome A','CNPJ CPF A','Identidade A','Endereço A','Cidade A','UF A','CEP A','Data Ativação A',
    'Data Cancel A','Identidade B','Nome B','CNPJ CPF B','Endereço B','Status','Cidade B','UF B','CEP B','Data Ativação B',
    'Data Cancel B','SetorID RE'])
    
    #pegando numero com maior reincidência (alvo)
    q = dff['Número A']
    y = dff['Número B']
    z = pd.DataFrame(pd.concat([q,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #trabalhando as coordenadas do alvo
    df2 = dff[dff['Número A'] == alvo]
    df3 = dff[dff['Número B'] == alvo]
    LatA = df2['Latitude RE']
    LatB = df3['Latitude RS']
    LonA = df2['Longitude RE'] 
    LonB = df3['Longitude RS']
    LatCon = pd.DataFrame(pd.concat([LatA,LatB],axis=0))
    LatCon = LatCon.rename(columns={0:'Latitude'})
    LonCon = pd.DataFrame(pd.concat([LonA,LonB],axis=0))
    LonCon = LonCon.rename(columns={0:'Longitude'})
    Coordenadas = pd.DataFrame(pd.concat([LatCon,LonCon],axis=1)).dropna()
    Coordenadas = Coordenadas.dropna()

    #pegando a coodenada referência para a abertura do mapa
    r = pd.DataFrame(Coordenadas.value_counts())
    r = r.reset_index()
    x1 = r.iloc[0,0]
    x2 = r.iloc[0,1]

    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000, height=700, corner_radius=0)
    map.pack()
    map.set_position(float(x1),float(x2)) #posição inicial
    map.set_zoom(100)
    for i in Coordenadas.itertuples():
        map.set_marker(float(i.Latitude), float(i.Longitude), marker_color_circle='black', marker_color_outside='darkblue')
    map.set_marker(float(x1), float(x2), text='MAIOR REINCIDÊNCIA')

#CLARO CHAMADAS SEM CADASTRO

def claro_sc_ch():
    global dff
    global x
    limpar()
    #importação planilha dados de loc
    df1 = pd.read_excel(f'{x}')
    df1 = df1.drop(columns=['SiteID RE','Raio RE','Rota Entrada','SiteID RS','SetorID RS', 'Raio RS','Rota Saída','Central','Tipo Registro','Tipo Rota'])
    df1['Azimute RE'] = df1['Azimute RE'].astype(str)
    df1['Azimute RS'] = df1['Azimute RS'].astype(str)
    df1['Número A'] = df1['Número A'].astype(str)
    df1['Número B'] = df1['Número B'].astype(str)
    df1['Endereço Chamador'] = df1['Endereço ERB RE'] +', ' + df1['Município ERB RE'] + ' - Azimute: ' + df1['Azimute RE']
    df1['Endereço Chamado'] = df1['Endereço ERB RS'] +', ' + df1['Município ERB RS'] + ' - Azimute: ' + df1['Azimute RS']
    dff = df1.drop(columns=['Endereço ERB RE','Município ERB RE','Azimute RE','Endereço ERB RS','Município ERB RS','Azimute RS','IMEI A',
    'IMEI B','Latitude RE','Longitude RE','Latitude RS','Longitude RS','SetorID RE'])
   
def claro_sc_loc():
    global x
    global dff
    limpar()
    claro_sc_ch()
    #pegando numero com maior reincidência (alvo)
    q = dff['Número A']
    y = dff['Número B']
    z = pd.DataFrame(pd.concat([q,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #calculando
    df2 = dff[dff['Número A'] == alvo]
    df3 = dff[dff['Número B'] == alvo]
    a = df2['Endereço Chamador']
    b = df3['Endereço Chamado']
    c = pd.DataFrame(pd.concat([a,b],axis=0).value_counts())
    c.rename(columns={'index':'Endereços', 0:'Reincidência'}, inplace=True)
    dff = c.reset_index()

def claro_sc_imei():
    global x
    global dff
    limpar()
    dff = pd.read_excel(f'{x}')
    q = dff['Número A']
    y = dff['Número B']
    z = pd.DataFrame(pd.concat([q,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]
    
    dff['IMEI A'] = dff['IMEI A'].astype(str)
    dff['IMEI B'] = dff['IMEI B'].astype(str)
    df2 = dff[dff['Número A'] == alvo] 
    df3 = dff[dff['Número B'] == alvo] 
    dff = pd.DataFrame((pd.concat([df2['IMEI A'],df3['IMEI B']])).unique())
        
def claro_sc_mapa():
    global x
    global dff
        
    dff = pd.read_excel(f'{x}')

    q = dff['Número A']
    y = dff['Número B']
    z = pd.DataFrame(pd.concat([q,y])).value_counts()
    c = pd.DataFrame(data=z, columns=['Reincidencia'])
    c = c.reset_index()
    alvo = c.iloc[0,0]

    #trabalhando as coordenadas do alvo
    df2 = dff[dff['Número A'] == alvo]
    df3 = dff[dff['Número B'] == alvo]
    LatA = df2['Latitude RE']
    LatB = df3['Latitude RS']
    LonA = df2['Longitude RE'] 
    LonB = df3['Longitude RS']
    LatCon = pd.DataFrame(pd.concat([LatA,LatB],axis=0))
    LatCon = LatCon.rename(columns={0:'Latitude'})
    LonCon = pd.DataFrame(pd.concat([LonA,LonB],axis=0))
    LonCon = LonCon.rename(columns={0:'Longitude'})
    Coordenadas = pd.DataFrame(pd.concat([LatCon,LonCon],axis=1)).dropna()
    Coordenadas = Coordenadas.dropna()

    #pegando a coodenada referência para a abertura do mapa
    x = pd.DataFrame(Coordenadas.value_counts())
    x = x.reset_index()
    x1=x.iloc[0,0]
    x2=x.iloc[0,1]

    #plotando mapa
    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000, height=700, corner_radius=0)
    map.pack()
    map.set_position(float(x1),float(x2)) #posição inicial
    map.set_zoom(100)
    for i in Coordenadas.itertuples():
        map.set_marker(float(i.Latitude), float(i.Longitude), marker_color_circle='black', marker_color_outside='darkblue')
    map.set_marker(float(x1), float(x2), text='MAIOR REINCIDÊNCIA')

#CLARO IP
def claro_ip():
    limpar()
    global x
    global dff
    dff = pd.read_excel(f'{x}')
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    dff = dff[['Data Início', 'Hora Início', 'Endereço ERB RS',  'Município ERB RS']]

def claro_ip_loc():
    limpar()
    global x
    global dff
    dff = pd.read_excel(f'{x}')
    dff = pd.DataFrame((dff['Endereço ERB RS'] + ', ' + dff['Município ERB RS']).value_counts())
    dff = dff.reset_index()

def claro_ip_imei():
    limpar()
    global x
    global dff
    dff = pd.read_excel(f'{x}')
    dff = pd.DataFrame(dff['IMEI A'].unique())
           
#VIVO IP
def vivo_IP():
    limpar()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_colwidth', None)
    global dff
    global x

    dff = pd.read_excel(f'{x}',header=5)
    dff = dff.drop(columns=['IP', 'IPv6','Even.Ini','Data Fim','Hora Fim','Even.Fim','Vol.Ent','Vol.Sai','Status', 'Tráfego','Serviço','Localização'])
    
def vivo_ip_imei():
    limpar()
    vivo_IP()
    global dff
    dff = pd.DataFrame(dff['IMEI'].unique())
    
def vivo_ip_loc():
    limpar()
    global dff
    vivo_IP()
    dff = pd.DataFrame(dff['Endereço'].value_counts())
    dff = dff.reset_index()

def vivo_ip_mapa():
    vivo_IP()
    limpar()
    global dff
    global x
    df = pd.read_excel(f'{x}',header=5)
    df = df['Endereço'].str.split('-',expand=True)
    dff = df[3]+','+df[2]+','+df[1]+','+df[0]
    dff.dropna(inplace=True)
    dff = pd.DataFrame(dff)

    #pegando a coodenada referência para a abertura do mapa
    x = pd.DataFrame(dff[0].value_counts())
    x = x.rename(columns={0:'Reincidência'})
    x = x.reset_index()
    MaiRei = x.iloc[0,0]
    
    dff=dff.drop_duplicates()
    my_map = Toplevel(root)
    map = tkintermapview.TkinterMapView(my_map, width=1000,height=700, corner_radius=0)
    map.pack()
    map.set_address(MaiRei, marker=True, text='MAIOR REINCIDÊNCIA') #posição inicial
    map.set_zoom(100)
    for i in dff.itertuples():
        map.set_address(i, marker_color_circle='black', marker_color_outside='darkblue',marker=True, text='')
    
    

#config inicial
root.title('Desenvolvido por DANIEL ROQUE')
root.geometry('1195x700')
root.resizable(0,0)
root.iconbitmap('icon.ico')
root['bg'] = '#363636'
foto_draco = PhotoImage(file='draco1.png')
foto_claro = PhotoImage(file='claro.png')
foto_tim = PhotoImage(file='tim.png')
foto_vivo = PhotoImage(file='vivo.png')
a = IntVar()
b = IntVar()

def ver_radio():
    index_a = a.get()
    index_b = b.get()
    if (index_a == 1) & (index_b == 1):
        claro_chamadas()
    elif (index_a == 1) & (index_b == 2):
        claro_ch_imei()
    elif (index_a == 1) & (index_b == 3):
        claro_ch_loc()
    elif (index_a == 1) & (index_b == 4):
        claro_ch_map()
    elif (index_a == 7) & (index_b == 1):
        claro_sc_ch()
    elif (index_a == 7) & (index_b == 2):
        claro_sc_imei()
    elif (index_a == 7) & (index_b == 3):
        claro_sc_loc()
    elif (index_a == 7) & (index_b == 4):
        claro_sc_mapa()
    elif (index_a == 2) & (index_b == 1):
        tim_chamadas()
    elif (index_a == 2) & (index_b == 2):
        tim_ch_imei()
    elif (index_a == 2) & (index_b == 3):
        tim_ch_loc()
    elif (index_a == 3) & (index_b == 1):
        tim_ip()
    elif (index_a == 2) & (index_b == 4):
        tim_mapa()
    elif (index_a == 3) & (index_b == 2):
        messagebox.showerror('erro','Planilha de IP da TIM não retorna IMEI')
    elif (index_a == 3) & (index_b == 3):
        tim_ip_loc()
    elif (index_a == 3) & (index_b == 4):
        tim_ip_mapa()
    elif (index_a == 6) & (index_b == 1):
        claro_ip()
    elif (index_a == 6) & (index_b == 2):
        claro_ip_imei()
    elif (index_a == 6) & (index_b == 3):
        claro_ip_loc()
    elif (index_a == 4) & (index_b == 1):
        vivo_chamadas()
    elif (index_a == 4) & (index_b == 2):
        vivo_ch_imei()
    elif (index_a == 4) & (index_b == 3):
        vivo_ch_loc()
    elif (index_a ==4) & (index_b == 4):
        vivo_mapa()
    elif (index_a == 5) & (index_b == 1):
        vivo_IP()
    elif (index_a == 5) & (index_b == 2):
        vivo_ip_imei()
    elif (index_a == 5) & (index_b == 3):
        vivo_ip_loc()
    elif (index_a == 5) & (index_b == 4):
        vivo_ip_mapa()

    #config da planilha
    tv1['column']=list(dff.columns)
    tv1['show']='headings'
    for column in tv1['columns']:
        tv1.heading(column, text=column)

    dff_row = dff.to_numpy().tolist()
    for row in dff_row:
        tv1.insert('','end',values=row)

#widgets    
imagem_draco = Label(root, image= foto_draco, borderwidth=False)
imagem_claro = Label(root, image= foto_claro, borderwidth=False)
imagem_tim = Label(root, image= foto_tim, borderwidth=False)
imagem_vivo = Label(root, image= foto_vivo, borderwidth=False)
Frame_novo = LabelFrame(root,text='PLANILHA DE DADOS', bg="grey")
titulo = Label(root, text='ANALYSER', bg='black', font='Monaco 20 bold', fg='white', width=74, bd=5, relief='ridge', justify=CENTER)
Lb_arq = Label(root, bg='black', font='Monaco 7', fg='white', width=83, relief='ridge', justify=CENTER, text=f'{x}')
cmd = Button(root,text='EXECUTAR',command=ver_radio,font='Monaco 13 bold',bg='grey',width=18)
BtDeSelecao =  Button(text='CARREGAR ARQUIVO',command=subir_arquivo,font='Monaco 13 bold',bg='grey',width=18)

#botoes operadoras
Rb_claro1 = Radiobutton(root, text='Chamadas (c/ cad)',font='Monaco 13 bold',bg='black',fg='#FF0000', justify='left',variable=a, value=1, indicatoron=0, selectcolor='grey', width=20)
Rb_claro_sc = Radiobutton(root, text='Chamadas(s/ cad)',font='Monaco 13 bold',bg='black',fg='#FF0000', justify='left',variable=a, value=7, indicatoron=0, selectcolor='grey', width=20)
Rb_claro2 = Radiobutton(root, text='Conexões',font='Monaco 13 bold',bg='black',fg='#FF0000', justify='left',variable=a, value=6, indicatoron=0, selectcolor='grey', width=20)
Rb_tim1 = Radiobutton(root, text='Chamadas',font='Monaco 13 bold',bg='black',fg='#0000CD', justify='left',variable=a, value=2, indicatoron=0, selectcolor='grey', width=20)
Rb_tim2 = Radiobutton(root, text='Conexões',font='Monaco 13 bold',bg='black',fg='#0000CD', justify='left',variable=a, value=3, indicatoron=0, selectcolor='grey', width=20)
Rb_vivo1 = Radiobutton(root, text='Chamadas',font='Monaco 13 bold',bg='black',fg='#9400D3', justify='left',variable=a, value=4, indicatoron=0, selectcolor='grey', width=20)
Rb_vivo2 = Radiobutton(root, text='Conexões',font='Monaco 13 bold',bg='black',fg='#9400D3', justify='left',variable=a, value=5, indicatoron=0, selectcolor='grey', width=20)
   
#botoes tipo de analise
Lb_bilh = Radiobutton(root, text='Bilhetagem',font='Monaco 10 bold',bg='black',fg='white', justify='left',variable=b, value=1, indicatoron=0, selectcolor='grey', width=12)
Lb_imei = Radiobutton(root, text='IMEIs',font='Monaco 10 bold',bg='black',fg='white', justify='left',variable=b, value=2, indicatoron=0, selectcolor='grey', width=12)
Lb_loc = Radiobutton(root, text='Localizações',font='Monaco 10 bold',bg='black',fg='white', justify='left',variable=b, value=3, indicatoron=0, selectcolor='grey', width=12)
Lb_map = Radiobutton(root, text='Mapa',font='Monaco 10 bold',bg='black',fg='white', justify='left',variable=b, value=4, indicatoron=0, selectcolor='grey', width=12)

#input da planilha
tv1 = ttk.Treeview(Frame_novo)
tv1.place(relheight=1, relwidth=1)

rolagemx = Scrollbar(Frame_novo, orient='horizontal',command=tv1.xview)
rolagemy = Scrollbar(Frame_novo, orient='vertical',command=tv1.yview)

tv1.configure(xscrollcommand=rolagemx.set, yscrollcommand=rolagemy.set) 
             
#layout
titulo.place(y= 20, x=0)
Rb_claro1.place(y= 130, x=0)
Rb_claro2.place(y=190, x=0)
Rb_claro_sc.place(y=160, x=0)
Rb_tim1.place(y= 130, x=220)
Rb_tim2.place(y= 160, x=220)
Rb_vivo1.place(y= 130, x=440)
Rb_vivo2.place(y= 160, x=440)
cmd.place(y=255, x=460)
BtDeSelecao.place(y= 142, x=660)
imagem_draco.place(x=910, y=75)
Frame_novo.place(height=390,width=1150,y=300, x=25)
rolagemx.pack(side='bottom', fill='x')
rolagemy.pack(side='right', fill='y')
Lb_bilh.place(y=270, x=25)
Lb_imei.place(y=270, x=130)
Lb_loc.place(y=270, x=235)
Lb_map.place(y=270, x=340)
Lb_arq.place(x=25, y=253)
imagem_claro.place(x=76,y=74)
imagem_tim.place(x=290,y=66)
imagem_vivo.place(x=510,y=83)

root.mainloop()
