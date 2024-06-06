#Bibliotecas necessárias
import pandas as pd
import numpy as np
import datetime
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import streamlit as st


st.set_page_config(page_title='Dashboard de Atendimento')
Arquivo = st.file_uploader(label = 'Faça o Upload do Relatório de Atendimentos Maxbot. O relatório pode ser obtido acessando sua conta maxbot, no menu lateral clique em "Relatório", em seguida "Atendimento". Selecione o período de referência e gere o relatório')
st.title('Dashboard Análise de Atendimentos Maxbot 🤖')

#Reading exported report
if Arquivo:
    df = pd.read_excel(Arquivo)


    #Data Cleaning
    df = df.drop(index = 0) #Deleting the first row with NaNs
    df.columns = df.iloc[0] #Updating the column names
    df = df.drop(index = 1) #Deleting the colnames remaining in the DF
    df = df.reset_index(drop = True) #Resetting index
    #df.head()

    #Adjusting data types
    df['PROTOCOLO'] = pd.to_numeric(df['PROTOCOLO'], errors = 'coerce')
    df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y')
    df['ABERTURA'] = pd.to_datetime(df['ABERTURA'], format='%d/%m/%Y %H:%M:%S')
    df['1ª MSG ATENDENTE'] = pd.to_datetime(df['1ª MSG ATENDENTE'], format='%d/%m/%Y %H:%M:%S')
    df['ORIGEM'] = df['ORIGEM'].astype('category')
    df['CANAL'] = df['CANAL'].astype('category')
    df['SETOR'] = df['SETOR'].astype('category')
    df['ATENDENTE'] = df['ATENDENTE'].astype('category')
    df['CONTATO'] = df['CONTATO'].astype(str)
    #df['INFO. CONTATO'] = df['INFO. CONTATO'].astype(str)
    df['SITUAÇÃO'] = df['SITUAÇÃO'].astype('category')


    #
    #
    #
    #
    #Aqui tem o salvamento do excel, está com comentário para não ficar salvando a todo momento.
    #with pd.ExcelWriter('Protocolos2024.xlsx', engine = 'openpyxl', mode = 'a', if_sheet_exists = 'replace') as writer:
    #    df.to_excel(writer, sheet_name='DadosTratados')

    #
    #
    #
    #Esse dado quando sai do pandas e vai para o excel como timedelta dá pau. 
    #Por isso é preciso alterar essas colunas após o salvamento no excel
    df['TEMPO DE FILA'] = pd.to_timedelta(df['TEMPO DE FILA'], errors = 'coerce')
    df['DURAÇÃO H:M:S'] = pd.to_timedelta(df['DURAÇÃO H:M:S'], errors = 'coerce')





    st.subheader('Métricas de atendimento')

    #
    #
    #
    #Metrica de protocolos em atendimento no momento
    df_Em_Atendimento = df[(df['SITUAÇÃO'] == 'Em Atendimento')]
    count_em_atendimento = df_Em_Atendimento.shape[0]
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Protocolos Em Atendimento", count_em_atendimento, 'Abertos no momento')

    #
    #
    #
    #Aqui se inicia a quantidade de protocolos abertos hoje
    df_hoje = df[df['DATA'] == datetime.today().strftime('%Y-%m-%d')]
    protocolos_hoje = df_hoje.shape[0]
    col2.metric("Protocolos Iniciados hoje", protocolos_hoje, 'Iniciados hoje', delta_color='off')

    #
    #
    #
    #Aqui se inicia a média de tempo de fila hoje
    media = df_hoje['TEMPO DE FILA'].mean()
    media_format = pd.Timedelta(media)
    hours = media_format.components.hours
    minutes = media_format.components.minutes
    seconds = media_format.components.seconds
    formatted_time = f'{hours}h{minutes}m{seconds}s'
    col3.metric("Tempo médio de fila hoje", formatted_time, 'Iniciados hoje', delta_color='off')

    #
    #
    #
    #Aqui inicia-se a contagem de protocolos aguardando
    aguardando = df['SITUAÇÃO'][df['SITUAÇÃO'] == 'Aguardando Atendimento'].count()
    aguardando
    col4.metric("Protocolos na fila", aguardando, 'Aguardando Atendimento', delta_color="inverse")



    #
    #
    #
    #Seção para protocoloc em atendimento por setor.
    st.subheader('Protocolos em atendimento por setor')

    df_by_setor = df_Em_Atendimento.groupby(['SETOR'])['SETOR'].count()
    #df_by_setor.index_name = 'Protocolos Abertos' ISSO AQUI PODE DELETAR QUE NÃO FUNCIONOU

    df_setores = df_by_setor.to_frame()
    df_setores = df_setores.rename(columns = {'SETOR': 'EM ATEDIMENTO'})
    df_setores = df_setores.sort_values('EM ATEDIMENTO', ascending = False)
    df_setores2 = df_setores.reset_index()
    col1, col2 = st.columns([1,1])
    col1.dataframe(df_setores2, hide_index=True)
    col2.bar_chart(df_setores)



    #
    #
    #
    #Aqui se inicia os protocolos em atendimento por atendete.
    st.subheader('Protocolos em atendimento por atendente')
    df_atendentes = df_Em_Atendimento.groupby(['ATENDENTE'], observed=False)['ATENDENTE'].count()
    df_atendentes = df_atendentes.to_frame()
    df_atendentes = df_atendentes.rename(columns = {'ATENDENTE': 'EM ATEDIMENTO'})
    df_atendentes = df_atendentes.sort_values('EM ATEDIMENTO', ascending = False)
    df_atendentes2 = df_atendentes.reset_index()
    col1, col2 = st.columns([1,1])
    col1.dataframe(df_atendentes2, hide_index=True, column_config={
        'ATENDENTE' : st.column_config.Column(width="small")
    })#.style.set_properties(**{'width': '100px'}, subset=['ATENDENTE'])
                                #.set_properties(**{'width': '200px'}, subset=['EM ATENDIMENTO']))
    col2.bar_chart(df_atendentes)



    # 
    #
    #
    # Aqui se inicia a parte de protocolos abertos por dias com o elemento slider
    st.header('Protocolos abertos por dia')
    df_data = df.groupby(['DATA'], observed=False)['PROTOCOLO'].count()
    df_data = df_data.to_frame()
    df_data = df_data.rename(columns = {'PROTOCOLO': 'PROTOCOLOS ABERTOS'})

    df_data.reset_index(inplace=True)
    df_data['DATA'] = pd.to_datetime(df_data['DATA'])

    # Extract the date part (this returns a Series of datetime.date objects)
    df_data['DATE_ONLY'] = df_data['DATA'].dt.date

    # Set 'DATE_ONLY' column as the index
    df_data.set_index('DATE_ONLY', inplace=True)

    # Optionally, drop the original 'DATA' column
    df_data.drop('DATA', axis=1, inplace=True)

    # Print the DataFrame
    print(df_data)
    start_date, end_date = st.select_slider(
        "Selecione o range do período",
        options=df_data.index,
        value=(df_data.index[0], df_data.index[-1]))#,
        #format='MM/DD/YYYY')
    #df_data_sliced = df_data.loc[start_date, end_date]
    st.bar_chart(df_data)
    #st.bar_chart(df_data_sliced)
    #st.bar_chart(df_data_sliced)

    # Create a datetime slider with custom format and options
    start_date = datetime(2020, 1, 1)
    end_date = start_date + timedelta(weeks=1)
    
    # selected_date = st.slider(
    #     "Select a date range",
    #     min_value=start_date,
    #     max_value=end_date,
    #     value=(start_date, end_date),
    #     step=timedelta(days=1),
    #     format="MM/DD/YYYY",
    #     options=["Day", "Week", "Month", "Year"]
    # )

