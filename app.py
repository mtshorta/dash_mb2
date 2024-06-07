#Bibliotecas necessárias
import pandas as pd
import numpy as np
import datetime
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import streamlit as st


st.set_page_config(page_title='Dashboard de Atendimento')

#
#
#
#Upload do arquivo
Arquivo = st.file_uploader(label = 'Faça o Upload do Relatório de Atendimentos Maxbot. O relatório pode ser obtido acessando sua conta maxbot, no menu lateral clique em "Relatório", em seguida "Atendimento". Selecione o período de referência e gere o relatório')
#Arquivo = 'Protocolos2024-06-07.xlsx'

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

        #
    #
    #
    #
    #
    #
    #
    #
    #
    #AQUI MEU AMIGO, um malabarismo imenso para conseguir fazer o gráfico de aberturas e fechamentos
    #Pego o a menor e maior data do df_fechamento, que nada mais é que o df normal mais uma coluna de data de fechamento 
    # E outra coluna de data e hora de fechamento dos protocolos
    #Criando a data de fechamento
    st.subheader('Número de protocolos abertos e fechados no período')
    df_fechamento = df
    #
    #Calculando a data do fechamento
    def calculate_data_fechamento(row):
        if row['SITUAÇÃO'] == 'Encerrado':
            return row['ABERTURA'] + row['DURAÇÃO H:M:S']
        else:
            return pd.NaT

    # Apply the function to create 'DATA FECHAMENTO' column
    df_fechamento['DATA E HORA FECHAMENTO'] = df.apply(calculate_data_fechamento, axis=1)
    df_fechamento['DATA FECHAMENTO']= df['DATA E HORA FECHAMENTO'].apply(lambda x: x.date() if pd.notnull(x) else x)

    # Display the DataFrame
    #df_fechamento


    min = df_fechamento['DATA'].min().date()
    max = df_fechamento['DATA'].max().date()
    print(min, max)

    #Aqui eu crio um dataframe com todas as datas entre o mínimo e o máximo que é o df_openings
    date_range = pd.date_range(start=min, end=max)


    #Aqui, eu crio um df com o daterange acima, e nomeio as colunas como Date
    df_openings = pd.DataFrame(date_range, columns=['Date'])
    print('este é o df openings')
    print(df_openings)

    #Aqui eu faço o date_counts, que é um DF que conta as datas de abertura, o resultado final é um df com Date e Count dessa respectiva data
    date_counts = df_fechamento['DATA'].value_counts().reset_index()
    date_counts.columns = ['Date', 'Count']
    print('Este é o date counts')
    print(date_counts)


    print('Este é o df_openings ANTES do merge')
    print(df_openings)
    # Merge the counts into df_openings. Eu preciso fazer isso pois o date_counts pode ter datas vazias
    df_openings = df_openings.merge(date_counts, on='Date', how='left')
    # Fill NaN values with 0
    df_openings['Count'] = df_openings['Count'].fillna(0).astype(int)
    print('Este é o df_openings DEPOIS do merge')
    print(df_openings)


    # Rename the 'Count' column to 'abertura'
    df_openings = df_openings.rename(columns={'Count': 'aberturas'})

    # Display the DataFrame
    print('Esse é o df openings depois do rename aberturas: ')
    print(df_openings)

    #
    #
    #Segunda parte do code para encontrar os fechamentos por dia

    #Aqui eu faço o date_counts, que é um DF que conta as datas de abertura, o resultado final é um df com Date e Count dessa respectiva data
    date_counts2 = df_fechamento['DATA FECHAMENTO'].value_counts().reset_index()
    #print(date_counts2)
    date_counts2.columns = ['Date', 'Count']
    #print(date_counts2)

    #converting to date
    date_counts2['Date'] = pd.to_datetime(date_counts2['Date'])
    #print(df_openings)
    # Merge the counts into df_openings. Eu preciso fazer isso pois o date_counts pode ter datas vazias
    df_openings = df_openings.merge(date_counts2, on='Date', how='left')

    df_openings = df_openings.rename(columns={'Count': 'fechamentos'})
    df_openings['fechamentos'] = df_openings['fechamentos'].fillna(0).astype(int)
    df_openings['DateOnly'] = df_openings['Date'].dt.date
    print(df_openings)

    #
    #
    #Printando o gráfico
    col1, col2 = st.columns([8,2])
    col1.line_chart(
    df_openings,x='Date', y=["aberturas", "fechamentos"], color=["#0033CC", "#CC3333"])
    cont_aberturas=df_openings['aberturas'].sum()
    cont_fechamentos=df_openings['fechamentos'].sum()
    col2.metric('Abertos no período', cont_aberturas, "Protocolos")
    col2.metric('Fechados no período', cont_fechamentos, "-Protocolos")


    #Encerramento do gráfico de linhas
    #
    #
    #
    #
    #
    #
    #
    #



    st.subheader('Métricas de atendimento')

    #
    #
    #
    #Metrica de protocolos em atendimento no momento
    df_Em_Atendimento = df[(df['SITUAÇÃO'] == 'Em Atendimento')]
    count_em_atendimento = df_Em_Atendimento.shape[0]
    col1, col2, col3 = st.columns(3)
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
    # media = df_hoje['TEMPO DE FILA'].mean()
    # media_format = pd.Timedelta(media)
    # hours = media_format.components.hours
    # minutes = media_format.components.minutes
    # seconds = media_format.components.seconds
    # formatted_time = f'{hours}h{minutes}m{seconds}s'
    def check_fila_hoje():
        if protocolos_hoje == True:
            media = df_hoje['TEMPO DE FILA'].mean()
            media_format = pd.Timedelta(media)
            hours = media_format.components.hours
            minutes = media_format.components.minutes
            seconds = media_format.components.seconds
            formatted_time = f'{hours}h{minutes}m{seconds}s'
            return formatted_time
        else:
            return 0
        
    media_tempo_fila_hoje=df_hoje['TEMPO DE FILA'].mean()
    media_format = pd.Timedelta(media_tempo_fila_hoje)
    hours = media_format.components.hours
    minutes = media_format.components.minutes
    seconds = media_format.components.seconds
    formatted_tempo_fila_hoje = f'{hours}h{minutes}m{seconds}s'

    col3.metric("Tempo médio de fila hoje", formatted_tempo_fila_hoje, 'Iniciados hoje', delta_color='off')

    #
    #
    #
    #Aqui inicia-se a contagem de protocolos aguardando
    aguardando = df['SITUAÇÃO'][df['SITUAÇÃO'] == 'Aguardando Atendimento'].count()
    aguardando
    col1.metric("Protocolos na fila", aguardando, 'Aguardando Atendimento', delta_color="inverse")


    #
    #
    #
    #Aqui inicia-se a apuração do tempo médio de duração de todos os protocolos
    df_fechados = df_fechamento[df_fechamento['SITUAÇÃO'] == 'Encerrado']
    tempo_medio = df_fechados['DURAÇÃO H:M:S'].mean()
    tempo_medio_format = pd.Timedelta(tempo_medio)
    hours = tempo_medio_format.components.hours
    minutes = tempo_medio_format.components.minutes
    seconds = tempo_medio_format.components.seconds
    tempo_medio_formatado = f'{hours}h{minutes}m{seconds}s'
    col2.metric('Tempo médio resolução', tempo_medio_formatado, 'Prot. encerrados' ,delta_color="inverse")

    #
    #
    #
    #Aqui inicia-se a apuração do tempo médio de duração de todos os protocolos abertos
    now = datetime.now()
    df_Em_Atendimento['DURAÇÃO H:M:S'] = (now - df_Em_Atendimento['ABERTURA'])
    tempo_medio_format = df_Em_Atendimento['DURAÇÃO H:M:S'].mean()
    hours = tempo_medio_format.components.hours
    minutes = tempo_medio_format.components.minutes
    seconds = tempo_medio_format.components.seconds
    tempo_medio_formatado = f'{hours}h{minutes}m{seconds}s'
    col3.metric('Duração Atendimento', tempo_medio_formatado, 'Prot. em atendimento' ,delta_color="inverse")



    #
    #
    #
    #Seção para protocoloc em atendimento por setor.
    st.subheader('Protocolos em atendimento por setor')

    df_by_setor = df_Em_Atendimento.groupby(['SETOR'])['SETOR'].count()
    #df_by_setor.index_name = 'Protocolos Abertos' ISSO AQUI PODE DELETAR QUE NÃO FUNCIONOU

    df_setores = df_by_setor.to_frame()
    df_setores = df_setores.rename(columns = {'SETOR': 'EM ATENDIMENTO'})
    df_setores = df_setores.sort_values('EM ATENDIMENTO', ascending = False)
    df_setores2 = df_setores.reset_index()
    col1, col2 = st.columns([1,1])
    col1.dataframe(df_setores2, hide_index=True)
    col2.bar_chart(df_setores)



    #
    #
    #
    #Aquui se inicia os protocolos em atendimento por atendete.
    st.subheader('Protocolos em atendimento por atendente')
    df_atendentes = df_Em_Atendimento.groupby(['ATENDENTE'], observed=False)['ATENDENTE'].count()
    df_atendentes = df_atendentes.to_frame()
    df_atendentes = df_atendentes.rename(columns = {'ATENDENTE': 'EM ATENDIMENTO'})
    df_atendentes = df_atendentes.sort_values('EM ATENDIMENTO', ascending = False)
    #df_atendentes2 = df_atendentes.reset_index()
    df_atendentes = df_atendentes[df_atendentes['EM ATENDIMENTO'] > 0]
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
    #
    #Aqui vou inserir o df dos atendimentos mais demorados
    df_em_at = df_Em_Atendimento[['PROTOCOLO', 'ATENDENTE', 'CANAL', 'CONTATO', 'DURAÇÃO H:M:S']]

    st.header('Atendimentos mais antigos ainda não encerrados')

    st.dataframe(df_em_at.head(32), hide_index=True)






    # 
    #
    #
    # # Aqui se inicia a curva de atendimento
    list_setor = list(df['SETOR'].value_counts().index.categories)
    

    #Menu de opções
    options = st.multiselect(
    "Selecione os setores para comparar com os valores gerais",
    list_setor,
    list_setor)


    col1, col2 = st.columns(2)
    df_analitic = df
    df_analitic['ABERTURA'] = pd.to_datetime(df_analitic['ABERTURA'])
    df_analitic['HOUR'] = df_analitic['ABERTURA'].dt.hour
    hour_counts = df_analitic['HOUR'].value_counts().sort_index()
    col1.write('Curva de atendimento geral')

    # Create a DataFrame with all hours of the day (0 to 23)
    all_hours = pd.DataFrame({'HOUR': range(24)})
    # Merge the DataFrames to ensure all hours are included
    hour_counts_merged = pd.merge(all_hours, hour_counts, on='HOUR', how='left').fillna(0)
    col1.bar_chart(hour_counts_merged.set_index('HOUR')) 

  
    

    df_analitic2 = df[df['SETOR'].isin(options)]
    df_analitic2['ABERTURA'] = pd.to_datetime(df_analitic['ABERTURA'])
    df_analitic2['HOUR'] = df_analitic2['ABERTURA'].dt.hour
    hour_counts2 = df_analitic2['HOUR'].value_counts().sort_index()
    col2.write('Curva de atendimento por setor')

    # Merge the DataFrames to ensure all hours are included
    hour_counts_merged2 = pd.merge(all_hours, hour_counts2, on='HOUR', how='left').fillna(0)
    col2.bar_chart(hour_counts_merged2.set_index('HOUR')) 
    #Ploting results
    # 
