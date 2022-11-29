from datetime import datetime
from random import choice
import smtplib
import email.message
from time import sleep
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

DATA = datetime.now().strftime('%d/%m/%Y')
HORA = datetime.now().strftime('%H:%M - Horário de Brasília')
EMAIL = 'joaomarcosndasilva.lna@gmail.com'
SENHA = 'nctntrhwzqkvxnyh'
destinatario = ['joaomarcosndasilva@gmail.com', 'jmndbruto@gmail.com', 'joaomarcosndasilva.economista@gmail.com']

#df = st.download_button('Selecione o relatório do Jorge!')
df = pd.read_excel('Relatorio2.xlsx')
#df = st.file_uploader('Selecione o Arquivo para o gerenciador', type=['xls', 'xlsx'])
#arquivo_imagem = st.file_uploader('Envie Sua Imagem e role para baixo para ver o resultado', type=['jpg', 'png', 'jpeg'] )


def leia_relatorio():
    import pandas as pd
    import numpy as np
    from datetime import datetime
    df = pd.read_excel('Relatório.xlsx')
    df['Prazo_editado'] = pd.to_datetime(df['Prazo'])
    df['Atual'] = datetime.now()
    df2 = df['Atual'] - df['Prazo']
    df['Diferenca'] = df2


def cobranca_por_regiao():
    leia_relatorio()
    regional_lista = ['TODAS','NRT', 'NRO', 'LES', 'OES', 'CSL']
    funcionario_lista = ['telêmaco', 'flávio', 'leandro', 'joão', 'gabriela']
    regional_escolhida = choice(regional_lista)
    funcionario_escolhido = choice(funcionario_lista)
    funcionario = funcionario_escolhido.title()
    regional = regional_escolhida.upper()
    return f'{funcionario} - {regional}'
    '''
    if regional == 'TODAS':
        print(df)
        print('Aqui são todas as regiões')
    elif regional == 'NRT':
        print(df[df['Regional'] == 'NRT'])
        print('Todas da NRT')
    elif regional == 'NRO':
        print([df['Regional'] == 'NRO'])
        print('todas da nro')
    elif regional == 'LES':
        print(df[df[df['Regional'] == 'LES']])
    elif regional == 'OES':
        print(df[df['Regional'] == 'OES'])
    elif regional == 'CSL':
        print(df[df['Regional'] == 'CSL'])
    '''

def dados_regionais():
    return print('Dados regionais')
def enviar_email():
    corpo_email = f"""
    <p>{cobranca_por_regiao()}</p> 
    """

    msg = email.message.Message()
    msg['Subject'] = f'Email enviado na data: {DATA}, às {HORA}'
    msg['From'] = EMAIL
    msg['To'] = email_para_envio
    password = SENHA
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], msg['To'], msg.as_string().encode('utf-8'))
    st.write(corpo_email)

st.title('Gerenciador automático de APDS')
regiao = st.selectbox('Qual Região você quer analisar?', ('TODAS','NRT', 'NRO', 'LES', 'OES', 'CSL'))
if regiao == 'TODAS':
    st.dataframe(df)
elif regiao == 'NRT':
    st.dataframe(df[df['Regional']=='NRT'])
    st.success('Funcionário Responsável: João Marcos')
elif regiao == 'NRO':
    st.dataframe(df[df['Regional']=='NRO'])
    st.success('Funcionário Responsável: Leandro')
elif regiao == 'OES':
    st.dataframe(df[df['Regional']=='OES'])
    st.success('Funcionário Responsável: Responsável Por Cascavél')
elif regiao == 'CSL':
    st.dataframe(df[df['Regional'] == 'CSL'])
    st.success('Funcionário Responsável: Flávio Augusto')
elif regiao == 'LES':
    st.dataframe(df[df['Regional'] == 'LES'])
    st.success('Funcionário Responsável: Telêmaco')

st.sidebar.subheader('Digite os dados e preferências do gerenciador dos apds')
email_para_envio = st.sidebar.selectbox('Selecione o e-mail do gerenciador', destinatario, help='Esse email vai receber o feedback, isto é, todo o retorno da atividade de cobrança e alerta desse gerenciador')
st.sidebar.success(f'Retorno será enviado para: {email_para_envio}')
cb_altera_emails = st.sidebar.checkbox('Alterar os e-mails dos responsávies')
if cb_altera_emails:
    selectbox_emails = st.sidebar.selectbox('Selecione a Região para alterar os responsáveis', ('TODAS','NRT', 'NRO', 'LES', 'OES', 'CSL'), help='Os responsáveis já estão definidos, porém com férias, pode-se atribuir outra pessoa para cuidar dos apds')
    if selectbox_emails == 'NRT':
        st.sidebar.info('Email Cadstrado do João Marcos: joao.marcos@copel.com')
        cb_novo_email = st.sidebar.checkbox('Quer Selecionar outro responsável?')
        if cb_novo_email:
            st.sidebar.selectbox('Selecione o novo responsável', ('Flávio', 'Telêmaco', 'Gabriele', 'Andressa', 'Carlos Àvila', 'Foganholi'))


enviar_email_radio = st.sidebar.radio('Enviar Email', ('Cobrador Manual', 'Cobrador Automático'))


if enviar_email_radio == 'Cobrador Automático':
    slider_vezes = st.sidebar.slider('Quantos e-mails enviar? (Clique na interrogação)', 1, 20, 5, help='Na quantidade 5, ele vai mandar 1 email quando faltar 3 dias, 2 quando faltar 2 dias(um 08:00 e outro 14:30) e no dia do vencimento será enviado um no primeiro horário e otro uma hora antes de vencer')
    emails_auto = st.sidebar.checkbox('Iniciar Disparos de e-mails controlados')
    msg_text = st.sidebar.checkbox('Iniciar cobranças com mensagens de texto')
    whats_app = st.sidebar.checkbox('Iniciar cobranças com mensagem de WhatsApp')
    if emails_auto:
        minutos = st.sidebar.number_input('Selecione o intervalo de quantos minutos será o envio', format='%d', step=1)
        cont = 0
        if minutos:
            btn_iniciar = st.sidebar.button('INICIAR')
            st.info('Iniciando o processo de cobrança...')
            if btn_iniciar:
                while cont < slider_vezes:
                    with st.spinner('Wait for it...'):
                        enviar_email()
                        sleep(minutos*60)
                        cont += 1
                    st.success(f'Fim dos Envios de e-mails. Total de {cont} emails enviados!')
