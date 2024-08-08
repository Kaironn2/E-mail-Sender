import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import uuid

def enviar_email(destinatario, assunto, corpo, remetente, senha, awb, pedido):
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg['Message-ID'] = f"<{uuid.uuid4()}@{remetente.split('@')[1]}>"

    msg.attach(MIMEText(corpo, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        text = msg.as_string()
        server.sendmail(remetente, destinatario, text)
        server.quit()
        print(f'[Pedido:#{pedido}] [AWB:{awb}] -> enviado para {destinatario}')
        return True
    except Exception as e:
        print(f'Falha ao enviar email para {destinatario}. Erro: {str(e)}')
        return False

def main():
    remetente_email = ''
    remetente_senha = ''

    arquivo_excel = "C:\\Users\\Kaironn2\\Desktop\\sender\\xlsx\\teste.xlsx"
    df = pd.read_excel(arquivo_excel)

    top_image_url = "https://i.imgur.com/4la6UoL.jpeg"
    jet1_image_url = "https://i.imgur.com/D1YJ4px.jpeg"
    jet2_image_url = "https://i.imgur.com/mZjCFv9.jpeg"
    azul1_image_url = "https://i.imgur.com/TjBIxFt.png"
    azul2_image_url = "https://i.imgur.com/rZImoWB.png"
    correios1_image_url = "https://i.imgur.com/U3YZTQo.png"
    correios2_image_url = "https://i.imgur.com/BhJUFdZ.jpeg"

    total_linhas = len(df)
    emails_nao_enviados = []
    contador = 1

    for index, row in df.iterrows():
        transportadora = row['transportadora']
        awb = row['awb']
        destinatario = row['destinatario']
        pedido = row['pedido']

        if transportadora == 'J&T':
            assunto = f'Instruções de Rastreio do Pedido #{pedido}123456'
            corpo = f"""
                <html>
                <head>
                    <style>
                        .center {{
                            display: block;
                            margin-left: auto;
                            margin-right: auto;
                        }}
                        .center-text {{
                            text-align: center;
                        }}
                        .large-font {{
                            font-size: 8em;
                            color: red;
                            font-weight: bold;
                        }}
                        .small-font {{
                            font-size: 1.5em;
                        }}
                        .no-bold {{
                            font-weight: normal;
                            font-size: 1.5em;
                        }}
                        .mini-image {{
                            width: 40%;
                            height: auto;
                        }}
                        .small-image {{
                            width: 50%;
                            height: auto;
                        }}
                        .medium-image {{
                            width: 75%;
                            height: auto;
                        }}
                        .bg-text {{
                            background-color: red;
                            padding: 2px 5px;
                            display: inline;
                            font-size: 1.5em;
                        }}
                        .responsive-image {{
                            width: 50%;
                        }}
                        .responsive-text {{
                            font-size: 1.5em;
                        }}
                        .responsive-large-font {{
                            font-size: 4em;
                            color: red;
                            font-weight: bold;
                        }}
                        @media (max-width: 600px) {{
                            .responsive-image {{
                                width: 85%;
                            }}
                            .responsive-text {{
                                font-size: 1.5em;
                            }}
                            .responsive-large-font {{
                                font-size: 2.4em;
                            }}
                        }}
                        @media (min-width: 601px) {{
                            .responsive-image {{
                                width: 50%;
                            }}
                            .responsive-text {{
                                font-size: 2em;
                            }}
                            .responsive-large-font {{
                                font-size: 3em;
                            }}
                        }}
                    </style>
                </head>
                <body>
                    <img src="{top_image_url}" alt="Top Image" class="center responsive-image">
                    <div class="center-text">
                        <h1 class="responsive-text">Seu pedido já está a caminho!<br>Rastreie seguindo os passos abaixo:</h1>
                        <br><br>
                        <h1 class="responsive-text"><strong>RASTREAMENTO DA CARTA DE PORTE:</strong><br>
                        <span class="responsive-large-font">{awb}</span></h1>
                        <br><br>
                        <h1 class="responsive-text small-font">Para rastrear, basta pesquisar por "J&T Express" e acessar o site da transportadora.</h1>
                        <h1 class="responsive-text no-bold">Por questões de segurança, não podemos inserir os links aqui pelo e-mail :(</h1>
                        <img src="{jet1_image_url}" alt="J&T Instructions 1" class="center responsive-image">
                        <h1 class="responsive-text no-bold">Após acessar o site, basta inserir o seu CPF (sem pontos e sem traços) e colocar o<br>código que te informamos acima e em seguida clique na lupa para rastrear.</h1>
                        <img src="{jet2_image_url}" alt="J&T Instructions 2" class="center responsive-image">
                        <br>
                        <h1 class="responsive-text small-font">Em caso de mais dúvidas, estamos sempre à disposição!</h1>
                        <br>
                        <h1 class="responsive-text bg-text">ATENÇÃO: Caso a embalagem esteja violada, recomendamos recusar o <br>recebimento e entrar em contato com nosso SAC, <br>através desse email ou do WhatsApp.</h1>
                        <br><br>
                        <h1 class="responsive-text small-font">Com amor,<br>Equipe Loja Sangue Latino</h1>
                    </div>
                </body>
                </html>
            """
        elif transportadora == 'AZUL':
            assunto = f'Instruções de Rastreio do Pedido #{pedido}123456'
            corpo = f"""
                <html>
                <head>
                    <style>
                        .center {{
                            display: block;
                            margin-left: auto;
                            margin-right: auto;
                        }}
                        .center-text {{
                            text-align: center;
                        }}
                        .large-font {{
                            font-size: 8em;
                            color: blue;
                            font-weight: bold;
                        }}
                        .small-font {{
                            font-size: 1.5em;
                        }}
                        .no-bold {{
                            font-weight: normal;
                            font-size: 1.5em;
                        }}
                        .mini-image {{
                            width: 40%;
                            height: auto;
                        }}
                        .small-image {{
                            width: 50%;
                            height: auto;
                        }}
                        .medium-image {{
                            width: 75%;
                            height: auto;
                        }}
                        .bg-text {{
                            background-color: red;
                            padding: 2px 5px;
                            display: inline;
                            font-size: 1.5em;
                        }}
                        .responsive-image {{
                            width: 50%;
                        }}
                        .responsive-text {{
                            font-size: 1.5em;
                        }}
                        .responsive-large-font {{
                            font-size: 4em;
                            color: blue;
                            font-weight: bold;
                        }}
                        @media (max-width: 600px) {{
                            .responsive-image {{
                                width: 85%;
                            }}
                            .responsive-text {{
                                font-size: 1.5em;
                            }}
                            .responsive-large-font {{
                                font-size: 2.4em;
                            }}
                        }}
                        @media (min-width: 601px) {{
                            .responsive-image {{
                                width: 50%;
                            }}
                            .responsive-text {{
                                font-size: 2em;
                            }}
                            .responsive-large-font {{
                                font-size: 3em;
                            }}
                        }}
                    </style>
                </head>
                <body>
                    <img src="{top_image_url}" alt="Top Image" class="center responsive-image">
                    <div class="center-text">
                        <h1 class="responsive-text">Seu pedido já está a caminho!<br>Rastreie seguindo os passos abaixo:</h1>
                        <br><br>
                        <h1 class="responsive-text"><strong>CÓDIGO AZUL:</strong><br>
                        <span class="responsive-large-font">{awb}</span></h1>
                        <br><br>
                        <h1 class="responsive-text small-font">Para rastrear, basta pesquisar por "Rastreio Azul Cargo" e acessar o site da transportadora.</h1>
                        <h1 class="responsive-text no-bold">Por questões de segurança, não podemos inserir os links aqui pelo e-mail :(</h1>
                        <img src="{azul1_image_url}" alt="Azul Instructions 1" class="center responsive-image">
                        <h1 class="responsive-text no-bold">Após acessar o site, basta inserir o código informado acima e clicar em adicionar.<br>Depois, marque "Não sou um robô" e em seguida clique em Rastrear.</h1>
                        <img src="{azul2_image_url}" alt="Azul Instructions 2" class="center responsive-image">
                        <br>
                        <h1 class="responsive-text small-font">Em caso de mais dúvidas, estamos sempre à disposição!</h1>
                        <br>
                        <h1 class="responsive-text bg-text">ATENÇÃO: Caso a embalagem esteja violada, recomendamos recusar o <br>recebimento e entrar em contato com nosso SAC, <br>através desse email ou do WhatsApp.</h1>
                        <br><br>
                        <h1 class="responsive-text no-bold">Com amor,<br>Equipe Loja Sangue Latino</h1>
                    </div>
                </body>
                </html>
            """
        elif transportadora == 'CORREIOS':
            assunto = f'Instruções de Rastreio do Pedido #{pedido}123456'
            corpo = f"""
                <html>
                <head>
                    <style>
                        .center {{
                            display: block;
                            margin-left: auto;
                            margin-right: auto;
                        }}
                        .center-text {{
                            text-align: center;
                        }}
                        .large-font {{
                            font-size: 8em;
                            color: #DAA520;
                            font-weight: bold;
                        }}
                        .small-font {{
                            font-size: 1.5em;
                        }}
                        .no-bold {{
                            font-weight: normal;
                            font-size: 1.5em;
                        }}
                        .mini-image {{
                            width: 40%;
                            height: auto;
                        }}
                        .small-image {{
                            width: 50%;
                            height: auto;
                        }}
                        .medium-image {{
                            width: 75%;
                            height: auto;
                        }}
                        .bg-text {{
                            background-color: red;
                            padding: 2px 5px;
                            display: inline;
                            font-size: 1.5em;
                        }}
                        .responsive-image {{
                            width: 50%;
                        }}
                        .responsive-text {{
                            font-size: 1.5em;
                        }}
                        .responsive-large-font {{
                            font-size: 4em;
                            color: #DAA520;
                            font-weight: bold;
                        }}
                        @media (max-width: 600px) {{
                            .responsive-image {{
                                width: 85%;
                            }}
                            .responsive-text {{
                                font-size: 1.5em;
                            }}
                            .responsive-large-font {{
                                font-size: 2.4em;
                            }}
                        }}
                        @media (min-width: 601px) {{
                            .responsive-image {{
                                width: 50%;
                            }}
                            .responsive-text {{
                                font-size: 2em;
                            }}
                            .responsive-large-font {{
                                font-size: 3em;
                            }}
                        }}
                    </style>
                </head>
                <body>
                    <img src="{top_image_url}" alt="Top Image" class="center responsive-image">
                    <div class="center-text">
                        <h1 class="responsive-text">Seu pedido já está a caminho!<br>Rastreie seguindo os passos abaixo:</h1>
                        <br><br>
                        <h1 class="responsive-text"><strong>CÓDIGO DOS CORREIOS:</strong><br>
                        <span class="responsive-large-font">{awb}</span></h1>
                        <br><br>
                        <h1 class="responsive-text small-font">Para rastrear, basta pesquisar por "Rastreamento Correios" e acessar o site.</h1>
                        <h1 class="responsive-text no-bold">Por questões de segurança, não podemos inserir os links aqui pelo e-mail :(</h1>
                        <img src="{correios1_image_url}" alt="Correios Instructions 1" class="center responsive-image">
                        <h1 class="responsive-text no-bold">Após acessar o site, basta inserir o código de rastreio acima enviado.</h1>
                        <img src="{correios2_image_url}" alt="Correios Instructions 2" class="center responsive-image">
                        <br>
                        <h1 class="responsive-text small-font">Em caso de mais dúvidas, estamos sempre à disposição!</h1>
                        <br>
                        <h1 class="responsive-text bg-text">ATENÇÃO: Caso a embalagem esteja violada, recomendamos recusar o <br>recebimento e entrar em contato com nosso SAC, <br>através desse email ou do WhatsApp.</h1>
                        <br><br>
                        <h1 class="responsive-text no-bold">Com amor,<br>Equipe Loja Sangue Latino</h1>
                    </div>
                </body>
                </html>
            """

        else:
            print(f'Linha inválida na linha {index + 2}')  # Adiciona 2 para corresponder ao índice do Excel (1-based) e ao cabeçalho
            continue

        if not enviar_email(destinatario, assunto, corpo, remetente_email, remetente_senha, awb=awb, pedido=pedido):
            emails_nao_enviados.append(destinatario)

        print(f'{contador}/{total_linhas}\n')
        contador += 1
        time.sleep(2)  # TESTAR DEPOIS COM 1 SEC SLEEP APENAS

    print(f'\n\nE-mails não enviados: {emails_nao_enviados}')

if __name__ == "__main__":
    main()
