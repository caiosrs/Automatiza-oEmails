from flask import Flask
import pyodbc
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
from smtplib import SMTPException
from datetime import datetime, timedelta

import sys, os

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Exemplo de uso:
img = get_resource_path('Diario de Bordo LOGO.jpeg')
senha = get_resource_path(r'C:\Reviant\Documentos\Codes\Python\Automatização de Emails\senha temporaria.txt')

app = Flask(__name__)

emails_enviados = set()

def enviar_email():
    hoje = datetime.now()

    if hoje.hour >= 17:
        return

    # Conexão para verificar os feriados
    conn_old = pyodbc.connect('Driver={SQL Server};Server=sql.informatecservicos.com.br;Database=informatec01;UID=informatec01;PWD=AkY27Vf6xRr_4Nf;')
    cursor_old = conn_old.cursor()

    # Verificar se o dia anterior é feriado
    dia_anterior = hoje - timedelta(days=1)

    cursor_old.execute("SELECT COUNT(*) FROM Feriados WHERE Data = ?", dia_anterior.strftime('%Y-%m-%d'))
    feriado = cursor_old.fetchone()[0] > 0

    if hoje.weekday() == 0:  # Segunda-feira
        dia_anterior = hoje - timedelta(days=3)  # Voltar para Sexta-feira
        while True:
            cursor_old.execute("SELECT COUNT(*) FROM Feriados WHERE Data = ?", dia_anterior.strftime('%Y-%m-%d'))
            if cursor_old.fetchone()[0] == 0:  # Se não for feriado
                break
            dia_anterior -= timedelta(days=1)  # Voltar mais um dia
    else:
        while dia_anterior.weekday() >= 5:  # Sábado ou domingo
            dia_anterior -= timedelta(days=1)  # Voltar um dia
        while True:
            cursor_old.execute("SELECT COUNT(*) FROM Feriados WHERE Data = ?", dia_anterior.strftime('%Y-%m-%d'))
            if cursor_old.fetchone()[0] == 0:  # Se não for feriado
                break
            dia_anterior -= timedelta(days=1)  # Voltar mais um dia

    conn_old.close()

    data_titulo = dia_anterior.strftime('%d/%m')

    conn_new = pyodbc.connect('Driver={SQL Server};Server=sql.informatecservicos.com.br;Database=informatec01;UID=informatec01;PWD=AkY27Vf6xRr_4Nf;')
    cursor_new = conn_new.cursor()

    query = """
    SELECT
        A.CODUSUARIO, A.NOME, E.EMAIL,
        SUM(CASE WHEN B.DATA = ? THEN 1 ELSE 0 END) AS INFO
    FROM
        TBDB001 A
    LEFT JOIN
        TBDB010 B ON
        A.CODUSUARIO = B.CODUSUARIO AND
        B.DATA = ?
    LEFT JOIN
        dpcontrole.DBO.TBUSER001 E ON A.CPF = E.CPF COLLATE SQL_Latin1_General_CP1_CI_AS
    WHERE
        A.AUTOMATICO = 0 AND A.ATIVO = 1 AND E.STATUS = 'A'
    GROUP BY
        A.CODUSUARIO, A.NOME, E.EMAIL
    HAVING
        SUM(CASE WHEN B.DATA = ? THEN 1 ELSE 0 END) = 0
    """

    dia_anterior_str = dia_anterior.strftime('%Y-%m-%d')
    cursor_new.execute(query, dia_anterior_str, dia_anterior_str, dia_anterior_str)
    clientes = cursor_new.fetchall()

    conn_new.close()

    with open(r'senha temporaria.txt') as f:
        senha = f.read().strip()

    try:
        for cliente in clientes:
            if len(cliente) == 4:
                _, nome, email, _ = cliente
            else:
                continue

            if email in emails_enviados:
                continue

            emails_enviados.add(email)

            msg = MIMEMultipart()

            msg['Subject'] = '[Diário de Bordo] Preenchimento Pendente!'
            msg['From'] = "noreply@informatecservicos.com.br"
            msg['To'] = email

            html_content = f"""
<html>
    <head>
        <style>
        .container {{
            font-family: Arial, sans-serif;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f4f4f4;
            border-radius: 5px;
        }}
        .message {{
            font-size: 16px;
            line-height: 1.6;
        }}
        .button {{
            display: inline-block;
            padding: 10px 20px;
            background-color: #007bff;
            color: #fff;
            text-decoration: none;
            border-radius: 5px;
        }}
        .button:hover {{
            background-color: #0056b3;
        }}
        .image {{
            margin-bottom: 20px;
            text-align: left;
        }}
        .image img {{
            max-width: 50%;
            height: auto;
        }}
        .signature {{
            margin-top: 20px;
            font-size: 14px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="image">
            <img src="cid:diario_bordo_logo" alt="Diário de Bordo">
        </div>
        <p>Olá, {nome}.</p>
        <p class="message"><strong>Verificamos que existem pendências em seu Diário de Bordo.</p>
        <br/>Por favor, realize o preenchimento o mais rápido possível.</strong></p>
        <p class="message">Clique no botão abaixo para acessar o Diário de Bordo:</p>
        <a href="https://dpcontrole.informatecservicos.com.br/" class="button"><strong>Acessar o Diário de Bordo</strong></a>
        <p class="message">Obrigado pela sua atenção!</p>
        <p class="signature">Atenciosamente,<br>Informatec Serviços</p>
        <br/>
    </div>
</body>
</html>
            """

            with open(r'C:\Reviant\Documentos\Codes\Python\Automatização de Emails\Diario de Bordo LOGO.jpeg', 'rb') as img_file:
                img = MIMEImage(img_file.read())
                img.add_header('Content-Disposition', 'inline', filename='Diario de Bordo LOGO.jpeg')
                img.add_header('Content-ID', '<diario_bordo_logo>')
                msg.attach(img)

            msg.attach(MIMEText(html_content, 'html'))

            with smtplib.SMTP('smtp.gmail.com', port=587) as server:
                server.starttls()
                server.login('noreply@informatecservicos.com.br', senha)
                server.sendmail(msg['From'], msg['To'], msg.as_string())

            print('Email enviado para', nome)
    except SMTPException as e:
        print("Erro ao enviar email:", e)

@app.route('/diario')
def diario():
    enviar_email()
    return 'E-mails enviados para o Diário de Bordo'

if __name__ == '__main__':
    app.run(debug=True)
