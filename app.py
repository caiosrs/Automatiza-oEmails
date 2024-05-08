from flask import Flask
import pyodbc
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
from smtplib import SMTPException
from datetime import datetime, timedelta

app = Flask(__name__)

def enviar_email():
    hoje = datetime.now()

    if hoje.hour >= 17:
        return

    conn = pyodbc.connect('Driver={SQL Server};Server=ISPD-DP062\\SQLEXPRESS01;Database=DatabaseTeste;Trusted_Connection=yes;')
    cursor = conn.cursor()

    cursor.execute("""select Nome, Email from Usuarios""")
    clientes = cursor.fetchall()

    dia_anterior = hoje - timedelta(days=1)
    cursor.execute("SELECT COUNT(*) FROM Feriados WHERE Data = ?", dia_anterior.strftime('%Y-%m-%d'))
    feriado = cursor.fetchone()[0] > 0

    conn.close()

    if hoje.weekday() == 0 or feriado:
        return

    if hoje.weekday() == 0:
        dia_anterior = hoje - timedelta(days=3)
    else:
        dia_anterior = hoje - timedelta(days=1)

    # dia e mes
    data_titulo = dia_anterior.strftime('%d/%m')

    with open(r'D:\1Desktop\Documentos\My Web Sites\App py\Automatização de Emails\senha temporaria.txt') as f:
        senha = f.read().strip()

    try:
        for cliente in clientes:
            nome, email = cliente
            
            msg = MIMEMultipart()

            msg['Subject'] = f'Preencha o Diário de Bordo do dia {data_titulo}!'
            msg['From'] = "caiosilva.informatec@gmail.com"
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
        <p>Olá Caio,</p>
        <p class="message"><strong>Parece que você ainda não preencheu o seu Diário de Bordo desde a data {data_titulo}.
        <br/>Por favor, não se esqueça de fazer isso o mais rápido possível.</strong></p>
        <p class="message">O Diário de Bordo é uma ferramenta importante para registrar suas atividades diárias e garantir o sucesso de nosso projeto.</p>
        <p class="message">Clique no botão abaixo para acessar o Diário de Bordo:</p>
        <a href="https://dpcontrole.informatecservicos.com.br/" class="button"><strong>Acessar o Diário de Bordo</strong></a>
        <p class="message">Obrigado pela sua atenção!</p>
        <p class="signature">Atenciosamente,<br>Informatec Serviços</p>
        <br/>
    </div>
</body>
</html>
    """

            with open(r'D:\1Desktop\Documentos\My Web Sites\App py\Automatização de Emails\Diario de Bordo LOGO.jpeg', 'rb') as img_file:
                img = MIMEImage(img_file.read())
                img.add_header('Content-Disposition', 'inline', filename='Diario de Bordo LOGO.jpeg')
                img.add_header('Content-ID', '<diario_bordo_logo>')
                msg.attach(img)

            msg.attach(MIMEText(html_content, 'html'))

            with smtplib.SMTP('smtp.gmail.com', port=587) as server:
                server.starttls()
                server.login('caiosilva.informatec@gmail.com', senha)
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