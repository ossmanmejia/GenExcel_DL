import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd

# Configuración de la cuenta de Gmail
sender_email = "ossmanmejia@gmail.com"   # Cambia por tu correo de Gmail
# Cambia por tu contraseña de aplicación
password = "lkxl fzkn eowg ymyy"

# Edita el asunto y el cuerpo del correo
subject = "Otro ensayo - Experiencias para valorar - DL Latam 2025"
body = """\
Bonito día, estimad@ evaluador/@.

Esperamos que este correo te encuentre en bienestar.

En este correo, te adjuntamos el archivo con las experiencias que vas a valorar. 

Vas a encontrar un excel con la información general de la postulación y, además, un enlace para conocer el documento con la planeación de la sesión. Recuerda revisarlo a detalle, pues es allí donde vas a comprender de qué se trata realmente la propuesta.

Además, te pedimos que revises bien las dos pestañas inferiores, allí puedes encontrar las Inmersiones y los Talleres que vas a evaluar

No dudes en preguntarnos cualquier duda que tengas sobre el proceso o sobre las experiencias que vas a evaluar. Puedes escribirnos a este correo o al WhatsApp de Cristina: (57)3174319882.

Nos entusiasma contar contigo. 

Con agradecimiento,
Equipo de aprendizaje - DL Latam 2025
"""

# Cargar la hoja Evaluadores del archivo Excel
archivo = 'postulaciones_actualizadoII.xlsx'
hoja_evaluadores = pd.read_excel(archivo, sheet_name='Evaluadores')

# Iterar sobre cada evaluador
for index, row in hoja_evaluadores.iterrows():
    codigo = row['Código']
    # Ajusta si la columna de email tiene otro nombre
    recipient_email = row['Contacto']

    # Verificar que exista una dirección de correo
    if pd.isna(recipient_email):
        print(f"El evaluador {codigo} no tiene dirección de correo. Se omite.")
        continue

    # Nombre del archivo que se adjuntará
    filename = f'{codigo}.xlsx'

    # Crear el mensaje
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg.set_content(body)

    # Adjuntar el archivo
    try:
        with open(filename, 'rb') as f:
            file_data = f.read()
        msg.add_attachment(file_data,
                           maintype='application',
                           subtype='octet-stream',
                           filename=filename)
    except FileNotFoundError:
        print(
            f"El archivo {filename} no se encontró. Se omite el envío a {recipient_email}.")
        continue

    # Enviar el correo a través del servidor SMTP de Gmail
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.send_message(msg)
        print(f"Correo enviado a {recipient_email} con el archivo {filename}.")
    except Exception as e:
        print(f"Error al enviar correo a {recipient_email}: {e}")

print("Envío de correos completado.")
