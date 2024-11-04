import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Cargar el archivo .ods en un DataFrame
file_path = 'Prueba.ods'
df = pd.read_excel(file_path, engine='odf')

# Extraer los correos electrónicos en la columna C a partir de la segunda fila
email_list = df.iloc[0:, 2].dropna().tolist()

# iloc es un accesor de índice de pandas que permite seleccionar filas y columnas en un DataFrame usando índices numéricos
# DataFrame.iloc[fila, columna]

# Imprimir la lista de correos electrónicos
print(email_list)

remitente = 'cognivia@gmail.com'
password = 'aqtt zdxg ilin vtid'
asunto = 'asuntillo'
cuerpo = 'holi caracoli'

# aqtt zdxg ilin vtid


# Configuración del servidor SMTP de Gmail
with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
    servidor.starttls() # Iniciar conexión TLS segura
    servidor.login(remitente, password) #Iniciar sesión con la contraseña de la aplicación

    # Enviar correo a cada destinatario en la lista
    for destinatario in email_list:
        mensaje = MIMEMultipart()
        mensaje['FROM'] = remitente
        mensaje['TO'] = destinatario
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        # Enviar el mensaje
        servidor.sendmail(remitente, destinatario, mensaje.as_string())
    

    print("ya está todo enviado bro")
