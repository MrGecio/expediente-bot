import smtplib
import imaplib
import email
import os
import json
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- Configuración de entorno / secrets ---
GMAIL_USER = os.environ["GMAIL_USER"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
TO_EMAIL = os.environ["TO_EMAIL"]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
RANGE_NAME = "'Hoja 1'!A2:F100"

#Hasta aqui esta bien

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
service = build("sheets", "v4", credentials=credentials)

#Prueba hasta aqui



# --- Enviar correos ---
def enviar_correos():
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get("values", [])

    if not values:
        print("⚠️ No hay expedientes en la hoja.")
        return

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(GMAIL_USER, GMAIL_APP_PASSWORD)

    for row in values:
        if len(row) < 6:
            continue
        num, tipo, quejoso, fecha_inicio, fecha_limite, descripcion = row
        mensaje = f"""
Expediente pendiente:
Número: {num}
Tipo: {tipo}
Quejoso: {quejoso}
Fecha de inicio: {fecha_inicio}
Fecha límite: {fecha_limite}
Descripción: {descripcion}

👉 Para borrar este expediente, responde a este correo con:
BORRAR {num}
"""
        msg = MIMEText(mensaje)
        msg["From"] = GMAIL_USER
        msg["To"] = TO_EMAIL
        msg["Subject"] = f"Expediente pendiente: {num}"

        server.sendmail(GMAIL_USER, TO_EMAIL, msg.as_string())
        print(f"📤 Expediente {num} enviado a {TO_EMAIL}")

    server.quit()

# --- Revisar respuestas ---
def leer_respuestas():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, '(UNSEEN)')
    correos_ids = messages[0].split()

    for num in correos_ids:
        status, data = mail.fetch(num, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Leer cuerpo del mensaje
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode()
                    procesar_comando(body.strip())
        else:
            body = msg.get_payload(decode=True).decode()
            procesar_comando(body.strip())

    mail.logout()

# --- Procesar comandos de respuesta ---
def procesar_comando(body):
    if body.upper().startswith("BORRAR"):
        try:
            _, num_exp = body.split()
            num_exp = num_exp.strip()

            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
            values = result.get("values", [])

            for i, row in enumerate(values, start=2):
                if row and row[0] == num_exp:
                    service.spreadsheets().values().clear(
                        spreadsheetId=SPREADSHEET_ID,
                        range=f"'Hoja 1'!A{i}:F{i}"
                    ).execute()
                    print(f"🗑️ Expediente {num_exp} borrado de Sheets")
                    return
            print(f"⚠️ No encontré expediente {num_exp}")

        except Exception as e:
            print("❌ Error procesando comando:", e)

if __name__ == "__main__":
    print("📤 Enviando expedientes...")
    enviar_correos()

    print("📥 Revisando respuestas...")
    leer_respuestas()
