import os
import logging
import asyncio
import json
import re
import base64
import pytz
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_JUSTIFY
import io

from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import anthropic

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
TELEGRAM_TOKEN    = os.environ.get('TELEGRAM_TOKEN')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
GMAIL_USER        = os.environ.get('GMAIL_USER')
TELEGRAM_CHAT_ID  = os.environ.get('TELEGRAM_CHAT_ID', '')
GOOGLE_TOKEN_B64  = os.environ.get('GOOGLE_TOKEN_B64')
SHEETS_ID         = os.environ.get('SHEETS_ID', '1saZ98ZEqj46nxcvQC5V0oKJ0GLL-ly0MeuxZqsZHNKI')
TIMEZONE          = 'Europe/Madrid'
SCOPES            = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/tasks',
    'https://www.googleapis.com/auth/spreadsheets'
]

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# Historial de conversación (memoria de contexto, máx 10 turnos)
conversation_history = []
MAX_HISTORY = 10

# ─────────────────────────────────────────────
# GOOGLE SERVICES
# ─────────────────────────────────────────────
def get_credentials():
    if not GOOGLE_TOKEN_B64:
        return None
    try:
        token_data = json.loads(base64.b64decode(GOOGLE_TOKEN_B64).decode())
        creds = Credentials.from_authorized_user_info(token_data, SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    except Exception as e:
        logger.error(f"Error obteniendo credenciales: {e}")
        return None

def get_calendar_service():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return build('calendar', 'v3', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando con Google Calendar: {e}")
        return None

def get_gmail_service():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return build('gmail', 'v1', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando con Gmail: {e}")
        return None

def get_tasks_service():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return build('tasks', 'v1', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando con Google Tasks: {e}")
        return None

# ─────────────────────────────────────────────
# GOOGLE TASKS
# ─────────────────────────────────────────────
def get_tasks(tasklist='@default'):
    service = get_tasks_service()
    if not service:
        return []
    try:
        result = service.tasks().list(
            tasklist=tasklist,
            showCompleted=False,
            maxResults=20
        ).execute()
        return result.get('items', [])
    except Exception as e:
        logger.error(f"Error obteniendo tareas: {e}")
        return []

def create_task(title, notes='', due_date=None):
    service = get_tasks_service()
    if not service:
        return None
    try:
        task = {'title': title}
        if notes:
            task['notes'] = notes
        if due_date:
            task['due'] = due_date + 'T00:00:00.000Z'
        return service.tasks().insert(tasklist='@default', body=task).execute()
    except Exception as e:
        logger.error(f"Error creando tarea: {e}")
        return None

def delete_task(task_id, tasklist='@default'):
    service = get_tasks_service()
    if not service:
        return False
    try:
        service.tasks().delete(tasklist=tasklist, task=task_id).execute()
        return True
    except Exception as e:
        logger.error(f"Error eliminando tarea: {e}")
        return False

def complete_task(task_id, tasklist='@default'):
    service = get_tasks_service()
    if not service:
        return False
    try:
        task = service.tasks().get(tasklist=tasklist, task=task_id).execute()
        task['status'] = 'completed'
        service.tasks().update(tasklist=tasklist, task=task_id, body=task).execute()
        return True
    except Exception as e:
        logger.error(f"Error completando tarea: {e}")
        return False

def find_task_by_name(name):
    tasks = get_tasks()
    name_lower = name.lower()
    for task in tasks:
        if name_lower in task.get('title', '').lower():
            return task
    return None

def format_tasks(tasks):
    if not tasks:
        return "No hay tareas pendientes."
    txt = ""
    for t in tasks:
        due = ''
        if t.get('due'):
            try:
                dt  = datetime.fromisoformat(t['due'].replace('Z', '+00:00'))
                due = f" — vence {dt.strftime('%d/%m')}"
            except Exception:
                pass
        txt += f"• {t['title']}{due}\n"
    return txt

# ─────────────────────────────────────────────
# GOOGLE CALENDAR
# ─────────────────────────────────────────────
def get_events(days=1):
    service = get_calendar_service()
    if not service:
        return []
    try:
        tz  = pytz.timezone(TIMEZONE)
        now = datetime.now(tz)
        end = now + timedelta(days=days)
        result = service.events().list(
            calendarId='primary',
            timeMin=now.isoformat(),
            timeMax=end.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        return result.get('items', [])
    except Exception as e:
        logger.error(f"Error obteniendo eventos: {e}")
        return []

def create_event(summary, date_str, time_str, duration_hours=1, description=''):
    service = get_calendar_service()
    if not service:
        return None
    try:
        tz       = pytz.timezone(TIMEZONE)
        dt_str   = f"{date_str} {time_str}"
        start_dt = tz.localize(datetime.strptime(dt_str, "%Y-%m-%d %H:%M"))
        end_dt   = start_dt + timedelta(hours=duration_hours)
        event    = {
            'summary': summary,
            'description': description,
            'start': {'dateTime': start_dt.isoformat(), 'timeZone': TIMEZONE},
            'end':   {'dateTime': end_dt.isoformat(),   'timeZone': TIMEZONE},
        }
        return service.events().insert(calendarId='primary', body=event).execute()
    except Exception as e:
        logger.error(f"Error creando evento: {e}")
        return None

def find_event_by_name(name):
    service = get_calendar_service()
    if not service:
        return None
    try:
        tz  = pytz.timezone(TIMEZONE)
        now = datetime.now(tz)
        end = now + timedelta(days=30)
        result = service.events().list(
            calendarId='primary',
            timeMin=now.isoformat(),
            timeMax=end.isoformat(),
            q=name,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        items = result.get('items', [])
        return items[0] if items else None
    except Exception as e:
        logger.error(f"Error buscando evento: {e}")
        return None

def update_event(event_id, date_str=None, time_str=None, duration_hours=1):
    service = get_calendar_service()
    if not service:
        return None
    try:
        event = service.events().get(calendarId='primary', eventId=event_id).execute()
        if date_str and time_str:
            tz       = pytz.timezone(TIMEZONE)
            start_dt = tz.localize(datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M"))
            end_dt   = start_dt + timedelta(hours=duration_hours)
            event['start'] = {'dateTime': start_dt.isoformat(), 'timeZone': TIMEZONE}
            event['end']   = {'dateTime': end_dt.isoformat(),   'timeZone': TIMEZONE}
        return service.events().update(calendarId='primary', eventId=event_id, body=event).execute()
    except Exception as e:
        logger.error(f"Error actualizando evento: {e}")
        return None

def delete_event(event_id):
    service = get_calendar_service()
    if not service:
        return False
    try:
        service.events().delete(calendarId='primary', eventId=event_id).execute()
        return True
    except Exception as e:
        logger.error(f"Error eliminando evento: {e}")
        return False

def format_events(events):
    if not events:
        return "No hay eventos."
    tz  = pytz.timezone(TIMEZONE)
    txt = ""
    for ev in events:
        if 'dateTime' not in ev['start']:
            date_str = ev['start'].get('date', '')
            try:
                dt  = datetime.strptime(date_str, '%Y-%m-%d')
                fmt = dt.strftime('%d/%m — Todo el día')
            except Exception:
                fmt = date_str
        else:
            start = ev['start']['dateTime']
            try:
                dt  = datetime.fromisoformat(start.replace('Z', '+00:00')).astimezone(tz)
                fmt = dt.strftime('%d/%m a las %H:%M')
            except Exception:
                fmt = start
        txt += f"• *{ev['summary']}* — {fmt}\n"
    return txt

# ─────────────────────────────────────────────
# GMAIL API
# ─────────────────────────────────────────────
def send_email(to_addr, subject, body_text):
    try:
        service = get_gmail_service()
        if not service:
            logger.error("No se pudo conectar con Gmail API")
            return False

        # Convertir **texto** a <strong>texto</strong> para negrita en HTML
        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')

        html_body = f"""
        <html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">
          <div style="border-top:3px solid #b8960c;padding-top:20px;">
            {formatted}
          </div>
          <div style="margin-top:30px;padding-top:15px;border-top:1px solid #ddd;
                      font-size:0.85em;color:#666;">
            <em>Secretaria Virtual — AP Estudio Jurídico</em>
          </div>
        </body></html>"""

        msg = MIMEText(html_body, 'html', 'utf-8')
        msg['Subject'] = subject
        msg['From']    = GMAIL_USER
        msg['To']      = to_addr

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(
            userId='me',
            body={'raw': raw}
        ).execute()
        logger.info(f"Email enviado a {to_addr}")
        return True
    except Exception as e:
        logger.error(f"Error enviando email via Gmail API: {e}")
        return False


# ─────────────────────────────────────────────
# GENERADOR DE FACTURAS PDF
# ─────────────────────────────────────────────
def generar_factura(num_factura, cliente_nombre, cliente_nif, cliente_domicilio,
                    concepto, base_imponible, iva=21, retencion=0):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=2.5*cm, leftMargin=2.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    story = []

    normal   = ParagraphStyle('normal',  fontSize=10, fontName='Helvetica', leading=14)
    negrita  = ParagraphStyle('negrita', fontSize=10, fontName='Helvetica-Bold', leading=14)
    derecha  = ParagraphStyle('derecha', fontSize=10, fontName='Helvetica', alignment=TA_RIGHT)
    der_bold = ParagraphStyle('derbold', fontSize=10, fontName='Helvetica-Bold', alignment=TA_RIGHT)
    centrado = ParagraphStyle('centrado',fontSize=10, fontName='Helvetica-Bold', alignment=TA_CENTER)
    small    = ParagraphStyle('small',   fontSize=9,  fontName='Helvetica', alignment=TA_CENTER)
    justif   = ParagraphStyle('justif',  fontSize=10, fontName='Helvetica', leading=15, alignment=TA_JUSTIFY)

    tz  = pytz.timezone(TIMEZONE)
    hoy = datetime.now(tz).strftime('%d-%m-%Y')

    cab_data = [[
        Paragraph('', normal),
        Paragraph('<b>Adrià Paños Ruiz</b><br/>Carrer Comte Ramon Berenguer 1-3, esc. B, 2º 1ª<br/>08204 Sabadell Barcelona<br/>47182626N', derecha)
    ]]
    cab_table = Table(cab_data, colWidths=[8*cm, 8*cm])
    cab_table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(cab_table)
    story.append(Spacer(1, 1.5*cm))

    story.append(Paragraph(hoy, derecha))
    story.append(Spacer(1, 0.8*cm))
    story.append(Paragraph(f'Factura núm. {num_factura}', normal))
    story.append(Spacer(1, 0.8*cm))

    texto_cuerpo = (
        f'<b>MINUTA DE HONORARIOS</b> devengados por el despacho a cargo de '
        f'<b>{cliente_nombre.upper()}</b>, con domicilio en {cliente_domicilio}, '
        f'provisto de NIF/CIF núm. {cliente_nif}, comprensiva de la siguiente actuación profesional:'
    )
    story.append(Paragraph(texto_cuerpo, justif))
    story.append(Spacer(1, 0.8*cm))

    concepto_lines = concepto.replace('\\n', '\n').replace('\n', '<br/>')
    concepto_data = [[
        Paragraph(concepto_lines, normal),
        Paragraph(f'{base_imponible:.2f} €', derecha)
    ]]
    concepto_table = Table(concepto_data, colWidths=[13*cm, 3*cm])
    concepto_table.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
    ]))
    story.append(concepto_table)
    story.append(Spacer(1, 0.8*cm))

    iva_importe   = round(base_imponible * iva / 100, 2)
    retencion_imp = round(base_imponible * retencion / 100, 2)
    total         = round(base_imponible + iva_importe - retencion_imp, 2)

    totales_rows = [
        [Paragraph('<b>TOTAL HONORARIOS</b>', negrita), Paragraph(f'<b>{base_imponible:.2f} €</b>', der_bold)],
        [Paragraph(f'{iva}% de IVA, euros', normal),    Paragraph(f'{iva_importe:.2f} €', derecha)],
    ]
    if retencion > 0:
        totales_rows.append([Paragraph(f'{retencion}% de retención IRPF, euros', normal), Paragraph(f'{retencion_imp:.2f} €', derecha)])
    totales_rows.append([Paragraph('<b>TOTAL MINUTA, EUROS</b>', negrita), Paragraph(f'<b>{total:.2f} €</b>', der_bold)])

    tot_table = Table(totales_rows, colWidths=[13*cm, 3*cm])
    tot_table.setStyle(TableStyle([
        ('VALIGN',  (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',    (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LINEABOVE', (0,0),  (-1,0),  0.5, colors.black),
        ('LINEBELOW', (0,0),  (-1,0),  0.5, colors.black),
        ('LINEABOVE', (0,-1), (-1,-1), 0.5, colors.black),
        ('LINEBELOW', (0,-1), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#f0f0f0')),
    ]))
    story.append(tot_table)
    story.append(Spacer(1, 1.5*cm))

    story.append(Paragraph('<b>Forma de pago: transferencia bancaria</b>', centrado))
    story.append(Paragraph('<b>C/C núm. ES1015632626313269891055</b>', centrado))
    story.append(Spacer(1, 2*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.black, spaceAfter=6))
    story.append(Paragraph('Teléfono: 603690659', small))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue(), total

def send_email_with_pdf(to_addr, subject, body_text, pdf_bytes, pdf_filename):
    try:
        service = get_gmail_service()
        if not service:
            return False

        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From']    = GMAIL_USER
        msg['To']      = to_addr

        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')
        html_body = (
            '<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">'
            f'<div style="text-align:center;margin-bottom:24px;">'  
            f'<img src="data:image/png;base64,{LOGO_B64}" alt="AP Estudio Juridico" style="height:80px;width:auto;"/>'
            '</div>'
            f'<div>{formatted}</div>'
            '<hr style="border:none;border-top:1px solid #ddd;margin:24px 0;">'
            '<div style="font-size:0.85em;color:#666;text-align:center;"><em>AP Estudio Juridico</em></div>'
            '</body></html>')
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        part = MIMEBase('application', 'pdf')
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
        part.add_header('Content-Type', 'application/pdf', name=pdf_filename)
        msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()
        logger.info(f"Email con PDF enviado a {to_addr}")
        return True
    except Exception as e:
        logger.error(f"Error enviando email con PDF: {e}")
        return False

# ─────────────────────────────────────────────
# CLAUDE
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# GOOGLE SHEETS — BASE DE DATOS DESPACHO
# ─────────────────────────────────────────────
def get_sheets_service():
    try:
        creds = get_credentials()
        return build('sheets', 'v4', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando Sheets: {e}")
        return None

def sheets_read(rango):
    try:
        svc = get_sheets_service()
        if not svc:
            return []
        if '!' in rango:
            hoja, celdas = rango.split('!', 1)
            meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
            hojas_reales = {s['properties']['title'].lower(): s['properties']['title'] for s in meta.get('sheets', [])}
            hoja_real = hojas_reales.get(hoja.lower(), hoja)
            rango = f"{hoja_real}!{celdas}"
        result = svc.spreadsheets().values().get(spreadsheetId=SHEETS_ID, range=rango).execute()
        return result.get('values', [])
    except Exception as e:
        logger.error(f"sheets_read error: {e}")
        return []

def sheets_append(hoja, valores):
    try:
        svc = get_sheets_service()
        if not svc:
            return False
        meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
        hojas_reales = {s['properties']['title'].lower(): s['properties']['title'] for s in meta.get('sheets', [])}
        hoja_real = hojas_reales.get(hoja.lower(), hoja)
        svc.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID, range=f"{hoja_real}!A1",
            valueInputOption='USER_ENTERED', body={'values': [valores]}
        ).execute()
        return True
    except Exception as e:
        logger.error(f"sheets_append error: {e}")
        return False

def sheets_update_cell(rango, valor):
    try:
        svc = get_sheets_service()
        if not svc:
            return False
        if '!' in rango:
            hoja, celdas = rango.split('!', 1)
            meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
            hojas_reales = {s['properties']['title'].lower(): s['properties']['title'] for s in meta.get('sheets', [])}
            hoja_real = hojas_reales.get(hoja.lower(), hoja)
            rango = f"{hoja_real}!{celdas}"
        svc.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID, range=rango,
            valueInputOption='USER_ENTERED', body={'values': [[valor]]}
        ).execute()
        return True
    except Exception as e:
        logger.error(f"sheets_update error: {e}")
        return False

def normalizar(texto):
    """Elimina acentos y pasa a minúsculas para comparación."""
    import unicodedata
    texto = str(texto).lower()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def get_cliente(nombre):
    """Columnas: A=ID(0), B=Cliente(1), C=NIF(2), D=Dir(3), E=CP(4), F=Pobl(5), G=Prov(6), H=País(7), I=Email(8), J=Tel(9)"""
    rows = sheets_read("Clientes!A2:J200")
    nombre_norm = normalizar(nombre)
    for row in rows:
        if len(row) < 2:
            continue
        nombre_norm2 = normalizar(str(row[1])) if len(row) > 1 else ''
        if nombre_norm in nombre_norm2 or nombre_norm2 in nombre_norm:
            def col(i): return str(row[i]).strip() if len(row) > i else ''
            return {
                'id':        col(0),
                'nombre':    col(1),
                'nif':       col(2),
                'direccion': col(3),
                'cp':        col(4),
                'poblacion': col(5),
                'provincia': col(6),
                'pais':      col(7),
                'email':     col(8),
                'telefono':  col(9),
                'apellidos': '', 'tipo': '', 'fecha_alta': '', 'notas': '',
            }
    return None

def get_todos_clientes():
    rows = sheets_read("Clientes!A2:J200")
    clientes = []
    for row in rows:
        if len(row) >= 2 and row[0]:
            clientes.append({'id': row[0], 'nombre': row[1] if len(row) > 1 else '',
                'nif': row[2] if len(row) > 2 else '',
                'email': row[8] if len(row) > 8 else '',
                'telefono': row[5] if len(row) > 5 else '',
                'tipo': row[9] if len(row) > 9 else ''})
    return clientes

def get_casos_cliente(nombre_cliente=None):
    rows = sheets_read("Casos!A2:N200")
    casos = []
    for row in rows:
        if not row or not row[0]:
            continue
        if nombre_cliente and nombre_cliente.lower() not in (row[2] if len(row) > 2 else '').lower():
            continue
        casos.append({
            'id': row[0], 'cliente': row[2] if len(row) > 2 else '',
            'tipo': row[3] if len(row) > 3 else '', 'materia': row[4] if len(row) > 4 else '',
            'descripcion': row[5] if len(row) > 5 else '', 'juzgado': row[6] if len(row) > 6 else '',
            'autos': row[7] if len(row) > 7 else '', 'estado': row[8] if len(row) > 8 else '',
            'fecha_apertura': row[9] if len(row) > 9 else '',
            'proxima_actuacion': row[10] if len(row) > 10 else '',
            'fecha_actuacion': row[11] if len(row) > 11 else '',
            'honorarios': row[12] if len(row) > 12 else '',
            'cobrado': row[13] if len(row) > 13 else ''})
    return casos

def get_facturas(estado=None):
    rows = sheets_read("Facturas!A2:K200")
    facturas = []
    for row in rows:
        if not row or not row[0]:
            continue
        est = row[9] if len(row) > 9 else ''
        if estado and estado.lower() not in est.lower():
            continue
        facturas.append({
            'num': row[0], 'cliente': row[2] if len(row) > 2 else '',
            'fecha': row[3] if len(row) > 3 else '', 'concepto': row[4] if len(row) > 4 else '',
            'base': row[5] if len(row) > 5 else '', 'total': row[8] if len(row) > 8 else '',
            'estado': est, 'fecha_cobro': row[10] if len(row) > 10 else ''})
    return facturas

def siguiente_id_cliente():
    rows = sheets_read("Clientes!A2:A200")
    ids = [int(r[0]) for r in rows if r and str(r[0]).isdigit()]
    return max(ids) + 1 if ids else 1

def siguiente_num_factura():
    """Lee todas las facturas y devuelve el siguiente número correlativo (solo número)."""
    rows = sheets_read("Facturas!C2:C200")  # Columna C = Numero factura
    nums = []
    for r in rows:
        if r and r[0]:
            try:
                nums.append(int(str(r[0]).strip()))
            except:
                pass
    return str(max(nums) + 1 if nums else 1)


def insertar_factura_en_sheets(fecha, num_factura, cliente, nif, base, iva_amount, ret_amount, total):
    """Inserta una factura antes de la fila de totales, desplaza totales y actualiza sumas."""
    try:
        svc = get_sheets_service()
        if not svc:
            return False

        hoja_real = 'Facturas'

        # Obtener metadatos para saber el sheetId
        meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
        sheet_id = None
        for s in meta.get('sheets', []):
            if s['properties']['title'] == hoja_real:
                sheet_id = s['properties']['sheetId']
                break

        # Leer columna A para encontrar última fila con datos
        rows = sheets_read(f"{hoja_real}!A2:A200")
        ultima_fila_datos = 1
        for i, r in enumerate(rows, 2):
            if r and str(r[0]).strip().isdigit():
                ultima_fila_datos = i

        # La fila de inserción es después de la última factura
        fila_insercion = ultima_fila_datos + 1  # fila en blanco
        fila_nueva_factura = ultima_fila_datos + 2  # nueva factura aquí... 
        # En realidad: insertar ANTES del total, dejando 1 fila en blanco entre última factura y total

        # Obtener ID correlativo
        id_rows = sheets_read(f"{hoja_real}!A2:A200")
        ids = [int(str(r[0]).strip()) for r in id_rows if r and str(r[0]).strip().isdigit()]
        nuevo_id = max(ids) + 1 if ids else 36

        # Insertar 1 fila en blanco + 1 para la nueva factura (2 filas) antes de la fila de totales
        # Primero, insertar 1 fila vacía después de la última factura
        insert_request = {
            'insertDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': ultima_fila_datos,  # 0-indexed
                    'endIndex': ultima_fila_datos + 1
                },
                'inheritFromBefore': True
            }
        }
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SHEETS_ID,
            body={'requests': [insert_request]}
        ).execute()

        # Escribir datos de la nueva factura en la fila insertada
        nueva_fila = ultima_fila_datos + 1  # 1-indexed, justo la que acabamos de insertar
        valores = [[
            nuevo_id, fecha, num_factura, 0,
            cliente, nif,
            round(base, 2), round(iva_amount, 2),
            round(ret_amount, 2), round(total, 2),
            '', '', 'Emitida', 'ORDINARIA'
        ]]
        svc.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"{hoja_real}!A{nueva_fila}:N{nueva_fila}",
            valueInputOption='USER_ENTERED',
            body={'values': valores}
        ).execute()

        # Actualizar fórmulas de totales (fila totales = nueva_fila + 2, con 1 en blanco)
        fila_total = nueva_fila + 2
        # Columnas con sumas: G=base, H=iva, I=irpf, J=total, D=pendiente
        for col_letra, col_num in [('D',4),('G',7),('H',8),('I',9),('J',10)]:
            formula = f"=SUM({col_letra}2:{col_letra}{nueva_fila})"
            svc.spreadsheets().values().update(
                spreadsheetId=SHEETS_ID,
                range=f"{hoja_real}!{col_letra}{fila_total}",
                valueInputOption='USER_ENTERED',
                body={'values': [[formula]]}
            ).execute()

        # Escribir etiqueta "Total" si no existe
        svc.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"{hoja_real}!A{fila_total}",
            valueInputOption='USER_ENTERED',
            body={'values': [['Total']]}
        ).execute()

        # Aplicar formato moneda (€) a columnas G, H, I, J, D
        format_requests = []
        currency_format = {
            "numberFormat": {
                "type": "CURRENCY",
                "pattern": '#,##0.00 "€"'
            }
        }
        for col_index in [3, 6, 7, 8, 9]:  # D=3, G=6, H=7, I=8, J=9 (0-indexed)
            format_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": nueva_fila - 1,  # 0-indexed
                        "endRowIndex": nueva_fila,
                        "startColumnIndex": col_index,
                        "endColumnIndex": col_index + 1
                    },
                    "cell": {"userEnteredFormat": currency_format},
                    "fields": "userEnteredFormat.numberFormat"
                }
            })
        if format_requests:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SHEETS_ID,
                body={'requests': format_requests}
            ).execute()

        logger.info(f"Factura {num_factura} insertada en fila {nueva_fila}, totales en fila {fila_total}")
        return True

    except Exception as e:
        logger.error(f"Error insertando factura en Sheets: {e}")
        return False

def get_bbdd_context():
    try:
        clientes = get_todos_clientes()
        casos = get_casos_cliente()
        facturas_pend = get_facturas(estado='Pendiente')
        ctx = f"BASE DE DATOS DEL DESPACHO:\n"
        ctx += f"- {len(clientes)} clientes, {len(casos)} casos, {len(facturas_pend)} facturas pendientes\n\n"
        if clientes:
            ctx += "CLIENTES:\n"
            for c in clientes[:20]:
                ctx += f"  [{c['id']}] {c['nombre']} | NIF: {c['nif']} | Email: {c['email']} | Tel: {c['telefono']}\n"
        if casos:
            ctx += "\nCASOS:\n"
            for c in casos[:20]:
                ctx += f"  [{c['id']}] {c['cliente']} — {c['materia']} | {c['autos']} | {c['estado']} | Próx: {c['proxima_actuacion']} ({c['fecha_actuacion']})\n"
        if facturas_pend:
            ctx += "\nFACTURAS PENDIENTES:\n"
            for f in facturas_pend:
                ctx += f"  {f['num']} | {f['cliente']} | {f['total']} € | {f['concepto'][:40]}\n"
        return ctx
    except Exception as e:
        logger.error(f"Error get_bbdd_context: {e}")
        return ""

SYSTEM_PROMPT = """Eres la secretaria virtual del despacho de abogados AP Estudio Jurídico.
Ayudas al abogado Adrià con agenda, emails, tareas y facturación.

Responde SIEMPRE con uno o varios JSON válidos, uno por línea. NUNCA añadas texto fuera de los JSON.
Para respuestas conversacionales usa: {{"action":"none","response":"texto"}}

══════════════════════════════════════
ACCIONES DISPONIBLES
══════════════════════════════════════

AGENDA:
{{"action":"create_event","summary":"título","date":"YYYY-MM-DD","time":"HH:MM","duration_hours":1,"description":""}}
{{"action":"update_event","event_name":"nombre","date":"YYYY-MM-DD","time":"HH:MM","duration_hours":1}}
{{"action":"delete_event","event_name":"nombre"}}
{{"action":"query_calendar","days":7}}

EMAIL:
{{"action":"send_email","to":"email@ejemplo.com","subject":"Asunto","body":"Cuerpo"}}

TAREAS:
{{"action":"query_tasks"}}
{{"action":"create_task","title":"título","notes":"","due_date":"YYYY-MM-DD"}}
{{"action":"delete_task","task_name":"nombre"}}
{{"action":"complete_task","task_name":"nombre"}}

FACTURAS (CRÍTICO — leer bien):
{{"action":"create_invoice_bd","cliente":"nombre del cliente","concepto":"descripción","base_imponible":500.00,"es_base":true,"iva":21,"retencion":0}}
- USA SIEMPRE create_invoice_bd cuando el usuario mencione un nombre de cliente. El sistema obtiene NIF, domicilio y número de factura AUTOMÁTICAMENTE de la base de datos. NUNCA los pidas.
- "es_base":true → el importe indicado es la base (sin IVA). Por defecto si el usuario dice "honorarios" o no especifica.
- "es_base":false → el importe indicado es el TOTAL final con IVA. Úsalo cuando el usuario diga "total X€", "que el total sea X€", "importe total X€".
- Ejemplo: "factura con total 40€" → {{"action":"create_invoice_bd","cliente":"nombre","concepto":"visita","base_imponible":40,"es_base":false,"iva":21,"retencion":0}}
- IVA: SIEMPRE 21% salvo que el abogado indique otro valor expresamente.
- Retención IRPF: SOLO si el abogado lo indica expresamente. Por defecto siempre 0.

BASE DE DATOS:
{{"action":"query_cliente","nombre":"nombre"}}
{{"action":"query_clientes"}}
{{"action":"add_cliente","nombre":"","nif":"","email":"","telefono":"","direccion":"","poblacion":"","cp":"","provincia":"","pais":"España"}}
{{"action":"query_casos","cliente":"nombre opcional"}}
{{"action":"add_caso","cliente":"","materia":"","descripcion":"","juzgado":"","autos":"","estado":"Activo","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD"}}
{{"action":"update_caso_estado","autos":"PA 1/2026","estado":"","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD"}}
{{"action":"query_facturas","estado":"Pendiente"}}
{{"action":"cobrar_factura","num_factura":"15","fecha_cobro":"YYYY-MM-DD"}}

══════════════════════════════════════
REGLAS GENERALES
══════════════════════════════════════
- Responde siempre en español, de forma concisa y profesional.
- MÚLTIPLES ACCIONES: si el abogado pide varias cosas, responde con varios JSON seguidos, uno por línea.
- Fechas relativas: calcúlalas a partir de hoy {today}. Día de la semana actual: {weekday} (0=lunes…6=domingo).
- "el lunes" = próximo lunes. Si hoy es lunes, es el lunes de la semana que viene.
- Para mover evento: update_event. Para cancelar: delete_event.
- due_date en create_task es opcional, solo si el abogado indica fecha límite.
- NUNCA pidas datos que el sistema puede obtener automáticamente de la BD (NIF, domicilio, número de factura, email del cliente).
- Solo pide datos con action:none si son estrictamente necesarios y no están en la BD (por ejemplo, un cliente nuevo).
"""

def ask_claude(user_msg, calendar_context=""):
    global conversation_history
    tz    = pytz.timezone(TIMEZONE)
    today   = datetime.now(tz).strftime('%d/%m/%Y, %A')
    weekday = str(datetime.now(tz).weekday())
    system  = SYSTEM_PROMPT.replace('{today}', today).replace('{weekday}', weekday)

    content = user_msg
    bd_keywords = ['cliente','clientes','caso','casos','factura','facturas','expediente',
                   'cobrar','pendiente','debe','honorario','autos','juzgado','nif','email',
                   'teléfono','dirección','añade','añadir','nuevo cliente','nuevo caso']
    bbdd_ctx = ""
    if any(kw in user_msg.lower() for kw in bd_keywords):
        bbdd_ctx = get_bbdd_context()
    if calendar_context and bbdd_ctx:
        content = f"Contexto actual del calendario:\n{calendar_context}\n\n{bbdd_ctx}\n\nMensaje del abogado: {user_msg}"
    elif calendar_context:
        content = f"Contexto actual del calendario:\n{calendar_context}\n\nMensaje del abogado: {user_msg}"
    elif bbdd_ctx:
        content = f"{bbdd_ctx}\n\nMensaje del abogado: {user_msg}"

    # Añadir mensaje del usuario al historial
    conversation_history.append({"role": "user", "content": content})

    # Mantener solo los últimos MAX_HISTORY turnos
    if len(conversation_history) > MAX_HISTORY * 2:
        conversation_history = conversation_history[-(MAX_HISTORY * 2):]

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        system=system,
        messages=conversation_history
    )
    reply = resp.content[0].text.strip()

    # Guardar respuesta en historial
    conversation_history.append({"role": "assistant", "content": reply})

    return reply

# ─────────────────────────────────────────────
# RESUMEN DIARIO (Lunes a Viernes 7:00 AM)
# ─────────────────────────────────────────────
async def daily_summary(bot):
    chat_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    tz      = pytz.timezone(TIMEZONE)
    today   = datetime.now(tz)

    events_today = get_events(days=1)
    events_week  = get_events(days=7)
    tasks        = get_tasks()

    ctx = (
        f"Hoy es {today.strftime('%A %d/%m/%Y')}.\n\n"
        f"Eventos de hoy:\n{format_events(events_today)}\n\n"
        f"Resto de la semana:\n{format_events(events_week)}"
    )

    prompt = (
        f"{ctx}\n\n"
        "Genera un resumen matutino MUY esquemático. "
        "Formato: primero los eventos de hoy con hora y título, luego los de la semana. "
        "Sin saludos, sin frases motivadoras. Solo datos. Máximo 80 palabras."
    )

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}]
    )
    calendar_summary = resp.content[0].text.strip()

    tasks_text = format_tasks(tasks)
    body = f"{calendar_summary}\n\n**TAREAS PENDIENTES**\n\n{tasks_text}"

    send_email(
        GMAIL_USER,
        f"📅 Agenda {today.strftime('%d/%m/%Y')}",
        body
    )
    logger.info("Resumen diario enviado.")

# ─────────────────────────────────────────────
# RESUMEN SEMANAL (Sábados 9:00 AM)
# ─────────────────────────────────────────────
async def weekly_summary(bot):
    tz    = pytz.timezone(TIMEZONE)
    today = datetime.now(tz)

    # Calcular lunes y viernes de la semana siguiente
    days_until_monday = (7 - today.weekday()) % 7 or 7
    next_monday = today + timedelta(days=days_until_monday)
    next_friday = next_monday + timedelta(days=4)

    monday_start = next_monday.replace(hour=0,  minute=0,  second=0, microsecond=0)
    friday_end   = next_friday.replace(hour=23, minute=59, second=59, microsecond=0)

    service = get_calendar_service()
    events = []
    if service:
        try:
            result = service.events().list(
                calendarId='primary',
                timeMin=monday_start.isoformat(),
                timeMax=friday_end.isoformat(),
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            events = result.get('items', [])
        except Exception as e:
            logger.error(f"Error obteniendo eventos semanales: {e}")

    subject = f"\U0001f4c5 Agenda semanal \u2014 {next_monday.strftime('%d/%m')} al {next_friday.strftime('%d/%m/%Y')}"

    events_text = format_events(events) if events else "Sin eventos."
    prompt = (
        f"Agenda semana {next_monday.strftime('%d/%m')} al {next_friday.strftime('%d/%m/%Y')}:\n\n"
        f"{events_text}\n\n"
        "Genera un resumen esquemático. Solo datos: día, hora y evento. "
        "Sin saludos ni frases. Máximo 80 palabras."
    )

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}]
    )
    calendar_summary = resp.content[0].text.strip()

    tasks      = get_tasks()
    tasks_text = format_tasks(tasks)
    body       = f"**AGENDA SEMANA {next_monday.strftime('%d/%m')} AL {next_friday.strftime('%d/%m/%Y')}**\n\n{calendar_summary}\n\n**TAREAS PENDIENTES**\n\n{tasks_text}"

    send_email(GMAIL_USER, subject, body)
    logger.info("Resumen semanal enviado.")

# ─────────────────────────────────────────────
# HANDLERS TELEGRAM
# ─────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    allowed_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    if allowed_id and chat_id != str(allowed_id):
        logger.warning(f"Acceso denegado en /start a chat_id: {chat_id}")
        return
    os.environ['TELEGRAM_CHAT_ID'] = chat_id
    text = (
        "👋 *¡Buenos días, Adrià!* Soy su secretaria virtual.\n\n"
        "Puedo ayudarle con:\n"
        "📅 *Agenda* — _\"Cita con García el lunes a las 10\"_\n"
        "🔍 *Consultas* — _\"¿Qué tengo esta semana?\"_\n"
        "📧 *Emails* — _\"Envía email al procurador sobre el caso López\"_\n"
        "🔔 *Recordatorios* — _\"¿Qué plazos tengo esta semana?\"_\n\n"
        "¿En qué le puedo ayudar?"
    )
    await update.message.reply_text(text, parse_mode='Markdown')


async def cmd_initbbdd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Inicializa la base de datos con hojas y datos de ejemplo."""
    allowed_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    if allowed_id and str(update.effective_chat.id) != str(allowed_id):
        return

    await update.message.reply_text("⚙️ Inicializando base de datos...")

    try:
        svc = get_sheets_service()
        if not svc:
            await update.message.reply_text("❌ No se pudo conectar con Google Sheets.")
            return

        # Obtener hojas existentes
        meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
        hojas_existentes = [s['properties']['title'] for s in meta.get('sheets', [])]

        requests = []

        # Renombrar "Hoja 1" a "Clientes" si existe
        for s in meta.get('sheets', []):
            if s['properties']['title'] in ['Hoja 1', 'Sheet1', 'Hoja1']:
                requests.append({
                    'updateSheetProperties': {
                        'properties': {'sheetId': s['properties']['sheetId'], 'title': 'Clientes'},
                        'fields': 'title'
                    }
                })
                hojas_existentes = ['Clientes' if h in ['Hoja 1', 'Sheet1', 'Hoja1'] else h for h in hojas_existentes]
                break

        # Añadir hojas que falten
        for hoja in ['Casos', 'Facturas']:
            if hoja not in hojas_existentes:
                requests.append({'addSheet': {'properties': {'title': hoja}}})

        if requests:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SHEETS_ID,
                body={'requests': requests}
            ).execute()

        # Cabeceras y datos
        datos = {
            'Clientes!A1:L1': [['ID Cliente','Nombre','Apellidos','NIF/CIF','Email','Teléfono','Dirección','Población','CP','Tipo','Fecha Alta','Notas']],
            'Clientes!A2:L11': [
                ['1','Marc','Berbel Saura','12345678A','berbelmarc04@gmail.com','611 234 567','Carrer Margenat 30, 2º 1ª','Sabadell','08203','Particular','2024-01-15','Caso laboral pendiente'],
                ['2','Sindicato Reformista de Policías','','B12345678','contacto@srp.es','93 456 78 90','Avenida Foietes 10, esc. 5, 1A','Benidorm','03501','Empresa','2024-02-03','Asesoramiento legal continuado'],
                ['3','Rhama','Belouafi Chakir','23456789B','rhama.belouafi@gmail.com','622 345 678','Carrer Sant Joan 12, 3º 2ª','Terrassa','08221','Particular','2024-03-10','Querella por negligencia médica'],
                ['4','Francisco','Rodríguez Contreras','34567890C','frc@hotmail.com','633 456 789','Avinguda Catalunya 45, 1º','Sabadell','08204','Particular','2024-04-22','Juicio del Jurado 2/2024'],
                ['5','Antonia','Martínez Vidal','45678901D','antonia.mv@gmail.com','644 567 890','Carrer Gràcia 8, 4º 1ª','Sabadell','08205','Particular','2024-05-05','Causa penal en curso'],
                ['6','Construcciones García SL','','B23456789','admin@construgarcia.com','93 567 89 01','Polígon Industrial Nord, Nau 3','Barberà del Vallès','08210','Empresa','2024-06-18','Reclamación subcontrata'],
                ['7','Marta','Ramón Vidal','56789012E','marta.ramon@outlook.com','655 678 901','Passeig de la Concòrdia 22, 2º','Sabadell','08206','Particular','2024-07-30','Vista oral programada 12/03'],
                ['8','Miquel','Moreno Puig','67890123F','miquel.moreno@gmail.com','666 789 012','Carrer del Nord 5, bajos','Cerdanyola del Vallès','08290','Particular','2024-08-14','Demanda civil pendiente'],
                ['9','Consuelo','Álvarez García','78901234G','consuelo.ag@yahoo.es','677 890 123','Rambla Pompeu Fabra 33, 5º 3ª','Sabadell','08207','Particular','2024-09-01','Recurso C-A en tramitación'],
                ['10','Transportes Kabouri SL','','B34567890','info@tkabouri.com','93 678 90 12','Carrer Indústria 15, Nau 7','Sabadell','08208','Empresa','2024-10-20','Documentación pendiente'],
            ],
            'Casos!A1:N1': [['ID Caso','ID Cliente','Cliente','Tipo','Materia','Descripción','Juzgado','Nº Autos','Estado','Fecha Apertura','Próxima Actuación','Fecha Act.','Honorarios (€)','Cobrado (€)']],
            'Casos!A2:N11': [
                ['1','3','Belouafi','Penal','Negligencia médica','Querella contra médicos del Consorci Sanitari','Juzgado Instrucción 3 Terrassa','DP 45/2024','En instrucción','2024-03-10','Declaración imputados','2026-03-20','3500','1500'],
                ['2','4','Rodríguez','Penal','Violencia de género','Juicio del Jurado — defensa acusado','Juzgado Violencia Mujer 1 Sabadell','TJ 2/2024','Juicio oral','2024-04-22','Vista oral','2026-04-15','4200','2000'],
                ['3','5','Martínez','Penal','Estafa','Causa penal por presunta estafa inmobiliaria','Juzgado Instrucción 5 Sabadell','PA 112/2024','Fase intermedia','2024-05-05','Escrito acusación','2026-03-25','2800','2800'],
                ['4','7','Ramón','Civil','Divorcio contencioso','Divorcio con hijos menores y bienes','Juzgado Familia 2 Sabadell','JV 78/2024','Juicio oral','2024-07-30','Vista oral','2026-03-12','2200','1100'],
                ['5','8','Moreno','Civil','Reclamación cantidad','Demanda por impago de servicios profesionales','Juzgado 1ª Instancia 4 Cerdanyola','JO 234/2024','Admitida','2024-08-14','Contestación demanda','2026-03-30','1500','750'],
                ['6','9','Álvarez','Penal','Recurso multa','Recurso contra auto imposición multa — art. 52 CP','Sección Instrucción TI Sabadell','PA 89/2023','Recurso pendiente','2024-09-01','Resolución recurso','2026-04-01','900','900'],
                ['7','10','Kabouri','Laboral','Documentación','Incumplimiento obligaciones documentales empresa','Juzgado Social 3 Sabadell','AS 567/2024','Pendiente','2024-10-20','Entrega documentación','2026-03-13','1800','0'],
                ['8','2','SRP','Penal','Asesoramiento','Asesoramiento legal continuado al sindicato','-','-','Activo','2024-02-03','Reunión mensual','2026-04-07','743.76','743.76'],
                ['9','6','Construcciones G.','Civil','Reclamación subcontrata','Reclamación por trabajos no pagados — 45.000€','Juzgado 1ª Instancia 6 Barberà','JO 123/2025','Admitida','2025-01-10','Audiencia previa','2026-03-10','3200','1600'],
                ['10','1','Berbel','Laboral','Despido improcedente','Despido objetivo impugnado — reclamación 18.000€','Juzgado Social 1 Sabadell','AS 234/2025','Conciliación','2025-02-20','Acto conciliación','2026-03-18','1800','900'],
            ],
            'Facturas!A1:K1': [['Nº Factura','ID Cliente','Cliente','Fecha','Concepto','Base (€)','IVA 21% (€)','Retención IRPF (€)','Total (€)','Estado','Fecha Cobro']],
            'Facturas!A2:K9': [
                ['1/2026','2','SRP','2026-03-07','Honorarios asesoramiento legal Marzo 2026','61.98','13.02','4.34','70.66','Cobrada','2026-03-07'],
                ['2/2026','4','Rodríguez','2026-02-28','Provisión de fondos juicio oral TJ 2/2024','1000','210','150','1060','Cobrada','2026-03-01'],
                ['3/2026','7','Ramón','2026-02-15','Minuta honorarios preparación vista oral JV 78/2024','550','115.5','82.5','583','Cobrada','2026-02-20'],
                ['4/2026','3','Belouafi','2026-02-01','Provisión de fondos diligencias DP 45/2024','750','157.5','0','907.5','Pendiente',''],
                ['5/2026','9','Álvarez','2026-01-20','Honorarios recurso auto multa PA 89/2023','900','189','135','954','Cobrada','2026-01-25'],
                ['6/2026','5','Martínez','2026-01-10','Honorarios fase intermedia PA 112/2024','800','168','120','848','Cobrada','2026-01-15'],
                ['7/2026','1','Berbel','2026-01-05','Provisión de fondos acto conciliación AS 234/2025','450','94.5','67.5','477','Pendiente',''],
                ['8/2026','6','Construcciones G.','2026-03-01','Honorarios audiencia previa JO 123/2025','600','126','90','636','Emitida',''],
            ],
        }

        for rango, valores in datos.items():
            svc.spreadsheets().values().update(
                spreadsheetId=SHEETS_ID,
                range=rango,
                valueInputOption='USER_ENTERED',
                body={'values': valores}
            ).execute()

        await update.message.reply_text(
            "✅ *Base de datos inicializada*\n\n"
            "• 10 clientes\n"
            "• 10 casos\n"
            "• 8 facturas\n\n"
            "Ya puede consultar: _\"Dame los datos de Berbel\"_",
            parse_mode='Markdown'
        )

    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

async def cmd_bbdd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    allowed_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    if allowed_id and str(update.effective_chat.id) != str(allowed_id):
        return
    await update.message.reply_text("🔍 Comprobando conexión con la base de datos...")
    try:
        svc = get_sheets_service()
        if not svc:
            await update.message.reply_text("❌ No se pudo conectar (servicio nulo).")
            return
        meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
        titulo = meta.get('properties', {}).get('title', '?')
        hojas = [s['properties']['title'] for s in meta.get('sheets', [])]
        msg = f"✅ Conectado a: *{titulo}*\nHojas: {', '.join(hojas)}\n\n"
        for h in hojas:
            try:
                rows = svc.spreadsheets().values().get(
                    spreadsheetId=SHEETS_ID, range=f"{h}!A2:B5"
                ).execute().get('values', [])
                msg += f"• *{h}*: {len(rows)} filas\n"
            except Exception as e2:
                msg += f"• *{h}*: error — {e2}\n"
        await update.message.reply_text(msg, parse_mode='Markdown')
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

async def cmd_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        f"Su Chat ID es: `{update.effective_chat.id}`",
        parse_mode='Markdown'
    )

async def cmd_resumen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Generando resumen...")
    await daily_summary(context.bot)
    await update.message.reply_text("✅ Resumen enviado a su email.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # ── SEGURIDAD: solo responde al chat autorizado ──
    allowed_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    if allowed_id and str(update.effective_chat.id) != str(allowed_id):
        logger.warning(f"Acceso denegado a chat_id: {update.effective_chat.id}")
        return

    user_msg = update.message.text
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action="typing"
    )

    keywords = ['agenda','cita','evento','reunión','juicio','vista',
                'mañana','semana','hoy','pendiente','calendario','tengo',
                'mueve','cambia','modifica','cancela','tarea','tareas','recordatorio']
    calendar_ctx = ""
    if any(kw in user_msg.lower() for kw in keywords):
        events = get_events(days=7)
        if events:
            calendar_ctx = format_events(events)

    try:
        raw = ask_claude(user_msg, calendar_ctx)

        # Extraer todos los bloques JSON (soporte multi-acción)
        # Buscar tanto JSONs sueltos como dentro de bloques ```json ... ```
        # Extraer JSONs: primero dentro de bloques ```json```, luego sueltos
        json_matches = re.findall(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', raw, re.DOTALL)

        # Texto limpio = raw sin bloques JSON ni ```
        clean_text = re.sub(r'```json.*?```', '', raw, flags=re.DOTALL)
        clean_text = re.sub(r'```.*?```', '', clean_text, flags=re.DOTALL)
        clean_text = re.sub(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', clean_text, flags=re.DOTALL)
        clean_text = clean_text.strip()

        if not json_matches:
            await update.message.reply_text(raw)
            return

        actions_list = []
        for match in json_matches:
            try:
                actions_list.append(json.loads(match))
            except:
                pass

        if not actions_list:
            await update.message.reply_text(raw)
            return

        acciones_completadas = []

        for data in actions_list:
            action = data.get('action', 'none')

            if action == 'create_event':
                ev = create_event(
                    data['summary'], data['date'], data['time'],
                    data.get('duration_hours', 1), data.get('description', '')
                )
                if ev:
                    await update.message.reply_text(
                        f"✅ Evento creado:\n📌 *{data['summary']}*\n📅 {data['date']} a las {data['time']}",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ No se pudo crear el evento.")

            elif action == 'update_event':
                event = find_event_by_name(data.get('event_name', ''))
                if event:
                    ev = update_event(
                        event['id'],
                        date_str=data.get('date'),
                        time_str=data.get('time'),
                        duration_hours=data.get('duration_hours', 1)
                    )
                    if ev:
                        await update.message.reply_text(
                            f"✅ Evento actualizado:\n📌 *{event['summary']}*\n📅 {data.get('date')} a las {data.get('time')}",
                            parse_mode='Markdown'
                        )
                    else:
                        await update.message.reply_text("❌ No se pudo actualizar el evento.")
                else:
                    await update.message.reply_text("❌ No encontré ese evento en los próximos 30 días.")

            elif action == 'delete_event':
                event = find_event_by_name(data.get('event_name', ''))
                if event:
                    ok = delete_event(event['id'])
                    if ok:
                        await update.message.reply_text(
                            f"✅ Evento eliminado: *{event['summary']}*",
                            parse_mode='Markdown'
                        )
                    else:
                        await update.message.reply_text("❌ No se pudo eliminar el evento.")
                else:
                    await update.message.reply_text("❌ No encontré ese evento en los próximos 30 días.")

            elif action == 'query_calendar':
                days   = data.get('days', 7)
                events = get_events(days=days)
                if not events:
                    await update.message.reply_text(f"📅 No hay eventos en los próximos {days} días.")
                else:
                    msg = f"📅 *Agenda — próximos {days} días:*\n\n{format_events(events)}"
                    await update.message.reply_text(msg, parse_mode='Markdown')

            elif action == 'query_tasks':
                tasks = get_tasks()
                msg   = f"📋 *Tareas pendientes:*\n\n{format_tasks(tasks)}"
                await update.message.reply_text(msg, parse_mode='Markdown')

            elif action == 'create_task':
                t = create_task(data['title'], data.get('notes',''), data.get('due_date'))
                if t:
                    due_txt = f" (vence {data['due_date']})" if data.get('due_date') else ''
                    await update.message.reply_text(
                        f"✅ Tarea creada: *{data['title']}*{due_txt}",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ No se pudo crear la tarea.")

            elif action == 'delete_task':
                task = find_task_by_name(data.get('task_name',''))
                if task:
                    ok = delete_task(task['id'])
                    if ok:
                        await update.message.reply_text(
                            f"✅ Tarea eliminada: *{task['title']}*",
                            parse_mode='Markdown'
                        )
                    else:
                        await update.message.reply_text("❌ No se pudo eliminar la tarea.")
                else:
                    await update.message.reply_text("❌ No encontré esa tarea.")

            elif action == 'complete_task':
                task = find_task_by_name(data.get('task_name',''))
                if task:
                    ok = complete_task(task['id'])
                    if ok:
                        await update.message.reply_text(
                            f"✅ Tarea completada: *{task['title']}*",
                            parse_mode='Markdown'
                        )
                    else:
                        await update.message.reply_text("❌ No se pudo completar la tarea.")
                else:
                    await update.message.reply_text("❌ No encontré esa tarea.")

            elif action == 'send_email':
                ok = send_email(data['to'], data['subject'], data['body'])
                if ok:
                    await update.message.reply_text(
                        f"✅ Email enviado a `{data['to']}`", parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ Error al enviar el email.")

            elif action == 'create_invoice':
                base = data['base_imponible']
                iva  = data.get('iva', 21)
                ret  = data.get('retencion', 0)
                if not data.get('es_base', True):
                    factor = 1 + iva/100 - ret/100
                    base   = round(base / factor, 2)
                pdf_bytes, total = generar_factura(
                    num_factura=data['num_factura'],
                    cliente_nombre=data['cliente_nombre'],
                    cliente_nif=data['cliente_nif'],
                    cliente_domicilio=data['cliente_domicilio'],
                    concepto=data['concepto'],
                    base_imponible=base,
                    iva=iva,
                    retencion=ret
                )
                num_safe     = data['num_factura'].replace('/', '-').replace(' ', '')
                cliente_safe = data['cliente_nombre'].replace(' ', '_').encode('ascii', 'ignore').decode()
                pdf_name     = f"Factura_{num_safe}_{cliente_safe}.pdf"
                body_email   = f"Adjunto encontrará la factura núm. {data['num_factura']} por importe de {total:.2f} €."
                ok = send_email_with_pdf(
                    data['cliente_email'],
                    f"Factura {data['num_factura']} — AP Estudio Jurídico",
                    body_email, pdf_bytes, pdf_name
                )
                if ok:
                    await update.message.reply_text(
                        f"✅ Factura enviada a `{data['cliente_email']}`\n"
                        f"📄 *{data['num_factura']}* — Total: *{total:.2f} €*",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ Error al enviar la factura.")


            elif action == 'query_cliente':
                c = get_cliente(data.get('nombre', ''))
                if c:
                    def campo(etiqueta, valor):
                        return f"{etiqueta}: {valor}\n" if valor else f"{etiqueta}:\n"
                    msg = (campo("Cliente",   c['nombre']) +
                           campo("NIF/CIF",   c['nif']) +
                           campo("Dirección", c['direccion']) +
                           campo("CP",        c['cp']) +
                           campo("Población", c['poblacion']) +
                           campo("Provincia", c['provincia']) +
                           campo("País",      c['pais']) +
                           campo("Email",     c['email']) +
                           campo("Teléfono",  c['telefono']))
                    await update.message.reply_text(msg.strip())
                else:
                    await update.message.reply_text(f"❌ No encontré el cliente '{data.get('nombre')}'.")

            elif action == 'query_clientes':
                clientes = get_todos_clientes()
                if not clientes:
                    await update.message.reply_text("No hay clientes registrados.")
                else:
                    msg = f"Clientes del despacho ({len(clientes)}):\n\n"
                    for c in clientes:
                        msg += f"• {c['nombre']} — {c['tipo']} | {c['email']}\n"
                    await update.message.reply_text(msg)

            elif action == 'add_cliente':
                tz = pytz.timezone(TIMEZONE)
                fecha_alta = datetime.now(tz).strftime('%Y-%m-%d')
                nuevo_id = siguiente_id_cliente()
                fila = [
                    str(nuevo_id), data.get('nombre',''), data.get('apellidos',''),
                    data.get('nif',''), data.get('email',''), data.get('telefono',''),
                    data.get('direccion',''), data.get('poblacion',''), data.get('cp',''),
                    data.get('tipo','Particular'), fecha_alta, data.get('notas','')
                ]
                ok = sheets_append('Clientes', fila)
                if ok:
                    nombre = f"{data.get('nombre','')} {data.get('apellidos','')}".strip()
                    await update.message.reply_text(f"✅ Cliente añadido: {nombre} (ID: {nuevo_id})")
                else:
                    await update.message.reply_text("❌ Error al añadir el cliente.")

            elif action == 'query_casos':
                casos = get_casos_cliente(data.get('cliente'))
                if not casos:
                    await update.message.reply_text("No hay casos registrados.")
                else:
                    nombre_filtro = data.get('cliente', '')
                    titulo = f"Casos de {nombre_filtro}:" if nombre_filtro else f"Todos los casos ({len(casos)}):"
                    msg = titulo + "\n\n"
                    for c in casos:
                        msg += (f"• {c['cliente']} — {c['materia']}\n"
                                f"  {c['autos']} | {c['estado']}\n"
                                f"  Próx: {c['proxima_actuacion']} ({c['fecha_actuacion']})\n\n")
                    await update.message.reply_text(msg)

            elif action == 'add_caso':
                rows = sheets_read("Casos!A2:A200")
                ids = [int(r[0]) for r in rows if r and str(r[0]).isdigit()]
                nuevo_id = max(ids) + 1 if ids else 1
                fila = [
                    str(nuevo_id), data.get('id_cliente',''), data.get('cliente',''),
                    data.get('tipo',''), data.get('materia',''), data.get('descripcion',''),
                    data.get('juzgado',''), data.get('autos',''), data.get('estado','Activo'),
                    data.get('fecha_apertura',''), data.get('proxima_actuacion',''),
                    data.get('fecha_actuacion',''), data.get('honorarios','0'), data.get('cobrado','0')
                ]
                ok = sheets_append('Casos', fila)
                if ok:
                    await update.message.reply_text(f"✅ Caso añadido: {data.get('materia','')} — {data.get('cliente','')}")
                else:
                    await update.message.reply_text("❌ Error al añadir el caso.")

            elif action == 'update_caso_estado':
                rows = sheets_read("Casos!A2:N200")
                autos_buscar = normalizar(data.get('autos', ''))
                encontrado = False
                for i, row in enumerate(rows, 2):
                    if len(row) > 7 and autos_buscar in normalizar(row[7]):
                        sheets_update_cell(f"Casos!I{i}", data.get('estado', row[8]))
                        if data.get('proxima_actuacion'):
                            sheets_update_cell(f"Casos!K{i}", data['proxima_actuacion'])
                        if data.get('fecha_actuacion'):
                            sheets_update_cell(f"Casos!L{i}", data['fecha_actuacion'])
                        encontrado = True
                        break
                if encontrado:
                    await update.message.reply_text(f"✅ Caso {data.get('autos','')} actualizado.")
                else:
                    await update.message.reply_text(f"❌ No encontré el caso '{data.get('autos')}'.")

            elif action == 'query_facturas':
                estado = data.get('estado')
                facturas = get_facturas(estado=estado)
                if not facturas:
                    est_txt = f" {estado}" if estado else ""
                    await update.message.reply_text(f"No hay facturas{est_txt}.")
                else:
                    titulo = f"Facturas{' ' + estado if estado else ''} ({len(facturas)}):"
                    msg = titulo + "\n\n"
                    total_pend = 0
                    for f in facturas:
                        msg += f"• Fac.{f['num']} — {f['cliente']} | {f['total']} € | {f['estado']}\n"
                        try:
                            total_pend += float(str(f['total']).replace(',','.').replace('€','').strip())
                        except:
                            pass
                    if estado and 'pendiente' in estado.lower():
                        msg += f"\nTotal pendiente: {total_pend:.2f} €"
                    await update.message.reply_text(msg)

            elif action == 'cobrar_factura':
                rows = sheets_read("Facturas!A2:N200")
                num_buscar = str(data.get('num_factura', '')).strip()
                encontrado = False
                tz = pytz.timezone(TIMEZONE)
                fecha_cobro = data.get('fecha_cobro', datetime.now(tz).strftime('%Y-%m-%d'))
                for i, row in enumerate(rows, 2):
                    # Buscar por columna C (num factura) o columna A (id)
                    num_fila = str(row[2]).strip() if len(row) > 2 else ''
                    if row and (num_fila == num_buscar or str(row[0]).strip() == num_buscar):
                        sheets_update_cell(f"Facturas!M{i}", "Cobrada")
                        encontrado = True
                        break
                if encontrado:
                    await update.message.reply_text(f"✅ Factura {data.get('num_factura','')} marcada como cobrada.")
                else:
                    await update.message.reply_text(f"❌ No encontré la factura '{data.get('num_factura')}'.")

            elif action == 'create_invoice_bd':
                c = get_cliente(data.get('cliente', ''))
                if not c:
                    await update.message.reply_text(f"❌ No encontré el cliente '{data.get('cliente')}' en la base de datos.")
                else:
                    nombre_completo = f"{c['nombre']} {c['apellidos']}".strip()
                    domicilio = f"{c['direccion']}, {c['cp']} {c['poblacion']}{', ' + c.get('provincia','') if c.get('provincia') else ''}"
                    num_factura = siguiente_num_factura()
                    base = data['base_imponible']
                    iva  = data.get('iva', 21)
                    ret  = data.get('retencion', 0)
                    if not data.get('es_base', True):
                        factor = 1 + iva/100 - ret/100
                        base   = round(base / factor, 2)
                    pdf_bytes, total = generar_factura(
                        num_factura=num_factura, cliente_nombre=nombre_completo,
                        cliente_nif=c['nif'], cliente_domicilio=domicilio,
                        concepto=data['concepto'], base_imponible=base, iva=iva, retencion=ret
                    )
                    tz = pytz.timezone(TIMEZONE)
                    fecha = datetime.now(tz).strftime('%Y-%m-%d')
                    iva_amount = round(base * iva / 100, 2)
                    ret_amount = round(base * ret / 100, 2)
                    # Insertar factura antes de la fila de totales
                    insertar_factura_en_sheets(
                        fecha=fecha,
                        num_factura=num_factura,
                        cliente=nombre_completo,
                        nif=c['nif'],
                        base=base,
                        iva_amount=iva_amount,
                        ret_amount=ret_amount,
                        total=total
                    )
                    num_safe    = num_factura.replace('/', '-').replace(' ', '')
                    pdf_name    = f"Factura_{num_safe}_{nombre_completo.replace(' ','_')}.pdf"
                    email_dest  = c['email'] if c['email'] else GMAIL_USER
                    body_email  = f"Adjunto la factura núm. {num_factura} por importe de {total:.2f} €."
                    ok = send_email_with_pdf(email_dest, f"Factura {num_factura} — AP Estudio Jurídico", body_email, pdf_bytes, pdf_name)
                    if ok:
                        await update.message.reply_text(
                            f"✅ Factura {num_factura} creada y enviada a {email_dest}\n"
                            f"Cliente: {nombre_completo}\nTotal: {total:.2f} €"
                        )
                    else:
                        await update.message.reply_text(f"✅ Factura {num_factura} creada (error al enviar email)\nTotal: {total:.2f} €")

            else:
                response_text = data.get('response', raw)
                # Si la respuesta es un JSON, procesarlo de nuevo
                if response_text.strip().startswith('{'):
                    try:
                        inner = json.loads(response_text.strip())
                        inner_action = inner.get('action', 'none')
                        if inner_action != 'none':
                            data = inner
                            action = inner_action
                            # Re-dispatch: añadir a la lista para procesar
                            actions_list.append(inner)
                            continue
                        else:
                            response_text = inner.get('response', response_text)
                    except Exception:
                        pass
                response_text = re.sub(r'\*\*(.*?)\*\*', r'*\1*', response_text)
                try:
                    await update.message.reply_text(response_text, parse_mode='Markdown')
                except Exception:
                    await update.message.reply_text(response_text)


        # Enviar resumen final de todas las acciones
        if acciones_completadas:
            resumen = "\n".join(acciones_completadas)
            try:
                await update.message.reply_text(resumen, parse_mode='HTML')
            except Exception:
                await update.message.reply_text(resumen)
        elif clean_text:
            await update.message.reply_text(clean_text)

    except json.JSONDecodeError:
        await update.message.reply_text(raw)
    except Exception as e:
        logger.error(f"Error en handle_message: {e}")
        await update.message.reply_text("❌ Ha ocurrido un error. Inténtelo de nuevo.")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("id",      cmd_id))
    app.add_handler(CommandHandler("resumen", cmd_resumen))
    app.add_handler(CommandHandler("bbdd",    cmd_bbdd))
    app.add_handler(CommandHandler("initbbdd", cmd_initbbdd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    loop = asyncio.get_event_loop()

    def run_daily():
        loop.create_task(daily_summary(app.bot))

    def run_weekly():
        loop.create_task(weekly_summary(app.bot))

    scheduler.add_job(run_daily,  'cron', hour=7, minute=0, day_of_week='mon-fri')
    scheduler.add_job(run_weekly, 'cron', hour=9, minute=0, day_of_week='sat')

    scheduler.start()
    logger.info("✅ Bot Secretaria iniciado.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
