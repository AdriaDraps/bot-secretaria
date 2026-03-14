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
LOGO_B64          = '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'
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

# ─────────────────────────────────────────────
# GENERACIÓN DE FACTURAS PDF
# ─────────────────────────────────────────────
def generar_factura(num_factura, cliente_nombre, cliente_nif, cliente_domicilio,
                    concepto, base_imponible, iva=21, retencion=0):
    """Genera una factura en PDF y devuelve (bytes, total)."""
    import io
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    story = []
    estilos = {
        'normal':  ParagraphStyle('normal',  fontName='Helvetica',      fontSize=10, leading=14),
        'negrita': ParagraphStyle('negrita', fontName='Helvetica-Bold',  fontSize=10, leading=14),
        'titulo':  ParagraphStyle('titulo',  fontName='Helvetica-Bold',  fontSize=16, leading=20, spaceAfter=6),
        'derecha': ParagraphStyle('derecha', fontName='Helvetica',       fontSize=10, leading=14, alignment=TA_RIGHT),
        'der_bold':ParagraphStyle('der_bold',fontName='Helvetica-Bold',  fontSize=10, leading=14, alignment=TA_RIGHT),
        'pequeño': ParagraphStyle('pequeño', fontName='Helvetica',       fontSize=8,  leading=12, textColor=colors.grey),
        'center':  ParagraphStyle('center',  fontName='Helvetica',       fontSize=10, leading=14, alignment=TA_CENTER),
    }
    normal   = estilos['normal']
    negrita  = estilos['negrita']
    derecha  = estilos['derecha']
    der_bold = estilos['der_bold']

    # Cabecera
    story.append(Paragraph('<b>AP ESTUDIO JURÍDICO</b>', estilos['titulo']))
    story.append(Paragraph('Adrià Paños Ruiz — NIF: 47182626N', normal))
    story.append(Paragraph('Carrer Comte Ramon Berenguer 1-3, esc. B, 2º 1ª', normal))
    story.append(Paragraph('08204 Sabadell (Barcelona)', normal))
    story.append(Paragraph('Tel: 603690659', normal))
    story.append(Spacer(1, 0.4*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black))
    story.append(Spacer(1, 0.4*cm))

    # Número de factura y fecha
    import pytz as _pytz
    from datetime import datetime as _dt
    tz = _pytz.timezone(TIMEZONE)
    fecha_emision = _dt.now(tz).strftime('%d/%m/%Y')
    story.append(Paragraph(f'<b>FACTURA N.º {num_factura}</b>', negrita))
    story.append(Paragraph(f'Fecha de emisión: {fecha_emision}', normal))
    story.append(Spacer(1, 0.4*cm))

    # Datos cliente
    story.append(Paragraph('<b>DATOS DEL CLIENTE</b>', negrita))
    story.append(Paragraph(cliente_nombre, normal))
    story.append(Paragraph(f'NIF/CIF: {cliente_nif}', normal))
    story.append(Paragraph(cliente_domicilio, normal))
    story.append(Spacer(1, 0.4*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Spacer(1, 0.4*cm))

    # Tabla de conceptos
    tabla_data = [
        [Paragraph('<b>CONCEPTO</b>', negrita), Paragraph('<b>IMPORTE</b>', der_bold)],
        [Paragraph(concepto, normal), Paragraph(f'{base_imponible:.2f} €', derecha)],
    ]
    tabla = Table(tabla_data, colWidths=[13*cm, 4*cm])
    tabla.setStyle(TableStyle([
        ('GRID',        (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND',  (0,0), (-1,0),  colors.lightgrey),
        ('VALIGN',      (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING',  (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0),(-1,-1), 6),
    ]))
    story.append(tabla)
    story.append(Spacer(1, 0.3*cm))

    # Totales
    iva_importe   = round(base_imponible * iva / 100, 2)
    retencion_imp = round(base_imponible * retencion / 100, 2)
    total         = round(base_imponible + iva_importe - retencion_imp, 2)

    totales_data = [
        [Paragraph('<b>Base imponible</b>', negrita),        Paragraph(f'{base_imponible:.2f} €', der_bold)],
        [Paragraph(f'IVA ({iva}%)', normal),                 Paragraph(f'{iva_importe:.2f} €', derecha)],
    ]
    if retencion > 0:
        totales_data.append([
            Paragraph(f'Retención IRPF ({retencion}%)', normal),
            Paragraph(f'-{retencion_imp:.2f} €', derecha)
        ])
    totales_data.append([
        Paragraph('<b>TOTAL HONORARIOS</b>', negrita),
        Paragraph(f'<b>{total:.2f} €</b>', der_bold)
    ])
    tabla_tot = Table(totales_data, colWidths=[13*cm, 4*cm])
    tabla_tot.setStyle(TableStyle([
        ('LINEABOVE', (0,-1), (-1,-1), 1, colors.black),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(tabla_tot)
    story.append(Spacer(1, 1*cm))

    # Pie
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph('Los honorarios devengados deberán abonarse en el siguiente número de cuenta:', estilos['pequeño']))
    story.append(Paragraph('IBAN: ES10 1563 2626 3132 6989 1055', estilos['pequeño']))

    doc.build(story)
    return buffer.getvalue(), total


def encode_subject(subject):
    """Codifica el asunto del email en base64 UTF-8 para evitar caracteres corruptos."""
    import base64 as _b64
    encoded = _b64.b64encode(subject.encode('utf-8')).decode('ascii')
    return f'=?utf-8?b?{encoded}?='

def send_email(to_addr, subject, body_text):
    try:
        service = get_gmail_service()
        if not service:
            logger.error("No se pudo conectar con Gmail API")
            return False
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart as _MM
        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')
        logo_tag = '<img src="data:image/png;base64,' + LOGO_B64 + '" alt="AP Estudio Juridico" style="height:80px;width:auto;"/>'
        html_body = (
            '<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">'
            + '<div style="text-align:center;margin-bottom:24px;">' + logo_tag + '</div>'
            + '<div style="border-top:3px solid #b8960c;padding-top:20px;">'
            + formatted + '</div>'
            + '<div style="margin-top:30px;padding-top:15px;border-top:1px solid #ddd;font-size:0.85em;color:#666;">'
            + '<em>Secretar\u00eda \u2014 AP Estudio Jur\u00eddico</em></div></body></html>')
        msg = MIMEText(html_body, 'html', 'utf-8')
        msg['Subject'] = subject
        msg['From']    = GMAIL_USER
        msg['To']      = to_addr
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()
        logger.info(f"Email enviado a {to_addr}")
        return True
    except Exception as e:
        logger.error(f"Error enviando email via Gmail API: {e}")
        return False


def send_email_with_pdf(to_addr, subject, body_text, pdf_bytes, pdf_filename):
    try:
        service = get_gmail_service()
        if not service:
            return False
        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')
        logo_tag = '<img src="data:image/png;base64,' + LOGO_B64 + '" alt="AP Estudio Juridico" style="height:80px;width:auto;"/>'
        html_body = (
            '<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">'
            + '<div style="text-align:center;margin-bottom:24px;">' + logo_tag + '</div>'
            + '<div>' + formatted + '</div>'
            + '<hr style="border:none;border-top:1px solid #ddd;margin:24px 0;">'
            + '<div style="font-size:0.85em;color:#666;text-align:center;"><em>Secretar\u00eda \u2014 AP Estudio Jur\u00eddico</em></div>'
            + '</body></html>')
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders as _enc
        msg2 = MIMEMultipart()
        msg2['Subject'] = subject
        msg2['From']    = GMAIL_USER
        msg2['To']      = to_addr
        msg2.attach(MIMEText(html_body, 'html', 'utf-8'))
        part = MIMEBase('application', 'pdf')
        part.set_payload(pdf_bytes)
        _enc.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
        part.add_header('Content-Type', 'application/pdf', name=pdf_filename)
        msg2.attach(part)
        raw = base64.urlsafe_b64encode(msg2.as_bytes()).decode()
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
                body_email   = "Estimado/a cliente,\n\nLe adjunto la factura por los servicios prestados por este despacho.\n\nSaludos cordiales,\nAP Estudio Jurídico"
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
                    email_dest  = c['email'].strip() if c.get('email','').strip() else None
                    body_email  = "Estimado/a cliente,\n\nLe adjunto la factura por los servicios prestados por este despacho.\n\nSaludos cordiales,\nAP Estudio Jurídico"
                    if email_dest:
                        ok = send_email_with_pdf(email_dest, f"Factura {num_factura} — AP Estudio Jurídico", body_email, pdf_bytes, pdf_name)
                        if ok:
                            acciones_completadas.append(f"✅ Factura {num_factura} creada y enviada a {email_dest}\nCliente: {nombre_completo} | Total: {total:.2f} €")
                        else:
                            acciones_completadas.append(f"✅ Factura {num_factura} creada (error al enviar email)\nTotal: {total:.2f} €")
                    else:
                        acciones_completadas.append(f"✅ Factura {num_factura} creada — sin email en BD para {nombre_completo}\nTotal: {total:.2f} €")

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
