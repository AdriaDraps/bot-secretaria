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
        html_body = f"""<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">
          <div>{formatted}</div>
          <div style="margin-top:20px;font-size:0.85em;color:#666;"><em>AP Estudio Jurídico</em></div>
        </body></html>"""
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{pdf_filename}"')
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

def get_cliente(nombre):
    rows = sheets_read("Clientes!A2:L100")
    nombre_lower = nombre.lower()
    for row in rows:
        if len(row) < 2:
            continue
        nombre_completo = f"{row[1]} {row[2]}".lower() if len(row) > 2 else row[1].lower()
        if nombre_lower in nombre_completo or nombre_lower in row[1].lower():
            return {
                'id': row[0] if len(row) > 0 else '',
                'nombre': row[1] if len(row) > 1 else '',
                'apellidos': row[2] if len(row) > 2 else '',
                'nif': row[3] if len(row) > 3 else '',
                'email': row[4] if len(row) > 4 else '',
                'telefono': row[5] if len(row) > 5 else '',
                'direccion': row[6] if len(row) > 6 else '',
                'poblacion': row[7] if len(row) > 7 else '',
                'cp': row[8] if len(row) > 8 else '',
                'tipo': row[9] if len(row) > 9 else '',
                'fecha_alta': row[10] if len(row) > 10 else '',
                'notas': row[11] if len(row) > 11 else '',
            }
    return None

def get_todos_clientes():
    rows = sheets_read("Clientes!A2:L100")
    clientes = []
    for row in rows:
        if len(row) >= 2 and row[0]:
            nombre = f"{row[1]} {row[2]}".strip() if len(row) > 2 else row[1]
            clientes.append({'id': row[0], 'nombre': nombre,
                'nif': row[3] if len(row) > 3 else '',
                'email': row[4] if len(row) > 4 else '',
                'telefono': row[5] if len(row) > 5 else '',
                'tipo': row[9] if len(row) > 9 else ''})
    return clientes

def get_casos_cliente(nombre_cliente=None):
    rows = sheets_read("Casos!A2:N100")
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
    rows = sheets_read("Facturas!A2:K100")
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
    rows = sheets_read("Clientes!A2:A100")
    ids = [int(r[0]) for r in rows if r and r[0].isdigit()]
    return max(ids) + 1 if ids else 1

def siguiente_num_factura():
    rows = sheets_read("Facturas!A2:A100")
    nums = []
    for r in rows:
        if r and r[0]:
            try:
                nums.append(int(r[0].split('/')[0].strip()))
            except:
                pass
    return f"{max(nums) + 1 if nums else 1}/{datetime.now().year}"

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
Ayudas al abogado Adrià con su agenda, emails y tareas administrativas.

Cuando el usuario quiera crear un evento, modificar uno, cancelar uno, consultar agenda o enviar un email,
responde ÚNICAMENTE con un JSON válido en una de estas formas:

Para crear evento:
{{"action":"create_event","summary":"título","date":"YYYY-MM-DD","time":"HH:MM","duration_hours":1,"description":""}}

Para modificar evento:
{{"action":"update_event","event_name":"nombre del evento","date":"YYYY-MM-DD","time":"HH:MM","duration_hours":1}}

Para eliminar evento:
{{"action":"delete_event","event_name":"nombre del evento"}}

Para consultar agenda:
{{"action":"query_calendar","days":7}}

Para enviar email:
{{"action":"send_email","to":"email@ejemplo.com","subject":"Asunto","body":"Cuerpo del email"}}

Para consultar tareas:
{{"action":"query_tasks"}}

Para crear tarea:
{{"action":"create_task","title":"título de la tarea","notes":"","due_date":"YYYY-MM-DD"}}

Para eliminar tarea:
{{"action":"delete_task","task_name":"nombre de la tarea"}}

Para marcar tarea como completada:
{{"action":"complete_task","task_name":"nombre de la tarea"}}

Para crear factura:
{{"action":"create_invoice","num_factura":"14/ 2026","cliente_nombre":"Nombre Cliente","cliente_nif":"12345678A","cliente_domicilio":"Dirección completa","cliente_email":"cliente@email.com","concepto":"Descripción del servicio","base_imponible":500.00,"es_base":true,"iva":21,"retencion":0}}

- "es_base":true si el importe indicado es la base imponible
- "es_base":false si el importe indicado es el total a pagar (calculará la base)
- "retencion" SOLO se añade si el abogado lo indica expresamente. Por defecto siempre es 0, nunca asumas retención

Para cualquier otra respuesta conversacional:
{{"action":"none","response":"tu respuesta aquí"}}

Reglas:
- Responde siempre en español
- Sé concisa y profesional
- Si la fecha es relativa (mañana, el lunes...) calcúlala a partir de hoy: {today}
- Para calcular días de la semana: el número de día es {weekday} (0=lunes, 1=martes, 2=miércoles, 3=jueves, 4=viernes, 5=sábado, 6=domingo)
- "el lunes" significa el próximo lunes. Si hoy es domingo (6), el lunes es mañana. Si hoy es lunes (0), el lunes es el de la semana que viene
- Calcula siempre la fecha exacta YYYY-MM-DD antes de responder y verifica que el día de la semana coincide
- Si falta información necesaria, pídela con action:none
- Para mover un evento, usa action:update_event con el nombre del evento y la nueva hora/fecha
- Para eliminar o cancelar un evento, usa action:delete_event con el nombre del evento
- Para tareas (pendientes, recordatorios, to-do), usa las acciones de tasks
- due_date es opcional en create_task, solo si el abogado indica una fecha límite
- Para emails al procurador u otros contactos del despacho, redacta el cuerpo de forma formal
- Para facturas, si falta num_factura, cliente_nif, cliente_domicilio o concepto, pídelos con action:none

ACCIONES BASE DE DATOS:
Para buscar cliente: {"action":"query_cliente","nombre":"nombre"}
Para listar clientes: {"action":"query_clientes"}
Para añadir cliente: {"action":"add_cliente","nombre":"","apellidos":"","nif":"","email":"","telefono":"","direccion":"","poblacion":"","cp":"","tipo":"Particular","notas":""}
Para consultar casos: {"action":"query_casos","cliente":"nombre opcional"}
Para añadir caso: {"action":"add_caso","id_cliente":"","cliente":"","tipo":"","materia":"","descripcion":"","juzgado":"","autos":"","estado":"","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD","honorarios":"","cobrado":"0"}
Para actualizar estado caso: {"action":"update_caso_estado","autos":"PA 1/2026","estado":"","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD"}
Para consultar facturas: {"action":"query_facturas","estado":"Pendiente"}
Para marcar factura cobrada: {"action":"cobrar_factura","num_factura":"1/2026","fecha_cobro":"YYYY-MM-DD"}
Para crear factura con datos BD: {"action":"create_invoice_bd","cliente":"nombre","concepto":"","base_imponible":500.00,"es_base":true,"iva":21,"retencion":0}

MÚLTIPLES ACCIONES: Si el abogado pide varias cosas en un mensaje, responde con varios JSON seguidos, uno por línea. Ejemplo:
{"action":"create_event","summary":"...","date":"...","time":"...","duration_hours":1,"description":""}
{"action":"create_task","title":"...","notes":""}

Reglas BD:
- Cuando pida datos de cliente, caso o factura usa acciones BD
- Si pide factura y el cliente está en BD usa create_invoice_bd
- Si pide añadir cliente o caso, recoge los datos con action:none antes
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
        json_matches = re.findall(r'```json\s*(\{.*?\})\s*```', raw, re.DOTALL)
        if not json_matches:
            json_matches = re.findall(r'(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})', raw, re.DOTALL)

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
                # Si es total, calcular base
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
                num_safe    = data['num_factura'].replace('/', '-').replace(' ', '')
                cliente_safe = data['cliente_nombre'].replace(' ','_').encode('ascii','ignore').decode()
            pdf_name    = f"Factura_{num_safe}_{cliente_safe}.pdf"
            if not pdf_name or pdf_name == ".pdf":
                pdf_name = f"Factura_{num_safe}.pdf"
                body_email  = f"Adjunto encontrará la factura núm. {data['num_factura']} por importe de {total:.2f} €."
                ok = send_email_with_pdf(data['cliente_email'], f"Factura {data['num_factura']} — AP Estudio Jurídico", body_email, pdf_bytes, pdf_name)
                if ok:
                    await update.message.reply_text(
                        f"✅ Factura enviada a `{data['cliente_email']}`\n"
                        f"📄 *{data['num_factura']}* — Total: *{total:.2f} €*",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ Error al enviar la factura.")

            else:
                response_text = data.get('response', raw)
            response_text = re.sub(r'\*\*(.*?)\*\*', r'*\1*', response_text)
            await update.message.reply_text(response_text, parse_mode='Markdown')

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
