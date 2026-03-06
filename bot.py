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
TIMEZONE          = 'Europe/Madrid'
SCOPES            = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.send'
]

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

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
SYSTEM_PROMPT = """Eres la secretaria virtual del despacho de abogados APE Estudio Jurídico.
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
- Si falta información necesaria, pídela con action:none
- Para mover un evento, usa action:update_event con el nombre del evento y la nueva hora/fecha
- Para eliminar o cancelar un evento, usa action:delete_event con el nombre del evento
- Para emails al procurador u otros contactos del despacho, redacta el cuerpo de forma formal
- Para facturas, si falta num_factura, cliente_nif, cliente_domicilio o concepto, pídelos con action:none
"""

def ask_claude(user_msg, calendar_context=""):
    tz    = pytz.timezone(TIMEZONE)
    today = datetime.now(tz).strftime('%d/%m/%Y, %A')
    system = SYSTEM_PROMPT.replace('{today}', today)

    content = user_msg
    if calendar_context:
        content = f"Contexto actual del calendario:\n{calendar_context}\n\nMensaje del abogado: {user_msg}"

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        system=system,
        messages=[{"role": "user", "content": content}]
    )
    return resp.content[0].text.strip()

# ─────────────────────────────────────────────
# RESUMEN DIARIO (Lunes a Viernes 7:00 AM)
# ─────────────────────────────────────────────
async def daily_summary(bot):
    chat_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    tz      = pytz.timezone(TIMEZONE)
    today   = datetime.now(tz)

    events_today = get_events(days=1)
    events_week  = get_events(days=7)

    ctx = (
        f"Hoy es {today.strftime('%A %d/%m/%Y')}.\n\n"
        f"Eventos de hoy:\n{format_events(events_today)}\n\n"
        f"Resto de la semana:\n{format_events(events_week)}"
    )

    prompt = (
        f"{ctx}\n\n"
        "Genera un resumen matutino MUY esquemático para el abogado Adrià. "
        "Formato: primero los eventos de hoy con hora y título, luego los de la semana. "
        "Sin saludos, sin frases motivadoras. Solo datos. Máximo 80 palabras."
    )

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}]
    )
    summary = resp.content[0].text.strip()

    send_email(
        GMAIL_USER,
        f"📅 Agenda {today.strftime('%d/%m/%Y')}",
        summary
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
    summary = resp.content[0].text.strip()

    send_email(GMAIL_USER, subject, summary)
    logger.info("Resumen semanal enviado.")

# ─────────────────────────────────────────────
# HANDLERS TELEGRAM
# ─────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
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
    user_msg = update.message.text
    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action="typing"
    )

    keywords = ['agenda','cita','evento','reunión','juicio','vista',
                'mañana','semana','hoy','pendiente','calendario','tengo',
                'mueve','cambia','modifica','cancela']
    calendar_ctx = ""
    if any(kw in user_msg.lower() for kw in keywords):
        events = get_events(days=7)
        if events:
            calendar_ctx = format_events(events)

    try:
        raw = ask_claude(user_msg, calendar_ctx)

        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        if not json_match:
            await update.message.reply_text(raw)
            return

        data   = json.loads(json_match.group())
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
            pdf_name    = f"Factura_{num_safe}_{data['cliente_nombre'].replace(' ','_')}.pdf"
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
            await update.message.reply_text(data.get('response', raw))

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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    scheduler.add_job(
        lambda: asyncio.ensure_future(daily_summary(app.bot)),
        'cron', hour=7, minute=0, day_of_week='mon-fri'
    )
    scheduler.add_job(
        lambda: asyncio.ensure_future(weekly_summary(app.bot)),
        'cron', hour=9, minute=0, day_of_week='sat'
    )

    scheduler.start()
    logger.info("✅ Bot Secretaria iniciado.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
