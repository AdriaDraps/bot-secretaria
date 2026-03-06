import os
import logging
import asyncio
import json
import re
import base64
import smtplib
import pytz
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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
GMAIL_PASSWORD    = os.environ.get('GMAIL_PASSWORD')
TELEGRAM_CHAT_ID  = os.environ.get('TELEGRAM_CHAT_ID', '')
GOOGLE_TOKEN_B64  = os.environ.get('GOOGLE_TOKEN_B64')
TIMEZONE          = 'Europe/Madrid'
SCOPES            = ['https://www.googleapis.com/auth/calendar']

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

# ─────────────────────────────────────────────
# GOOGLE CALENDAR
# ─────────────────────────────────────────────
def get_calendar_service():
    if not GOOGLE_TOKEN_B64:
        return None
    try:
        token_data = json.loads(base64.b64decode(GOOGLE_TOKEN_B64).decode())
        creds = Credentials.from_authorized_user_info(token_data, SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return build('calendar', 'v3', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando con Google Calendar: {e}")
        return None

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

def update_event(event_id, summary=None, date_str=None, time_str=None, duration_hours=None):
    service = get_calendar_service()
    if not service:
        return None
    try:
        event = service.events().get(calendarId='primary', eventId=event_id).execute()
        if summary:
            event['summary'] = summary
        if date_str and time_str:
            tz = pytz.timezone(TIMEZONE)
            start_dt = tz.localize(datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M"))
            hours = duration_hours or 1
            end_dt = start_dt + timedelta(hours=hours)
            event['start'] = {'dateTime': start_dt.isoformat(), 'timeZone': TIMEZONE}
            event['end']   = {'dateTime': end_dt.isoformat(),   'timeZone': TIMEZONE}
        return service.events().update(calendarId='primary', eventId=event_id, body=event).execute()
    except Exception as e:
        logger.error(f"Error actualizando evento: {e}")
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

def format_events(events):
    if not events:
        return "No hay eventos."
    tz  = pytz.timezone(TIMEZONE)
    txt = ""
    for ev in events:
        start = ev['start'].get('dateTime', ev['start'].get('date'))
        try:
            dt  = datetime.fromisoformat(start.replace('Z', '+00:00')).astimezone(tz)
            fmt = dt.strftime('%d/%m a las %H:%M')
        except Exception:
            fmt = start
        txt += f"• *{ev['summary']}* — {fmt}\n"
    return txt

# ─────────────────────────────────────────────
# EMAIL
# ─────────────────────────────────────────────
def send_email(to_addr, subject, body_text):
    try:
        msg             = MIMEMultipart('alternative')
        msg['Subject']  = subject
        msg['From']     = GMAIL_USER
        msg['To']       = to_addr
        html = f"""
        <html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">
          <div style="border-top:3px solid #b8960c;padding-top:20px;">
            {body_text.replace(chr(10),'<br>')}
          </div>
          <div style="margin-top:30px;padding-top:15px;border-top:1px solid #ddd;
                      font-size:0.85em;color:#666;">
            <em>Secretaria Virtual — APE Estudio Jurídico</em>
          </div>
        </body></html>"""
        msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
        msg.attach(MIMEText(html,      'html',  'utf-8'))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as srv:
            srv.login(GMAIL_USER, GMAIL_PASSWORD)
            srv.sendmail(GMAIL_USER, to_addr, msg.as_string())
        return True
    except Exception as e:
        logger.error(f"Error enviando email: {e}")
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

Para modificar evento (mover a otra hora/fecha):
{{"action":"update_event","event_name":"nombre del evento","date":"YYYY-MM-DD","time":"HH:MM","duration_hours":1}}

Para consultar agenda:
{{"action":"query_calendar","days":7}}

Para enviar email:
{{"action":"send_email","to":"email@ejemplo.com","subject":"Asunto","body":"Cuerpo del email"}}

Para cualquier otra respuesta conversacional:
{{"action":"none","response":"tu respuesta aquí"}}

Reglas:
- Responde siempre en español
- Sé concisa y profesional
- Si la fecha es relativa (mañana, el lunes...) calcúlala a partir de hoy: {today}
- Si falta información necesaria, pídela con action:none
- Para mover un evento, usa action:update_event con el nombre del evento y la nueva hora/fecha
- Para emails al procurador u otros contactos del despacho, redacta el cuerpo de forma formal
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
    if not chat_id:
        logger.warning("TELEGRAM_CHAT_ID no configurado — no se puede enviar resumen.")
        return

    tz    = pytz.timezone(TIMEZONE)
    today = datetime.now(tz)

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
        f"📅 Agenda del día — {today.strftime('%d/%m/%Y')}",
        f"Buenos días,\n\n{summary}\n\nAPE Estudio Jurídico — Secretaria Virtual"
    )
    logger.info("Resumen diario enviado.")

# ─────────────────────────────────────────────
# RESUMEN SEMANAL (Sábados 9:00 AM)
# ─────────────────────────────────────────────
async def weekly_summary(bot):
    chat_id = os.environ.get('TELEGRAM_CHAT_ID', TELEGRAM_CHAT_ID)
    if not chat_id:
        return

    tz    = pytz.timezone(TIMEZONE)
    today = datetime.now(tz)
    events = get_events(days=7)

    prompt = (
        f"Hoy es sábado {today.strftime('%d/%m/%Y')}.\n\n"
        f"Agenda de la próxima semana:\n{format_events(events)}\n\n"
        "Genera un resumen de la semana que viene para el abogado Adrià. "
        "Salúdale, lista los eventos importantes y deséale buena semana. "
        "Máximo 200 palabras. Sin JSON, solo texto."
    )

    resp = claude_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}]
    )
    summary = resp.content[0].text.strip()

    try:
        await bot.send_message(
            chat_id=chat_id,
            text=f"📅 *Resumen semana próxima*\n\n{summary}",
            parse_mode='Markdown'
        )
    except Exception as e:
        logger.error(f"Error enviando resumen semanal: {e}")

    logger.info("Resumen semanal enviado.")

# ─────────────────────────────────────────────
# HANDLERS TELEGRAM
# ─────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = str(update.effective_chat.id)
    os.environ['TELEGRAM_CHAT_ID'] = chat_id
    logger.info(f"Chat ID registrado: {chat_id}")

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
        f"Su Chat ID es: `{update.effective_chat.id}`\n"
        "Guárdelo para configurar la variable TELEGRAM_CHAT_ID en Railway.",
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
                    f"✅ Evento creado:\n"
                    f"📌 *{data['summary']}*\n"
                    f"📅 {data['date']} a las {data['time']}",
                    parse_mode='Markdown'
                )
            else:
                await update.message.reply_text(
                    "❌ No se pudo crear el evento. Compruebe la conexión con Google Calendar."
                )

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
                        f"✅ Evento actualizado:\n"
                        f"📌 *{event['summary']}*\n"
                        f"📅 {data.get('date')} a las {data.get('time')}",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text("❌ No se pudo actualizar el evento.")
            else:
                await update.message.reply_text(
                    f"❌ No encontré ningún evento con ese nombre en los próximos 30 días."
                )

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

    # Resumen diario lunes a viernes a las 7:00
    scheduler.add_job(
        lambda: asyncio.ensure_future(daily_summary(app.bot)),
        'cron', hour=7, minute=0, day_of_week='mon-fri'
    )

    # Resumen semanal los sábados a las 9:00
    scheduler.add_job(
        lambda: asyncio.ensure_future(weekly_summary(app.bot)),
        'cron', hour=9, minute=0, day_of_week='sat'
    )

    scheduler.start()

    logger.info("✅ Bot Secretaria iniciado.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
