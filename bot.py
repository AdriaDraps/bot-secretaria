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
LOGO_BYTES        = b'\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00\xff\xe2\x01\xd8ICC_PROFILE\x00\x01\x01\x00\x00\x01\xc8\x00\x00\x00\x00\x040\x00\x00mntrRGB XYZ \x07\xe0\x00\x01\x00\x01\x00\x00\x00\x00\x00\x00acsp\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\xf6\xd6\x00\x01\x00\x00\x00\x00\xd3-\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\tdesc\x00\x00\x00\xf0\x00\x00\x00$rXYZ\x00\x00\x01\x14\x00\x00\x00\x14gXYZ\x00\x00\x01(\x00\x00\x00\x14bXYZ\x00\x00\x01<\x00\x00\x00\x14wtpt\x00\x00\x01P\x00\x00\x00\x14rTRC\x00\x00\x01d\x00\x00\x00(gTRC\x00\x00\x01d\x00\x00\x00(bTRC\x00\x00\x01d\x00\x00\x00(cprt\x00\x00\x01\x8c\x00\x00\x00<mluc\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00\x0cenUS\x00\x00\x00\x08\x00\x00\x00\x1c\x00s\x00R\x00G\x00BXYZ \x00\x00\x00\x00\x00\x00o\xa2\x00\x008\xf5\x00\x00\x03\x90XYZ \x00\x00\x00\x00\x00\x00b\x99\x00\x00\xb7\x85\x00\x00\x18\xdaXYZ \x00\x00\x00\x00\x00\x00$\xa0\x00\x00\x0f\x84\x00\x00\xb6\xcfXYZ \x00\x00\x00\x00\x00\x00\xf6\xd6\x00\x01\x00\x00\x00\x00\xd3-para\x00\x00\x00\x00\x00\x04\x00\x00\x00\x02ff\x00\x00\xf2\xa7\x00\x00\rY\x00\x00\x13\xd0\x00\x00\n[\x00\x00\x00\x00\x00\x00\x00\x00mluc\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00\x0cenUS\x00\x00\x00 \x00\x00\x00\x1c\x00G\x00o\x00o\x00g\x00l\x00e\x00 \x00I\x00n\x00c\x00.\x00 \x002\x000\x001\x006\xff\xdb\x00C\x00\x05\x03\x04\x04\x04\x03\x05\x04\x04\x04\x05\x05\x05\x06\x07\x0c\x08\x07\x07\x07\x07\x0f\x0b\x0b\t\x0c\x11\x0f\x12\x12\x11\x0f\x11\x11\x13\x16\x1c\x17\x13\x14\x1a\x15\x11\x11\x18!\x18\x1a\x1d\x1d\x1f\x1f\x1f\x13\x17"$"\x1e$\x1c\x1e\x1f\x1e\xff\xdb\x00C\x01\x05\x05\x05\x07\x06\x07\x0e\x08\x08\x0e\x1e\x14\x11\x14\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\x1e\xff\xc0\x00\x11\x08\x01\xf4\x01\xf4\x03\x01"\x00\x02\x11\x01\x03\x11\x01\xff\xc4\x00\x1d\x00\x01\x00\x03\x00\x03\x01\x01\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x07\x08\t\x04\x05\x06\x03\x01\x02\xff\xc4\x00`\x10\x00\x01\x03\x03\x02\x02\x03\x05\x13\x06\n\x07\x05\x08\x03\x00\x00\x01\x02\x03\x04\x05\x06\x07\x11\x12!\x08\x131AQaq\xb3\t\x14\x15\x16\x17\x18"267Vrtu\x81\x95\xb4\xd2\xd3BRU\x94\xb2\xd1#8Wbs\x82\x91\x93\xa1\xb134v\x83\x84\xa5\xc2$\x92\xa2\xa3\xc1%&DESc\xc4\xd4e\xc3\xf0\xff\xc4\x00\x14\x01\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\xff\xc4\x00\x14\x11\x01\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\xff\xda\x00\x0c\x03\x01\x00\x02\x11\x03\x11\x00?\x00\xb9`\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x07\xf1<\xb1A\x04\x93\xcf#c\x8a6\xab\xde\xf7.\xc8\xd6\xa2n\xaa\xab\xdcC\xfb!\x0e\x9b\x19\xc7\xa4\xed\r\xb8\xd2\xd3\xcd\xc1p\xbf;\xd0\xcatE\xe6\x8cz*\xcc\xef\x17V\x8en\xfd\xc5{@\xee\xf4\xbf_\xb4\xe3Q\xb2\x95\xc6\xb1\xaa\xfa\xc7\xdc:\x87\xce\xc6\xd4R\xacM\x91\xacT\xdd\x1a\xab\xda\xbb.\xfbw\x91{\xc4\xa8dv\x94\xe5\xb58&\xa3X\xb2\xda^5u\xb6\xb1\x92\xc8\xc6\xae\xcb$K\xecdg\xf5\x98\xaeo\xd2kU\xbe\xae\x9a\xe1AO_E3g\xa6\xa9\x89\xb3C#W\x93\xd8\xe4Ek\x93\xc0\xa8\xa8\xa0}\xc0\x00\x00\x00\x00\x00\x00\x00\x00\x00x\xda\xbdV\xd3*;\x8dM\xba\xb7P1\x8aJ\xcaY\x9f\x04\xf0\xd4]!\x8d\xf1H\xd5V\xb9\xaeG96TTTT>\xb0\xeav\x9a\xcd\xfe\x87P\xb1)7\xfc\xcb\xcd:\xff\x00\x93\xca\xc7\xa9}\x0e\xf2<\x8f6\xc82Kveib]nU5\xcc\x82\xa2\x9aF\xf5i,\xaez5\\\xd5v\xfbqm\xbe\xc4[\x96\xf4H\xd6\x0b\x15,\x954\x94v\xab\xf3\x18\x9b\xabm\xb5\x9b\xbfo\x03ek\x15W\xc0\x9b\xaf{p/\xd4Y\xd6\x11/\xfa,\xc7\x1d\x93\xe2\xdc\xe1_\xfa\x8f\xbae\xb8\xaa\xa6\xe9\x93Y\x7f_\x8b\xef\x19\x17u\xb7\\-7\x19\xed\xb7Z\x1a\x9a\x1a\xdawpMOS\x13\xa3\x927w\x9c\xd7"*/\x8c\xe2\x81\xaf\xfe\x9b1_\x84\xd6_\xd7\xa2\xfb\xc3\xd3f+\xf0\x9a\xcb\xfa\xf4_x\xc8\x00\x06\xbdI\x9aa\xd1\xff\x00\xa4\xcb,,\xf8\xd7\x18\x93\xfe\xa3\xe0\xedA\xc0\x9a\xbb;7\xc6\x9a\xbe\x1b\xac\x1fx\xc8\xf8"\x96y\x99\x04\x11\xbeYdr1\x8ccU\\\xe7*\xec\x88\x88\x9d\xaa\xa4\xed\x82\xf4O\xd5\xbc\x9a\x96:\xca\xca\x1a\x1cr\x9eNi\xe8\xac\xea\xc9U;\xfd[\x1a\xe7"\xf8\x1d\xc2\xa0_i5\x1fO#\xff\x00I\x9eb\xcc\xf8\xd7x\x13\xfe\xb3\x8e\xedU\xd2\xf6\xb9\x1a\xedH\xc3\x9a\xab\xdc[\xdd6\xff\x00\xb6U\x18:\x0f_\xd5\xbf\xc3\xe7\xf6\xc6/y\x94\x0fw\xf9\xb9\x0f\x85gB\x1c\xad\xad_:f\xf6Y]\xdcIi\xa5\x8d?\xc3\x88\x0b\x7fM\xa8\xba}R\xe4m>u\x8b\xcc\xe5\xecH\xee\xd09W\xfb\x1cw\xd45\xf45\xf1\xf5\x945\xb4\xd5L\xfc\xe8ek\xd3\xfc\x14\xcfl\x8b\xa1\xf6\xaf\xdb\x18\xaf\xa0e\x8a\xf7\xb2n\x8d\xa3\xae\xe0w\xfesX\x9f\xe2C\xd9V!\x9a\xe0w\x067!\xb1]\xec5\x1c[E,\xd0\xbe4r\xff\x002D\xe4\xef\x1bU@\xd70e63\xad\x1a\xad\x8e9\x8bj\xcf\xaf\xcck=\xacU\x15KQ\x12\x7fR^&\xff\x00\x812\xe9\xff\x00L\xfc\xda\xd93!\xcc\xac\xb6\xeb\xfd/$t\xb4\xe9\xe7Z\x84\xf0\xf2\xdd\x8b\xe2\xe1o\x8c\x0b\xe8\x08\xbfIu\xe7M\xf5)#\xa6\xb3^R\x8a\xea\xff\x00\xfeYp\xda\x1a\x85^\xf3Sul\x9f\xd4U\xf0\xecJ\x00\x00\x00\x00\x00\x01\xd4\\2\x9cf\xdfT\xfaJ\xfc\x8a\xd1IP\xcfo\x14\xf5\xb1\xb1\xed\xf1\xa2\xae\xe8q\xfd;\xe1\x9f\x0bl\x1fX\xc3\xf7\x80\xef\xc1\xd0zw\xc3>\x17X>\xb1\x87\xef\x0fN\xf8g\xc2\xdb\x07\xd60\xfd\xe0;\xf0t\x1e\x9d\xb0\xcf\x85\xb6\x0f\xaca\xfb\xc3\xd3\xbe\x19\xf0\xb6\xc1\xf5\x8c?x\x0e\xfc\x1d\x07\xa7|3\xe1u\x83\xeb\x18~\xf0\xf4\xed\x86|-\xb0}c\x0f\xde\x03\xbf\x07A\xe9\xdf\x0c\xf8]`\xfa\xc6\x1f\xbc=;\xe1\x9f\x0bl\x1fX\xc3\xf7\x80\xef\xc1\xd0zw\xc3>\x17X>\xb1\x87\xef\x0fN\xf8g\xc2\xdb\x07\xd60\xfd\xe0;\xf0t\x1e\x9d\xb0\xcf\x85\xb6\x0f\xaca\xfb\xc3\xd3\xbe\x19\xf0\xba\xc1\xf5\x8c?x\x0e\xfc\x1d\x07\xa7l3\xe1m\x83\xeb\x18~\xf0\xf4\xed\x86|-\xb0}c\x0f\xde\x03\xbf\x07A\xe9\xdb\x0c\xf8[`\xfa\xc6\x1f\xbc=;\xe1\x9f\x0bl\x1fX\xc3\xf7\x80\xef\xc1\xd0zw\xc3>\x17X>\xb1\x87\xef\x0fN\xf8g\xc2\xeb\x07\xd60\xfd\xe0;\xf0t\x1e\x9d\xb0\xcf\x85\xb6\x0f\xaca\xfb\xc7wK<\x15T\xec\xa8\xa6\x9a9\xe1\x918\x99$nG5\xc9\xdfENJ\x07\xd0\x00\x00\x00\x00\x00\x00\x00\x00\xcf.\x9fY\xc7\xa6]an7K7\x1d\x067\x07\x9d\xf6E\xdd\x16\xa6M\x9f*\xfd\t\xd5\xb1|,R\xf7j.OG\x85\xe0\xb7\xac\xaa\xbbe\x82\xd9G%G\n\xae\xddc\x91=\x83\x13\xc2\xe7p\xb5<*d\x8d\xe6\xe3Yx\xbcV\xdd\xae\x13,\xd5\x95\xb5\x0f\xa8\xa8\x91{_#\xdc\xaer\xfd*\xaa\x07\x10\xd1\xae\x82y\xc7\xa6\xbd\x14\x82\xcdS7\x1d\xc3\x1c\x97\xce/E_d\xb0/\xb2\x85\xde.\x1d\xd8\x9f\xd1\x99\xcaO\x9d\x05\xb3\x8fJz\xdbMi\xaa\x9b\x82\xdf\x91\xc5\xe7\t\x11W\xd8\xa4\xdb\xf1B\xef\x1f\x16\xecO\xe9\x14\r\x1c\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01[:z\xe9\xad\xbf"\xd3)s\x9aJV2\xf7`\xe1t\x92\xb1\xbe\xcazW9\x1a\xf6;\xbf\xc2\xaeG\xa2\xafb#\xf6\xf6\xc6}\x1a\xe1\xab\xb6\xf6]\xb4\xab,\xb6H\xdd\xd2\xa6\xcbW\x17\x89V\x17"/\x8d\x17e2<\x00\x00\x0bw\xe6v\xe9\xdd\r\xce\xe5w\xd4K\xa5+\'[t\xa9El\xe3j+Y2\xb5\x1d$\x89\xbf\xe55\xaeb"\xff\x00=\xde\x02\xee\x90oAkch:8Y*\x11\x11\x1dp\xa9\xaa\xa9\x7f\x8f\xaeti\xfe\x11\xa19\x00\x00\x008\x97\x9b]\xb6\xf3m\x9a\xdbw\xb7\xd2\xdc(\xa7o\x0c\xb4\xf50\xb6H\xde\x9d\xe5k\x91QNX\x02\xa0t\x83\xe8\x8bo\x9a\x8a\xab"\xd2\xc6\xbe\x9e\xad\x88\xb2Id\x91\xfcQ\xca\x9d\xab\xd49y\xb5\xdd\xe69U\x17\xb1\x15\xbc\x91iMD2\xd3\xcf$\x13\xc4\xf8\xa6\x8d\xca\xc9#{U\xaec\x91vTT^h\xa8\xbd\xc3eJi\xd3\xebGi\xe3\xa7\xf5U\xc7i\x1b\x13\xb8\xdb\x15\xf2\x18\x9b\xb2;\x89veN\xc9\xdd\xddQ\xae\xef\xee\xd5\xfc\xe5P\xa6ls\x98\xe4{\x1c\xadsWtT]\x95\x14\xb5\x1d\x1b\xbaV]\xb1\xfa\x9alkR\xea\xa7\xbaYWh\xe1\xba9\x15\xf54\x9d\xee\xb1{eg\x87\x9b\xd3\xf9\xdc\x90\xaa\xc0\r\x91\xb7V\xd2\\h)\xee\x14\x150\xd5RT\xc6\xd9`\x9a\'\xa3\x99#\x1c\x9b\xa3\x9a\xa9\xc9QS\xba}\xcc\xfc\xe8g\xaf2`\xd7\xa8\xb0\x9c\xb2\xb9\xcb\x8bWI\xc3M4\xae\xdd-\xd39{w\xeeD\xe5_d\x9d\x88\xab\xc5\xcb\xd9o\xa0i\xcd7@\x00\x003KQ\x1c\xf5\xd4\x1c\x8d\xca\xe5U[\xb5R\xaa\xef\xda\xbds\xce\xb2\x86\xddt\xaek\x9dCCYT\xd6\xae\xceXbs\xf6\xf1\xec\x9c\x8e\xc7P\xfd\xdf\xe4?:\xd5yW\x17_\xa2<QG\xa0v\x07\xb265\xd2IT\xe7\xaa5\x11\\\xbey\x957^\xfa\xec\x88\x9e$@)\x07\xa5\xfc\x8f\xf4\x1d\xdb\xf5I?q\xfa\x98\xf6G\xfa\n\xed\xfa\xa4\x9f\xb8\xd3m\x93\xbc6N\xf0\x19\x93\xe9\x7f#\xfd\x07v\xfdRO\xdc=/d\x7f\xa0\xae\xdf\xaaI\xfb\x8d6\xd9;\xc3d\xef\x01\x99>\x97\xb2O\xd0Wo\xd5$\xfd\xc3\xd2\xf6I\xfa\n\xed\xfa\xa4\x9f\xb8\xd3m\x93\xbc6N\xf0\x19\x93\xe9{$\xfd\x05w\xfdRO\xdc=/d\x7f\xa0\xae\xdf\xaaI\xfb\x8d6\xd9;\xc3d\xef\x01\x99+\x8fd\x9f\xa0\xae\xdf\xaaI\xfb\x87\xa5\xec\x93\xf4\x15\xdf\xf5I?q\xa6\xdb\'xl\x9d\xe03\'\xd2\xf6I\xfa\n\xed\xfa\xa4\x9f\xb8z_\xc8\xfb=\x02\xbb~\xa9\'\xee4\xdbd\xef\r\x93\xbc\x06dz^\xc8\xff\x00A]\xbfT\x93\xf7\x1f\xbe\x97\xb2?\xd0Wo\xd5$\xfd\xc6\x9bl\x9d\xe1\xb2w\x80\xcc\x9fK\xd9\'\xe8+\xb7\xea\x92~\xe3\xf3\xd2\xf6G\xd9\xe8\x15\xdb\xf5I?q\xa6\xfb\'xl\x9d\xe03#\xd2\xf6G\xfa\n\xed\xfa\xa4\x9f\xb8\xfe&\xb1\xdf\xe1\x89\xd3Mf\xb9\xc5\x1b\x11\\\xe7\xbe\x99\xe8\x8dN\xfa\xaa\xa7#N\xb6N\xf1\xf8\xe4N\x15\xe4\x9d\x80e\xaf\x13\xbf9\x7f\xb4\xbc\xdd\r^\xf7\xe8\x85"9\xcer6\xba\xa5\x1a\x8a\xbd\x89\xc7\xbf/\xa5U~\x92\x97fq\xb2,\xbe\xf5\x14Llq\xb2\xe1P\xd6\xb5\xa9\xb25\x12G""\'x\xb9\xfd\x0c\xfd\xe4\xa9\xbe_Q\xfb@L\xe0\x00\x00\x00\x00\x00\x00\x00*o\x9a1\x9c\xf9\xc3\x13\xb3`\x14\x93m=\xd2_?V\xb5\xab\xcd \x89v\x8d\xaa\x9d\xe7I\xcf\xc7\x11FI\xbbP\xab*u\xef\xa5c\xa8-\xf3:J*\xfb\x93m\xf4og4\x8e\x8a\x1d\xd1\xd2\xb7\xc1\xc2\xd9%\xfe\xb2\x9dOK\x0c\x02\r;\xd6\x8b\xa5\xaa\xdfL\x94\xf6\x9a\xd6\xb6\xbe\xdc\xc6\xa7\xb1lRo\xbbS\xc0\xd7\xa3\xda\x89\xdeD\x02\'>\xd45U\x145\xb0V\xd2L\xe8ji\xe4l\xb0\xc8\xd5\xd9X\xf6\xae\xedrxQQ\x14\xf8\x805\xbfIr\xea|\xf3M\xac9m7\n%\xc6\x91\xb2J\xc6\xf6G2{\x19Y\xfdW\xb5\xc9\xf4\x1e\xa4\xa7\xdeg\x16q\xd7\xdb/\xda{W6\xef\xa6r\\\xe8\x1a\xab\xcf\xabr\xa3&jx\x11\xddZ\xed\xdf{\x8b\x82\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x07[\x95\xb5\x1f\x8b]\x9a\xbd\x8bE2\x7f\xe0S\x1eM\x87\xca=\xcd]>G7\xec)\x8f\x00\x00\x00i\xff\x00C\xe6#:7a\xcdob\xd3J\xbf\xdb<\x8b\xff\x00\xa9,\x91GD/\xe2\xdf\x86\xfc\x92O/!+\x80\x00\x00\x00\x008\x19\x15\xa2\x86\xff\x00`\xaf\xb1\xdc\xe1I\xa8k\xe9\xa4\xa6\xa8\x8d\x7f)\x8fj\xb5\x7f\xc1Nx\x03 s\xbcv\xaf\x12\xcc\xef8\xc5r\xf1TZ\xebe\xa5{\xf6\xd9\x1f\xc0\xe5Drx\x156T\xf0)\xd2\x96\'\xcd\x02\xc7[g\xd7T\xbbE\x1a6;\xdd\xb6\x1a\x97*&\xc8\xb23x\x9c\x9e=\xa3b\xfd%v\x00h\xbfA\xddN\x939\xd3\x05\xb0\xddj\x96k\xde;\xc1L\xf7=}\x94\xd4\xca\x9f\xc0\xbd{\xea\x88\x8a\xc5\xf8\xa8\xab\xcd\xc6t\x13?C\x0c\xc1\xf8\x96\xbeY\x18\xe98i/J\xb6\xaa\x84\xdf\x92\xf5\xaa\x9d_\xfejG\xf4o\xdf\x03L@\x00f\x86\xa1{\xbe\xc8~u\xaa\xf2\xce.\xd7D\xafx\x0cw\xe3U\xfd\xaab\x92\xea\x17\xbb\xec\x8b\xe7Z\xaf,\xe2\xedtJ\xf7\x80\xc7~5_\xda\xa6\x02U<\x86e\xa9\x986\x1fse\xb3#\xc8 \xa0\xac|I*D\xb1H\xf7p*\xaa"\xaf\x03Wm\xf6^\xde\xf1\xeb\xca\x83\xd2\xf3\x07\xcc/:\xab\x1d\xce\xcd\x8d]\xaetr\xdb\xe2ce\xa3\xa4|\xcdG5\\\x8a\x8b\xc2\x8b\xb2\xf3N\xde\xf8\x13\xa7\xab\xc6\x93|0\x83\xf5Y\xff\x00\x0cz\xbci7\xc3\x08?T\x9f\xf0\xcaS\xeai\xa8\xbf\x012o\xaa\xe6\xfb\xa3\xd4\xd3Q~\x02d\xdfU\xcd\xf7@\xba\xde\xaf\x1aM\xf0\xbe\x0f\xd5*?\x0cz\xbci7\xc3\x08?T\x9f\xf0\xcaS\xeai\xa8\xbf\x012o\xaa\xe6\xfb\xa3\xd4\xd7Q>\x02d\xdfUM\xf7@\xba\xde\xaf\x1aM\xf0\xc2\x0f\xd5\'\xfc1\xea\xf1\xa4\xdf\x0c \xfdR\x7f\xc3)O\xa9\xa6\xa2\xfc\x04\xc9\xbe\xab\x9b\xee\x8fSMD\xf8\t\x93}W7\xdd\x02\xebz\xbci7\xc3\x08?T\x9f\xf0\xc7\xab\xc6\x93\xfc/\x83\xf5J\x8f\xc3)O\xa9\xa6\xa2|\x04\xc9\xbe\xaa\x9b\xee\x8fSMD\xf8\t\x93\xfdU7\xdd\x02\xebz\xbci?\xc3\x08?T\x9f\xf0\xc7\xab\xc6\x93|0\x83\xf5I\xff\x00\x0c\xa5>\xa6\x9a\x8b\xf0\x13&\xfa\xaeo\xba=M5\x13\xe0&M\xf5\\\xdft\x0b\xad\xea\xf1\xa4\xdf\x0c \xfdR\xa3\xf0\xc7\xab\xc6\x93|0\x83\xf5I\xff\x00\x0c\xa5>\xa6\x9a\x8b\xf0\x13&\xfa\xaao\xba=Mu\x13\xe0&O\xf5T\xdft\x0b\xad\xea\xf1\xa4\xdf\x0c \xfdR\x7f\xc3\x1e\xaf\x1aM\xf0\xc2\x0f\xd5\'\xfc2\x93\xbfM\xf5\t\x8dW;\x05\xc9\x91\xa8\x9b\xaa\xad\xaa~_\xf8O*\xa9\xb7 4\xcf\x0e\xca\xb1\xfc\xbe\xd2\xb7\\n\xe7\x15\xc6\x8d\xb2\xac.\x91\x88\xe6\xf0\xbd\x11\x15Z\xa8\xe4EE\xd9Qy\xa7b\xa2\x9d\xc3\xbd\xaa\xf8\x8a\xfb\xd0O\xde\xca\xf5\xf3\xdb\xfc\x84%\x82w\xb5_\x10\x19\x99\x9b\xfb\xb4\xbe|\xe5S\xe5\\\\\xde\x86~\xf2T\xdf/\xa8\xfd\xa2\x99f\xfe\xec\xef\x9f9T\xf9W\x177\xa1\x9f\xbc\x957\xcb\xea?i\x00\x99\xc0\x00\x00\x00\x00\x00\x08\xa3\xa5\x8eq\xe9\x0fC\xef\x97\x08&\xea\xae\x15\xec\xf46\x85Qv^\xb6TTW\'\x85\xacG\xbd<-BW(_\x9a\x1d\x9b\xad\xe7Q\xad\x98=\x14\xaa\xfa{\x1c\x1dmCZ\xbd\xb53".\xca\x9d\xde\x18\xf86\xf8\xee\x03\xb9\xf38\xf0\x7f<\xde\xef\xba\x83W\x0e\xf1\xd13\xd0\xda\x15T\xe5\xd6\xbd\x11\xd2\xb9<-g\x02x\xa4S\xdc\xf9\xa28O\xa2\xfaqm\xcdib\xde\xa6\xc5S\xd5T\xb9\x13\xb6\x9aeF\xee\xbd\xfd\xa4H\xf6\xf8\xee&>\x8f8Ji\xf6\x8f\xe3\xf8\xdc\x91#+#\xa7I\xeb\xb9sZ\x89=\x9c\x88\xbd\xfd\x95xS\xc0\xd4=>o\x8fQe\x98}\xdf\x1a\xb8\xa6\xf4\xb7:9)d]\xb7V\xf1\xb5Q\x1c\x9e\x14]\x95<(\x80c\xf89\xd9\r\xa6\xb6\xc3~\xb8Y.1\xf5U\xb6\xfa\x99)j\x19\xf9\xb21\xca\xd7\'\xf6\xa2\x9c\x10=\xd6\x81\xe6\xce\xd3\xed[\xc7\xf2\x87H\xe6\xd2\xc1R\x91\xd6\xa2wi\xe4\xf6\x12\xf2\xee\xec\xd7+\x91;\xedCWc{$cd\x8d\xcd{\x1c\x88\xadsWtT^\xea\x18\xd0i\x97C\\\xe3\xd3\xb6\x86Z|\xf17Yq\xb3\x7f\xec\xba\xbd\xd7\x9a\xf5h\x9d[\x97\xba\xbb\xc6\xac\xe7\xddTp\x130\x00\x00\x00\x00\x00\x00\x00\x00\x00\x0e\xbb(\xf75t\xf9\x1c\xdf\xb0\xa6<\x1b\x0f\x94{\x9a\xba|\x8eo\xd8S\x1e\x00\x00\x00\xd4\x1e\x88_\xc5\xbf\r\xf9$\x9e^BW"\x8e\x88_\xc5\xbf\r\xf9$\x9e^BW\x00\x00\x00\x00\x00\x00\x02\x9cy\xa5\xd6\xc6\xad.\x13xjl\xe6\xbe\xae\x99\xeb\xdfEH\x9c\xdf\xec\xd9\xdf\xdaR\xe2\xf9y\xa4p\xb5\xda[\x8eTm\xec\x99{\xe0E\xf0:\t\x17\xfe\x94(h\x03\x93j\xae\xa8\xb6]).T\x8f\xe0\xa9\xa4\x9d\x93\xc2\xef\xcd{\x1c\x8ej\xff\x00j!\xc6\x00l\x8d\xba\xaa*\xeb}5t+\xbcU\x116V/\xf3\\\x88\xa9\xfe`\xe8t\xa6e\xa8\xd2\xecNwo\xbc\x96J7\xae\xfe\x18\x18\xa0\x0c\xf5\xd4?w\xd9\x0f\xce\xb5^Y\xc5\xda\xe8\x95\xef\x01\x8e|j\xbf\xb5LR]B\xf7}\x91|\xebU\xe5\x9c]\xae\x89^\xf0\x18\xe7\xc6\xab\xfbT\xc0J\xa0\x1e\x07P\xb5\x7f\x04\xc1/1\xd9\xf2+\xa4\xd0V\xbe\x14\x9b\xaa\x8a\x92YxX\xaa\xa8\x8a\xaa\xd6\xaa&\xfb/.\xdf\xf0\x03\xdf\x02\x1e\xf5\xcai7\xe9\xaa\xdf\xabg\xfb\x83\xd7\'\xa4\xdf\xa6\xeb~\xad\x9f\xee\x010\x82\x1e\xf5\xc9\xe97\xe9\xaa\xdf\xabg\xfb\x83\xd7\'\xa4\xdf\xa6\xab~\xad\x9f\xee\x010\x82\x1e\xf5\xcai7\xe9\xaa\xdf\xabg\xfb\x83\xd7\'\xa4\xdf\xa6\xab~\xad\x9f\xee\x010\x82\x1e\xf5\xcai7\xe9\xaa\xdf\xabg\xfb\x83\xd7)\xa4\xdf\xa6\xab~\xad\x9f\xee\x010\x82\x1e\xf5\xcai7\xe9\xaa\xdf\xabg\xfb\x83\xd7\'\xa4\xdf\xa6\xab~\xad\x9f\xee\x010\x82\x1e\xf5\xcai7\xe9\xaa\xdf\xabg\xfb\x83\xd7)\xa4\xdf\xa6\xab~\xad\x9f\xee\x010\xafa\x99\xf9\xfb\x11\x99\xeeD\xc6\xa25\xad\xba\xd5""w\x13\xadqsd\xe9+\xa4\xedb\xabn\xf5\xefTNMKl\xc8\xab\xe0\xe6\xd4B\x93d\xb7\x06\xdd\xb2;\x9d\xd5\x91\xba6V\xd6MP\x8cr\xee\xadG\xbd\\\x88\xbe-\xc0\xb7=\x04\xfd\xec\xaf\x7f=\xbf\xc8BX\'{U\xf1\x15\xf3\xa0\x9f\xbd\x9d\xef\xe7\xb7\xf9\x08K\x06\xefj\xbe 336\xf7g|\xf9\xca\xa7\xca\xb8\xb9\xbd\x0c\xfd\xe4\xa9\xbe_Q\xfbHS,\xdf\xdd\x9d\xf3\xe7*\x9f*\xe2\xe6t2\xf7\x92\xa6\xf9}G\xed \x13@\x00\x00\x00\x00\x00\x0e\xbb(\xbdP\xe3\x98\xdd\xca\xff\x00s\x93\xab\xa2\xb7R\xc9U;\xbb\xbc\x0cj\xb9v\xf0\xf2\xe4\x9d\xf3;z9Y\xab\xb5\x8b\xa4\xfcw\xdb\xcc}tL\xad\x96\xfbr\xdf\x9bQ\x1a\xfe&3\x9fkz\xc7F\xdd\xbf7~\xf1b|\xd0\x9c\xe3\xd0-/\xa3\xc3\xe9&\xe1\xac\xc8j?\x86D^iM\n\xa3\x9d\xe2\xdd\xeb\x1axQ\x1c\x87\xf3\xe6z`\xfe\x82i\x8dvcW\x0f\r^AQ\xc3\x02\xaasJhUZ\xdf\x16\xefY\x17\xc2\x88\xd5\x02\xce\x00\x00\xcfo4\x03\t\xf4\xbb\xab\xf0\xe4\xd4\xd0\xf0Qdt\xc93\x95\x13dJ\x98\xb6d\x89\xf4\xb7\xabw\x85\\\xa5p4\xa3\xa6\xde\x13\xe9\xc3B\xee5t\xf0\xf1\xd7\xd8\x1e\x9786Nj\xc6"\xa4\xc9\xe2\xea\xd5\xce\xdb\xbe\xc43\\\x01e<\xcf\xbc\xe3\xd2\xfe\xab\xd4\xe2uSp\xd1dt\xfc\x11\xa2\xaf$\xa9\x89\x15\xec_\x06\xedY\x1b\xe1UiZ\xce~;v\xad\xb0_\xed\xf7\xcbl\xbdUm\xbe\xa6:\xaaw\xfek\xd8\xe4sW\xfbP\r\x89\x07M\x83\xe4TYn\x1dh\xc9\xad\xcb\xbd-\xce\x92:\x98\xd3}\xd5\x9cMEV\xaf\x85\xab\xba/\x85\x14\xee@\x00\x00\x00\x00\x00\x00\x00\x00\xeb\xb2\x8fsWO\x91\xcd\xfb\nc\xc1\xb0\xf9G\xb9\xab\xa7\xc8\xe6\xfd\x851\xe0\x00\x00\rA\xe8\x85\xfc[\xf0\xdf\x92I\xe5\xe4%r(\xe8\x85\xfc[\xf0\xdf\x92I\xe5\xe4%p\x00\x00\x00\x00\x00\x00*\xe7\x9aC*&\x93c\xd0\xf7]}k\xbf\xb2\tS\xff\x00R\x85\x17g\xcd-\xaf\xe0\xb1\xe1V\xb4w\xfaj\x9a\xba\x85O\x88\xd8\x9a\x8b\xff\x00\x98\xa5&\x00\x01\xe8\xf4\xc6\xc2\xec\xa7Q\xb1\xdcu\x18\xafm\xc6\xe7\x05;\xd1;\x8ct\x88\x8f_\x127u\xfa\x00\xd5\xcc\x12\x85\xd6\xcc"\xc3mzl\xeaKm<\n\x9e\x16D\xd6\xff\x00\xe8\x0e\xe4\x01\x9a\x1a\x85\xee\xfb!\xf9\xd6\xab\xcb8\xbb]\x12\xbd\xe01\xdf\x8dW\xf6\xa9\x8aK\xa8^\xef\xb2\x1f\x9dj\xbc\xb3\x8b\xb5\xd1+\xde\x07\x1d\xf8\xd5\x7fj\x98\tT\xaa]*t\xbf:\xca\xb5=\xb7\x8cw\x1f\x9e\xe3D\xea\x08\xa3\xeb#\x91\x89\xb3\x9a\xae\xdd\x15\x1c\xe4^\xea\x7fik@\x19\xf5\xea\x19\xaa\xff\x00\x03k?\xbd\x8b\xef\x0fP\xcdW\xf8\x1bY\xfd\xec_|\xd0P\x06`^\xedW\x1b%\xd6\xa2\xd5v\xa3\x96\x8e\xb6\x99\xea\xc9\xa1\x95\xbb9\x8b\xff\x00\xfb\x9e\xfd\xd4TS\x99\x89b\xf9\x06Yt[f9j\xa9\xb9U\xa4k#\xa3\x85\xbe\xd5\xa9\xb6\xeer\xaf$M\xd5\x13u^\xd5D\xee\x9e\xeb\xa5~\xde\xaf\xd9*\'~\x97\xec\xb0\x9e\xef\xa0\x8f\xbb\\\x83\xe6\xe6\xf9F\x81\x1cz\x86j\xbf\xc0\xda\xcf\xefb\xfb\xc3\xd47U\xfe\x06\xd6\xff\x00{\x17\xde4\x14\xf8\xd6USQ\xd3\xba\xa2\xb2\xa2*x[\xb7\x14\x92\xbd\x1a\xd4\xddvM\xd5y\x01@=C5_\xe0eg\xf7\xb1}\xe1\xea\x19\xaa\xff\x00\x03k\x7f\xbd\x8b\xef\x17\xd2\x96\xfbd\xab\xa8e=-\xe2\xdf<\xcf\xe4\xd8\xe3\xa9c\x9c\xee\xef$E\xe6v g\xd7\xa8n\xab\xfc\x0c\xad\xfe\xf6/\xbc=Cu_\xe0mg\xf7\xb1}\xe3AN\xb2\\\x86\xc1\x0c\xaf\x86[\xdd\xb69\x18\xaa\xd71\xd5LEj\xa7j*n\x05\r\xf5\x0c\xd5\x7f\x81\x95\x9f\xde\xc5\xf7\x87\xa8n\xab\xfc\r\xac\xfe\xf6/\xbch\x14\x12\xc5<,\x9a\x19\x19,Oj9\x8fc\xb7k\x91{\x15\x15;P\xfe\xc0\xcf\xafP\xcdW\xf8\x1b[\xfd\xec_xz\x86\xea\xbf\xc0\xca\xcf\xefb\xfb\xc6\x82\x80!\xbe\x89\x18vE\x86i\xf5\xc6\x8b%\xb7\xba\x82\xaa\xa6\xe8\xf9\xe3\x85\xcfk\x9d\xd5\xf5Q5\x15xUQ9\xb5\xdf\xd8Ln\xf6\xab\xe2?O\xc7{U\xf1\x01\x99\x99\xb7\xbb;\xe7\xceU>U\xc5\xcc\xe8e\xef%O\xf2\xfa\x8f\xdaB\x99\xe6\xde\xec\xef\x9f9T\xf9W\x177\xa1\x9f\xbc\x95?\xcb\xea?i\x00\x99\xc0\x00\x00\x00\x00<\x07Hl\xdd4\xf7G\xef\xf9$r\xa3+c\xa7X(y\xf3Z\x89=\x84j\x9d\xfd\x95x\x97\xc0\xd5\x02\x8d\xf4\x93\xbeWj\xff\x00I\xc9,vW\xf5\xd1GY\x15\x8a\xdb\xb76\xfb\x17\xf0\xbd\xfc\xbf%dt\x8e\xdf\xf3v\xef\x1a!\x8aY(q\xacf\xd9\x8f[\x19\xc1Gm\xa5\x8e\x96\x14\xee\xf0\xb1\xa8\xd4U\xf0\xae\xdb\xaa\xf7\xca5\xe6za\x0b|\xd4\xcb\x86k[\x12\xbe\x9a\xc3O\xc3\x03\x9d\xdd\xaa\x99\x15\xa8\xbe\x1d\x98\x92o\xdeW5K\xf0\x00\x00\x07\xce\xaa\x08j\xa9\xa5\xa6\xa8\x89\x92\xc33\x169\x18\xf4\xdd\xaej\xa6\xca\x8a\x9d\xe5C%\xf5\x87\x10\x9b\x03\xd4\xeb\xfe\'*;\x86\xdfX\xe6@\xe7v\xbe\x17{(\x9c\xbe69\xab\xf4\x9a\xd8R\x8f4{\t\xea.\x96\r@\xa4\x87fT\xb1m\x95\xceD\xe5\xd67w\xc4\xab\xe1V\xf5\x89\xe2b\x01O\x80\x00_\x0f3\xb38\xf4S\x03\xba\xe0\xb5soQe\x9f\xcf4\x88\xab\xdbO2\xaa\xb9\x11?\x9b\'\x12\xaf\xf4\x88Z\x83.:,g\x1e\x90u\xb6\xc5v\x9an\xaa\xdfW\'\xa1\xf5\xea\xab\xb3z\x99U\x1b\xc4\xbe\x06\xbb\x81\xff\x00\xd45\x1c\x00\x00\x00\x00\x00\x00\x00\x00\x0e\xbb(\xf75t\xf9\x1c\xdf\xb0\xa6<\x1b\x0f\x94{\x9a\xba|\x8eo\xd8S\x1e\x00\x00\x00\xd4\x1e\x88_\xc5\xbf\r\xf9$\x9e^BW"\x8e\x88_\xc5\xbf\r\xf9$\x9e^BW\x00\x00\x00\x00\x00\x01\xfcO,pB\xf9\xa6\x91\xb1\xc5\x1bU\xcf{\x97dj"n\xaa\xab\xdc@(\'\x9a+}m~\xb1[,\xb1I\xc4\xcbU\xa5\x9dcw\xf6\xb2\xca\xf79S\xfe\xe2F\xa5e=~\xb4e\xbe\x9e\xb5W#\xca\xda\xaeXk\xeb\x9e\xea~.\xd4\x81\xbb2$_\n1\xadC\xc8\x00,\x8f\x99\xf5\x87\xbe\xfb\xacsd\xd2\xc7\xbd&;H\xe9\x11\xca\x9b\xa7_2:6\'\xfd\xde\xb5\xde6\xa1[\xda\x8a\xe7#Z\x8a\xaa\xab\xb2"wM9\xe8\x93\xa6\xae\xd3]"\xa2\xa5\xaf\xa7\xea\xafwE\xf3\xf5\xc9\x1c\x9e\xc9\x8fr\'\x04K\xf1\x1b\xb2*~w\x17|\tx\x00\x06hj\x17\xbb\xec\x87\xe7Z\xaf,\xe2\xedtK\xf7\x80\xc7>5_\xda\xa6).\xa1{\xbe\xc8~u\xaa\xf2\xce.\xd7D\xafx\x0cw\xe3U\xfd\xaa`%P\x00\x00\x00\x14\x1b\xa5\x7f\xbf\xfeK\xff\x00\x0b\xf6XOw\xd0E?\xf7\xd7 \xf9\xb9\xbeQ\xa7\x84\xe9_\xef\xff\x00\x92\xff\x00\xc2\xfd\x96\x13\xdet\x11\xf7k\x90\xfc\xdc\xdf(\xd0-\xe9\x00\xf4\xe8s\x93J-LG*5\xd7\xc8\x91\xc8\x8b\xc9S\xa8\x9dy\xfd$\xfc@\x1d:=\xea\xad?>G\xe4\'\x02\xabiJ\xacz\xa1\x8a>5V9/t{*r_\xf4\xec4\xa0\xcd}-\xf7\xce\xc5>z\xa3\xf2\xec4\xa1\x00\xf9U\xae\xd4\xb3*.\xca\x8c_\xf22\xeedGJ\xe797U]\xd5W\xbaj%g\xfa\xa4\xdf\xd1\xbb\xfc\x8c\xbb\x93\xdb\xaf\x8c\x0b\xcf\xd0\xe9\xcev\x86\xdb\x91\xces\x91\xb5u(\x9b\xafbu\xae\xe4\x84\xc4C}\x0e=\xe3\xa8>YU\xe5\\L\x80\x00\x00\x0f\xc7{U\xf1\x1f\xa7\xe3\xbd\xaa\xf8\x80\xcc\xcc\xdf\xdd\x9d\xf3\xe7*\x9f*\xe2\xe6t2\xf7\x92\xa7\xf9}G\xed!L\xf3\x7fvw\xcf\x9c\xaa|\xab\x8b\x9b\xd0\xcf\xdeJ\x9b\xe5\xf5\x1f\xb4\x80L\xe0\x00\x00\x00\x05#\xf3G3\x8f=_lZ}I6\xf1Q3\xd1\x1a\xe6\xa2\xf2\xeb^\x8a\xd8\x9a\xbe\x16\xb3\x8d|R!u.U\xb4\xb6\xdbuM\xc6\xbaf\xc1KK\x0b\xe6\x9eWv1\x8dEs\x9c\xbe\x04DU3[\x06\xa3\xab\xd7\xae\x94\xac\xaa\xb8D\xf7R\xdd.n\xae\xacc\xb9\xf5tQsH\xd5\x7f\xa3k#E\xef\xaa\x01uz"\xe0\xfe\x91t6\xcbK<=]\xc2\xe6\xdfD\xebwM\x97\x8eTEkW\xbc\xad\x8d#j\xa7}\x14\x97\x02""""""v"\x00\x00\x00\x07\x80\xe9\x0f\x85&\xa0h\xeeC\x8d\xc7\x12IY%2\xcfC\xcb\x9f\x9e#\xf6q\xa2w\xb8\x95\xbc+\xe0r\x9e\xfc\x01\x8d\n\x8a\x8a\xa8\xa8\xa8\xa9\xda\x8a~\x12\xd7K|\'\xd26\xba_(\xe0\x87\xab\xa0\xb8\xbf\xd1:$D\xd9:\xb9\x95U\xc8\x9e\x06\xc8\x925<\rB%\x00j_E\xfc\xe3\xd3\xfe\x8aXo3M\xd6\xdc)\xe2\xf3\x8d\xc1Uww_\x16\xcdW/\x85\xcd\xe1\x7f\xf5\xcc\xb4-o\x99\xd5\x9cz\x1b\x9b]\xb0:\xb9\xb6\xa7\xbcC\xe7\xba6\xaa\xf6TD\x9e\xc9\x11;\xee\x8fu_\xe8\x90\x0b\xd6\x00\x00\x00\x00\x00\x00\x00\x03\xae\xca=\xcd]>G7\xec)\x8f\x06\xc3\xe5\x1e\xe6\xae\x9f#\x9b\xf6\x14\xc7\x80\x00\x005\x07\xa2\x17\xf1o\xc3~I\'\x97\x90\x95\xc8\xa3\xa2\x17\xf1o\xc3~I\'\x97\x90\x95\xc0\x00\x00\x00\x00\x15\xdb\xa7.\xaa\xc5\x85\xe9\xd4\x98\x85\xb2\xa1\xbe\x8e\xe4Q:\x175\xae\xf6PR/\xb1\x91\xeb\xdeWsc|nT\xf6\xa7\xb7\xe9\x01\xad\x98\xc6\x92X\xd5\xd5\xd26\xbe\xfdQ\x1a\xad\r\xae\'\xfb9;\x88\xf7\xaf\xe4G\xbfuy\xae\xca\x88\x8b\xb2\xed\x9ay\xc6S{\xcd2\x9a\xec\x97!\xacu]\xc6\xb6N9^\xbc\x91\xa9\xd8\x8djw\x1a\xd4\xd9\x11;\x88\x80t\xa0\x16K\xa3\x17F[\xbesQG\x94\xe6\xb0Ml\xc5QRX\xa9\xdd\xbb\'\xb8\xa7j#S\xb5\x91\xafu\xfd\xaa\x9e\xd7\xb7\x89\x03\xb2\xe89\xa2R\xe4\xb7\xe85#$\xa5V\xd8\xed\x93q[a\x91\xbc\xab*Z\xbe\xdf\xc2\xc8\xd7\xbb\xddr"~K\x90\xbe\x87\x1e\xd9CGl\xb7\xd3\xdb\xad\xd4\xb0\xd2Q\xd3F\xd8\xa0\x82\x16#\x19\x1b\x1a\x9b#Z\x89\xc9\x11\x10\xe4\x00\x00\x01\x9a\x1a\x85\xee\xfb"\xf9\xd6\xab\xcb8\xbb]\x12\xbd\xe01\xdf\x8dW\xf6\xa9\x8aK\xa8^\xef\xb2\x1f\x9dj\xbc\xab\x8b\xb5\xd1+\xde\x07\x1d\xf8\xd5\x7fj\x98\tT\x00\x00\x00\x05\x06\xe9_\xef\xff\x00\x93x\xe9~\xcb\t\xef:\x08\xfb\xb5\xc8>no\x94i\xe0\xfaW\xfb\xff\x00\xe4\xbe:_\xb2\xc2{\xce\x82>\xedr\x1f\x9b\x9b\xe5\x1a\x05\xbd \x1e\x9d\x1e\xf5V\x9f\x9f#\xf2\x13\x93\xf1\x00\xf4\xe8\xf7\xaa\xb4\xfc\xf9\x1f\x90\x9c\n\xab\xa5\xbe\xf9\xd8\xa7\xcft~]\x86\x94!\x9a\xfaY\xef\x9f\x8a|\xf7G\xe5\xd8iB\x01\xf2\xac\xff\x00T\x9b\xfa7\x7f\x91\x97r{u\xf1\x9a\x89Y\xfe\xa97\xf4n\xff\x00#.\xe4\xf6\xeb\xe3\x02\xf3t8\xf7\x8e\xa0\xf9eW\x95q1\x90\xe7C\x8fx\xdb\x7f\xca\xea\xbc\xab\x89\x8c\x00\x00\x01\xf8\xefj\xbe#\xf4\xfcw\xb5_\x10\x19\x99\x9b\xfb\xb3\xbe|\xe5S\xe5\\\\\xde\x86~\xf2T\xdf/\xa8\xfd\xa4)\x96o\xee\xce\xf9\xf3\x95O\x95qsz\x19\xfb\xc9S|\xbe\xa3\xf6\x90\t\x9c\x00\x00\x00\x05~\xe9\xe1\x9c\xfaU\xd1il\x94\xb3\xf5w\x0c\x92o91\x11vr@\x9b:gx\xb6\xe1b\xff\x00Hx\x7f3\x93\x07\xf3\xa6?|\xd4\n\xb8v\x96\xbeOC\xa8\\\xa9\xcf\xa9b\xa3\xa5rx\x1c\xfe\x14\xf1\xc6\xa43\xd3\x979\xf4\xdd\xad\xd5v\xcai\xb8\xed\xf8\xec~\x87D\x88\xbc\x96d]\xe6w\x8f\x8dx\x17\xfa4=\xee\x99\xf4\xb7\xc7\xb0l\x02\xcb\x89Pi\xe5c\xe1\xb6R6\x15\x91.MoZ\xfe\xd7\xc9\xb7W\xc9\\\xf5s\xb6\xf0\x81x\x81P=|v\xaf\xe4\xea\xb7\xebF\xfe\x18\xf5\xf1\xda\xbf\x93\xaa\xdf\xad\x1b\xf8`[\xf0T\x0f_\x1d\xab\xf9:\xad\xfa\xd1\xbf\x86=|v\xaf\xe4\xea\xb7\xebF\xfe\x18\x16\xfc\x15\x03\xd7\xc7j\xfeN\xab~\xb4o\xe1\x8f_\x1d\xab\xf9:\xad\xfa\xd1\xbf\x86\x07u\xe6\x89\xe1>\x8a\xe9\xed\xaf6\xa5\x87z\x8b\x1dOQT\xe4O\xfe\x1eeDEU\xfel\x88\xc4O\xe9\x14\xa1\x85\xbf\xce\xfa`c\xd9v\x19x\xc6.\x1auX\x94\xd7:9)\x9e\xefD\xd8\xaa\xce&\xaa#\xd3\xf8>\xd6\xae\xca\x9e\x14B\xa0\x00;\xcc\x07$\xad\xc3\xf3[>Qo\xff\x00Y\xb6VGR\xd6\xef\xb2=\x1a\xbb\xb9\x8b\xe0rn\xd5\xf0*\x9d\x18\x03b\xec7J+\xe5\x8e\x82\xf5m\x95&\xa2\xaf\xa6\x8e\xa6\x9eO\xce\x8d\xedG5\x7f\xb1P\xe6\x95\xbf\xcc\xff\x00\xce=1i\x1c\xd8\xbdT\xdcu\xd8\xe5GT\xd4U\xddV\x9aUW\xc6\xbfC\xba\xc6\xf8\x11\xad,\x80\x00\x00\x00\x00\x00\x00\x1dvQ\xeej\xe9\xf29\xbfaLx6\x1f(\xf75t\xf9\x1c\xdf\xb0\xa6<\x00\x00\x01\xa4}\x153L:\xd9\xd1\xf7\x12\xa1\xb8\xe5\x96\x1a*\xb8\xa9dI \xa8\xb8\xc5\x1c\x8c^\xbaE\xd9Z\xe7"\xa7%BO\xf5C\xc0>\x1cc?[A\xf7\x8a\x0b\xa7\x1d\x15\xf5\x0b;\xc2m\x99m\xa2\xf1\x8b\xc1Cq\x8d\xd2C\x1dUL\xed\x95\xa8\x8eV\xfb$l.D\xe6\xd5\xecU=\x0f\xac\xabT\xff\x00O\xe1\x9f\xaeT\xff\x00\xfa\xe0]y5\x1bObb\xbeL\xf3\x16cS\xba\xeb\xbc\x08\x9f\xb6uU\xfa\xcf\xa4\xb4Q\xbaI\xb5\x1f\x16r57^\xa6\xe7\x14\xcb\xf4#\x15UJ{\xeb*\xd5?\xd3\xf8g\xeb\x95?\xfe\xb9\xcc\xa6\xe8M\xa8.\xdb\xcf9V/\x1f\x7f\xabt\xef\xff\x008\xd0\t\xe3$\xe9o\xa3V\x94zQ\xdc\xee\x97\xb7\xb7\xf2hh\x1e\x88\xab\xe3\x97\x81>\x9d\xc8#S\xfae\xe5\xb7\xaai\xed\xf8M\x9a\x0cr\t\x11[\xe7\xc9\x9f\xe7\x8a\xad\xbb\xedM\x91\x8c_\xa1\xdbw\x14\xeem\x9d\x07n\xafT\xf4OP\xe8\xa9\xd3\xba\x94\xf6\xc7K\xfbR4\xf6X\xf7B|\x1a\x96D}\xf3*\xbf\\\xf6\xfc\x88\x1b\x1d3W\xc7\xc9\xeb\xb7\x89P\n1v\xb8\xdc.\xd7\x19\xeeWZ\xea\x9a\xea\xda\x87q\xcdQQ*\xc9$\x8e\xef\xb9\xcb\xcdT\xf7\xda]\xa2\x1a\x93\xa8\xb2\xc6\xeb\x0e;<4\x0f\xdb{\x8dr,\x14\xc8\x9d\xf4z\xa6\xef\xf11\x1c\xbe\x03B0=\x0b\xd2\x9c*FOe\xc3\xa8\x1dV\xdeiUX\x8bU*/}\xae\x91W\x81~.\xc4\x90\x05y\xd0\xde\x8a\xb8f\x0c\xf8n\xf9;\xa3\xcao\xacTs\x16h\xb6\xa4\xa7w\xf3"]\xf8\x95?9\xfb\xf7\x15\x11\xaaXdDD\xd99 \x00\x00\x00\x00\x00f\x86\xa1\xfb\xbe\xc8~u\xaa\xf2\xce.\xd7D\xbfx\x1cw\xe3U\xfd\xaab\x92\xea\x1f\xbb\xfc\x8b\xe7Z\xaf*\xe2\xedtJ\xf7\x80\xc7~5_\xda\xa6\x02U\x00\x00\x00\x01A\xfaW\xfb\xff\x00\xe4\xbf\xf0\xbfe\x84\xf7}\x04}\xda\xe4?77\xca!\xe0\xfaX{\xff\x00d\xbf\xf0\xbfe\x84\xf7\x9d\x04}\xda\xe4?77\xca \x16\xf4\x80zt{\xd5\xda~|\x8f\xc8NO\xc4\x01\xd3\xa3\xde\xae\xd1\xf3\xe4~Bp*\xbe\x96\xfb\xe7b\x9f=\xd1\xf9v\x1aN\x86k\xe9o\xbev)\xf3\xdd\x1f\x97a\xa5\x08\x07\xca\xb3\xfdRo\xe8\xdd\xfeF]\xc9\xed\xd4\xd4J\xcf\xf5I\xbf\xa3w\xf9\x19w\'\xb7_\x18\x17\x9b\xa1\xc7\xbcm\xbf\xe5u^U\xc4\xc6C\x9d\x0e=\xe3m\xff\x00,\xaa\xf2\xae&0\x00\xe1\xde.\x96\xdb5\xbe[\x85\xda\xbe\x9a\x86\x92$\xdeI\xaa%H\xd8\xd4\xf0\xaa\xf2+\xfe\xa6t\xa3\xb1\xdbY=\x0e\x0fB\xb7z\xbe\x1d\x99]R\xd7GL\xc7w\xd1\x9c\x9e\xfd\xbb\xde\xc5;\xca\xa0X[\x85m\x1d\xba\x8eZ\xda\xfa\xa8))\xa2o\x14\x93M"1\x8cN\xfa\xaa\xf2B\x07\xd4\xde\x93\x98\xc5\x93\xad\xa1\xc4i\x96\xff\x00Z\xd7+\x16w*\xc7J\xcf\n.\xdcRs\xef""\xf7\x1cU\xfc\xc75\xcd\xb5\x12\xee\xdfE\xee5\xd79d\x91:\x8a(Qz\xb6\xbb\xb1\x128\x9b\xcb\x7f\n&\xeb\xddU$\xcd6\xe8\xcd\x97\xdf\xd6:\xcc\x9efc\xb4*\xe4U\x89\xed\xeb*^\xde\xd5\xd9\x88\xbb3\xbd\xbb\x97t_\xc9P \xfa\xc9\xeanW)\xaa\xa5E\x92\xa6\xaawH\xee\x16\xfbg\xbd\xdb\xae\xc8\x9d\xf5^\xc2\xf6\xf4R\xb2]l:9CIx\xa1\x9e\x86\xa6Z\x99\xe6Hgb\xb1\xe8\xc7?\xd8\xaa\xb5y\xa6\xfbo\xcf\xb8\xa8w\xbas\xa4\xf8>\x06\xc6\xbe\xc9hc\xebQwZ\xea\xad\xa5\xa8U\xdbnNT\xd9\xbc\xbb\x8dF\xa1\xee@\x00\x00\x1eSW\xb2\xf80-4\xbfe\xb3\xf0*\xdb\xe9\x1c\xf8Z\xee\xc7\xcc\xbe\xc6&}/sS\xe9=YN\xfc\xd1\xdc\xe3\xaa\xa1\xb0\xe9\xe5\x1c\xde\xcag-\xce\xbd\xa8\xbf\x90\x9b\xb2\x16\xaf\x81W\xac]\xbf\x9a\xd5\x02 \xe8\x87\xa6\x94z\xb7\xaa\xb5\xf5\x19m<\xb7+5\r<\x95w\x14t\xafg_4\xaa\xa9\x1bU\xecTr*\xb9\\\xfeJ\x9b\xf5j[\xdfZ\xe6\x84\xfc\x06\xff\x00\x9bV\xfe1\xd5\xf4\x18\xc1\xfd)h\x8d-\xd6\xa6\x1e\x0b\x86E\'\xa2\x12*\xa74\x87n\x18[\xe2\xe0N4\xfe\x91I\xe8\x08g\xd6\xb9\xa1?\x01\xbf\xe6\xd5\xbf\x8c=k\x9a\x13\xf0\x1b\xfem[\xf8\xc4\xcc\x00\x86}k\x9a\x13\xf0\x1b\xfem[\xf8\xc3\xd6\xb9\xa1?\x01\xbf\xe6\xd5\xbf\x8cL\xc0\x08g\xd6\xb9\xa1?\x01\xbf\xe6\xd5\xbf\x8c=k\x9a\x13\xf0\x1b\xfem[\xf8\xc4\xcc\x00\x86}k\x9a\x13\xf0\x1b\xfem[\xf8\xc5D\xe9\xa1\xa5\x16\x8d1\xcf\xed\xee\xc6(\x1fE\x8f\xdd\xa8\xf8\xe9\xe2Y_"G4k\xc3+\x11\xcfUr\xf2V;\x9a\xaf\xb7T\xecCHH/\xa6\xfe\x11\xe9\xbbC+\xeb\xa9\xa1\xe3\xaf\xc7\xde\x97(U\x13\x9a\xc6\xd4T\x99<\\\n\xae\xf1\xb1\x00\xcd\xb0\x00\x13WB\xfc\xe3\xd2V\xb9\xda\xe3\xa8\x9b\xab\xb7_\x13\xd0\xba\xad\xd7\x92,\x8a\x9dS\xbb\xdc\xa4F&\xfd\xc4s\x8d.1\xa6)\x1f\x14\xac\x96\'\xb9\x921\xc8\xe6\xb9\xab\xb2\xb5S\xb1QM_\xd0\x9c\xd5\x9a\x83\xa4\xf8\xfeS\xc6\xd7T\xd5R\xa3+\x11?&\xa1\x9e\xc2T\xdb\xb8\x9cMUO\x02\xa0\x1e\xdc\x00\x00\x00\x00\x00\x07]\x94{\x9a\xba|\x8eo\xd8S\x1e\r\x87\xca=\xcd]>G7\xec)\x8f\x00\x00\x00j\x0fD/\xe2\xdf\x86\xfc\x92O/!+\x91GD/\xe2\xdf\x86\xfc\x92O/!+\x80\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\x9a\x1a\x85\xee\xfb"\xf9\xd6\xab\xcb8\xbb]\x12\xbd\xe0q\xdf\x8dW\xf6\xa9\x8aK\xa8^\xef\xb2\x1f\x9dj\xbc\xb3\x8b\xb5\xd1+\xde\x03\x1d\xf8\xd5\x7fj\x98\tT\x00\x00\x00\x05\x06\xe9a\xef\xfd\x92\xff\x00\xc2\xfd\x96\x13\xdet\x11\xf7i\x90\xfc\xdc\xdf(\x87\x84\xe9a\xef\xfd\x92\xff\x00\xc2\xfd\x96\x13\xdd\xf4\x11\xf7k\x90\xfc\xdc\xdf(\xd0-\xe9\x00t\xe8\xf7\xab\xb4\xfc\xf9\x1f\x90\x9c\x9f\xc8\x07\xa7G\xbdU\xa7\xe7\xc8\xfc\x84\xe0U]-\xf7\xce\xc5>{\xa2\xf2\xec4\xa1\x0c\xd7\xd2\xdf|\xfcS\xe7\xba?.\xc3J\x10\x0f\x95g\xfa\xa4\xdf\xd1\xbb\xfc\x8c\xbb\x93\xdb\xaf\x8c\xd1\xedG\xce\xf1l&\xcd5NCv\x82\x99\xee\x8d\xddU:;\x8ai\x97n\xc6\xb19\xaf\x8f\xb1;\xaa\x86p9wr\xa8\x173\xa2\x96[\x8cYt:$\xbb_\xed\xb4.\xa4\xab\x9f\xcf\r\x9e\xa1\xacs8\xa4W7\x92\xae\xfc\xd1Sn\xf9\xd3j_Jz\nU\x92\x8b\x03\xb6y\xfeDUO?\xd75\xcc\x89;\xca\xc8\xf99\xdf\xd6V\xed\xdeR\xb0c\x18\xed\xf7&\xb9\xb2\xdb`\xb5U\xdcj\x9e\xbb# \x8d]\xc2\x8b\xddr\xf65<*\xa8\x85\x8a\xd3N\x8b\x15\x12\xac5\xf9\xed\xcf\xce\xecT\xe2[u\x0b\x91_\xe0G\xca\xbc\x93\xc2\x8dE\xf09\x00\x82\xef7\xcc\xebR\xf2\x08Y]Wt\xc8.\x0e\xdf\xa8\xa7\x8d\x8a\xfe\x14\xee\xf0F\xc4\xe1jvn\xa8\x89\xe1&\r6\xe8\xb7}\xb9\xb5\x95\x99\xb5\xc1,\xd4\xeej*RR\xabe\xa9]\xfb\x8a\xeelg\xfe?\x12\x16\x8f\r\xc3\xf1\xac>\xda\xda\x0cr\xcfKo\x89\x1a\x8ds\xa3f\xf2I\xb7u\xef_d\xf5\xf0\xb9T\xef@\xf2\xd8\x1e\x9e\xe1\xf8=?W\x8d\xd9)\xe9%s\x11\xb2T\xaaq\xcf"\'\xe7H\xed\xdd\xb7wn\xce\xf2\x1e\xa4\x00\x00\x00\x00\x00?\x99^\xc8\xa3t\xb2=\xacc\x11\\\xe79vDD\xedUS2\xaf\xb3\xd5\xeb\xefJ\'GL\xf9V\x9a\xf5uH r\'8\xa8c\xe5\xc4\x89\xdcT\x85\x8a\xf5N\xfe\xfd\xf2\xe7\xf4\xce\xce=%he\xd5\x94\xd3uw\x1b\xda\xfa\x17K\xb2\xf3D\x91\x17\xadw\x7f\x94h\xf4\xdf\xb8\xaa\xd2\x10\xf38\xb0~\xbe\xe9~\xd4*\xb8we3R\xd9B\xe5N]c\xb6|\xceO\n7\xabO\x13\xdc\x05\xd1\xa1\xa5\xa7\xa1\xa2\x82\x8a\x92\x16\xc3OO\x1bb\x8a6\xa6\xc8\xc656DO\x02""\x1f`\x00\x00\x00\x00\x00\x00\x00\x1f*\xcah+)&\xa4\xaa\x89\xb3A<n\x8eX\xdc\x9b\xa3\xda\xe4\xd9Q|\n\x8a}@\x19#\xab\x98\x8c\xf8&\xa5_\xf19\xd1\xdb[\xab\x1f\x1c.wk\xe1_e\x13\xfe\x969\xab\xf4\x9eT\xb7\xbeh\xee\x11\xe7k\xd5\x87P)"\xda:\xc8\xd6\xdb\\\xe4N]c7|N_\x0b\x9b\xc6\x9e(\xd0\xa8@\x0b\x8b\xe6qg\x1d]e\xfbO+&\xf62\xa7\xa2t\rU\xfc\xa4\xd9\x935<*\x9dZ\xed\xfc\xd7)N\x8fY\xa3\xf9\x84\xf8\x0e\xa6Xr\xd88\xd5\xb6\xfa\xb6\xbev7\xb6H]\xecebx\xd8\xe7\'\xd2\x06\xb6\x03\xe5GS\x05e$5t\xb2\xb6h\'\x8d\xb2E#Wt{\\\x9b\xa2\xa7\x81QO\xa8\x00\x00\x00\x00\x1dvQ\xeej\xe9\xf29\xbfaLx6\x1f(\xf75t\xf9\x1c\xdf\xb0\xa6<\x00\x00\x01\xa8=\x10\xbf\x8b~\x1b\xf2I<\xbc\x84\xaeE\x1d\x10\xbf\x8b~\x1b\xf2I<\xbc\x84\xae\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x06hj\x17\xbb\xec\x8b\xe7Z\xaf,\xe2\xedtJ\xf7\x80\xc7~5_\xda\xa6)6\xa1\xfb\xbf\xc8\xbeu\xaa\xf2\xce.\xcfD\xafx\x0cw\xe3U\xfd\xaa`%P\x00\x00\x00\x14\x1f\xa5\x87\xbf\xf6K\xff\x00\x0b\xf6XOw\xd0G\xdd\xaeC\xf3s|\xa3O\x07\xd2\xc3\xdf\xfb%\xff\x00\x85\xfb,\'\xbb\xe8$\xadL\xdb j\xaa"\xad\xb5\xaa\x89\xbf\xff\x00q\xbf\xbd\x00\xb7\xc4\x03\xd3\xa3\xde\xaa\xd3\xf3\xe4~BrE\xd4mW\xc208\xd5\xb7\xbb\xbb\x1f[\xcf\x86\x86\x97ij\x17\x97u\xa8\xbe\xc7\xc6\xe5jxJ\x8b\xae\xfa\xd5r\xd4\xf8\xa9\xadm\xb6El\xb3\xd2\xcf\xe7\x88\xe1\xe2\xeb%\x92M\x9c\xd4s\x9d\xb2"l\xd7/\xb1D\xee\xaf5\xe5\xb0F\xb8\xf5\xcek-\xfe\xddy\xa7c$\x9a\x82\xae*\xa8\xda\xfd\xf8\\\xe8\xde\x8eD]\xbb\x9b\xa17\xea\x0fJ\x0c\xb6\xf7H\xb4x\xcd\x048\xe4Oj#\xe7I:\xfa\x8d\xfb\xbc.V\xa3Z\x9d\xcfj\xab\xddEE"\xdc\x03N\xf3\x0c\xea\xa7\xa9\xc7,\xb3\xd4\xc6\x9e\xde\xa5\xff\x00\xc1\xc0\xce{s\x91\xdc\xb7\xf0&\xeb\xe0,\xb6\x9a\xf4\\\xb0[\x1b\x1dnk\\\xeb\xcdV\xc8\xbet\xa7U\x8a\x99\x8b\xddEw\'\xbf\xff\x00\nw\xd1@\xac\x98\xf67\x9aj%\xf6\xa1\xd6\xaa\x0b\x8d\xf2\xbeW\xa3\xaaj\x1c\xe5v\xce^\xc5\x92W.\xcd\xf1\xb9S\xb0\xb1zk\xd1b\x8a\x99\xd1\xd7gwE\xab\x91\x1c\x8e\xf3\x85\x0b\x95\xb1m\xde|\x8a\x88\xe5\xfe\xaa7\xc6\xa5\x90\xb4\xdbm\xd6\x9a\x18\xe8mT\x14\xb44\xb1\xa6\xcc\x86\x9e&\xc6\xc6\xf8\x9a\xd4D9@u\x98\xd6?d\xc6\xadl\xb5\xd8-t\x96\xda6\xb9\\\x91S\xc6\x8dEr\xf6\xb9{\xee^\xea\xae\xea\xa7f\x00\x00\x00\x00\x00\x00\x00\x00\x0e\x93=\xc9(\xb0\xfc.\xf1\x94\\U<\xedl\xa3\x92\xa5\xed\xdfez\xb5\xbb\xa3\x13\xc2\xe5\xd9\xa9\xe1T\x02\x89\xf9\xa09\xc7\xa6-[\x83\x16\xa4\x9b\x8e\x8b\x1c\xa7\xea\xde\x88\xbb\xa2\xd4\xca\x88\xf9\x17\xe8oV\xdf\x02\xb5\xc5\xc8\xe8\xf3\x867\x01\xd1\xccw\x1ct]][)Rz\xde\\\xfc\xf1/\xb3\x91\x17\xbf\xb2\xbb\x85<\rC<4f\x82\xafRzEXYvrT\xcdu\xbd\xf9\xfa\xbft\xe5#Z\xe5\x9ed\xfaZ\xd7\x1a\x9c\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x04{\xd27\tM@\xd1\xbc\x83\x1d\x8a.\xb2\xb5i\xd6\xa6\x87d\xe7\xe7\x88\xbd\x9b\x11;\xdcJ\x9c\x1e\')\x95+\xc9vSf\x0c\xbe\xe9e\x84\xfaE\xd7+\xed\xbe\x08z\xba\n\xf9=\x12\xa1DM\x93\xaa\x99UU\xa9\xe0k\xd1\xecO\x8a\x04P\x00\x03Hz\x0eg\x1e\x9b\xb4B\x8e\xdbS7\x1d\xc3\x1d\x93\xd0\xe9QW\x9a\xc4\x89\xbc.\xf1p/\x07\xfb\xb5\'s:z\x07g\x1e\x95\xb5\xa2;\x1dT\xdc\x14\x19$>sr*\xec\xd4\x9d\xbb\xba\x17x\xf7\xe2b\x7fHh\xb0\x00\x00\x00\x00\x1dvQ\xeej\xe9\xf29\xbfaLx6\x1f(\xf75t\xf9\x1c\xdf\xb0\xa6<\x00\x00\x01\xa8=\x10\xbf\x8b~\x1b\xf2I<\xbc\x84\xaeE\x1d\x10\xbf\x8b~\x1b\xf2I<\xbc\x84\xae\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x06h\xea\x17\xbb\xfc\x8b\xe7Z\xaf,\xe2\xectJ\xf7\x80\xc7~5_\xda\xa6)>\xa1\xfb\xbf\xc8\xbeu\xaa\xf2\xae.\xc7D\xafx\x0cw\xe3U\xfd\xaa`%`\x02\xaa5\x15UQ\x11;U@\x05\xe5\xcdH\x8bSzA`\xb8\x83f\xa4\xa2\xaa\xf4~\xea\xc6\xfb\x1az\'"\xc4\xd7w\x11\xf2\xfbT\xf0\xa3x\x95;\xc5]\xd4\xedo\xce\xb3\x95\x9a\x9az\xf5\xb5\xda\x9f\xff\x00\xc0\xd0\xaa\xb1\x8a\xde\xf3\xdd\xed\x9f\xe1E^\x1f\x02\x01\xf9\xd2\x86\xb6\x92\xe1\xae\xd9-E\x15DU\x10\xf5\x90G\xd6F\xe4sx\x99O\x1b\x1c\x9b\xa7u\x1c\xd5E\xf0\xa2\x9e\x06\xcbv\xba\xd9k|\xfbg\xb8\xd5\xdb\xea\xb8\x15\x9du,\xce\x8d\xfc+\xda\x9cMT]\x97d\xfe\xc3\xdf\xe9\xb6\x87g\xb9\xb7WQ\x05\xb5mv\xd5r"\xd6W\xa2\xc4\xd5j\xf3Ucv\xe2\x7f.\xc5D\xe1^\xce$,\xf6\x9a\xf4x\xc1q5\x8e\xae\xe3\x02\xe4W&/\x12MZ\xc4\xea\x98\xbf\xcd\x8b\x9b\x7f\xefq.\xfd\x8a\x05Q\xd3\xbd(\xce\xf5\x06vTZ\xadr\xb6\x86W\xaa\xbe\xe5X\xab\x1c\x1b\xef\xec\x97\x89y\xbdw\xfc\xd4r\xef\xdaY\x8d4\xe8\xcf\x88c\xfdUnO+\xb2*\xf6\xa6\xeb\x1b\xdb\xc1J\xd5\xf03}\xdf\xb7g\xb2UE\xfc\xd4\'X\xd8\xc8\xd8\xd6F\xd6\xb1\x8dM\x91\xadM\x91\x10\xfe\x80\xf9Q\xd3SQ\xd2\xc5IGO\x15=<-FG\x14LF1\x8dN\xc4DNH\x9e\x04>\xa0\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00*G\x9a\'\xa8l\xa3\xc7m\x9aoo\xa8\xff\x00\xb4\xdc\x1e\xda\xeb\x93Z\xbe\xd6\x06/\xf0L_\x8c\xf4\xe2\xff\x00v\x9d\xf2\xc7\xea\x8esa\xd3\xac.\xb7(\xc8*[\x1d=;U"\x89\x1c\x88\xfa\x99U\x17\x86&\'u\xceT\xfa\x13u]\x91\x15L\xb0\xd4l\xba\xed\x9d\xe6\xd7L\xb2\xf6\xf4um\xc2e\x91\xcdo\xb5\x8d\xa8\x9b26\xff\x005\xadDjx\x10\tw\xa054S\xf4\x88\xa2\x96DEu5\xba\xaaX\xf7\xee9X\x8c\xe5\xf4=M\x1a2\xfb\xa2VWK\x87\xeb\xe67q\xb8L\xd8hje}\rC\xdc\xbb5\xa93\x15\x8dr\xaa\xf6"=X\xaa\xbd\xe4SP@\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x15?\xcd\x1c\xc3\xd2\xbb\x0b\xb1f\xd4\xf1\xa2\xcdk\xaaZ*\x95D\xe6\xb0\xcc\x9b\xb5W\xc0\xd7\xb7o\xf7\x85\xb0"N\x98\x94\xf4\xf5=\x1b\xb3\x06\xd4\xa7\xb1e<25{\xcfl\xf1\xab\x7f\xc5\x10\x0c\xc2\x00\x01\xc8\xb6V\xd5[.T\xb7*\x19\x9d\x05],\xcc\x9e\t[\xda\xc7\xb5\xc8\xe6\xb9<(\xa8\x8ak^\x96e\xb4y\xde\x9eY2\xda%oWr\xa4l\xafcWt\x8eT\xe5#?\xaa\xf4s~\x83#\x0bs\xe6}j\xa46\xcb\x9dV\x98\xde\xaa\x9b\x1d=|\x8bSh|\x8e\xd9\x12}\x91\x1f\x0e\xeb\xf9\xe8\x88\xe6\xa7}\xaeNj\xe4\x02\xef\x80\x00\x00\x00\xe9s\xca\x86\xd2`\xf7\xea\xa7\xae\xcd\x86\xdbQ"\xaf\x81"r\x99\x02j_J|\x86\x1co@2\xfa\xc9eF>\xa6\xde\xfa\x08S}\x95\xcf\x9f\xf8$D\xf0\xa2=W\xc4\x8a\xbd\xc3-\x00\x00\x00\xd3~\x86\x15-\xa9\xe8\xd3\x88\xb9\xab\xcd\x91\xd4\xc6\xe4\xef+j\xa5O\xfd\t\x84\xac\x1egVQOq\xd2\xbb\xa6,\xf9\x93\xcf\x96\x8b\x8b\xa5lj\xbc\xfa\x89\x9a\x8a\xd5D\xf8\xed\x93\x7f\xa3\xbeY\xf0\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x003GQ=\xdf\xe4_:\xd5yW\x17_\xa2c\xda\xde\x8f\xd8\xf3\x9e\xe4kQ\xd5j\xaa\xab\xb2\'\xfd\xaab\x94j7\xbe\x0eG\xf3\xadW\x96q\xf2\xa0\xbc\xe4\xd3\xdaY\x8bP\xdc\xae\xb2\xd0O6\xec\xb6\xc33\xd69$^\xf4h\xbb*\xaa\xf76\xed\x02\xe6joH\xbc#\x14W\xd1\xda$\xf4\xc9rk\x95\xab\x1d$\x88\x90F\xa9\xf9\xd2\xec\xa9\xdb\xcb\xd8\xa3\xbc;\x15\x7fRu\x8f;\xcf\xb8\xe9\xae7%\xa4\xb7\xbfv\xf9\xc2\x87x\xa2r/q\xdc\xd5\xcf\xfe\xb2\xaaw\x91\x0f[\xa6}\x1b3,\x91\xd0\xd6\xe4Nn9nrq*L\xde:\xa7\xa7sh\xf7\xf6;\xff\x009QS\xf3T\xb3\x9ao\xa48.\x08\xc8e\xb4\xda\x19Qq\x8d\x9c.\xb8U\xed,\xee^\xea\xa2\xf63\xfa\x88\xd0*\x8e\x9a\xf4z\xce\xf2\xe6\xb2\xae\xba\x04\xc7\xad\xcej9\xb3\xd7\xc6\xbdc\xd1\x7f6.N_\xebp\xa7yK;\xa6\xba\x19\x81a=]LV\xefE\xaem\xe1_>W\xa2H\xadrwX\xcfj\xce}\xd4N/\n\x92x\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00C\x1a\xc3\xd2KN4\xfa\x9ezh\xaeQ\xe47\xb6n\xd6\xdb\xed\xd2\xa3\xf8^\x9d\xc9eM\xdb\x1e\xcb\xc9S\x9b\x93\xf3T\x99\xcaC]\xd0\xbf0\xba^\xee\x17\n\xbc\xc2\xc5J\xca\x9a\xa9&kb\x8aYU\x11\xceUM\xf7F\xf3\xe6\x04\x01\xac\xda\xab\x96j\xaeE\xe8\xaeIV\x89\x04;\xa5\x1d\x04;\xa4\x14\xad^\xd4jw\\\xbb&\xee]\xd5v\xef""xB\xdf\xa7A\xcb\xb6\xdc\xf5\x12\x8b\x7f\x9a\xdd\xf8\x83\xd69u\xfeQh\xbe\xabw\xe2\x01P\x0bW\xa0}.\xae8\xd5\xb6\x9b\x1d\xd4J*\xab\xcd\x04\rH\xe0\xb9\xd3\xaa-TlNH\x925\xca\x89*\'\xe7n\x8e\xd99\xf1)\xdaz\xc7.\xbf\xca-\x17\xd5n\xfc@\xbd\x07n\xdbr\xd4J%_\r\xad\xdf\x88\x05\x84\xc5\xbaDh\xd6B\xc6\xad6sn\xa2\x91{c\xb8\xf1R+W\xbc\xab*#W\xe8UC\xdeYr\xbcZ\xf73a\xb3d\xb6k\x94\xaeEsYI]\x14\xceTN\xea#\\\xa54\x9f\xa1\x06N\x88\xbdFugz\xf78\xe9$o\xf9*\x9e\xd7\xa3\x9fF\\\xc7L\xb5n\x83,\xba_,5\xb6\xfax\'\x8d\xec\xa6|\xbdj\xab\xe3V\xa6\xc8\xe8\xd1;W\x9f\xb2\x02\xd7\x80\x00\x00\x00\x00\x00\x1d&C\x98bX\xe4\x9d^A\x94Y-\x0f\xe1\xe2\xe0\xae\xaf\x8a\x07m\xdf\xd9\xeeC\xbb*\xe7I\x9e\x8d\xd9f\xaaj\x97\xa6{U\xf2\xc9AC\xe7\x08i\xb8j]*\xcb\xc4\xc5r\xaa\xec\xd6*m\xec\x93\xba\x04\x8d\x93\xf4\x93\xd1\x8b\x0b\x1f\xd6\xe6t\xd7\t\x9a\x9b\xa4V\xe8d\xa9W\xf8\x11\xcdN\x0f\xedr\x15;\xa5\x0fIY\xf5B\xd4\x98\xae5n\xa8\xb5\xe3\x8b+e\xa8uJ\xa7\x9e*\xdc\xd5\xdd\xa8\xe4j\xabZ\xc4]\x97\x87uUTE\xdd6\xd8\xf5\x90\xf4 \xc9\x15\x13\xae\xce\xed,\xef\xf0Q\xc8\xef\xf3T9)\xd0r\xed\xb7=D\xa2\xdf\xe6\xb7~ \x15\x00\x16\xff\x00\xd69u\xfeQh\xbe\xabw\xe2\x0fX\xe5\xd7\xf9E\xa2\xfa\xad\xdf\x88\x05@>\x94\xb3\xcfKS\x15M4\xd2A</I"\x927+\\\xc7"\xee\x8eENh\xa8\xbc\xf7-\xdf\xacr\xeb\xfc\xa2\xd1}V\xef\xc4\x1e\xb1\xcb\xaf\xf2\x89E\xf5[\xbf\x10\x0e\xe3B:`\xdb\x9do\xa5\xb1j\x94sS\xd5D\xd4b^\xe0\x8ddd\xbbv,\xb1\xb58\x9a\xee\xfb\x9a\x8a\x8a\xbd\xc6\x96\x87\r\xcd\xb1\x1c\xca\x99\xd5\x18\xaeIk\xbc1\x88\x8b"R\xd4\xb5\xef\x8f~\xce6\xa2\xf17\xe9D)\xec\xfd\x082F\xa2\xf5\x19\xdd\xa6E\xeeq\xd1\xc8\xdf\xf2U%\x8e\x89\xfa\x0b\x93\xe9\x06M}\xb8^\xee\xd6z\xfa{\x85$pC\xe77I\xc6\x8ek\xf8\x95\\\x8eb"&\xdd\xe5P&+\xee\xa2\xe9\xfd\x8ai`\xbc\xe6\xf8\xe5\x04\xf0\xb9[$3\xdc\xe1d\x8dr.\xca\x9c\n\xee-\xd3\xbd\xb1\xe0rn\x93\xfa/c\x85\xea\x99g\xa2\x937\xb2\x0b},\x92\xb9\xde\'*#?\xb5\xc8B9\xcfC\xfc\xdb\'\xd4<\x93"\xf4\xcf\x8fRR].\xd5U\xb0\xb3y\x9f#c\x96g=\xa8\xe4\xe0D\xe2\xd9\xc9\xbe\xca\xa9\xbfuN$=\x07\xf2\x05\xff\x00M\x9f[\x19\xf1h\x1e\xef\xfa\x90\x08\xc7\xa4\xe6\xbe\\\xf5~\xb6\x9a\xddIF\xfbV5C"\xcbOH\xf7\xa3\xa4\x9aM\x95:\xd9U9n\x88\xaa\x88\xd4\xe4\x9cK\xcdw!B\xdf\xfa\xc7.\xbf\xca-\x17\xd5n\xfcA\xeb\x1c\xba\xff\x00(\xb4_U\xbb\xf1\x00\xa8\x00\xb7\xfe\xb1\xcb\xaf\xf2\x8bE\xf5[\xbf\x10z\xc7.\xbf\xca-\x17\xd5n\xfc@+n\x94j\x0eG\xa6y\x84\x1965P\xc6T1\xab\x1c\xd0\xca\x8a\xb1T\xc4\xaa\x8a\xe8\xe4DT\xdd\xab\xb2/%EEDTTT.\xb6\x0b\xd3#M\xae\xd4\xd13(\xa4\xbacu\x9b\'Z\xab\n\xd4\xd3\xa2\xff\x005\xd1\xa7\x1a\xa7\x8d\x88G~\xb1\xcb\xaf\xf2\x8bE\xf5[\xbf\x10\xfc\x7fA\xdb\xc2\'\xb0\xd4:\x15_\r\xb1\xe9\xff\x00\xf6\x01c\xed:\xf9\xa3w4E\xa6\xd4;${\xf6y\xe6U\xa7\xf2\x88\xdd\x8f]b\xcc\xf0\xfb\xf4\xec\xa7\xb1\xe5v+\xa4\xcfEV\xc7Gp\x8ag9\x117]\x91\xaeU^\\\xcaq7B\x0c\x9d\x11z\x9c\xe6\xce\xf5\xeeq\xd2H\xdf\xf2U=wG\xfe\x8b\xd9\x96\x9b\xeb\x15\x970\xb8\xdf\xac\x15\xb6\xfa\x04\xa8Ic\xa7|\xc92\xf5\x90I\x1ap\xa3\xa3\xdb\xb5\xe9\xbe\xeeN[\x81m\x80\x00\x00\x00\x00\x00\x00\x00\x00\x00U\xe7\xf4^\xb9^3\xeb\xad\xda\xff\x00\x91RAi\xa9\xb8KQ\x1ctMs\xe7\x927\xbd\xce\xe1Ur#X\xbc\xd17\xf6}\xdeD\xe7\xa7\xdam\x86\xe0\x94\xe8\xccv\xcb\x0c\x13\xaa/\x1d\\\x9f\xc2T?~\xdd\xe4w4O\xe6\xa6\xc9\xe0=p\x00\x00\x00\x01\xc2\xbfS\\+,\xf5T\xb6\xab\x97\xa1u\xb2F\xad\x82\xaf\xa8l\xddK\xbf;\x81\xdc\x9d\xe2P9\xa0\xa7\xd3\xe7\xfa\xe7I\xd2V\r\x1c\xab\xd4*\x04l\xd2\'\x05\xce;\x14\x0b\xbck\x02\xcc\x8b\xd5\xafwd\xe1T\xe2\xed\xee\xa96Uaz\xcf\xd4\xbb\xce\xba\xd9K\xd6\xed\xecRLN\x9f\x85W\xc3\xb4\x9c\x80\x95A\t\xe85\xff\x00T\xd3Ss\x0c\'S\xeeTW\t\xad\xb4\xb4\xb5V\xf9\xe9)\x99\x14r\xc5#\xa4Ezp\xb5\x15QxQ6^\xc5j\xa1\xcf\xd7\xcdS\xba\xe2w\xac{\x04\xc3(\xa9+\xb3,\x96^\xae\x91*\xd5R\nH\xb7\xd9f\x91\x13\x9a\xf3\xdfdO\xcdr\xf76P\x97A\x0c\xcf\x88k\xfd\xba\x8dn\x94:\xbbk\xbd\xdcX\xde5\xb5V\xe3\xb0\xc1G*\xf6\xf5i,k\xd65;\x88\xee\xefwc\x9f\xd1\xb7Q2-D\xb4d\xf5y=\xaa;E}\xb2\xfb%\x07\x9c\x1b\xcdi\x9a\xc8b\xdd\x8eU\xe6\xe5\xe3Y\x17\x7f\x0e\xdd\x88\x04\xae\n\xd3\xa1St\x85\xd4|:\x0c\xb2\xe5\xa9\x16\xcb\x05\x05[\x9f\xe78\xd2\xc1\rD\xb2\xb1\xaeV\xab\xd57b5\x15Qv\xedU\xdb~\xea\x1ec\xa4\xf6}\xad\x9a1-\x81c\xd4z+\xdcW\x84\xa8\xdb\x8a\xc1\x04\x0e\x89a\xea\xf7\xee\xbb\x89\x17\xadN\xf7`\x16\xf4\x10\xe5\x8f\x18\xd6\xbb\x969Eqv\xb4QEQUM\x1c\xfdZb\x909\x8cW5\x1d\xc3\xbfX\x8a\xbbo\xb6\xfb\'\x88\xe94\xae\xef\xacu\x99\xeeq\xa7y\xaeUE\x1d\xca\x86\x86\x9a\xa6\xd5s\xa4\xb6\xc4\xacFH\xf7o"1Q8\xb7D\xe1Twb\xa2\xf8\xc0\x9f\x81Hu#Vu\xff\x00\x10\xd6\xd6i\x8cy\x8d\xb6\xba\xa2\xa2\xae\x9a\x9e\x8e\xab\xd0\x88cl\xa9?\n1U\xbc+\xc3\xb2\xbbeM\xd7\xb1K\x81\x83\xdb2KU\x9d\xd4\xf9FN\xdc\x8e\xb9\xd2\xf1\xa5Sh\x19H\x8do\x0bS\x81\x18\xc5T\xdbtr\xee\xab\xbf\xb2\xf0\x01\xdf\x02\xbct\xa9\xbd\xea\xf6\x9dcuY\xc61\x9cR\xc9im[#\x92\xdb5\x9e\x15u3\x1e\xbc-V\xca\xbb\xab\xfd\x96\xc9\xcd\x13\xdbv\x92.\x8eP\xea#\xb18.y\xcea\r\xce\xba\xe5C\x14\xcc\xa7\x86\xd9\x1c\r\xa1{\x9b\xc5\xb7\x13y\xc8\xa9\xba"\xee\x88\x9c\x80\x90\x81O\xb5\xb7?\xd7=9\xd5[\x06\x18\xcdA\xb7\xdc)\xef\xbdO\x9d\xeb\x1dc\x827G\xc7/T\xa8\xe6s\xdf\x85v^K\xcd\x17\xb8LO\xc3\xb5\xe66\xab\xe1\xd6{=C\xd3\xb29qX\x98\xd7x\xd5\xb2*\xa0\x12\xf8+\xd5F\xb6\xe6Zc\x93Q\xd85\xc7\x1e\xa1\x86\xdf\\\xfe\nL\x92\xcb\xc6\xeaW\xed\xff\x00\xd4\x8d\xdb\xb9\x159o\xb6\xca\x9d\xa8\xd5Ne\x80\xa3\xa9\xa7\xac\xa4\x86\xae\x92x\xea)\xe7\x8d\xb2E,nG1\xecrn\x8eENJ\x8a\x8b\xbe\xe0}AM\xfaIj6\xbbi\x16Wk\xb5\xc1\x9c\xd0\xdei.\xd1,\x94\x92\xad\x96\x08\x9e\x8ek\x91\xaec\x9b\xb2\xa7.&\xf3\xdf\x9e\xfd\x88Z\r7\xb5\xe66\xbb\x1f\x06o\x94\xc1\x90\\\xe5V\xbd_\x05\x03)\xa3\x83\x9765\x1b\xed\x93}\xfd\x92\xec\xab\xde@=@"-w\xd7+N\x9cV\xd2cv\xbbl\xd9&cqV\xb6\x8e\xd3L\xbc\xd3\x89vk\xa4TET\xdd{\x1a\x88\xaa\xbe\x04\xe6u\x16\xccc\xa4\x8eIL\xdb\x85\xf3Sl\x98T\x92\xa7\x1a[mv8\xabz\xae\xf3\\\xf9W\xb5;\xbb9\xc9\xe1\x02t\x05t\xcbr\xedv\xd1\xba\x7fG2\xdfBu\x17\x13\x8d\xc8\x95u\x94T\xa9E[J\x8a\xbbq\xb9\x8d\xf6\x1c?B\xa7}[\xdaM\x1aq\x9a\xe3\xba\x81\x89\xd2\xe4\xd8\xc5jUP\xd4n\xd5EN\x19"z{h\xde\xdf\xc9rw\xbcJ\x9b\xa2\xa2\xa8z0\x0f7\xa86\xac\xb2\xebglX~U\x169pc\x95\xddt\xb6\xf6U\xb2T\xd9vb\xb5\xca\x9c)\xbe\xcb\xbas\x03\xd2\x02\x9d\xf4u\xd4=u\xd5|\xce\xf3c\xaa\xcf(,\xf0Y\xa3\xe2\xaa\x95\x96X&s\x9f\xc6\xacF56jw\x1c\xbb\xaa\xf7;9\x93WH\xb7\xeaU\x8f\x10\xbbf8FeOo\x8a\xd3@\xb5\x12\xdb\'\xb5\xc52L\x8c\xdd\xd2=%w6\xaf\x0fseOc\xdc\xdfp%\xa0T\xee\x8a\xd9\xce\xb5k\x04\xd7K\x85\xc3<\xa3\xb6[-R\xc2\xc7\xb6;,\x12>\xa1\xce\xddU\xa8\xbb\'\n"7\x9a\xf3_d\x85\x9f\xca(\xee\xb7\x0b\x1dE%\x92\xf3\xe8-\xc2N\x1e\xa6\xb7\xce\xcd\xa8\xea\xb6r*\xff\x00\x06\xeeN\xdd\x11S\x9ff\xfb\xf7\x00\xec\x81U1\xcc\x87\xa4\x05\xcf\xa45\xdbJ\xaa5\x02\xdd\x0c\x16\xaao>\xcdse\x8e\x07u\x90*F\xacV\xc7\xdcr\xac\xadEN.[;\xb7ns6\xadPj,\x18B\xdc1\x0c\xde\x0b}\xca\xd5n\x92Z\x94\x9e\xd5\x14\xac\xb8H\xc6"\xee\xbb\xaf\xf0[\xf0\xbb\xda\xa2\xa7\xb2\xec\xe4\x04\x8a\n\x9d\xd1\xb3.\xd6\xedd\xb1]\xae\xbe\xa9\xf4\x164\xb7\xd56\x9f\xab\xf4\xb7\x05GY\xbbx\xb7\xdf\x89\x9b\x7f\x89\xddj\xae\xa2kF\x87z\x1fy\xcag\xc7\xf3\x8cb\xaa\xa1)\xa5\x9e\n7P\xd4\xc4\xf5Er&\xcds\x9a\x9b\xa3]\xb2\xec\xe4\xe5\xb2\xed\xcbp\xb2\xe7\x99\xd4\x0c\xca\x93\x10\xb76wZoW\x9a\xd9\xb7Jj\x0bM\x0c\x953L\xef\xea\xa7\x0b\x13\xf9\xceTC\x93\x80ev\x8c\xdf\r\xb6eV)_%\xbe\xe3\x0f[\x17\x1al\xe6.\xea\x8ec\x93\xb8\xe6\xb9\x15\xab\xe1E+\xff\x00Kl\xbfX\xb4\xaa\x9e\x9f)\xb0\xe6\xb4s\xd8\xee\x15\xebJ\xca9-\x10\xa3\xe9\x1c\xads\xd8\xde5\xdf\xacM\x98\xeek\xb2\xf2\xee\xef\xc8%]\x1f\xc9\xb5G%\xaf\xb9U\xe6\xf8%\x1e)h\xe1j\xdbcu_[V\xf5\xdf\x9fX\x88\xbb"m\xdfF\xae\xfd\xc5\xedI \x8b::?Q\xae\xf8]\xb7-\xce\xf2\xcak\x9a^h#\xaa\xa7\xa0\x82\xdb\x1c\tN\xd9\x11\x1e\xc7,\x8d\xe6\xf5V*n\x9b"&\xfd\xdd\xb7"\xbe\x969\xb6\xb2i#\xa8\xef\x96|\xea\x8e\xb2\xd1u\xac\x96(\xa9d\xb3B\xd7\xd2~S\x19\xc5\xcf\xacN\x1d\xd3\x89v_c\xdd\xdc\x0bN\x08O\x1f\xc6\xf5\xde\xeb\x8f\xdb\xee\xcd\xd6kTn\xac\xa5\x8a\xa1!v)\n\xa3x\xd8\x8e\xe1\xe2\xeb9\xed\xbe\xdb\xect\xb9>\xa7\xea\xc6\x8f\xcf\x05V\xa8X\xadY.+,\xad\x89\xd7\xcb\x13\x1d\x14\xb4\xca\xab\xb2u\xb1=v\xe7\xdc\xdbd\xe7\xb7\x12\xaf ,0:\xbcO!\xb3ex\xf5\x1eA\x8f\xd7\xc5_m\xac\x8f\xac\x82x\xfb\x1c\x9d\x8a\x8a\x8b\xcd\x15\x17tT^h\xa8\xa8\xa5v\xe9m\x97\xeb\x16\x95S\xd3\xe56\x1c\xd6\x8e{\x1d\xc2\xbdiYG%\xa2\x14}#\x95\xae{\x1b\xc6\xbb\xf5\x89\xb3\x1d\xcdv^]\xdd\xf9\x05\x9d\x04Y\xd1\xd1\xfa\x8dw\xc2\xed\xb9nw\x96S\\\xd2\xf3A\x1dU=\x04\x16\xd8\xe0Jv\xc8\x88\xf69do7\xaa\xb1St\xd9\x117\xee\xed\xb9\xfb\xd2\x0fZ\xb1\xfd"\xb2@\xfa\xb8\x1ds\xbd\xd7"\xf9\xc2\xd9\x13\xf8]&\xdc\x95\xefv\xcb\xc0\xc4^[\xec\xaa\xab\xc9\x11y\xec\x12\x90 LZ\xcf\xd2;4\xa0\x8a\xf9}\xcfm\x9a\x7f\x15J$\xb0\xda\xa8lqU\xcb\x1b\x17\xb1$Y\x97v\xbbn\xd4\xdd|(\x8b\xc98\xd9\xe6W\xad\xda;e\x9e\xf9~\x92\xd3\xa8\x98\xeclT\x9a\xb6\x9e\x93\xce5\xb4\x8e^M{\xe3f\xect{\xaao\xb2x\xd5\xa9\xcc\x0b\x08\x0f-\xa4W\xba\xdc\x93J\xf1\\\x82\xe5#e\xae\xb8Z)\xaaj^\xd6\xa3Q\xd2\xbe&\xab\xd5\x119\'\xb2U\xe4\x84{\xac\xfa\xef\x16-\x95S\xe9\xf6\x0fevW\x9c\xd5\xb9\x18\xca&;hi\x95StY\\\x9d\xdd\xbd\x92\xb5\x156o79\xa9\xda\x13X \xea,3\xa4]\xe2\x9d\xb5\xb7\x8db\xb4cu/N/8[1\xd8j\xa2\x8dW\xf2VIU\x1c\xbbvw|jt96\xa4k\x0e\x8bTAW\xa9\x946\xdc\xcf\x11\x9aV\xc4\xfb\xd5\xa6\x0f;\xd4\xd3*\xae\xc9\xd6G\xed9\xf7\x13dE^\\{\xf2\x02\xc7\x83\xaa\xc4r+6Y\x8eQd8\xfdtu\xd6\xda\xd8\xfa\xc8&gu;\x15\x15\x17\x9a*.\xe8\xa8\xbc\xd1QQN\xd4\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\n_\x9e\\(-^h\xa5\xb2\xe1s\xad\xa6\xa1\xa3\x868\x96Z\x8a\x89[\x1clE\xa1r&\xeer\xa2\'5D\xe7\xdf\'\x1c\xa7]qF\xe7x\x96\x1d\x88_m\x17\xdb\x8d\xe2\xec\xcaz\xcf;I\xd7\xc7\x057\x0b\xb8\x97\x8d\x8b\xc2\x8f\xe2\xe1DM\xd7\xf2\xb7N\xc2\x0f\xce\xa8h\xae^h\xb5\xb2\x86\xe3GOYK,q$\x90TD\x92F\xf4\xf3\x8b\x97\x9b\\\x8a\x8b\xcd\x11I\xc7,\xd0\xdcVL\xef\x12\xccq\x1b\x15\xa6\xc7r\xb4]\x99QX\xb4\xd1\xa4\x11\xd4Sp\xbb\x898\x18\x9c*\xfe.\x05E\xd9?+u\x02WJ:4\xafu\xc1)`J\xc7D\x90\xba\xa3\xabN\xb1cEUF+\xbbxQUWn\xcd\xd5J\x81\xd3\x96\xdd\x97bz\xad\x89\xeb\x16?\x0b\xe4\xa6\xb6\xd3\xc5L\xe9x\x15\xec\x82h\xe5\x91\xe8\x92\'q\x8fI\x15\xbfB\xa6\xe8\xaa\x85\xc3|\xd12X\xe2|\xacl\x92*\xa4mW"+\xf6M\xd7d\xee\xf2:\xca[\x9d\x87"}\xde\xcf\x1c\xb4\xd7\x0f8\xcd\xe7;\x954\x91\xf15\x8es\x1a\xfe\x075\xc9\xb2\xa2\xb5\xed^\xea.\xfe0!\x8d\x1d\xe9Q\xa7\x99\xb4t\xf4\x17\xda\x84\xc5\xafoDk\xa1\xad\x7f\xfd\x9aG\xff\x002n\xc4O\x03\xf8W\xb9\xcc\x9al\x96+-\xae\xb6\xe7r\xb5Q\xc5\x04\xd7y\xdbU[$nUI\xe4F5\x88\xfd\xb7\xd9\x17\x85\xad\xec\xdb~\xde\xd2\xbek\x0fD<+$\x8a\xa2\xe3\x84H\xb8\xc5\xd9Q\\\xd8\x13w\xd1H\xee\xf2\xb3\xb6=\xfb\xed]\x93\xf3T\xf1=\x08sL\xca\xc1\xa9\x97]\x19\xca\xa4\x9eX(\xe2\x9b\xa8\x86g\xf1\xad\x14\xd0\xb9\x11\xccc\xbf\xfanEU\xdb\xb3tEN\xd5\xdc.\x05\x8e\xd5n\xb1\xdai\xad6\x8a8h\xa8)Y\xd5\xc1\x04M\xd9\x91\xb7\xbc\x88T\x1f4\xd3\xda\xe9\xff\x00\x8e\xe5\xff\x00\xe2\x97,\xa6\x9ei\xa7\xb5\xd3\xff\x00\x1d\xcb\xff\x00\xc5\x02\xdaa^\xe3l\x9f7\xc1\xe4\xdar[h\xb66\xfd%\xf9\xb40\xa5\xd2Jf\xd2>\xa9\x1b\xfc#\xa1k\x95\xedb\xafy\x1c\xe7/\xd2\xa7\x1b\n\xf7\x1bd\xf9\xbe\x0f&\xd3\xb7\x02\x8d\xeb\xf2\'\xaf\xe7\x10\xe5\xff\x00\xccl\xcb\xff\x00\x9c\xc2\xf2\x14s\xa4/\xf0\x1d=0\xe9e^\x16:\xbe\xcc\xe4U\xec\xdb\xafjo\xfd\xa8\xa5\xe3\x02\x17\xe9\xba\xd4^\x8c\x99R\xaask\xa8\xd5?\\\x84\x97l\xa8\x8d\xb3\xd15\x13dJx\xd1?\xee\xa1\x10\xf4\xdcrz\xdarh\x93\x9c\x93ID\xc8\xda\x9d\xaew\x9e\xe1]\x93\xe8E&+|N\x82\x82\x9e\x17{h\xe2kW\xc6\x88\x88\x05?\xe9\xab\xfcc\xf4\xb7\xe3\xd3\xfd\xb1\x0b\x8eS\x8e\x9a\xbf\xc6?K~=?\xdb\x10\xb8\xe0F\xbd\'q:L\xc7Cr\x8buD-|\xd4\xd42W\xd29S\x9b&\x85\xab#U\x17\xb9\xbe\xca\xd5\xf09H\x9f\xcc\xee\xcd\xab/\x9a}w\xc4+\xe6t\xab\x8f\xcf\x1b\xa9\x1c\xe5\xddR\t\xb8\xd5\x19\xe2k\xd8\xff\x00\x129\x13\xb8\x84\xe7\xadWjk\x1e\x90\xe5\xb7J\xb7\xb5\x91\xc1g\xa9\xdbu\xf6\xcfX\xdc\xd67\xc6\xaeTO\xa4\xad\x9ef\xb5\x86\xae\x1b^_\x92\xcd\x1b\x9bKU-=\x1d;\x959=\xd1\xa3\xdf\'\xf6u\x8c\xfe\xd5\x03\xe5\xe6\x85{\xbb\xd3O\x8f?\x95\x80\xb79\x05\xca\x1b-\x86\xe1x\xa9EX(ie\xa9\x91\x13\xf3X\xc5r\xff\x00\x82\x15\x1b\xcd\n\xf7w\xa6\x9f\x1e\x7f+\x01ls;K\xaf\xd8}\xea\xc6\xd7\xa3\x1dq\xb7\xcfH\x8e^\xc4Y#s7\xff\x00\x10)\xb7A\xcay\xf5\x0f]r\xcdK\xc8\xd5*\xee\x14\x91u\xacW&\xe9\x1c\xd5\x0er"\xb7\xbc\x8d\x8d\x8fb\'q\x14\xbb\xe5&\xf3:k\x16\xc9\x9fg\x18}\xce5\xa5\xb9\xcb\x04/t\x12rs]M$\x8c\x91\xbbw\xd1e\xec\xf0)v@\xf8\\(\xe9n\x14\x15\x14\x15\xd0GQKS\x13\xa1\x9a)\x13v\xc8\xc7&\xcej\xa7u\x15\x15P\xa4\x1d\x10\xee\xd5zw\xd2\x7f#\xd2\xc5\x9d\xef\xb5\xd6\xd5U\xd21\x8e]\xd3\xad\xa6W\xba9<k\x1b\x1e\x8b\xdf\xe2N\xf2\x17\x99U\x117U\xd9\x10\xa2\x1d\x1d\xa9_\x9et\xdb\xbef6\xd6\xac\x96\xaa\n\xeb\x85z\xcc\xdfj\xb1\xbf\xac\x8a.}\xf7q\xa2\xed\xdeE\xef\x01{\xc0\x00S\x8e\x80\x9e\xfaz\x9b\xf1\xd9\xe5\xe5,\x96\xbd&\xfa\x1d\x9d\xef\xf0r\xe1\xf6w\x95\xb7\xa0\'\xbe\x9e\xa6\xfcvyyK+\xae\x91\xba]\x13\xce\xa3bn\xe7c\xb7\x04D\xef\xaf\x9d\xe4\x02\x01\xf36=\xc1ek\xff\x00\xf2\x91\xf9$-\x81S\xbc\xcdw\xb5p\\\xb24T\xe2m\xce%T\xef"\xc5\xcb\xfc\x94\xb6 A\xb8\xa3\x1a\x9d5\xb3\'"s\\V\x93\x9f\xfb\xc6~\xe4%\xbc\xdf\xdce\xf3\xe6\xea\x8f&\xe2(\xc3\xd8\xb2\xf4\xcb\xcegg6\xc1\x8dQE"\xa7\xe4\xb9\xcek\x91\x17\xe8M\xc9_7\xf7\x19|\xf9\xba\xa3\xc9\xb8\n\xb1\xe6u^m\x16\xcc\x13)e\xca\xebCD\xe7\\\xe3sR\xa2\xa1\x91\xaa\xa7T\x9c\xd3\x89P\xfbt\xe8\xd4\xbc[!\xc2\xa8t\xf7\x14\xb9S\xe47\xba\xcb\x94R\xc9\x15\xb9\xe9P\x915\x88\xed\x9b\xbb7E{\x9c\xe6\xa259\xed\xbe\xfbr\xdf\xc8\xf4*\xd2\xac\x1fR\xb4\xd7,\x8b*\xb2\xc5UP\xda\xd6\xc3OX\xd7+\'\xa7E\x8b}\xd8\xe4\xef/=\x97t\xef\xa2\x9f\x9a/v^\x8d\xba\xe1[\x81\xe7\xf6\xea\x05\xb5\xdcdE\xa3\xbe\xad+RH\xda\xefb\xc9RM\xb8\xba\xa7m\xc2\xf6\xef\xb3\x1c\x8a\xbd\xc7n\x16k\xa2\xee\x1ds\xc14?\x1f\xc7\xefLX\xaeMd\x95\x151*\xef\xd5:Y\x1d\'\x07\x8d\xa8\xe4E\xf0\xa2\x91\xaf\x9a5\xef\x1di\xff\x00h\xe0\xfb=Ie\x98\xe6\xbd\x88\xf69\x1c\xd7&\xe8\xa8\xbb\xa2\xa1Z|\xd1\xafx\xebO\xfbG\x07\xd9\xea@\x99\xf4C\xde[\x06\xff\x00gm\xff\x00f\x8c\x81<\xd2_{\x8cc\xe7wy\x17\x13\xde\x88{\xcb`\xff\x00\xec\xed\xbf\xec\xd1\x90\'\x9aK\xefq\x8c|\xee\xef"\xe0,~\x9f{\x82\xc7\xbek\xa6\xf2M9\x19m\x86\xdd\x94c\x17,v\xed\x0bf\xa1\xb8\xd3>\x9ef\xaao\xc9\xc9\xb6\xe9\xe1N\xd4^\xe2\xa2)\xc7\xd3\xefpX\xf7\xcdt\xdeI\xa7sQ4T\xf0I<\xf26(\xa3j\xbd\xefr\xec\x8dj&\xea\xaa\xbd\xed\x80\xa5^g\xeeQr\xb2j\x0eK\xa5\xf7\t\x95\xf4\xca\xd9jaj\xaf(\xea!zG\'\x0f\xc6j\xee\xbf\xd1\xa1 \xf9\xa3^\xf1\xd6\x9f\xf6\x8e\x0f\xb3\xd4\x91OA\xda\x1a\x8c\x97\xa4\x86M\x98\xd3\xc4\xf4\xb7\xd3\xc3U;\xa4\xdb\x92>\xa2_\xe0\xd9\xe3V\xf1\xaf\xf5T\x95\xbc\xd1\xafx\xebO\xfbG\x07\xd9\xea@\x99\xf4C\xde[\x06\xff\x00gm\xff\x00f\x8c\xa7x\xadJ\xea\xd7O\'U]v\xa8\xb7\xdbn3\xac\x11;\x9bR\x1a4zB\x88\x9d\xe5{Z\xe5N\xfb\x9c\\M\x10\xf7\x96\xc1\xff\x00\xd9\xdb\x7f\xd9\xa3)\xa6\x831p\xce\x9d\x15\xf6k\x9f\xf0+5\xc6\xe3H\xc7;\x92/\x1a=\xf1/?\xceDn\xdf\x19\x00\xbfg\x16\xefo\xa4\xbbZ\xaa\xedu\xf0\xb6zJ\xc8\x1f\x04\xf1\xb99=\x8fj\xb5\xc8\xbe4U9G\xe3\x95\x1a\xd5s\x95\x11\x117U^\xc4\x03\xc1\xe4R\xd1i\x0e\x85W>\xd4\xb2\xcdM\x8dY\x9c\xda$\xa9r9\xcfV3h\xd1\xea\x88\x88\xbb\xbb\x85\x17dB\xba\xf9\x9d\xf65\xbc\xde3\x1dF\xbc=k.\xb2N\xdaVTK\xcd\xdcRo,\xee\xdf\xbe\xe5X\xf9\xf8\xfb\xe4\xed\x9fT[\xf5\x83\xa3\xa6C.\x1d;\xeb\xa0\xbb[\xaaY@\xee\xadZ\xb3K\x0b\xdc\xd4j"\xf3\xe7$J\x89\xe3\xdc\x85\xbc\xcd\x9b\xc5?\xa5\xdc\xbb\x1a\x91\xc8\xca\xc8+b\xac\xea\xdd\xc9\xca\xc7\xb3\x81Wo\x02\xc6\x9b\xfcd\x02\xdd\x1dNc`\xb7\xe5X\xad\xd3\x1c\xba\xc4\xd9h\xae4\xaf\xa7\x95\x157\xd9\x1c\x9bq\'\x85\x17eE\xee*"\x9d\xb1\xf0\xb8\xd6S[\xed\xf55\xf5\xb36\x1ajh\x9d4\xd29vF1\xa8\xaa\xe7/\x81\x11\x15@\xa5\xfeg\xb6Uq\xb4\xe6\xf9.\x99W\xca\xae\xa6\xea\xe4\xac\x85\x8a\xbb\xa4s\xc4\xf6\xc7"7\xe35Q\x7f\xdd\x97\\\xa2\xfd\x02-U\x99\x0e\xb8\xe4\xf9\xca@\xf8\xe8 \xa7\x9fw*r\xeb\xaa%G5\x9b\xfcV\xbdW\xe8\xef\x97\xa0\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x01\xd7d\xb7h\xecV:\x9b\xb4\xb47\x1a\xe6S\xb5\x15i\xed\xf4\xaf\xa8\xa8~\xea\x89\xb3#b+\x9c\xbc\xf7\xe5\xdc\xdd{\x87b\x00\xa1\x99}\xd3Q*zU\xc3\xabv\x8d!\xce\xe5\xb6\xd2O\x13b\xa6\x9a\xc7P\xc9e\x85\xb0\xa4O^LTj\xaa+\x95;v\xe5\xb9f\xa2\xd7\nwR$\xcf\xd2\xcdX\x8eM\xb7XW\x14\x99\\\x8b\xde\xdd\x17\x87\xfcId\x01Z\xb0\xfc\xa3:\xd4\x8e\x93\x96\x0b\xbd\xcbO2\x8c_\x16\xb0\xd0Vy\xcd\xd7[t\xb0\xab\xe5\x95\x88\xc7=\xeeV\xa3Q\xca\x9b"5\x15vD^k\xba\x9c\x9b\x95\xbfVp}|\xces\xcck\x17u\xff\x00\x13\xb9>\x89\xb5V\xd6N\x8c\xa9\xa8\xe1\xa6b,\xf4\xe8\xbc\x9c\xacr9\x15?+\x8bdE\xdbv\xd8\xc0\x045.\xbf\xd0>\x9d\xd1\xdb\xf4\xd3R\xaa\xae\xbbl\xdb{\xb1\xf7\xc6\xf4\x7fy\xceU\xe1jo\xda\xbc\xf6\xf0\x9d\x1fF\xfd*\xc9-\xf9\xfeG\xab\xda\x81K\x05\x06E\x7f|\x9ew\xb6D\xf4zQD\xf7#\x97\x89\xc9\xc9]\xb3Z\xd4\xdb\xb1\x11w\xe6\xe5D\xb0 \x0f\x85\xca\xa9\xb46\xea\x9a\xd7\xc3Q;i\xe2|\xab\x1d<K$\xafF\xa2\xaf\x0b\x18\x9c\xdc\xe5\xdbdD\xe6\xab\xc8\xa3\xdd1\xeb\xb3=Z\xbdX!\xc6t\xafP\x19o\xb3G?\xf0\xf5X\xfdDn\x99\xf3,{\xec\xd4j\xec\xd4H\xdb\xcdv]\xd5yw\xef@\x02-\xd0\x8c\xfe\xb7!\xb0ZlW\x9c\x1b1\xc7n\xb4\x96\xf62\xa5\xd7;<\xb0S9\xd1\xb5\xad^\t\\\x88\x8b\xbfj"\xec\xbd\xbc\xb9\x12\x90\x00V\x9e\x99Z-\x91\xe6\xd5\xd6\x8c\xf7\x04o[\x90\xda#H\xa4\xa6k\xd1\x92K\x1b\x1e\xb2F\xf8\xd5yq\xb1\xca\xeeK\xda\x8b\xcb\x9a"/+\x14\xe9\x1f}\xa4\xb6\xc3E\x9e\xe8\xee\xa0R\xde\xa3j2e\xb7Z\x1d,S9\x13\x9b\x9a\x8fV+w\xed\xdb\xd9\'\x85K\x18\x00\xaf\xb50\xe6\xfa\xe7\x93X\x9bx\xc3\xee8v\x9f\xda+\xa3\xb8\xcf\x15\xd9\x11\x95\xb7Y\xa3\xe7\x1ck\x12sdh\xab\xcf~\xde\xe2\xef\xb6\xd3\xc5\xde\xb5\x96\xdbU]\xc2Jz\xaa\x86SB\xf9\x9d\x15,.\x96i\x11\xa8\xab\xc2\xc67\x9b\x9c\xbbl\x8dNj\xbc\x8eP\x02\x89t\xa4\xaf\xcfs\xfdV\xb0d\xb8\x86\x96\xe7\x8c\xa6\xb0\xc3\x1fQ%f=R\xd5\x96f\xca\xb2o\xc2\x8d]\x9b\xedS\x9f>\xd2\xc0\xd9u\xee\xe1=\xae\'\xddtKU\xe9n\n\xd4\xeb`\x83\x1e|\x91\xf1wxdr\xb7t\xf0\xaa!6\x00+\x1e\xa0X\xb5\x83\xa4\n\xd3\xd8+l\x12i\xde\x08\xc9\x9b-R\xd7\xca\xd7\xd7V\xf0\xae\xe8\x9dS{6^h\xd7l\x9b\xec\xaa\xae\xd9\x11\'\xdc#\x17\xb2`X]\x1e9\x8f\xd1\xbe\x1bu\xba\x15F1\xa9\xc7$\x8b\xcd\\\xe5\xdb\xdb=\xcb\xba\xaf}T\xef\xc0\x14_\xa5\xddnk\xa9\xb9\xa6=Y\x8ai~}\xe7+$N\xdaj\xacz\xa65\x96W=\xae]\x9b\xc2\xab\xc2\x88\xc6\xf3]\x97\xb7\x97~\xde\xe9\xb6h\x99\x9d\xad\xf5O\xc62\\zxZ\xce\xba\x9e\xf5m\x92\x95\xdcNEUF+\x93g\xa2l\xbb\xaax;7=X\x02\xbf\xebf\x87\xde\xaa3\xeamX\xd2z\xfakVeJ\xf4\x92\xa2\x96ocO_\xcbe\xdd{\x8es}\x8b\x91y9;U\xab\xba\xaf:\xdf\xaf\xb7K\\\r\xa4\xcft\x8f;\xb3\xdd\x18\x9bH\xb4\x16\xd5\xac\xa5\x91\xdd\xd5\x8eV\xaan\x9e\x0e{o\xda\xbd\xa4\xe4\x00\xaeY\xbeg\xab:\xb7j\x9f\x15\xd3\x8c\x0e\xf5\x89\xdb+\x98\xb1V\xdf\xf2&y\xd1\xed\x85\xdc\x9c\x91G\xcd\xdc\xd3t\xe2o\x12\xed\xdcE\xe6\x92^\x83iE\x87I0\xff\x00Am.Z\xaa\xca\x87$\xb7\n\xf7\xb3\x85\xf52"r\xe5\xf9-n\xea\x8dn\xfc\xb7^\xd5UU\x90\x80\x00uy}}}\xab\x13\xbb\xdd-V\xf7\\\xab\xe8\xe8f\x9e\x96\x8d\xbb\xefQ+\x18\xaedi\xb7\xe7*"}$\x11\xa4})\xb1K\xe6\tt\xb9\xe7\xb56\xfcz\xfdjs\xfa\xea\x06\xb9\xcdZ\xa6\xa2n\xde\xa5\xaeUr\xbb}\xda\xad\xddU\x157]\x91@\xf0}\x01=\xf4\xb57\xfaFyy\x8b\x81p\xa4\xa7\xb8PTPU\xc6\x92\xd3\xd4\xc4\xe8eb\xf69\x8eEEO\xa5\x15J\xdd\xd03\t\xbc\xd9\xf1\xcc\x8b8\xbfQIE>OT\xc7\xd2\xc3+U\xae\xea\x18\xafw\x1e\xcb\xcd\x11\xce\x91v\xdf\xb5\x1a\x8b\xd8\xa8Y\x80)F\x11\x8d\xea\xafF]F\xbb\xc9m\xc3\xeeY\x9e\x19rTk\xdfnb\xc9"\xb1\xaa\xab\x1c\x8a\x8dEVH\xd4r\xa2\xa3\x93\x85w]\x97\xb1Ra\x8f\xa4=e\xd6\x0f;\xe3\x1a5\xa9\x15\xf7W&\xcc\x8a\xae\xda\xdajv\xbb\xf9\xf3q9\x1a\x9b\xf7U\t\xd8\x01\x18h6\x0f\x7f\xc7\x9b\x7f\xcbsii\xe5\xcb\xf2\x9a\xa6\xd5\\[N\xbcQR\xc6\xc6\xf0\xc3N\xc5\xee\xa3\x1a\xaa\x9b\xee\xbd\xed\xd7m\xd7\x93\xae\x99\xa5F9\x8b\xdc-\x96\xfcC*\xc8n5\xf6\xf9\x99L\xdbM\xa6Z\x98\x9a\xf75X\x9dd\x8dM\x99\xcdw\xdb\xb7n\xc4RF\x00R\xee\x86w\x9c\x9fK-\xd7\xdb6[\xa6\x1a\x85\x1c\x15\xf3\xc7Q\x05E69S*5\xc8\xd5k\x9a\xe4\xe1EO\xc9TT\xdf\xbb\xd8O]!\xf4\xaa\xdb\xac\x9ar\xc8\x12%\xa3\xbd\xd3\xc5\xe7\x9bMD\xf1\xab\x1f\x13\xdc\xd4U\x8aD^h\xd7rG\'qQ\x17m\xdb\xb1+\x80*WF\xfdJ\xd5\x0c\x16\xc8\xdc\'Q4\xbb>\xb8P\xd09a\xa1\xb8\xd1\xd9g\x9d\xf0\xb19un\xe5\xb3\xd8\x9b{\x175W\x96\xc9\xcd6\xdb\xe5\xd3b\xf5\x92j\x0e;C\x86b\x9aq\x9d\xd6y\xc6\xec\xb5UU\x8b`\xa8H\\\xb1\xb2H\xda\x91\xb9\x1a\xbch\xbdc\x97\x8b\xb3dM\xb7\xdc\xb7@\x08\x8b\xa3.[_p\xc0\xac\x18\x95\xeb\x0c\xcbl\x17[=\xa2*y\x9fs\xb4MOO"B\x8c\x89\x15\x92\xb9\x11\x15\xceM\x97\x87\x92\xfbn\\\xb7!\x1e\x99\xf7\x8c\x9bS\xadV[\x06\'\xa6Z\x814t\x15rTTT\xcf\x8e\xd4\xc6\xc5^\x1e\x06\xb5\x89\xc2\xaa\xbd\xaeUU\xdb\xb9\xb6\xfd\xcb\x96\x00\xae\xdaS\xad\x99%\xbb\x06\xb5\xda2\xed\x17\xd4\xe8\xeet\x14\x91\xd3:Z\x1czYb\x9f\x81\xa8\xd4w\xb2\xe1V\xaa\xa2"\xaam\xb2/t\xf9j\x15\xe7Yu\x86\xd3>#\x88`\xd7\x0c\x1a\xc5\\\xde\xaa\xe1w\xc8\x1e\x90T:\x15\xf6\xd1\xb2\x16\xaa\xb9\xbb\xa7%\xdb}\xd1v\xdd\xbd\xa5\x8e\x00x]\x11\xd3\x0b\x0e\x94\xe1q\xe3\xb6^)\xe5{\xba\xda\xda\xd9\x1a\x89%T\xaa\x9b+\x95;\x88\x89\xc9\x1b\xdcN\xfa\xee\xab\x00\xf4\xdc\xbcd\xb9\xf5\x86\x8f\t\xc54\xe7:\xad\xf4>\xee\xea\x9a\xaa\xd4\xb0\xd4u\x0fX\xd9$MH\x9c\x8d^6\xaf\x1b\x97\x899l\x88\xa9\xbe\xe5\xb8\x00D]\x19r\xda\xfb\x86\x05`\xc4\xafXf[`\xba\xd9\xed\x11S\xcc\xfb\x9d\xa2jzy\x12\x14dH\xac\x95\xc8\x88\xaerl\xbc<\x97\xdbr\xe5\xb9\xe5zQh-\xcb5\xbd\xd1j\x16\x01Y\x1d\xbf2\xb7\xacnV\xb9\xfd[j\xba\xb5E\x8d\xc8\xfe\xc6\xc8\xdd\xb6E^J\x9b"\xaalXp\x04\x01\x8a\xeb\xdeIk\xa0\x8a\xdd\xa9\xbaO\x9b[\xaf15\x19-E\xb6\xd6\xea\x9aZ\x87\'.&\xaa/-\xfb\xc8\xaeO\t\xf4\xcc2\x8dI\xd5\x9b4\xf8\xae\x03\x86\xde\xb1+]\xc1\x8b\r~A\x91C\xe7W\xc7\x0b\xb99 \x83u{\x9c\xe4\xdd8\xb9m\xe0\xdd\x1c\x93\xd8\x02+\xe8\x9ba\xbbc\x1a\x07\x8f\xd8o\x945\x147\n9kc\x9a\x19\xa3V9?\xed\x93\xaa.\xca\x88\xbb**9\x17\xba\x8a\x8a\x9c\x94\xf0\xda\x8b\xa2\xb9N+\xa9\xce\xd5\xad\x14\x92\x92;\xac\xaa\xe7\\\xecU\x0e\xe0\x86\xb5\x1d\xceDb\xf2D\xe2TEV\xaa\xa2#\xbd\x92*.\xc8X\xd0\x04!I\xd2\r\xd4p\xa4\x19^\x93\xea-\x9e\xe6\xd4\xdaHa\xb4-L._\xfe\xdc\xa8\xa9\xc4\x9e\x1d\x90\xf2\xf9\xf5\xc7W5\xd6\xdc\xfcK\x18\xc4.8\x1e%V\xa8\xdb\x8d\xd6\xfc\xce\xa6\xaax\xb7\xe6\xc6B\x9e\xc9\x11S\xb5\x13twb\xb9\xa8\xab\xbd\x97\x00y\r"\xd3\xdb\x06\x99aT\xd8\xbe?\x1b\xba\xa8\xd5d\xa8\xa8\x93n\xb2\xa6eD\xe2\x91\xfe\x15\xd9\x11\x13\xb8\x88\x89\xdc=x\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x0e\xaa\xa3\x1a\xc7j.>\x89O`\xb5K]\xbf\x17\x9e_G\x1b\xa5\xdf\xbf\xc4\xa9\xb9\xda\x80\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00?\xff\xd9'
TELEGRAM_CHAT_ID  = os.environ.get('TELEGRAM_CHAT_ID', '')
GOOGLE_TOKEN_B64  = os.environ.get('GOOGLE_TOKEN_B64')
SHEETS_ID         = os.environ.get('SHEETS_ID', '1saZ98ZEqj46nxcvQC5V0oKJ0GLL-ly0MeuxZqsZHNKI')
TIMEZONE          = 'Europe/Madrid'
SCOPES            = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/tasks',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# ─────────────────────────────────────────────
# PROCURADORES Y CONFIGURACIÓN RESOLUCIONES
# ─────────────────────────────────────────────
PROCURADORES = {
    'monica@lopezmanso.com': 'Mónica López Manso',
    'adriaparu@icloud.com': 'Adrià Paños (test)',
}

DRIVE_FOLDER_RESOLUCIONES = 'Resoluciones_Test'

diario_secretaria = []  # Se r
CORREOS_ACTIVOS = True  # Se puede activar/desactivar con /correos_on y /correos_offesetea cada día al enviar el diario

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

def get_drive_service():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        logger.error(f"Error conectando con Google Drive: {e}")
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
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.image import MIMEImage
        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')
        html_body = (
            '<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">'
            '<div style="text-align:center;margin-bottom:24px;"><img src="cid:logo_ap" alt="AP Estudio Juridico" style="height:80px;width:auto;"/></div>'
            '<div style="border-top:3px solid #b8960c;padding-top:20px;">' + formatted + '</div>'
            '<div style="margin-top:30px;padding-top:15px;border-top:1px solid #ddd;font-size:0.85em;color:#666;">'
            '<em>Secretar\u00eda \u2014 AP Estudio Jur\u00eddico</em></div></body></html>')
        msg_root = MIMEMultipart('related')
        msg_root['Subject'] = subject
        msg_root['From']    = GMAIL_USER
        msg_root['To']      = to_addr
        msg_alt = MIMEMultipart('alternative')
        msg_root.attach(msg_alt)
        msg_alt.attach(MIMEText(html_body, 'html', 'utf-8'))
        logo_mime = MIMEImage(LOGO_BYTES, _subtype='png')
        logo_mime.add_header('Content-ID', '<logo_ap>')
        logo_mime.add_header('Content-Disposition', 'inline', filename='logo.png')
        msg_root.attach(logo_mime)
        raw = base64.urlsafe_b64encode(msg_root.as_bytes()).decode()
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
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.image import MIMEImage
        from email.mime.base import MIMEBase
        from email import encoders as _enc
        formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', body_text)
        formatted = formatted.replace('\n', '<br>')
        html_body = (
            '<html><body style="font-family:Georgia,serif;max-width:600px;margin:0 auto;padding:20px;">'
            '<div style="text-align:center;margin-bottom:24px;"><img src="cid:logo_ap" alt="AP Estudio Juridico" style="height:80px;width:auto;"/></div>'
            '<div>' + formatted + '</div>'
            '<hr style="border:none;border-top:1px solid #ddd;margin:24px 0;">'
            '<div style="font-size:0.85em;color:#666;text-align:center;"><em>Secretar\u00eda \u2014 AP Estudio Jur\u00eddico</em></div>'
            '</body></html>')
        msg_root = MIMEMultipart('mixed')
        msg_root['Subject'] = subject
        msg_root['From']    = GMAIL_USER
        msg_root['To']      = to_addr
        msg_rel = MIMEMultipart('related')
        msg_root.attach(msg_rel)
        msg_alt = MIMEMultipart('alternative')
        msg_rel.attach(msg_alt)
        msg_alt.attach(MIMEText(html_body, 'html', 'utf-8'))
        logo_mime = MIMEImage(LOGO_BYTES, _subtype='png')
        logo_mime.add_header('Content-ID', '<logo_ap>')
        logo_mime.add_header('Content-Disposition', 'inline', filename='logo.png')
        msg_rel.attach(logo_mime)
        part = MIMEBase('application', 'pdf')
        part.set_payload(pdf_bytes)
        _enc.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
        part.add_header('Content-Type', 'application/pdf', name=pdf_filename)
        msg_root.attach(part)
        raw = base64.urlsafe_b64encode(msg_root.as_bytes()).decode()
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
            hoja = hoja.strip("'")  # quitar comillas simples si las hay
            meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
            hojas_reales = {s['properties']['title'].lower(): s['properties']['title'] for s in meta.get('sheets', [])}
            hoja_real = hojas_reales.get(hoja.lower(), hoja)
            rango = f"'{hoja_real}'!{celdas}" if ' ' in hoja_real else f"{hoja_real}!{celdas}"
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
        hoja_rango = f"'{hoja_real}'!A1" if ' ' in hoja_real else f"{hoja_real}!A1"
        svc.spreadsheets().values().append(
            spreadsheetId=SHEETS_ID, range=hoja_rango,
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


def sheets_insert_row_at(hoja, fila_valores, posicion=2):
    """Inserta una fila en la posición indicada (1-indexed), desplazando las demás hacia abajo."""
    try:
        svc = get_sheets_service()
        if not svc:
            return False
        meta = svc.spreadsheets().get(spreadsheetId=SHEETS_ID).execute()
        sheet_id = None
        for s in meta.get('sheets', []):
            if s['properties']['title'] == hoja:
                sheet_id = s['properties']['sheetId']
                break
        if sheet_id is None:
            return False
        # Insertar fila vacía en posición (0-indexed = posicion-1)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SHEETS_ID,
            body={'requests': [{'insertDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': posicion - 1, 'endIndex': posicion},
                'inheritFromBefore': False
            }}]}
        ).execute()
        # Escribir valores en la nueva fila
        svc.spreadsheets().values().update(
            spreadsheetId=SHEETS_ID,
            range=f"{hoja}!A{posicion}",
            valueInputOption='USER_ENTERED',
            body={'values': [fila_valores]}
        ).execute()
        return True
    except Exception as e:
        logger.error(f"Error insertando fila en {hoja}: {e}")
        return False

def siguiente_id_cliente():
    """Obtiene el ID más alto de la hoja Clientes y devuelve el siguiente."""
    rows = sheets_read("Clientes!A2:A200")
    ids = [int(str(r[0]).strip()) for r in rows if r and str(r[0]).strip().isdigit()]
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


# ─────────────────────────────────────────────
# FACTURAS RECIBIDAS
# Columnas: A=Id, B=Fecha, C=Proveedor, D=NIF, E=Concepto,
#           F=Base, G=IVA%, H=CuotaIVA, I=IRPF%, J=CuotaIRPF, K=Total
# ─────────────────────────────────────────────
def get_facturas_recibidas(trimestre=None, año=None):
    """Lee facturas recibidas con filtro opcional por trimestre/año."""
    import pytz as _pytz
    from datetime import datetime as _dt
    tz   = _pytz.timezone(TIMEZONE)
    year = año or _dt.now(tz).year
    rows = sheets_read("'Facturas Recibidas'!A2:K200")
    result = []
    for row in rows:
        if not row or len(row) < 3:
            continue
        def col(i): return str(row[i]).strip() if len(row) > i else ''
        fecha = col(1)
        if trimestre and fecha:
            try:
                mes = int(fecha[5:7]) if len(fecha) >= 7 else 0
                t   = (mes - 1) // 3 + 1
                if str(t) != str(trimestre):
                    continue
                if str(year) not in fecha:
                    continue
            except:
                pass
        result.append({
            'id':        col(0), 'fecha':    col(1), 'proveedor': col(2),
            'nif':       col(3), 'concepto': col(4), 'base':      col(5),
            'iva_pct':   col(6), 'cuota_iva': col(7),
            'irpf_pct':  col(8), 'cuota_irpf': col(9), 'total': col(10),
        })
    return result


def get_nif_proveedor(proveedor):
    """Busca el NIF de un proveedor en facturas recibidas anteriores."""
    rows = sheets_read("'Facturas Recibidas'!A2:K200")
    prov_norm = normalizar(proveedor)
    for row in rows:
        if len(row) < 4:
            continue
        nombre_row = normalizar(str(row[2]))
        nif_row    = str(row[3]).strip()
        if prov_norm in nombre_row or nombre_row in prov_norm:
            if nif_row:
                return nif_row
    return ''

def siguiente_id_factura_recibida():
    rows = sheets_read("'Facturas Recibidas'!A2:A200")
    ids  = [int(str(r[0]).strip()) for r in rows if r and str(r[0]).strip().isdigit()]
    if ids:
        return max(ids) + 1
    # Si no hay IDs, contar filas con datos
    rows_all = sheets_read("'Facturas Recibidas'!B2:B200")
    count = sum(1 for r in rows_all if r and r[0].strip())
    return count + 1

def calcular_trimestre(trimestre, año=None):
    """Calcula el resumen IVA/IRPF de un trimestre."""
    import pytz as _pytz
    from datetime import datetime as _dt
    tz   = _pytz.timezone(TIMEZONE)
    year = año or _dt.now(tz).year

    meses = {1: (1,3), 2: (4,6), 3: (7,9), 4: (10,12)}
    mes_inicio, mes_fin = meses.get(int(trimestre), (1,3))

    def mes_en_rango(fecha_str):
        try:
            mes = int(fecha_str[5:7])
            return mes_inicio <= mes <= mes_fin and str(year) in fecha_str
        except:
            return False

    # ── Facturas emitidas ──
    rows_emit = sheets_read("Facturas!A2:N200")
    iva_repercutido  = 0.0
    irpf_retenido    = 0.0
    base_emit        = 0.0
    facturas_emit    = 0

    for row in rows_emit:
        if not row or not str(row[0]).strip().isdigit():
            continue
        def ce(i): return str(row[i]).strip() if len(row) > i else ''
        fecha = ce(1)  # B=Fecha
        if not mes_en_rango(fecha):
            continue
        try: base_emit       += float(ce(6).replace(',','.').replace('€','').strip() or 0)
        except: pass
        try: iva_repercutido += float(ce(7).replace(',','.').replace('€','').strip() or 0)
        except: pass
        try: irpf_retenido   += float(ce(8).replace(',','.').replace('€','').strip() or 0)
        except: pass
        facturas_emit += 1

    # ── Facturas recibidas ──
    rows_recib = sheets_read("'Facturas Recibidas'!A2:K200")
    iva_soportado = 0.0
    base_recib    = 0.0
    facturas_recib = 0

    for row in rows_recib:
        if not row or len(row) < 3:
            continue
        def cr(i): return str(row[i]).strip() if len(row) > i else ''
        fecha = cr(1)
        if not fecha or not mes_en_rango(fecha):
            continue
        try: base_recib     += float(cr(5).replace(',','.').replace('€','').strip() or 0)
        except: pass
        try: iva_soportado  += float(cr(7).replace(',','.').replace('€','').strip() or 0)
        except: pass
        facturas_recib += 1

    iva_liquidar = round(iva_repercutido - iva_soportado, 2)

    nombres_t = {1:'1T (Ene-Mar)', 2:'2T (Abr-Jun)', 3:'3T (Jul-Sep)', 4:'4T (Oct-Dic)'}
    return {
        'trimestre':       f"{nombres_t.get(int(trimestre), str(trimestre))} {year}",
        'facturas_emit':   facturas_emit,
        'base_emit':       round(base_emit, 2),
        'iva_repercutido': round(iva_repercutido, 2),
        'irpf_retenido':   round(irpf_retenido, 2),
        'facturas_recib':  facturas_recib,
        'base_recib':      round(base_recib, 2),
        'iva_soportado':   round(iva_soportado, 2),
        'iva_liquidar':    iva_liquidar,
    }

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

FORMATO OBLIGATORIO DEL TÍTULO (campo "summary"):
Siempre usa: [TIPO_EVENTO] - [CLIENTE] - [DESCRIPCIÓN]

Tipos permitidos: REUNIÓN · JUICIO · LLAMADA · VIDEOLLAMADA · GESTIÓN · VENCIMIENTO · AUDIENCIA · COMPARECENCIA
- Si hay cliente identificable: usa su nombre exacto
- Si no hay cliente: usa GENERAL
- Descripción: máximo 5 palabras

Ejemplos:
- "reunión con Juan Pérez sobre concurso" → REUNIÓN - Juan Pérez - Consulta concurso
- "juicio García contra Banco Santander audiencia previa" → JUICIO - García - Audiencia previa
- "llamar a María seguimiento" → LLAMADA - María - Seguimiento caso
- "preparar escritos" → GESTIÓN - GENERAL - Preparar escritos
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
{{"action":"add_cliente","nombre":"","nif":"","direccion":"","cp":"",""poblacion":"","provincia":"","pais":"España","email":"","telefono":""}}
- Campos obligatorios: nombre, nif. Pide los que falten antes de proceder.
- Comprueba duplicados automáticamente. Si faltan datos no obligatorios déjalos vacíos.
{{"action":"query_casos","cliente":"nombre opcional"}}
{{"action":"add_caso","cliente":"","materia":"","descripcion":"","juzgado":"","autos":"","estado":"Activo","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD"}}
{{"action":"update_caso_estado","autos":"PA 1/2026","estado":"","proxima_actuacion":"","fecha_actuacion":"YYYY-MM-DD"}}
{{"action":"query_facturas","estado":"Pendiente"}}
{{"action":"cobrar_factura","num_factura":"15","fecha_cobro":"YYYY-MM-DD"}}

FACTURAS RECIBIDAS (gastos del despacho):
{"action":"add_factura_recibida","fecha":"YYYY-MM-DD","proveedor":"nombre empresa","nif":"","concepto":"","base_imponible":0,"iva":21,"irpf":0}
- USA SIEMPRE add_factura_recibida cuando el usuario diga "factura recibida", "me han facturado", "factura de [empresa/proveedor]".
- NUNCA uses create_invoice_bd ni create_invoice para facturas recibidas. Son completamente distintas.
- create_invoice_bd = facturas que TU emites a clientes.
- add_factura_recibida = facturas que recibes de proveedores.
- Si falta NIF, ponlo vacio "" — el sistema lo rellena solo.
- Si falta fecha, usa hoy.

CALCULO TRIMESTRAL IVA/IRPF:
{"action":"calculo_trimestral","trimestre":1,"año":2026}
- Usa calculo_trimestral cuando pidan "liquidacion", "trimestre", "IVA trimestral", "modelo 303", "calcular IVA".
- trimestre: 1=Ene-Mar, 2=Abr-Jun, 3=Jul-Sep, 4=Oct-Dic. Si no indica año, usa el actual.
- NUNCA respondas con action:none para estas peticiones. Ejecuta directamente.

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
- En los correos enviados a clientes NUNCA incluyas el nombre del abogado. Firma siempre como "Secretaría AP Estudio Jurídico".
- Para emails de confirmación de cita: "Estimado/a [nombre], le confirmamos su cita programada para el [fecha] a las [hora]h. Saludos cordiales, Secretaría AP Estudio Jurídico"
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

# ─────────────────────────────────────────────
# RECORDATORIOS DE CITAS (Lunes a Viernes 8:00 AM)
# ─────────────────────────────────────────────
async def appointment_reminders(bot):
    """Envía recordatorios de citas de mañana por Telegram y por email a los clientes."""
    tz       = pytz.timezone(TIMEZONE)
    today    = datetime.now(tz)
    tomorrow = today + timedelta(days=1)

    # Obtener eventos de mañana
    t_start = tomorrow.replace(hour=0,  minute=0,  second=0,  microsecond=0)
    t_end   = tomorrow.replace(hour=23, minute=59, second=59, microsecond=0)

    service = get_calendar_service()
    if not service:
        return

    try:
        result = service.events().list(
            calendarId='primary',
            timeMin=t_start.isoformat(),
            timeMax=t_end.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = result.get('items', [])
    except Exception as e:
        logger.error(f"Error obteniendo eventos recordatorio: {e}")
        return

    if not events:
        return

    # Construir mensaje Telegram
    fecha_str = tomorrow.strftime('%d/%m/%Y')
    msg = f"📅 Recordatorio de citas para mañana ({fecha_str}):\n\n"
    for ev in events:
        start = ev.get('start', {})
        hora  = ''
        if 'dateTime' in start:
            dt   = datetime.fromisoformat(start['dateTime'])
            hora = dt.astimezone(tz).strftime('%H:%M')
        elif 'date' in start:
            hora = 'Todo el día'
        titulo      = ev.get('summary', 'Sin título')
        descripcion = ev.get('description', '')
        msg += f"🕐 {hora} — {titulo}\n"
        if descripcion:
            msg += f"   {descripcion}\n"
        msg += "\n"

    try:
        await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=msg)
        logger.info("Recordatorio de citas enviado por Telegram.")
    except Exception as e:
        logger.error(f"Error enviando recordatorio Telegram: {e}")

    # Enviar emails a clientes detectados
    clientes = get_todos_clientes()

    for ev in events:
        start  = ev.get('start', {})
        titulo = ev.get('summary', '')
        hora   = ''
        fecha_ev = ''
        if 'dateTime' in start:
            dt       = datetime.fromisoformat(start['dateTime'])
            dt_local = dt.astimezone(tz)
            hora     = dt_local.strftime('%H:%M')
            fecha_ev = dt_local.strftime('%d/%m/%Y')
        elif 'date' in start:
            hora     = 'Todo el día'
            fecha_ev = datetime.fromisoformat(start['date']).strftime('%d/%m/%Y')

        titulo_norm = normalizar(titulo)

        # Buscar clientes cuyo nombre aparezca en el título
        for c in clientes:
            nombre_c      = c.get('nombre', '')
            nombre_c_norm = normalizar(nombre_c)
            if not nombre_c_norm:
                continue

            # Coincidencia flexible: el nombre del cliente está en el título
            if nombre_c_norm in titulo_norm:
                email_c = c.get('email', '').strip()
                if not email_c:
                    logger.info(f"Cliente {nombre_c} sin email — no se envía recordatorio.")
                    continue

                # Recuperar datos completos del cliente
                c_full = get_cliente(nombre_c)
                nombre_display = c_full['nombre'] if c_full else nombre_c

                asunto = f"Recordatorio de reunión mañana — {fecha_ev}"
                cuerpo = (
                    f"Estimado/a {nombre_display},\n\n"
                    f"Le recordamos que mañana tenemos una reunión programada.\n\n"
                    f"Detalles de la cita:\n"
                    f"Fecha: {fecha_ev}\n"
                    f"Hora: {hora}\n"
                    f"Asunto: {titulo}\n\n"
                    f"Si necesita modificar la cita, por favor indíquelo respondiendo a este correo.\n\n"
                    f"Un cordial saludo.\n"
                    f"Secretaría AP Estudio Jurídico"
                )
                ok = send_email(email_c, asunto, cuerpo)
                if ok:
                    logger.info(f"Recordatorio enviado a {nombre_display} ({email_c})")
                else:
                    logger.error(f"Error enviando recordatorio a {nombre_display} ({email_c})")

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

async def cmd_correos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fuerza la revisión de correos manualmente."""
    await update.message.reply_text("⏳ Revisando correos...")
    await procesar_correos(context.bot)
    await update.message.reply_text("✅ Revisión completada.")

async def cmd_correos_on(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Activa el procesado automático de correos de procuradores."""
    global CORREOS_ACTIVOS
    CORREOS_ACTIVOS = True
    await update.message.reply_text("✅ Procesado de correos de procuradores ACTIVADO.")

async def cmd_correos_off(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Desactiva el procesado automático de correos de procuradores."""
    global CORREOS_ACTIVOS
    CORREOS_ACTIVOS = False
    await update.message.reply_text("⏸ Procesado de correos de procuradores DESACTIVADO.")

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
                body_email   = "Estimado/a cliente,\n\nLe adjunto la factura por los servicios prestados por este despacho.\n\nSaludos cordiales,\nSecretaría AP Estudio Jurídico"
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
                nombre_nuevo = data.get('nombre', '').strip()
                nif_nuevo    = normalizar(data.get('nif', '').strip())

                # Verificar duplicados
                rows_check = sheets_read("Clientes!A2:J200")
                duplicado = False
                for row in rows_check:
                    if len(row) < 2:
                        continue
                    nombre_exist = normalizar(str(row[1]))
                    nif_exist    = normalizar(str(row[2])) if len(row) > 2 else ''
                    if normalizar(nombre_nuevo) == nombre_exist:
                        await update.message.reply_text(
                            f"⚠️ Ya existe un cliente con el nombre '{row[1]}'. No se ha registrado.")
                        duplicado = True
                        break
                    if nif_nuevo and nif_exist and nif_nuevo == nif_exist:
                        await update.message.reply_text(
                            f"⚠️ Ya existe un cliente con NIF/CIF '{row[2]}' ({row[1]}). No se ha registrado.")
                        duplicado = True
                        break

                if not duplicado:
                    nuevo_id = siguiente_id_cliente()
                    # Columnas: Id, Cliente, NIF, Dirección, CP, Población, Provincia, País, Email, Teléfono
                    fila = [
                        str(nuevo_id),
                        nombre_nuevo,
                        data.get('nif', ''),
                        data.get('direccion', ''),
                        data.get('cp', ''),
                        data.get('poblacion', ''),
                        data.get('provincia', ''),
                        data.get('pais', 'España'),
                        data.get('email', ''),
                        data.get('telefono', ''),
                    ]
                    # Insertar en fila 2 (debajo del encabezado), desplazando los demás
                    ok = sheets_insert_row_at('Clientes', fila, posicion=2)
                    if ok:
                        msg = (f"✅ Cliente registrado correctamente\n"
                               f"ID: {nuevo_id}\n"
                               f"Nombre: {nombre_nuevo}\n"
                               f"NIF/CIF: {data.get('nif','')}\n"
                               f"Email: {data.get('email','')}\n"
                               f"Teléfono: {data.get('telefono','')}")
                        await update.message.reply_text(msg)
                    else:
                        await update.message.reply_text("❌ Error al registrar el cliente.")

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

            elif action == 'add_factura_recibida':
                import pytz as _pytz2
                from datetime import datetime as _dt2
                proveedor = data.get('proveedor', '')
                # Auto-completar NIF si no se ha indicado
                nif = data.get('nif', '').strip()
                if not nif and proveedor:
                    nif = get_nif_proveedor(proveedor)
                base     = float(data.get('base_imponible', 0))
                iva_pct  = float(data.get('iva', 21))
                irpf_pct = float(data.get('irpf', 0))
                cuota_iva  = round(base * iva_pct  / 100, 2)
                cuota_irpf = round(base * irpf_pct / 100, 2)
                total      = round(base + cuota_iva - cuota_irpf, 2)
                nuevo_id   = siguiente_id_factura_recibida()
                fecha      = data.get('fecha', _dt2.now(_pytz2.timezone(TIMEZONE)).strftime('%Y-%m-%d'))
                fila = [
                    str(nuevo_id), fecha,
                    proveedor, nif, data.get('concepto',''),
                    base, iva_pct, cuota_iva,
                    irpf_pct, cuota_irpf, total
                ]
                ok = sheets_append('Facturas Recibidas', fila)
                if ok:
                    nif_txt = f" | NIF: {nif}" if nif else ""
                    acciones_completadas.append(
                        f"✅ Factura recibida registrada\n"
                        f"Proveedor: {proveedor}{nif_txt}\n"
                        f"Base: {base:.2f}€ | IVA: {cuota_iva:.2f}€ | Total: {total:.2f}€"
                    )
                else:
                    acciones_completadas.append("❌ Error al registrar la factura recibida.")

            elif action == 'calculo_trimestral':
                t    = data.get('trimestre', 1)
                anyo = data.get('año', None)
                res  = calcular_trimestre(t, anyo)
                signo = "a ingresar 🔴" if res['iva_liquidar'] > 0 else "a devolver 🟢"
                msg = (
                    f"📊 Liquidación {res['trimestre']}\n\n"
                    f"FACTURAS EMITIDAS ({res['facturas_emit']})\n"
                    f"  Base imponible:    {res['base_emit']:.2f} €\n"
                    f"  IVA repercutido:   {res['iva_repercutido']:.2f} €\n"
                    f"  IRPF retenido:     {res['irpf_retenido']:.2f} €\n\n"
                    f"FACTURAS RECIBIDAS ({res['facturas_recib']})\n"
                    f"  Base imponible:    {res['base_recib']:.2f} €\n"
                    f"  IVA soportado:     {res['iva_soportado']:.2f} €\n\n"
                    f"RESULTADO IVA\n"
                    f"  {res['iva_repercutido']:.2f} - {res['iva_soportado']:.2f} = "
                    f"{abs(res['iva_liquidar']):.2f} € {signo}\n\n"
                    f"IRPF RETENIDO (informativo — lo ingresa el pagador)\n"
                    f"  {res['irpf_retenido']:.2f} €"
                )
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
                    body_email  = "Estimado/a cliente,\n\nLe adjunto la factura por los servicios prestados por este despacho.\n\nSaludos cordiales,\nSecretaría AP Estudio Jurídico"
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
# GOOGLE DRIVE — HELPERS
# ─────────────────────────────────────────────

def drive_find_folder(nombre, parent_id=None):
    """Busca una carpeta en Drive por nombre. Devuelve el id o None."""
    service = get_drive_service()
    if not service:
        return None
    try:
        q = f"mimeType='application/vnd.google-apps.folder' and name='{nombre}' and trashed=false"
        if parent_id:
            q += f" and '{parent_id}' in parents"
        result = service.files().list(q=q, fields='files(id, name)', pageSize=1).execute()
        files = result.get('files', [])
        return files[0]['id'] if files else None
    except Exception as e:
        logger.error(f"Error buscando carpeta Drive '{nombre}': {e}")
        return None


def drive_create_folder(nombre, parent_id=None):
    """Crea una carpeta en Drive. Devuelve el id o None."""
    service = get_drive_service()
    if not service:
        return None
    try:
        metadata = {
            'name': nombre,
            'mimeType': 'application/vnd.google-apps.folder',
        }
        if parent_id:
            metadata['parents'] = [parent_id]
        folder = service.files().create(body=metadata, fields='id').execute()
        return folder.get('id')
    except Exception as e:
        logger.error(f"Error creando carpeta Drive '{nombre}': {e}")
        return None


def drive_get_or_create_folder(nombre, parent_id=None):
    """Obtiene una carpeta o la crea si no existe. Devuelve el id o None."""
    folder_id = drive_find_folder(nombre, parent_id)
    if folder_id:
        return folder_id
    return drive_create_folder(nombre, parent_id)


def drive_upload_file(nombre, file_bytes, mime_type, parent_id):
    """Sube un archivo a Drive. Devuelve el id o None."""
    service = get_drive_service()
    if not service:
        return None
    try:
        from googleapiclient.http import MediaIoBaseUpload
        metadata = {'name': nombre}
        if parent_id:
            metadata['parents'] = [parent_id]
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
        file = service.files().create(body=metadata, media_body=media, fields='id').execute()
        return file.get('id')
    except Exception as e:
        logger.error(f"Error subiendo archivo Drive '{nombre}': {e}")
        return None


def drive_verify_file(file_id):
    """Verifica que un archivo existe en Drive. Devuelve True/False."""
    service = get_drive_service()
    if not service:
        return False
    try:
        service.files().get(fileId=file_id, fields='id').execute()
        return True
    except Exception as e:
        logger.error(f"Error verificando archivo Drive {file_id}: {e}")
        return False


# ─────────────────────────────────────────────
# GMAIL — LECTURA
# ─────────────────────────────────────────────

def gmail_get_unread(max_results=20):
    """Obtiene correos no leídos de las últimas 48h. Devuelve lista de {id, threadId}."""
    service = get_gmail_service()
    if not service:
        return []
    try:
        desde = (datetime.now() - timedelta(hours=48)).strftime('%Y/%m/%d')
        result = service.users().messages().list(
            userId='me',
            q=f'is:unread after:{desde}',
            maxResults=max_results
        ).execute()
        return result.get('messages', [])
    except Exception as e:
        logger.error(f"Error obteniendo correos no leídos: {e}")
        return []


def gmail_get_message(msg_id):
    """Obtiene el mensaje completo por id."""
    service = get_gmail_service()
    if not service:
        return None
    try:
        return service.users().messages().get(userId='me', id=msg_id, format='full').execute()
    except Exception as e:
        logger.error(f"Error obteniendo mensaje {msg_id}: {e}")
        return None


def gmail_mark_read(msg_id):
    """Marca un mensaje como leído."""
    service = get_gmail_service()
    if not service:
        return
    try:
        service.users().messages().modify(
            userId='me', id=msg_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
    except Exception as e:
        logger.error(f"Error marcando como leído {msg_id}: {e}")


def gmail_reply(msg_id, to, subject, body):
    """Responde a un mensaje en el mismo hilo. Devuelve True/False."""
    service = get_gmail_service()
    if not service:
        return False
    try:
        original = gmail_get_message(msg_id)
        if not original:
            return False
        thread_id = original.get('threadId', '')
        headers = {h['name']: h['value'] for h in original.get('payload', {}).get('headers', [])}
        message_id_header = headers.get('Message-ID', '')

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject if subject.startswith('Re:') else f'Re: {subject}'
        msg['From'] = GMAIL_USER
        msg['To'] = to
        if message_id_header:
            msg['In-Reply-To'] = message_id_header
            msg['References'] = message_id_header

        html_body = (
            '<html><body style="font-family:Georgia,serif;font-size:14px;">'
            + body.replace('\n', '<br>')
            + '<br><br><hr style="border:none;border-top:1px solid #ddd;">'
            '<em style="color:#666;">Secretaría AP Estudio Jurídico</em>'
            '</body></html>'
        )
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(
            userId='me',
            body={'raw': raw, 'threadId': thread_id}
        ).execute()
        return True
    except Exception as e:
        logger.error(f"Error enviando respuesta al mensaje {msg_id}: {e}")
        return False


def gmail_get_attachments(msg):
    """Extrae adjuntos de un mensaje. Devuelve lista de {nombre, bytes, mime}."""
    service = get_gmail_service()
    if not service or not msg:
        return []
    adjuntos = []
    try:
        def _procesar_parte(parte):
            filename = parte.get('filename', '')
            mime = parte.get('mimeType', '')
            body = parte.get('body', {})
            if filename and body.get('attachmentId'):
                att = service.users().messages().attachments().get(
                    userId='me', messageId=msg['id'], id=body['attachmentId']
                ).execute()
                data = base64.urlsafe_b64decode(att['data'])
                adjuntos.append({'nombre': filename, 'bytes': data, 'mime': mime})
            for sub in parte.get('parts', []):
                _procesar_parte(sub)

        payload = msg.get('payload', {})
        _procesar_parte(payload)
    except Exception as e:
        logger.error(f"Error extrayendo adjuntos: {e}")
    return adjuntos


# ─────────────────────────────────────────────
# ANÁLISIS DE RESOLUCIONES
# ─────────────────────────────────────────────

def extraer_texto_pdf(pdf_bytes):
    """Extrae texto de un PDF en bytes. Usa pdfminer.six."""
    try:
        from pdfminer.high_level import extract_text
        texto = extract_text(io.BytesIO(pdf_bytes))
        return texto.strip()
    except Exception as e:
        logger.error(f"Error extrayendo texto PDF: {e}")
        return ''


def extraer_num_procedimiento(texto):
    """Extrae número de procedimiento del texto (PA 112/2024, AS 234/2025, etc.)."""
    patron = r'\b(PA|AS|POA|JVP|DP|D\.P\.|EXP|ROJ)\s*\.?\s*(\d+/\d{4})\b'
    match = re.search(patron, texto, re.IGNORECASE)
    if match:
        return f"{match.group(1).upper()} {match.group(2)}"
    match2 = re.search(r'\b(\d{1,5}/\d{4})\b', texto)
    if match2:
        return match2.group(1)
    return 'SIN_NUM'


def analizar_resolucion_con_claude(texto_pdf, asunto, remitente):
    """
    Analiza una resolución judicial con Claude Haiku.
    Devuelve dict con tipo, juzgado, num_procedimiento, parte_dispositiva,
    plazos, proxima_actuacion, nombre_archivo.
    """
    prompt = (
        "Eres un asistente jurídico especializado en derecho español.\n"
        "Analiza la siguiente resolución judicial y responde EXCLUSIVAMENTE con un JSON válido,\n"
        "sin markdown, sin texto adicional.\n\n"
        f"Asunto del correo: {asunto}\n"
        f"Remitente: {remitente}\n\n"
        "Texto de la resolución:\n"
        f"{texto_pdf[:6000]}\n"
        f"{'...' if len(texto_pdf) > 6000 else ''}"
        f"{texto_pdf[-3000:] if len(texto_pdf) > 6000 else ''}\n\n"
        'INSTRUCCIONES CRÍTICAS para el campo "parte_dispositiva":\n'
        '- Extrae ÚNICAMENTE el contenido de la sección "PARTE DISPOSITIVA" o "FALLO" del documento.\n'
        '- NO resumas los antecedentes ni los fundamentos jurídicos.\n'
        '- Si el auto DENIEGA algo, el fallo debe decir claramente que se deniega.\n'
        '- Si el auto ESTIMA algo, el fallo debe decir claramente que se estima.\n'
        '- Copia literalmente las primeras líneas del fallo si es posible.\n\n'
        'JSON de salida (exactamente este formato):\n'
        '{\n'
        '  "tipo": "Auto|Sentencia|Decreto|Providencia|Diligencia",\n'
        '  "juzgado": "nombre del juzgado o tribunal",\n'
        '  "num_procedimiento": "PA 112/2024 o similar",\n'
        '  "parte_dispositiva": "texto literal o resumen fiel del fallo/parte dispositiva",\n'
        '  "plazos": [\n'
        '    {"dias": 5, "desde": "notificacion", "actuacion": "interponer recurso de reforma"}\n'
        '  ],\n'
        '  "proxima_actuacion": "descripción de la próxima actuación necesaria",\n'
        '  "nombre_archivo": "Auto_Deniega_recurso_reforma"\n'
        '}\n\n'
        'Si no hay plazos, "plazos" debe ser [].\n'
        '"nombre_archivo" sin espacios, máximo 10 palabras separadas por guiones bajos, sin extensión.'
    )
    try:
        response = claude_client.messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=1024,
            messages=[{'role': 'user', 'content': prompt}]
        )
        raw = response.content[0].text.strip()
        if raw.startswith('```'):
            raw = re.sub(r'^```[a-z]*\n?', '', raw)
            raw = re.sub(r'```$', '', raw).strip()
        return json.loads(raw)
    except json.JSONDecodeError as e:
        logger.error(f"Claude devolvió JSON inválido: {e}")
        return {
            'tipo': 'Resolución',
            'juzgado': 'Desconocido',
            'num_procedimiento': extraer_num_procedimiento(texto_pdf),
            'parte_dispositiva': 'No se pudo analizar automáticamente.',
            'plazos': [],
            'proxima_actuacion': 'Revisar manualmente.',
            'nombre_archivo': 'Resolucion_sin_analizar',
        }
    except Exception as e:
        logger.error(f"Error analizando resolución con Claude: {e}")
        return None


# ─────────────────────────────────────────────
# PROCESADO DE CORREOS (cada 2h, L-V)
# ─────────────────────────────────────────────

async def procesar_correos(bot):
    """Revisa correos no leídos, procesa resoluciones de procuradores."""
    global diario_secretaria, CORREOS_ACTIVOS
    if not CORREOS_ACTIVOS:
        logger.info("Revisión de correos desactivada.")
        return
    logger.info("Iniciando revisión de correos...")

    mensajes = gmail_get_unread(max_results=20)
    if not mensajes:
        logger.info("No hay correos no leídos.")
        return

    tz = pytz.timezone(TIMEZONE)

    for ref in mensajes:
        msg = gmail_get_message(ref['id'])
        if not msg:
            continue

        headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}
        remitente_raw = headers.get('From', '')
        asunto = headers.get('Subject', '(sin asunto)')

        match_email = re.search(r'<(.+?)>', remitente_raw)
        remitente_email = match_email.group(1).lower() if match_email else remitente_raw.lower().strip()

        if remitente_email in PROCURADORES:
            nombre_proc = PROCURADORES[remitente_email]
            adjuntos = gmail_get_attachments(msg)
            pdfs = [a for a in adjuntos if a['mime'] == 'application/pdf' or a['nombre'].lower().endswith('.pdf')]

            if not pdfs:
                diario_secretaria.append({
                    'tipo': 'procurador_sin_pdf',
                    'remitente': nombre_proc,
                    'asunto': asunto,
                })
                gmail_mark_read(ref['id'])
                continue

            for pdf in pdfs:
                texto_pdf = extraer_texto_pdf(pdf['bytes'])
                if not texto_pdf:
                    logger.warning(f"No se pudo extraer texto del PDF: {pdf['nombre']}")
                    continue

                analisis = analizar_resolucion_con_claude(texto_pdf, asunto, nombre_proc)
                if not analisis:
                    continue

                num_proc = analisis.get('num_procedimiento', extraer_num_procedimiento(texto_pdf))
                tipo = analisis.get('tipo', 'Resolución')
                fallo = analisis.get('parte_dispositiva', '')
                plazos = analisis.get('plazos', [])
                proxima = analisis.get('proxima_actuacion', '')
                nombre_base = analisis.get('nombre_archivo', 'Resolucion')

                num_proc_safe = re.sub(r'[^A-Za-z0-9]', '_', num_proc)
                nombre_pdf = f"{nombre_base}_{num_proc_safe}.pdf"

                eventos_creados = []
                for plazo in plazos:
                    try:
                        dias = int(plazo.get('dias', 0))
                        actuacion = plazo.get('actuacion', 'Actuación procesal')
                        desde = plazo.get('desde', 'notificación')
                        if dias > 0:
                            fecha_venc = (datetime.now(tz) + timedelta(days=dias)).date()
                            service_cal = get_calendar_service()
                            if service_cal:
                                event = {
                                    'summary': f'VENCIMIENTO - {num_proc} - {actuacion}',
                                    'description': f'Plazo: {dias} días desde {desde}\nActuación: {actuacion}\nProcurador: {nombre_proc}',
                                    'start': {'date': str(fecha_venc)},
                                    'end': {'date': str(fecha_venc)},
                                }
                                service_cal.events().insert(calendarId='primary', body=event).execute()
                                eventos_creados.append(f"{actuacion} ({fecha_venc})")
                    except Exception as e:
                        logger.error(f"Error creando evento calendario: {e}")

                drive_ok = False
                file_id_drive = None
                try:
                    carpeta_raiz_id = drive_get_or_create_folder(DRIVE_FOLDER_RESOLUCIONES)
                    if carpeta_raiz_id:
                        carpeta_proc_id = drive_get_or_create_folder(num_proc_safe, carpeta_raiz_id)
                        if carpeta_proc_id:
                            file_id_drive = drive_upload_file(
                                nombre_pdf, pdf['bytes'], 'application/pdf', carpeta_proc_id
                            )
                            if file_id_drive:
                                drive_ok = drive_verify_file(file_id_drive)
                except Exception as e:
                    logger.error(f"Error en flujo Drive: {e}")

                recibido_ok = False
                if drive_ok:
                    recibido_ok = gmail_reply(ref['id'], remitente_email, asunto, 'Recibido.')

                plazos_txt = '\n'.join([
                    f"  - {p.get('dias')}d desde {p.get('desde')}: {p.get('actuacion')}"
                    for p in plazos
                ]) if plazos else '  (ninguno detectado)'

                eventos_txt = '\n'.join([f'  - {e}' for e in eventos_creados]) if eventos_creados else '  (ninguno)'

                telegram_msg = (
                    f'📬 <b>Nueva resolución de {nombre_proc}</b>\n'
                    f'<b>Asunto:</b> {asunto}\n'
                    f'<b>Tipo:</b> {tipo}\n'
                    f'<b>Procedimiento:</b> {num_proc}\n'
                    f'<b>Juzgado:</b> {analisis.get("juzgado", "—")}\n\n'
                    f'<b>Fallo:</b>\n{fallo}\n\n'
                    f'<b>Plazos:</b>\n{plazos_txt}\n\n'
                    f'<b>Próxima actuación:</b> {proxima}\n\n'
                    f'<b>Eventos en calendario:</b>\n{eventos_txt}\n'
                    f'<b>Guardado en Drive:</b> {"✅" if drive_ok else "❌"}\n'
                    f'<b>Recibido enviado:</b> {"✅" if recibido_ok else "❌"}'
                )
                try:
                    await bot.send_message(chat_id=5682841007, text=telegram_msg, parse_mode='HTML')
                except Exception as e:
                    logger.error(f"Error enviando notificación Telegram: {e}")

                diario_secretaria.append({
                    'tipo': 'resolucion',
                    'remitente': nombre_proc,
                    'asunto': asunto,
                    'analisis': analisis,
                    'num_proc': num_proc,
                    'drive_ok': drive_ok,
                    'recibido_ok': recibido_ok,
                    'eventos': eventos_creados,
                })

        else:
            diario_secretaria.append({
                'tipo': 'otro',
                'remitente': remitente_raw,
                'asunto': asunto,
            })

        gmail_mark_read(ref['id'])

    logger.info(f"Revisión de correos completada. Procesados: {len(mensajes)}")


# ─────────────────────────────────────────────
# DIARIO DE SECRETARÍA (19h, L-V)
# ─────────────────────────────────────────────

async def enviar_diario_secretaria(bot):
    """Envía el diario diario de correos procesados y resetea la lista."""
    global diario_secretaria

    hoy = datetime.now(pytz.timezone(TIMEZONE)).strftime('%d/%m/%Y')
    resoluciones = [e for e in diario_secretaria if e.get('tipo') == 'resolucion']
    procurador_sin_pdf = [e for e in diario_secretaria if e.get('tipo') == 'procurador_sin_pdf']
    otros = [e for e in diario_secretaria if e.get('tipo') == 'otro']
    total = len(diario_secretaria)

    lineas = [f'<h2>Diario de Secretaría — {hoy}</h2>']
    lineas.append(f'<p><b>Total de correos procesados:</b> {total}</p>')

    if resoluciones:
        lineas.append('<h3>📄 Resoluciones judiciales</h3>')
        for r in resoluciones:
            an = r.get('analisis', {})
            plazos = an.get('plazos', [])
            plazos_html = ('<ul>' + ''.join(
                f"<li>{p.get('dias')}d desde {p.get('desde')}: {p.get('actuacion')}</li>"
                for p in plazos
            ) + '</ul>') if plazos else '<p>(sin plazos)</p>'
            eventos_html = ('<ul>' + ''.join(
                f'<li>{e}</li>' for e in r.get('eventos', [])
            ) + '</ul>') if r.get('eventos') else '<p>(sin eventos)</p>'
            lineas.append(
                '<hr>'
                f'<p><b>De:</b> {r["remitente"]}<br>'
                f'<b>Asunto:</b> {r["asunto"]}<br>'
                f'<b>Tipo:</b> {an.get("tipo", "—")}<br>'
                f'<b>Procedimiento:</b> {r["num_proc"]}<br>'
                f'<b>Juzgado:</b> {an.get("juzgado", "—")}<br>'
                f'<b>Fallo:</b> {an.get("parte_dispositiva", "—")}</p>'
                f'<b>Plazos:</b>{plazos_html}'
                f'<b>Próxima actuación:</b> {an.get("proxima_actuacion", "—")}<br>'
                f'<b>Eventos calendario:</b>{eventos_html}'
                f'<b>Drive:</b> {"✅ Guardado" if r["drive_ok"] else "❌ Error"} | '
                f'<b>Recibido:</b> {"✅ Enviado" if r["recibido_ok"] else "❌ No enviado"}'
            )

    if procurador_sin_pdf:
        lineas.append('<h3>📭 Correos de procuradores sin PDF adjunto</h3><ul>')
        for r in procurador_sin_pdf:
            lineas.append(f'<li><b>{r["remitente"]}</b> — {r["asunto"]}</li>')
        lineas.append('</ul>')

    if otros:
        lineas.append('<h3>📥 Otros correos recibidos</h3><ul>')
        for r in otros:
            lineas.append(f'<li><b>{r["remitente"]}</b> — {r["asunto"]}</li>')
        lineas.append('</ul>')

    if not diario_secretaria:
        lineas.append('<p><em>No se procesaron correos hoy.</em></p>')

    lineas.append('<br><hr><em>Secretaría AP Estudio Jurídico</em>')
    cuerpo_html = '\n'.join(lineas)

    try:
        service = get_gmail_service()
        if service:
            from email.mime.multipart import MIMEMultipart as _MM
            from email.mime.text import MIMEText as _MT
            msg_root = _MM('alternative')
            msg_root['Subject'] = f'📋 Diario de Secretaría — {hoy}'
            msg_root['From'] = GMAIL_USER
            msg_root['To'] = GMAIL_USER
            msg_root.attach(_MT(cuerpo_html, 'html', 'utf-8'))
            raw = base64.urlsafe_b64encode(msg_root.as_bytes()).decode()
            service.users().messages().send(userId='me', body={'raw': raw}).execute()
            logger.info("Diario de secretaría enviado.")
    except Exception as e:
        logger.error(f"Error enviando diario de secretaría: {e}")

    try:
        await bot.send_message(
            chat_id=5682841007,
            text=f'📋 Diario de secretaría enviado ({hoy}). Total correos: {total}, resoluciones: {len(resoluciones)}.',
        )
    except Exception as e:
        logger.error(f"Error notificando Telegram (diario): {e}")

    diario_secretaria = []


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    scheduler = AsyncIOScheduler(timezone=pytz.timezone(TIMEZONE))

    async def post_init(application: Application) -> None:
        loop = asyncio.get_running_loop()

        def run_daily():
            loop.create_task(daily_summary(application.bot))

        def run_weekly():
            loop.create_task(weekly_summary(application.bot))

        def run_reminders():
            loop.create_task(appointment_reminders(application.bot))

        def run_correos():
            loop.create_task(procesar_correos(application.bot))

        def run_diario():
            loop.create_task(enviar_diario_secretaria(application.bot))

        scheduler.add_job(run_daily,     'cron', hour=7,  minute=0, day_of_week='mon-fri')
        scheduler.add_job(run_weekly,    'cron', hour=9,  minute=0, day_of_week='sat')
        scheduler.add_job(run_reminders, 'cron', hour=8,  minute=0, day_of_week='mon-fri')
        scheduler.add_job(run_correos,   'cron', hour='8,10,12,14,16,18,20', minute=0, day_of_week='mon-fri')
        scheduler.add_job(run_diario,    'cron', hour=19, minute=0, day_of_week='mon-fri')

        scheduler.start()
        logger.info("✅ Bot Secretaria iniciado.")

    app = Application.builder().token(TELEGRAM_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("id",      cmd_id))
    app.add_handler(CommandHandler("resumen", cmd_resumen))
    app.add_handler(CommandHandler("correos", cmd_correos))
    app.add_handler(CommandHandler("correos_on", cmd_correos_on))
    app.add_handler(CommandHandler("correos_off", cmd_correos_off))
    app.add_handler(CommandHandler("bbdd",    cmd_bbdd))
    app.add_handler(CommandHandler("initbbdd", cmd_initbbdd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()
