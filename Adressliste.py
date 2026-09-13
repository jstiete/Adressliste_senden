#!/usr/bin/env python
# coding: utf-8
# /// script
# dependencies = [
#   "pandas >= 0.21",
#   "openpyxl >= 3.0",
# ]
# ///
'''
Adressliste.py
Das Skript enthält Funktionen, um eine Adressliste automatisiert zu verwalten:
 - Gespeichterte Daten an die jeweilige Person senden (E-Mail).
 - Komplette Liste als Dateianhang senden.

Author: (c) 2026 jstiete (https://github.com/jstiete)
Licence: MIT
'''

import argparse
import configparser
import logging
import sys
import pathlib
from typing import Optional
import pandas as pd
import re
import time

import smtplib
import ssl
import certifi
from email.message import EmailMessage
import mimetypes

# Logger für Konsole erstellen:
logger = logging.getLogger("Adressliste")
logger.propagate = False
logFormatter = logging.Formatter('%(asctime)s\t- %(levelname)s\t- %(filename)s:(%(lineno)d) - %(message)s')
console_handler = logging.StreamHandler(stream=sys.stdout)
console_handler.setFormatter(logFormatter)
logger.addHandler(console_handler)

class ConfigData:
    """Speichert alle Konfigurationsparameter für SMTP, E‑Mail‑Texte und Tabellenstruktur, etc."""
    def __init__(
            self,
            smtp_host: str,
            smtp_port: int = 587,
            *,
            username: Optional[str] = None,
            password: Optional[str] = None,
            use_tls: bool = True,
            use_ssl: bool = False,
            timeout: Optional[float ] = 10.0,
            from_addr: str,
            reply_to: Optional[str] = None,
            subject: str,
            body_plaintext_userdata: str,
            body_html_userdata: Optional[str] = None,
            body_plaintext_complete_list: str,
            body_html_complete_list: Optional[str] = None,
            header_row:int = 1,
            send_log_to: Optional[bool] = None
    ):

        if use_ssl and use_tls:
            raise ValueError("use_ssl und use_tls schließen sich aus. Bitte nur eines setzen.")

        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.username = username
        self.password = password
        self.use_tls = use_tls
        self.use_ssl = use_ssl
        self.timeout = timeout
        self.from_addr = from_addr
        self.reply_to = reply_to
        self.subject = subject
        self.body_plaintext_userdata = body_plaintext_userdata
        self.body_html_userdata = body_html_userdata
        self.body_plaintext_complete_list = body_plaintext_complete_list
        self.body_html_complete_list = body_html_complete_list
        self.header_row = header_row
        self.send_log_to = send_log_to


def add_email_loghandler(config:ConfigData):
    """ E-Mail Logging Handler hinzufügen,
    um Benachrichtigungen über Fehler und Warnungen per E-Mail zu erhalten."""
    try:
        from email_loghandler import make_email_handler
    except ImportError:
        logger.warning("E-Mail Logging Handler konnte nicht importiert werden.'pip install -r requirements.txt' oder 'uv sync' ausführen.")
        return 1

    mail_formatter = logging.Formatter(
        fmt="%(asctime)s\t- %(levelname)s\t- %(filename)s:(%(lineno)d) - %(message)s",
        datefmt="%YYYY-%mm-%dd %HH:%MM:%SS",
    )

    handler = make_email_handler(
        smtp_server=config.smtp_host,
        smtp_port=config.smtp_port,
        use_tls=config.use_tls,
        use_ssl=config.use_ssl,
        username=config.username,
        password=config.password,
        from_addr=config.from_addr,
        to_addrs=[config.send_log_to],
        subject_template="[{levelname_max}] {program_name} @ {hostname} – {count} entries (from {min_level})",
        header=(
            "Hello,\n\n"
            "here is the summarized log report:\n"
            "Program: {program_name}\nHost: {hostname}\n"
            "Time: {date:%Y-%m-%d %H:%M:%S}\n"
            "Threshold: {min_level}\n"
            "Count: {count}\n"
        ),
        footer=(
            "\n---\n"
            "Automatically generated. Highest level: {levelname_max}\n"
            "Best regards"
        ),
        level=logging.INFO,                    # Alle Logs >= INFO in der Mail
        attachment_min_level=logging.DEBUG,    # Alle Logs >= DEBUG als Anhang. (Wenn logger.level <= DEBUG)
        formatter=mail_formatter,
        send_if_empty=False,         # Nur senden, wenn Logs oberhalb  von 'level' auftreten.
        max_buffer=None,             # optional: e.g. send immediately at 500 entries
        register_atexit=True,        # automatically send mail at program end (off in debugger)
    )
    handler.set_name("email_logging")
    logger.addHandler(handler)


class SMTPConnection:
    """Wrapper für eine robuste SMTP‑Verbindung und Re-Connect mit Context-Manager."""
    def __init__(self, config:ConfigData):
        self.config = config
        self.server = None

    def _connect(self):
        """ Verbindung zum SMTP Server herstellen."""
        try:
            context = ssl.create_default_context(cafile=certifi.where())
            if self.config.use_ssl:
                self.server = smtplib.SMTP_SSL(self.config.smtp_host, self.config.smtp_port, timeout=self.config.timeout, context=context)
            else:
                self.server = smtplib.SMTP(self.config.smtp_host, self.config.smtp_port, timeout=self.config.timeout)
            self.server.ehlo()
            if self.config.use_tls:
                self.server.starttls(context=context)
                self.server.ehlo()
            if self.config.username:
                self.server.login(self.config.username, self.config.password or "")
            logger.debug("Connected to SMTP server")
        except Exception as e:
            logger.error(f"Fehler beim Verbindungsaufbau zum SMTP Server: {e}")
            raise e

    def _reconnect(self):
        try:
            self.server.quit()
        except Exception:
            pass
        time.sleep(self.config.timeout)
        self._connect()

    def send_message(self, msg):
        """ Nachricht senden, falls Verbindung getrennt ist, wird automatisch neu verbunden."""
        try:
            self.server.send_message(msg)
        except smtplib.SMTPServerDisconnected:
            logger.warning("SMTP disconnected. Reconnecting.")
            self._reconnect()
            self.server.send_message(msg)

    def __enter__(self):
        """ Context-Manager betreten. """
        self._connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """ Context-Manager verlassen. """
        try:
            self.server.quit()
        except Exception:
            pass


def read_excel_as_clean_strings(file_path, date_format='%d.%m.%Y'):
    """
    Liest eine Excel-Datei generisch ein. Datumsspalten werden im Wunschformat
    als String formatiert, alle anderen Spalten werden ebenfalls zu Strings.
    Leere Zellen werden zu ''.
    """
    # Rohdaten einlesen (Leere Zeilen durch '' ersetzen)
    df = pd.read_excel(file_path, keep_default_na=False, na_values=[])

    # Generischer Loop über alle Spalten
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            # Datumsspalten werden im gewünschten Format als String formatiert
            df[col] = df[col].dt.strftime(date_format)
        else:
            # Alles andere wird als String konvertiert
            df[col] = df[col].astype(str)

    # NaT/NaN Artefakte bereinigen
    return df.replace(['NaT', 'nan', '<NaT>'], '')


def main(addressfile:pathlib.Path, send_list:bool, config:ConfigData):
    """
    Hauptfunktion zur Verarbeitung der Adressliste.
      1. Excel‑Datei einlesen
      2. Datumsfelder bereinigen
      3. E‑Mail‑Spalten erkennen (`E-Mail*`)
      4. Für jede Datenzeile (je nach Angabe von sendlist)
          - E‑Mail‑Nachricht erstellen
          - Platzhalter ersetzen
          - Optional Anhang hinzufügen
          - Nachricht versenden
      5. Logging der Ergebnisse
    """
    try:
        addressdata = pd.read_excel(addressfile,
                                    #sheet_name="Tabellenblatt",
                                    header= config.header_row-1,
                                    dtype=str,              # Deaktiviert das Umwandeln von Zahlen
                                    parse_dates=False,      # Deaktiviert das Umwandeln von Datumsangaben
                                    keep_default_na=False,  # Deaktiviert die Standard-NaN-Erkennung
                                    na_values=[]  # Definiert keine zusätzlichen Werte als NaN
        ) # ggf. engine='openpyxl'

        # Hack:
        # Datumsangaben (21.08.1950) aus Excel werden bei Excel intern als 21.08.1950 00:00:00 gespeichert.
        # Daher bei allen Datumsangaben, wo der Zeitstempel 00:00:00 ist, die Uhrzeit entfernen.
        # Für das gesamte DataFrame alle " 00:00:00" Zeitstempel am Ende entfernen
        # Alternativ read_excel_as_clean_strings() aufrufen.
        addressdata = addressdata.astype(str).replace(r'\s+00:00:00$', '', regex=True)

    except FileNotFoundError:
        logger.error(f"{addressfile.name} nicht gefunden.")
        raise

    attachment_data = None
    if send_list:
        with open(addressfile, "rb") as f:
            attachment_data = f.read()

    # get all columns with e-mail addresses
    r = re.compile("E-Mail.*", re.IGNORECASE)
    email_addresses = list(filter(r.match, addressdata.columns))
    if len(email_addresses) == 0:
        logger.error(f"Keine gültigen E-Mail-Adressen gefunden. Mindestens ein Spaltennamen muss mit 'E-Mail' beginnen.")
        raise ValueError("Keine gültigen E-Mail-Adressen gefunden.")

    # initialize server connection
    with SMTPConnection(config) as smtp:
        send_mails = 0
        mail_errors = 0
        for row, data in enumerate(addressdata.to_dict(orient="records"), start=1):
            try:
                msg = EmailMessage()
                msg["From"] = config.from_addr
                msg['Subject'] = config.subject.format(**data)
                if config.reply_to is not None:
                    msg["Reply-To"] = config.reply_to
                msg["To"] = ", ".join(str(data[x]).strip() for x in email_addresses if pd.notna(data[x]))

                if msg["To"] == "":
                    logger.warning(f"Überspringe Datenreihe {row}, da keine gültige E-Mail-Adresse gefunden wurde.")
                    mail_errors += 1
                    continue

                if send_list:
                    body_text = config.body_plaintext_complete_list.format(**data)
                    msg.set_content(body_text, subtype="plain", charset="utf-8")
                    if config.body_html_complete_list:
                        body_html = config.body_html_complete_list.format(**data)
                        msg.add_alternative(body_html, subtype="html", charset="utf-8")

                    mime_type, _ = mimetypes.guess_type(addressfile.name)
                    maintype, subtype = mime_type.split("/", 1)
                    # Add attachment to message
                    msg.add_attachment(
                        attachment_data,
                        maintype=maintype,
                        subtype=subtype,
                        filename=addressfile.name
                        )
                else:
                    body_text = config.body_plaintext_userdata.format(**data)
                    msg.set_content(body_text, subtype="plain", charset="utf-8")
                    if config.body_html_userdata:
                        body_html = config.body_html_userdata.format(**data)
                        msg.add_alternative(body_html, subtype="html", charset="utf-8")

                smtp.send_message(msg)
                logger.debug(f"E-Mail an {msg['To']} gesendet. (Datenreihe {row})")
                send_mails += 1

            except Exception as e:
                    mail_errors += 1
                    logger.error(f"Fehler bei Datenreihe {row}: {e}")
        logger.info(f"{send_mails} E-Mails erfolgreich gesendet. {mail_errors} Fehler.")


# Beispiel Konfiguration:
# ConfigParser entfernt führende spaces oder tabs in der Zeile. So können Werte über mehrere Zeilen
# eingegeben werden.
config_example = (
"[SMTP]\n"
"from_addr=noreply@example.com\n"
"smtp_host=smtp.example.com\n"
";<optional parameters>\n"
"smtp_port=587\n"
"username=user@example.com\n"
"password=***PASSWORD***\n"
"use_tls=True\n"
"use_ssl=False\n"
"timeout=10\n\n"
"[E-MAIL]\n"
";<Alle Spaltennamen der Tabelle können als Platzhalter verwendet werden.>\n"
"subject=Aktuelle Adressliste\n"
"text_actual_data = Hallo {Vorname},\n"
                   "    aktuell sind folgende Daten von dir gespeichert:\n"
                   "    Name, Vorname: {Name}, {Vorname}\n"
                   "    E-Mail (privat): {E-Mail privat}\n"
                   "    E-Mail 2: {E-Mail 2}\n"
                   "    Telefon: {Telefon}\n"
                   "    Geburtsdatum: {Geburtsdatum}\n    \n"
                   "    Wenn die Daten nicht mehr aktuell sind, bitte aktuelle Daten an mich senden.\n"
                   "    Falls du aus der Liste entfernt werden möchtest, ebenfalls melden.\n"
                   "    Viele Grüße\n\n"
"html_actual_data = <html>\n"
"                     <body>\n"
"                      <p>Hallo {Vorname},<br>\n"
"                         aktuell sind folgende Daten von dir gespeichert:<br>\n"
"                         Name, Vorname: {Name}, {Vorname}<br>\n"
"                         E-Mail (privat): {E-Mail privat}<br>\n"
"                         E-Mail 2: {E-Mail 2}<br>\n"
"                         Telefon: {Telefon}<br>\n"
"                         Geburtsdatum: {Geburtsdatum}</p>\n"
"                      <p>Wenn die Daten nicht mehr aktuell sind, bitte aktuelle Daten an mich senden.<br>\n"
"                         Falls du aus der Liste entfernt werden möchtest, ebenfalls melden.<br>\n"
"                         Viele Grüße</p>\n"
"                    </body>\n"
"                  </html>\n\n"
"text_actual_list = Hallo {Vorname},\n "
                   "    im Anhang ist die aktuelle Adressliste.\n"
                   "    Alternativ kann sie auch über den folgenden Link heruntergeladen werden:\n"
                   "    http://mycloud.net/adresslist\n    \n"
                   "    Wenn deine Daten nicht mehr aktuell sind, bitte aktuelle Daten an mich senden.\n"
                   "    Falls du aus der Liste entfernt werden möchtest, ebenfalls melden.\n"
                   "    Viele Grüße\n\n"
"html_actual_list = <html>\n"
"                    <body>\n"
"                      <p>Hallo {Vorname},<br>\n"
"                         im Anhang ist die aktuelle Adressliste.<br>\n"
"                         Alternativ kann sie auch über den folgenden Link heruntergeladen werden:<br>\n"
"                         <a href=\"http://mycloud.net/adresslist\">http://mycloud.net/adresslist</a></p>\n"
"                      <p>Wenn deine Daten nicht mehr aktuell sind, bitte aktuelle Daten an mich senden.<br>\n"
"                         Falls du aus der Liste entfernt werden möchtest, ebenfalls melden.<br>\n"
"                         Viele Grüße</p>\n"
"                    </body>\n"
"                  </html>\n\n"
"[TABLE]\n"
"headline_row=3\n\n"
";<Diese Sektion ist optional.>\n"
"[EMAIL_LOGGING]\n"
"send_log_to=user@example.com\n"
)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("-f", "--file", default='./Testdaten.xlsx',
                        help="Pfad der Adressliste als xls oder xlsx Datei.")
    parser.add_argument("-c", "--config", default='./config.ini',
                        help="Konfiguration und E-Mail Zugang. (Siehe Beispiel unten.)")
    parser.add_argument("-l", "--sendlist", action='store_true', help="Sende die gesamte Liste. Wenn nicht angegeben, werden nur die persönlichen Daten an die Person gesendet.")
    parser.add_argument("--trace", default="info", choices=["warning", "info", "debug"], help="Logging level")
    parser.description = "Beispielkonfiguration:\n\n" + config_example
    args = parser.parse_args()

    print(args.sendlist)

    # logger.info(...)
    #   ↓
    # Logger - Level prüfen (an Handler weiterleiten oder verwerfen)
    #   ↓
    # Handler - Level prüfen (verarbeiten oder verwerfen)
    #   ↓
    # Ausgabe
    logger.setLevel(logging.DEBUG)                   #Logger auf das kleinste Level setzen.
    match str.lower(args.trace):                     #Ausgaben auf Handler-Ebene steuern.
        case "debug":
            console_handler.setLevel(logging.DEBUG)
        case "info":
            console_handler.setLevel(logging.INFO)
        case _:
            console_handler.setLevel(logging.WARNING)

    config = configparser.ConfigParser(delimiters=('='))
    configfile = pathlib.Path(args.config)

    try:
        if configfile.exists():
            config.read(configfile, encoding='utf-8')
        else:
            logger.error("Kann Config-Datei nicht lesen. Siehe 'python Adressliste.py --help'")
            with open("./config_template.ini", 'w', encoding='utf-8') as configtemplate:
                configtemplate.write(config_example)
            raise FileNotFoundError(f"Kann {configfile} nicht lesen.")

        #parse all optional values:
        kwargs = {}
        if (value := config.getint("SMTP", "smtp_port", fallback=None)) is not None:
            kwargs["smtp_port"] = value
        if (value := config.get("SMTP", "username", fallback=None)) is not None:
            kwargs["username"] = value
        if (value := config.get("SMTP", "password", fallback=None)) is not None:
            kwargs["password"] = value
        if (value := config.getboolean("SMTP", "use_tls", fallback=None)) is not None:
            kwargs["use_tls"] = value
        if (value := config.getboolean("SMTP", "use_ssl", fallback=None)) is not None:
            kwargs["use_ssl"] = value
        if (value := config.getfloat("SMTP", "timeout", fallback=None)) is not None:
            kwargs["timeout"] = value
        if (value := config.get("SMTP", "reply_to", fallback=None)) is not None: kwargs["reply_to"] = value
        if (value := config.get("E-MAIL", "html_actual_data", fallback=None)) is not None: kwargs["body_html_userdata"] = value
        if (value := config.get("E-MAIL", "html_actual_list", fallback=None)) is not None: kwargs["body_html_complete_list"] = value
        if (value := config.getint("TABLE", "headline_row", fallback=None)) is not None: kwargs["header_row"] = value
        if (value := config.get("EMAIL_LOGGING", "send_log_to", fallback=None)) is not None: kwargs["send_log_to"] = value

        config_data=ConfigData(smtp_host=str(config["SMTP"]["smtp_host"]),
                               from_addr=str(config["SMTP"]["from_addr"]),
                               subject=str(config["E-MAIL"]["subject"]),
                               body_plaintext_userdata=str(config["E-MAIL"]["text_actual_data"]),
                               body_plaintext_complete_list=str(config["E-MAIL"]["text_actual_list"]),
                               **kwargs)
    except Exception as e:
        logger.error(f"Fehler beim Einlesen der Konfiguration: {e}", exc_info=True) #log with complete traceback
        sys.exit(1)

    if config_data.send_log_to:
        add_email_loghandler(config_data)
        if h:= logging.getHandlerByName("email_logging"):
            h.setLevel(console_handler.level)

    main(addressfile=pathlib.Path(args.file), send_list=args.sendlist, config=config_data)

    sys.exit(0)