# Adressliste – Automatisiertes Versenden von Adresslisten per E‑Mail

Dieses Projekt enthält ein Python‑Skript, das eine Excel‑Adressliste einliest und
automatisiert personalisierte E‑Mails an alle darin enthaltenen Personen versendet.
Alternativ kann zusätzlich die komplette Adressliste als Datei‑Anhang verschickt werden.

1. Informieren aller Personen über ihre gespeicherte Daten.  
   Z.B.: "Hallo, wir haben folgende Daten von dir gespeichert: Name, E-Mail Adresse, Wohnort."
2. Versenden der gesamten Adressliste als Dateianhang an alle Personen.  
   Z.B.: "Hallo, im Anhang findest du die aktuelle Adressliste."

Das Skript eignet sich besonders für Vereine, Gruppen oder Organisationen, die regelmäßig
aktualisierte Kontaktdaten verteilen möchten.

---

## ✨ Funktionen

- Einlesen von Excel‑Dateien (xls/xlsx)
- Automatische Bereinigung von Datumsfeldern und leeren Zellen
- Versand personalisierter E‑Mails basierend auf Platzhaltern
- Optionaler Versand der kompletten Adressliste als Anhang (siehe `--sendlist`)
- TLS/SSL‑Unterstützung für SMTP‑Server
- Robuste SMTP‑Verbindung mit automatischer Wiederverbindung
- Konfigurierbares Logging (Konsole + optional E‑Mail‑Logging)
- Vollautomatisierbar über Cron/Anacron

---
1. [📦 Installation & Voraussetzungen](#installation--voraussetzungen)
2. [🚀 Quickstart](#quickstart)
3. [️⚙️ Konfiguration](#konfiguration)
4. [🔄 Platzhalter / Wildcards / Überschriften](#platzhalter--wildcards--überschriften)
5. [❗E-Mail Adressen](#e-mail-adressen)
6. [ ▶️ Nutzung / CLI-Interface](#nutzung--cli-interface)
7. [🕒 Automatischer Versand (Cron / Anacron)](#automatischer-versand-cron--anacron)
8. [🛠️ Testmodus](#testmodus)
9. [🧩 Ablauf](#ablauf)
10. [⚖️ Lizenz](#lizenz)


---

## 📦 Installation & Voraussetzungen
Projektdateien herunterladen oder ```git clone https://github.com/jstiete/Adressliste_senden```

### Native Python‑Installation
Benötigt werden:
- Python 3.10 oder neuer
- Abhängigkeiten:
  - `pandas`
  - `openpyxl`
  - `certifi`

Installation der Abhängigkeiten:
```bash
pip install -r requirements.txt
```

### Alternativ: Installation mit uv
```bash
uv sync
```

## 🚀 Quickstart
Für einen ersten Test wird das aiosmtpd Modul aus dem Abschnitt [Testmodus](#testmodus) verwendet.
Zunächst die Schritte im Abschnitt [Installation & Voraussetzungen](#--installation--voraussetzungen)
durchführen.

Anschließend eine Konsole starten und in den Ordner navigieren, in dem die Daten des Repos liegen.

`aiosmtpd` installieren und starten:
```shell
$ python -m pip install aiosmtpd
$ python -m aiosmtpd -n
```

Eine zweite Konsole im Repo öffnen und das Python Skript mit den Testdaten starten:
```shell
$ python Adressliste.py -f ./Testdaten.xlsx -c ./config_aiosmtpd.ini 
```
bzw.
```shell
$ python Adressliste.py -f ./Testdaten.xlsx -c ./config_aiosmtpd.ini --sendlist
```

Nun sollten in der ersten Konsole die E-Mails an alle Empfänger als Rohdaten erscheinen.

## ⚙️ Konfiguration
Die Konfiguration erfolgt über eine INI‑Datei.
Die Struktur entspricht dem Beispiel in `config_template.ini`.

Die folgenden Zeilen geben die Konfiguration und Verbindungsdaten des SMTP-Servers an
und sind im Allgemeinen beim Anbieter zu finden.
```ini
[SMTP]
from_addr=
    E-Mail Adresse, von der die Listen gesendet werden sollen.
smtp_host=
    URL des SMTP-Servers

smtp_port= (optional, default: 587):
    Port vom SMTP server.
username= (optional)
    Benutzername für SMTP Authentifikation.
password= (optional)
    Passwort für Benutzerauthentifikation am SMTP Server.
use_tls= (True/False bzw. 1/0) (optional, default=True)
    TLS Verschlüsselung für den Verbindungsaufbau.
use_ssl= (True/False, bzw. 1/0) (optional, default=False)
    SSL Verschlüsselung für den Verbindungsaufbau.
timeout= (optional, default=10)
    Timeout für den Verbindungsaufbau in Sekunden.
reply_to= (optional)
    Antwort Adresse, falls diese von der Sender-Adresse abweichen soll. (Wird beim Empfänger häufiger als Spam erkannt.)
```

Konfiguration des E-Mail Textes mit Platzhaltern. (siehe Abschnitt [Platzhalter](#platzhalter--wildcards--überschriften))  
Zeilenumbrüche können erstellt werden, indem eine neue Zeile mit mindestens zwei Leerzeichen
oder einem Tabulator-Zeichen begonnen werden.
```ini
[E-MAIL]
subject=
    Betreffzeile der E-Mail.
text_actual_data=
    Plain-Text Variante für den E-Mail Text der gespeicherten Daten.
text_actual_list=
    Plain-Text Variante für den E-Mail beim Versenden der gesamten Liste.
    
html_actual_data= (optional)
    HTML-Text Variante für den E-Mail Text der gespeicherten Daten.
html_actual_list= (optional)
    HTML-Text Variante für den E-Mail beim Versenden der gesamten Liste.
```
Einstellungen für die zu lesende Tabelle:
```ini
[TABLE]
headline_row= (optional, default=1)
    Tabellenzeile, welche die Überschriften enthält. Zählweise beginnend bei 1.
```
Optional: Log-Nachrichten per E-Mail senden.
```ini
[EMAIL_LOGGING] (optional)
send_log_to=
    E-Mailadresse des Empfängers. Mehrere Adressen können per Komma getrennt werden.
```

## 🔄 Platzhalter / Wildcards / Überschriften
Alle Überschriften (Spaltennamen aus der konfigurierten Zeile `headline_row`)
können im E-Mail Text per `{...}` als Platzhalter verwendet werden.  
- Leere Tabellen-Felder werden durch einen leeren Platzhalter ersetzt.
- Zahlen (Telefonnummern) werden als String formatiert.
- Datumsangaben (Geburtsdaten) werden im Format YYYY-MM-DD als String konvertiert.

In Bezug auf die Tabelle in der Datei `Testdaten.xlsx`,
wird zum Beispiel aus dem Text *"Hallo \{Vorname\} \{Name\}"* der Text *"Hallo Max Mustermann"*

## ❗E-Mail Adressen
Zum Versenden der E-Mails werden alle Spalten genutzt die mit "E-Mail" beginnen. Die Tabelle muss **zwingend**
mindestens eine Spalte enthalten. Zum Beispiel "E-Mail" oder "E-Mail (privat)".  
Die Mails werden an alle Adressen gesendet, die dem Muster entsprechen.


## ▶️ Nutzung / CLI-Interface
Beispielaufruf: 
```
python Adressliste.py -f 'Meine Testdaten.xlsx' -c config.ini
```

### Kommandozeilen-Parameter
```
  -f,  --file [Pfad]
      Pfad zur Adressliste (xls/xlsx-Datei). Siehe Testdaten.xlsx. Default: 

  -c, --config [Pfad]
      Pfad zur Konfigurationsdatei. Siehe config_template.ini. Default: ./config.ini

  -l, --sendlist
      Wird dieser Parameter angegeben, wird die gesamte Liste an alle Teilnehmer versendet. 

  --trace [debug, info, warning]
      Bestimmt das Trace-level für Log-Ausgaben. Default: info
```


## 🕒 Automatischer Versand (Cron / Anacron)
Um die Adressliste z. B. **einmal jährlich automatisch** zu versenden,
kann ein Cronjob eingerichtet werden.

### Beispiel: Cronjob (jährlich am 1. Januar)
```bash
0 8 1 1 * /usr/bin/python3 /pfad/zu/Adressliste.py -f /pfad/Testdaten.xlsx -c /pfad/config.ini --sendlist
```
### Beispiel: Anacron (jährlich)
In `/etc/anacrontab`:
```text
365 10 adressliste-job /usr/bin/python3 /pfad/Adressliste.py -f /pfad/Testdaten.xlsx -c /pfad/config.ini --sendlist
```
Damit läuft das Skript zuverlässig auch auf Systemen, die nicht dauerhaft eingeschaltet sind.

### Windows Aufgabenplaner
Unter Windows kann ein komplett automatischer Versand mittels *Aufgabenplanung* realisiert werden. 

## 🛠️ Testmodus
Für einen Test-/Dry-Run-Modus, bei dem keine E-Mails tatsächlich versendet werden, kann das Modul
[aiosmtpd](https://pypi.org/project/aiosmtpd/) verwendet werden.
Es erstellt einen lokalen SMTP Server. Anstatt die E-Mails zu senden, gibt es den Inhalt der Mail auf der Konsole aus.  
`aiosmtpd` verwendet keine Verschlüsselung und hört auf localhost:8025.

### Installation
```shell
$ python -m pip install aiosmtpd
```
### Start von aiosmtpd in einem separaten Konsolenfenster
```shell
$ python -m aiosmtpd -n
```

Es muss folgende Konfiguration in der `config.ini` genutzt werden:
```python
[SMTP]
from_addr = test@example.com
smtp_host = localhost
smtp_port = 8025
use_tls=False
use_ssl=False
```

## 🧩 Ablauf
Vereinfacht arbeitet das Programm nach folgendem Schema:

             ┌───────────────────────┐
             │ Excel-Datei einlesen: │
             │   Adressliste.xlsx    │
             └──────────┬────────────┘
                        │
                        ▼
             ┌─────────────────────┐
             │ E-Mail-Spalten      │
             │ erkennen            │
             └──────────┬──────────┘
                        │
                        ▼
             ┌─────────────────────┐
             │ SMTP-Verbindung     │
             │ herstellen          │
             └──────────┬──────────┘
                        │
                        ▼
              ┌───────────────────┐
              │ Jede Tabellenzeile│
              │ verarbeiten       │
              └────────┬──────────┘
                       │
              ┌────────▼─────────┐
              │ Platzhalter      │
              │ ersetzen         │
              └────────┬─────────┘
                       │
                ┌──────▼──────┐
                │ --sendlist? │
                └──────┬─┬────┘
                    Ja │ │Nein
               ┌───────┘ └───────┐
               ▼                 ▼
        ┌──────────────┐   ┌────────────────┐
        │ Excel-Datei  │   │ Nur persönliche│
        │ anhängen     │   │ Daten senden   │
        └──────┬───────┘   └──────┬─────────┘
               │                  │
               └────────┬─────────┘
                        ▼
               ┌─────────────────┐
               │ E-Mail senden   │
               └────────┬────────┘
                        │
                        ▼
               ┌─────────────────┐
               │ Logging /       │
               │ Fehlerbehandlung│
               └─────────────────┘

## ⚖️ Lizenz
Dieses Projekt steht unter der MIT License. Siehe dazu die Lizenzdatei des Repositorys.