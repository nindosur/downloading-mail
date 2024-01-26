import imaplib
import email
from email.header import decode_header
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from bs4 import BeautifulSoup
import re
import html


def strip_html_tags(text):
    soup = BeautifulSoup(text, 'html.parser')
    plain_text = soup.get_text()
    return re.sub(r'[^A-Za-z0-9\s]+', '', plain_text)


def get_email_sender(msg):
    try:
        sender = msg.get("From")
        if sender:
            decoded_sender, encoding = email.header.decode_header(sender)[0]
            if isinstance(decoded_sender, bytes):
                return decoded_sender.decode(encoding or 'utf-8', errors='replace')
            else:
                return decoded_sender
    except Exception as e:
        print(f"Ошибка обработки отправителя электронной почты: {e}")
    return ""


def get_email_recipient(msg):
    try:
        recipient = msg.get("To")
        if recipient:
            decoded_recipient, encoding = email.header.decode_header(recipient)[0]
            if isinstance(decoded_recipient, bytes):
                return decoded_recipient.decode(encoding or 'utf-8', errors='replace')
            else:
                return decoded_recipient
    except Exception as e:
        print(f"Ошибка обработки получателя электронной почты: {e}")
    return ""


def get_email_subject(msg):
    try:
        subject = msg.get("Subject")
        if subject:
            decoded_subject, encoding = email.header.decode_header(subject)[0]
            if isinstance(decoded_subject, bytes):
                return decoded_subject.decode(encoding or 'utf-8', errors='replace')
            else:
                return decoded_subject
    except Exception as e:
        print(f"Ошибка обработки темы письма.: {e}")
    return ""


def get_email_content(msg):
    body = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        try:
                            body = payload.decode(errors='replace')
                        except (LookupError, UnicodeDecodeError) as e:
                            print(f"Ошибка декодирования: {e}")
                    break
                elif part.get_content_type() == "text/html":
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        try:
                            html_body = payload.decode(errors='replace')
                            body = strip_html_tags(html.unescape(html_body))
                        except (LookupError, UnicodeDecodeError) as e:
                            print(f"Ошибка декодирования данных HTML: {e}")
        else:
            payload = msg.get_payload(decode=True)
            if isinstance(payload, bytes):
                try:
                    body = payload.decode(errors='replace')
                except (LookupError, UnicodeDecodeError) as e:
                    print(f"Ошибка декодирования данных: {e}")
    except Exception as e:
        print(f"Ошибка обработки содержимого электронной почты: {e}")
    return body


def import_emails(email_address, password, excel_filename):
    # Подключение к почтовому ящику
    mail = imaplib.IMAP4_SSL("imap.yandex.ru")  # Измени на свой почтовый сервер
    mail.login(email_address, password)
    mail.select("inbox")

    # Создание Excel файла и листов
    wb = Workbook()
    incoming_sheet = wb.active
    incoming_sheet.title = "Входящие письма"
    outgoing_sheet = wb.create_sheet(title="Исходящие письма")

    # Заголовки столбцов
    headers = ["Номер письма", "Дата", "Время", "Отправитель/Адресат", "Тема", "Содержание", "Кол-во вложений", "Объем"]
    incoming_sheet.append(headers)
    outgoing_sheet.append(headers)

    # Получение всех сообщений
    result, data = mail.uid('search', None, 'ALL')
    if result == 'OK':
        uids = data[0].split()
        for uid in uids:
            result, msg_data = mail.uid('fetch', uid, '(RFC822)')
            if result == 'OK':
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                subject = get_email_content(msg)
                if isinstance(subject, bytes):
                    subject = subject.decode('utf-8') if isinstance(subject, bytes) else subject
                body = get_email_content(msg)
                attachments_count = len(msg.get_payload()) - 1 if msg.is_multipart() else 0
                date_time = email.utils.parsedate(msg["Date"])
                if date_time:
                    formatted_date = datetime(*date_time[:6]).strftime("%d.%m.%Y")
                    formatted_time = "{:02d}:{:02d}".format(date_time[3], date_time[4])
                else:
                    formatted_date = formatted_time = "N/A"
                volume_mb = len(raw_email) / (1024 * 1024)
                sender = get_email_sender(msg)
                if email_address in sender:
                    row = [uid.decode('utf-8'), formatted_date, formatted_time, sender, subject, body,
                           attachments_count, f"{volume_mb:.2f} MB"]
                    outgoing_sheet.append(row)
                else:
                    row = [uid.decode('utf-8'), formatted_date, formatted_time, sender, subject, body,
                           attachments_count, f"{volume_mb:.2f} MB"]
                    incoming_sheet.append(row)
                print(sender)

    wb.save(excel_filename)
    print(f"Импорт завершен. Данные сохранены в файл {excel_filename}")

    mail.logout()


if __name__ == "__main__":
    email_address = input("Введите адрес электронной почты: ")
    password = input("Введите пароль: ")
    excel_filename = input("Введите имя файла Excel для сохранения данных: ")

    import_emails(email_address, password, excel_filename)
