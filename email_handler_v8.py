#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль работы с электронной почтой v8.0
Получение писем и отправка результатов
ИСПРАВЛЕНИЯ v8.0:
- Поддержка настраиваемых email шаблонов из OneDrive
- Расширенные переменные для подстановки
- Улучшенное логирование отправки
"""

import smtplib
import ssl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
from imap_tools import MailBox, AND
from logger import get_logger

logger = get_logger(__name__)

class EmailHandler:
    """Обработчик электронной почты v8.0"""

    def __init__(self, config):
        """Инициализация обработчика"""
        self.config = config

    def test_imap_connection(self):
        """Тестирование IMAP соединения"""
        try:
            with MailBox(self.config.imap_server, port=self.config.imap_port).login(
                self.config.imap_user,
                self.config.imap_password,
                'INBOX'
            ) as mailbox:
                logger.info("IMAP соединение работает")
                return True
        except Exception as e:
            logger.error(f"Ошибка IMAP соединения: {str(e)}")
            return False

    def test_smtp_connection(self):
        """Тестирование SMTP соединения"""
        try:
            if self.config.smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(self.config.smtp_server, self.config.smtp_port, context=context) as server:
                    server.login(self.config.smtp_user, self.config.smtp_password)
                    logger.info("SMTP соединение работает")
                    return True
            else:
                with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
                    server.starttls()
                    server.login(self.config.smtp_user, self.config.smtp_password)
                    logger.info("SMTP соединение работает")
                    return True
        except Exception as e:
            logger.error(f"Ошибка SMTP соединения: {str(e)}")
            return False

    def get_unread_emails_with_excel(self):
        """Получение непрочитанных писем с Excel вложениями"""
        try:
            emails_with_excel = []
            with MailBox(self.config.imap_server, port=self.config.imap_port).login(
                self.config.imap_user,
                self.config.imap_password,
                'INBOX'
            ) as mailbox:
                # Поиск непрочитанных писем
                for msg in mailbox.fetch(AND(seen=False)):
                    excel_attachments = []
                    # Проверка вложений
                    for attachment in msg.attachments:
                        if attachment.filename and attachment.filename.lower().endswith(('.xlsx', '.xls')):
                            file_size_mb = len(attachment.payload) / (1024 * 1024)
                            if file_size_mb <= self.config.max_file_size_mb:
                                excel_attachments.append({
                                    'filename': attachment.filename,
                                    'content': attachment.payload,
                                    'size_mb': round(file_size_mb, 2)
                                })
                                logger.info(f"Найдено Excel вложение: {attachment.filename} ({file_size_mb:.2f} МБ)")
                            else:
                                logger.warning(f"Файл {attachment.filename} превышает максимальный размер ({file_size_mb:.2f} МБ)")

                    if excel_attachments:
                        emails_with_excel.append({
                            'uid': msg.uid,
                            'sender': msg.from_,
                            'subject': msg.subject,
                            'date': msg.date,
                            'attachments': excel_attachments
                        })
                        logger.info(f"Письмо от {msg.from_} содержит {len(excel_attachments)} Excel файлов")

            return emails_with_excel
        except Exception as e:
            logger.error(f"Ошибка при получении писем: {str(e)}")
            return []

    def send_processed_file_v8(self, file_path, original_filename, sender_email, email_template=None, processing_stats=None):
        """
        Отправка обработанного файла v8.0 с поддержкой настраиваемых шаблонов
        """
        try:
            # Подготовка данных для шаблона v8.0
            now = datetime.now()
            
            # Расширенный набор переменных для v8.0
            template_data = {
                'source_filename': original_filename,
                'output_filename': os.path.basename(file_path),
                'sender_email': sender_email,
                'processing_date': now.strftime("%Y-%m-%d %H:%M:%S"),
                'processing_date_short': now.strftime("%Y-%m-%d"),
                'processing_time': now.strftime("%H:%M"),
                'processed_rows': processing_stats.get('processed_rows', 0) if processing_stats else 0,
                'applied_rules': processing_stats.get('applied_rules', 0) if processing_stats else 0,
                'created_columns': processing_stats.get('created_columns', 0) if processing_stats else 0,
                'custom_content': self._get_custom_content(processing_stats),
                'footer_text': self._get_footer_text(email_template)
            }

            # Формирование темы и текста письма v8.0
            if email_template and 'body_template' in email_template:
                # Новый подход v8.0: полный шаблон из OneDrive
                subject = email_template.get('subject', 'Обработанный файл: {output_filename}').format(**template_data)
                body = email_template['body_template'].format(**template_data)
                logger.info("✅ Используется настраиваемый шаблон письма v8.0 из OneDrive")
            elif email_template:
                # Совместимость с v6.0
                subject = email_template.get('subject', 'Обработанный файл: {output_filename}').format(**template_data)
                body = self._build_legacy_email_body(email_template, template_data)
                logger.info("📧 Используется совместимый шаблон письма v6.0")
            else:
                # Стандартный шаблон
                subject = f"Обработанный файл: {os.path.basename(file_path)}"
                body = self._get_default_email_body_v8().format(**template_data)
                logger.info("📧 Используется стандартный шаблон письма")

            # Создание сообщения
            msg = MIMEMultipart()
            msg['From'] = f"{self.config.sender_name} <{self.config.smtp_user}>"
            msg['To'] = self.config.recipient_email
            msg['Subject'] = subject

            # Добавление текста
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Добавление вложения
            with open(file_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='xlsx')
                attachment.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=os.path.basename(file_path)
                )
                msg.attach(attachment)

            # Отправка письма
            if self.config.smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(self.config.smtp_server, self.config.smtp_port, context=context) as server:
                    server.login(self.config.smtp_user, self.config.smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
                    server.starttls()
                    server.login(self.config.smtp_user, self.config.smtp_password)
                    server.send_message(msg)

            logger.info(f"✅ Email отправлен успешно на {self.config.recipient_email}")
            logger.info(f"📧 Тема: {subject}")
            logger.info(f"📎 Вложение: {os.path.basename(file_path)}")

            # Удаление временного файла
            if os.path.exists(file_path):
                os.unlink(file_path)
                logger.info(f"🗑️ Временный файл удален: {file_path}")

        except Exception as e:
            logger.error(f"Ошибка при отправке email: {str(e)}")
            raise

    def _get_custom_content(self, processing_stats):
        """Генерация кастомного контента для письма"""
        if not processing_stats:
            return ""
            
        content_parts = []
        
        if processing_stats.get('applied_rules', 0) > 0:
            content_parts.append(f"✅ Успешно применены правила замены для {processing_stats['applied_rules']} типов записей")
            
        if processing_stats.get('created_columns', 0) > 0:
            content_parts.append(f"📊 Добавлено новых колонок: {processing_stats['created_columns']}")
            
        return "\n".join(content_parts)

    def _get_footer_text(self, email_template):
        """Получение текста футера"""
        if email_template and 'footer_text' in email_template:
            return email_template['footer_text']
        return "Автоматическая обработка Excel Mail Processor v8.0\nООО ТРАНСФОРА"

    def _build_legacy_email_body(self, email_template, template_data):
        """Построение письма в стиле v6.0 для совместимости"""
        header = email_template.get('body_header', 'Результат автоматической обработки файла дислокации вагонов.')
        footer = email_template.get('body_footer', 'ООО ТРАНСФОРА')
        
        body = f"""{header}

Исходный файл: {template_data['source_filename']}
Обработанный файл: {template_data['output_filename']}
Отправитель: {template_data['sender_email']}
Дата и время обработки: {template_data['processing_date']}

Статистика обработки:
- Обработано строк: {template_data['processed_rows']}
- Применено правил замены: {template_data['applied_rules']}
- Создано колонок: {template_data['created_columns']}


---
{footer}"""
        return body

    def _get_default_email_body_v8(self):
        """Получение стандартного шаблона письма v8.0"""
        template = """Результат автоматической обработки файла дислокации вагонов.


Приложение к письму: {output_filename}

---
{footer_text}"""
        return template

    def mark_emails_as_read(self, emails_data):
        """Пометка писем как прочитанных"""
        try:
            with MailBox(self.config.imap_server, port=self.config.imap_port).login(
                self.config.imap_user,
                self.config.imap_password,
                'INBOX'
            ) as mailbox:
                uids = [email_data['uid'] for email_data in emails_data]
                mailbox.flag(uids, ['\\Seen'], True)
                logger.info(f"Помечено как прочитанных: {len(uids)} писем")
        except Exception as e:
            logger.error(f"Ошибка при отметке писем как прочитанных: {str(e)}")

    # Обратная совместимость с v6.0
    def send_processed_file(self, file_path, original_filename, sender_email, email_template=None):
        """Обратная совместимость с v6.0"""
        return self.send_processed_file_v8(file_path, original_filename, sender_email, email_template)