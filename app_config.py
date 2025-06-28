#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль конфигурации Excel Mail Processor
Загрузка и валидация переменных окружения
"""

import os
from dotenv import load_dotenv
from logger import get_logger

logger = get_logger(__name__)

class Config:
    """Класс конфигурации системы"""

    def __init__(self):
        """Инициализация конфигурации"""
        load_dotenv()
        self._load_config()
        self._validate_config()
        logger.info("Конфигурация загружена и валидирована успешно")

    def _load_config(self):
        """Загрузка всех параметров конфигурации"""

        # IMAP настройки
        self.imap_server = os.getenv('IMAP_SERVER', 'mail.hosting.reg.ru')
        self.imap_port = int(os.getenv('IMAP_PORT', '993'))
        self.imap_user = os.getenv('IMAP_USER')
        self.imap_password = os.getenv('IMAP_PASSWORD')
        self.imap_use_ssl = os.getenv('IMAP_USE_SSL', 'true').lower() == 'true'

        # SMTP настройки
        self.smtp_server = os.getenv('SMTP_SERVER', 'mail.hosting.reg.ru')
        self.smtp_port = int(os.getenv('SMTP_PORT', '465'))
        self.smtp_user = os.getenv('SMTP_USER')
        self.smtp_password = os.getenv('SMTP_PASSWORD')
        self.smtp_use_ssl = os.getenv('SMTP_USE_SSL', 'true').lower() == 'true'
        self.smtp_use_tls = os.getenv('SMTP_USE_TLS', 'false').lower() == 'true'

        # OneDrive настройки
        self.onedrive_instruction_url = os.getenv('ONEDRIVE_INSTRUCTION_URL')

        # Email настройки
        self.recipient_email = os.getenv('RECIPIENT_EMAIL')
        self.sender_name = os.getenv('SENDER_NAME', 'Transfora Mail Processor')

        # Дополнительные настройки
        self.temp_dir = os.getenv('TEMP_DIR', '/tmp')
        self.log_level = os.getenv('LOG_LEVEL', 'INFO')
        self.max_file_size_mb = int(os.getenv('MAX_FILE_SIZE_MB', '10'))

    def _validate_config(self):
        """Валидация обязательных параметров"""
        required_params = [
            ('IMAP_USER', self.imap_user),
            ('IMAP_PASSWORD', self.imap_password),
            ('SMTP_USER', self.smtp_user),
            ('SMTP_PASSWORD', self.smtp_password),
            ('ONEDRIVE_INSTRUCTION_URL', self.onedrive_instruction_url),
            ('RECIPIENT_EMAIL', self.recipient_email)
        ]

        missing_params = []
        for param_name, param_value in required_params:
            if not param_value:
                missing_params.append(param_name)

        if missing_params:
            error_msg = f"Отсутствуют обязательные параметры в .env: {', '.join(missing_params)}"
            logger.error(error_msg)
            raise ValueError(error_msg)

    def get_imap_config(self):
        """Получение конфигурации IMAP"""
        return {
            'host': self.imap_server,
            'port': self.imap_port,
            'username': self.imap_user,
            'password': self.imap_password,
            'ssl': self.imap_use_ssl
        }

    def get_smtp_config(self):
        """Получение конфигурации SMTP"""
        return {
            'host': self.smtp_server,
            'port': self.smtp_port,
            'username': self.smtp_user,
            'password': self.smtp_password,
            'ssl': self.smtp_use_ssl,
            'tls': self.smtp_use_tls
        }
