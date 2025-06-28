#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль логирования Excel Mail Processor
Централизованная система логирования с ротацией файлов
"""

import logging
import os
from logging.handlers import RotatingFileHandler
from datetime import datetime

def setup_logging():
    """Настройка системы логирования"""

    # Создаем директорию для логов
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Настройка форматирования
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Настройка основного логгера
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Очистка существующих обработчиков
    root_logger.handlers.clear()

    # Консольный обработчик
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # Файловый обработчик для всех логов
    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'excel_processor.log'),
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    # Файловый обработчик только для ошибок
    error_handler = RotatingFileHandler(
        os.path.join(log_dir, 'errors.log'),
        maxBytes=5*1024*1024,  # 5MB
        backupCount=3,
        encoding='utf-8'
    )
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)
    root_logger.addHandler(error_handler)

def get_logger(name):
    """Получение логгера для модуля"""

    # Настройка логирования при первом вызове
    if not logging.getLogger().handlers:
        setup_logging()

    return logging.getLogger(name)
