#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Mail Processor v8.0 Final

Автоматизация обработки Excel файлов из электронной почты
Специально адаптирован для хостинга REG.RU

ИСПРАВЛЕНИЯ v8.0:
- Исправлено именование результирующих файлов 
- Добавлена защита от Segmentation fault
- Поддержка настраиваемых email шаблонов
- Цветовое форматирование из OneDrive
"""

import sys
import os
import argparse
from datetime import datetime

# ЗАЩИТА ОТ SEGMENTATION FAULT v8.0
# Установка переменных окружения для ограничения потоков
# Решение проверенное пользователем
os.environ['OPENBLAS_NUM_THREADS'] = '1'
os.environ['OMP_NUM_THREADS'] = '1'
os.environ['MKL_NUM_THREADS'] = '1'

# Дополнительная защита от ошибок памяти
import resource
try:
    # Ограничение количества процессов (эквивалент ulimit -u 36)
    resource.setrlimit(resource.RLIMIT_NPROC, (36, 36))
except:
    # Если не удается установить лимит, просто логируем
    pass

# Локальные импорты
from app_config import Config
from excel_processor_v8 import ExcelProcessor  # Обновленный модуль v8.0
from email_handler_v8 import EmailHandler  # Обновленный модуль v8.0
from onedrive_handler_v8 import OneDriveHandler  # Обновленный модуль v8.0
from logger import get_logger

# Настройка логирования
logger = get_logger(__name__)

def test_system():
    """Тестирование всех компонентов системы v8.0"""
    print("🔧 Тестирование системы Excel Mail Processor v8.0...")
    try:
        # Загрузка конфигурации
        config = Config()
        logger.info("✅ Конфигурация загружена успешно")

        # Тестирование почтовых соединений
        email_handler = EmailHandler(config)
        print("📧 Тестирование IMAP соединения...")
        if email_handler.test_imap_connection():
            logger.info("✅ IMAP соединение работает")
        else:
            logger.error("❌ IMAP соединение не работает")

        print("📤 Тестирование SMTP соединения...")
        if email_handler.test_smtp_connection():
            logger.info("✅ SMTP соединение работает")
        else:
            logger.error("❌ SMTP соединение не работает")

        # Тестирование OneDrive v8.0
        print("☁️ Тестирование загрузки инструкции с OneDrive v8.0...")
        onedrive_handler = OneDriveHandler(config.onedrive_instruction_url)
        instructions = onedrive_handler.get_processing_instructions()
        
        if instructions:
            logger.info("✅ Инструкция OneDrive загружена успешно v8.0")
            print(f" - Колонок для обработки: {len(instructions['columns'])}")
            print(f" - Правил замены: {len(instructions['replace_rules'])}")
            print(f" - Email настроек: {len(instructions['email_template'])}")
            print(f" - Настроек форматирования: {len(instructions['formatting'])}")
            
            # Тестирование новых возможностей v8.0
            if 'body_template' in instructions['email_template']:
                print(" - ✅ Найден настраиваемый email шаблон v8.0")
            if instructions['formatting']:
                print(" - ✅ Найдены настройки цветового форматирования v8.0")
        else:
            logger.error("❌ Не удалось загрузить инструкцию с OneDrive")

        # Тестирование защиты от Segmentation fault
        print("🛡️ Тестирование защиты от Segmentation fault...")
        test_segfault_protection()

    except Exception as e:
        logger.error(f"❌ Критическая ошибка: {str(e)}")

    print("\n🔧 Тестирование v8.0 завершено!")

def test_segfault_protection():
    """Тестирование защиты от Segmentation fault"""
    try:
        # Проверка установки переменных окружения
        required_vars = ['OPENBLAS_NUM_THREADS', 'OMP_NUM_THREADS', 'MKL_NUM_THREADS']
        for var in required_vars:
            value = os.environ.get(var)
            if value == '1':
                print(f" - ✅ {var} = {value}")
            else:
                print(f" - ⚠️ {var} = {value} (ожидается '1')")
        
        # Проверка ограничения процессов
        try:
            soft, hard = resource.getrlimit(resource.RLIMIT_NPROC)
            print(f" - ✅ Лимит процессов: {soft}/{hard}")
        except:
            print(" - ⚠️ Не удалось проверить лимит процессов")
            
        # Простой тест импорта pandas с защитой
        import pandas as pd
        test_df = pd.DataFrame({'test': [1, 2, 3]})
        print(f" - ✅ Pandas работает корректно (тест: {len(test_df)} строк)")
        
    except Exception as e:
        logger.warning(f"⚠️ Ошибка тестирования защиты: {str(e)}")

def process_emails():
    """Основная обработка писем v8.0"""
    logger.info("🚀 Запуск обработки писем v8.0")
    try:
        # Загрузка конфигурации
        config = Config()
        logger.info("📋 Конфигурация загружена")

        # Загрузка инструкции с OneDrive v8.0
        logger.info("☁️ Загрузка инструкции с OneDrive v8.0...")
        onedrive_handler = OneDriveHandler(config.onedrive_instruction_url)
        instructions = onedrive_handler.get_processing_instructions()

        if not instructions:
            logger.error("❌ Не удалось загрузить инструкцию с OneDrive")
            return

        # Инициализация обработчиков v8.0
        email_handler = EmailHandler(config)
        excel_processor = ExcelProcessor(instructions)

        # Проверка новых писем
        logger.info("📧 Проверка новых писем...")
        emails_with_excel = email_handler.get_unread_emails_with_excel()

        if not emails_with_excel:
            logger.info("📪 Новых писем с Excel файлами не найдено")
            return

        logger.info(f"📬 Найдено {len(emails_with_excel)} писем для обработки")

        # Обработка каждого письма
        processed_files = []
        for i, email_data in enumerate(emails_with_excel, 1):
            logger.info(f"⚙️ Обработка письма {i}/{len(emails_with_excel)}")

            # Обработка каждого Excel файла в письме
            for attachment in email_data['attachments']:
                try:
                    # Обработка файла v8.0
                    output_path = excel_processor.process_file(
                        attachment['content'],
                        attachment['filename']
                    )

                    # Получение статистики обработки v8.0
                    processing_stats = excel_processor.get_processing_statistics()

                    # Отправка результата v8.0 с новыми возможностями
                    email_handler.send_processed_file_v8(
                        output_path,
                        attachment['filename'],
                        email_data['sender'],
                        instructions.get('email_template', {}),
                        processing_stats
                    )

                    processed_files.append(attachment['filename'])
                    logger.info(f"✅ Файл {attachment['filename']} обработан и отправлен v8.0")

                except Exception as e:
                    logger.error(f"❌ Ошибка обработки файла {attachment['filename']}: {str(e)}")

        # Пометка писем как прочитанных
        if processed_files:
            email_handler.mark_emails_as_read(emails_with_excel)
            logger.info(f"✅ Обработка v8.0 завершена успешно: {len(processed_files)} файлов")

    except Exception as e:
        logger.error(f"❌ Критическая ошибка v8.0: {str(e)}")

def main():
    """Главная функция v8.0"""
    parser = argparse.ArgumentParser(description='Excel Mail Processor v8.0')
    parser.add_argument('--test', action='store_true', help='Тестировать систему v8.0')
    parser.add_argument('--config', action='store_true', help='Показать конфигурацию')
    parser.add_argument('--version', action='store_true', help='Показать версию')

    args = parser.parse_args()

    if args.version:
        print("📧 Excel Mail Processor v8.0 Final")
        print("🚀 Автоматизация обработки дислокации вагонов")
        print("🏢 ООО ТРАНСФОРА")
        print("📅 2025")
        print("\n🆕 Новое в v8.0:")
        print(" - ✅ Исправлено именование результирующих файлов")
        print(" - 🛡️ Защита от Segmentation fault")
        print(" - 📧 Настраиваемые email шаблоны")
        print(" - 🎨 Цветовое форматирование из OneDrive")

    elif args.test:
        test_system()

    elif args.config:
        try:
            config = Config()
            print("📋 Текущая конфигурация v8.0:")
            print(f" - IMAP Server: {config.imap_server}:{config.imap_port}")
            print(f" - SMTP Server: {config.smtp_server}:{config.smtp_port}")
            print(f" - OneDrive URL: {config.onedrive_instruction_url[:50]}...")
            print(f" - Recipient: {config.recipient_email}")
            print("\n🛡️ Защита от Segmentation fault:")
            for var in ['OPENBLAS_NUM_THREADS', 'OMP_NUM_THREADS', 'MKL_NUM_THREADS']:
                print(f" - {var}: {os.environ.get(var, 'НЕ УСТАНОВЛЕНО')}")
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки конфигурации: {str(e)}")

    else:
        process_emails()

if __name__ == "__main__":
    main()