#!/bin/bash

# Обертка для запуска Excel Mail Processor v8.0 из cron
# Включает защиту от Segmentation fault

# Переход в рабочую директорию
cd "/var/www/u2995595/data/www/sonpic.com/fgk"

# ЗАЩИТА ОТ SEGMENTATION FAULT v8.0
# Установка переменных окружения для ограничения потоков
export OPENBLAS_NUM_THREADS=1
export OMP_NUM_THREADS=1
export MKL_NUM_THREADS=1

# Ограничение количества процессов (ulimit -u 36)
ulimit -u 36

# Дополнительные ограничения памяти для стабильности
ulimit -m 512000  # 512MB память
ulimit -v 1048576 # 1GB виртуальная память

# Загрузка переменных окружения
if [ -f ".env" ]; then
    export $(cat .env | grep -v '^#' | xargs)
fi

# Логирование начала работы
echo "🚀 $(date): Запуск Excel Mail Processor v8.0" >> logs/cron.log
echo "🛡️ Защита Segmentation fault: OPENBLAS=$OPENBLAS_NUM_THREADS OMP=$OMP_NUM_THREADS MKL=$MKL_NUM_THREADS" >> logs/cron.log

# Активация виртуального окружения и запуск
"/var/www/u2995595/data/www/sonpic.com/fgk/venv/bin/python" "/var/www/u2995595/data/www/sonpic.com/fgk/main.py" >> logs/cron.log 2>&1

# Логирование завершения
echo "✅ $(date): Завершение Excel Mail Processor v8.0" >> logs/cron.log
echo "---" >> logs/cron.log

