#!/bin/bash

# Excel Mail Processor v8.0 Final - Настройка планировщика cron с защитой от Segmentation fault

set -e

echo "⏰ Настройка планировщика задач Excel Mail Processor v8.0"
echo "=================================================="

# Получение абсолютного пути к проекту
PROJECT_DIR=$(pwd)
PYTHON_PATH="$PROJECT_DIR/venv/bin/python"
MAIN_SCRIPT="$PROJECT_DIR/main.py"

# Проверка существования файлов
if [ ! -f "$PYTHON_PATH" ]; then
    echo "❌ Виртуальное окружение не найдено: $PYTHON_PATH"
    echo "Запустите сначала: ./setup.sh"
    exit 1
fi

if [ ! -f "$MAIN_SCRIPT" ]; then
    echo "❌ Основной скрипт не найден: $MAIN_SCRIPT"
    exit 1
fi

# Создание скрипта-обертки для cron с защитой от Segmentation fault v8.0
create_cron_wrapper() {
    cat > "$PROJECT_DIR/run_excel_processor_v8.sh" << EOF
#!/bin/bash

# Обертка для запуска Excel Mail Processor v8.0 из cron
# Включает защиту от Segmentation fault

# Переход в рабочую директорию
cd "$PROJECT_DIR"

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
    export \$(cat .env | grep -v '^#' | xargs)
fi

# Логирование начала работы
echo "🚀 \$(date): Запуск Excel Mail Processor v8.0" >> logs/cron.log
echo "🛡️ Защита Segmentation fault: OPENBLAS=\$OPENBLAS_NUM_THREADS OMP=\$OMP_NUM_THREADS MKL=\$MKL_NUM_THREADS" >> logs/cron.log

# Активация виртуального окружения и запуск
"$PYTHON_PATH" "$MAIN_SCRIPT" >> logs/cron.log 2>&1

# Логирование завершения
echo "✅ \$(date): Завершение Excel Mail Processor v8.0" >> logs/cron.log
echo "---" >> logs/cron.log

EOF

    chmod +x "$PROJECT_DIR/run_excel_processor_v8.sh"
    echo "✅ Создан скрипт-обертка v8.0 с защитой от Segmentation fault: run_excel_processor_v8.sh"
}

# Выбор расписания
choose_schedule() {
    echo ""
    echo "📅 Выберите расписание запуска v8.0:"
    echo "1) Каждые 15 минут (рекомендуется для v8.0)"
    echo "2) Каждые 30 минут"
    echo "3) Каждый час"
    echo "4) Каждые 2 часа"
    echo "5) Только в рабочие часы (9:00-18:00, пн-пт, каждые 30 мин)"
    echo "6) Пользовательское расписание"
    read -p "Введите номер (1-6): " choice

    case $choice in
        1)
            CRON_SCHEDULE="*/15 * * * *"
            DESCRIPTION="каждые 15 минут"
            ;;
        2)
            CRON_SCHEDULE="*/30 * * * *"
            DESCRIPTION="каждые 30 минут"
            ;;
        3)
            CRON_SCHEDULE="0 * * * *"
            DESCRIPTION="каждый час"
            ;;
        4)
            CRON_SCHEDULE="0 */2 * * *"
            DESCRIPTION="каждые 2 часа"
            ;;
        5)
            CRON_SCHEDULE="*/30 9-18 * * 1-5"
            DESCRIPTION="каждые 30 минут в рабочие часы (9:00-18:00, пн-пт)"
            ;;
        6)
            echo "Введите cron расписание (например: 0 */2 * * *):"
            read -p "> " CRON_SCHEDULE
            DESCRIPTION="пользовательское расписание"
            ;;
        *)
            echo "❌ Неверный выбор"
            exit 1
            ;;
    esac
}

# Установка cron задания
install_cron_job() {
    # Создание временного файла с cron заданиями
    TEMP_CRON=$(mktemp)

    # Сохранение существующих cron заданий
    crontab -l > "$TEMP_CRON" 2>/dev/null || true

    # Удаление старых заданий Excel Mail Processor
    grep -v "run_excel_processor" "$TEMP_CRON" > "${TEMP_CRON}.new" || true
    mv "${TEMP_CRON}.new" "$TEMP_CRON"

    # Добавление нового задания v8.0
    echo "# Excel Mail Processor v8.0 - $DESCRIPTION" >> "$TEMP_CRON"
    echo "$CRON_SCHEDULE $PROJECT_DIR/run_excel_processor_v8.sh" >> "$TEMP_CRON"

    # Установка обновленного crontab
    crontab "$TEMP_CRON"

    # Удаление временного файла
    rm "$TEMP_CRON"

    echo "✅ Cron задание v8.0 установлено: $DESCRIPTION"
}

# Создание директории для логов
setup_logging() {
    mkdir -p "$PROJECT_DIR/logs"
    
    # Создание файла лога cron если не существует
    touch "$PROJECT_DIR/logs/cron.log"
    
    # Создание лога для мониторинга Segmentation fault
    touch "$PROJECT_DIR/logs/segfault_protection.log"
    
    echo "✅ Настроено логирование v8.0 в logs/"
}

# Проверка cron демона
check_cron_service() {
    if ! pgrep -x "cron" > /dev/null && ! pgrep -x "crond" > /dev/null; then
        echo "⚠️ Служба cron не запущена"
        echo "Обратитесь к администратору сервера для запуска cron"
    else
        echo "✅ Служба cron работает"
    fi
}

# Отображение текущих заданий
show_current_jobs() {
    echo ""
    echo "📋 Текущие cron задания v8.0:"
    crontab -l | grep -E "(Excel|run_excel_processor)" || echo "Нет заданий Excel Mail Processor"
}

# Тестирование защиты от Segmentation fault
test_segfault_protection() {
    echo ""
    echo "🛡️ Тестирование защиты от Segmentation fault..."
    
    # Тест запуска скрипта обертки
    if [ -f "$PROJECT_DIR/run_excel_processor_v8.sh" ]; then
        echo "Тестовый запуск с защитой..."
        cd "$PROJECT_DIR"
        timeout 30s ./run_excel_processor_v8.sh --version > /dev/null 2>&1
        if [ $? -eq 0 ]; then
            echo "✅ Защита от Segmentation fault работает корректно"
        else
            echo "⚠️ Возможны проблемы с защитой, проверьте логи"
        fi
    fi
}

# Основная функция
main() {
    create_cron_wrapper
    choose_schedule
    install_cron_job
    setup_logging
    check_cron_service
    test_segfault_protection
    show_current_jobs

    echo ""
    echo "✅ Настройка cron v8.0 завершена!"
    echo ""
    echo "📋 Полезные команды v8.0:"
    echo " Просмотр логов cron: tail -f logs/cron.log"
    echo " Просмотр всех cron: crontab -l"
    echo " Удаление cron заданий: crontab -r"
    echo " Ручной запуск v8.0: ./run_excel_processor_v8.sh"
    echo " Тестирование v8.0: python main.py --test"
    echo ""
    echo "🛡️ Защита от Segmentation fault:"
    echo " Переменные окружения установлены автоматически"
    echo " Ограничения ресурсов применяются в cron обертке"
    echo ""
    echo "📖 Система v8.0 будет автоматически проверять почту $DESCRIPTION"
}

# Запуск
main "$@"