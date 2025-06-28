#!/bin/bash
# Excel Mail Processor v6.0 Final - Автоматическая установка для REG.RU

set -e  # Остановка при любой ошибке

echo "🚀 Excel Mail Processor v6.0 Final - Установка для REG.RU"
echo "=================================================="

# Функция проверки версии Python
check_python_version() {
    echo "🔍 Проверяем версию Python..."

    # Попробуем найти подходящий Python
    for python_cmd in python3 python; do
        if command -v $python_cmd &> /dev/null; then
            # Получаем версию через сам Python
            version_check=$($python_cmd -c "
import sys
major, minor = sys.version_info.major, sys.version_info.minor
if major >= 3 and minor >= 7:
    print('OK')
    print(f'{major}.{minor}')
else:
    print('LOW')
    print(f'{major}.{minor}')
" 2>/dev/null)

            if echo "$version_check" | head -n1 | grep -q "OK"; then
                python_version=$(echo "$version_check" | tail -n1)
                echo "Найдена версия Python: $python_version"
                echo "✅ Версия Python совместима"
                PYTHON_CMD=$python_cmd
                return 0
            fi
        fi
    done

    echo "❌ Не найдена совместимая версия Python (требуется 3.7+)"
    echo "Установите Python 3.7 или выше"
    exit 1
}

# Создание виртуального окружения
create_venv() {
    echo "📦 Создание виртуального окружения..."

    if [ -d "venv" ]; then
        echo "⚠️  Виртуальное окружение уже существует"
        read -p "Пересоздать? (y/N): " recreate
        if [[ $recreate =~ ^[Yy]$ ]]; then
            rm -rf venv
        else
            echo "ℹ️  Используется существующее окружение"
            return 0
        fi
    fi

    $PYTHON_CMD -m venv venv
    echo "✅ Виртуальное окружение создано"
}

# Активация виртуального окружения
activate_venv() {
    echo "🔧 Активация виртуального окружения..."
    source venv/bin/activate
    echo "✅ Виртуальное окружение активировано"
}

# Установка зависимостей
install_dependencies() {
    echo "📚 Установка зависимостей..."

    # Обновление pip
    pip install --upgrade pip

    # Установка пакетов
    pip install -r requirements.txt

    echo "✅ Зависимости установлены"
}

# Настройка конфигурации
setup_config() {
    echo "⚙️  Настройка конфигурации..."

    if [ ! -f ".env" ]; then
        cp .env.example .env
        chmod 600 .env
        echo "✅ Создан файл конфигурации .env"
        echo "⚠️  ВАЖНО: Отредактируйте .env файл с вашими данными!"
    else
        echo "ℹ️  Файл .env уже существует"
    fi
}

# Создание директорий
create_directories() {
    echo "📁 Создание рабочих директорий..."

    mkdir -p logs
    chmod 755 logs

    echo "✅ Директории созданы"
}

# Тестирование установки
test_installation() {
    echo "🧪 Тестирование установки..."

    if python main.py --config > /dev/null 2>&1; then
        echo "✅ Основные модули загружаются корректно"
    else
        echo "⚠️  Проблемы с загрузкой модулей (проверьте .env файл)"
    fi
}

# Основная функция установки
main() {
    echo "📍 Рабочая директория: $(pwd)"

    check_python_version
    create_venv
    activate_venv
    install_dependencies
    setup_config
    create_directories
    test_installation

    echo ""
    echo "✅ Установка завершена!"
    echo ""
    echo "📋 Следующие шаги:"
    echo "1. Отредактируйте .env файл: nano .env"
    echo "2. Загрузите FGK.xlsx на OneDrive и обновите ссылку в .env"
    echo "3. Протестируйте систему: source venv/bin/activate && python main.py --test"
    echo "4. Запустите обработку: python main.py"
    echo "5. Настройте автозапуск: ./cron_setup.sh"
    echo ""
    echo "📖 Полная документация в файле README.md"
}

# Запуск основной функции
main "$@"
