import os
import sys

# Базовая директория проекта (с поддержкой PyInstaller .exe)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Пути к папкам
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
DATA_DIR = os.path.join(BASE_DIR, 'data')
RESULT_DIR = os.path.join(BASE_DIR, 'result')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')

# Файлы
LOG_FILE = os.path.join(LOGS_DIR, 'generator.log')

# Настройки генерации
CREATE_SUBFOLDERS_BY_FIO = True   # Создавать подпапки для каждого человека
CLEAN_TEMP_FILES = True           # Удалять временные .docx после конвертации в PDF

# Информация о приложении и автообновлении
APP_NAME = "Yahoo"
APP_VERSION = "1.0.1"
GITHUB_REPO = "thejuzvr/app-excel"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_CHECK_ON_STARTUP = True