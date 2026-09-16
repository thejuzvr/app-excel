import os
import sys
import re
import json
import ssl
import time
import zipfile
import tempfile
import logging
import subprocess
import urllib.request
import urllib.error
from typing import Dict, Any, Optional, Tuple, Callable

from config import APP_VERSION, GITHUB_REPO, BASE_DIR, LOG_FILE

logger = logging.getLogger(__name__)


def parse_version(version_str: str) -> Tuple[int, ...]:
    """
    Преобразует строку версии (например, 'v1.2.3' или '1.0') в кортеж чисел для сравнения.
    """
    cleaned = re.sub(r'^[^\d]*', '', version_str.strip())
    parts = []
    for part in re.split(r'[\.\-_]', cleaned):
        match = re.match(r'(\d+)', part)
        if match:
            parts.append(int(match.group(1)))
        else:
            break
    return tuple(parts) if parts else (0,)


def check_for_updates(
    current_version: str = APP_VERSION,
    repo: str = GITHUB_REPO,
    timeout: int = 8
) -> Dict[str, Any]:
    """
    Проверяет наличие новой версии на GitHub Releases.
    Возвращает словарь с информацией о статусе обновления.
    """
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    headers = {
        "User-Agent": "Yahoo-DocGen-App",
        "Accept": "application/vnd.github.v3+json"
    }

    ctx = ssl.create_default_context()

    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as response:
            data = json.loads(response.read().decode('utf-8'))
            
            tag_name = data.get("tag_name", "")
            release_name = data.get("name") or tag_name
            changelog = data.get("body", "")
            html_url = data.get("html_url", "")
            published_at = data.get("published_at", "")
            
            remote_ver = parse_version(tag_name)
            curr_ver = parse_version(current_version)
            
            has_update = remote_ver > curr_ver
            
            # Поиск подходящего архива или установщика среди assets
            assets = data.get("assets", [])
            download_url = None
            asset_name = None
            asset_size = 0
            
            # Приоритет: .zip архив с программой
            for asset in assets:
                name = asset.get("name", "")
                if name.lower().endswith(".zip"):
                    download_url = asset.get("browser_download_url")
                    asset_name = name
                    asset_size = asset.get("size", 0)
                    break
            
            # Вторичный приоритет: .exe файл (если собран единый установщик/бинарник)
            if not download_url:
                for asset in assets:
                    name = asset.get("name", "")
                    if name.lower().endswith(".exe"):
                        download_url = asset.get("browser_download_url")
                        asset_name = name
                        asset_size = asset.get("size", 0)
                        break
            
            # Fallback: архив с исходным кодом релиза
            if not download_url:
                download_url = data.get("zipball_url")
                asset_name = f"{repo.replace('/', '-')}-{tag_name}.zip"
            
            return {
                "success": True,
                "has_update": has_update,
                "current_version": current_version,
                "latest_version": tag_name.lstrip('v'),
                "tag_name": tag_name,
                "release_name": release_name,
                "changelog": changelog,
                "html_url": html_url,
                "published_at": published_at,
                "download_url": download_url,
                "asset_name": asset_name,
                "asset_size": asset_size
            }
            
    except urllib.error.HTTPError as e:
        if e.code == 404:
            # Релизов в репозитории пока нет — у пользователя актуальная версия
            return {
                "success": True,
                "has_update": False,
                "current_version": current_version,
                "latest_version": current_version,
                "message": f"У вас установлена последняя версия (v{current_version})."
            }
        elif e.code == 403:
            return {
                "success": False,
                "has_update": False,
                "error": "Слишком много запросов к серверу обновлений. Попробуйте позже."
            }
        else:
            return {
                "success": False,
                "has_update": False,
                "error": f"Ошибка сервера обновлений (код {e.code})"
            }
            
    except urllib.error.URLError as e:
        return {
            "success": False,
            "has_update": False,
            "error": "Не удалось проверить обновления. Проверьте подключение к интернету."
        }
    except Exception as e:
        logger.exception("Ошибка при проверке обновлений")
        return {
            "success": False,
            "has_update": False,
            "error": f"Не удалось проверить обновления: {e}"
        }


def download_update(
    download_url: str,
    target_path: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None
) -> bool:
    """
    Скачивает файл обновления по ссылке с отслеживанием прогресса.
    progress_callback(downloaded_bytes, total_bytes, speed_str)
    """
    ctx = ssl.create_default_context()
    headers = {"User-Agent": "Yahoo-DocGen-App"}
    req = urllib.request.Request(download_url, headers=headers)
    
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    
    start_time = time.time()
    last_update_time = 0.0
    downloaded = 0
    chunk_size = 64 * 1024  # 64 KB

    try:
        with urllib.request.urlopen(req, timeout=30, context=ctx) as response:
            total_size = int(response.headers.get('Content-Length', 0))
            
            with open(target_path, 'wb') as out_file:
                while True:
                    if cancel_flag and cancel_flag():
                        out_file.close()
                        if os.path.exists(target_path):
                            os.remove(target_path)
                        return False
                        
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                        
                    out_file.write(chunk)
                    downloaded += len(chunk)
                    
                    now = time.time()
                    if progress_callback and (now - last_update_time > 0.1 or downloaded == total_size):
                        elapsed = now - start_time
                        speed = downloaded / elapsed if elapsed > 0 else 0
                        speed_str = f"{speed / (1024 * 1024):.1f} МБ/с" if speed > 1024 * 1024 else f"{speed / 1024:.0f} КБ/с"
                        progress_callback(downloaded, total_size, speed_str)
                        last_update_time = now
                        
        return True
    except Exception as e:
        logger.exception(f"Ошибка при скачивании обновления: {e}")
        if os.path.exists(target_path):
            try:
                os.remove(target_path)
            except Exception:
                pass
        raise e


def apply_update_and_restart(zip_path: str, target_dir: str = BASE_DIR) -> None:
    """
    Распаковывает обновление и запускает скрипт безопасной замены файлов на Windows.
    Гарантированно сохраняет существующие пользовательские шаблоны, данные и настройки!
    """
    temp_extract_dir = os.path.join(tempfile.gettempdir(), f"yahoo_update_{int(time.time())}")
    os.makedirs(temp_extract_dir, exist_ok=True)
    
    # 1. Распаковка архива во временную папку
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(temp_extract_dir)
        
    # Определяем корневую папку с файлами обновления (если архив упакован с подпапкой)
    source_dir = temp_extract_dir
    entries = [os.path.join(temp_extract_dir, e) for e in os.listdir(temp_extract_dir)]
    dirs = [e for e in entries if os.path.isdir(e)]
    if len(dirs) == 1 and len(entries) == 1:
        source_dir = dirs[0]
        
    pid = os.getpid()
    is_frozen = getattr(sys, 'frozen', False)
    exe_name = os.path.basename(sys.executable) if is_frozen else "Yahoo.exe"
    
    bat_path = os.path.join(tempfile.gettempdir(), f"yahoo_updater_{pid}.bat")
    
    # Создаем bat-скрипт обновления:
    # 1) Ожидает завершения текущего процесса Yahoo.exe
    # 2) Копирует новые файлы в target_dir, исключая templates, data, logs, result, settings.json
    # 3) Запускает обновленный Yahoo.exe
    # 4) Очищает временные файлы и самоудаляется
    bat_content = f"""@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

set "PID={pid}"
set "SOURCE_DIR={source_dir}"
set "TARGET_DIR={target_dir}"
set "EXE_NAME={exe_name}"
set "ZIP_FILE={zip_path}"

:: Ожидание завершения текущего процесса программы
:WAIT_PID
tasklist /FI "PID eq !PID!" 2>nul | findstr "!PID!" >nul
if %ERRORLEVEL% == 0 (
    timeout /t 1 /nobreak > nul
    goto WAIT_PID
)

:: Небольшая пауза для снятия блокировок файлов операционной системой
timeout /t 1 /nobreak > nul

:: Безопасное копирование обновленных файлов
:: Исключаем перезапись пользовательских данных: data, templates, result, logs, settings.json
robocopy "!SOURCE_DIR!" "!TARGET_DIR!" /E /IS /IT /XF settings.json /XD data templates result logs > nul 2>&1

:: Если robocopy недоступен, выполняем xcopy
if errorlevel 8 (
    xcopy "!SOURCE_DIR!\\*" "!TARGET_DIR!\\" /E /Y /I > nul 2>&1
)

:: Запуск обновленного приложения
cd /d "!TARGET_DIR!"
if exist "!TARGET_DIR!\\!EXE_NAME!" (
    start "" "!TARGET_DIR!\\!EXE_NAME!"
) else (
    if exist "!TARGET_DIR!\\Yahoo.exe" (
        start "" "!TARGET_DIR!\\Yahoo.exe"
    )
)

:: Очистка временных папок и архива обновления
rmdir /s /q "{temp_extract_dir}" > nul 2>&1
if exist "!ZIP_FILE!" del /f /q "!ZIP_FILE!" > nul 2>&1

:: Самоудаление bat-файла
(goto) 2>nul & del "%~f0"
"""

    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)
        
    logger.info(f"Запуск скрипта обновления: {bat_path}")
    
    # Запуск bat-скрипта в отдельном detached-процессе
    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        creationflags=CREATE_NO_WINDOW,
        close_fds=True
    )
    
    # Немедленное завершение работы текущей программы
    sys.exit(0)
