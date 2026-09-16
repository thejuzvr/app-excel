import os
import json
import threading
import datetime
import tempfile
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry

from config import TEMPLATES_DIR, RESULT_DIR, ASSETS_DIR, APP_VERSION, UPDATE_CHECK_ON_STARTUP
from generator import get_template_files, load_excel, generate_documents, sanitize_filename, format_fill_date
from updater import check_for_updates, download_update, apply_update_and_restart

# ==========================================
# СОХРАНЕНИЕ И ЗАГРУЗКА НАСТРОЕК (SETTINGS)
# ==========================================
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')

def load_settings():
    """Загружает сохраненные настройки пользователя из JSON"""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_settings():
    """Сохраняет текущие параметры (папку, формат, тему, стиль даты) в JSON"""
    try:
        out_folder = entry_output_path.get().strip() if 'entry_output_path' in globals() else RESULT_DIR
        out_format = output_format_var.get() if 'output_format_var' in globals() else 'word'
        app_mode = ctk.get_appearance_mode().lower()
        date_fmt = date_format_var.get() if 'date_format_var' in globals() else '29.07.2026 (Числовой)'
        
        data = {
            'output_folder': out_folder,
            'output_format': out_format,
            'appearance_mode': app_mode,
            'date_format_style': date_fmt
        }
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

# ==========================================
# БАЗОВЫЕ НАСТРОЙКИ СТИЛЯ
# ==========================================
saved_user_settings = load_settings()
initial_theme = saved_user_settings.get('appearance_mode', 'dark')
ctk.set_appearance_mode(initial_theme)
ctk.set_default_color_theme("blue")

df_data = None
all_templates = {}            # {template_name: BooleanVar}
template_checkboxes = {}      # {template_name: CTkCheckBox}
last_output_folder = saved_user_settings.get('output_folder', RESULT_DIR)


# ==========================================
# СТИЛИЗАЦИЯ ТАБЛИЦЫ ДЛЯ DARK/LIGHT
# ==========================================
def apply_treeview_theme(mode="dark"):
    """Настройка современного оформления ttk.Treeview под цвет CustomTkinter"""
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    if mode == "dark":
        bg_color = "#242831"
        fg_color = "#f1f2f6"
        header_bg = "#1b1e25"
        header_fg = "#dfe4ea"
        select_bg = "#2980b9"
        select_fg = "#ffffff"
    else:
        bg_color = "#ffffff"
        fg_color = "#2f3542"
        header_bg = "#f1f2f6"
        header_fg = "#2f3542"
        select_bg = "#3498db"
        select_fg = "#ffffff"

    style.configure(
        "Modern.Treeview",
        background=bg_color,
        foreground=fg_color,
        fieldbackground=bg_color,
        rowheight=28,
        font=("Segoe UI", 10),
        borderwidth=0
    )
    style.configure(
        "Modern.Treeview.Heading",
        background=header_bg,
        foreground=header_fg,
        font=("Segoe UI", 10, "bold"),
        relief="flat",
        borderwidth=1
    )
    style.map(
        "Modern.Treeview",
        background=[("selected", select_bg)],
        foreground=[("selected", select_fg)]
    )
    style.map(
        "Modern.Treeview.Heading",
        background=[("active", "#2f3542" if mode == "dark" else "#e4e7eb")]
    )


# ==========================================
# ЛОГИКА РАБОТЫ С EXCEL И ТАБЛИЦЕЙ
# ==========================================
def load_excel_gui():
    global df_data
    
    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=[("Файлы Excel", "*.xlsx *.xls")]
    )
    
    if not file_path:
        return
    
    try:
        with open(file_path, 'rb') as f:
            f.read(1)
    except PermissionError:
        messagebox.showerror(
            "Ошибка доступа",
            "Файл Excel открыт в другой программе.\nПожалуйста, закройте его и попробуйте снова."
        )
        return
    
    df_data = load_excel(file_path)
    
    if df_data is None or df_data.empty:
        messagebox.showerror("Ошибка", "Не удалось прочитать данные из файла или файл пуст.")
        return
    
    # Сброс поисковой строки
    search_entry.delete(0, ctk.END)
    
    # Заполнение таблицы
    render_table(df_data)
    
    lbl_file_info.configure(
        text=f"📄 {os.path.basename(file_path)} ({len(df_data)} чел.)",
        text_color="#2ecc71"
    )
    status_label.configure(
        text=f"✅ Загружен файл: {os.path.basename(file_path)} | Строк: {len(df_data)}",
        text_color="#2ecc71"
    )
    update_selection_counter()


def render_table(df_to_show):
    """Отрисовывает переданный датафрейм в Treeview"""
    tree_widget.delete(*tree_widget.get_children())
    
    if df_to_show is None or df_to_show.empty:
        return
    
    cols = list(df_to_show.columns)
    tree_widget["columns"] = cols
    
    # Настройка читаемой ширины колонок
    col_widths = {
        'табель_номер': 105,
        'ФИО': 200,
        'Дата_рождения': 110,
        'Должность': 150,
        'Подразделение': 150,
        'Паспорт_серия': 85,
        'Паспорт_номер': 95,
        'Паспорт_кем_выдан': 240,
        'Паспорт_дата_выдачи': 110,
        'Паспорт_адрес': 240,
        'Патент_серия': 90,
        'Патент_номер': 95,
        'Патент_дата_выдачи': 115,
        'Патент_действителен_до': 120,
    }
    
    for col in cols:
        tree_widget.heading(col, text=col, anchor="w")
        w = col_widths.get(col, 130)
        tree_widget.column(col, width=w, minwidth=70, anchor="w")
    
    for index, row in df_to_show.iterrows():
        values = [str(v) if pd.notna(v) else '' for v in row]
        tree_widget.insert("", "end", iid=index, values=values)


def filter_table(event=None):
    """Живой поиск и фильтрация сотрудников"""
    if df_data is None:
        return
        
    query = search_entry.get().strip().lower()
    
    if not query:
        render_table(df_data)
        update_selection_counter()
        return
    
    mask = df_data.astype(str).apply(lambda row: row.str.lower().str.contains(query, na=False).any(), axis=1)
    filtered_df = df_data[mask]
    render_table(filtered_df)
    update_selection_counter()


def select_all_employees():
    """Выбрать всех отображаемых сотрудников"""
    children = tree_widget.get_children()
    tree_widget.selection_set(children)
    update_selection_counter()


def deselect_all_employees():
    """Снять выделение со всех сотрудников"""
    tree_widget.selection_remove(tree_widget.selection())
    update_selection_counter()


def update_selection_counter(event=None):
    """Обновляет бейдж количества выбранных людей"""
    selected = len(tree_widget.selection())
    total = len(df_data) if df_data is not None else 0
    lbl_selected_people.configure(text=f"Выбрано: {selected} из {total}")


# ==========================================
# ЛОГИКА РАБОТЫ С ШАБЛОНАМИ
# ==========================================
def reload_templates_list():
    """Загружает шаблоны из папки и перестраивает список чекбоксов"""
    global all_templates, template_checkboxes
    
    template_files = get_template_files()
    
    for widget in templates_scroll_frame.winfo_children():
        widget.destroy()
        
    all_templates.clear()
    template_checkboxes.clear()
    
    for template_name in template_files:
        var = ctk.BooleanVar(value=True)
        all_templates[template_name] = var
        
        is_excel = template_name.lower().endswith('.xlsx')
        icon = "📊 " if is_excel else "📄 "
        
        cb = ctk.CTkCheckBox(
            templates_scroll_frame,
            text=f"{icon}{template_name}",
            variable=var,
            font=ctk.CTkFont(size=12),
            command=update_templates_counter
        )
        cb.pack(fill="x", padx=6, pady=3, anchor="w")
        template_checkboxes[template_name] = cb
        
    lbl_templates_header.configure(text=f"📄 Шаблоны документов ({len(template_files)})")
    update_templates_counter()


def select_all_templates():
    for var in all_templates.values():
        var.set(True)
    update_templates_counter()


def deselect_all_templates():
    for var in all_templates.values():
        var.set(False)
    update_templates_counter()


def update_templates_counter():
    selected = sum(1 for var in all_templates.values() if var.get())
    total = len(all_templates)
    lbl_selected_templates.configure(text=f"Выбрано: {selected} из {total}")


def filter_templates(event=None):
    """Поиск/фильтрация шаблонов в левой панели"""
    query = search_template_entry.get().strip().lower()
    for name, cb in template_checkboxes.items():
        if query in name.lower():
            cb.pack(fill="x", padx=6, pady=3, anchor="w")
        else:
            cb.pack_forget()


# ==========================================
# ПАПКА СОХРАНЕНИЯ И СТОРОННИЕ ДЕЙСТВИЯ
# ==========================================
def select_output_folder():
    global last_output_folder
    folder = filedialog.askdirectory(title="Выберите папку для сохранения документов", initialdir=last_output_folder)
    if folder:
        last_output_folder = folder
        entry_output_path.delete(0, ctk.END)
        entry_output_path.insert(0, folder)
        status_label.configure(text=f"📁 Папка сохранения изменена на: {folder}", text_color="#3498db")
        save_settings()


def open_output_folder():
    """Открывает папку сохранения в проводнике Windows"""
    target = entry_output_path.get().strip() or RESULT_DIR
    if not os.path.exists(target):
        try:
            os.makedirs(target)
        except Exception:
            pass
    try:
        os.startfile(target)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")


# ==========================================
# ПРОЦЕСС ГЕНЕРАЦИИ
# ==========================================
def generate_docs_gui():
    if df_data is None or df_data.empty:
        messagebox.showwarning("Внимание", "Сначала выберите файл Excel с сотрудниками!")
        return
    
    selected_items = tree_widget.selection()
    if not selected_items:
        messagebox.showwarning("Внимание", "Не выбрано ни одного сотрудника в таблице!")
        return
    
    selected_templates = [name for name, var in all_templates.items() if var.get()]
    if not selected_templates:
        messagebox.showwarning("Внимание", "Выберите хотя бы один шаблон для генерации!")
        return
    
    output_folder = entry_output_path.get().strip() or RESULT_DIR
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать целевую папку:\n{e}")
            return
    
    raw_date = date_picker.get()
    chosen_style = date_format_var.get() if 'date_format_var' in globals() else 'numeric'
    if "Прописью" in chosen_style:
        style_key = "text"
    elif "Официальный" in chosen_style:
        style_key = "quotes"
    else:
        style_key = "numeric"
    fill_date = format_fill_date(raw_date, style_key)
    output_format = output_format_var.get()
    
    people_count = len(selected_items)
    templates_count = len(selected_templates)
    total_files = people_count * templates_count * (2 if output_format == "both" else 1)
    format_map = {"word": "Word / Excel", "pdf": "PDF (.pdf)", "both": "Word / Excel + PDF"}
    
    confirm = messagebox.askyesno(
        "Подтверждение генерации",
        f"👥 Выбрано сотрудников: {people_count}\n"
        f"📄 Выбрано шаблонов: {templates_count}\n"
        f"📁 Будет создано файлов: {total_files}\n"
        f"⚙️ Формат: {format_map.get(output_format, output_format)}\n"
        f"📅 Дата в документах: {fill_date}\n"
        f"📂 Папка назначения: {output_folder}\n\n"
        f"Начать формирование документов?"
    )
    
    if not confirm:
        return
    
    set_ui_state(disabled=True)
    btn_open_result.configure(state="disabled")
    progress_bar.set(0)
    progress_label.configure(text="Подготовка к генерации...")
    status_label.configure(text="⏳ Идет генерация документов...", text_color="#f39c12")
    
    thread = threading.Thread(
        target=process_generation_thread,
        args=(selected_items, selected_templates, output_folder, output_format, fill_date),
        daemon=True
    )
    thread.start()


def process_generation_thread(selected_items, templates, output_folder, output_format, fill_date):
    def update_progress(current, total):
        progress = current / total if total > 0 else 0
        root.after(0, lambda: progress_bar.set(progress))
        root.after(0, lambda: progress_label.configure(
            text=f"Обработано {current} из {total} ({int(progress * 100)}%)"
        ))
    
    results = generate_documents(
        df_data,
        [int(i) for i in selected_items],
        templates,
        output_folder,
        output_format=output_format,
        fill_date=fill_date,
        progress_callback=update_progress
    )
    
    root.after(0, lambda: show_results(results, output_folder))


def show_results(results, output_folder):
    btn_open_result.configure(state="normal")
    set_ui_state(disabled=False)
    progress_bar.set(1.0)
    progress_label.configure(text="Генерация завершена!")
    
    if results['errors']:
        err_snippet = "\n".join(f"• {e}" for e in results['errors'][:4])
        messagebox.showwarning(
            "Генерация завершена с замечаниями",
            f"✅ Успешно создано: {results['success']}\n"
            f"❌ Ошибок: {len(results['errors'])}\n\n"
            f"Некоторые ошибки:\n{err_snippet}"
        )
        status_label.configure(
            text=f"⚠️ Завершено: {results['success']} успешно, {len(results['errors'])} ошибок.",
            text_color="#f39c12"
        )
    else:
        messagebox.showinfo(
            "✅ Генерация завершена",
            f"Все документы успешно сформированы!\n\n"
            f"Файлов создано: {results['success']}\n"
            f"Папка: {output_folder}"
        )
        status_label.configure(
            text=f"✅ Успешно создано файлов: {results['success']} в папке: {output_folder}",
            text_color="#2ecc71"
        )


def set_ui_state(disabled):
    state = "disabled" if disabled else "normal"
    btn_load_excel.configure(state=state)
    btn_select_output.configure(state=state)
    btn_select_all_templates.configure(state=state)
    btn_deselect_all_templates.configure(state=state)
    btn_reload_templates.configure(state=state)
    btn_select_all_emp.configure(state=state)
    btn_deselect_all_emp.configure(state=state)
    btn_generate.configure(state=state)
    
    for cb in template_checkboxes.values():
        cb.configure(state=state)


def toggle_theme():
    current = ctk.get_appearance_mode()
    new_mode = "light" if current == "Dark" else "dark"
    ctk.set_appearance_mode(new_mode)
    apply_treeview_theme(new_mode)
    btn_theme.configure(text="☀️ Светлая" if new_mode == "dark" else "🌙 Тёмная")
    save_settings()


# ==========================================
# ИНИЦИАЛИЗАЦИЯ ИНТЕРФЕЙСА
# ==========================================
root = ctk.CTk()
root.title("📄 Генератор документов Word / PDF PRO")
root.geometry("1220x860")
root.minsize(1050, 750)

icon_path = os.path.join(ASSETS_DIR, 'ico.ico')
if os.path.exists(icon_path):
    try:
        root.iconbitmap(icon_path)
    except Exception:
        pass

apply_treeview_theme("dark")

# ==========================================
# ДИАЛОГ И ЛОГИКА АВТООБНОВЛЕНИЯ (UPDATER)
# ==========================================
latest_update_info = None

def show_update_dialog(update_info, parent=None):
    """Модальное окно с описанием новой версии, прогрессом загрузки и установкой"""
    dialog = ctk.CTkToplevel(parent or root)
    dialog.title("🔄 Обновление программы")
    dialog.geometry("520x460")
    dialog.resizable(False, False)
    dialog.transient(parent or root)
    dialog.grab_set()

    # Центрирование относительно родительского окна
    try:
        p_x = (parent or root).winfo_x()
        p_y = (parent or root).winfo_y()
        p_w = (parent or root).winfo_width()
        p_h = (parent or root).winfo_height()
        d_x = p_x + (p_w - 520) // 2
        d_y = p_y + (p_h - 460) // 2
        dialog.geometry(f"520x460+{max(0, d_x)}+{max(0, d_y)}")
    except Exception:
        pass

    icon_p = os.path.join(ASSETS_DIR, 'ico.ico')
    if os.path.exists(icon_p):
        try:
            dialog.iconbitmap(icon_p)
        except Exception:
            pass

    header_box = ctk.CTkFrame(dialog, fg_color="transparent")
    header_box.pack(fill="x", padx=20, pady=(16, 8))

    lbl_new_ver = ctk.CTkLabel(
        header_box,
        text=f"🎉 Доступна новая версия {update_info.get('tag_name', '')}!",
        font=ctk.CTkFont(size=17, weight="bold"),
        text_color="#2ecc71"
    )
    lbl_new_ver.pack(anchor="w")

    lbl_curr_ver = ctk.CTkLabel(
        header_box,
        text=f"Текущая версия: v{APP_VERSION}  •  Новая: {update_info.get('tag_name', '')}",
        font=ctk.CTkFont(size=12),
        text_color="gray"
    )
    lbl_curr_ver.pack(anchor="w", pady=(2, 0))

    ctk.CTkLabel(
        dialog,
        text="📋 Что нового:",
        font=ctk.CTkFont(size=13, weight="bold")
    ).pack(anchor="w", padx=20, pady=(6, 4))

    txt_changelog = ctk.CTkTextbox(dialog, height=140, corner_radius=8, font=ctk.CTkFont(size=12))
    txt_changelog.pack(fill="x", padx=20, pady=(0, 10))
    raw_body = update_info.get('changelog') or 'Описание изменений отсутствует.'
    txt_changelog.insert("1.0", raw_body.strip())
    txt_changelog.configure(state="disabled")

    progress_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    progress_frame.pack(fill="x", padx=20, pady=(0, 10))

    lbl_prog_status = ctk.CTkLabel(
        progress_frame,
        text="Готово к загрузке",
        font=ctk.CTkFont(size=12),
        text_color="gray",
        anchor="w"
    )
    lbl_prog_status.pack(fill="x", pady=(0, 4))

    update_pbar = ctk.CTkProgressBar(progress_frame, height=12)
    update_pbar.pack(fill="x")
    update_pbar.set(0)

    btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
    btn_frame.pack(fill="x", padx=20, pady=(8, 16))

    is_downloading = [False]
    is_cancelled = [False]

    def on_cancel():
        if is_downloading[0]:
            is_cancelled[0] = True
            lbl_prog_status.configure(text="Отмена скачивания...", text_color="#e74c3c")
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    def start_download():
        download_url = update_info.get('download_url')
        if not download_url:
            messagebox.showerror("Ошибка", "Ссылка на файл обновления не найдена.", parent=dialog)
            return

        is_downloading[0] = True
        btn_update_action.configure(state="disabled")
        btn_close.configure(text="Отмена")
        lbl_prog_status.configure(text="Подключение к серверу...", text_color="#3498db")

        def download_worker():
            asset_name = update_info.get('asset_name') or 'update.zip'
            dest_file = os.path.join(tempfile.gettempdir(), asset_name)

            def progress_hook(downloaded, total, speed_str):
                if is_cancelled[0]:
                    return
                pct = downloaded / total if total > 0 else 0
                mb_down = downloaded / (1024 * 1024)
                mb_total = total / (1024 * 1024)
                dialog.after(0, lambda: update_pbar.set(pct))
                dialog.after(0, lambda: lbl_prog_status.configure(
                    text=f"Загрузка: {mb_down:.1f} / {mb_total:.1f} МБ ({int(pct * 100)}%) • {speed_str}",
                    text_color="#3498db"
                ))

            try:
                success = download_update(
                    download_url,
                    dest_file,
                    progress_callback=progress_hook,
                    cancel_flag=lambda: is_cancelled[0]
                )
                if not success or is_cancelled[0]:
                    return

                dialog.after(0, lambda: lbl_prog_status.configure(
                    text="Обновление скачано! Перезапуск программы...",
                    text_color="#27ae60"
                ))
                dialog.after(1000, lambda: apply_update_and_restart(dest_file))

            except Exception as e:
                is_downloading[0] = False
                dialog.after(0, lambda: btn_update_action.configure(state="normal"))
                dialog.after(0, lambda: lbl_prog_status.configure(
                    text=f"Ошибка: {e}",
                    text_color="#e74c3c"
                ))

        threading.Thread(target=download_worker, daemon=True).start()

    btn_update_action = ctk.CTkButton(
        btn_frame,
        text="⬇️ Скачать и установить",
        command=start_download,
        height=38,
        font=ctk.CTkFont(size=13, weight="bold"),
        fg_color="#27ae60",
        hover_color="#219150"
    )
    btn_update_action.pack(side="left", fill="x", expand=True, padx=(0, 8))

    btn_close = ctk.CTkButton(
        btn_frame,
        text="Закрыть",
        command=on_cancel,
        width=100,
        height=38,
        fg_color=("#94a3b8", "#34495e"),
        hover_color=("#64748b", "#2c3e50")
    )
    btn_close.pack(side="right")


def check_updates_gui(manual=True):
    """Запуск проверки обновлений (ручной или фоновый при старте)"""
    global latest_update_info

    if manual and 'btn_update' in globals():
        btn_update.configure(text="⏳ Проверка...", state="disabled")

    def worker():
        res = check_for_updates()

        def on_complete():
            global latest_update_info
            if 'btn_update' in globals():
                if manual:
                    btn_update.configure(text="🔄 Обновления", state="normal")

                if res.get("has_update"):
                    latest_update_info = res
                    tag = res.get("tag_name", "")
                    btn_update.configure(
                        text=f"🚀 Доступно {tag}",
                        fg_color="#e67e22",
                        hover_color="#d35400",
                        text_color="#ffffff",
                        state="normal"
                    )
                    show_update_dialog(res, root)
                elif not res.get("success"):
                    if manual:
                        messagebox.showwarning(
                            "Проверка обновлений",
                            res.get("error", "Не удалось связаться с сервером обновлений.")
                        )
                else:
                    if manual:
                        msg = res.get("message") or f"У вас установлена последняя версия (v{APP_VERSION})."
                        messagebox.showinfo("Обновления", msg)

        root.after(0, on_complete)

    threading.Thread(target=worker, daemon=True).start()


# ----------------------------------------------------
# 1. ШАПКА (HEADER)
# ----------------------------------------------------
header_card = ctk.CTkFrame(root, corner_radius=10, fg_color=("#e2e8f0", "#1e222a"))
header_card.pack(fill="x", padx=12, pady=(10, 6))

header_left = ctk.CTkFrame(header_card, fg_color="transparent")
header_left.pack(side="left", padx=14, pady=10)

lbl_title = ctk.CTkLabel(
    header_left,
    text="📄 Генератор документов",
    font=ctk.CTkFont(size=20, weight="bold")
)
lbl_title.pack(side="left", padx=(0, 10))

lbl_badge = ctk.CTkLabel(
    header_left,
    text="PRO",
    font=ctk.CTkFont(size=11, weight="bold"),
    fg_color="#3498db",
    text_color="#ffffff",
    corner_radius=6,
    padx=8,
    pady=2
)
lbl_badge.pack(side="left")

lbl_version = ctk.CTkLabel(
    header_left,
    text=f"v{APP_VERSION}",
    font=ctk.CTkFont(size=11),
    text_color="gray",
    padx=8
)
lbl_version.pack(side="left")

header_right = ctk.CTkFrame(header_card, fg_color="transparent")
header_right.pack(side="right", padx=14, pady=10)

btn_theme = ctk.CTkButton(
    header_right,
    text="☀️ Светлая",
    width=105,
    height=32,
    command=toggle_theme,
    fg_color=("#cbd5e1", "#2d3436"),
    hover_color=("#94a3b8", "#3d4852"),
    text_color=("#0f172a", "#f1f2f6")
)
btn_theme.pack(side="right")

btn_update = ctk.CTkButton(
    header_right,
    text="🔄 Обновления",
    width=125,
    height=32,
    command=lambda: check_updates_gui(manual=True),
    fg_color=("#cbd5e1", "#2d3436"),
    hover_color=("#94a3b8", "#3d4852"),
    text_color=("#0f172a", "#f1f2f6"),
    font=ctk.CTkFont(size=11, weight="bold")
)
btn_update.pack(side="right", padx=(0, 8))



# ----------------------------------------------------
# 2. ПАНЕЛЬ ПАРАМЕТРОВ (EXCEL, ДАТА, ПАПКА)
# ----------------------------------------------------
params_card = ctk.CTkFrame(root, corner_radius=10)
params_card.pack(fill="x", padx=12, pady=6)

# Блок выбора Excel
p_excel_frame = ctk.CTkFrame(params_card, fg_color="transparent")
p_excel_frame.pack(side="left", padx=10, pady=10)

btn_load_excel = ctk.CTkButton(
    p_excel_frame,
    text="📥 Загрузить Excel",
    command=load_excel_gui,
    width=140,
    height=36,
    font=ctk.CTkFont(weight="bold")
)
btn_load_excel.pack(side="left", padx=(0, 8))

lbl_file_info = ctk.CTkLabel(
    p_excel_frame,
    text="Файл не выбран",
    font=ctk.CTkFont(size=12),
    text_color="gray"
)
lbl_file_info.pack(side="left")

# Блок выбора даты
p_date_frame = ctk.CTkFrame(params_card, fg_color="transparent")
p_date_frame.pack(side="left", padx=12, pady=10)

ctk.CTkLabel(
    p_date_frame,
    text="📅 Дата:",
    font=ctk.CTkFont(weight="bold", size=12)
).pack(side="left", padx=(0, 6))

today = datetime.datetime.now().strftime('%d.%m.%Y')
date_picker = DateEntry(
    p_date_frame,
    width=11,
    background='#2980b9',
    foreground='white',
    borderwidth=1,
    date_pattern='dd.mm.yyyy',
    locale='ru_RU',
    font=('Segoe UI', 10)
)
date_picker.set_date(today)
date_picker.pack(side="left", padx=(0, 6))

DATE_STYLE_OPTIONS = [
    "29.07.2026 (Числовой)",
    "29 июля 2026 г. (Прописью)",
    "«29» июля 2026 г. (Официальный)"
]

saved_date_style = saved_user_settings.get('date_format_style', DATE_STYLE_OPTIONS[0])
if saved_date_style not in DATE_STYLE_OPTIONS:
    saved_date_style = DATE_STYLE_OPTIONS[0]

date_format_var = ctk.StringVar(value=saved_date_style)

date_format_menu = ctk.CTkOptionMenu(
    p_date_frame,
    values=DATE_STYLE_OPTIONS,
    variable=date_format_var,
    command=lambda _: save_settings(),
    width=195,
    height=36,
    font=ctk.CTkFont(size=11)
)
date_format_menu.pack(side="left")

# Блок папки сохранения
p_output_frame = ctk.CTkFrame(params_card, fg_color="transparent")
p_output_frame.pack(side="left", fill="x", expand=True, padx=10, pady=10)

ctk.CTkLabel(
    p_output_frame,
    text="📁 Папка сохранения:",
    font=ctk.CTkFont(weight="bold", size=12)
).pack(side="left", padx=(0, 6))

entry_output_path = ctk.CTkEntry(
    p_output_frame,
    height=36,
    placeholder_text=RESULT_DIR
)
entry_output_path.pack(side="left", fill="x", expand=True, padx=(0, 6))
entry_output_path.insert(0, last_output_folder)

btn_select_output = ctk.CTkButton(
    p_output_frame,
    text="Обзор...",
    command=select_output_folder,
    width=75,
    height=36
)
btn_select_output.pack(side="left", padx=(0, 6))

btn_open_folder = ctk.CTkButton(
    p_output_frame,
    text="📂 Открыть",
    command=open_output_folder,
    width=80,
    height=36,
    fg_color=("#94a3b8", "#34495e"),
    hover_color=("#64748b", "#2c3e50")
)
btn_open_folder.pack(side="left")


# ----------------------------------------------------
# 3. ОСНОВНАЯ РАБОЧАЯ ОБЛАСТЬ (ШАБЛОНЫ СЛЕВА, ТАБЛИЦА СПРАВА)
# ----------------------------------------------------
main_work_frame = ctk.CTkFrame(root, fg_color="transparent")
main_work_frame.pack(fill="both", expand=True, padx=12, pady=6)

# === ЛЕВАЯ КОЛОНКА: ШАБЛОНЫ ===
templates_panel = ctk.CTkFrame(main_work_frame, width=340, corner_radius=10)
templates_panel.pack(side="left", fill="y", padx=(0, 8))
templates_panel.pack_propagate(False)

# Шапка блока шаблонов
t_header_frame = ctk.CTkFrame(templates_panel, fg_color="transparent")
t_header_frame.pack(fill="x", padx=10, pady=(10, 4))

lbl_templates_header = ctk.CTkLabel(
    t_header_frame,
    text="📄 Шаблоны",
    font=ctk.CTkFont(size=14, weight="bold")
)
lbl_templates_header.pack(side="left")

lbl_selected_templates = ctk.CTkLabel(
    t_header_frame,
    text="0/0",
    font=ctk.CTkFont(size=11),
    text_color="#3498db"
)
lbl_selected_templates.pack(side="right")

# Поле поиска шаблона
search_template_entry = ctk.CTkEntry(
    templates_panel,
    height=30,
    placeholder_text="🔍 Фильтр шаблонов..."
)
search_template_entry.pack(fill="x", padx=10, pady=4)
search_template_entry.bind("<KeyRelease>", filter_templates)

# Кнопки быстрого выбора шаблонов
t_btn_frame = ctk.CTkFrame(templates_panel, fg_color="transparent")
t_btn_frame.pack(fill="x", padx=10, pady=4)

btn_select_all_templates = ctk.CTkButton(
    t_btn_frame,
    text="Выбрать все",
    command=select_all_templates,
    height=28,
    font=ctk.CTkFont(size=11)
)
btn_select_all_templates.pack(side="left", fill="x", expand=True, padx=(0, 4))

btn_deselect_all_templates = ctk.CTkButton(
    t_btn_frame,
    text="Снять все",
    command=deselect_all_templates,
    height=28,
    font=ctk.CTkFont(size=11),
    fg_color=("#94a3b8", "#34495e"),
    hover_color=("#64748b", "#2c3e50")
)
btn_deselect_all_templates.pack(side="left", fill="x", expand=True, padx=(0, 4))

btn_reload_templates = ctk.CTkButton(
    t_btn_frame,
    text="🔄",
    width=32,
    height=28,
    command=reload_templates_list
)
btn_reload_templates.pack(side="left")

# Скроллируемый контейнер для чекбоксов
templates_scroll_frame = ctk.CTkScrollableFrame(templates_panel, fg_color="transparent")
templates_scroll_frame.pack(fill="both", expand=True, padx=6, pady=6)


# === ПРАВАЯ КОЛОНКА: ТАБЛИЦА СОТРУДНИКОВ ===
table_panel = ctk.CTkFrame(main_work_frame, corner_radius=10)
table_panel.pack(side="left", fill="both", expand=True)

# Шапка таблицы с поиском и действиями
table_top_bar = ctk.CTkFrame(table_panel, fg_color="transparent")
table_top_bar.pack(fill="x", padx=12, pady=(10, 6))

lbl_table_header = ctk.CTkLabel(
    table_top_bar,
    text="👥 Сотрудники",
    font=ctk.CTkFont(size=14, weight="bold")
)
lbl_table_header.pack(side="left", padx=(0, 10))

lbl_selected_people = ctk.CTkLabel(
    table_top_bar,
    text="Выбрано: 0",
    font=ctk.CTkFont(size=12, weight="bold"),
    text_color="#3498db"
)
lbl_selected_people.pack(side="left", padx=(0, 15))

# Поле живого поиска сотрудников
search_entry = ctk.CTkEntry(
    table_top_bar,
    height=32,
    width=240,
    placeholder_text="🔍 Поиск (ФИО, должность...)"
)
search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
search_entry.bind("<KeyRelease>", filter_table)

btn_select_all_emp = ctk.CTkButton(
    table_top_bar,
    text="Выбрать всех",
    command=select_all_employees,
    height=30,
    width=100,
    font=ctk.CTkFont(size=11)
)
btn_select_all_emp.pack(side="right", padx=(4, 0))

btn_deselect_all_emp = ctk.CTkButton(
    table_top_bar,
    text="Снять выбор",
    command=deselect_all_employees,
    height=30,
    width=90,
    font=ctk.CTkFont(size=11),
    fg_color=("#94a3b8", "#34495e"),
    hover_color=("#64748b", "#2c3e50")
)
btn_deselect_all_emp.pack(side="right", padx=(0, 4))

# Контейнер для Treeview с вертикальным и горизонтальным скроллом
table_container = ctk.CTkFrame(table_panel, fg_color="transparent")
table_container.pack(fill="both", expand=True, padx=12, pady=(0, 10))

table_container.grid_rowconfigure(0, weight=1)
table_container.grid_columnconfigure(0, weight=1)

tree_widget = ttk.Treeview(
    table_container,
    show="headings",
    selectmode="extended",
    style="Modern.Treeview"
)
tree_widget.grid(row=0, column=0, sticky="nsew")

vsb = ctk.CTkScrollbar(table_container, orientation="vertical", command=tree_widget.yview)
vsb.grid(row=0, column=1, sticky="ns")

hsb = ctk.CTkScrollbar(table_container, orientation="horizontal", command=tree_widget.xview)
hsb.grid(row=1, column=0, sticky="ew")

tree_widget.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
tree_widget.bind("<<TreeviewSelect>>", update_selection_counter)


# ----------------------------------------------------
# 4. НИЖНЯЯ ПАНЕЛЬ ДЕЙСТВИЙ (FOOTER)
# ----------------------------------------------------
bottom_card = ctk.CTkFrame(root, corner_radius=10)
bottom_card.pack(fill="x", padx=12, pady=(6, 10))

# Верхний ряд нижнего блока: выбор формата
fmt_row = ctk.CTkFrame(bottom_card, fg_color="transparent")
fmt_row.pack(fill="x", padx=14, pady=(8, 4))

ctk.CTkLabel(
    fmt_row,
    text="📄 Формат документов:",
    font=ctk.CTkFont(size=12, weight="bold")
).pack(side="left", padx=(0, 12))

output_format_var = ctk.StringVar(value=saved_user_settings.get('output_format', 'word'))

radio_word = ctk.CTkRadioButton(
    fmt_row,
    text="Word / Excel",
    variable=output_format_var,
    value="word",
    font=ctk.CTkFont(size=12),
    command=save_settings
)
radio_word.pack(side="left", padx=8)

radio_pdf = ctk.CTkRadioButton(
    fmt_row,
    text="Только PDF",
    variable=output_format_var,
    value="pdf",
    font=ctk.CTkFont(size=12),
    command=save_settings
)
radio_pdf.pack(side="left", padx=8)

radio_both = ctk.CTkRadioButton(
    fmt_row,
    text="Word / Excel + PDF",
    variable=output_format_var,
    value="both",
    font=ctk.CTkFont(size=12),
    command=save_settings
)
radio_both.pack(side="left", padx=8)

# Ряд прогресса
prog_row = ctk.CTkFrame(bottom_card, fg_color="transparent")
prog_row.pack(fill="x", padx=14, pady=4)

progress_bar = ctk.CTkProgressBar(prog_row, height=12)
progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 12))
progress_bar.set(0)

progress_label = ctk.CTkLabel(
    prog_row,
    text="Ожидание запуска...",
    font=ctk.CTkFont(size=12),
    width=220,
    anchor="e"
)
progress_label.pack(side="right")

# Ряд кнопок генерации
action_row = ctk.CTkFrame(bottom_card, fg_color="transparent")
action_row.pack(fill="x", padx=14, pady=(4, 8))

btn_generate = ctk.CTkButton(
    action_row,
    text="🚀 Сгенерировать выбранные документы",
    command=generate_docs_gui,
    height=44,
    font=ctk.CTkFont(size=15, weight="bold"),
    fg_color="#27ae60",
    hover_color="#219150"
)
btn_generate.pack(side="left", fill="x", expand=True, padx=(0, 8))

btn_open_result = ctk.CTkButton(
    action_row,
    text="📂 Открыть результаты",
    command=open_output_folder,
    height=44,
    width=180,
    font=ctk.CTkFont(size=13, weight="bold"),
    fg_color=("#94a3b8", "#2c3e50"),
    hover_color=("#64748b", "#1a252f"),
    state="disabled"
)
btn_open_result.pack(side="left")

# Статусная строка внизу
status_bar = ctk.CTkFrame(root, fg_color="transparent")
status_bar.pack(fill="x", padx=14, pady=(0, 6))

status_label = ctk.CTkLabel(
    status_bar,
    text="Загрузите файл Excel, выберите людей и шаблоны для генерации",
    text_color="gray",
    font=ctk.CTkFont(size=11),
    anchor="w"
)
status_label.pack(side="left", fill="x", expand=True)

lbl_hint = ctk.CTkLabel(
    status_bar,
    text="💡 Ctrl/Shift+клик для выбора нескольких | Поиск фильтрует список в реальном времени",
    text_color="gray",
    font=ctk.CTkFont(size=11),
    anchor="e"
)
lbl_hint.pack(side="right")

def on_closing():
    save_settings()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# Начальная загрузка списка шаблонов
reload_templates_list()

# Фоновая проверка наличия обновлений
if UPDATE_CHECK_ON_STARTUP:
    root.after(1500, lambda: check_updates_gui(manual=False))

if __name__ == "__main__":
    print(f"\n=== ЗАПУСК ГЕНЕРАТОРА PRO (Modern UI) ===")
    print(f"Версия: v{APP_VERSION}")
    print(f"Шаблонов найдено: {len(all_templates)}")
    print(f"Тема: {ctk.get_appearance_mode()}")
    print("=" * 50)
    root.mainloop()