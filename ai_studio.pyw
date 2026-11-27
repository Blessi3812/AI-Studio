import winsound
import asyncio
import re
import edge_tts
import base64
import importlib.util
import io
import json
import os
import subprocess
import sys
import threading
import time
from tkinter import filedialog
import customtkinter as ctk
import cv2
import numpy as np
from PIL import Image, ImageGrab
from pypdf import PdfReader
import docx
import openpyxl
import requests
import scipy.io.wavfile as wav
import sounddevice as sd
import speech_recognition as sr

OLLAMA_BASE = "http://127.0.0.1:11434"
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "ai_assistant_data")
CHATS_DIR = os.path.join(DATA_DIR, "chats")
PLUGINS_DIR = os.path.join(DATA_DIR, "plugins")
MEMORY_FILE = os.path.join(DATA_DIR, "memory.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")

os.makedirs(CHATS_DIR, exist_ok=True)
os.makedirs(PLUGINS_DIR, exist_ok=True)

THEMES = {
    "Claude Тёмная (Уютная)": {
        "mode": "dark", "bg_main": "#21201e", "bg_sidebar": "#191817",
        "bg_card": "#2a2927", "user_bubble": "#32302d", "ai_bubble": "#21201e",
        "border": "#3a3834", "accent": "#d97757", "accent_hover": "#c56444",
        "text": "#eceae4", "muted": "#9c988f"
    },
    "Claude Светлая (Кремовая)": {
        "mode": "light", "bg_main": "#faf9f5", "bg_sidebar": "#f0eee6",
        "bg_card": "#ffffff", "user_bubble": "#f0eee6", "ai_bubble": "#faf9f5",
        "border": "#e2dfd5", "accent": "#d97757", "accent_hover": "#c56444",
        "text": "#2b2a27", "muted": "#7c7970"
    },
    "Фиолетовый Неон": {
        "mode": "dark", "bg_main": "#151320", "bg_sidebar": "#0f0d17",
        "bg_card": "#1f1c2e", "user_bubble": "#2a2540", "ai_bubble": "#151320",
        "border": "#362f52", "accent": "#8b5cf6", "accent_hover": "#7c3aed",
        "text": "#f3f0ff", "muted": "#948fb3"
    },
    "Изумрудная Ночь": {
        "mode": "dark", "bg_main": "#101816", "bg_sidebar": "#0b110f",
        "bg_card": "#172420", "user_bubble": "#1e332d", "ai_bubble": "#101816",
        "border": "#28423b", "accent": "#10b981", "accent_hover": "#059669",
        "text": "#ecfdf5", "muted": "#86a89e"
    },
    "Полночный Океан": {
        "mode": "dark", "bg_main": "#0f172a", "bg_sidebar": "#0a101f",
        "bg_card": "#182238", "user_bubble": "#23314d", "ai_bubble": "#0f172a",
        "border": "#304266", "accent": "#38bdf8", "accent_hover": "#0284c7",
        "text": "#f8fafc", "muted": "#94a3b8"
    }
}



class SkillManager:
    def __init__(self):
        self.skills = {}
        self.enabled_skills = {}
        self.create_sample_plugin()
        self.reload_all()

    def create_sample_plugin(self):
        sample_path = os.path.join(PLUGINS_DIR, "system_info.py")
        if not os.path.exists(sample_path):
            code = (
                "import platform, datetime\n"
                "def run_tool(args):\n"
                "    return f'Время: {datetime.datetime.now().strftime(\"%Y-%m-%d %H:%M:%S\")}. ОС: {platform.system()} {platform.release()}'\n"
                "def register_skill():\n"
                "    return {'name': 'get_system_time', 'title': 'Время и инфо о ПК', 'description': 'Узнать текущее время и версию Windows.', 'parameters': {'type': 'object', 'properties': {}}, 'handler': run_tool}\n"
            )
            with open(sample_path, "w", encoding="utf-8") as f:
                f.write(code)

    def reload_all(self):
        old_states = dict(self.enabled_skills)
        self.skills = {}
        self.register(
            "manage_files",
            "Полный доступ к файлам ПК (Claude Code)",
            "Работа с любыми файлами и папками на ПК. action: 'read' (прочитать), 'write' (создать/перезаписать файл), 'append' (дописать в конец), 'list' (список файлов в папке), 'search' (поиск файла по имени в папке), 'mkdir' (создать папку), 'delete' (удалить файл). path: полный путь или путь от домашней папки (например 'Desktop/file.txt'). content: текст или маска поиска.",
            {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["read", "write", "append", "list", "search", "mkdir", "delete"]},
                    "path": {"type": "string"},
                    "content": {"type": "string"}
                },
                "required": ["action", "path"]
            },
            self._tool_manage_files,
            old_states.get("manage_files", True)
        )

        self.register("run_terminal", "Терминал Windows (Shell)", "Выполняет консольную команду в Windows.", {"type": "object", "properties": {"command": {"type": "string"}}, "required": ["command"]}, self._tool_run_terminal, old_states.get("run_terminal", False))
        self.register("web_search", "Поиск в интернете", "Ищет актуальную информацию в сети.", {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}, self._tool_web_search, old_states.get("web_search", True))
        self.register(
            "office_docs",
            "Word и Excel документы (.docx / .xlsx)",
            "Создаёт, читает и редактирует файлы Word (.docx) и Excel (.xlsx). action: 'read' или 'write', path: путь к файлу, content: для Word — текст (строки через \\n, заголовки начинаются с '# '), для Excel — JSON-массив строк, например '[[\"Имя\", \"Возраст\"], [\"Анна\", 25]]'.",
            {
                "type": "object",
                "properties": {
                    "action": {"type": "string", "enum": ["read", "write"]},
                    "path": {"type": "string"},
                    "content": {"type": "string"}
                },
                "required": ["action", "path"]
            },
            self._tool_office_docs,
            old_states.get("office_docs", True)
        )

        for filename in os.listdir(PLUGINS_DIR):
            if filename.endswith(".py") and not filename.startswith("_"):
                filepath = os.path.join(PLUGINS_DIR, filename)
                try:
                    spec = importlib.util.spec_from_file_location(filename[:-3], filepath)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    if hasattr(module, "register_skill"):
                        info = module.register_skill()
                        name = info["name"]
                        self.register(name, info.get("title", name), info.get("description", ""), info.get("parameters", {"type": "object", "properties": {}}), info["handler"], old_states.get(name, True))
                except Exception as e:
                    print(f"Ошибка плагина {filename}: {e}")

    def register(self, name, title, description, parameters, handler, default_enabled=True):
        self.skills[name] = {"name": name, "title": title, "description": description, "parameters": parameters, "handler": handler}
        self.enabled_skills[name] = default_enabled

    def get_ollama_tools(self):
        return [{"type": "function", "function": {"name": s["name"], "description": s["description"], "parameters": s["parameters"]}} for n, s in self.skills.items() if self.enabled_skills.get(n, False)]

    def execute(self, name, arguments):
        if name not in self.skills or not self.enabled_skills.get(name, False):
            return f"Скилл '{name}' отключён."
        try:
            return str(self.skills[name]["handler"](arguments))
        except Exception as e:
            return f"Ошибка скилла '{name}': {e}"

    def _resolve_pc_path(self, raw_path):
        """Преобразует путь от модели в точный путь Windows, включая Desktop, Documents и т.д."""
        if not raw_path:
            return os.path.expanduser("~")
        p = os.path.expandvars(os.path.expanduser(raw_path.strip()))
        # Если путь относительный (например "Desktop/note.txt"), привязываем к папке пользователя
        if not os.path.isabs(p):
            p = os.path.join(os.path.expanduser("~"), p)
        return os.path.normpath(p)

    def _tool_manage_files(self, args):
        action = args.get("action")
        path = self._resolve_pc_path(args.get("path", ""))
        content = args.get("content", "")

        try:
            if action == "list":
                if not os.path.exists(path):
                    return f"Папка не найдена: {path}"
                if not os.path.isdir(path):
                    return f"Это файл, а не папка: {path}"
                items = []
                for entry in sorted(os.listdir(path))[:200]:
                    full = os.path.join(path, entry)
                    tag = "[ПАПКА]" if os.path.isdir(full) else f"[ФАЙЛ, {os.path.getsize(full)} Б]"
                    items.append(f"{tag} {entry}")
                return f"Содержимое {path}:\n" + ("\n".join(items) if items else "(папка пуста)")

            elif action == "search":
                if not os.path.exists(path):
                    return f"Папка для поиска не найдена: {path}"
                query = content.lower().strip()
                matches = []
                for root, dirs, files in os.walk(path):
                    # Пропускаем тяжёлые системные директории для скорости
                    dirs[:] = [d for d in dirs if d not in ("Windows", "Program Files", "Program Files (x86)", "$Recycle.Bin", "AppData", ".git", "node_modules")]
                    for name in files + dirs:
                        if not query or query in name.lower():
                            matches.append(os.path.join(root, name))
                            if len(matches) >= 50:
                                break
                    if len(matches) >= 50:
                        break
                return f"Найдено в {path}:\n" + ("\n".join(matches) if matches else "Совпадений не найдено.")

            elif action == "mkdir":
                os.makedirs(path, exist_ok=True)
                return f"Папка создана: {path}"

            elif action == "read":
                if not os.path.exists(path):
                    return f"Файл не найден: {path}"
                if os.path.isdir(path):
                    return f"Это папка, используйте action='list': {path}"
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()[:15000]

            elif action == "write":
                folder = os.path.dirname(path)
                if folder:
                    os.makedirs(folder, exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                return f"Файл успешно создан/перезаписан: {path}"

            elif action == "append":
                folder = os.path.dirname(path)
                if folder:
                    os.makedirs(folder, exist_ok=True)
                with open(path, "a", encoding="utf-8") as f:
                    f.write(content)
                return f"Текст дописан в файл: {path}"

            elif action == "delete":
                if not os.path.exists(path):
                    return f"Объект не найден: {path}"
                if os.path.isfile(path):
                    os.remove(path)
                    return f"Файл удалён: {path}"
                return "Для безопасности через delete можно удалять только отдельные файлы."

            return "Неизвестное действие для manage_files."
        except PermissionError:
            return f"Нет прав доступа к: {path}"
        except Exception as e:
            return f"Ошибка при работе с {path}: {e}"


    def _tool_run_terminal(self, args):
        try:
            res = subprocess.run(args.get("command", ""), shell=True, capture_output=True, text=True, timeout=20)
            out = (res.stdout + "\n" + res.stderr).strip()
            return out[:6000] if out else "Выполнено без вывода."
        except Exception as e:
            return f"Ошибка: {e}"

    def _tool_web_search(self, args):
        try:
            from duckduckgo_search import DDGS
            with DDGS() as ddgs:
                results = [f"{r['title']}: {r['body']} ({r['href']})" for r in ddgs.text(args.get("query", ""), max_results=4)]
            return "\n\n".join(results) if results else "Ничего не найдено."
        except ImportError:
            return "Установите: pip install duckduckgo-search"
        except Exception as e:
            return f"Ошибка поиска: {e}"
    def _tool_office_docs(self, args):
        action = args.get("action")
        path = self._resolve_pc_path(args.get("path", ""))
        content = args.get("content", "")
        ext = os.path.splitext(path).lower()

        try:
            if ext == ".docx":
                if action == "read":
                    if not os.path.exists(path):
                        return f"Файл не найден: {path}"
                    doc = docx.Document(path)
                    parts = [p.text for p in doc.paragraphs if p.text.strip()]
                    for table in doc.tables:
                        for row in table.rows:
                            parts.append(" | ".join(cell.text.strip() for cell in row.cells))
                    return "\n".join(parts)[:15000] or "(документ пуст)"

                elif action == "write":
                    folder = os.path.dirname(path)
                    if folder:
                        os.makedirs(folder, exist_ok=True)
                    doc = docx.Document()
                    for line in content.splitlines():
                        line_s = line.strip()
                        if line_s.startswith("# "):
                            doc.add_heading(line_s[2:], level=1)
                        elif line_s.startswith("## "):
                            doc.add_heading(line_s[3:], level=2)
                        elif line_s:
                            doc.add_paragraph(line_s)
                    doc.save(path)
                    return f"Документ Word сохранён: {path}"

            elif ext == ".xlsx":
                if action == "read":
                    if not os.path.exists(path):
                        return f"Файл не найден: {path}"
                    wb = openpyxl.load_workbook(path, data_only=True)
                    out = []
                    for sheet in wb.sheetnames:
                        ws = wb[sheet]
                        out.append(f"--- Лист: {sheet} ---")
                        for row in ws.iter_rows(values_only=True):
                            if any(v is not None for v in row):
                                out.append(" | ".join("" if v is None else str(v) for v in row))
                    return "\n".join(out)[:15000] or "(таблица пуста)"

                elif action == "write":
                    folder = os.path.dirname(path)
                    if folder:
                        os.makedirs(folder, exist_ok=True)
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    try:
                        rows = json.loads(content)
                        if isinstance(rows, list):
                            for r in rows:
                                ws.append(r if isinstance(r, list) else [str(r)])
                        else:
                            ws.append([str(content)])
                    except Exception:
                        for line in content.splitlines():
                            ws.append([cell.strip() for cell in line.split(";")])
                    wb.save(path)
                    return f"Таблица Excel сохранена: {path}"

            return "Поддерживаются только файлы с расширением .docx и .xlsx."
        except Exception as e:
            return f"Ошибка при работе с офисным файлом: {e}"

class AIStudioApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.config_data = self.load_config()
        self.current_theme_name = self.config_data.get("theme", "Claude Тёмная (Уютная)")
        if self.current_theme_name not in THEMES:
            self.current_theme_name = "Claude Тёмная (Уютная)"

        self.title("AI Studio")
        self.geometry("1180x780")
        self.minsize(920, 600)

        self.skill_manager = SkillManager()
        self.current_chat_id = None
        self.chats = {}
        self.memory_facts = self.load_memory()
        self.attached_images = []
        self.attached_files_text = []
        self.attached_labels = []
        self.is_generating = False
        self.is_recording = False
        self.active_ai_label = None

        self.setup_ui()
        self.apply_theme(self.current_theme_name)
        self.refresh_models()
        self.load_all_chats()

        if not self.chats:
            self.create_new_chat()
        else:
            first_id = sorted(self.chats.keys(), key=lambda k: self.chats[k].get("updated_at", 0), reverse=True)
            self.select_chat(first_id[0])

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"theme": "Claude Тёмная (Уютная)"}

    def save_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=265, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(3, weight=1)
        self.sidebar.grid_propagate(False)

        self.brand_lbl = ctk.CTkLabel(self.sidebar, text="✦ AI Studio", font=("Segoe UI", 20, "bold"))
        self.brand_lbl.grid(row=0, column=0, padx=18, pady=(18, 12), sticky="w")

        self.new_chat_btn = ctk.CTkButton(self.sidebar, text="+  Новый чат", font=("Segoe UI", 14, "bold"), height=40, corner_radius=10, command=self.create_new_chat)
        self.new_chat_btn.grid(row=1, column=0, padx=14, pady=(0, 12), sticky="ew")

        self.chats_label = ctk.CTkLabel(self.sidebar, text="НЕДАВНИЕ ЧАТЫ", font=("Segoe UI", 11, "bold"))
        self.chats_label.grid(row=2, column=0, padx=18, pady=(4, 2), sticky="w")

        self.chat_list_frame = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.chat_list_frame.grid(row=3, column=0, padx=8, pady=4, sticky="nsew")

        self.skills_btn = ctk.CTkButton(self.sidebar, text="🧩  Скиллы и плагины", anchor="w", height=38, corner_radius=8, command=self.open_skills_window)
        self.skills_btn.grid(row=4, column=0, padx=14, pady=(6, 4), sticky="ew")

        self.memory_btn = ctk.CTkButton(self.sidebar, text="🧠  Память ассистента", anchor="w", height=38, corner_radius=8, command=self.open_memory_window)
        self.memory_btn.grid(row=5, column=0, padx=14, pady=(0, 16), sticky="ew")

        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=20, pady=14)
        self.main_area.grid_rowconfigure(1, weight=1)
        self.main_area.grid_columnconfigure(0, weight=1)

        self.top_bar = ctk.CTkFrame(self.main_area, height=44, corner_radius=14, fg_color="transparent")
        self.top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.model_var = ctk.StringVar(value="qwen3.5:4b")
        self.model_menu = ctk.CTkOptionMenu(self.top_bar, variable=self.model_var, values=["qwen3.5:4b"], width=175, height=32, corner_radius=12, font=("Segoe UI", 12, "bold"))
        self.model_menu.pack(side="left", padx=(4, 4), pady=6)

        self.refresh_models_btn = ctk.CTkButton(self.top_bar, text="↻", width=32, height=32, corner_radius=12, command=self.refresh_models)
        self.refresh_models_btn.pack(side="left", padx=2, pady=6)

        self.theme_var = ctk.StringVar(value=self.current_theme_name)
        self.theme_menu = ctk.CTkOptionMenu(self.top_bar, variable=self.theme_var, values=list(THEMES.keys()), width=200, height=32, corner_radius=12, font=("Segoe UI", 12), command=self.apply_theme)
        self.theme_menu.pack(side="left", padx=(10, 4), pady=6)

        self.wake_var = ctk.BooleanVar(value=False)
        self.wake_switch = ctk.CTkSwitch(
            self.top_bar, text="👂 Отклик:", variable=self.wake_var,
            font=("Segoe UI", 12), command=self.toggle_wake_listener
        )
        self.wake_switch.pack(side="left", padx=(14, 4), pady=6)

        self.wake_word_entry = ctk.CTkEntry(
            self.top_bar, width=95, height=30, corner_radius=10, border_width=1,
            font=("Segoe UI", 12), placeholder_text="Слово..."
        )
        self.wake_word_entry.insert(0, "чупик")
        self.wake_word_entry.pack(side="left", padx=2, pady=6)

        self.tts_var = ctk.BooleanVar(value=True)
        self.tts_switch = ctk.CTkSwitch(
            self.top_bar, text="🔊 Озвучка", variable=self.tts_var,
            font=("Segoe UI", 12), command=self.on_tts_toggle
        )
        self.tts_switch.pack(side="left", padx=(14, 4), pady=6)

        self.stop_tts_btn = ctk.CTkButton(
            self.top_bar, text="🔇", width=32, height=30, corner_radius=10,
            fg_color="transparent", border_width=1, command=self.stop_speaking
        )
        self.stop_tts_btn.pack(side="left", padx=(2, 4), pady=6)


        self.delete_chat_btn = ctk.CTkButton(self.top_bar, text="🗑", width=34, height=32, corner_radius=12, fg_color="transparent", border_width=1, command=self.delete_current_chat)
        self.delete_chat_btn.pack(side="right", padx=4, pady=6)

        self.chat_scroll = ctk.CTkScrollableFrame(self.main_area, fg_color="transparent")
        self.chat_scroll.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        self.chat_scroll.grid_columnconfigure(0, weight=1)

        # Мягкая плавающая капсула ввода в стиле Claude
        self.input_box_wrapper = ctk.CTkFrame(self.main_area, corner_radius=22, border_width=1)
        self.input_box_wrapper.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 6))
        self.input_box_wrapper.grid_columnconfigure(0, weight=1)

        self.attachments_frame = ctk.CTkFrame(self.input_box_wrapper, fg_color="transparent")

        self.input_textbox = ctk.CTkTextbox(self.input_box_wrapper, height=68, font=("Segoe UI", 14), wrap="word", fg_color="transparent", border_width=0)
        self.input_textbox.grid(row=1, column=0, padx=18, pady=(12, 4), sticky="ew")
        self.input_textbox.bind("<Return>", self.on_enter_pressed)
        self.input_textbox.bind("<FocusIn>", lambda e: self.input_box_wrapper.configure(border_color=THEMES[self.current_theme_name]["accent"]))
        self.input_textbox.bind("<FocusOut>", lambda e: self.input_box_wrapper.configure(border_color=THEMES[self.current_theme_name]["border"]))
        self.bind_all("<Control-KeyPress>", self.on_global_ctrl_key)

        self.toolbar = ctk.CTkFrame(self.input_box_wrapper, fg_color="transparent", height=38)
        self.toolbar.grid(row=2, column=0, padx=14, pady=(0, 10), sticky="ew")

        self.attach_btn = ctk.CTkButton(self.toolbar, text="＋ Файл", width=88, height=32, corner_radius=14, font=("Segoe UI", 12), command=self.attach_media)
        self.attach_btn.pack(side="left", padx=3)

        self.voice_btn = ctk.CTkButton(self.toolbar, text="🎙 Голос", width=88, height=32, corner_radius=14, font=("Segoe UI", 12), command=self.toggle_voice_recording)
        self.voice_btn.pack(side="left", padx=3)

        self.hint_lbl = ctk.CTkLabel(self.toolbar, text="Ctrl+V: скриншот или файл  •  Shift+Enter: новая строка", font=("Segoe UI", 11))
        self.hint_lbl.pack(side="left", padx=12)

        self.send_btn = ctk.CTkButton(self.toolbar, text="↑", width=36, height=36, corner_radius=18, font=("Segoe UI", 16, "bold"), command=self.send_message)
        self.send_btn.pack(side="right", padx=2)


    def on_enter_pressed(self, event):
        if event.state & 0x1:
            return None
        self.send_message()
        return "break"

    def on_global_ctrl_key(self, event):
        # keycode 86 — физическая клавиша V (работает и в русской, и в английской раскладке)
        if event.keycode == 86:
            return self.handle_clipboard_paste()
        return None

    def handle_clipboard_paste(self):
        try:
            clip = ImageGrab.grabclipboard()
            if isinstance(clip, Image.Image):
                img = clip.convert("RGB")
                img.thumbnail((1024, 1024))
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=85)
                self.attached_images.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
                self.attached_labels.append(f"🖼 Скриншот ({img.width}x{img.height})")
                self.update_attachments_bar()
                return "break"
            elif isinstance(clip, list):
                self.process_file_paths(clip)
                return "break"
        except Exception:
            pass
        try:
            text = self.clipboard_get()
            if text:
                self.input_textbox.insert("insert", text)
                return "break"
        except Exception:
            pass
        return None

    def apply_theme(self, theme_name):
        self.current_theme_name = theme_name
        self.config_data["theme"] = theme_name
        self.save_config()
        t = THEMES[theme_name]
        ctk.set_appearance_mode(t["mode"])

        self.configure(fg_color=t["bg_main"])
        self.sidebar.configure(fg_color=t["bg_sidebar"])
        self.brand_lbl.configure(text_color=t["text"])
        self.chats_label.configure(text_color=t["muted"])

        # Принудительно красим оба скролл-фрейма и их внутренние холсты
        self.chat_list_frame.configure(fg_color=t["bg_sidebar"])
        self.chat_list_frame._parent_canvas.configure(bg=t["bg_sidebar"])
        self.chat_scroll.configure(fg_color=t["bg_main"])
        self.chat_scroll._parent_canvas.configure(bg=t["bg_main"])

        self.main_area.configure(fg_color=t["bg_main"])
        self.top_bar.configure(fg_color=t["bg_main"])
        self.input_box_wrapper.configure(fg_color=t["bg_card"], border_color=t["border"])
        self.input_textbox.configure(text_color=t["text"])
        self.toolbar.configure(fg_color=t["bg_card"])
        self.hint_lbl.configure(text_color=t["muted"])
        self.wake_word_entry.configure(fg_color=t["bg_card"], border_color=t["border"], text_color=t["text"])

        for sw in [self.wake_switch, self.tts_switch]:
            sw.configure(text_color=t["text"], progress_color=t["accent"])

        for btn in [self.new_chat_btn, self.send_btn]:
            btn.configure(fg_color=t["accent"], hover_color=t["accent_hover"], text_color="#ffffff")
        self.refresh_models_btn.configure(fg_color=t["user_bubble"], hover_color=t["accent"], text_color=t["text"])
        for menu in [self.model_menu, self.theme_menu]:
            menu.configure(fg_color=t["user_bubble"], button_color=t["user_bubble"], button_hover_color=t["accent"], text_color=t["text"])
        for btn in [self.skills_btn, self.memory_btn, self.attach_btn, self.voice_btn]:
            btn.configure(fg_color=t["user_bubble"], hover_color=t["accent"], text_color=t["text"])
        self.delete_chat_btn.configure(border_color=t["border"], text_color=t["muted"], hover_color="#991b1b")
        self.stop_tts_btn.configure(border_color=t["border"], text_color=t["muted"], hover_color=t["accent"])

        self.render_sidebar_chats()
        if self.current_chat_id:
            self.render_messages_list()


    def toggle_wake_listener(self):
        if self.wake_var.get():
            threading.Thread(target=self._wake_word_loop, daemon=True).start()
        else:
            self.wake_switch.configure(text="👂 Отклик:")

    def _wake_word_loop(self):
        fs = 16000
        block_size = 1600  # 0.1 секунды
        recognizer = sr.Recognizer()
        pre_buffer = []  # Хранит 0.4 сек до начала слова, чтобы не обрезать первую букву
        speech_chunks = []
        silence_blocks = 0
        is_speaking = False
        ui_tick = 0

        try:
            with sd.InputStream(samplerate=fs, channels=1, dtype="int16", blocksize=block_size) as stream:
                while self.wake_var.get():
                    if self.is_recording or self.is_generating:
                        speech_chunks.clear()
                        pre_buffer.clear()
                        is_speaking = False
                        time.sleep(0.2)
                        continue

                    data, _ = stream.read(block_size)
                    volume = int(np.abs(data).mean())

                    # Показываем текущую громкость микрофона прямо на переключателе
                    ui_tick += 1
                    if ui_tick % 3 == 0:
                        self.after(0, lambda v=volume: self.wake_switch.configure(text=f"👂 ({v}):"))

                    if volume > 120:  # Сниженный порог чувствительности
                        if not is_speaking:
                            is_speaking = True
                            speech_chunks.extend(pre_buffer)
                            pre_buffer.clear()
                        silence_blocks = 0
                        speech_chunks.append(data.copy())
                    elif is_speaking:
                        speech_chunks.append(data.copy())
                        silence_blocks += 1

                        # Пауза 0.6 сек после слова
                        if silence_blocks >= 6:
                            audio_np = np.concatenate(speech_chunks, axis=0)
                            speech_chunks.clear()
                            is_speaking = False
                            silence_blocks = 0

                            if len(audio_np) > fs * 0.4:
                                threading.Thread(
                                    target=self._check_wake_phrase,
                                    args=(audio_np, fs, recognizer),
                                    daemon=True,
                                ).start()
                    else:
                        pre_buffer.append(data.copy())
                        if len(pre_buffer) > 4:
                            pre_buffer.pop(0)
        except Exception as e:
            self.after(0, lambda err=e: self.add_message_card("system", f"Ошибка фонового микрофона: {err}"))
        finally:
            self.after(0, lambda: self.wake_switch.configure(text="👂 Отклик:"))

    def _check_wake_phrase(self, audio_np, fs, recognizer):
        try:
            # Добавляем 0.3 сек тишины по краям для точного распознавания
            pad = np.zeros((int(fs * 0.3), 1), dtype="int16")
            padded_audio = np.concatenate([pad, audio_np, pad], axis=0)

            buf = io.BytesIO()
            wav.write(buf, fs, padded_audio)
            buf.seek(0)
            with sr.AudioFile(buf) as source:
                audio = recognizer.record(source)
            heard = recognizer.recognize_google(audio, language="ru-RU").lower()
            trigger = self.wake_word_entry.get().strip().lower()
            if trigger and trigger in heard and not self.is_recording:
                self.after(0, self._activate_by_wake_word)
        except Exception:
            pass
    def on_tts_toggle(self):
        if not self.tts_var.get():
            self.stop_speaking()

    def stop_speaking(self):
        self.tts_session_id = getattr(self, "tts_session_id", 0) + 1
        try:
            import ctypes
            ctypes.windll.winmm.mciSendStringW("stop tts_snd", None, 0, None)
            ctypes.windll.winmm.mciSendStringW("close tts_snd", None, 0, None)
        except Exception:
            pass

    def speak_text(self, text):
        if not self.tts_var.get():
            return
        clean = re.sub(r"```[\s\(\S\)]*?```", " Код приведён в чате. ", text)
        clean = re.sub(r"[*#`~_✦⚙️🖼🎬📄📝]", "", clean).strip()
        if not clean:
            return

        self.stop_speaking()
        session_id = getattr(self, "tts_session_id", 0)

        def _tts_worker():
            mp3_path = os.path.join(DATA_DIR, "tts_reply.mp3")
            try:
                import ctypes

                async def _generate():
                    communicate = edge_tts.Communicate(clean[:1500], "ru-RU-DmitryNeural", rate="+10%")
                    await communicate.save(mp3_path)

                asyncio.run(_generate())

                if not self.tts_var.get() or getattr(self, "tts_session_id", 0) != session_id:
                    return

                ctypes.windll.winmm.mciSendStringW("close tts_snd", None, 0, None)
                ctypes.windll.winmm.mciSendStringW(f'open "{mp3_path}" type mpegvideo alias tts_snd', None, 0, None)
                ctypes.windll.winmm.mciSendStringW("play tts_snd", None, 0, None)
            except Exception as e:
                print(f"Ошибка озвучки: {e}")

        threading.Thread(target=_tts_worker, daemon=True).start()


    def _ensure_soft_wake_sound(self, wav_path):
        """Создаёт мягкий двутональный звук с плавным затуханием, если файла ещё нет."""
        if os.path.exists(wav_path):
            return
        try:
            fs = 44100
            # Мягкий восходящий интервал: До5 (523 Гц) -> Соль5 (784 Гц)
            t1 = np.linspace(0, 0.11, int(fs * 0.11), endpoint=False)
            t2 = np.linspace(0, 0.22, int(fs * 0.22), endpoint=False)

            env1 = np.sin(np.pi * t1 / 0.11) ** 1.5
            env2 = np.exp(-t2 * 14) * (1 - np.exp(-t2 * 120))

            tone1 = (0.7 * np.sin(2 * np.pi * 523.25 * t1) + 0.3 * np.sin(2 * np.pi * 1046.5 * t1)) * env1
            tone2 = (0.7 * np.sin(2 * np.pi * 783.99 * t2) + 0.3 * np.sin(2 * np.pi * 1567.98 * t2)) * env2

            signal = np.concatenate([tone1, tone2]) * 0.28
            wav.write(wav_path, fs, np.int16(signal * 32767))
        except Exception:
            pass

    def play_wake_sound(self):
        wav_path = os.path.join(DATA_DIR, "wake.wav")
        mp3_path = os.path.join(DATA_DIR, "wake.mp3")
        if os.path.exists(mp3_path):
            def _play_mp3():
                import ctypes
                ctypes.windll.winmm.mciSendStringW("close wake_snd", None, 0, None)
                ctypes.windll.winmm.mciSendStringW(f'open "{mp3_path}" type mpegvideo alias wake_snd', None, 0, None)
                ctypes.windll.winmm.mciSendStringW("play wake_snd", None, 0, None)
            threading.Thread(target=_play_mp3, daemon=True).start()
        else:
            self._ensure_soft_wake_sound(wav_path)
            if os.path.exists(wav_path):
                winsound.PlaySound(wav_path, winsound.SND_FILENAME | winsound.SND_ASYNC)


    def _activate_by_wake_word(self):
        self.stop_speaking()
        self.deiconify()
        self.state("normal")
        self.lift()
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.focus_force()
        self.play_wake_sound()
        if not self.is_recording:
            self.start_voice_recording(auto_send=True)

    def toggle_voice_recording(self):
        if self.is_recording:
            self.is_recording = False
            self.voice_btn.configure(text="⏳...")
            return
        self.stop_speaking()

        self.start_voice_recording(auto_send=False)

    def start_voice_recording(self, auto_send=False):
        self.is_recording = True
        self.voice_btn.configure(text="⏹ Стоп", fg_color="#dc2626", hover_color="#b91c1c")
        threading.Thread(
            target=self._record_and_transcribe, args=(auto_send,), daemon=True
        ).start()

    def _record_and_transcribe(self, auto_send=False):
        fs = 16000
        block_size = 1600  # 0.1 секунды
        chunks = []
        silence_blocks = 0
        has_spoken = False

        try:
            with sd.InputStream(samplerate=fs, channels=1, dtype="int16", blocksize=block_size) as stream:
                while self.is_recording:
                    data, _ = stream.read(block_size)
                    chunks.append(data.copy())
                    volume = np.abs(data).mean()

                    if volume > 120:
                        has_spoken = True
                        silence_blocks = 0
                    elif auto_send:
                        silence_blocks += 1
                        # Если человек закончил говорить и молчит 1.5 сек (или молчит 4 сек с начала)
                        if (has_spoken and silence_blocks >= 15) or (not has_spoken and silence_blocks >= 40):
                            self.is_recording = False
                            break

            if not chunks or (auto_send and not has_spoken):
                return

            self.after(0, lambda: self.voice_btn.configure(text="⏳..."))
            audio_np = np.concatenate(chunks, axis=0)
            buf = io.BytesIO()
            wav.write(buf, fs, audio_np)
            buf.seek(0)

            recognizer = sr.Recognizer()
            with sr.AudioFile(buf) as source:
                audio = recognizer.record(source)
            text = recognizer.recognize_google(audio, language="ru-RU")
            self.after(0, lambda t=text: self._insert_voice_text(t, auto_send))
        except Exception as e:
            if not auto_send:
                self.after(0, lambda err=e: self.add_message_card("system", f"Ошибка микрофона: {err}"))
        finally:
            self.is_recording = False
            t = THEMES[self.current_theme_name]
            self.after(0, lambda: self.voice_btn.configure(text="🎙 Голос", fg_color=t["user_bubble"], hover_color=t["accent"]))

    def _insert_voice_text(self, text, auto_send=False):
        current = self.input_textbox.get("1.0", "end").strip()
        new_text = f"{current} {text}".strip() if current else text
        self.input_textbox.delete("1.0", "end")
        self.input_textbox.insert("1.0", new_text)
        if auto_send and new_text:
            self.send_message()


    def open_skills_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Скиллы и плагины")
        win.geometry("600x520")
        win.grab_set()
        t = THEMES[self.current_theme_name]
        win.configure(fg_color=t["bg_main"])
        ctk.CTkLabel(win, text="Подключённые скиллы и плагины", font=("Segoe UI", 16, "bold"), text_color=t["text"]).pack(padx=15, pady=(15, 5), anchor="w")
        scroll = ctk.CTkScrollableFrame(win, fg_color=t["bg_card"])
        scroll.pack(fill="both", expand=True, padx=15, pady=8)

        def render_list():
            for w in scroll.winfo_children():
                w.destroy()
            for name, skill in self.skill_manager.skills.items():
                row = ctk.CTkFrame(scroll, fg_color=t["bg_main"], corner_radius=10)
                row.pack(fill="x", padx=6, pady=5)
                var = ctk.BooleanVar(value=self.skill_manager.enabled_skills.get(name, False))
                ctk.CTkCheckBox(row, text=skill["title"], variable=var, font=("Segoe UI", 13, "bold"), text_color=t["text"], fg_color=t["accent"], command=lambda n=name, v=var: self.skill_manager.enabled_skills.update({n: v.get()})).pack(anchor="w", padx=12, pady=(8, 2))
                ctk.CTkLabel(row, text=skill["description"], font=("Segoe UI", 12), text_color=t["muted"], wraplength=500, justify="left").pack(anchor="w", padx=36, pady=(0, 8))

        render_list()
        bf = ctk.CTkFrame(win, fg_color="transparent")
        bf.pack(fill="x", padx=15, pady=12)
        ctk.CTkButton(bf, text="↻ Обновить плагины", fg_color=t["accent"], hover_color=t["accent_hover"], command=lambda: (self.skill_manager.reload_all(), render_list())).pack(side="left")
        ctk.CTkButton(bf, text="📂 Папка plugins", fg_color=t["bg_card"], hover_color=t["accent"], text_color=t["text"], command=lambda: os.startfile(os.path.abspath(PLUGINS_DIR))).pack(side="right")

    def refresh_models(self):
        def fetch():
            try:
                resp = requests.get(f"{OLLAMA_BASE}/api/tags", timeout=5)
                if resp.status_code == 200:
                    models = [m["name"] for m in resp.json().get("models", [])]
                    if models:
                        self.after(0, lambda: self.update_model_dropdown(models))
            except Exception:
                pass
        threading.Thread(target=fetch, daemon=True).start()

    def update_model_dropdown(self, models):
        self.model_menu.configure(values=models)
        if self.model_var.get() not in models:
            self.model_var.set(models[0])

    def load_memory(self):
        if os.path.exists(MEMORY_FILE):
            try:
                with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return []
        return ["Отвечай всегда на русском языке.", "Используй подключённые скиллы, когда это нужно для задачи."]

    def save_memory(self, facts):
        self.memory_facts = facts
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(self.memory_facts, f, ensure_ascii=False, indent=2)

    def open_memory_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Память ассистента")
        win.geometry("520x420")
        win.grab_set()
        t = THEMES[self.current_theme_name]
        win.configure(fg_color=t["bg_main"])
        ctk.CTkLabel(win, text="Факты и правила (по одному на строку):", font=("Segoe UI", 13), text_color=t["text"]).pack(padx=15, pady=(15, 5), anchor="w")
        tb = ctk.CTkTextbox(win, font=("Segoe UI", 13), fg_color=t["bg_card"], text_color=t["text"])
        tb.pack(fill="both", expand=True, padx=15, pady=10)
        tb.insert("1.0", "\n".join(self.memory_facts))
        ctk.CTkButton(win, text="Сохранить память", height=38, fg_color=t["accent"], hover_color=t["accent_hover"], command=lambda: (self.save_memory([l.strip() for l in tb.get("1.0", "end").splitlines() if l.strip()]), win.destroy())).pack(padx=15, pady=(0, 15), fill="x")
    def load_all_chats(self):
        self.chats = {}
        for filename in os.listdir(CHATS_DIR):
            if filename.endswith(".json"):
                try:
                    with open(os.path.join(CHATS_DIR, filename), "r", encoding="utf-8") as f:
                        data = json.load(f)
                        self.chats[data["id"]] = data
                except Exception:
                    pass
        self.render_sidebar_chats()

    def save_chat(self, chat_id):
        if chat_id in self.chats:
            self.chats[chat_id]["updated_at"] = time.time()
            with open(os.path.join(CHATS_DIR, f"{chat_id}.json"), "w", encoding="utf-8") as f:
                json.dump(self.chats[chat_id], f, ensure_ascii=False, indent=2)
            self.render_sidebar_chats()

    def create_new_chat(self):
        chat_id = str(int(time.time() * 1000))
        self.chats[chat_id] = {"id": chat_id, "title": "Новый чат", "updated_at": time.time(), "messages": []}
        self.save_chat(chat_id)
        self.select_chat(chat_id)

    def delete_current_chat(self):
        if not self.current_chat_id:
            return
        path = os.path.join(CHATS_DIR, f"{self.current_chat_id}.json")
        if os.path.exists(path):
            os.remove(path)
        del self.chats[self.current_chat_id]
        self.current_chat_id = None
        if not self.chats:
            self.create_new_chat()
        else:
            next_id = sorted(self.chats.keys(), key=lambda k: self.chats[k].get("updated_at", 0), reverse=True)
            self.select_chat(next_id[0])

    def render_sidebar_chats(self):
        for w in self.chat_list_frame.winfo_children():
            w.destroy()
        t = THEMES[self.current_theme_name]
        for chat in sorted(self.chats.values(), key=lambda c: c.get("updated_at", 0), reverse=True):
            is_active = chat["id"] == self.current_chat_id
            ctk.CTkButton(self.chat_list_frame, text=chat["title"], anchor="w", corner_radius=8, fg_color=t["user_bubble"] if is_active else "transparent", hover_color=t["bg_card"], text_color=t["text"], height=36, command=lambda cid=chat["id"]: self.select_chat(cid)).pack(fill="x", pady=2)

    def select_chat(self, chat_id):
        self.current_chat_id = chat_id
        self.render_sidebar_chats()
        self.clear_attachments()
        self.render_messages_list()

    def render_messages_list(self):
        for w in self.chat_scroll.winfo_children():
            w.destroy()
        chat = self.chats[self.current_chat_id]
        if not chat["messages"]:
            self.add_message_card("system", "✦ Добро пожаловать в AI Studio! Задайте вопрос, вставьте скриншот через Ctrl+V или прикрепите файл.")
        else:
            for msg in chat["messages"]:
                self.add_message_card(msg["role"], msg.get("display_content", msg["content"]))

    def add_message_card(self, role, text):
        t = THEMES[self.current_theme_name]
        outer = ctk.CTkFrame(self.chat_scroll, fg_color=t["bg_main"])
        outer.pack(fill="x", pady=10, padx=24)

        container = ctk.CTkFrame(outer, fg_color=t["bg_main"])
        container.pack(fill="x", padx=20)


        if role == "user":
            card = ctk.CTkFrame(container, fg_color=t["user_bubble"], corner_radius=18)
            card.pack(side="right", anchor="e", padx=(90, 0))
            lbl = ctk.CTkLabel(
                card, text=text, font=("Segoe UI", 14), text_color=t["text"],
                wraplength=580, justify="left"
            )
            lbl.pack(padx=16, pady=11)
        elif role == "assistant":
            # В стиле Claude: без тяжёлой рамки, текст прямо на фоне с мягким заголовком
            card = ctk.CTkFrame(container, fg_color="transparent")
            card.pack(side="left", anchor="w", fill="x", expand=True, padx=(0, 40))

            hdr = ctk.CTkFrame(card, fg_color="transparent")
            hdr.pack(fill="x", pady=(0, 4))

            ctk.CTkLabel(
                hdr, text="✦  AI Studio", font=("Segoe UI", 13, "bold"),
                text_color=t["accent"]
            ).pack(side="left")

            copy_btn = ctk.CTkButton(
                hdr, text="Копировать", width=78, height=24, corner_radius=10,
                font=("Segoe UI", 11), fg_color="transparent",
                hover_color=t["user_bubble"], text_color=t["muted"],
                command=lambda: (
                    self.clipboard_clear(),
                    self.clipboard_append(lbl.cget("text")),
                    copy_btn.configure(text="✓ Скопировано"),
                    self.after(1500, lambda: copy_btn.configure(text="Копировать"))
                )
            )
            copy_btn.pack(side="right")

            lbl = ctk.CTkLabel(
                card, text=text, font=("Segoe UI", 14), text_color=t["text"],
                wraplength=680, justify="left", anchor="w"
            )
            lbl.pack(fill="x", pady=(2, 6))
            self.active_ai_label = lbl
        else:
            welcome = ctk.CTkFrame(container, fg_color=t["bg_card"], corner_radius=16, border_width=1, border_color=t["border"])
            welcome.pack(pady=24, padx=40)
            lbl = ctk.CTkLabel(
                welcome, text=text, font=("Segoe UI", 13), text_color=t["muted"],
                wraplength=540, justify="center"
            )
            lbl.pack(padx=24, pady=16)

        self.chat_scroll._parent_canvas.yview_moveto(1.0)

    def attach_media(self):
        paths = filedialog.askopenfilenames(title="Выберите файлы", filetypes=[("Все поддерживаемые", "*.png *.jpg *.jpeg *.webp *.bmp *.mp4 *.avi *.mov *.mkv *.txt *.py *.md *.json *.csv *.pdf")])
        if paths:
            self.process_file_paths(paths)

    def process_file_paths(self, filepaths):
        for path in filepaths:
            if not os.path.isfile(path):
                continue
            ext = os.path.splitext(path)[1].lower()
            filename = os.path.basename(path)
            if ext in [".png", ".jpg", ".jpeg", ".webp", ".bmp"]:
                try:
                    with Image.open(path) as img:
                        img = img.convert("RGB")
                        img.thumbnail((1024, 1024))
                        buf = io.BytesIO()
                        img.save(buf, format="JPEG", quality=85)
                        self.attached_images.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
                        self.attached_labels.append(f"🖼 {filename}")
                except Exception:
                    pass
            elif ext in [".mp4", ".avi", ".mov", ".mkv"]:
                frames = self.extract_video_frames(path, 5)
                if frames:
                    self.attached_images.extend(frames)
                    self.attached_labels.append(f"🎬 {filename} ({len(frames)} к.)")
            elif ext == ".pdf":
                try:
                    reader = PdfReader(path)
                    txt = "\n".join([p.extract_text() or "" for p in reader.pages[:15]])[:15000]
                    self.attached_files_text.append((filename, txt))
                    self.attached_labels.append(f"📄 {filename}")
                except Exception:
                    pass
            elif ext == ".docx":
                try:
                    doc = docx.Document(path)
                    txt = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])[:15000]
                    self.attached_files_text.append((filename, txt))
                    self.attached_labels.append(f"📘 {filename}")
                except Exception:
                    pass
            elif ext == ".xlsx":
                try:
                    wb = openpyxl.load_workbook(path, data_only=True)
                    ws = wb.active
                    rows = [" | ".join("" if v is None else str(v) for v in r) for r in ws.iter_rows(values_only=True)]
                    self.attached_files_text.append((filename, "\n".join(rows)[:15000]))
                    self.attached_labels.append(f"📊 {filename}")
                except Exception:
                    pass

            else:
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        self.attached_files_text.append((filename, f.read()[:15000]))
                    self.attached_labels.append(f"📝 {filename}")
                except Exception:
                    pass
        self.update_attachments_bar()

    def extract_video_frames(self, video_path, max_frames=5):
        frames_b64 = []
        try:
            cap = cv2.VideoCapture(video_path)
            total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            step = max(1, total // max_frames)
            for i in range(0, total, step):
                if len(frames_b64) >= max_frames:
                    break
                cap.set(cv2.CAP_PROP_POS_FRAMES, i)
                ret, frame = cap.read()
                if ret:
                    pil_img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                    pil_img.thumbnail((800, 800))
                    buf = io.BytesIO()
                    pil_img.save(buf, format="JPEG", quality=80)
                    frames_b64.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
            cap.release()
        except Exception:
            pass
        return frames_b64

    def update_attachments_bar(self):
        for w in self.attachments_frame.winfo_children():
            w.destroy()
        if not self.attached_labels:
            self.attachments_frame.grid_forget()
            return
        t = THEMES[self.current_theme_name]
        self.attachments_frame.grid(row=0, column=0, padx=14, pady=(8, 0), sticky="ew")
        ctk.CTkLabel(self.attachments_frame, text="Прикреплено: " + ", ".join(self.attached_labels), font=("Segoe UI", 12), text_color=t["accent"]).pack(side="left")
        ctk.CTkButton(self.attachments_frame, text="✕", width=26, height=22, fg_color=t["user_bubble"], command=self.clear_attachments).pack(side="right")

    def clear_attachments(self):
        self.attached_images, self.attached_files_text, self.attached_labels = [], [], []
        self.update_attachments_bar()

    def send_message(self):
        if self.is_generating:
            return
        self.stop_speaking()

        user_text = self.input_textbox.get("1.0", "end").strip()
        if not user_text and not self.attached_labels:
            return
        if not user_text:
            user_text = "Проанализируй прикреплённые материалы."

        chat = self.chats[self.current_chat_id]
        display_msg = f"[{', '.join(self.attached_labels)}]\n{user_text}" if self.attached_labels else user_text
        full_prompt = user_text
        if self.attached_files_text:
            files_block = "\n\n".join([f"--- Файл {fn} ---\n{ft}" for fn, ft in self.attached_files_text])
            full_prompt = f"{user_text}\n\n{files_block}"

        if not chat["messages"]:
            chat["title"] = user_text[:28] + ("..." if len(user_text) > 28 else "")

        msg_obj = {"role": "user", "content": full_prompt, "display_content": display_msg}
        if self.attached_images:
            msg_obj["images"] = list(self.attached_images)

        chat["messages"].append(msg_obj)
        self.save_chat(self.current_chat_id)
        self.input_textbox.delete("1.0", "end")
        self.clear_attachments()
        self.add_message_card("user", display_msg)
        self.add_message_card("assistant", "Думаю...")

        self.is_generating = True
        self.send_btn.configure(state="disabled")
        threading.Thread(target=self.process_ai_turn, daemon=True).start()

    def process_ai_turn(self):
        chat = self.chats[self.current_chat_id]
        model_name = self.model_var.get()
        memory_text = "\n".join([f"- {fact}" for fact in self.memory_facts])
        user_home = os.path.expanduser("~")
        drives = [f"{d}:\\" for d in "CDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{d}:\\")]
        sys_prompt = (
            f"Ты локальный AI-агент с полным доступом к файловой системе ПК через инструмент manage_files.\n"
            f"Домашняя папка пользователя: {user_home}\n"
            f"Рабочий стол: {os.path.join(user_home, 'Desktop')}\n"
            f"Документы: {os.path.join(user_home, 'Documents')}\n"
            f"Загрузки: {os.path.join(user_home, 'Downloads')}\n"
            f"Доступные диски: {', '.join(drives)}\n"
            f"Ты можешь свободно создавать, читать, искать и редактировать файлы в любой папке по запросу пользователя.\n"
            f"Твоя память:\n{memory_text}"
        )
        api_messages = [{"role": "system", "content": sys_prompt}]

        for m in chat["messages"]:
            entry = {"role": m["role"], "content": m["content"]}
            if "images" in m:
                entry["images"] = m["images"]
            api_messages.append(entry)

        tools = self.skill_manager.get_ollama_tools()
        try:
            for _ in range(4):
                payload = {"model": model_name, "messages": api_messages, "stream": False, "think": False}
                if tools:
                    payload["tools"] = tools
                resp = requests.post(f"{OLLAMA_BASE}/api/chat", json=payload, timeout=180)
                resp.raise_for_status()
                msg = resp.json().get("message", {})
                tool_calls = msg.get("tool_calls")
                if tool_calls:
                    api_messages.append(msg)
                    for tc in tool_calls:
                        fn = tc.get("function", {})
                        sn, sa = fn.get("name", ""), fn.get("arguments", {})
                        self.after(0, lambda s=sn: self.active_ai_label.configure(text=f"⚙️ Выполняю скилл: {s}..."))
                        res = self.skill_manager.execute(sn, sa)
                        api_messages.append({"role": "tool", "content": res})
                else:
                    reply = msg.get("content", "")
                    self.after(0, lambda r=reply: self.active_ai_label.configure(text=r))
                    chat["messages"].append({"role": "assistant", "content": reply})
                    if self.tts_var.get():
                        self.speak_text(reply)
                    self.after(0, lambda: self.save_chat(self.current_chat_id))
                    break
        except Exception as e:
            self.after(0, lambda err=e: self.active_ai_label.configure(text=f"Ошибка: {err}"))
        finally:
            self.is_generating = False
            self.after(0, lambda: self.send_btn.configure(state="normal"))


if __name__ == "__main__":
    app = AIStudioApp()
    app.mainloop()