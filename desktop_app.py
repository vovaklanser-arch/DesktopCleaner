"""
Desktop Cleaner — standalone приложение
Портативное: работает из любой папки
"""
import json
import os
import hashlib
import subprocess
import threading
from pathlib import Path
import customtkinter as ctk
import psutil

# Все пути относительно расположения скрипта — работает при переносе
APP_DIR = Path(__file__).resolve().parent
os.chdir(APP_DIR)
SETTINGS_FILE = APP_DIR / "settings.json"

# ============ ПРЕМИУМ ДИЗАЙН ============
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

COLORS = {
    "bg": "#09090b",
    "card": "#18181b",
    "card_hover": "#2a2a35",
    "sidebar": "#131316",
    "accent": "#6366f1",
    "accent_hover": "#4f46e5",
    "success": "#10b981",
    "danger": "#f43f5e",
    "warning": "#f59e0b",
    "text": "#f8fafc",
    "text_dim": "#94a3b8",
    "border": "#27272a"
}

# ============ ПУТИ ============
USERPROFILE = Path(os.environ.get("USERPROFILE", ""))
DESKTOP = USERPROFILE / "Desktop"
PROGRAM_FOLDER = "Program"
SKIP_EXTENSIONS = {".lnk", ".tmp", ".temp", ".cache", ".log"}
MIN_FILE_SIZE = 1024

PROTECTED_PROCESSES = {
    "system", "idle", "csrss", "smss", "wininit", "services", "lsass",
    "svchost", "explorer", "dwm", "fontdrvhost", "sihost", "taskmgr",
    "searchapp", "startmenuexperiencehost", "runtimebroker", "ctfmon",
    "conhost", "python", "pythonw", "cmd", "powershell", "code", "cursor",
}
# Оставляем при закрытии фоновых
KEEP_PROCESSES = {"chrome", "cursor", "explorer", "python", "pythonw"}

CLEANUP_COMMANDS = {
    "temp": ("Очистить TEMP", "Remove-Item -Path $env:TEMP\\* -Recurse -Force -ErrorAction SilentlyContinue; Write-Host 'TEMP очищен'"),
    "prefetch": ("Очистить Prefetch", "Remove-Item -Path C:\\Windows\\Prefetch\\* -Force -ErrorAction SilentlyContinue; Write-Host 'Prefetch очищен'"),
    "recycle": ("Очистить корзину", "Clear-RecycleBin -Force -ErrorAction SilentlyContinue; Write-Host 'Корзина очищена'"),
    "browser_cache": ("Кэш Chrome", "Remove-Item -Path \"$env:LOCALAPPDATA\\Google\\Chrome\\User Data\\Default\\Cache\\*\" -Recurse -Force -ErrorAction SilentlyContinue; Write-Host 'Кэш Chrome очищен'"),
    "thumbnails": ("Миниатюры", "Remove-Item -Path \"$env:LOCALAPPDATA\\Microsoft\\Windows\\Explorer\\thumbcache_*.db\" -Force -ErrorAction SilentlyContinue; Write-Host 'Миниатюры очищены'"),
}


def get_desktop_path():
    return DESKTOP


def get_program_path():
    return get_desktop_path() / PROGRAM_FOLDER


def load_settings():
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def file_hash(path, size_limit=10 * 1024 * 1024):
    path = Path(path)
    if not path.is_file() or path.suffix.lower() in SKIP_EXTENSIONS:
        return None
    size = path.stat().st_size
    if size < MIN_FILE_SIZE:
        return None
    if size > size_limit:
        return f"size:{size}"
    try:
        with open(path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except (OSError, PermissionError):
        return None


# ============ ГЛАВНОЕ ОКНО ============
class DesktopCleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Desktop Cleaner")
        self.geometry("1150x750")
        self.minsize(950, 650)
        self.configure(fg_color=COLORS["bg"])
        self.after(100, self._center_window)

        # Верхний заголовок и статистика
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=28, pady=(28, 16))
        
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(title_box, text="⚡", font=ctk.CTkFont(size=32), text_color=COLORS["accent"]).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(title_box, text="Desktop Cleaner", font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"), text_color=COLORS["text"]).pack(side="left")
        
        stats_box = ctk.CTkFrame(header, fg_color=COLORS["sidebar"], corner_radius=12, border_width=1, border_color=COLORS["border"])
        stats_box.pack(side="right", fill="y")
        self.stats_label = ctk.CTkLabel(stats_box, text="RAM: — | Диск: —", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color=COLORS["text_dim"])
        self.stats_label.pack(padx=24, pady=10)

        # Контейнер: левый Sidebar + правая часть
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=24, pady=(0, 24))

        # Левый Sidebar
        self.sidebar_frame = ctk.CTkFrame(main_container, width=240, fg_color=COLORS["sidebar"], corner_radius=16, border_width=1, border_color=COLORS["border"])
        self.sidebar_frame.pack(side="left", fill="y", padx=(0, 24))
        self.sidebar_frame.pack_propagate(False)

        # Контентная часть справа
        self.content_container = ctk.CTkFrame(main_container, fg_color="transparent")
        self.content_container.pack(side="right", fill="both", expand=True)

        self.frames = {}
        self.nav_buttons = {}

        # Пути
        paths_frame = ctk.CTkFrame(self.content_container, fg_color=COLORS["sidebar"], corner_radius=16, border_width=1, border_color=COLORS["border"])
        paths_frame.pack(fill="x", pady=(0, 20))
        paths_inner = ctk.CTkFrame(paths_frame, fg_color="transparent")
        paths_inner.pack(fill="x", padx=20, pady=14)
        ctk.CTkLabel(paths_inner, text="Рабочий стол:", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLORS["text_dim"]).pack(side="left")
        self.path_desktop = ctk.CTkLabel(paths_inner, text=str(get_desktop_path()), font=ctk.CTkFont(size=13), text_color=COLORS["text"])
        self.path_desktop.pack(side="left", padx=(8, 24))
        ctk.CTkLabel(paths_inner, text="Program:", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLORS["text_dim"]).pack(side="left")
        self.path_program = ctk.CTkLabel(paths_inner, text=str(get_program_path()), font=ctk.CTkFont(size=13), text_color=COLORS["accent"])
        self.path_program.pack(side="left", padx=(8, 0))

        # Фрейм для активного таба
        self.active_tab_frame = ctk.CTkFrame(self.content_container, fg_color=COLORS["sidebar"], corner_radius=16, border_width=1, border_color=COLORS["border"])
        self.active_tab_frame.pack(fill="both", expand=True)

        # Настройка табов
        self._add_nav_item("⚙️", "Процессы", self._build_processes_tab)
        self._add_nav_item("🛡️", "Фоновые", self._build_background_tab)
        self._add_nav_item("🗂️", "Дубликаты", self._build_duplicates_tab)
        self._add_nav_item("🖥️", "Рабочий стол", self._build_desktop_tab)
        self._add_nav_item("🧹", "Очистка", self._build_cleanup_tab)

        # Кнопка остановить фоновые внизу сайдбара
        kill_btn = ctk.CTkButton(self.sidebar_frame, text="Остановить фоновые", command=self._kill_all_background, fg_color=COLORS["danger"], hover_color="#dc2626", font=ctk.CTkFont(size=14, weight="bold"))
        kill_btn.pack(side="bottom", fill="x", padx=20, pady=24)

        self.select_frame("Процессы")

        self._load_saved_settings()
        self._refresh_stats()
        self.after(15000, self._stats_loop)

    def _add_nav_item(self, icon, name, build_func):
        btn = ctk.CTkButton(
            self.sidebar_frame, 
            text=f"   {icon}   {name}", 
            anchor="w", 
            fg_color="transparent",
            text_color=COLORS["text_dim"], 
            hover_color=COLORS["card_hover"],
            font=ctk.CTkFont(size=15, weight="bold"),
            height=46,
            corner_radius=8,
            command=lambda: self.select_frame(name)
        )
        btn.pack(fill="x", padx=16, pady=4)
        if name == "Процессы": 
            btn.pack(pady=(24, 4))
        self.nav_buttons[name] = btn

        frame = ctk.CTkFrame(self.active_tab_frame, fg_color="transparent")
        self.frames[name] = frame
        build_func(frame)

    def select_frame(self, name):
        for btn_name, btn in self.nav_buttons.items():
            if btn_name == name:
                btn.configure(fg_color=COLORS["card_hover"], text_color=COLORS["text"])
            else:
                btn.configure(fg_color="transparent", text_color=COLORS["text_dim"])
        for frame_name, frame in self.frames.items():
            if frame_name == name:
                frame.pack(fill="both", expand=True, padx=28, pady=28)
            else:
                frame.pack_forget()

    def _center_window(self):
        self.update_idletasks()
        w, h = 1150, 750
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _refresh_stats(self):
        try:
            mem = psutil.virtual_memory()
            disk = psutil.disk_usage("C:\\")
            self.stats_label.configure(text=f"RAM: {mem.used / (1024**3):.1f} GB ({mem.percent}%) | Диск C: {disk.used / (1024**3):.1f} GB")
        except Exception:
            self.stats_label.configure(text="Ошибка загрузки статистики")

    def _stats_loop(self):
        self._refresh_stats()
        self.after(15000, self._stats_loop)

    def _load_saved_settings(self):
        s = load_settings()
        if s.get("last_dup_folder") and hasattr(self, "dup_folder"):
            self.dup_folder.insert(0, s["last_dup_folder"])

    def _build_processes_tab(self, tab):
        pass
        search_frame = ctk.CTkFrame(tab, fg_color="transparent")
        search_frame.pack(fill="x", pady=(0, 8))
        self.process_search = ctk.CTkEntry(search_frame, placeholder_text="Поиск по имени или PID...", width=400)
        self.process_search.pack(side="left", padx=(0, 8))
        self.process_search.bind("<KeyRelease>", lambda e: self._filter_processes())
        btn_refresh = ctk.CTkButton(search_frame, text="Обновить", command=self._load_processes, width=100)
        btn_refresh.pack(side="left")

        self.process_frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self.process_frame.pack(fill="both", expand=True)
        self._load_processes()

    def _load_processes(self):
        for w in self.process_frame.winfo_children():
            w.destroy()
        try:
            procs = []
            for p in psutil.process_iter(["pid", "name", "memory_percent", "cpu_percent"]):
                try:
                    info = p.info
                    name = (info.get("name") or "?").lower()
                    procs.append({
                        "pid": info.get("pid"),
                        "name": info.get("name", "?"),
                        "memory": round(info.get("memory_percent") or 0, 1),
                        "cpu": round(info.get("cpu_percent") or 0, 1),
                        "protected": any(prot in name for prot in PROTECTED_PROCESSES),
                    })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            procs.sort(key=lambda x: (x["memory"], x["cpu"]), reverse=True)
            self._processes_data = procs
            self._filter_processes()
        except Exception as e:
            ctk.CTkLabel(self.process_frame, text=f"Ошибка: {e}", text_color=COLORS["danger"]).pack(anchor="w")

    def _filter_processes(self):
        for w in self.process_frame.winfo_children():
            w.destroy()
        search = (self.process_search.get() or "").lower()
        procs = [p for p in getattr(self, "_processes_data", []) if search in (p["name"] or "").lower() or search in str(p["pid"])]
        for p in procs[:200]:
            row = ctk.CTkFrame(self.process_frame, fg_color=COLORS["sidebar"], corner_radius=8, height=44)
            row.pack(fill="x", pady=2)
            row.pack_propagate(False)
            ctk.CTkLabel(row, text=str(p["pid"]), width=70, font=ctk.CTkFont(family="Consolas"), text_color=COLORS["text_dim"]).pack(side="left", padx=12, pady=8)
            ctk.CTkLabel(row, text=(p["name"] or "?")[:50], width=250, anchor="w").pack(side="left", padx=4, pady=8)
            ctk.CTkLabel(row, text=f"{p['memory']}%", width=60, text_color=COLORS["warning"]).pack(side="left", padx=4, pady=8)
            ctk.CTkLabel(row, text=f"{p['cpu']}%", width=50, text_color=COLORS["accent"]).pack(side="left", padx=4, pady=8)
            btn = ctk.CTkButton(row, text="Завершить", width=90, fg_color=COLORS["danger"], hover_color="#dc2626", command=lambda pid=p["pid"], prot=p["protected"]: self._kill_process(pid, prot))
            if p["protected"]:
                btn.configure(state="disabled")
            btn.pack(side="right", padx=12, pady=6)

    def _kill_process(self, pid, protected):
        if protected:
            return
        if not self._confirm("Завершить процесс " + str(pid) + "?"):
            return
        try:
            p = psutil.Process(pid)
            p.terminate()
            p.wait(timeout=5)
            self._toast("Процесс завершён")
        except psutil.TimeoutExpired:
            try:
                p.kill()
            except Exception:
                pass
            self._toast("Процесс принудительно завершён")
        except Exception as e:
            self._toast(str(e), error=True)
        self._load_processes()

    def _build_background_tab(self, tab):
        pass
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(top, text="Закрыть фоновые приложения (кроме Chrome, Cursor, Explorer)", font=ctk.CTkFont(size=12), text_color=COLORS["text_dim"]).pack(anchor="w")
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.pack(fill="x", pady=8)
        ctk.CTkButton(btn_frame, text="Обновить список", command=self._load_background, width=140).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Закрыть все фоновые", command=self._kill_all_background, fg_color=COLORS["danger"], width=180).pack(side="left")
        self.bg_frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self.bg_frame.pack(fill="both", expand=True)
        self._load_background()

    def _load_background(self):
        for w in self.bg_frame.winfo_children():
            w.destroy()
        try:
            procs = []
            for p in psutil.process_iter(["pid", "name", "memory_percent"]):
                try:
                    info = p.info
                    name = (info.get("name") or "?").lower()
                    if any(prot in name for prot in PROTECTED_PROCESSES):
                        continue
                    if any(keep in name for keep in KEEP_PROCESSES):
                        continue
                    mem = round(info.get("memory_percent") or 0, 1)
                    if mem < 0.1:
                        continue
                    procs.append({"pid": info.get("pid"), "name": info.get("name", "?"), "memory": mem})
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            procs.sort(key=lambda x: x["memory"], reverse=True)
            self._bg_processes = procs
            for p in procs[:80]:
                row = ctk.CTkFrame(self.bg_frame, fg_color=COLORS["sidebar"], corner_radius=8, height=40)
                row.pack(fill="x", pady=2)
                row.pack_propagate(False)
                ctk.CTkLabel(row, text=(p["name"] or "?")[:45], width=280, anchor="w").pack(side="left", padx=12, pady=6)
                ctk.CTkLabel(row, text=f"{p['memory']}% RAM", width=70, text_color=COLORS["warning"]).pack(side="left", padx=4, pady=6)
                ctk.CTkButton(row, text="Закрыть", width=80, fg_color=COLORS["danger"], command=lambda pid=p["pid"]: self._kill_one_bg(pid)).pack(side="right", padx=8, pady=4)
        except Exception as e:
            ctk.CTkLabel(self.bg_frame, text=f"Ошибка: {e}", text_color=COLORS["danger"]).pack(anchor="w")

    def _kill_one_bg(self, pid):
        try:
            p = psutil.Process(pid)
            p.terminate()
            p.wait(timeout=3)
            self._toast("Процесс завершён")
        except Exception as e:
            self._toast(str(e), error=True)
        self._load_background()

    def _kill_all_background(self):
        if not hasattr(self, "_bg_processes") or not self._bg_processes:
            self._load_background()
        if not getattr(self, "_bg_processes", []):
            self._toast("Нет фоновых процессов для закрытия")
            return
        if not self._confirm("Закрыть ВСЕ фоновые приложения (кроме Chrome, Cursor, Explorer)? Несохранённые данные будут потеряны."):
            return
        killed = 0
        for p in self._bg_processes:
            try:
                proc = psutil.Process(p["pid"])
                proc.terminate()
                proc.wait(timeout=2)
                killed += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                try:
                    proc.kill()
                    killed += 1
                except Exception:
                    pass
            except Exception:
                pass
        self._toast(f"Закрыто: {killed} процессов")
        self._load_background()

    def _build_duplicates_tab(self, tab):
        pass
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))
        self.dup_folder = ctk.CTkEntry(top, placeholder_text="Путь к папке (пусто = рабочий стол)", width=450)
        self.dup_folder.pack(side="left", padx=(0, 8))
        ctk.CTkButton(top, text="Сканировать", command=self._scan_duplicates, width=120).pack(side="left")
        self.dup_frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self.dup_frame.pack(fill="both", expand=True)

    def _scan_duplicates(self):
        for w in self.dup_frame.winfo_children():
            w.destroy()
        folder = (self.dup_folder.get() or "").strip() or str(get_desktop_path())
        folder = Path(folder)
        if not folder.exists():
            ctk.CTkLabel(self.dup_frame, text=f"Папка не найдена: {folder}", text_color=COLORS["danger"]).pack(anchor="w")
            return
        if not folder.is_dir():
            ctk.CTkLabel(self.dup_frame, text=f"Указан файл, а не папка: {folder}", text_color=COLORS["danger"]).pack(anchor="w")
            return
        save_settings({"last_dup_folder": str(folder)})
        ctk.CTkLabel(self.dup_frame, text=f"Сканирование: {folder}", text_color=COLORS["text_dim"], font=ctk.CTkFont(size=11)).pack(anchor="w")
        self.dup_progress = ctk.CTkProgressBar(self.dup_frame, width=400, mode="indeterminate")
        self.dup_progress.pack(anchor="w", pady=8)
        self.dup_progress.start()
        self.update()

        def scan():
            try:
                by_hash = {}
                files = list(folder.rglob("*"))
                total = len([x for x in files if x.is_file()])
                for i, f in enumerate(files):
                    if f.is_file():
                        h = file_hash(f)
                        if h:
                            by_hash.setdefault(h, []).append(str(f))
                dups = [p for p in by_hash.values() if len(p) > 1]
                self.after(0, lambda: self._scan_done(dups))
            except Exception as e:
                self.after(0, lambda: self._scan_error(str(e)))

        threading.Thread(target=scan, daemon=True).start()

    def _scan_done(self, dups):
        if hasattr(self, "dup_progress") and self.dup_progress.winfo_exists():
            self.dup_progress.stop()
            self.dup_progress.destroy()
        for w in self.dup_frame.winfo_children():
            w.destroy()
        if not dups:
            ctk.CTkLabel(self.dup_frame, text="Дубликатов не найдено", text_color=COLORS["success"]).pack(anchor="w")
            return
        self._duplicates_data = dups
        for group in dups:
            card = ctk.CTkFrame(self.dup_frame, fg_color=COLORS["sidebar"], corner_radius=8)
            card.pack(fill="x", pady=4)
            for j, path in enumerate(group):
                color = COLORS["success"] if j == 0 else COLORS["text_dim"]
                ctk.CTkLabel(card, text=("✓ Оставить: " if j == 0 else "Удалить: ") + path[:80] + ("..." if len(path) > 80 else ""), font=ctk.CTkFont(size=11), text_color=color, anchor="w").pack(anchor="w", padx=12, pady=2)
            ctk.CTkButton(card, text=f"Удалить дубликаты ({len(group)-1} шт.)", fg_color=COLORS["danger"], command=lambda g=group: self._delete_duplicates(g)).pack(anchor="w", padx=12, pady=(8, 12))

    def _scan_error(self, msg):
        if hasattr(self, "dup_progress") and self.dup_progress.winfo_exists():
            self.dup_progress.stop()
            self.dup_progress.destroy()
        for w in self.dup_frame.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.dup_frame, text=f"Ошибка сканирования: {msg}", text_color=COLORS["danger"]).pack(anchor="w")

    def _delete_duplicates(self, group):
        if not self._confirm(f"Удалить {len(group)-1} дубликат(ов)? Восстановить будет невозможно."):
            return
        deleted, failed = 0, 0
        for p in group[1:]:
            try:
                Path(p).unlink()
                deleted += 1
            except PermissionError:
                failed += 1
            except Exception:
                failed += 1
        self._toast(f"Удалено: {deleted}" + (f", ошибок: {failed}" if failed else ""), error=bool(failed))
        self._scan_duplicates()

    def _build_desktop_tab(self, tab):
        pass
        self.desk_output = ctk.CTkTextbox(tab, height=80, fg_color=COLORS["sidebar"], corner_radius=8)
        self.desk_output.pack(fill="x", pady=(0, 12))

        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.pack(fill="x", pady=4)
        ctk.CTkButton(btn_frame, text="Создать папку Program", command=self._create_structure).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Перенести ярлыки", command=self._organize_shortcuts).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Показать лишние файлы", command=self._show_loose).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Перенести в Разное", command=self._organize_loose, fg_color=COLORS["accent"]).pack(side="left")

    def _create_structure(self):
        try:
            get_program_path().mkdir(parents=True, exist_ok=True)
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", f"Папка создана: {get_program_path()}")
            self._toast("Папка Program создана")
        except Exception as e:
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", f"Ошибка: {e}")

    def _organize_shortcuts(self):
        self.desk_output.delete("1.0", "end")
        self.desk_output.insert("1.0", "Выполняется...")
        self.update()
        try:
            root = get_program_path()
            root.mkdir(parents=True, exist_ok=True)
            desktop = get_desktop_path()
            moved = []
            exclude = ["chrome", "cursor", "program", "desktop cleaner"]
            for f in desktop.glob("*.lnk"):
                if not any(ex in f.stem.lower() for ex in exclude):
                    try:
                        (root / f.stem).mkdir(exist_ok=True)
                        dest = root / f.stem / f.name
                        if not dest.exists():
                            f.rename(dest)
                            moved.append(f.stem)
                    except Exception:
                        pass
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", f"Перенесено: {', '.join(moved)}" if moved else "Нечего переносить")
            if moved:
                self._toast(f"Перенесено {len(moved)} ярлыков")
        except Exception as e:
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", f"Ошибка: {e}")

    def _show_loose(self):
        root = get_program_path()
        if not root.exists():
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", "Папка Program не найдена")
            return
        files = sorted([f.name for f in root.iterdir() if f.is_file()])
        self.desk_output.delete("1.0", "end")
        self.desk_output.insert("1.0", f"Лишние файлы ({len(files)}):\n" + "\n".join(files) if files else "Лишних файлов нет")

    def _organize_loose(self):
        root = get_program_path()
        if not root.exists():
            self.desk_output.delete("1.0", "end")
            self.desk_output.insert("1.0", f"Папка Program не найдена: {root}\nСначала создайте её кнопкой «Создать папку Program».")
            return
        dest = root / "Разное"
        dest.mkdir(exist_ok=True)
        moved = []
        for f in root.iterdir():
            if f.is_file():
                try:
                    if not (dest / f.name).exists():
                        f.rename(dest / f.name)
                        moved.append(f.name)
                except Exception:
                    pass
        self.desk_output.delete("1.0", "end")
        self.desk_output.insert("1.0", f"Перенесено в Разное: {', '.join(moved)}" if moved else "Нечего переносить")
        if moved:
            self._toast(f"Перенесено {len(moved)} файлов")

    def _build_cleanup_tab(self, tab):
        pass
        self.cleanup_output = ctk.CTkTextbox(tab, height=120, fg_color=COLORS["sidebar"], corner_radius=8)
        self.cleanup_output.pack(fill="x", pady=(0, 12))
        for cmd_id, (name, _) in CLEANUP_COMMANDS.items():
            row = ctk.CTkFrame(tab, fg_color=COLORS["sidebar"], corner_radius=8, height=50)
            row.pack(fill="x", pady=4)
            row.pack_propagate(False)
            ctk.CTkLabel(row, text=name, font=ctk.CTkFont(size=14)).pack(side="left", padx=16, pady=12)
            ctk.CTkButton(row, text="Выполнить", width=100, command=lambda c=cmd_id: self._run_cleanup(c)).pack(side="right", padx=16, pady=8)

    def _run_cleanup(self, cmd_id):
        _, cmd = CLEANUP_COMMANDS[cmd_id]
        self.cleanup_output.delete("1.0", "end")
        self.cleanup_output.insert("1.0", "Выполняется...")
        self.update()

        def run():
            try:
                r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd], capture_output=True, text=True, timeout=120, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0)
                out = (r.stdout or r.stderr or "").strip() or "Готово"
                self.after(0, lambda: self._cleanup_done(out))
            except subprocess.TimeoutExpired:
                self.after(0, lambda: self._cleanup_done("Таймаут: операция заняла слишком много времени."))
            except Exception as e:
                self.after(0, lambda: self._cleanup_done(str(e)))

        threading.Thread(target=run, daemon=True).start()

    def _cleanup_done(self, text):
        self.cleanup_output.delete("1.0", "end")
        self.cleanup_output.insert("1.0", text)
        self._toast("Команда выполнена")

    def _confirm(self, msg):
        from tkinter import messagebox
        return messagebox.askyesno("Подтверждение", msg)

    def _toast(self, msg, error=False):
        top = ctk.CTkToplevel(self)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        lbl = ctk.CTkLabel(top, text=msg, fg_color=COLORS["danger"] if error else COLORS["success"], corner_radius=8, padx=16, pady=10, font=ctk.CTkFont(size=12))
        lbl.pack()
        self.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() - 320
        y = self.winfo_rooty() + self.winfo_height() - 80
        top.geometry(f"+{max(0, x)}+{max(0, y)}")
        top.after(2500, top.destroy)


if __name__ == "__main__":
    app = DesktopCleanerApp()
    app.mainloop()
