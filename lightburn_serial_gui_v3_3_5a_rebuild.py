# LightBurn Serial GUI (v3.7 — 5 jobs + pin-select + pulse toggles + fixes + updates)
# Implements:
# - 5 jobs (A..E); jobs hidden until a mapped pin goes HIGH (EXT1 GPIO 0,8,9,16,17)
# - When a mapped pin is HIGH, that job auto-selects and is the only button shown
# - Relay 0 and Relay 1 "Toggle" buttons now perform a pulse (off→on→off) using PULSE
# - Default Serial COM port set to COM5
# - Compatible with the ESP32-C6-EVB firmware you’re flashing
# - Fixes:
#   * START LASER is disabled while engraving is active
#   * LightBurn bring-to-front sequence mirrors Notepad test (robust foregrounding)
#   * "Settings saved" dialog appears in front of Settings (not behind)
#   * Renamed 4th relay to "Door Lock" (alias supports legacy "Stack Light" name)

import os, json, csv, sys, subprocess, platform, glob, shutil, time, re
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ---------- Windows helpers (ctypes) ----------
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    GetWindowTextW = user32.GetWindowTextW
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    IsWindowVisible = user32.IsWindowVisible
    GetClassNameW = user32.GetClassNameW
    SetForegroundWindow = user32.SetForegroundWindow
    BringWindowToTop = user32.BringWindowToTop
    ShowWindow = user32.ShowWindow
    MoveWindow = user32.MoveWindow
    GetSystemMetrics = user32.GetSystemMetrics
    GetWindowRect = user32.GetWindowRect
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    AttachThreadInput = user32.AttachThreadInput
    SetFocus = user32.SetFocus
    keybd_event = user32.keybd_event  # legacy but fine
    SW_RESTORE = 9
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    # Virtual key codes
    VK = {
        "ALT": 0x12, "MENU": 0x12, "CTRL": 0x11, "CONTROL": 0x11, "SHIFT": 0x10, "WIN": 0x5B,
        "ENTER": 0x0D, "RETURN": 0x0D, "SPACE": 0x20, "TAB": 0x09, "ESC": 0x1B, "ESCAPE": 0x1B,
        "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74, "F6": 0x75, "F7": 0x76,
        "F8": 0x77, "F9": 0x78, "F10": 0x79, "F11": 0x7A, "F12": 0x7B
    }
    for i, ch in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ"): VK[ch] = 0x41 + i
    for i, ch in enumerate("0123456789"): VK[ch] = 0x30 + i

def _parse_hotkey(spec: str):
    if not IS_WINDOWS: return []
    seq = []
    try:
        parts = [p.strip().upper() for p in spec.split("+") if p.strip()]
        for p in parts:
            if p in VK: seq.append(VK[p])
            elif len(p) == 1 and p.isalnum(): seq.append(VK[p.upper()])
        mods = [VK["CTRL"], VK["SHIFT"], VK["ALT"], VK["WIN"]]
        seq = [k for k in mods if k in seq] + [k for k in seq if k not in mods]
    except Exception:
        seq = [VK["ALT"], VK["S"]]
    return seq

def _send_hotkey_to_foreground(vkeys):
    if not IS_WINDOWS or not vkeys: return
    pressed = []
    for vk in vkeys:
        keybd_event(vk, 0, 0, 0)
        pressed.append(vk)
        time.sleep(0.01)
    for vk in reversed(pressed):
        keybd_event(vk, 0, 2, 0)
        time.sleep(0.01)

def _enum_windows():
    wins = []
    if not IS_WINDOWS: return wins
    def callback(hwnd, lParam):
        if not IsWindowVisible(hwnd): return True
        length = GetWindowTextLengthW(hwnd)
        title = ctypes.create_unicode_buffer(length + 1)
        GetWindowTextW(hwnd, title, length + 1)
        cls = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, cls, 256)
        wins.append((hwnd, title.value, cls.value))
        return True
    EnumWindows(EnumWindowsProc(callback), 0)
    return wins

# ---------- REPLACED: robust LightBurn window finder ----------
def _find_lightburn_window():
    if not IS_WINDOWS:
        return None
    # Prefer the largest visible Qt window whose title includes "LightBurn"
    best = None
    best_area = -1
    rect = wintypes.RECT()
    for hwnd, title, cls in _enum_windows():
        t = (title or "").lower()
        if "lightburn" not in t:
            continue
        if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            w = max(0, rect.right - rect.left)
            h = max(0, rect.bottom - rect.top)
            area = w * h
            if area > best_area:
                best_area = area
                best = hwnd
    return best

def _find_notepad_window():
    if not IS_WINDOWS: return None
    for hwnd, title, cls in _enum_windows():
        t = (title or "").lower()
        if "notepad" in t: return hwnd
    return None

def _move_window_bottom_right(hwnd, width, height, margin=10):
    if not IS_WINDOWS or not hwnd: return
    sw = GetSystemMetrics(SM_CXSCREEN)
    sh = GetSystemMetrics(SM_CYSCREEN)
    x = max(0, sw - width - margin)
    y = max(0, sh - height - margin)
    try:
        ShowWindow(hwnd, SW_RESTORE)
        MoveWindow(hwnd, x, y, width, height, True)
    except Exception:
        pass

def _focus_window(hwnd):
    if not IS_WINDOWS or not hwnd: return
    try:
        ShowWindow(hwnd, SW_RESTORE)
        BringWindowToTop(hwnd)
        SetForegroundWindow(hwnd)
        SetFocus(hwnd)
    except Exception:
        pass

# Stronger foreground enforcement
def _force_foreground(hwnd):
    if not IS_WINDOWS or not hwnd:
        return
    try:
        ShowWindow(hwnd, SW_RESTORE)
        BringWindowToTop(hwnd)
        SetForegroundWindow(hwnd)
        SetFocus(hwnd)
        time.sleep(0.05)

        cur_tid = kernel32.GetCurrentThreadId()
        pid = wintypes.DWORD()
        target_tid = GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if target_tid and target_tid != cur_tid:
            AttachThreadInput(cur_tid, target_tid, True)
            try:
                BringWindowToTop(hwnd)
                SetForegroundWindow(hwnd)
                SetFocus(hwnd)
            finally:
                AttachThreadInput(cur_tid, target_tid, False)
    except Exception:
        pass

# ---------- NEW foreground helpers ----------
def _allow_set_foreground_for_all():
    try:
        ASFW_ANY = -1
        user32.AllowSetForegroundWindow(ASFW_ANY)
    except Exception:
        pass

def _press_and_release_alt():
    try:
        keybd_event(VK["ALT"], 0, 0, 0)
        time.sleep(0.02)
        keybd_event(VK["ALT"], 0, 2, 0)
        time.sleep(0.02)
    except Exception:
        pass

# ----------------- Default Config -----------------
def today_code(): return datetime.now().strftime("%y%m%d")
def serial4(n): return f"{n:04d}"

DEFAULT_CONFIG = {
    "ROOT": r"S:\101 - Engineering\15 - Engineering Personal Folders\Riley Brugger\LaserSerials\LaserSerials",
    "JOBS": {
        "Job A": {"display_name": "Job A", "part_number": "PN123", "default_batch": 28, "lightburn_file": r"S:\101 - Engineering\15 - Engineering Personal Folders\Riley Brugger\LaserSerials\LaserSerials\Jobs\JobA.lbrn2", "focus_height": 0.0, "select_pin": 0},
        "Job B": {"display_name": "Job B", "part_number": "PN456", "default_batch": 28, "lightburn_file": r"S:\101 - Engineering\15 - Engineering Personal Folders\Riley Brugger\LaserSerials\LaserSerials\Jobs\JobB.lbrn2", "focus_height": 0.0, "select_pin": 8},
        "Job C": {"display_name": "Job C", "part_number": "PN789", "default_batch": 28, "lightburn_file": r"S:\101 - Engineering\15 - Engineering Personal Folders\Riley Brugger\LaserSerials\LaserSerials\Jobs\JobC.lbrn2", "focus_height": 0.0, "select_pin": 9},
        "Job D": {"display_name": "Job D", "part_number": "PN000", "default_batch": 28, "lightburn_file": "", "focus_height": 0.0, "select_pin": 16},
        "Job E": {"display_name": "Job E", "part_number": "PN000", "default_batch": 28, "lightburn_file": "", "focus_height": 0.0, "select_pin": 17},
    },
    "OPEN_LB_FILE_ON_START": True,
    "LIGHTBURN": {
        "exe_path": "",
        "start_delay_sec": 2,
        "post_open_delay_sec": 3,
        "window_width": 520,
        "window_height": 380,
        "window_margin": 8,
        "start_macro_hotkey": "ALT+S",
        "resize_on_job_change": False,
        "enable_positioning": False,
        "enable_start_hotkey": True,
        "test_hotkey_with_notepad": False  # Set to True to test hotkey with Notepad instead of LightBurn
    },
    "SERIAL": {
        "enabled": False,
        "port": "COM5",  # default COM for serial
        "baud": 112500,
        "start_command": "START\n",
        "stop_command": "STOP\n",
        "autofocus_command": "AF\n",
        "done_token": "DONE",
        "poll_ms": 50
    },
    "RELAYS": {
        "names": ["Auto Focus", "Foot Pedal", "Air Exhaust", "Door Lock"],
        "modes": ["toggle", "toggle", "toggle", "switch"],
        "pulse_ms": [250, 150, 150, 0],
        "pins": [10, 11, 22, 23]
    },
    "INPUTS": {
        "names": ["Job Done", "Laser Error", "E-Stop", "Door Interlock"],
        "pins": [1, 2, 3, 15]
    },
    "SIMULATE": {"enabled": False, "batch_done_on_done": True},
    "MACHINE": {"enabled": False, "code": ""},
    "UI": {
        "up_next_tail": 28,
        "job_button_colors": {"Job A": "#2ecc71", "Job B": "#3498db", "Job C": "#e67e22", "Job D": "#9b59b6", "Job E": "#7f8c8d"},
        "show_simulate_done": True,
        "show_test_buttons": True,
        "show_io_sim": True
    },
    "LOGGING": {"daily_max_rows": 20000, "retain_mode": "off", "retain_days": 7, "write_planned": True},
    "HISTORY": {
        "path": r"S:\101 - Engineering\15 - Engineering Personal Folders\Riley Brugger\LaserSerials\LaserSerials\Logs\All_Jobs_History.csv",
        "enabled": True
    },
    "LAST_DATE_CODE": today_code()
}

def deep_merge(a, b):
    out = dict(a)
    for k, v in b.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out

def build_fullcode(pn, dc, s4, cfg):
    mc = ""
    try:
        if cfg.get("MACHINE", {}).get("enabled"):
            raw = (cfg.get("MACHINE", {}).get("code") or "").strip().upper()
            if raw: mc = f"-{raw[:2]}"
    except Exception:
        pass
    return f"{pn}-{dc}{mc}-{s4}"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LightBurn Serial GUI v3.7")
        self.state('zoomed')
        self.attributes('-topmost', True)
        self.resizable(True, True)
        
        self.bind("<Unmap>", self._on_unmap_restore)
        self.CONFIG = self.load_or_init_config()
        self.DIRS, self.LB_BATCH, self.WORKING_BATCH, self.WORKING_COMPLETED_TODAY = self.derive_paths()
        self.ensure_dirs()
        self.filter_completed_today_old_dates()
        self.enforce_retention_on_startup()
        # ESP32 pin levels (for job auto-select)
        self.pin_levels = {}  # {int pin: 0/1}
        # Runtime IO state
        self.relay_state = [0] * len(self.CONFIG["RELAYS"]["names"])
        self.input_state = [0] * len(self.CONFIG["INPUTS"]["names"])
        self.input_override = [False] * len(self.CONFIG["INPUTS"]["names"])  # override flags for inputs
        self.door_closed = False
        # NEW: engraving state to lock out START button
        self.is_engraving = False

        # Banner
        self.banner = tk.Label(self, text="Idle", font=("Segoe UI", 14, "bold"), fg="white", bg="#7f8c8d", padx=10, pady=6)
        self.banner.pack(fill="x")
        # Top row
        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        self.clock_var = tk.StringVar(value="--:--:--")
        ttk.Label(top, textvariable=self.clock_var, font=("Segoe UI", 18, "bold")).pack(side="left")
        ms = ttk.Frame(top); ms.pack(side="right")
        ttk.Label(ms, text="Door:").pack(side="left", padx=(0,6))
        self.door_label = tk.Label(ms, text="Unknown", fg="white", bg="#7f8c8d", padx=10, pady=4)
        self.door_label.pack(side="left", padx=(0,10))
        ttk.Label(ms, text="ESP32:").pack(side="left", padx=(0,6))
        self.conn_label = tk.Label(ms, text="Disabled", fg="white", bg="#f39c12", padx=10, pady=4)
        self.conn_label.pack(side="left")
        ttk.Button(ms, text="Settings", command=self.open_settings).pack(side="left", padx=6)
        self.after(200, self.tick_clock)
        # Jobs frame
        jobs = ttk.LabelFrame(self, text="Jobs", padding=10); jobs.pack(fill="x", padx=8, pady=4)
        self.jobs_frame = jobs
        self.job_btns = {}
        self.job_keys = list(self.CONFIG["JOBS"].keys())
        # No selected job until a pin says so; we keep it empty
        self.selected_job = tk.StringVar(value="")
        self.last_selected_job = ""  # track last non-empty job for AF trigger
        self._rebuild_job_buttons()
        # Info row
        info = ttk.Frame(self, padding=8); info.pack(fill="x")
        ttk.Label(info, text="Selected Job:").grid(row=0, column=0, sticky="e")
        initial_job = self.selected_job.get() or ""
        disp_name = self.CONFIG["JOBS"][initial_job]["display_name"] if initial_job else "--"
        self.sel_job_lbl = ttk.Label(info, text=disp_name, font=("Segoe UI", 11, "bold"))
        self.sel_job_lbl.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(info, text="Part Number:").grid(row=0, column=2, sticky="e")
        self.part_var = tk.StringVar(value=(self.CONFIG["JOBS"][initial_job]["part_number"] if initial_job else ""))
        ttk.Label(info, textvariable=self.part_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=3, sticky="w", padx=6)
        ttk.Label(info, text="Date Code:").grid(row=0, column=4, sticky="e")
        self.date_var = tk.StringVar(value="" if not initial_job else today_code())
        ttk.Label(info, textvariable=self.date_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=5, sticky="w", padx=6)
        ttk.Label(info, text="Next Serial:").grid(row=0, column=6, sticky="e")
        self.next_var = tk.StringVar(value="" if not initial_job else "0001")
        ttk.Label(info, textvariable=self.next_var, font=("Segoe UI", 11, "bold")).grid(row=0, column=7, sticky="w", padx=6)
        # Batch
        batch = ttk.Frame(self, padding=8); batch.pack(fill="x")
        ttk.Label(batch, text="Batch Size:").grid(row=0, column=0, sticky="e")
        self.batch_var = tk.StringVar(value=str(self.CONFIG["JOBS"][initial_job]["default_batch"]) if initial_job else "")
        be = ttk.Entry(batch, textvariable=self.batch_var, width=8); be.grid(row=0, column=1, sticky="w", padx=6)
        be.bind("<KeyRelease>", lambda e: self.refresh_preview_upnext_and_lb())
        ttk.Button(batch, text="Set as Default for Job", command=self.set_job_default_batch).grid(row=0, column=2, padx=8)
        # Up Next / Preview
        upnext_frame = ttk.Labelframe(self, text="Up Next (Preview / Current Batch) [0]", padding=8)
        self.left_hdr = upnext_frame; upnext_frame.pack(fill="both", expand=True, padx=8, pady=4)
        inner = ttk.Frame(upnext_frame); inner.pack(fill="both", expand=True)
        self.upnext_list = tk.Listbox(inner, height=24, font=("Consolas", 10))
        self.upnext_list.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(inner, orient="vertical", command=self.upnext_list.yview)
        sb.pack(side="right", fill="y")
        self.upnext_list.config(yscrollcommand=sb.set)
        # Controls
        controls = ttk.Frame(self, padding=8); controls.pack(fill="x")
        self.start_btn = tk.Button(controls, text="START LASER", width=18, height=2, bg="#c0392b", fg="white", font=("Segoe UI", 14, "bold"), command=self.start_laser_flow)
        self.start_btn.pack(side="left")
        self.abort_btn = tk.Button(controls, text="ABORT / STOP", width=18, height=2, bg="#7f8c8d", fg="white", font=("Segoe UI", 14, "bold"), command=self.abort_stop_flow)
        self.abort_btn.pack(side="left", padx=10)
        self.af_btn = ttk.Button(controls, text="Auto Focus", command=self.do_autofocus)
        self.af_btn.pack(side="left", padx=10)
        self.sim_btn = ttk.Button(controls, text="Simulate DONE", command=self.sim_done)
        self.sim_btn.pack(side="left", padx=10)
        # IO sim strip
        self.io_sim_frame = ttk.Frame(controls)
        self.sim_relays = []
        for i, name in enumerate(self.CONFIG["RELAYS"]["names"]):
            b = ttk.Checkbutton(self.io_sim_frame, text=f"Sim {name}", command=lambda ii=i: self.sim_toggle_relay(ii))
            b.pack(side="left", padx=4)
            self.sim_relays.append(b)
        self.sim_inputs = []
        pulse_inputs = {"Job Done", "E-Stop", "Door Interlock"}
        for i, name in enumerate(self.CONFIG["INPUTS"]["names"]):
            if name in pulse_inputs:
                b = ttk.Button(self.io_sim_frame, text=f"Sim {name}", command=lambda ii=i: self.sim_pulse_input(ii))
                b.pack(side="left", padx=4)
                self.sim_inputs.append((b, None))
            else:
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(self.io_sim_frame, text=f"Sim {name}", variable=var, command=lambda ii=i, v=var: self.sim_set_input(ii, v.get()))
                cb.pack(side="left", padx=4)
                self.sim_inputs.append((cb, var))
        self.io_sim_frame.pack(side="left", padx=6)
        self.status_var = tk.StringVar(value="Idle.")
        ttk.Label(self, textvariable=self.status_var).pack(fill="x", padx=8, pady=(0,8))
        # Serial helper
        self.serial = SerialHelper(self.CONFIG["SERIAL"], self.update_conn_pill, self.on_serial_line)
        if self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]:
            self.serial.try_open(); self.after(self.CONFIG["SERIAL"]["poll_ms"], self.poll_serial)
        else:
            self.update_conn_pill("SIMULATE" if self.CONFIG["SIMULATE"]["enabled"] else "Disabled", warn=True)
        # Status pills
        self.io_status_frame = ttk.LabelFrame(self, text="Machine Status", padding=8)
        self.io_status_frame.pack(fill="x", padx=8, pady=4)
        self.status_pills = {}
        door_name_found = False
        for name in self.CONFIG["INPUTS"]["names"]:
            nl = name.lower()
            if "door" in nl or "interlock" in nl: door_name_found = True
            pill = tk.Label(self.io_status_frame, text=f"{name}: --", fg="white", bg="#7f8c8d", padx=10, pady=4)
            pill.pack(side="left", padx=6)
            self.status_pills[name] = pill
        for name in self.CONFIG["RELAYS"]["names"]:
            pill = tk.Label(self.io_status_frame, text=f"{name}: --", fg="white", bg="#7f8c8d", padx=10, pady=4)
            pill.pack(side="left", padx=6)
            self.status_pills[name] = pill
        if not door_name_found:
            self.door_label.config(text="No Door Input", bg="#7f8c8d")
        # Initial updates
        self.refresh_next_serial_label()
        self.refresh_preview_upnext_and_lb()
        self.update_visibility_from_settings()
        self.update_door_ui()  # initialize
        # Auto-open LB
        if self.CONFIG.get("OPEN_LB_FILE_ON_START"):
            delay = int(self.CONFIG.get("LIGHTBURN", {}).get("start_delay_sec", 2) * 1000)
            self.after(delay, self._open_and_position_lb_for_current_job)
        # Date rollover
        self.after(30_000, self._watch_date_rollover)

    # ---------- Fullscreen minimize guard ----------
    def _on_unmap_restore(self, event=None):
        try:
            self.after(120, lambda: self.state('zoomed') if IS_WINDOWS else self.attributes("-fullscreen", True))
        except Exception:
            pass

    # ---------- Paths / Storage ----------
    def derive_paths(self):
        ROOT = self.CONFIG["ROOT"]
        DIRS = {"config": os.path.join(ROOT, "Config"), "lightburn": os.path.join(ROOT, "LightBurn"), "working": os.path.join(ROOT, "Working"), "logs": os.path.join(ROOT, "Logs"), "jobs": os.path.join(ROOT, "Jobs")}
        LB_BATCH = os.path.join(DIRS["lightburn"], "NextBatch.csv")
        WORKING_BATCH = os.path.join(DIRS["working"], "CurrentBatch.csv")
        WORKING_COMPLETED_TODAY = os.path.join(DIRS["working"], "Completed_Today.csv")
        return DIRS, LB_BATCH, WORKING_BATCH, WORKING_COMPLETED_TODAY

    def ensure_dirs(self):
        for d in self.DIRS.values(): os.makedirs(d, exist_ok=True)
        y = os.path.join(self.DIRS["logs"], datetime.now().strftime("%Y"))
        os.makedirs(y, exist_ok=True); os.makedirs(os.path.join(y, today_code()), exist_ok=True)

    def daily_dir(self):
        return os.path.join(self.DIRS["logs"], datetime.now().strftime("%Y"), today_code())

    def config_file_path(self):
        return os.path.join(self.DIRS["config"], "gui_config.json")

    def load_or_init_config(self):
        cfg_dir = os.path.join(DEFAULT_CONFIG["ROOT"], "Config"); os.makedirs(cfg_dir, exist_ok=True)
        cfg_path = os.path.join(cfg_dir, "gui_config.json")
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f: return deep_merge(DEFAULT_CONFIG, json.load(f))
            except Exception:
                return dict(DEFAULT_CONFIG)
        else:
            with open(cfg_path, "w", encoding="utf-8") as f: json.dump(DEFAULT_CONFIG, f, indent=2)
            return dict(DEFAULT_CONFIG)

    def save_config(self):
        try:
            with open(self.config_file_path(), "w", encoding="utf-8") as f: json.dump(self.CONFIG, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Save Config", f"Failed to save config:\n{e}"); return False

    # ---------- Retention / housekeeping ----------
    def enforce_retention_on_startup(self):
        mode = self.CONFIG["LOGGING"].get("retain_mode", "off")
        if mode == "off":
            if os.path.exists(self.DIRS["logs"]):
                try: shutil.rmtree(self.DIRS["logs"])
                except Exception: pass
        elif mode == "today_only":
            y_root = self.DIRS["logs"]
            if os.path.exists(y_root):
                for y in os.listdir(y_root):
                    ypath = os.path.join(y_root, y)
                    if not os.path.isdir(ypath): continue
                    for d in os.listdir(ypath):
                        if d != today_code(): shutil.rmtree(os.path.join(ypath, d), ignore_errors=True)
        elif mode == "days":
            keep_days = max(1, int(self.CONFIG["LOGGING"].get("retain_days", 7)))
            self._purge_older_than_days(keep_days)

    def _purge_older_than_days(self, n_days):
        cutoff = date.today() - timedelta(days=n_days-1)
        y_root = self.DIRS["logs"]
        if not os.path.exists(y_root): return
        for y in os.listdir(y_root):
            ypath = os.path.join(y_root, y)
            if not os.path.isdir(ypath): continue
            for d in os.listdir(ypath):
                dpath = os.path.join(ypath, d)
                if not os.path.isdir(dpath): continue
                try:
                    dt = datetime.strptime(d, "%y%m%d").date()
                    if dt < cutoff: shutil.rmtree(dpath, ignore_errors=True)
                except Exception: pass

    def filter_completed_today_old_dates(self):
        if os.path.exists(self.WORKING_COMPLETED_TODAY):
            kept = []
            curr_dc = today_code()
            with open(self.WORKING_COMPLETED_TODAY, "r", newline="", encoding="utf-8") as fh:
                r = csv.DictReader(fh)
                for row in r:
                    if row.get("DateCode") == curr_dc: kept.append(row)
            with open(self.WORKING_COMPLETED_TODAY, "w", newline="", encoding="utf-8") as fh:
                w = csv.DictWriter(fh, fieldnames=self._completed_header())
                w.writeheader(); w.writerows(kept)

    # ---------- UI helpers ----------
    def set_banner(self, text, color): self.banner.config(text=text, bg=color)
    def banner_idle(self): self.set_banner("Idle", "#7f8c8d")
    def banner_engraving(self): self.set_banner("Engraving… please wait", "#c0392b")
    def banner_warning(self, msg="Warning"): self.set_banner(msg, "#f39c12")
    def banner_ok(self, msg="Complete"): self.set_banner(msg, "#27ae60")

    def tick_clock(self):
        self.clock_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S")); self.after(200, self.tick_clock)

    def update_conn_pill(self, text, ok=False, warn=False):
        bg = "#27ae60" if ok else ("#f39c12" if warn else "#e74c3c")
        self.conn_label.config(text=text, bg=bg)

    def update_door_ui(self):
        if self.door_closed:
            self.door_label.config(text="Closed", bg="#27ae60");
            if not self.is_engraving:
                self.start_btn.config(state="normal")
        else:
            self.door_label.config(text="Open", bg="#e74c3c"); self.start_btn.config(state="disabled")

    def update_visibility_from_settings(self):
        if not self.CONFIG["UI"].get("show_simulate_done", True):
            try: self.sim_btn.pack_forget()
            except Exception: pass
        else:
            if not self.sim_btn.winfo_ismapped(): self.sim_btn.pack(side="left", padx=10)
        if self.CONFIG["UI"].get("show_io_sim", True) and self.CONFIG["UI"].get("show_test_buttons", True):
            if not self.io_sim_frame.winfo_ismapped(): self.io_sim_frame.pack(side="left", padx=6)
        else:
            try: self.io_sim_frame.pack_forget()
            except Exception: pass

    # ---------- Jobs & preview ----------
    def _rebuild_job_buttons(self):
        for w in self.jobs_frame.grid_slaves(): w.destroy()
        self.job_btns.clear()
        self.job_keys = list(self.CONFIG["JOBS"].keys())
        for i, key in enumerate(self.job_keys):
            color = self.CONFIG["UI"]["job_button_colors"].get(key, "#bdc3c7")
            disp = self.CONFIG["JOBS"][key].get("display_name", key)
            b = tk.Button(self.jobs_frame, text=disp, width=16, height=2, bg=color, fg="white", font=("Segoe UI", 12, "bold"), command=lambda k=key: self.on_job_clicked(k))
            b.grid(row=0, column=i, padx=6, pady=4)
            self.job_btns[key] = b
        # Hide all by default; will show when a pin triggers
        self._apply_job_visibility_for_pin_selection()

    def on_job_clicked(self, key):
        if key and key != self.last_selected_job:
            self.do_autofocus()
        if key:
            self.last_selected_job = key
        self.selected_job.set(key)
        self.sel_job_lbl.config(text=self.CONFIG["JOBS"][key].get("display_name", key))
        self.part_var.set(self.CONFIG["JOBS"][key]["part_number"])
        self.date_var.set(today_code())
        self.batch_var.set(str(self.CONFIG["JOBS"][key]["default_batch"]))
        self.refresh_next_serial_label()
        self.refresh_preview_upnext_and_lb()
        self.status_var.set(f"Selected {self.CONFIG['JOBS'][key].get('display_name', key)}.")
        self.banner_idle()
        if self.CONFIG.get("OPEN_LB_FILE_ON_START"):
            self.after(100, self._open_and_position_lb_for_current_job)
        # keep visibility consistent
        self._apply_job_visibility_for_pin_selection()

    def compute_next_serial_from_completed(self, pn):
        mode = self.CONFIG["LOGGING"].get("retain_mode", "off")
        max_ser = 0
        if mode == "off":
            if os.path.exists(self.WORKING_COMPLETED_TODAY):
                with open(self.WORKING_COMPLETED_TODAY, "r", newline="", encoding="utf-8") as fh:
                    for row in csv.DictReader(fh):
                        if row.get("PartNumber")==pn and row.get("DateCode")==today_code():
                            try: max_ser = max(max_ser, int(row.get("Serial4","0000")))
                            except: pass
        else:
            for f in self.completed_files_today():
                if os.path.exists(f):
                    with open(f, "r", newline="", encoding="utf-8") as fh:
                        for row in csv.DictReader(f):
                            if row.get("PartNumber")==pn and row.get("DateCode")==today_code():
                                try: max_ser = max(max_ser, int(row.get("Serial4","0000")))
                                except: pass
        return max_ser + 1

    def refresh_next_serial_label(self):
        try:
            if not self.selected_job.get(): self.next_var.set(""); return
            pn = self.CONFIG["JOBS"][self.selected_job.get()]["part_number"]
        except Exception: pn = ""
        next_n = self.compute_next_serial_from_completed(pn) if pn else 1
        try: self.next_var.set(f"{next_n:04d}")
        except Exception: pass

    # ---------- Log file helpers (rotation) ----------
    def completed_files_today(self):
        base = os.path.join(self.daily_dir(), "Completed.csv")
        pattern = os.path.join(self.daily_dir(), "Completed*.csv")
        files = sorted(glob.glob(pattern), key=self._completed_sort_key)
        return files if files else [base]

    def planned_files_today(self):
        base = os.path.join(self.daily_dir(), "Planned.csv")
        pattern = os.path.join(self.daily_dir(), "Planned*.csv")
        files = sorted(glob.glob(pattern), key=self._completed_sort_key)
        return files if files else [base]

    @staticmethod
    def _completed_sort_key(path):
        name = os.path.basename(path)
        if "_" in name:
            try: n = int(name.split("_",1)[1].split(".")[0])
            except Exception: n = 1
        else: n = 1
        return n

    def _next_chunk_path(self, base_path):
        folder, base = os.path.split(base_path)
        stem, ext = os.path.splitext(base)
        existing = sorted(glob.glob(os.path.join(folder, stem+"*.csv")), key=self._completed_sort_key)
        if not existing: return base_path
        last = os.path.basename(existing[-1])
        if "_" in last:
            try: idx = int(last.split("_",1)[1].split(".")[0]) + 1
            except Exception: idx = 2
            return os.path.join(folder, f"{stem}_{idx}.csv")
        else:
            return os.path.join(folder, f"{stem}_2.csv")

    def _needs_rollover(self, path):
        max_rows = int(self.CONFIG.get("LOGGING",{}).get("daily_max_rows", 20000))
        if not os.path.exists(path): return False
        try:
            with open(path, "r", newline="", encoding="utf-8") as fh:
                r = csv.reader(fh); next(r, None)
                for i, _ in enumerate(r, start=1):
                    if i >= max_rows: return True
        except Exception: pass
        return False

    def _append_row(self, path, fieldnames, row):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        write_header = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fieldnames)
            if write_header: w.writeheader()
            w.writerow(row)

    def append_planned(self, rows):
        if self.CONFIG["LOGGING"].get("retain_mode","off") == "off": return
        if not self.CONFIG["LOGGING"].get("write_planned", True): return
        base = os.path.join(self.daily_dir(), "Planned.csv")
        target = self._next_chunk_path(base) if self._needs_rollover(base) else base
        fns = ["Date","JobID","JobName","PartNumber","DateCode","Serial4","FullCode"]
        with open(target, "a", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fns)
            if fh.tell() == 0: w.writeheader()
            for r in rows:
                w.writerow({
                    "Date": datetime.now().strftime("%Y-%m-%d"),
                    "JobID": r["JobID"],
                    "JobName": r["JobName"],
                    "PartNumber": r["PartNumber"],
                    "DateCode": r["DateCode"],
                    "Serial4": r["Serial4"],
                    "FullCode": r["FullCode"],
                })

    def _completed_header(self):
        return ["Time","JobName","PartNumber","DateCode","Serial4","Serial Number","Result"]

    def append_completed_many(self, rows, result="OK"):
        mode = self.CONFIG["LOGGING"].get("retain_mode", "off")
        fns = self._completed_header()
        out_rows = [{
            "Time": datetime.now().strftime("%H:%M:%S"),
            "JobName": r["JobName"],
            "PartNumber": r["PartNumber"],
            "DateCode": r["DateCode"],
            "Serial4": r["Serial4"],
            "Serial Number": r["FullCode"],
            "Result": result
        } for r in rows]
        if mode == "off":
            path = self.WORKING_COMPLETED_TODAY
        else:
            base = os.path.join(self.daily_dir(), "Completed.csv")
            path = self._next_chunk_path(base) if self._needs_rollover(base) else base
        os.makedirs(os.path.dirname(path), exist_ok=True)
        write_header = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fns)
            if write_header: w.writeheader()
            for row in out_rows: w.writerow(row)

    # Rolling history CSV
    def append_history_csv(self, rows, result="OK"):
        if not self.CONFIG["HISTORY"].get("enabled", False): return
        hist_path = (self.CONFIG.get("HISTORY", {}) or {}).get("path", "").strip()
        if not hist_path: return
        os.makedirs(os.path.dirname(hist_path), exist_ok=True)
        cols = ["Time","JobName","PartNumber","DateCode","Serial Number","Result"]
        write_header = not os.path.exists(hist_path)
        with open(hist_path, "a", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=cols)
            if write_header: w.writeheader()
            now_t = datetime.now().strftime("%H:%M:%S")
            for r in rows:
                w.writerow({
                    "Time": now_t,
                    "JobName": r["JobName"],
                    "PartNumber": r["PartNumber"],
                    "DateCode": r["DateCode"],
                    "Serial Number": r["FullCode"],
                    "Result": result if result in ("OK","SIM","Sim","sim") else "OK"
                })

    # ---------- Working batch ----------
    def read_working_batch(self):
        if not os.path.exists(self.WORKING_BATCH): return []
        with open(self.WORKING_BATCH, "r", newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))

    def write_working_batch(self, rows):
        with open(self.WORKING_BATCH, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["Date","JobID","JobName","PartNumber","DateCode","Serial4","FullCode"])
            w.writeheader(); w.writerows(rows)

    def write_lightburn_batch(self, codes):
        with open(self.LB_BATCH, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f); w.writerow(["CODE"]); [w.writerow([c]) for c in codes]

    # ---------- Preview ----------
    def build_preview_rows(self):
        if not self.selected_job.get(): return [], []
        key = self.selected_job.get(); cfg = self.CONFIG["JOBS"][key]
        job_name = cfg.get("display_name", key)
        pn, dc = cfg["part_number"], today_code()
        try: n = max(1, int(self.batch_var.get()))
        except Exception: n = cfg["default_batch"]
        start_ser = self.compute_next_serial_from_completed(pn)
        rows, codes = [], []
        for i in range(n):
            s4 = serial4(start_ser + i)
            fc = build_fullcode(pn, dc, s4, self.CONFIG)
            rows.append({"Date": datetime.now().strftime("%Y-%m-%d"), "JobID": key, "JobName": job_name, "PartNumber": pn, "DateCode": dc, "Serial4": s4, "FullCode": fc})
            codes.append(fc)
        return rows, codes

    def refresh_preview_upnext_and_lb(self):
        rows, codes = self.build_preview_rows()
        self.upnext_list.delete(0, tk.END)
        for r in rows[:self.CONFIG["UI"]["up_next_tail"]]:
            self.upnext_list.insert(tk.END, f'{r["FullCode"]}')
        try: self.left_hdr.config(text=f"Up Next (Preview / Current Batch) [{len(rows)}]")
        except Exception: pass
        self.write_lightburn_batch(codes)

    # ---------- LightBurn orchestration ----------
    def _current_job_lb_path(self):
        try: return self.CONFIG["JOBS"][self.selected_job.get()]["lightburn_file"]
        except Exception: return ""

    def _open_lightburn_file(self, path):
        if not path or not os.path.exists(path):
            messagebox.showerror("LightBurn File", f"File not found:\n{path}"); return False
        try:
            exe = (self.CONFIG.get("LIGHTBURN", {}) or {}).get("exe_path", "").strip()
            if IS_WINDOWS and exe and os.path.exists(exe):
                subprocess.Popen([exe, path], shell=False)
            else:
                if IS_WINDOWS: os.startfile(path)  # type: ignore
                elif platform.system()=="Darwin": subprocess.Popen(["open", path])
                else: subprocess.Popen(["xdg-open", path])
            return True
        except Exception as e:
            messagebox.showerror("Open File", f"Could not open LightBurn file.\n{e}")
            return False

    def _open_and_position_lb_for_current_job(self):
        lb_path = self._current_job_lb_path()
        if not lb_path or not os.path.exists(lb_path):
            self.status_var.set("LightBurn project missing for current job. Open Settings → Jobs to fix path.")
            return
        ok = self._open_lightburn_file(lb_path)
        if not ok: return
        delay_ms = int(self.CONFIG.get("LIGHTBURN", {}).get("post_open_delay_sec", 3) * 1000)
        self.after(delay_ms, self._do_position_bring_send_hotkey)

    def _position_lb_small_bottom_right(self):
        if not IS_WINDOWS: return
        if not self.CONFIG.get("LIGHTBURN", {}).get("enable_positioning", False): return
        hwnd = _find_lightburn_window()
        if not hwnd: return
        cfg = self.CONFIG.get("LIGHTBURN", {})
        w = int(cfg.get("window_width", 520)); h = int(cfg.get("window_height", 380)); m = int(cfg.get("window_margin", 8))
        _move_window_bottom_right(hwnd, w, h, m)
        try: self.lift(); self.focus_force(); self.attributes('-topmost', True)
        except Exception: pass

    # ---------- REPLACED: robust bring-to-front that mirrors Notepad flow ----------
    def _do_position_bring_send_hotkey(self):
        if not IS_WINDOWS:
            return

        cfg = self.CONFIG.get("LIGHTBURN", {})
        use_notepad = bool(cfg.get("test_hotkey_with_notepad", False))

        # 1) Minimize our GUI to clear the way
        try:
            self.attributes('-topmost', False)
            self.state('iconic')
        except Exception:
            pass
        time.sleep(1) # Give it a moment to minimize

        # 2) Find/launch target
        if use_notepad:
            subprocess.Popen(["notepad.exe"])
            time.sleep(0.8)
            hwnd_target = _find_notepad_window()
        else:
            hwnd_target = None
            for _ in range(30):  # Wait up to 3 seconds for LightBurn to appear
                hwnd_target = _find_lightburn_window()
                if hwnd_target:
                    break
                time.sleep(0.10)
        if not hwnd_target:
            time.sleep(0.4)
            hwnd_target = _find_notepad_window() if use_notepad else _find_lightburn_window()
            if not hwnd_target:
                # If we couldn't find the target, restore our GUI anyway
                self.state('zoomed')
                self.attributes('-topmost', True)
                return  # quietly bail

        # 3) Force target window to fullscreen and foreground
        try:
            user32.ShowWindow(hwnd_target, ctypes.c_int(SW_RESTORE))
            user32.SetForegroundWindow(hwnd_target)
            user32.ShowWindow(hwnd_target, ctypes.c_int(SW_SHOW))
        except Exception:
            pass
        
        time.sleep(2) # Give LightBurn a moment to fully appear

        # 4) Send the hotkey (if enabled)
        if cfg.get("enable_start_hotkey", True):
            spec = cfg.get("start_macro_hotkey", "ALT+S")
            vks = _parse_hotkey(spec)
            _send_hotkey_to_foreground(vks)
            time.sleep(0.40)

        # 5) Return focus to our GUI and restore topmost/fullscreen
        try:
            self.attributes('-topmost', True)
            self.state('zoomed')
            self.lift()
            self.focus_force()
        except Exception:
            pass

    # ---------- Start / Stop / Done ----------
    def _reenable_start_button(self):
        self.is_engraving = False
        try:
            # Only enable if door is closed
            self.start_btn.config(state=("normal" if self.door_closed else "disabled"))
            self.af_btn.config(state="normal") # Re-enable Autofocus button
        except Exception:
            pass

    def start_laser_flow(self):
        # Guard: disallow if already engraving
        if self.is_engraving:
            messagebox.showinfo("Engraving in Progress", "An engraving batch is already running. Please wait for it to finish or press ABORT / STOP.")
            return

        if not self.door_closed:
            messagebox.showwarning("Door Open", "Close the door (interlock) to start the laser.")
            return
        if not self.selected_job.get():
            messagebox.showwarning("No Job", "No job selected. Trigger a job select pin or pick a job in Settings → Jobs.")
            return

        rows, codes = self.build_preview_rows()
        if not rows:
            messagebox.showwarning("Nothing to Engrave", "No items are queued for engraving.")
            return

        self.write_working_batch(rows); self.append_planned(rows)

        # Enter engraving state: disable START
        self.is_engraving = True
        self.start_btn.config(state="disabled", disabledforeground="white")
        self.af_btn.config(state="disabled") # Disable Autofocus button

        self.status_var.set("Batch queued. Engraving… please wait."); self.banner_engraving()

        # Pulse Air Exhaust to toggle it on and turn on Door Lock/Air Assist
        self.do_pulse_relay_by_name("Air Exhaust")
        self.set_relay_by_name("Door Lock", 1)
        self.set_relay_by_name("Air Assist", 1)

        if self.CONFIG.get("OPEN_LB_FILE_ON_START"):
            lb_path = self._current_job_lb_path()
            if lb_path and os.path.exists(lb_path):
                self._open_lightburn_file(lb_path)
                delay_ms = int(self.CONFIG.get("LIGHTBURN", {}).get("post_open_delay_sec", 3) * 1000) + 500
                self.after(delay_ms, self._do_position_bring_send_hotkey)


        if self.CONFIG["SIMULATE"]["enabled"]:
            self.update_conn_pill("SIMULATE", warn=True)
        else:
            if self.CONFIG["SERIAL"]["enabled"]:
                if not self.serial.send_start():
                    self.status_var.set("Could not signal ESP32. Press foot pedal or check serial.")
            else:
                self.update_conn_pill("Disabled", warn=True)

    def abort_stop_flow(self):
        if self.CONFIG["SIMULATE"]["enabled"]:
            self.status_var.set("Simulated STOP. Batch canceled."); self.banner_warning("Aborted")
        else:
            if self.CONFIG["SERIAL"]["enabled"]:
                ok = self.serial.send_stop()
                self.status_var.set("STOP sent to ESP32." if ok else "Failed to send STOP to ESP32.")
            else:
                self.update_conn_pill("Disabled", warn=True)
        self.cancel_batch()
        self._reenable_start_button()

    def do_autofocus(self):
        # Pulse the Auto Focus relay (Relay 0) or send AF command
        if self.CONFIG["SIMULATE"]["enabled"]:
            # Simulate: Pulse Relay 0 (Auto Focus) for configured duration
            ms = int(self.CONFIG["RELAYS"]["pulse_ms"][0])  # Auto Focus relay is index 0
            self.relay_state[0] = 1
            self.on_serial_line(f"RELAY:{0}:1")
            self.after(ms, lambda: self._sim_relay_off_after_pulse(0))
            self.status_var.set(f"Simulated Auto Focus pulse for {ms}ms")
        else:
            if self.CONFIG["SERIAL"]["enabled"]:
                # Send AF command to ESP32 and pulse Relay 0
                ok = self.serial.send_autofocus()
                ms = int(self.CONFIG["RELAYS"]["pulse_ms"][0])
                ok_pulse = self.serial.pulse(0, ms)
                self.status_var.set("AF command and relay pulse sent." if (ok and ok_pulse) else "Failed to send AF command or pulse relay.")
            else:
                self.update_conn_pill("Disabled", warn=True)
                self.status_var.set("Serial disabled; cannot trigger Auto Focus.")

        # Get focus height for the current job
        try:
            if not self.selected_job.get():
                focus_x = 0
            else:
                focus_x = float(self.CONFIG["JOBS"][self.selected_job.get()].get("focus_height", 0))
        except Exception:
            focus_x = 0

        # Show confirmation dialog
        if messagebox.askyesno("Focus Check", f"Was the Height set to {focus_x}?"):
            self.status_var.set(f"Focus confirmed ({focus_x}).")
        else:
            self.status_var.set("Focus NOT confirmed. Retrying Auto Focus.")
            self.after(500, self.do_autofocus)  # Re-run auto-focus after a short delay

    def sim_done(self):
        if self.CONFIG["SIMULATE"]["batch_done_on_done"]:
            self.complete_whole_batch(result="SIM")
        else:
            self.complete_one_item(result="SIM")

    def cancel_batch(self):
        if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
        self.write_lightburn_batch([]); self.refresh_preview_upnext_and_lb()
        self.status_var.set("Batch aborted / canceled. Queue cleared."); self.banner_warning("Aborted")
        self.refresh_next_serial_label()
        # Turn off relays on abort
        self.set_relay_by_name("Door Lock", 0)
        self.set_relay_by_name("Air Assist", 0)
        self.do_pulse_relay_by_name("Air Exhaust")
        self._reenable_start_button()

    def on_serial_line(self, line):
        line = (line or "").strip()
        if not line: return
        if line == self.CONFIG["SERIAL"]["done_token"]:
            if self.CONFIG["SIMULATE"]["batch_done_on_done"]:
                self.complete_whole_batch()
            else:
                self.complete_one_item()
            return
        if line.upper().startswith("STATE:"): return
        if line.upper().startswith("DOOR:"):
            val = line.split(":",1)[1].strip().upper()
            self.door_closed = (val == "CLOSED"); self.update_door_ui(); return
        if line.upper().startswith("RELAY:"):
            try:
                _, rest = line.split(":",1)
                idx, v = rest.split(":")
                i = int(idx.strip()); val = int(v.strip())
                if 0 <= i < len(self.relay_state):
                    self.relay_state[i] = 1 if val else 0
                    name = self.CONFIG["RELAYS"]["names"][i]
                    if name in self.status_pills: self._set_pill(name, self.relay_state[i])
            except Exception: pass
            return
        if line.upper().startswith("INPUT:"):
            try:
                _, rest = line.split(":",1)
                parts = rest.split(":")
                if len(parts)==2 and parts[0].strip().isdigit():
                    ii = int(parts[0].strip()); val = int(parts[1].strip())
                    name = self.CONFIG["INPUTS"]["names"][ii]
                else:
                    name_token = parts[0].strip()
                    val = int(parts[1].strip())
                    name_map = {n.upper().replace(" ","_"): idx for idx, n in enumerate(self.CONFIG["INPUTS"]["names"])}
                    ii = name_map.get(name_token.upper(), 0)
                    name = self.CONFIG["INPUTS"]["names"][ii]
                if self.input_override[ii]: return  # Ignore if overriding
                nl = name.lower()
                if "door" in nl or "interlock" in nl:
                    self.door_closed = bool(val); self.update_door_ui()
                if name in self.status_pills: self._set_pill(name, val)
                if 0 <= ii < len(self.input_state): self.input_state[ii] = val
            except Exception: pass
            return
        m = re.match(r"^\s*PIN\s*:\s*(\d+)\s*:\s*([01])\s*$", line, re.IGNORECASE)
        if m:
            p = int(m.group(1)); val = int(m.group(2))
            self.pin_levels[p] = val
            self._maybe_select_job_from_pin(p, val)
            return
        if line.startswith("{") and line.endswith("}"):
            try:
                js = json.loads(line)
                ins = js.get("inputs", [0]*len(self.input_state))
                if isinstance(ins, list) and len(ins) >= len(self.input_state):
                    for ii, v in enumerate(ins):
                        val = 1 if v else 0
                        if self.input_override[ii]: continue  # Ignore if overriding
                        name = self.CONFIG["INPUTS"]["names"][ii]
                        nl = name.lower()
                        if "door" in nl or "interlock" in nl:
                            self.door_closed = bool(val); self.update_door_ui()
                        if name in self.status_pills: self._set_pill(name, val)
                        self.input_state[ii] = val
                rs = js.get("relays",[0]*len(self.relay_state))
                if isinstance(rs, list) and len(rs) >= len(self.relay_state):
                    for ii, v in enumerate(rs):
                        val = 1 if v else 0
                        self.relay_state[ii] = val
                        name = self.CONFIG["RELAYS"]["names"][ii]
                        if name in self.status_pills: self._set_pill(name, val)
                pins = js.get("pins", {})
                if isinstance(pins, dict):
                    for k, v in pins.items():
                        try: pk = int(k); pv = 1 if int(v) else 0; self.pin_levels[pk] = pv
                        except Exception: pass
                    self._apply_job_visibility_for_pin_selection()
            except Exception: pass

    def _set_pill(self, name, active):
        pill = self.status_pills.get(name)
        if not pill: return
        if active:
            pill.config(text=f"{name}: ON", bg="#27ae60")
        else:
            pill.config(text=f"{name}: OFF", bg="#7f8c8d")

    def complete_one_item(self, result="OK"):
        rows = self.read_working_batch()
        if not rows:
            self.status_var.set("Idle. (No items in queue)"); self.banner_idle(); return
        head, rest = rows[0], rows[1:]
        self.append_completed_many([head], result=result)
        self.append_history_csv([head], result=result)
        if rest:
            self.write_working_batch(rest); self.write_lightburn_batch([r["FullCode"] for r in rest])
            self.status_var.set("Next ready…"); self.banner_engraving()
        else:
            if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
            self.refresh_preview_upnext_and_lb()
            self.status_var.set("Batch complete."); self.banner_ok("Batch complete")
            self.beep_batch_complete()
            # Turn off relays on completion
            self.set_relay_by_name("Door Lock", 0)
            self.set_relay_by_name("Air Assist", 0)
            self.do_pulse_relay_by_name("Air Exhaust")
            self.refresh_next_serial_label()
            self._reenable_start_button()

    def complete_whole_batch(self, result="OK"):
        rows = self.read_working_batch()
        if not rows:
            self.status_var.set("Idle. (No items in queue)"); self.banner_idle(); return
        self.append_completed_many(rows, result=result)
        self.append_history_csv(rows, result=result)
        if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
        self.refresh_preview_upnext_and_lb(); self.refresh_next_serial_label()
        self.status_var.set("Batch complete (all items written)."); self.banner_ok("Batch complete")
        self.beep_batch_complete()
        # Turn off relays on completion
        self.set_relay_by_name("Door Lock", 0)
        self.set_relay_by_name("Air Assist", 0)
        self.do_pulse_relay_by_name("Air Exhaust")
        self._reenable_start_button()

    def beep_batch_complete(self):
        try:
            if IS_WINDOWS:
                import winsound; winsound.MessageBeep(winsound.MB_ICONASTERISK)
            else:
                print("\a", end="", flush=True)
        except Exception: pass

    def open_lightburn_file(self, path):
        return self._open_lightburn_file(path)

    def set_job_default_batch(self):
        if not self.selected_job.get():
            messagebox.showinfo("No Job", "No job selected yet."); return
        key = self.selected_job.get()
        try: n = max(1, int(self.batch_var.get() or 0))
        except Exception: n = 1
        self.CONFIG["JOBS"][key]["default_batch"] = n
        self.save_config()
        disp = self.CONFIG["JOBS"][key].get("display_name", key)
        self.status_var.set(f"Default batch for {disp} set to {n}.")
        self.refresh_preview_upnext_and_lb()

    # ---------- Settings ----------
    def open_settings(self):
        self.attributes('-topmost', False)
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("1140x820")
        win.attributes('-topmost', True)
        win.lift()

        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True)
        
        # Jobs tab
        jobs_tab = ttk.Frame(nb, padding=10); nb.add(jobs_tab, text="Jobs")
        self._build_job_settings_tab(jobs_tab)
        
        # System tab
        sys_tab = ttk.Frame(nb, padding=10); nb.add(sys_tab, text="System")
        self._root_var = tk.StringVar(value=self.CONFIG["ROOT"])
        self._open_lb_var = tk.BooleanVar(value=self.CONFIG["OPEN_LB_FILE_ON_START"])
        ttk.Label(sys_tab, text="Root Folder:").grid(row=0, column=0, sticky="e")
        ttk.Entry(sys_tab, textvariable=self._root_var, width=60).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(sys_tab, text="Browse…", command=lambda: self._root_var.set(filedialog.askdirectory(title="Pick root folder") or self._root_var.get())).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(sys_tab, text="Open LightBurn file automatically", variable=self._open_lb_var).grid(row=1, column=1, sticky="w", pady=8)
        
        # LightBurn tab
        lb_tab = ttk.Frame(nb, padding=10); nb.add(lb_tab, text="LightBurn")
        self._lb_exe_var = tk.StringVar(value=self.CONFIG["LIGHTBURN"].get("exe_path",""))
        self._lb_start_delay = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("start_delay_sec",2)))
        self._lb_post_delay = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("post_open_delay_sec",3)))
        self._lb_hotkey = tk.StringVar(value=self.CONFIG["LIGHTBURN"].get("start_macro_hotkey","ALT+S"))
        self._lb_w = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("window_width",520)))
        self._lb_h = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("window_height",380)))
        self._lb_m = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("window_margin",8)))
        self._lb_enable_hotkey = tk.BooleanVar(value=bool(self.CONFIG["LIGHTBURN"].get("enable_start_hotkey",True)))
        self._lb_test_hotkey = tk.BooleanVar(value=bool(self.CONFIG["LIGHTBURN"].get("test_hotkey_with_notepad",False)))
        
        ttk.Label(lb_tab, text="LightBurn EXE (optional):").grid(row=0, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_exe_var, width=70).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(lb_tab, text="Browse…", command=lambda: self._lb_exe_var.set(filedialog.askopenfilename(title="Select LightBurn.exe", filetypes=[("EXE","*.exe"),("All files","*.*")]) or self._lb_exe_var.get())).grid(row=0, column=2, sticky="w")
        ttk.Label(lb_tab, text="Hotkey to Start (e.g., ALT+S):").grid(row=1, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_hotkey, width=16).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(lb_tab, text="Open after (sec):").grid(row=2, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_start_delay, width=10).grid(row=2, column=1, sticky="w", padx=6)
        ttk.Label(lb_tab, text="Resize after open (sec):").grid(row=3, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_post_delay, width=10).grid(row=3, column=1, sticky="w", padx=6)
        ttk.Label(lb_tab, text="LB window W x H:").grid(row=4, column=0, sticky="e")
        wh = ttk.Frame(lb_tab); wh.grid(row=4, column=1, sticky="w")
        ttk.Entry(wh, textvariable=self._lb_w, width=8).pack(side="left")
        ttk.Label(wh, text=" x ").pack(side="left")
        ttk.Entry(wh, textvariable=self._lb_h, width=8).pack(side="left")
        ttk.Label(lb_tab, text="Bottom-right margin (px):").grid(row=5, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_m, width=8).grid(row=5, column=1, sticky="w")
        
        ttk.Checkbutton(lb_tab, text="Enable Sending Start Hotkey to LB", variable=self._lb_enable_hotkey).grid(row=6, column=1, sticky="w", pady=6)
        ttk.Checkbutton(lb_tab, text="Test Hotkey with Notepad (instead of LB)", variable=self._lb_test_hotkey).grid(row=7, column=1, sticky="w", pady=6)
        
        # Serial & Sim tab
        ser_tab = ttk.Frame(nb, padding=10); nb.add(ser_tab, text="Serial & Simulate")
        self._ser_en_var = tk.BooleanVar(value=self.CONFIG["SERIAL"]["enabled"])
        self._sim_en_var = tk.BooleanVar(value=self.CONFIG["SIMULATE"]["enabled"])
        self._batch_done_var = tk.BooleanVar(value=self.CONFIG["SIMULATE"]["batch_done_on_done"])
        self._ser_port_var = tk.StringVar(value=self.CONFIG["SERIAL"]["port"])
        self._ser_baud_var = tk.IntVar(value=self.CONFIG["SERIAL"]["baud"])
        self._show_sim_var = tk.BooleanVar(value=self.CONFIG["UI"].get("show_simulate_done", True))
        self._show_tests_var = tk.BooleanVar(value=self.CONFIG["UI"].get("show_test_buttons", True))
        self._show_io_sim_var = tk.BooleanVar(value=self.CONFIG["UI"].get("show_io_sim", True))
        ttk.Checkbutton(ser_tab, text="Enable serial (ESP32 pedal/stop/AF)", variable=self._ser_en_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(ser_tab, text="Simulate Mode (no hardware)", variable=self._sim_en_var).grid(row=0, column=1, sticky="w", padx=12)
        ttk.Checkbutton(ser_tab, text="Treat DONE as 'whole batch complete'", variable=self._batch_done_var).grid(row=0, column=2, sticky="w", padx=12)
        ttk.Label(ser_tab, text="Port:").grid(row=1, column=0, sticky="e"); ttk.Entry(ser_tab, textvariable=self._ser_port_var, width=12).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(ser_tab, text="Baud:").grid(row=1, column=2, sticky="e"); ttk.Entry(ser_tab, textvariable=self._ser_baud_var, width=10).grid(row=1, column=3, sticky="w", padx=6)
        ttk.Separator(ser_tab, orient="horizontal").grid(row=2, column=0, columnspan=6, sticky="ew", pady=10)
        ttk.Checkbutton(ser_tab, text="Show Simulate DONE button", variable=self._show_sim_var).grid(row=3, column=0, sticky="w")
        ttk.Checkbutton(ser_tab, text="Show test buttons (Sim ESP signals)", variable=self._show_tests_var).grid(row=3, column=1, sticky="w", padx=12)
        ttk.Checkbutton(ser_tab, text="Show IO sim strip", variable=self._show_io_sim_var).grid(row=3, column=2, sticky="w", padx=12)
        
        # I/O tab
        io_tab = ttk.Frame(nb, padding=10); nb.add(io_tab, text="I/O")
        self._build_io_settings_tab(io_tab)
        
        # History Log tab
        hist_tab = ttk.Frame(nb, padding=10); nb.add(hist_tab, text="History Log")
        self._hist_path_var = tk.StringVar(value=(self.CONFIG.get("HISTORY", {}) or {}).get("path",""))
        self._hist_en_var = tk.BooleanVar(value=self.CONFIG["HISTORY"].get("enabled", True))
        ttk.Label(hist_tab, text="History CSV path:").grid(row=0, column=0, sticky="e")
        self._hist_entry = ttk.Entry(hist_tab, textvariable=self._hist_path_var, width=80)
        self._hist_entry.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(hist_tab, text="Browse…", command=lambda: self._hist_path_var.set(
            filedialog.asksaveasfilename(title="Select All_Jobs_History.csv", defaultextension=".csv", filetypes=[("CSV","*.csv"),("All files","*.*")]) or self._hist_path_var.get())
        ).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(hist_tab, text="Enable History Log", variable=self._hist_en_var).grid(row=1, column=1, sticky="w", pady=6)
        
        ttk.Button(win, text="Save", command=self._settings_save_and_close(win)).pack(side="right", padx=10, pady=10)
        ttk.Button(win, text="Cancel", command=lambda: (win.destroy(), self.attributes('-topmost', True))).pack(side="right", pady=10)

    def _build_job_settings_tab(self, jobs_tab):
        for w in jobs_tab.winfo_children(): w.destroy()
        row = 0
        self._job_entries = {}
        for key in list(self.CONFIG["JOBS"].keys()):
            jcfg = self.CONFIG["JOBS"][key]
            ttk.Label(jobs_tab, text=key, font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(6,2))
            dn_var = tk.StringVar(value=jcfg.get("display_name", key))
            pn_var = tk.StringVar(value=jcfg["part_number"])
            db_var = tk.IntVar(value=jcfg["default_batch"])
            lb_var = tk.StringVar(value=jcfg["lightburn_file"])
            fx_var = tk.DoubleVar(value=float(jcfg.get("focus_height", 0.0)))
            sp_var = tk.StringVar(value="" if jcfg.get("select_pin") is None else str(jcfg.get("select_pin")))
            color_var = tk.StringVar(value=self.CONFIG["UI"]["job_button_colors"].get(key, "#bdc3c7"))
            ttk.Label(jobs_tab, text="Button Text:").grid(row=row+1, column=0, sticky="e")
            ttk.Entry(jobs_tab, textvariable=dn_var, width=20).grid(row=row+1, column=1, sticky="w", padx=6)
            ttk.Label(jobs_tab, text="Part Number:").grid(row=row+1, column=2, sticky="e")
            ttk.Entry(jobs_tab, textvariable=pn_var, width=20).grid(row=row+1, column=3, sticky="w", padx=6)
            ttk.Label(jobs_tab, text="Default Batch:").grid(row=row+1, column=4, sticky="e")
            ttk.Entry(jobs_tab, textvariable=db_var, width=8).grid(row=row+1, column=5, sticky="w", padx=6)
            ttk.Label(jobs_tab, text=".lbrn2 File:").grid(row=row+2, column=0, sticky="e")
            ttk.Entry(jobs_tab, textvariable=lb_var, width=76).grid(row=row+2, column=1, columnspan=4, sticky="we", padx=6)
            def browse_file(var=lb_var):
                p = filedialog.askopenfilename(title="Select LightBurn file", filetypes=[("LightBurn","*.lbrn2"),("All files","*.*")])
                if p: var.set(p)
            ttk.Button(jobs_tab, text="Browse…", command=browse_file).grid(row=row+2, column=5, sticky="w")
            ttk.Label(jobs_tab, text="Focus Height (X):").grid(row=row+3, column=0, sticky="e")
            ttk.Entry(jobs_tab, textvariable=fx_var, width=10).grid(row=row+3, column=1, sticky="w", padx=6)
            ttk.Label(jobs_tab, text="ESP32 select pin:").grid(row=row+3, column=2, sticky="e")
            ttk.Entry(jobs_tab, textvariable=sp_var, width=10).grid(row=row+3, column=3, sticky="w", padx=6)
            ttk.Label(jobs_tab, text="Button Color:").grid(row=row+3, column=4, sticky="e")
            ttk.Entry(jobs_tab, textvariable=color_var, width=10).grid(row=row+3, column=5, sticky="w", padx=6)
            def remove_this(k=key):
                if messagebox.askyesno("Remove Job", f"Remove {k}?"):
                    del self.CONFIG["JOBS"][k]
                    if self.selected_job.get() == k:
                        new_keys = list(self.CONFIG["JOBS"].keys())
                        self.selected_job.set(new_keys[0] if new_keys else "")
                    self.save_config()
                    self._rebuild_job_buttons()
                    self._build_io_settings_tab(jobs_tab)
            ttk.Button(jobs_tab, text="Remove", command=remove_this).grid(row=row, column=6, pady=4)
            ttk.Separator(jobs_tab, orient="horizontal").grid(row=row+4, column=0, columnspan=7, sticky="ew", pady=8)
            self._job_entries[key] = (dn_var, pn_var, db_var, lb_var, fx_var, sp_var, color_var)
            row += 5
        def add_new_job():
            base = "Job "
            idx = 1
            existing = set(self.CONFIG["JOBS"].keys())
            while f"{base}{idx}" in existing: idx += 1
            key = f"{base}{idx}"
            self.CONFIG["JOBS"][key] = {
                "display_name": key,
                "part_number": f"PN{idx:03d}",
                "default_batch": 28,
                "lightburn_file": "",
                "focus_height": 0.0,
                "select_pin": None
            }
            self.CONFIG["UI"]["job_button_colors"][key] = "#bdc3c7"
            self.save_config()
            self._rebuild_job_buttons()
            self._build_io_settings_tab(jobs_tab)
            self.status_var.set(f"Added {key}. Set its LightBurn file/path before use.")
        ttk.Button(jobs_tab, text="Add Job", command=add_new_job).grid(row=row, column=0, pady=8, sticky="w")

    def _build_io_settings_tab(self, io_tab):
        for w in io_tab.winfo_children(): w.destroy()
        row = 0
        ttk.Label(io_tab, text="Relay Names / Modes", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(0,6))
        self._relay_vars = []
        for i in range(len(self.CONFIG["RELAYS"]["names"])):
            row += 1
            ttk.Label(io_tab, text=f"Relay {i}").grid(row=row, column=0, sticky="e")
            nv = tk.StringVar(value=self.CONFIG["RELAYS"]["names"][i])
            ttk.Entry(io_tab, textvariable=nv, width=18).grid(row=row, column=1, sticky="w", padx=6)
            mv = tk.StringVar(value=self.CONFIG["RELAYS"]["modes"][i])
            ttk.Combobox(io_tab, textvariable=mv, width=10, values=["toggle", "switch"]).grid(row=row, column=2, sticky="w")
            pv = tk.IntVar(value=int(self.CONFIG["RELAYS"]["pulse_ms"][i]))
            ttk.Label(io_tab, text="Pulse (ms)").grid(row=row, column=3, sticky="e")
            ttk.Entry(io_tab, textvariable=pv, width=8).grid(row=row, column=4, sticky="w", padx=6)
            pinv = tk.IntVar(value=int(self.CONFIG["RELAYS"]["pins"][i]))
            ttk.Label(io_tab, text="Pin").grid(row=row, column=5, sticky="e")
            ttk.Entry(io_tab, textvariable=pinv, width=8).grid(row=row, column=6, sticky="w", padx=6)
            
            def mk_toggle(ii=i):
                ms = int(self.CONFIG["RELAYS"]["pulse_ms"][ii])
                if ms > 0:
                    ok = self.serial.pulse(ii, ms)  # use pulse for momentary
                else:
                    ok = self.serial.relay_toggle(ii)
                if not self.CONFIG["SIMULATE"]["enabled"]:
                    self.status_var.set(f"Sent toggle to relay {ii}." if ok else f"Failed to send toggle to relay {ii}.")
            
            ttk.Button(io_tab, text="Toggle", command=mk_toggle).grid(row=row, column=7, padx=6)
            
            if self.CONFIG["RELAYS"]["modes"][i] == "switch":
                frame = ttk.Frame(io_tab); frame.grid(row=row, column=8, padx=6)
                
                def mk_on(ii=i):
                    ok = self.serial.relay_set(ii, 1)
                    if not self.CONFIG["SIMULATE"]["enabled"]:
                        self.status_var.set(f"Sent ON to relay {ii}." if ok else f"Failed to send ON to relay {ii}.")

                def mk_off(ii=i):
                    ok = self.serial.relay_set(ii, 0)
                    if not self.CONFIG["SIMULATE"]["enabled"]:
                        self.status_var.set(f"Sent OFF to relay {ii}." if ok else f"Failed to send OFF to relay {ii}.")

                ttk.Button(frame, text="On", command=mk_on).pack(side="left")
                ttk.Button(frame, text="Off", command=mk_off).pack(side="left", padx=4)
            
            def remove_relay(j=i):
                if messagebox.askyesno("Remove Relay", f"Remove Relay {j}?"):
                    for k in ["names", "modes", "pulse_ms", "pins"]: del self.CONFIG["RELAYS"][k][j]
                    self._build_io_settings_tab(io_tab)
            ttk.Button(io_tab, text="Remove", command=remove_relay).grid(row=row, column=9, padx=6)
            self._relay_vars.append((nv, mv, pv, pinv))
        row += 1
        def add_relay():
            self.CONFIG["RELAYS"]["names"].append("New Relay")
            self.CONFIG["RELAYS"]["modes"].append("toggle")
            self.CONFIG["RELAYS"]["pulse_ms"].append(0)
            self.CONFIG["RELAYS"]["pins"].append(0)
            self._build_io_settings_tab(io_tab)
        ttk.Button(io_tab, text="Add Relay", command=add_relay).grid(row=row, column=0, pady=8, sticky="w")
        ttk.Separator(io_tab, orient="horizontal").grid(row=row+1, column=0, columnspan=10, sticky="ew", pady=8)
        row += 2
        ttk.Label(io_tab, text="Input Names (read-only status shown on main screen)", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(0,6))
        self._input_vars = []
        for i in range(len(self.CONFIG["INPUTS"]["names"])):
            row += 1
            ttk.Label(io_tab, text=f"Input {i}").grid(row=row, column=0, sticky="e")
            nv = tk.StringVar(value=self.CONFIG["INPUTS"]["names"][i])
            ttk.Entry(io_tab, textvariable=nv, width=22).grid(row=row, column=1, sticky="w", padx=6)
            pinv = tk.IntVar(value=int(self.CONFIG["INPUTS"]["pins"][i]))
            ttk.Label(io_tab, text="Pin").grid(row=row, column=2, sticky="e")
            ttk.Entry(io_tab, textvariable=pinv, width=8).grid(row=row, column=3, sticky="w", padx=6)
            def remove_input(j=i):
                if messagebox.askyesno("Remove Input", f"Remove Input {j}?"):
                    for k in ["names", "pins"]: del self.CONFIG["INPUTS"][k][j]
                    self._build_io_settings_tab(io_tab)
            ttk.Button(io_tab, text="Remove", command=remove_input).grid(row=row, column=4, padx=6)
            self._input_vars.append((nv, pinv))
        row += 1
        def add_input():
            self.CONFIG["INPUTS"]["names"].append("New Input")
            self.CONFIG["INPUTS"]["pins"].append(0)
            self._build_io_settings_tab(io_tab)
        ttk.Button(io_tab, text="Add Input", command=add_input).grid(row=row, column=0, pady=8, sticky="w")

    def _settings_save_and_close(self, win):
        def inner():
            # Jobs
            for key, tpl in self._job_entries.items():
                dn, pn, db, lb, fx, sp, color = tpl
                self.CONFIG["JOBS"][key]["display_name"] = dn.get().strip() or key
                self.CONFIG["JOBS"][key]["part_number"] = pn.get().strip()
                try: self.CONFIG["JOBS"][key]["default_batch"] = max(1, int(db.get()))
                except Exception: self.CONFIG["JOBS"][key]["default_batch"] = 1
                self.CONFIG["JOBS"][key]["lightburn_file"] = lb.get().strip()
                try: self.CONFIG["JOBS"][key]["focus_height"] = float(fx.get())
                except Exception: self.CONFIG["JOBS"][key]["focus_height"] = 0.0
                sp_raw = sp.get().strip()
                self.CONFIG["JOBS"][key]["select_pin"] = (int(sp_raw) if sp_raw.isdigit() else None)
                self.CONFIG["UI"]["job_button_colors"][key] = color.get().strip() or "#bdc3c7"
            for key, btn in self.job_btns.items():
                btn.config(text=self.CONFIG["JOBS"][key].get("display_name", key))
                btn.config(bg=self.CONFIG["UI"]["job_button_colors"].get(key, "#bdc3c7"))
            if self.selected_job.get() in self.CONFIG["JOBS"]:
                self.sel_job_lbl.config(text=self.CONFIG["JOBS"][self.selected_job.get()].get("display_name", self.selected_job.get()))
                self.part_var.set(self.CONFIG["JOBS"][self.selected_job.get()]["part_number"])
                self.batch_var.set(str(self.CONFIG["JOBS"][self.selected_job.get()]["default_batch"]))
            self.CONFIG["ROOT"] = (self._root_var.get().strip() or self.CONFIG["ROOT"]).rstrip("\\/")
            self.CONFIG["OPEN_LB_FILE_ON_START"] = bool(self._open_lb_var.get())
            self.CONFIG["LIGHTBURN"]["exe_path"] = self._lb_exe_var.get().strip()
            self.CONFIG["LIGHTBURN"]["start_delay_sec"] = int(self._lb_start_delay.get())
            self.CONFIG["LIGHTBURN"]["post_open_delay_sec"] = int(self._lb_post_delay.get())
            self.CONFIG["LIGHTBURN"]["start_macro_hotkey"] = self._lb_hotkey.get().strip() or "ALT+S"
            self.CONFIG["LIGHTBURN"]["window_width"] = int(self._lb_w.get())
            self.CONFIG["LIGHTBURN"]["window_height"] = int(self._lb_h.get())
            self.CONFIG["LIGHTBURN"]["window_margin"] = int(self._lb_m.get())
            self.CONFIG["LIGHTBURN"]["resize_on_job_change"] = bool(self._lb_resize_on_change.get())
            self.CONFIG["LIGHTBURN"]["enable_positioning"] = bool(self._lb_enable_positioning.get())
            self.CONFIG["LIGHTBURN"]["enable_start_hotkey"] = bool(self._lb_enable_hotkey.get())
            self.CONFIG["LIGHTBURN"]["test_hotkey_with_notepad"] = bool(self._lb_test_hotkey.get())
            self.CONFIG["SERIAL"]["enabled"] = bool(self._ser_en_var.get())
            self.CONFIG["SERIAL"]["port"] = self._ser_port_var.get().strip() or self.CONFIG["SERIAL"]["port"]
            try: self.CONFIG["SERIAL"]["baud"] = int(self._ser_baud_var.get())
            except Exception: pass
            self.CONFIG["SIMULATE"]["enabled"] = bool(self._sim_en_var.get())
            self.CONFIG["SIMULATE"]["batch_done_on_done"] = bool(self._batch_done_var.get())
            self.CONFIG["UI"]["show_simulate_done"] = bool(self._show_sim_var.get())
            self.CONFIG["UI"]["show_test_buttons"] = bool(self._show_tests_var.get())
            self.CONFIG["UI"]["show_io_sim"] = bool(self._show_io_sim_var.get())
            if hasattr(self, "_relay_vars"):
                self.CONFIG["RELAYS"]["names"] = [v[0].get().strip() for v in self._relay_vars]
                self.CONFIG["RELAYS"]["modes"] = [v[1].get() for v in self._relay_vars]
                self.CONFIG["RELAYS"]["pulse_ms"] = [int(v[2].get()) for v in self._relay_vars]
                self.CONFIG["RELAYS"]["pins"] = [int(v[3].get()) for v in self._relay_vars]
            if hasattr(self, "_input_vars"):
                self.CONFIG["INPUTS"]["names"] = [v[0].get().strip() for v in self._input_vars]
                self.CONFIG["INPUTS"]["pins"] = [int(v[1].get()) for v in self._input_vars]
            hp = self._hist_path_var.get().strip()
            self.CONFIG.setdefault("HISTORY", {})
            if hp: self.CONFIG["HISTORY"]["path"] = hp
            self.CONFIG["HISTORY"]["enabled"] = bool(self._hist_en_var.get())
            if self.save_config():
                # SHOW IN FRONT OF SETTINGS
                messagebox.showinfo("Saved", "Settings saved.", parent=win)
            self.DIRS, self.LB_BATCH, self.WORKING_BATCH, self.WORKING_COMPLETED_TODAY = self.derive_paths(); self.ensure_dirs()
            if self.CONFIG["SIMULATE"]["enabled"]:
                self.update_conn_pill("SIMULATE", warn=True)
            elif self.CONFIG["SERIAL"]["enabled"]:
                self.serial.enabled = True; self.serial.port = self.CONFIG["SERIAL"]["port"]; self.serial.baud = self.CONFIG["SERIAL"]["baud"]; self.serial.reconnect()
            else:
                self.serial.enabled = False; self.update_conn_pill("Disabled", warn=True)
            self.refresh_next_serial_label(); self.refresh_preview_upnext_and_lb(); self.update_visibility_from_settings()
            self._rebuild_job_buttons()
            self._apply_job_visibility_for_pin_selection()
            win.destroy()
            self.attributes('-topmost', True)
        return inner

    def poll_serial(self):
        if self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]:
            line = self.serial.readline()
            if line: self.on_serial_line(line)
        self.after(self.CONFIG["SERIAL"]["poll_ms"], self.poll_serial)

    # ---------- Sim controls ----------
    def sim_toggle_relay(self, i):
        ms = int(self.CONFIG["RELAYS"]["pulse_ms"][i])
        name = self.CONFIG["RELAYS"]["names"][i]
        if self.CONFIG["SIMULATE"]["enabled"]:
            if ms > 0:
                self.relay_state[i] = 1
                self.on_serial_line(f"RELAY:{i}:1")
                self.after(ms, lambda: self._sim_relay_off_after_pulse(i))
                self.status_var.set(f"Simulated pulse on {name} for {ms}ms")
            else:
                self.relay_state[i] = 1 - self.relay_state[i]
                self.on_serial_line(f"RELAY:{i}:{self.relay_state[i]}")
                self.status_var.set(f"Simulated toggle on {name} to {'ON' if self.relay_state[i] else 'OFF'}")
        else:
            if ms > 0:
                ok = self.serial.pulse(i, ms)
            else:
                ok = self.serial.relay_toggle(i)
            self.status_var.set(f"Sent {'pulse' if ms>0 else 'toggle'} to {name}." if ok else f"Failed to send to {name}.")

    def _sim_relay_off_after_pulse(self, i):
        self.relay_state[i] = 0
        self.on_serial_line(f"RELAY:{i}:0")

    def sim_set_input(self, i, value):
        name = self.CONFIG["INPUTS"]["names"][i]
        self.input_override[i] = value
        val = 1 if value else 0
        self.input_state[i] = val
        nl = name.lower()
        if "door" in nl or "interlock" in nl:
            self.door_closed = bool(val)
            self.update_door_ui()
        if name in self.status_pills:
            self._set_pill(name, val)
        if self.CONFIG["SIMULATE"]["enabled"]:
            self.on_serial_line(f"INPUT:{i}:{val}")
            self.status_var.set(f"Simulated {name} to {'ON' if value else 'OFF'}")
        else:
            ok = self.serial.sim_input(i, val)
            self.status_var.set(f"Sent sim input to {name}." if ok else f"Failed to send sim input to {name}.")

    def sim_pulse_input(self, i):
        name = self.CONFIG["INPUTS"]["names"][i]
        self.input_override[i] = True
        val = 1  # assume active high for simulation
        self.input_state[i] = val
        nl = name.lower()
        if "door" in nl or "interlock" in nl:
            self.door_closed = True
            self.update_door_ui()
        if name in self.status_pills:
            self._set_pill(name, val)
        if self.CONFIG["SIMULATE"]["enabled"]:
            self.on_serial_line(f"INPUT:{i}:1")
        else:
            self.serial.sim_input(i, 1)
        self.status_var.set(f"Simulated pulse {name} ON")
        self.after(1000, lambda: self._end_sim_pulse_input(i))

    def _end_sim_pulse_input(self, i):
        name = self.CONFIG["INPUTS"]["names"][i]
        self.input_override[i] = False
        val = 0
        self.input_state[i] = val
        nl = name.lower()
        if "door" in nl or "interlock" in nl:
            self.door_closed = False
            self.update_door_ui()
        if name in self.status_pills:
            self._set_pill(name, val)
        if self.CONFIG["SIMULATE"]["enabled"]:
            self.on_serial_line(f"INPUT:{i}:0")
        else:
            self.serial.sim_input(i, 0)
        self.status_var.set(f"Simulated pulse {name} OFF")

    # ---------- Date rollover watcher ----------
    def _watch_date_rollover(self):
        try:
            if self.CONFIG.get("LAST_DATE_CODE") != today_code():
                self.filter_completed_today_old_dates()
                self.CONFIG["LAST_DATE_CODE"] = today_code()
                self.save_config()
                self.date_var.set(today_code())
                self.refresh_next_serial_label()
        finally:
            self.after(30_000, self._watch_date_rollover)

    # ---------- Auto-select job via ESP32 pin ----------
    def _maybe_select_job_from_pin(self, pin, level):
        for key, jcfg in self.CONFIG["JOBS"].items():
            sp = jcfg.get("select_pin")
            if sp is not None and int(sp) == int(pin) and int(level) == 1:
                if self.selected_job.get() != key:
                    self.on_job_clicked(key)
                self._apply_job_visibility_for_pin_selection()
                return
        self._apply_job_visibility_for_pin_selection()

    def _apply_job_visibility_for_pin_selection(self):
        active_key = None
        for key, jcfg in self.CONFIG["JOBS"].items():
            sp = jcfg.get("select_pin")
            if sp is not None and self.pin_levels.get(int(sp), 0) == 1:
                active_key = key; break
        if active_key:
            for key, btn in self.job_btns.items():
                try:
                    if key == active_key: btn.grid()
                    else: btn.grid_remove()
                except Exception: pass
            if self.selected_job.get() != active_key:
                self.on_job_clicked(active_key)
        else:
            for btn in self.job_btns.values():
                try: btn.grid_remove()
                except Exception: pass
            if self.selected_job.get():
                self.selected_job.set("")
                self.sel_job_lbl.config(text="--")
                self.part_var.set("")
                self.date_var.set("")
                self.next_var.set("")
                self.batch_var.set("")
                self.upnext_list.delete(0, tk.END)
                self.left_hdr.config(text="Up Next (Preview / Current Batch) [0]")
                self.write_lightburn_batch([])
                self.refresh_next_serial_label()

    def set_relay_by_name(self, name, val):
        if not self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]: return
        try_names = [name]
        if name == "Door Lock":
            try_names.append("Stack Light")  # alias, if user config still has old name
        
        relays_in_config = self.CONFIG["RELAYS"]["names"]
        
        try:
            for nm in try_names:
                if nm in relays_in_config:
                    i = relays_in_config.index(nm)
                    if self.CONFIG["SIMULATE"]["enabled"]:
                        self.relay_state[i] = val
                        self.on_serial_line(f"RELAY:{i}:{val}")
                    else:
                        ok = self.serial.relay_set(i, val)
                        if ok:
                            self.relay_state[i] = val
                            self._set_pill(nm, val)
                    return
        except ValueError:
            pass  # Name not found

    def do_pulse_relay_by_name(self, name):
        if not self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]: return
        try_names = [name]

        relays_in_config = self.CONFIG["RELAYS"]["names"]

        try:
            for nm in try_names:
                if nm in relays_in_config:
                    i = relays_in_config.index(nm)
                    ms = int(self.CONFIG["RELAYS"]["pulse_ms"][i])
                    if ms > 0:
                        if self.CONFIG["SIMULATE"]["enabled"]:
                            self.relay_state[i] = 1
                            self.on_serial_line(f"RELAY:{i}:1")
                            self.after(ms, lambda: self._sim_relay_off_after_pulse(i))
                            self.status_var.set(f"Simulated pulse on {name} for {ms}ms")
                        else:
                            ok = self.serial.pulse(i, ms)
                            if ok:
                                self.status_var.set(f"Sent pulse to {name}.")
                            else:
                                self.status_var.set(f"Failed to send pulse to {name}.")
                    return
        except (ValueError, IndexError):
            pass # Name not found

class SerialHelper:
    def __init__(self, cfg, status_cb, line_cb):
        self.enabled = bool(cfg.get("enabled", False))
        self.port = cfg.get("port", "COM5")  # default fallback COM5
        self.baud = int(cfg.get("baud", 112500))
        self.start_cmd = cfg.get("start_command", "START\n")
        self.stop_cmd = cfg.get("stop_command", "STOP\n")
        self.af_cmd = cfg.get("autofocus_command", "AF\n")
        self._status_cb = status_cb; self._line_cb = line_cb
        self.ser = None

    def _update_status(self, text, ok=False, warn=False):
        if self._status_cb: self._status_cb(text, ok=ok, warn=warn)

    def try_open(self):
        if not self.enabled: self._update_status("Disabled", warn=True); return False
        try:
            import serial
            if self.ser is None or not self.ser.is_open:
                self.ser = serial.Serial(self.port, self.baud, timeout=0)
            self._update_status(f"Connected ({self.port})", ok=True); return True
        except Exception:
            self._update_status("Disconnected", warn=True); return False
        return True

    def reconnect(self):
        self.close(); return self.try_open()

    def close(self):
        try:
            if self.ser and self.ser.is_open: self.ser.close()
        except Exception: pass
        self._update_status("Disconnected", warn=True)

    def _write(self, data: bytes):
        try:
            if self.try_open() and self.ser: self.ser.write(data); return True
        except Exception:
            self._update_status("Disconnected", warn=True)
        return False

    # High-level commands
    def send_start(self): return self._write(self.start_cmd.encode("utf-8"))
    def send_stop(self): return self._write(self.stop_cmd.encode("utf-8"))
    def send_autofocus(self): return self._write(self.af_cmd.encode("utf-8"))
    def relay_set(self, i, v): return self._write(f"RSET {int(i)} {1 if v else 0}\n".encode())
    def relay_toggle(self, i): return self._write(f"RTGL {int(i)}\n".encode())
    def relay_all(self, bits): return self._write(f"RALL {bits}\n".encode())
    def pulse(self, i, ms): return self._write(f"PULSE {int(i)} {int(ms)}\n".encode())
    def sim_input(self, i, v): return self._write(f"SIMI {int(i)} {1 if v else 0}\n".encode())
    def request_status(self): return self._write(b"STATUS\n")

    def readline(self):
        if not self.enabled or not self.try_open() or not self.ser: return ""
        try:
            line = self.ser.readline().decode(errors="ignore").strip()
            if line and self._line_cb: self._line_cb(line)
            return line
        except Exception:
            self._update_status("Disconnected", warn=True); return ""

if __name__ == "__main__":
    try: App().mainloop()
    except KeyboardInterrupt: sys.exit(0)