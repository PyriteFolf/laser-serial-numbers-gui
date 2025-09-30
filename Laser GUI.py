import os, json, csv, sys, subprocess, platform, glob, shutil, time, re
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports

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
    keybd_event = user32.keybd_event
    SW_RESTORE = 9
    SW_SHOW = 5
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
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

def _find_lightburn_window():
    if not IS_WINDOWS: return None
    for hwnd, title, cls in _enum_windows():
        t = (title or "").lower()
        if "lightburn" in t: return hwnd
    return None

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
        target_tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
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

# ----------------- Default Config -----------------
def today_code(): return datetime.now().strftime("%y%m%d")
def serial4(n): return f"{n:04d}"

DEFAULT_CONFIG = {
    "ROOT": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials",
    "JOBS": {
        "Job 8": {
            "display_name": "Idler 1",
            "part_number": "PN123",
            "default_batch": 28,
            "lightburn_file": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/Jobs/JobA.lbrn2",
            "focus_height": 0.0,
            "select_pattern": "100",
            "job_delay_sec": 5,
            "air_before_sec": 5,
            "air_after_sec": 5
        },
        "Job 9": {
            "display_name": "Idler 2",
            "part_number": "PN456",
            "default_batch": 9,
            "lightburn_file": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/Jobs/JobB.lbrn2",
            "focus_height": 0.0,
            "select_pattern": "010",
            "job_delay_sec": 5,
            "air_before_sec": 5,
            "air_after_sec": 5
        },
        "Job 10": {
            "display_name": "Idler 3",
            "part_number": "PN789",
            "default_batch": 28,
            "lightburn_file": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/Jobs/JobC.lbrn2",
            "focus_height": 0.0,
            "select_pattern": "110",
            "job_delay_sec": 5,
            "air_before_sec": 5,
            "air_after_sec": 5
        },
        "Job 11": {
            "display_name": "Job 11",
            "part_number": "PN000",
            "default_batch": 28,
            "lightburn_file": "",
            "focus_height": 0.0,
            "select_pattern": "111",
            "job_delay_sec": 5,
            "air_before_sec": 5,
            "air_after_sec": 5
        },
        "Job 12": {
            "display_name": "Job 12",
            "part_number": "PN000",
            "default_batch": 28,
            "lightburn_file": "",
            "focus_height": 0.0,
            "select_pattern": "011",
            "job_delay_sec": 5,
            "air_before_sec": 5,
            "air_after_sec": 5
        }
    },
    "OPEN_LB_FILE_ON_START": True,
    "LIGHTBURN": {
        "exe_path": "",
        "start_delay_sec": 2,
        "post_open_delay_sec": 3,
        "enable_start_hotkey": True,
        "test_hotkey_with_notepad": False
    },
    "SERIAL": {
        "enabled": True,
        "port": "COM3",
        "baud": 115200,
        "done_token": "DONE",
        "poll_ms": 10,
        "auto_connect": True # NEW: Auto-connect on startup
    },
    "ESP32_PINS": {
        "board": "esp32-c6-evb",
        "relays": [10, 11, 22, 23, 9],
        "opto_inputs": [1, 2, 3, 15],
    },
    "RELAYS": {
        "names": ["Auto Focus", "Air", "Start", "Door Lock", "Stack Light"],
        "modes": ["pulse", "pulse", "pulse", "pulse", "switch"],
        "pulse_ms": [250, 250, 250, 250, 0],
        "pins": [10, 11, 22, 23, 9],
        "disabled": [False, False, False, False, False]
    },
    "INPUTS": {
        "names": ["Job Sensor 1", "Job Sensor 2", "Job Sensor 3", "Door Sensor"],
        "pins": [1, 2, 3, 17]
    },
    "SIMULATE": {
        "enabled": False,
        "batch_done_on_done": True
    },
    "MACHINE": {
        "enabled": False,
        "code": ""
    },
    "UI": {
        "up_next_tail": 28,
        "job_button_colors": {
            "Job 8": "#2ecc71",
            "Job 9": "#3498db",
            "Job 10": "#e67e22",
            "Job 11": "#9b59b6",
            "Job 12": "#7f8c8d"
        }
    },
    "LOGGING": {
        "daily_max_rows": 20000,
        "retain_mode": "off",
        "retain_days": 7,
        "write_planned": True
    },
    "HISTORY": {
        "enabled": True
    },
    "FILE_PATHS": {
        "next_batch_path": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/LightBurn/NextBatch.csv",
        "completed_today_path": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/Working/Completed_Today.csv",
        "entire_history_path": "S:/101 - Engineering/15 - Engineering Personal Folders/Riley Brugger/LaserSerials/LaserSerials/Logs/All_Jobs_History.csv"
    },
    "AUTOFOCUS": {
        "enabled": True
    },
    "LAST_DATE_CODE": today_code()
}

def load_config_file(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded_config = json.load(f)
            return deep_merge(DEFAULT_CONFIG, loaded_config)
        except Exception:
            return dict(DEFAULT_CONFIG)
    else:
        return dict(DEFAULT_CONFIG)

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
        self.title("LightBurn Serial GUI v4.13")
        try:
            if IS_WINDOWS:
                self.state('zoomed')
            else:
                self.attributes("-fullscreen", True)
            self.attributes('-topmost', True)
        except Exception:
            pass
        self.bind("<Unmap>", self._on_unmap_restore)
        
        self.config_path = self.get_config_path()
        self.CONFIG = load_config_file(self.config_path)
        self.DIRS, self.LB_BATCH, self.WORKING_BATCH, self.WORKING_COMPLETED_TODAY = self.derive_paths()
        self.ensure_dirs()
        self.filter_completed_today_old_dates()
        self.enforce_retention_on_startup()
        
        self.relay_state = [0] * len(self.CONFIG["RELAYS"]["names"])
        self.input_state = [1] * len(self.CONFIG["INPUTS"]["names"])
        self.input_override = [False] * len(self.CONFIG["INPUTS"]["names"])
        
        self.door_closed = False
        self.is_engraving = False
        self.is_air_on = False
        self.af_enabled = self.CONFIG.get("AUTOFOCUS", {}).get("enabled", True)
        self._io_status_labels = {}
        self._io_update_after_id = None
        
        self.ports_dict = {}
        self.port_var = tk.StringVar(value=self.CONFIG["SERIAL"]["port"])
        
        self.last_job_pattern = ""
        self.last_stable_job_time = 0
        self.job_cooldown_ms = 500

        self.banner = tk.Label(self, text="Idle", font=("Segoe UI", 14, "bold"), fg="white", bg="#7f8c8d", padx=10, pady=6)
        self.banner.pack(fill="x")
        
        top = ttk.Frame(self, padding=8); top.pack(fill="x")
        
        self.clock_var = tk.StringVar(value="--:--:--")
        ttk.Label(top, textvariable=self.clock_var, font=("Segoe UI", 18, "bold")).pack(side="left")
        
        right_controls = ttk.Frame(top); right_controls.pack(side="right")
        
        self._create_port_controls_in_frame(right_controls) 

        ms = ttk.Frame(right_controls); ms.pack(side="left", padx=(10, 0))
        ttk.Label(ms, text="Door:").pack(side="left", padx=(0,6))
        self.door_label = tk.Label(ms, text="Unknown", fg="white", bg="#7f8c8d", padx=10, pady=4)
        self.door_label.pack(side="left", padx=(0,10))
        ttk.Label(ms, text="ESP32:").pack(side="left", padx=(0,6))
        self.conn_label = tk.Label(ms, text="Disabled", fg="white", bg="#f39c12", padx=10, pady=4)
        self.conn_label.pack(side="left")
        ttk.Button(ms, text="Settings", command=self.open_settings).pack(side="left", padx=6)
        
        self.after(200, self.tick_clock)
        
        main_content_frame = ttk.Frame(self, padding=8)
        main_content_frame.pack(fill="both", expand=True)

        left_frame = ttk.Frame(main_content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        jobs_container = ttk.Frame(left_frame)
        jobs_container.pack(fill="x")
        
        self.jobs_frame = ttk.LabelFrame(jobs_container, text="Active Job", padding=10)
        self.jobs_frame.pack(fill="x")
        
        self.job_btns = {}
        self.job_keys = list(self.CONFIG["JOBS"].keys())
        self.selected_job = tk.StringVar(value="")
        self.last_selected_job = ""
        self._rebuild_job_buttons()

        info = ttk.Frame(left_frame, padding=8)
        info.pack(fill="x")
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

        batch = ttk.Frame(left_frame, padding=8)
        batch.pack(fill="x")
        ttk.Label(batch, text="Batch Size:").grid(row=0, column=0, sticky="e")
        self.batch_var = tk.StringVar(value=str(self.CONFIG["JOBS"][initial_job]["default_batch"]) if initial_job else "")
        be = ttk.Entry(batch, textvariable=self.batch_var, width=8)
        be.grid(row=0, column=1, sticky="w", padx=6)
        be.bind("<KeyRelease>", lambda e: self.refresh_preview_upnext_and_lb())
        ttk.Button(batch, text="Set as Default for Job", command=self.set_job_default_batch).grid(row=0, column=2, padx=8)

        right_frame = ttk.Frame(main_content_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(8, 0))
        
        upnext_frame = ttk.Labelframe(right_frame, text="Up Next (Preview / Current Batch) [0]", padding=8)
        self.left_hdr = upnext_frame
        upnext_frame.pack(fill="both", expand=True)
        inner = ttk.Frame(upnext_frame)
        inner.pack(fill="both", expand=True)
        self.upnext_list = tk.Listbox(inner, font=("Consolas", 10))
        self.upnext_list.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(inner, orient="vertical", command=self.upnext_list.yview)
        sb.pack(side="right", fill="y")
        self.upnext_list.config(yscrollcommand=sb.set)

        controls = ttk.Frame(right_frame, padding=8)
        controls.pack(fill="x")
        self.open_door_btn = tk.Button(controls, text="OPEN DOOR", width=18, height=2, bg="#3498db", fg="white", font=("Segoe UI", 14, "bold"), command=self.open_door_pulse)
        self.open_door_btn.pack(side="left", padx=10)
        self.af_btn = ttk.Button(controls, text="Auto Focus", command=self.do_autofocus)
        self.af_btn.pack(side="left", padx=10)
        self.sim_btn = ttk.Button(controls, text="Simulate DONE", command=self.sim_done)
        self.sim_btn.pack(side="left", padx=10)

        self.input_status_var = tk.StringVar(value="Inputs: N/A")
        ttk.Label(self, textvariable=self.input_status_var, anchor="w", font=("Segoe UI", 9)).pack(fill="x", padx=8, pady=(4, 0))

        self.status_var = tk.StringVar(value="Idle.")
        ttk.Label(self, textvariable=self.status_var).pack(fill="x", padx=8, pady=(0,8))
        
        self.serial = SerialHelper(self.CONFIG["SERIAL"], self.update_conn_pill, self.on_serial_line, self)
        
        self.selected_job.set("")
        self._apply_job_visibility_for_pin_selection(active_key=None) 
        
        self.update_visibility_from_settings()
        self.update_door_ui()
        self.update_af_button_state()
        self._update_input_status_display()

        if self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]:
            self.after(self.CONFIG["SERIAL"]["poll_ms"], self.poll_serial)
        else:
            self.update_conn_pill("SIMULATE" if self.CONFIG["SIMULATE"]["enabled"] else "Disabled", warn=True)
            
        self.refresh_ports()
        self.after(30_000, self._watch_date_rollover)
        self.after(100, self._check_for_job_select_on_start)
        
        # NEW: Auto-connect logic
        self.after(500, self._auto_connect_on_startup)
        
        if self.CONFIG.get("OPEN_LB_FILE_ON_START"):
            delay = int(self.CONFIG.get("LIGHTBURN", {}).get("start_delay_sec", 2) * 1000)
            self.after(delay, self._open_and_position_lb_for_current_job)
            
    def _auto_connect_on_startup(self):
        """If auto-connect is enabled, find the last used port and connect to it."""
        if self.CONFIG["SERIAL"].get("auto_connect", False):
            last_port = self.CONFIG["SERIAL"].get("port")
            if not last_port:
                return

            # Find the full display name for the stored port (e.g., "COM3")
            full_port_name = ""
            for display_name, device_name in self.ports_dict.items():
                if device_name == last_port:
                    full_port_name = display_name
                    break
            
            if full_port_name:
                print(f"Auto-connecting to last used port: {full_port_name}...")
                self.port_var.set(full_port_name)
                self.toggle_connection()
            else:
                print(f"Could not find last used port '{last_port}' in available ports list.")


    def _get_available_ports(self):
        self.ports_dict = {}
        ports = []
        for port in serial.tools.list_ports.comports():
            display_name = f"{port.device} - {port.description}"
            self.ports_dict[display_name] = port.device
            ports.append(display_name)
        
        if not ports:
            ports = ["No Ports Found"]
            
        return ports

    def refresh_ports(self):
        new_ports = self._get_available_ports()
        
        menu = self.port_menu["menu"]
        menu.delete(0, "end")
        
        for port in new_ports:
            menu.add_command(label=port, command=tk._setit(self.port_var, port))
        
        if not self.port_var.get() or self.port_var.get() not in new_ports:
            config_port_display = next((p for p in new_ports if p.startswith(self.CONFIG["SERIAL"]["port"])), new_ports[0])
            self.port_var.set(config_port_display)

    def _create_port_controls_in_frame(self, container_frame):
        conn_frame = ttk.Frame(container_frame)
        conn_frame.pack(side="right", padx=(10, 0))

        initial_ports = self._get_available_ports()
        self.port_var.set(initial_ports[0])

        self.port_menu = ttk.OptionMenu(conn_frame, self.port_var, initial_ports[0], *initial_ports)
        self.port_menu.config(width=20)
        self.port_menu.pack(side=tk.LEFT, padx=5)

        self.connect_button = tk.Button(conn_frame, text="Connect", command=self.toggle_connection, bg="#4CAF50", fg="white", width=8, font=('Arial', 10, 'bold'))
        self.connect_button.pack(side=tk.LEFT, padx=5)
        
        self.refresh_button = tk.Button(conn_frame, text="Refresh", command=self.refresh_ports, width=8, font=('Arial', 10))
        self.refresh_button.pack(side=tk.LEFT, padx=5)

    def toggle_connection(self):
        if self.serial.ser and self.serial.ser.is_open:
            self.serial.close()
            self.connect_button.config(text="Connect", bg="#4CAF50")
            self.update_conn_pill("Disconnected", warn=True)
            print("Disconnected from serial port via GUI.")
        else:
            selected_display_name = self.port_var.get()
            port_name = self.ports_dict.get(selected_display_name)
            
            if port_name is None or port_name == "No Ports Found":
                messagebox.showerror("Connection Error", "No valid COM port selected.")
                return

            if self.serial.try_open_manual(port_name):
                self.connect_button.config(text="Disconnect", bg="#E74C3C")
                self.initialize_relays_off()
                
                # --- THIS IS THE HANDSHAKE ---
                # After connecting, wait a moment then ask the ESP32 for all current states.
                self.after(250, self.serial.get_all_states)

                # --- NEW: SAVE LAST SUCCESSFUL PORT ---
                self.CONFIG["SERIAL"]["port"] = port_name
                self.save_config()

    def initialize_relays_off(self):
        relay_pins = self.CONFIG["ESP32_PINS"]["relays"] 
        print("Sending initial RSET 0 commands to ensure all relays are OFF...")
        
        if not (self.serial.ser and self.serial.ser.is_open):
            print("Error: Serial connection lost, cannot send RSET commands.")
            return

        for i in range(len(relay_pins)):
            command = f"RSET {i} 0\n"
            try:
                self.serial.ser.write(command.encode('utf-8'))
                time.sleep(0.01) 
            except serial.SerialException as e:
                print(f"Error sending RSET OFF command for Relay {i}: {e}")
                
        print("Initial relays state set to OFF.")

    def get_config_path(self):
        if len(sys.argv) > 1 and sys.argv[1].lower().endswith('.json'):
            return os.path.abspath(sys.argv[1])
        return os.path.join(DEFAULT_CONFIG["ROOT"], "Config", "gui_config.json")
        
    def _on_unmap_restore(self, event=None):
        try:
            self.after(120, lambda: self.state('zoomed') if IS_WINDOWS else self.attributes("-fullscreen", True))
        except Exception:
            pass

    def derive_paths(self):
        ROOT = self.CONFIG["ROOT"]
        DIRS = {"config": os.path.join(ROOT, "Config"), "lightburn": os.path.join(ROOT, "LightBurn"), "working": os.path.join(ROOT, "Working"), "logs": os.path.join(ROOT, "Logs"), "jobs": os.path.join(ROOT, "Jobs")}
        
        LB_BATCH = self.CONFIG.get("FILE_PATHS", {}).get("next_batch_path", os.path.join(DIRS["lightburn"], "NextBatch.csv"))
        WORKING_BATCH = self.CONFIG.get("FILE_PATHS", {}).get("completed_today_path", os.path.join(DIRS["working"], "CurrentBatch.csv"))
        WORKING_COMPLETED_TODAY = self.CONFIG.get("FILE_PATHS", {}).get("entire_history_path", os.path.join(DIRS["working"], "Completed_Today.csv"))
        return DIRS, LB_BATCH, WORKING_BATCH, WORKING_COMPLETED_TODAY

    def ensure_dirs(self):
        for d in self.DIRS.values(): os.makedirs(d, exist_ok=True)
        y = os.path.join(self.DIRS["logs"], datetime.now().strftime("%Y"))
        os.makedirs(y, exist_ok=True); os.makedirs(os.path.join(y, today_code()), exist_ok=True)

    def daily_dir(self):
        return os.path.join(self.DIRS["logs"], datetime.now().strftime("%Y"), today_code())

    def save_config(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f: json.dump(self.CONFIG, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Save Config", f"Failed to save config:\n{e}"); return False

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

    def set_banner(self, text, color): self.banner.config(text=text, bg=color)
    def banner_idle(self): self.set_banner("Idle", "#7f8c8d")
    def banner_engraving(self): self.set_banner("Engraving… please wait", "#c0392b")
    def banner_warning(self): self.set_banner("Warning", "#f39c12")
    def banner_ok(self): self.set_banner("Complete", "#27ae60")

    def tick_clock(self):
        self.clock_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S")); self.after(200, self.tick_clock)

    def update_conn_pill(self, text, ok=False, warn=False):
        bg = "#27ae60" if ok else ("#f39c12" if warn else "#e74c3c")
        self.conn_label.config(text=text, bg=bg)

    def update_door_ui(self):
        if self.door_closed:
            self.door_label.config(text="Closed", bg="#27ae60")
        else:
            self.door_label.config(text="Open", bg="#e74c3c")
        
        if self.is_engraving:
            self.open_door_btn.config(text="JOB RUNNING", state="disabled", bg="#e74c3c")
        else:
            self.open_door_btn.config(text="OPEN DOOR", state="normal", bg="#3498db")

    def update_visibility_from_settings(self):
        if hasattr(self, 'sim_btn'):
            if not self.CONFIG["UI"].get("show_simulate_done", True):
                try: self.sim_btn.pack_forget()
                except Exception: pass
            else:
                if not self.sim_btn.winfo_ismapped(): self.sim_btn.pack(side="left", padx=10)
        
        if hasattr(self, 'open_door_btn'):
            if not self.CONFIG["UI"].get("show_open_door_btn", True):
                try: self.open_door_btn.pack_forget()
                except Exception: pass
            else:
                if not self.open_door_btn.winfo_ismapped(): self.open_door_btn.pack(side="left", padx=10)

        if hasattr(self, 'af_btn'):
            if not self.CONFIG["UI"].get("show_autofocus_btn", True):
                try: self.af_btn.pack_forget()
                except Exception: pass
            else:
                if not self.af_btn.winfo_ismapped(): self.af_btn.pack(side="left", padx=10)

    def _rebuild_job_buttons(self):
        for w in self.jobs_frame.winfo_children(): w.destroy()
        self.job_btns.clear()
        self.job_keys = list(self.CONFIG["JOBS"].keys())
        
        for key in self.job_keys:
            color = self.CONFIG["UI"]["job_button_colors"].get(key, "#bdc3c7")
            disp = self.CONFIG["JOBS"][key].get("display_name", key)
            
            b = tk.Label(self.jobs_frame, text=disp, width=16, height=2, bg=color, fg="white", font=("Segoe UI", 12, "bold"))
            
            b.grid(row=0, column=0, padx=6, pady=4)
            b.grid_remove() 
            self.job_btns[key] = b
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
            self.status_var.set(f"Selected {self.CONFIG['JOBS'][key].get('display_name', key)}. Serial batch loaded.")
            self.banner_idle()
            
            if self.CONFIG.get("OPEN_LB_FILE_ON_START"):
                self.after(100, self._open_and_position_lb_for_current_job)
                
        self._apply_job_visibility_for_pin_selection(active_key=key)

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
                        for row in csv.DictReader(fh):
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

    def append_history_csv(self, rows, result="OK"):
        if not self.CONFIG["HISTORY"].get("enabled", False): return
        hist_path = self.CONFIG.get("FILE_PATHS", {}).get("entire_history_path", "").strip()
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
                if IS_WINDOWS: os.startfile(path)
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
        self.after(delay_ms, self._position_lb_small_bottom_right)

    def _position_lb_small_bottom_right(self):
        if not IS_WINDOWS: return
        hwnd = _find_lightburn_window()
        if not hwnd: return
        
    def _force_foreground(self, hwnd):
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
            target_tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
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

    def _reenable_start_button(self):
        self.is_engraving = False
        self.update_door_ui()

    def start_laser_flow(self):
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
        self.is_engraving = True
        self.update_door_ui()
        self.status_var.set("Batch queued. Engraving… please wait."); self.banner_engraving()
        
        job_cfg = self.CONFIG["JOBS"][self.selected_job.get()]
        
        lb_path = self._current_job_lb_path()
        if lb_path and os.path.exists(lb_path):
            self._open_lightburn_file(lb_path)
            delay_ms = int(self.CONFIG.get("LIGHTBURN", {}).get("post_open_delay_sec", 3) * 1000)
            self.after(delay_ms, self._position_lb_small_bottom_right)

        self.after(500, self._start_air_and_laser_sequence)

    def _start_air_and_laser_sequence(self):
        if not self.is_engraving: return
        job_cfg = self.CONFIG["JOBS"][self.selected_job.get()]
        air_before_sec = job_cfg.get("air_before_sec", 0)
        
        self.status_var.set(f"Turning on air. Waiting for {air_before_sec} seconds before marking...")
        self.set_relay_by_name("Air", 1)
        
        self.after(int(air_before_sec * 1000), self._trigger_start_pulse)

    def _trigger_start_pulse(self):
        if not self.is_engraving: return
        job_cfg = self.CONFIG["JOBS"][self.selected_job.get()]
        job_delay_sec = job_cfg.get("job_delay_sec", 0)
        
        self.status_var.set("Sending START pulse to laser...")
        self.pulse_relay_by_name("Start", self.CONFIG["RELAYS"]["pulse_ms"][self.CONFIG["RELAYS"]["names"].index("Start")])

        self.set_relay_by_name("Door Lock", 1)
        self.set_relay_by_name("Stack Light", 1)
        
        self.after(int(job_delay_sec * 1000), self._unlock_door_and_turn_off_light)

    def _unlock_door_and_turn_off_light(self):
        if not self.is_engraving: return
        self.status_var.set("Job running. Door is now unlocked and stack light is on.")
        self.pulse_relay_by_name("Door Lock", self.CONFIG["RELAYS"]["pulse_ms"][self.CONFIG["RELAYS"]["names"].index("Door Lock")])
        
    def abort_stop_flow(self):
        self.status_var.set("Aborting job. Stopping all processes.")
        self.set_relay_by_name("Air", 0)
        self.set_relay_by_name("Door Lock", 0)
        self.set_relay_by_name("Stack Light", 0)
        self.cancel_batch()
        self._reenable_start_button()

    def do_autofocus(self):
        if not self.af_enabled:
            self.status_var.set("Autofocus is disabled in settings.")
            return

        if self.is_engraving:
            messagebox.showwarning("Job Running", "Cannot run Auto Focus while a job is in progress.")
            return
        
        if not self.selected_job.get():
            messagebox.showwarning("No Job", "Please select a job first.")
            return

        job_cfg = self.CONFIG["JOBS"][self.selected_job.get()]
        focus_x = float(job_cfg.get("focus_height", 0.0))
        
        self.status_var.set(f"Sending Auto Focus pulse...")
        self.pulse_relay_by_name("Auto Focus", self.CONFIG["RELAYS"]["pulse_ms"][self.CONFIG["RELAYS"]["names"].index("Auto Focus")])

        if messagebox.askyesno("Focus Check", f"Was the Height set to {focus_x}?"):
            self.status_var.set(f"Focus confirmed ({focus_x}).")
        else:
            self.status_var.set("Focus NOT confirmed. Please try again.")
            
    def open_door_pulse(self):
        if self.is_engraving:
            messagebox.showwarning("Job Running", "Cannot open the door while a job is in progress.")
            return
        self.pulse_relay_by_name("Door Lock", self.CONFIG["RELAYS"]["pulse_ms"][self.CONFIG["RELAYS"]["names"].index("Door Lock")])
        self.status_var.set("Door unlocked for a brief moment.")
        
    def sim_done(self):
        if self.CONFIG["SIMULATE"]["batch_done_on_done"]:
            self.complete_whole_batch(result="SIM")
        else:
            self.complete_one_item(result="SIM")

    def cancel_batch(self):
        if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
        self.write_lightburn_batch([]); self.refresh_preview_upnext_and_lb()
        self.status_var.set("Batch aborted / canceled. Queue cleared."); self.banner_warning()
        self.refresh_next_serial_label()
        self._reenable_start_button()

    def _update_input_status_display(self):
        """Generates and sets the string showing the current state of all input sensors."""
        status_parts = []
        input_names = self.CONFIG["INPUTS"]["names"]
        
        for i in range(len(input_names)):
            name = input_names[i]
            if i == 3: # Handle Door Sensor (index 3)
                 state_text = "Closed" if self.door_closed else "Open"
            else: # Handle Job Sensors
                # --- DEFINITIVELY FLIPPED LOGIC ---
                # This logic is now inverted to match your hardware's behavior
                if self.input_state[i] == 0:
                    state_text = "Active"
                else:
                    state_text = "Inactive"
            
            status_parts.append(f"{name}: {state_text}")
            
        self.input_status_var.set(" | ".join(status_parts))

    def on_serial_line(self, line):
        line = (line or "").strip()
        if not line: return
        if not line.upper().startswith("INFO:"):
            print(f"[RX] {line}")
        
        if line == self.CONFIG["SERIAL"]["done_token"]:
            if self.CONFIG["SIMULATE"]["batch_done_on_done"]:
                self.complete_whole_batch()
            else:
                self.complete_one_item()
            return

        if line.upper().startswith("RELAY:"):
            try:
                _, rest = line.split(":",1)
                idx, v = rest.split(":")
                i = int(idx.strip()); val = int(v.strip())
                if 0 <= i < len(self.CONFIG["RELAYS"]["names"]):
                    self.relay_state[i] = 1 if val else 0
            except Exception: pass
            return
            
        if line.upper().startswith("INPUT:"):
            try:
                _, rest = line.split(":", 1)
                idx_str, val_str = rest.split(":")
                ii = int(idx_str.strip())
                val = int(val_str.strip())

                if self.input_override[ii]: return
                self.input_state[ii] = val
                
                if ii == 3:
                    is_closed = (val == 0)
                    if is_closed != self.door_closed:
                        self.door_closed = is_closed
                        self.update_door_ui()
                
                self._maybe_select_job_from_input_pattern()
                self._update_input_status_display()
            except Exception: 
                print(f"[ERROR] Failed to parse INPUT line: {line}")
            return

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
            self._reenable_start_button()
        else:
            if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
            self.refresh_preview_upnext_and_lb()
            self.status_var.set("Batch complete."); self.banner_ok()
            self.beep_batch_complete()
            self.refresh_next_serial_label()
            self.set_relay_by_name("Air", 0)
            self.set_relay_by_name("Stack Light", 0)
            self._reenable_start_button()
            
    def complete_whole_batch(self, result="OK"):
        rows = self.read_working_batch()
        if not rows:
            self.status_var.set("Idle. (No items in queue)"); self.banner_idle(); return
        self.append_completed_many(rows, result=result)
        self.append_history_csv(rows, result=result)
        if os.path.exists(self.WORKING_BATCH): os.remove(self.WORKING_BATCH)
        self.refresh_preview_upnext_and_lb(); self.refresh_next_serial_label()
        self.status_var.set("Batch complete (all items written)."); self.banner_ok()
        self.beep_batch_complete()
        self.set_relay_by_name("Air", 0)
        self.set_relay_by_name("Stack Light", 0)
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

    def open_settings(self):
        win = tk.Toplevel(self); win.title("Settings"); win.geometry("1140x820")
        win.attributes('-topmost', True)
        win.lift()
        nb = ttk.Notebook(win); nb.pack(fill="both", expand=True)
        
        jobs_tab = ttk.Frame(nb, padding=10); nb.add(jobs_tab, text="Jobs"); self._build_job_settings_tab(jobs_tab)
        sys_tab = ttk.Frame(nb, padding=10); nb.add(sys_tab, text="System"); self._build_system_settings_tab(sys_tab)
        lb_tab = ttk.Frame(nb, padding=10); nb.add(lb_tab, text="LightBurn"); self._build_lightburn_settings_tab(lb_tab)
        ser_tab = ttk.Frame(nb, padding=10); nb.add(ser_tab, text="Serial & Simulate"); self._build_serial_settings_tab(ser_tab)
        io_tab = ttk.Frame(nb, padding=10); nb.add(io_tab, text="I/O"); self._build_io_settings_tab(io_tab, win)
        paths_tab = ttk.Frame(nb, padding=10); nb.add(paths_tab, text="File Paths"); self._build_paths_settings_tab(paths_tab)
        
        win.mainloop()
    
    def _start_io_status_updates(self):
        if self._io_update_after_id:
            self.after_cancel(self._io_update_after_id)
        
        self._update_io_tab_status()
        self._io_update_after_id = self.after(250, self._start_io_status_updates)

    def _stop_io_status_updates(self, event=None):
        if self._io_update_after_id:
            self.after_cancel(self._io_update_after_id)
            self._io_update_after_id = None
        self._io_status_labels.clear()

    def _update_io_tab_status(self):
        if not self._io_status_labels: return

        try:
            input_names = self.CONFIG["INPUTS"]["names"]
            
            for i, label in self._io_status_labels.items():
                if i < len(input_names):
                    if i == 3: # Handle Door Sensor
                        if self.door_closed:
                            label.config(text="Closed", bg="#27ae60", fg="white")
                        else:
                            label.config(text="Open", bg="#e74c3c", fg="white")
                    else: # Handle Job Sensors
                        # --- DEFINITIVELY FLIPPED LOGIC ---
                        if self.input_state[i] == 0:
                            label.config(text="Active", bg="#27ae60", fg="white")
                        else:
                            label.config(text="Inactive", bg="#7f8c8d", fg="white")
        except (IndexError, KeyError) as e:
            print(f"Error updating I/O status display: {e}")
        
    def _build_job_settings_tab(self, jobs_tab):
        for w in jobs_tab.winfo_children(): w.destroy()
        
        canvas = tk.Canvas(jobs_tab)
        scrollbar = ttk.Scrollbar(jobs_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        row = 0
        self._job_entries = {}
        for key in list(self.CONFIG["JOBS"].keys()):
            jcfg = self.CONFIG["JOBS"][key]
            ttk.Label(scrollable_frame, text=key, font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(6,2))
            dn_var = tk.StringVar(value=jcfg.get("display_name", key))
            pn_var = tk.StringVar(value=jcfg["part_number"])
            db_var = tk.IntVar(value=jcfg["default_batch"])
            lb_var = tk.StringVar(value=jcfg["lightburn_file"])
            fx_var = tk.DoubleVar(value=float(jcfg.get("focus_height", 0.0)))
            sp_var = tk.StringVar(value=str(jcfg.get("select_pattern", "")))
            jd_var = tk.IntVar(value=int(jcfg.get("job_delay_sec", 0)))
            ab_var = tk.IntVar(value=int(jcfg.get("air_before_sec", 0)))
            aa_var = tk.IntVar(value=int(jcfg.get("air_after_sec", 0)))
            color_var = tk.StringVar(value=self.CONFIG["UI"]["job_button_colors"].get(key, "#bdc3c7"))
            ttk.Label(scrollable_frame, text="Button Text:").grid(row=row+1, column=0, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=dn_var, width=20).grid(row=row+1, column=1, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Part Number:").grid(row=row+1, column=2, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=pn_var, width=20).grid(row=row+1, column=3, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Default Batch:").grid(row=row+1, column=4, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=db_var, width=8).grid(row=row+1, column=5, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text=".lbrn2 File:").grid(row=row+2, column=0, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=lb_var, width=76).grid(row=row+2, column=1, columnspan=4, sticky="we", padx=6)
            def browse_file(var=lb_var):
                p = filedialog.askopenfilename(title="Select LightBurn file", filetypes=[("LightBurn","*.lbrn2"),("All files","*.*")])
                if p: var.set(p)
            ttk.Button(scrollable_frame, text="Browse…", command=browse_file).grid(row=row+2, column=5, sticky="w")
            ttk.Label(scrollable_frame, text="Focus Height (X):").grid(row=row+3, column=0, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=fx_var, width=10).grid(row=row+3, column=1, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Select Pattern (001):").grid(row=row+3, column=2, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=sp_var, width=10).grid(row=row+3, column=3, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Button Color:").grid(row=row+3, column=4, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=color_var, width=10).grid(row=row+3, column=5, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Job Delay (sec):").grid(row=row+4, column=0, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=jd_var, width=8).grid(row=row+4, column=1, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Air Before (sec):").grid(row=row+4, column=2, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=ab_var, width=8).grid(row=row+4, column=3, sticky="w", padx=6)
            ttk.Label(scrollable_frame, text="Air After (sec):").grid(row=row+4, column=4, sticky="e")
            ttk.Entry(scrollable_frame, textvariable=aa_var, width=8).grid(row=row+4, column=5, sticky="w", padx=6)
            ttk.Separator(scrollable_frame, orient="horizontal").grid(row=row+5, column=0, columnspan=7, sticky="ew", pady=8)
            self._job_entries[key] = (dn_var, pn_var, db_var, lb_var, fx_var, sp_var, jd_var, ab_var, aa_var, color_var)
            row += 6
        
        ttk.Button(jobs_tab, text="Save", command=lambda: self._save_jobs_settings(jobs_tab)).pack(side="bottom", pady=10)

    def _save_jobs_settings(self, jobs_tab):
        for key, tpl in self._job_entries.items():
            dn, pn, db, lb, fx, sp, jd, ab, aa, color = tpl
            self.CONFIG["JOBS"][key]["display_name"] = dn.get().strip() or key
            self.CONFIG["JOBS"][key]["part_number"] = pn.get().strip()
            try: self.CONFIG["JOBS"][key]["default_batch"] = max(1, int(db.get()))
            except Exception: self.CONFIG["JOBS"][key]["default_batch"] = 1
            self.CONFIG["JOBS"][key]["lightburn_file"] = lb.get().strip()
            try: self.CONFIG["JOBS"][key]["focus_height"] = float(fx.get())
            except Exception: self.CONFIG["JOBS"][key]["focus_height"] = 0.0
            self.CONFIG["JOBS"][key]["select_pattern"] = sp.get().strip() or ""
            try: self.CONFIG["JOBS"][key]["job_delay_sec"] = int(jd.get())
            except Exception: self.CONFIG["JOBS"][key]["job_delay_sec"] = 0
            try: self.CONFIG["JOBS"][key]["air_before_sec"] = int(ab.get())
            except Exception: self.CONFIG["JOBS"][key]["air_before_sec"] = 0
            try: self.CONFIG["JOBS"][key]["air_after_sec"] = int(aa.get())
            except Exception: self.CONFIG["JOBS"][key]["air_after_sec"] = 0
            self.CONFIG["UI"]["job_button_colors"][key] = color.get().strip() or "#bdc3c7"
        if self.save_config():
            messagebox.showinfo("Saved", "Job settings saved.", parent=jobs_tab)
            self.refresh_preview_upnext_and_lb()
            self._rebuild_job_buttons()
            self._apply_job_visibility_for_pin_selection()
            
    def _build_system_settings_tab(self, sys_tab):
        for w in sys_tab.winfo_children(): w.destroy()
        self._root_var = tk.StringVar(value=self.CONFIG["ROOT"])
        self._open_lb_var = tk.BooleanVar(value=self.CONFIG["OPEN_LB_FILE_ON_START"])
        ttk.Label(sys_tab, text="Root Folder:").grid(row=0, column=0, sticky="e")
        ttk.Entry(sys_tab, textvariable=self._root_var, width=60).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(sys_tab, text="Browse…", command=lambda: self._root_var.set(filedialog.askdirectory(title="Pick root folder") or self._root_var.get())).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(sys_tab, text="Open LightBurn file automatically", variable=self._open_lb_var).grid(row=1, column=1, sticky="w", pady=8)
        ttk.Button(sys_tab, text="Save", command=lambda: self._save_system_settings(sys_tab)).grid(row=2, column=1, sticky="e", pady=10)

    def _save_system_settings(self, sys_tab):
        self.CONFIG["ROOT"] = (self._root_var.get().strip() or self.CONFIG["ROOT"]).rstrip("\\/")
        self.CONFIG["OPEN_LB_FILE_ON_START"] = bool(self._open_lb_var.get())
        if self.save_config():
            messagebox.showinfo("Saved", "System settings saved.", parent=sys_tab)
            self.DIRS, self.LB_BATCH, self.WORKING_BATCH, self.WORKING_COMPLETED_TODAY = self.derive_paths(); self.ensure_dirs()
            
    def _build_lightburn_settings_tab(self, lb_tab):
        for w in lb_tab.winfo_children(): w.destroy()
        self._lb_exe_var = tk.StringVar(value=self.CONFIG["LIGHTBURN"].get("exe_path",""))
        self._lb_start_delay = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("start_delay_sec",2)))
        self._lb_post_delay = tk.IntVar(value=int(self.CONFIG["LIGHTBURN"].get("post_open_delay_sec",3)))
        ttk.Label(lb_tab, text="LightBurn EXE (optional):").grid(row=0, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_exe_var, width=70).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(lb_tab, text="Browse…", command=lambda: self._lb_exe_var.set(filedialog.askopenfilename(title="Select LightBurn.exe", filetypes=[("EXE","*.exe"),("All files","*.*")]) or self._lb_exe_var.get())).grid(row=0, column=2, sticky="w")
        ttk.Label(lb_tab, text="Open after (sec):").grid(row=1, column=0, sticky="e")
        ttk.Entry(lb_tab, textvariable=self._lb_start_delay, width=10).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Button(lb_tab, text="Save", command=lambda: self._save_lightburn_settings(lb_tab)).grid(row=2, column=1, sticky="e", pady=10)

    def _save_lightburn_settings(self, lb_tab):
        self.CONFIG["LIGHTBURN"]["exe_path"] = self._lb_exe_var.get().strip()
        self.CONFIG["LIGHTBURN"]["start_delay_sec"] = int(self._lb_start_delay.get())
        self.CONFIG["LIGHTBURN"]["post_open_delay_sec"] = int(self._lb_post_delay.get())
        if self.save_config():
            messagebox.showinfo("Saved", "LightBurn settings saved.", parent=lb_tab)

    def _build_serial_settings_tab(self, ser_tab):
        for w in ser_tab.winfo_children(): w.destroy()
        
        # Define variables
        self._ser_en_var = tk.BooleanVar(value=self.CONFIG["SERIAL"]["enabled"])
        self._sim_en_var = tk.BooleanVar(value=self.CONFIG["SIMULATE"]["enabled"])
        self._batch_done_var = tk.BooleanVar(value=self.CONFIG["SIMULATE"]["batch_done_on_done"])
        self._ser_port_var = tk.StringVar(value=self.CONFIG["SERIAL"]["port"])
        self._ser_baud_var = tk.IntVar(value=int(self.CONFIG["SERIAL"]["baud"]))
        self._auto_connect_var = tk.BooleanVar(value=self.CONFIG["SERIAL"].get("auto_connect", True)) # NEW

        # Layout widgets
        f1 = ttk.Frame(ser_tab); f1.pack(fill="x", pady=2)
        ttk.Checkbutton(f1, text="Enable serial (ESP32 pedal/stop/AF)", variable=self._ser_en_var).pack(side="left")
        
        f_auto = ttk.Frame(ser_tab); f_auto.pack(fill="x", pady=2)
        ttk.Checkbutton(f_auto, text="Auto-connect to last used port on startup", variable=self._auto_connect_var).pack(side="left") # NEW

        f2 = ttk.Frame(ser_tab); f2.pack(fill="x", pady=2)
        ttk.Checkbutton(f2, text="Simulate Mode (no hardware)", variable=self._sim_en_var).pack(side="left")
        ttk.Checkbutton(f2, text="Treat DONE as 'whole batch complete'", variable=self._batch_done_var).pack(side="left", padx=12)
        
        f3 = ttk.Frame(ser_tab); f3.pack(fill="x", pady=10)
        ttk.Label(f3, text="Port:").pack(side="left")
        ttk.Entry(f3, textvariable=self._ser_port_var, width=12).pack(side="left", padx=6)
        ttk.Label(f3, text="Baud:").pack(side="left")
        ttk.Entry(f3, textvariable=self._ser_baud_var, width=10).pack(side="left", padx=6)
        
        ttk.Button(ser_tab, text="Save", command=lambda: self._save_serial_settings(ser_tab)).pack(side="right", pady=10)

    def _save_serial_settings(self, ser_tab):
        self.CONFIG["SERIAL"]["enabled"] = bool(self._ser_en_var.get())
        self.CONFIG["SERIAL"]["port"] = self._ser_port_var.get().strip() or self.CONFIG["SERIAL"]["port"]
        self.CONFIG["SERIAL"]["auto_connect"] = bool(self._auto_connect_var.get()) # NEW
        try: self.CONFIG["SERIAL"]["baud"] = int(self._ser_baud_var.get())
        except Exception: pass
        self.CONFIG["SIMULATE"]["enabled"] = bool(self._sim_en_var.get())
        self.CONFIG["SIMULATE"]["batch_done_on_done"] = bool(self._batch_done_var.get())
        
        if self.save_config():
            messagebox.showinfo("Saved", "Serial and Simulate settings saved.", parent=ser_tab)
            if self.CONFIG["SIMULATE"]["enabled"]:
                self.update_conn_pill("SIMULATE", warn=True)
            elif self.CONFIG["SERIAL"]["enabled"]:
                self.serial.enabled = True
                self.serial.baud = self.CONFIG["SERIAL"]["baud"]
            else:
                self.serial.enabled = False; self.update_conn_pill("Disabled", warn=True)
            self.update_af_button_state()

    def update_af_button_state(self):
        if hasattr(self, 'af_btn'):
            if self.af_enabled:
                self.af_btn.config(state="normal")
            else:
                self.af_btn.config(state="disabled")

    def _build_io_settings_tab(self, io_tab, parent_window):
        # This function remains the same as the previous version
        for w in io_tab.winfo_children(): w.destroy()
        main_frame = ttk.Frame(io_tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        bottom_frame = ttk.Frame(io_tab)
        bottom_frame.pack(fill="x", side="bottom")
        row = 0
        ttk.Label(main_frame, text="Relay Names / Modes", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0,6))
        row += 1
        self._relay_vars = []
        num_relays = len(self.CONFIG["RELAYS"]["names"])
        if "disabled" not in self.CONFIG["RELAYS"] or len(self.CONFIG["RELAYS"]["disabled"]) != num_relays:
            self.CONFIG["RELAYS"]["disabled"] = [False] * num_relays
        for i in range(num_relays):
            ttk.Label(main_frame, text=f"Relay {i+1}").grid(row=row, column=0, sticky="e")
            nv = tk.StringVar(value=self.CONFIG["RELAYS"]["names"][i])
            ttk.Entry(main_frame, textvariable=nv, width=18).grid(row=row, column=1, sticky="w", padx=6)
            mv = tk.StringVar(value=self.CONFIG["RELAYS"]["modes"][i])
            ttk.Combobox(main_frame, textvariable=mv, width=10, values=["pulse", "switch"]).grid(row=row, column=2, sticky="w")
            pv = tk.IntVar(value=int(self.CONFIG["RELAYS"]["pulse_ms"][i]))
            ttk.Label(main_frame, text="Pulse (ms)").grid(row=row, column=3, sticky="e")
            ttk.Entry(main_frame, textvariable=pv, width=8).grid(row=row, column=4, sticky="w", padx=6)
            pinv = tk.IntVar(value=int(self.CONFIG["RELAYS"]["pins"][i]))
            ttk.Label(main_frame, text="Pin").grid(row=row, column=5, sticky="e")
            ttk.Entry(main_frame, textvariable=pinv, width=8).grid(row=row, column=6, sticky="w", padx=6)
            test_frame = ttk.Frame(main_frame)
            test_frame.grid(row=row, column=7, sticky="w", padx=6)
            ttk.Button(test_frame, text="Test", command=lambda idx=i: self._test_relay(idx)).pack(side="left")
            dv = tk.BooleanVar(value=self.CONFIG["RELAYS"]["disabled"][i])
            ttk.Checkbutton(test_frame, text="Disable", variable=dv, command=lambda idx=i, v=dv: self._toggle_relay_disable(idx, v)).pack(side="left", padx=(6,0))
            self._relay_vars.append((nv, mv, pv, pinv, dv))
            row += 1
        ttk.Separator(main_frame, orient="horizontal").grid(row=row, column=0, columnspan=10, sticky="ew", pady=12)
        row += 1
        ttk.Label(main_frame, text="Input Names / Pins", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0,6))
        row += 1
        self._input_vars = []
        for i in range(len(self.CONFIG["INPUTS"]["names"])):
            ttk.Label(main_frame, text=f"Input {i+1}").grid(row=row, column=0, sticky="e")
            nv = tk.StringVar(value=self.CONFIG["INPUTS"]["names"][i])
            ttk.Entry(main_frame, textvariable=nv, width=22).grid(row=row, column=1, sticky="w", padx=6)
            pinv = tk.IntVar(value=int(self.CONFIG["INPUTS"]["pins"][i]))
            ttk.Label(main_frame, text="Pin").grid(row=row, column=2, sticky="e")
            ttk.Entry(main_frame, textvariable=pinv, width=8).grid(row=row, column=3, sticky="w", padx=6)
            self._input_vars.append((nv, pinv))
            row += 1
        ttk.Separator(main_frame, orient="horizontal").grid(row=row, column=0, columnspan=10, sticky="ew", pady=12)
        row += 1
        status_frame = ttk.LabelFrame(main_frame, text="Live Input Status", padding=10)
        status_frame.grid(row=row, column=0, columnspan=8, sticky="nsew", pady=(10, 10))
        self._io_status_labels.clear()
        status_row, status_col = 0, 0
        for i, name in enumerate(self.CONFIG["INPUTS"]["names"]):
            ttk.Label(status_frame, text=f"{name}:").grid(row=status_row, column=status_col, sticky="e", padx=(0, 5), pady=2)
            status_label = tk.Label(status_frame, text="--", width=10, bg="#bdc3c7", fg="white", relief="sunken", padx=5, pady=2, font=("Segoe UI", 9))
            status_label.grid(row=status_row, column=status_col + 1, sticky="w", padx=(0, 20))
            self._io_status_labels[i] = status_label
            status_col += 2
            if status_col >= 6:
                status_col = 0
                status_row += 1
        ttk.Button(bottom_frame, text="Save I/O Config", command=lambda: self._save_io_settings(io_tab)).pack(side="right", pady=10, padx=10)
        parent_window.bind("<Destroy>", self._stop_io_status_updates, add="+")
        self._start_io_status_updates()

    def _toggle_relay_disable(self, index, var):
        self.CONFIG["RELAYS"]["disabled"][index] = var.get()

    def _save_io_settings(self, io_tab):
        self.CONFIG["RELAYS"]["names"] = [v[0].get().strip() for v in self._relay_vars]
        self.CONFIG["RELAYS"]["modes"] = [v[1].get() for v in self._relay_vars]
        self.CONFIG["RELAYS"]["pulse_ms"] = [int(v[2].get()) for v in self._relay_vars]
        self.CONFIG["RELAYS"]["pins"] = [int(v[3].get()) for v in self._relay_vars]
        self.CONFIG["INPUTS"]["names"] = [v[0].get().strip() for v in self._input_vars]
        self.CONFIG["INPUTS"]["pins"] = [int(v[1].get()) for v in self._input_vars]
        
        if self.save_config():
            messagebox.showinfo("Saved", "I/O settings saved. Restart GUI for pin changes to take full effect.", parent=io_tab)

    def _build_paths_settings_tab(self, paths_tab):
        for w in paths_tab.winfo_children(): w.destroy()
        
        self._hist_en_var = tk.BooleanVar(value=self.CONFIG["HISTORY"].get("enabled", True))
        self._af_disabled_var = tk.BooleanVar(value=not self.CONFIG.get("AUTOFOCUS", {}).get("enabled", True))
        
        ttk.Label(paths_tab, text="File Paths:").grid(row=0, column=0, sticky="w", pady=(0,6), columnspan=3)
        self._nb_path_var = tk.StringVar(value=self.CONFIG.get("FILE_PATHS", {}).get("next_batch_path", ""))
        self._ct_path_var = tk.StringVar(value=self.CONFIG.get("FILE_PATHS", {}).get("completed_today_path", ""))
        self._eh_path_var = tk.StringVar(value=self.CONFIG.get("FILE_PATHS", {}).get("entire_history_path", ""))

        ttk.Label(paths_tab, text="Next Batch Path:").grid(row=1, column=0, sticky="e")
        ttk.Entry(paths_tab, textvariable=self._nb_path_var, width=80).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Button(paths_tab, text="Browse…", command=lambda: self._nb_path_var.set(filedialog.asksaveasfilename(defaultextension=".csv") or self._nb_path_var.get())).grid(row=1, column=2)

        ttk.Label(paths_tab, text="Completed Today Path:").grid(row=2, column=0, sticky="e")
        ttk.Entry(paths_tab, textvariable=self._ct_path_var, width=80).grid(row=2, column=1, sticky="w", padx=6)
        ttk.Button(paths_tab, text="Browse…", command=lambda: self._ct_path_var.set(filedialog.asksaveasfilename(defaultextension=".csv") or self._ct_path_var.get())).grid(row=2, column=2)

        ttk.Label(paths_tab, text="Entire History Path:").grid(row=3, column=0, sticky="e")
        ttk.Entry(paths_tab, textvariable=self._eh_path_var, width=80).grid(row=3, column=1, sticky="w", padx=6)
        ttk.Button(paths_tab, text="Browse…", command=lambda: self._eh_path_var.set(filedialog.asksaveasfilename(defaultextension=".csv") or self._eh_path_var.get())).grid(row=3, column=2)

        ttk.Checkbutton(paths_tab, text="Enable History Log (saves to Entire History Path)", variable=self._hist_en_var).grid(row=4, column=1, sticky="w", pady=10)
        ttk.Checkbutton(paths_tab, text="Disable Autofocus", variable=self._af_disabled_var, command=self.update_af_button_state).grid(row=5, column=1, sticky="w", pady=10)
        
        ttk.Button(paths_tab, text="Save", command=lambda: self._save_paths_settings(paths_tab)).grid(row=6, column=1, sticky="e", pady=10)

    def _save_paths_settings(self, paths_tab):
        self.CONFIG.setdefault("HISTORY", {})
        self.CONFIG["HISTORY"]["enabled"] = bool(self._hist_en_var.get())
        
        self.CONFIG.setdefault("FILE_PATHS", {})
        self.CONFIG["FILE_PATHS"]["next_batch_path"] = self._nb_path_var.get().strip()
        self.CONFIG["FILE_PATHS"]["completed_today_path"] = self._ct_path_var.get().strip()
        self.CONFIG["FILE_PATHS"]["entire_history_path"] = self._eh_path_var.get().strip()
        
        self.CONFIG.setdefault("AUTOFOCUS", {})
        self.CONFIG["AUTOFOCUS"]["enabled"] = not bool(self._af_disabled_var.get())
        self.af_enabled = self.CONFIG["AUTOFOCUS"]["enabled"]

        if self.save_config():
            messagebox.showinfo("Saved", "File Paths settings saved.", parent=paths_tab)
            self.DIRS, self.LB_BATCH, self.WORKING_BATCH, self.WORKING_COMPLETED_TODAY = self.derive_paths()
            self.update_af_button_state()

    def _test_relay(self, index):
        name = self.CONFIG["RELAYS"]["names"][index]
        mode = self.CONFIG["RELAYS"]["modes"][index]
        
        if self.CONFIG["RELAYS"]["disabled"][index]:
            self.status_var.set(f"Warning: {name} is disabled in settings.")
            return

        if self.CONFIG["SIMULATE"]["enabled"]:
            if mode == "pulse":
                ms = int(self.CONFIG["RELAYS"]["pulse_ms"][index])
                self.pulse_relay_by_name(name, ms)
                self.status_var.set(f"Simulating {name}: Pulse for {ms}ms.")
            elif mode == "switch":
                current_state = self.relay_state[index]
                new_state = 1 - current_state
                self.set_relay_by_name(name, new_state)
                self.status_var.set(f"Simulating {name}: Toggled to {'ON' if new_state else 'OFF'}.")
        else:
            if mode == "pulse":
                ms = int(self.CONFIG["RELAYS"]["pulse_ms"][index])
                ok = self.serial.pulse(index, ms)
                self.status_var.set(f"Testing {name}: Sent PULSE {index} {ms}." if ok else f"Failed to send PULSE to {name}.")
            elif mode == "switch":
                current_state = self.relay_state[index]
                new_state = 1 - current_state
                ok = self.serial.relay_set(index, new_state)
                if ok:
                    self.relay_state[index] = new_state
                self.status_var.set(f"Testing {name}: Sent RSET {index} {new_state}." if ok else f"Failed to send RSET to {name}.")

    def poll_serial(self):
        if self.CONFIG["SERIAL"]["enabled"] and not self.CONFIG["SIMULATE"]["enabled"]:
            line = self.serial.readline()
            if line: self.on_serial_line(line)
        self.after(self.CONFIG["SERIAL"]["poll_ms"], self.poll_serial)

    def sim_toggle_relay(self, i):
        ms = int(self.CONFIG["RELAYS"]["pulse_ms"][i])
        name = self.CONFIG["RELAYS"]["names"][i]
        
        if self.CONFIG["RELAYS"]["disabled"][i]:
             self.status_var.set(f"Warning: {name} is disabled in settings.")
             return

        if self.CONFIG["SIMULATE"]["enabled"]:
            if ms > 0:
                self.relay_state[i] = 1
                self.after(ms, lambda: self._sim_relay_off_after_pulse(i))
                self.status_var.set(f"Simulated pulse on {name} for {ms}ms")
            else:
                self.relay_state[i] = 1 - self.relay_state[i]
                self.status_var.set(f"Simulated toggle on {name} to {'ON' if self.relay_state[i] else 'OFF'}")
        else:
            if ms > 0:
                ok = self.serial.pulse(i, ms)
            else:
                ok = self.serial.relay_toggle(i)
            self.status_var.set(f"Sent {'pulse' if ms>0 else 'toggle'} to {name}." if ok else f"Failed to send to {name}.")

    def _sim_relay_off_after_pulse(self, i):
        self.relay_state[i] = 0

    def sim_set_input(self, i, value):
        name = self.CONFIG["INPUTS"]["names"][i]
        self.input_override[i] = value
        val = 1 if value else 0
        self.input_state[i] = val
        nl = name.lower()
        if "door" in nl or "sensor" in nl:
            self.door_closed = bool(val)
            self.update_door_ui()
        
        self._maybe_select_job_from_input_pattern()
        
        if self.CONFIG["SIMULATE"]["enabled"]:
            self.status_var.set(f"Simulated {name} to {'ON' if value else 'OFF'}")
        else:
            ok = self.serial.sim_input(i, val)
            self.status_var.set(f"Sent sim input to {name}." if ok else f"Failed to send sim input to {name}.")

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

    def _get_current_input_pattern(self):
        pattern = ""
        for i in range(3):
            pin_name = f"Job Sensor {i+1}"
            try:
                pin_index = [n.lower() for n in self.CONFIG["INPUTS"]["names"]].index(pin_name.lower())
                inverted_state = 1 - self.input_state[pin_index] 
                pattern += str(inverted_state)
            except (ValueError, IndexError):
                return None 
        return pattern

    def _check_for_job_select_on_start(self):
        self._maybe_select_job_from_input_pattern()
        self.after(200, self._check_for_job_select_on_start)

    def _maybe_select_job_from_input_pattern(self):
        current_pattern = self._get_current_input_pattern()
        current_time = int(time.time() * 1000)

        if current_pattern != self.last_job_pattern:
            self.last_stable_job_time = current_time
            self.last_job_pattern = current_pattern
            return
        
        if current_time - self.last_stable_job_time < self.job_cooldown_ms:
            return 

        if current_pattern is None or current_pattern == "000":
            if self.selected_job.get():
                self.selected_job.set("")
                self.last_selected_job = ""
                self.sel_job_lbl.config(text="--")
                self.part_var.set("")
                self.date_var.set("")
                self.next_var.set("")
                self.batch_var.set("")
                self.upnext_list.delete(0, tk.END)
                self.left_hdr.config(text="Up Next (Preview / Current Batch) [0]")
                self.write_lightburn_batch([])
                self.status_var.set("Idle. Waiting for job sensor input.")
            self._apply_job_visibility_for_pin_selection(active_key=None)
            return

        found_key = None
        for key, jcfg in self.CONFIG["JOBS"].items():
            sp = jcfg.get("select_pattern")
            if sp == current_pattern:
                found_key = key
                break
        
        if found_key:
            if self.selected_job.get() != found_key:
                self.on_job_clicked(found_key)
            self._apply_job_visibility_for_pin_selection(active_key=found_key)
        else:
            if self.selected_job.get() != "":
                self.selected_job.set("")
                self.last_selected_job = ""
                self.status_var.set(f"Warning: Input pattern '{current_pattern}' does not match any configured job. Cleared job selection.")
            self._apply_job_visibility_for_pin_selection(active_key=None)


    def _apply_job_visibility_for_pin_selection(self, active_key=None):
        for btn in self.job_btns.values():
            try: btn.grid_remove()
            except Exception: pass
            
        if active_key and active_key in self.job_btns:
            try: 
                self.job_btns[active_key].grid()
            except Exception: pass
        
        if active_key:
            self.jobs_frame.config(text=f"Active Job: {self.CONFIG['JOBS'][active_key]['display_name']}")
        else:
            self.jobs_frame.config(text="Active Job: (None Selected)")


    def set_relay_by_name(self, name, val):
        try:
            i = self.CONFIG["RELAYS"]["names"].index(name)
            
            if self.CONFIG["RELAYS"]["disabled"][i]: return

            if self.CONFIG["SIMULATE"]["enabled"]:
                self.relay_state[i] = val
            else:
                ok = self.serial.relay_set(i, val)
                if ok:
                    self.relay_state[i] = val
        except ValueError:
            pass

    def pulse_relay_by_name(self, name, ms):
        try:
            i = self.CONFIG["RELAYS"]["names"].index(name)

            if self.CONFIG["RELAYS"]["disabled"][i]: return

            if self.CONFIG["SIMULATE"]["enabled"]:
                self.relay_state[i] = 1
                self.after(ms, lambda: self._sim_relay_off_after_pulse(i))
            else:
                ok = self.serial.pulse(i, ms)
                if ok:
                    self.relay_state[i] = 1
        except ValueError:
            pass

class SerialHelper:
    def __init__(self, cfg, status_cb, line_cb, master):
        self.enabled = bool(cfg.get("enabled", False))
        self.baud = int(cfg.get("baud", 115200))
        self._status_cb = status_cb; self._line_cb = line_cb
        self.ser = None
        self.master = master

    def _update_status(self, text, ok=False, warn=False):
        if self._status_cb: self._status_cb(text, ok=ok, warn=warn)

    def try_open_manual(self, port_name):
        self.close()
        
        if not self.enabled:
             self._update_status("Disabled (in Settings)", warn=True)
             return False
        
        try:
            self.ser = serial.Serial(port_name, self.baud, timeout=0.5)
            
            print(f"Connecting to {port_name}. Waiting 2 seconds for ESP32 initialization...")
            time.sleep(2) 
            
            self.ser.reset_input_buffer() 
            print("Serial input buffer flushed.")
            
            if self.ser.is_open:
                self._update_status(f"Connected ({port_name})", ok=True)
                return True
            return False
            
        except Exception as e:
            self.ser = None
            self._update_status("Disconnected", warn=False)
            messagebox.showerror("Serial Connection Error", f"Could not open serial port {port_name}: {e}")
            return False

    def close(self):
        try:
            if self.ser and self.ser.is_open: 
                self.ser.close()
                self.ser = None
        except Exception: pass
        self._update_status("Disconnected", warn=True)

    def _write(self, data: bytes):
        try:
            if self.ser and self.ser.is_open:  
                self.ser.write(data);  
                return True
            else:
                self._update_status("Not Connected", warn=False)
                return False
        except Exception:
            self.close()
            return False

    # NEW method for the handshake
    def get_all_states(self): return self._write(b"GETSTATE\n")
    def relay_set(self, i, v): return self._write(f"RSET {int(i)} {1 if v else 0}\n".encode())
    def relay_toggle(self, i): return self._write(f"RTGL {int(i)}\n".encode())
    def pulse(self, i, ms): return self._write(f"PULSE {int(i)} {int(ms)}\n".encode())
    def sim_input(self, i, v): return self._write(f"SIMI {int(i)} {1 if v else 0}\n".encode())

    def readline(self):
        if not self.enabled or not self.ser or not self.ser.is_open: return ""
        try:
            line = self.ser.readline().decode(errors="ignore").strip()
            if line and self._line_cb: self._line_cb(line)
            return line
        except serial.SerialException:
            self.close()
            return ""
        except Exception:
            return ""

if __name__ == "__main__":
    try:
        try:
            import serial
            import serial.tools.list_ports
        except ImportError:
            print("ERROR: 'pyserial' library not found.")
            print("Please install it using: pip install pyserial")
            sys.exit(1)

        App().mainloop()
    except KeyboardInterrupt:
        sys.exit(0)