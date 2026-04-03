import re
import socket
import threading
import time
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
import sys
import os

from PIL import Image, ImageDraw
import pystray
import requests

TEM_STATUS_URL = "http://192.168.0.1/onsystemui/"
TEM_DATA_URL = "http://192.168.0.1/onsystemui/api/microscope/data/"
WEBSWITCH_BASE_URL = os.environ.get("LN2_WEBSWITCH_BASE_URL", "").rstrip("/")
POLL_INTERVAL_SECONDS = 30
REFILL_POLL_INTERVAL_SECONDS = 2
REFILL_THRESHOLD = 10.0
REFILL_TARGET = 95.0
REFILL_COUNTDOWN_SECONDS = 10
REFILL_READ_TIMEOUT_SECONDS = 10
SINGLE_INSTANCE_HOST = "127.0.0.1"
SINGLE_INSTANCE_PORT = 48527
SINGLE_INSTANCE_MESSAGE = b"SHOW_LN2_MONITOR"


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


def parse_nitrogen_level(html_text: str) -> float | None:
    m = re.search(
        r"<td[^>]*class=[\"']value[\"'][^>]*id=[\"']nitrogen-level[\"'][^>]*>([^<]*)</td>",
        html_text,
        re.IGNORECASE,
    )
    if not m:
        return None
    raw = m.group(1).strip()
    raw = raw.replace("%", "").strip()
    try:
        return float(raw)
    except ValueError:
        return None


def fetch_nitrogen_level() -> tuple[float | None, str | None]:
    try:
        resp = requests.get(TEM_DATA_URL, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        n2 = data.get("n2", {})
        if not n2.get("levelAvailable"):
            return (
                None,
                f"Nitrogen level unavailable: {n2.get('levelStatus', 'UNKNOWN')}",
            )

        raw_level = n2.get("level")
        if raw_level is None:
            return None, "Nitrogen level missing from microscope data"

        level = float(raw_level)
        if level <= 1:
            level *= 100
        return level, None
    except Exception as e:
        return None, str(e)


def trigger_webswitch_relay(relay: int, on: bool) -> tuple[bool, str | None]:
    try:
        if not WEBSWITCH_BASE_URL:
            return (
                False,
                "LN2_WEBSWITCH_BASE_URL is not set. Configure the WebSwitch URL before starting refills.",
            )
        command = 1 if on else 0
        url = f"{WEBSWITCH_BASE_URL}/state.xml?relay{relay}State={command}"
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        state_text = root.findtext(f"relay{relay}state")
        expected = str(command)
        if state_text != expected:
            return (
                False,
                f"Relay {relay} did not reach expected state {expected}; device reported {state_text!r}",
            )
        return True, None
    except Exception as e:
        return False, str(e)


class LN2MonitorApp:
    BACKGROUND = "#07111f"
    PANEL = "#0f1c2e"
    PANEL_ALT = "#16263c"
    PANEL_HIGHLIGHT = "#1d3350"
    TEXT = "#f4f7fb"
    MUTED = "#94a8c4"
    ACCENT = "#5eead4"
    ACCENT_SOFT = "#173a3d"
    WARNING = "#f59e0b"
    WARNING_SOFT = "#3d2a0b"
    DANGER = "#f97373"
    DANGER_SOFT = "#3c1d24"
    SUCCESS = "#34d399"
    SUCCESS_SOFT = "#12352f"
    BUTTON = "#1f6feb"
    BUTTON_ACTIVE = "#3b82f6"
    ABORT_BUTTON = "#7f1d2d"
    ABORT_BUTTON_ACTIVE = "#9f1239"
    ABORT_TEXT = "#fff7f8"

    def __init__(self, master: tk.Tk):
        self.master = master
        master.title("LN2 Control Center")
        self._configure_window()
        self._configure_styles()

        self.current_level_var = tk.StringVar(value="--")
        self.status_var = tk.StringVar(value="Starting...")
        self.level_caption_var = tk.StringVar(value="Waiting for first reading")
        self.status_badge_var = tk.StringVar(value="BOOTING")
        self.cryo_mode = tk.BooleanVar(value=False)
        self.refill_target_var = tk.StringVar(value=f"{REFILL_TARGET:.1f}")
        self.confirmed_refill_target = REFILL_TARGET
        self.latest_level = None
        self.latest_level_error = "No LN2 reading available yet."
        self.refill_in_progress = False
        self.refill_is_active = False
        self.active_refill_target = REFILL_TARGET
        self.refill_popup = None
        self.refill_popup_label = None
        self.refill_popup_level_var = tk.StringVar(value="--")
        self.level_meter_var = tk.DoubleVar(value=0.0)
        self.tray_icon = None
        self.tray_visible = False
        self.is_quitting = False
        self.instance_socket = None

        self._build_ui()
        self.refill_target_var.trace_add("write", self._on_refill_target_edit)
        self._on_refill_target_edit()
        self._update_visual_state()
        self._initialize_tray()
        self._start_instance_listener()
        self.master.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

        self.abort_event = threading.Event()

        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()

    def _configure_window(self):
        self.master.geometry("860x580")
        self.master.minsize(760, 540)
        self.master.configure(bg=self.BACKGROUND)
        icon_path = resource_path("assets/app_icon.png")
        if icon_path.exists():
            self.window_icon = tk.PhotoImage(file=str(icon_path))
            self.master.iconphoto(True, self.window_icon)

    def _initialize_tray(self):
        image = self._create_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("Show LN2 Monitor", self._tray_show_window, default=True),
            pystray.MenuItem("Hide Window", self._tray_hide_window),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self._tray_exit_application),
        )
        self.tray_icon = pystray.Icon("ln2-monitor", image, "LN2 Monitor", menu)
        self.tray_icon.run_detached()
        self.tray_visible = True

    def _start_instance_listener(self):
        self.instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.instance_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.instance_socket.bind((SINGLE_INSTANCE_HOST, SINGLE_INSTANCE_PORT))
        self.instance_socket.listen(1)
        threading.Thread(target=self._instance_listener_loop, daemon=True).start()

    def _instance_listener_loop(self):
        while not self.is_quitting:
            try:
                conn, _addr = self.instance_socket.accept()
            except OSError:
                break

            with conn:
                try:
                    message = conn.recv(1024)
                except OSError:
                    continue

                if message == SINGLE_INSTANCE_MESSAGE:
                    self.master.after(0, self._show_from_tray)

    def _create_tray_image(self) -> Image.Image:
        icon_path = resource_path("assets/app_icon.png")
        if icon_path.exists():
            return Image.open(icon_path)
        return self._create_fallback_tray_image()

    def _create_fallback_tray_image(self) -> Image.Image:
        image = Image.new("RGB", (64, 64), self.BACKGROUND)
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((6, 6, 58, 58), radius=14, fill=self.PANEL)
        draw.rectangle((18, 14, 46, 52), fill=self.PANEL_HIGHLIGHT)
        draw.rectangle((22, 18, 42, 48), fill=self.ACCENT_SOFT)
        draw.rectangle((22, 38, 42, 48), fill=self.ACCENT)
        draw.rectangle((26, 10, 38, 18), fill=self.TEXT)
        return image

    def _configure_styles(self):
        self.style = ttk.Style(self.master)
        self.style.theme_use("clam")

        self.style.configure(".", background=self.BACKGROUND, foreground=self.TEXT)
        self.style.configure("App.TFrame", background=self.BACKGROUND)
        self.style.configure("Panel.TFrame", background=self.PANEL)
        self.style.configure("PanelAlt.TFrame", background=self.PANEL_ALT)
        self.style.configure(
            "Eyebrow.TLabel",
            background=self.BACKGROUND,
            foreground=self.ACCENT,
            font=("Avenir Next", 11, "bold"),
        )
        self.style.configure(
            "Title.TLabel",
            background=self.BACKGROUND,
            foreground=self.TEXT,
            font=("Avenir Next", 28, "bold"),
        )
        self.style.configure(
            "Subtitle.TLabel",
            background=self.BACKGROUND,
            foreground=self.MUTED,
            font=("Avenir Next", 11),
        )
        self.style.configure(
            "SectionLabel.TLabel",
            background=self.PANEL,
            foreground=self.MUTED,
            font=("Avenir Next", 10, "bold"),
        )
        self.style.configure(
            "CardValue.TLabel",
            background=self.PANEL,
            foreground=self.TEXT,
            font=("Avenir Next", 36, "bold"),
        )
        self.style.configure(
            "CardHint.TLabel",
            background=self.PANEL,
            foreground=self.MUTED,
            font=("Avenir Next", 11),
        )
        self.style.configure(
            "StatusText.TLabel",
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            font=("Avenir Next", 12),
        )
        self.style.configure(
            "StatusBadge.TLabel",
            background=self.PANEL_HIGHLIGHT,
            foreground=self.TEXT,
            font=("Avenir Next", 10, "bold"),
            padding=(12, 6),
        )
        self.style.configure(
            "FormLabel.TLabel",
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            font=("Avenir Next", 11, "bold"),
        )
        self.style.configure(
            "FormHelp.TLabel",
            background=self.PANEL_ALT,
            foreground=self.MUTED,
            font=("Avenir Next", 10),
        )
        self.style.configure(
            "Action.TButton",
            background=self.BUTTON,
            foreground=self.TEXT,
            font=("Avenir Next", 11, "bold"),
            borderwidth=0,
            focuscolor=self.BUTTON,
            padding=(16, 12),
        )
        self.style.map(
            "Action.TButton",
            background=[("active", self.BUTTON_ACTIVE), ("disabled", "#29415d")],
            foreground=[("disabled", "#7d93af")],
        )
        self.style.configure(
            "Secondary.TButton",
            background=self.PANEL_HIGHLIGHT,
            foreground=self.TEXT,
            font=("Avenir Next", 11, "bold"),
            borderwidth=0,
            focuscolor=self.PANEL_HIGHLIGHT,
            padding=(16, 12),
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", "#2c4b73"), ("disabled", "#26364c")],
            foreground=[("disabled", "#7d93af")],
        )
        self.style.configure(
            "PressedSecondary.TButton",
            background=self.ACCENT_SOFT,
            foreground="#d4fffa",
            font=("Avenir Next", 11, "bold"),
            borderwidth=0,
            focuscolor=self.ACCENT_SOFT,
            padding=(16, 12),
        )
        self.style.map(
            "PressedSecondary.TButton",
            background=[("active", "#215255"), ("disabled", "#214447")],
            foreground=[("disabled", "#8fbab5")],
        )
        self.style.configure(
            "Danger.TButton",
            background="#2f1b24",
            foreground="#ffe5e7",
            font=("Avenir Next", 11, "bold"),
            borderwidth=0,
            focuscolor="#2f1b24",
            padding=(16, 12),
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", "#4b2230"), ("disabled", "#30242a")],
            foreground=[("disabled", "#b18a92")],
        )
        self.style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=self.PANEL_HIGHLIGHT,
            background=self.ACCENT,
            bordercolor=self.PANEL_HIGHLIGHT,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT,
            thickness=18,
        )
        self.style.configure(
            "Modern.TEntry",
            fieldbackground="#0c1728",
            foreground=self.TEXT,
            insertcolor=self.TEXT,
            bordercolor=self.PANEL_HIGHLIGHT,
            lightcolor=self.PANEL_HIGHLIGHT,
            darkcolor=self.PANEL_HIGHLIGHT,
            padding=8,
        )

    def _build_ui(self):
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(self.master, style="App.TFrame", padding=(28, 24, 28, 12))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ttk.Label(header, text="LN2 MONITOR", style="Eyebrow.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header, text="Talos LN2 Control Center", style="Title.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 4)
        )
        ttk.Label(
            header,
            text="Live nitrogen monitoring, refill controls, and cryo-cycle.",
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, sticky="w")

        content = ttk.Frame(self.master, style="App.TFrame", padding=(28, 8, 28, 34))
        content.grid(row=1, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(0, weight=1)

        left_panel = ttk.Frame(content, style="Panel.TFrame", padding=24)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left_panel.grid_columnconfigure(0, weight=1)

        ttk.Label(left_panel, text="CURRENT LEVEL", style="SectionLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            left_panel, textvariable=self.current_level_var, style="CardValue.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            left_panel, textvariable=self.level_caption_var, style="CardHint.TLabel"
        ).grid(row=2, column=0, sticky="w", pady=(4, 16))

        self.level_progress = ttk.Progressbar(
            left_panel,
            style="Modern.Horizontal.TProgressbar",
            variable=self.level_meter_var,
            maximum=100,
            mode="determinate",
        )
        self.level_progress.grid(row=3, column=0, sticky="ew")

        self.level_scale = tk.Canvas(
            left_panel,
            bg=self.PANEL,
            highlightthickness=0,
            height=28,
            bd=0,
        )
        self.level_scale.grid(row=4, column=0, sticky="ew", pady=(8, 24))
        self.level_scale.bind("<Configure>", lambda _event: self._draw_level_scale())

        status_card = ttk.Frame(left_panel, style="PanelAlt.TFrame", padding=18)
        status_card.grid(row=5, column=0, sticky="ew")
        status_card.grid_columnconfigure(0, weight=1)
        ttk.Label(status_card, text="SYSTEM STATUS", style="FormHelp.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.status_badge = ttk.Label(
            status_card, textvariable=self.status_badge_var, style="StatusBadge.TLabel"
        )
        self.status_badge.grid(row=0, column=1, sticky="e")
        ttk.Label(
            status_card,
            textvariable=self.status_var,
            style="StatusText.TLabel",
            wraplength=420,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(14, 0))

        right_panel = ttk.Frame(content, style="PanelAlt.TFrame", padding=24)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.grid_columnconfigure(0, weight=1)

        ttk.Label(right_panel, text="REFILL SETTINGS", style="FormLabel.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            right_panel,
            text="Choose the confirmed refill target before automatic or manual activation.",
            style="FormHelp.TLabel",
            wraplength=260,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 18))

        target_row = ttk.Frame(right_panel, style="PanelAlt.TFrame")
        target_row.grid(row=2, column=0, sticky="ew")
        target_row.grid_columnconfigure(0, weight=1)

        self.refill_target_entry = ttk.Entry(
            target_row,
            textvariable=self.refill_target_var,
            width=8,
            style="Modern.TEntry",
        )
        self.refill_target_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.refill_target_entry.bind("<Return>", self._on_refill_target_enter)
        self.confirm_target_button = ttk.Button(
            target_row,
            text="Confirm Target",
            command=self.confirm_refill_target,
            style="Secondary.TButton",
        )
        self.confirm_target_button.grid(row=0, column=1, sticky="e")

        ttk.Label(
            right_panel,
            text="Allowed range: 30% to 100%",
            style="FormHelp.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(10, 24))

        self.cryo_button = ttk.Button(
            right_panel,
            text="Enter Cryo-Cycle",
            command=self.toggle_cryo_cycle,
            style="Secondary.TButton",
        )
        self.cryo_button.grid(row=4, column=0, sticky="ew", pady=(0, 10))

        self.force_button = ttk.Button(
            right_panel,
            text="Force Refill Now",
            command=self.schedule_refill,
            style="Action.TButton",
        )
        self.force_button.grid(row=5, column=0, sticky="ew", pady=(0, 10))

        ttk.Button(
            right_panel,
            text="Hide To Tray",
            command=self._hide_to_tray,
            style="Danger.TButton",
        ).grid(row=6, column=0, sticky="ew", pady=(0, 6))

        ttk.Label(
            right_panel,
            text="Closing the window keeps monitoring in the system tray.",
            style="FormHelp.TLabel",
            wraplength=260,
            justify="left",
        ).grid(row=7, column=0, sticky="w", pady=(8, 0))

    def _draw_level_scale(self):
        canvas = self.level_scale
        canvas.delete("all")
        width = max(canvas.winfo_width(), 240)
        start_x = 8
        end_x = width - 8
        y = 10
        canvas.create_line(start_x, y, end_x, y, fill=self.PANEL_HIGHLIGHT, width=2)
        for value in (0, 25, 50, 75, 100):
            x = start_x + (end_x - start_x) * (value / 100)
            canvas.create_line(x, y - 5, x, y + 5, fill=self.MUTED, width=1)
            canvas.create_text(
                x,
                y + 14,
                text=str(value),
                fill=self.MUTED,
                font=("Avenir Next", 9),
            )

    def _set_status(self, message: str):
        self.status_var.set(message)
        self._update_visual_state()

    def _hide_to_tray(self):
        if self.is_quitting:
            return
        self.master.withdraw()
        self._set_status("Monitoring in background. Use the tray icon to reopen.")

    def _show_from_tray(self):
        if self.is_quitting:
            return
        self.master.deiconify()
        self.master.lift()
        self.master.focus_force()
        self._set_status("Monitoring...")

    def _tray_show_window(self, _icon=None, _item=None):
        self.master.after(0, self._show_from_tray)

    def _tray_hide_window(self, _icon=None, _item=None):
        self.master.after(0, self._hide_to_tray)

    def _tray_exit_application(self, _icon=None, _item=None):
        self.master.after(0, self._exit_application)

    def _exit_application(self):
        if self.is_quitting:
            return
        self.is_quitting = True
        if self.instance_socket is not None:
            try:
                self.instance_socket.close()
            except OSError:
                pass
            self.instance_socket = None
        if self.tray_icon is not None:
            self.tray_icon.stop()
            self.tray_icon = None
        self.master.destroy()

    def _refresh_control_states(self):
        cryo_active = self.cryo_mode.get()
        self.cryo_button.configure(
            style="PressedSecondary.TButton" if cryo_active else "Secondary.TButton"
        )
        force_disabled = cryo_active
        self.force_button.configure(state="disabled" if force_disabled else "normal")

    def _update_visual_state(self):
        level_text = self.current_level_var.get()
        badge_text = "MONITORING"
        badge_bg = self.PANEL_HIGHLIGHT
        badge_fg = self.TEXT
        fill_color = self.ACCENT

        if level_text == "ERR":
            badge_text = "ERROR"
            badge_bg = self.DANGER_SOFT
            badge_fg = "#ffd7db"
            fill_color = self.DANGER
            self.level_caption_var.set("Connection or parsing issue")
            self.level_meter_var.set(0)
        else:
            try:
                level_value = float(level_text.replace("%", ""))
            except ValueError:
                level_value = None

            if level_value is not None:
                self.level_meter_var.set(max(0, min(level_value, 100)))
                if level_value <= REFILL_THRESHOLD:
                    fill_color = self.WARNING
                    self.level_caption_var.set("Below threshold, refill watch active")
                elif level_value < 35:
                    fill_color = self.WARNING
                    self.level_caption_var.set("Reserve is getting low")
                elif level_value < 75:
                    fill_color = self.ACCENT
                    self.level_caption_var.set("Within normal operating band")
                else:
                    fill_color = self.SUCCESS
                    self.level_caption_var.set("Tank level is comfortably healthy")

        if self.refill_is_active:
            badge_text = "REFILLING"
            badge_bg = self.SUCCESS_SOFT
            badge_fg = "#d8fff2"
            fill_color = self.SUCCESS
        elif self.refill_in_progress:
            badge_text = "COUNTDOWN"
            badge_bg = self.WARNING_SOFT
            badge_fg = "#ffe6b4"
            fill_color = self.WARNING
        elif self.cryo_mode.get():
            badge_text = "CRYO MODE"
            badge_bg = self.ACCENT_SOFT
            badge_fg = "#d4fffa"

        self.status_badge_var.set(badge_text)
        self.style.configure(
            "StatusBadge.TLabel", background=badge_bg, foreground=badge_fg
        )
        self.style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=self.PANEL_HIGHLIGHT,
            background=fill_color,
            bordercolor=self.PANEL_HIGHLIGHT,
            lightcolor=fill_color,
            darkcolor=fill_color,
        )
        self._refresh_control_states()

    def toggle_cryo_cycle(self):
        self.cryo_mode.set(not self.cryo_mode.get())
        if self.cryo_mode.get():
            self.cryo_button.config(text="Exit Cryo-Cycle")
            self._set_status("Cryo-cycle ON: monitoring only, no refill.")
        else:
            self.cryo_button.config(text="Enter Cryo-Cycle")
            self._set_status("Cryo-cycle OFF: normal refill behavior restored.")
        self._update_visual_state()

    def _get_refill_target(self) -> float | None:
        raw_value = self.refill_target_var.get().strip()
        try:
            target = float(raw_value)
        except ValueError:
            messagebox.showerror(
                "Invalid Refill Target",
                "Please enter a numeric refill target between 30 and 100.",
            )
            self.refill_target_entry.focus_set()
            return None

        if not 30 <= target <= 100:
            messagebox.showerror(
                "Invalid Refill Target",
                "Refill target must be between 30% and 100%.",
            )
            self.refill_target_entry.focus_set()
            return None

        self.refill_target_var.set(f"{target:.1f}")
        return target

    def _on_refill_target_edit(self, *_args):
        current_text = self.refill_target_var.get().strip()
        confirmed_text = f"{self.confirmed_refill_target:.1f}"
        if current_text == confirmed_text:
            self.confirm_target_button.config(state="disabled")
        else:
            self.confirm_target_button.config(state="normal")

    def _on_refill_target_enter(self, _event):
        self.confirm_refill_target()

    def confirm_refill_target(self):
        if self.refill_in_progress:
            self._set_status(
                "Cannot change the confirmed refill target during an active refill."
            )
            return

        target = self._get_refill_target()
        if target is None:
            return

        self.confirmed_refill_target = target
        self.confirm_target_button.config(state="disabled")
        self._set_status(f"Confirmed refill target: {target:.1f}%.")

    def schedule_refill(self):
        if self.refill_in_progress:
            return
        if self.cryo_mode.get():
            self._set_status("Cryo-cycle is active: refill suppressed.")
            return
        if self.latest_level is None:
            self._set_status(
                f"Refill aborted because no current LN2 reading is available: {self.latest_level_error}"
            )
            return

        target = self.confirmed_refill_target

        self.refill_in_progress = True
        self.active_refill_target = target
        self.abort_event.clear()

        msg = f"Nitrogen <= {REFILL_THRESHOLD}%: refill to {target:.1f}% starting in {REFILL_COUNTDOWN_SECONDS} seconds."
        self._set_status(msg)

        self._show_popup_countdown(REFILL_COUNTDOWN_SECONDS, target)

    def _show_popup_countdown(self, seconds_left: int, target: float):
        popup = tk.Toplevel(self.master)
        popup.title("Refill Warning")
        popup.geometry("420x250")
        popup.minsize(420, 250)
        popup.configure(bg=self.PANEL)
        self.refill_popup = popup

        container = tk.Frame(popup, bg=self.PANEL, padx=22, pady=22)
        container.pack(fill="both", expand=True)

        tk.Label(
            container,
            text="Automatic Refill Pending",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Avenir Next", 18, "bold"),
        ).pack(anchor="w")

        label = tk.Label(
            container,
            text=f"Refill to {target:.1f}% will start in {seconds_left} seconds.",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Avenir Next", 12),
            wraplength=360,
            justify="left",
        )
        label.pack(anchor="w", pady=(8, 16))
        self.refill_popup_label = label

        tk.Label(
            container,
            text="Current nitrogen level",
            bg=self.PANEL,
            fg=self.MUTED,
            font=("Avenir Next", 10, "bold"),
        ).pack(anchor="w")
        tk.Label(
            container,
            textvariable=self.refill_popup_level_var,
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Avenir Next", 22, "bold"),
        ).pack(anchor="w", pady=(2, 18))

        abort_btn = tk.Label(
            container,
            text="Abort Refill",
            bg=self.ABORT_BUTTON,
            fg=self.ABORT_TEXT,
            padx=18,
            pady=10,
            font=("Avenir Next", 11, "bold"),
            cursor="hand2",
        )
        abort_btn.pack(anchor="w")
        abort_btn.bind("<Button-1>", lambda _event: self._abort_refill(popup))
        abort_btn.bind(
            "<Enter>",
            lambda _event: abort_btn.config(bg=self.ABORT_BUTTON_ACTIVE),
        )
        abort_btn.bind(
            "<Leave>",
            lambda _event: abort_btn.config(bg=self.ABORT_BUTTON),
        )

        def tick():
            nonlocal seconds_left
            if self.abort_event.is_set():
                self._close_refill_popup()
                return
            if seconds_left <= 0:
                if self.latest_level is None:
                    self.refill_in_progress = False
                    self._close_refill_popup()
                    self._set_status(
                        f"Refill aborted because no current LN2 reading is available: {self.latest_level_error}"
                    )
                    return
                if self.refill_popup_label is not None:
                    self.refill_popup_label.config(
                        text=f"Refill is active until {target:.1f}%."
                    )
                self._start_refill_thread(target)
                return
            seconds_left -= 1
            label.config(
                text=f"Refill to {target:.1f}% will start in {seconds_left} seconds."
            )
            popup.after(1000, tick)

        popup.protocol("WM_DELETE_WINDOW", lambda: None)
        popup.after(1000, tick)

    def _abort_refill(self, popup=None):
        self.abort_event.set()
        if popup is not None:
            self._close_refill_popup()
        if self.refill_is_active:
            self._set_status("Abort requested. Switching relay OFF...")
        else:
            self.refill_in_progress = False
            self._set_status("Refill aborted by user.")

    def _start_refill_thread(self, target: float):
        self.refill_is_active = True
        self.active_refill_target = target
        self._set_status(f"Performing refill to {target:.1f}%: switching relay ON")
        threading.Thread(
            target=self._execute_refill, args=(target,), daemon=True
        ).start()

    def _execute_refill(self, target: float):
        if self.cryo_mode.get():
            self.master.after(
                0,
                lambda: self._finish_refill(
                    "Cryo-cycle active, aborting forced refill."
                ),
            )
            return

        ok, err = trigger_webswitch_relay(relay=1, on=True)

        if not ok:
            self.master.after(
                0, lambda: self._finish_refill(f"Refill failed at ON command: {err}")
            )
            return

        missing_level_since = None
        while True:
            if self.abort_event.is_set():
                break

            level, err = fetch_nitrogen_level()
            if level is None:
                if missing_level_since is None:
                    missing_level_since = time.time()
                elif time.time() - missing_level_since > REFILL_READ_TIMEOUT_SECONDS:
                    ok, off_err = trigger_webswitch_relay(relay=1, on=False)
                    if ok:
                        self.master.after(
                            0,
                            lambda e=err: self._finish_refill(
                                f"Refill stopped because nitrogen level was unavailable for more than {REFILL_READ_TIMEOUT_SECONDS} seconds: {e}"
                            ),
                        )
                    else:
                        self.master.after(
                            0,
                            lambda e=off_err: self._finish_refill(
                                f"Refill stopped and relay OFF failed: {e}"
                            ),
                        )
                    return
            else:
                missing_level_since = None

            if level is not None and level >= target:
                break

            time.sleep(REFILL_POLL_INTERVAL_SECONDS)

        ok, err = trigger_webswitch_relay(relay=1, on=False)
        if not ok:
            self.master.after(
                0,
                lambda: self._finish_refill(
                    f"Refill completed with errors at OFF command: {err}"
                ),
            )
            return

        if self.abort_event.is_set():
            self.master.after(
                0,
                lambda: self._finish_refill(
                    "Refill aborted by user. Relay switched OFF."
                ),
            )
        else:
            self.master.after(
                0,
                lambda: self._finish_refill(
                    f"Refill complete. Nitrogen level reached {target:.1f}%."
                ),
            )

    def _finish_refill(self, message: str):
        self.refill_is_active = False
        self.refill_in_progress = False
        self.active_refill_target = self.confirmed_refill_target
        self._set_status(message)
        self._close_refill_popup()

    def _close_refill_popup(self):
        if self.refill_popup is not None and self.refill_popup.winfo_exists():
            self.refill_popup.destroy()
        self.refill_popup = None
        self.refill_popup_label = None

    def _apply_level_update(self, level: float | None, err: str | None):
        if level is not None:
            self.latest_level = level
            self.latest_level_error = ""
            formatted_level = f"{level:.1f}%"
            self.current_level_var.set(formatted_level)
            if self.refill_in_progress:
                self.refill_popup_level_var.set(formatted_level)
                if self.refill_is_active:
                    self._set_status(
                        f"Refilling to {self.active_refill_target:.1f}%... current nitrogen level {formatted_level}"
                    )
            if (
                level <= REFILL_THRESHOLD
                and not self.cryo_mode.get()
                and not self.refill_in_progress
            ):
                self.schedule_refill()
            elif not self.refill_in_progress:
                self._set_status("Monitoring...")
        else:
            self.latest_level = None
            self.latest_level_error = err or "Unknown LN2 read error."
            self.current_level_var.set("ERR")
            if self.refill_in_progress:
                self.refill_popup_level_var.set("ERR")
            self._set_status(f"Error reading level: {err}")
        self._update_visual_state()

    def _monitor_loop(self):
        while True:
            level, err = fetch_nitrogen_level()
            self.master.after(0, lambda l=level, e=err: self._apply_level_update(l, e))
            poll_interval = (
                REFILL_POLL_INTERVAL_SECONDS
                if self.refill_in_progress
                else POLL_INTERVAL_SECONDS
            )
            time.sleep(poll_interval)


def main():
    try:
        with socket.create_connection(
            (SINGLE_INSTANCE_HOST, SINGLE_INSTANCE_PORT), timeout=0.5
        ) as client:
            client.sendall(SINGLE_INSTANCE_MESSAGE)
        return
    except OSError:
        pass

    root = tk.Tk()
    app = LN2MonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
