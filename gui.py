"""
Did You Close the Speaker? — GUI Application
System tray icon with a compact control window.
"""

import sys
import os
import asyncio
import logging
import json
import time
import threading
from pathlib import Path
from enum import Enum

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QSystemTrayIcon, QMenu, QAction, QMessageBox,
    QCheckBox, QFrame, QComboBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon, QColor, QPixmap, QPainter

import platform
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    import ctypes
    import ctypes.wintypes
    WM_POWERBROADCAST = 0x0218
    WM_QUERYENDSESSION = 0x0011
    WM_ENDSESSION = 0x0016
    PBT_APMRESUMEAUTOMATIC = 0x0012
    PBT_APMRESUMESUSPEND = 0x0007
    PBT_APMSUSPEND = 0x0004

    # For idle detection via GetLastInputInfo
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_uint),
            ("dwTime", ctypes.c_uint),
        ]

from config import load_config
from tapo_control import TapoController
from power import shutdown_windows, sleep_windows, restart_windows

# ---------------------------------------------------------------------------
# Audio output detection (pycaw / WASAPI peak meter)
# ---------------------------------------------------------------------------
try:
    from pycaw.pycaw import AudioUtilities, IAudioMeterInformation, IMMDevice
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    PYCAW_AVAILABLE = True
except ImportError:
    PYCAW_AVAILABLE = False

# Threshold for considering audio as "playing".
# Values below this are treated as silence (e.g. noise floor, idle loopback).
_AUDIO_PEAK_THRESHOLD = 0.001

# Cache the meter interface — GetSpeakers() + Activate() is expensive to repeat every second.
_audio_meter_cache: object | None = None


def _get_audio_meter():
    """Get or create cached IAudioMeterInformation interface."""
    global _audio_meter_cache
    if _audio_meter_cache is None and PYCAW_AVAILABLE:
        try:
            speakers = AudioUtilities.GetSpeakers()
            imm_device = speakers._dev.QueryInterface(IMMDevice)
            interface = imm_device.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
            _audio_meter_cache = cast(interface, POINTER(IAudioMeterInformation))
        except Exception:
            _audio_meter_cache = None
    return _audio_meter_cache


def is_audio_playing() -> bool:
    """
    Check if audio is actually being output by reading the system peak meter.
    Returns True if the peak level exceeds the silence threshold.
    Returns False if pycaw is not available, on error, or silence.
    """
    if not PYCAW_AVAILABLE:
        return False
    try:
        meter = _get_audio_meter()
        if meter is None:
            return False
        peak = meter.GetPeakValue()
        return peak > _AUDIO_PEAK_THRESHOLD
    except Exception:
        # Meter may become invalid after audio device change — reset cache
        global _audio_meter_cache
        _audio_meter_cache = None
        return False


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).parent
LOG_DIR = APP_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
SETTINGS_PATH = APP_DIR / "settings.json"
ICON_PATH = APP_DIR / "assets" / "ico.ico"
TRAY_ICON_PATH = APP_DIR / "assets" / "sim_ico.ico"

if IS_WINDOWS:
    STARTUP_DIR = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    SHORTCUT_NAME = "DYCTSpeaker.lnk"
    SHORTCUT_PATH = STARTUP_DIR / SHORTCUT_NAME

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "dycts.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("dycts")


# ---------------------------------------------------------------------------
# Settings persistence
# ---------------------------------------------------------------------------
DEFAULT_SETTINGS = {
    "start_with_windows": False,
    "speaker_on_at_startup": False,
    "watchdog_enabled": False,
    "idle_timer_enabled": False,
    "idle_timer_minutes": 15,
    "idle_auto_on": False,
    "idle_audio_aware": False,
}


def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            return {**DEFAULT_SETTINGS, **saved}
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(settings: dict):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=4)


# ---------------------------------------------------------------------------
# Windows startup shortcut management
# ---------------------------------------------------------------------------
def create_startup_shortcut(speaker_on: bool = False):
    """Create a Windows startup .lnk via PowerShell (no pywin32 needed)."""
    if not IS_WINDOWS:
        return

    if getattr(sys, "frozen", False):
        target_path = sys.executable
        arguments = "--startup"
    else:
        # Use pythonw.exe to avoid console window
        python_dir = Path(sys.executable).parent
        pythonw = python_dir / "pythonw.exe"
        if not pythonw.exists():
            pythonw = Path(sys.executable)  # fallback
        target_path = str(pythonw)
        arguments = f'"{APP_DIR / "gui.py"}" --startup'

    if speaker_on:
        arguments += " --speaker-on"

    ps_script = (
        '$ws = New-Object -ComObject WScript.Shell; '
        f'$sc = $ws.CreateShortcut("{SHORTCUT_PATH}"); '
        f'$sc.TargetPath = "{target_path}"; '
        f"$sc.Arguments = '{arguments}'; "
        f'$sc.WorkingDirectory = "{APP_DIR}"; '
        f'$sc.IconLocation = "{ICON_PATH}"; '
        '$sc.Description = "Did You Close the Speaker?"; '
        '$sc.Save()'
    )
    import subprocess
    subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True)
    logger.info(f"Startup shortcut created: {SHORTCUT_PATH}")


def remove_startup_shortcut():
    if IS_WINDOWS and SHORTCUT_PATH.exists():
        SHORTCUT_PATH.unlink()
        logger.info(f"Startup shortcut removed: {SHORTCUT_PATH}")


# ---------------------------------------------------------------------------
# Async worker thread
# ---------------------------------------------------------------------------
class AsyncWorker(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, coro_func, *args, **kwargs):
        super().__init__()
        self._coro_func = coro_func
        self._args = args
        self._kwargs = kwargs

    def run(self):
        try:
            result = asyncio.run(self._coro_func(*self._args, **self._kwargs))
            self.finished.emit(result)
        except Exception as e:
            logger.error(f"AsyncWorker error: {e}")
            self.error.emit(str(e))


# ---------------------------------------------------------------------------
# Plug status enum
# ---------------------------------------------------------------------------
class PlugState(Enum):
    UNKNOWN = "unknown"
    ON = "on"
    OFF = "off"
    ERROR = "error"


# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
STYLE_SHEET = """
QMainWindow {
    background-color: #1a1a2e;
}
QLabel {
    color: #e0e0e0;
}
QLabel#title {
    font-size: 15px;
    font-weight: bold;
    color: #ffffff;
}
QLabel#sectionLabel {
    color: #888888;
    font-size: 11px;
}
QPushButton {
    font-size: 12px;
    font-weight: 600;
    padding: 10px 16px;
    border: none;
    border-radius: 8px;
    color: #ffffff;
}
QPushButton:hover {
    opacity: 0.9;
}
QPushButton:pressed {
    padding-top: 11px;
}
QPushButton:disabled {
    background-color: #3a3a5c;
    color: #888888;
}
QPushButton#toggleBtn {
    background-color: #0f3460;
    min-width: 100px;
}
QPushButton#toggleBtn:hover {
    background-color: #1a4a7a;
}
QPushButton#shutdownBtn {
    background-color: #c0392b;
}
QPushButton#shutdownBtn:hover {
    background-color: #e74c3c;
}
QPushButton#sleepBtn {
    background-color: #2980b9;
}
QPushButton#sleepBtn:hover {
    background-color: #3498db;
}
QPushButton#restartBtn {
    background-color: #d68910;
}
QPushButton#restartBtn:hover {
    background-color: #f39c12;
}
QPushButton#refreshBtn {
    background-color: transparent;
    color: #aaaaaa;
    font-size: 11px;
    padding: 4px 8px;
}
QPushButton#refreshBtn:hover {
    color: #ffffff;
}
QCheckBox {
    color: #cccccc;
    font-size: 11px;
    spacing: 6px;
}
QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #555577;
    border-radius: 3px;
    background-color: #2a2a4a;
}
QCheckBox::indicator:checked {
    background-color: #2980b9;
    border-color: #3498db;
}
QPushButton#settingsBtn {
    background-color: #2a2a4a;
    color: #cccccc;
    font-size: 12px;
    padding: 8px 16px;
}
QPushButton#settingsBtn:hover {
    background-color: #3a3a5c;
    color: #ffffff;
}
QFrame#separator {
    background-color: #2a2a4a;
}
QMessageBox {
    background-color: #1a1a2e;
    color: #e0e0e0;
}
QMessageBox QLabel {
    color: #e0e0e0;
    font-size: 12px;
}
QMessageBox QPushButton {
    background-color: #0f3460;
    color: #ffffff;
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    min-width: 80px;
}
QMessageBox QPushButton:hover {
    background-color: #1a4a7a;
}
QMessageBox QPushButton:pressed {
    background-color: #0a2b50;
}
"""


def make_circle_icon(color: str, size: int = 64) -> QIcon:
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor(color))
    painter.setPen(Qt.NoPen)
    painter.drawEllipse(4, 4, size - 8, size - 8)
    painter.end()
    return QIcon(pixmap)


def _make_separator() -> QFrame:
    sep = QFrame()
    sep.setObjectName("separator")
    sep.setFixedHeight(1)
    sep.setFrameShape(QFrame.HLine)
    return sep


# ---------------------------------------------------------------------------
# Settings window
# ---------------------------------------------------------------------------


class SettingsWindow(QMainWindow):
    """Separate settings window opened from the main window."""

    def __init__(self, parent: "MainWindow", settings: dict, idle_timer: QTimer):
        super().__init__(parent)
        self._parent = parent
        self.settings = settings          # shared reference — mutations are live
        self._idle_timer = idle_timer
        self._setup_ui()

    # ----- UI -----

    def _setup_ui(self):
        self.setWindowTitle("Settings")
        self.setFixedSize(370, 320)
        self.setStyleSheet(STYLE_SHEET)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        # --- Startup ---
        self.chk_startup = QCheckBox("Run at startup")
        self.chk_startup.setChecked(self.settings["start_with_windows"])
        self.chk_startup.stateChanged.connect(self._on_startup_changed)
        layout.addWidget(self.chk_startup)

        self.chk_startup_speaker = QCheckBox("Run at startup & Speaker ON")
        self.chk_startup_speaker.setChecked(self.settings["speaker_on_at_startup"])
        self.chk_startup_speaker.setEnabled(self.settings["start_with_windows"])
        self.chk_startup_speaker.stateChanged.connect(self._on_startup_speaker_changed)
        layout.addWidget(self.chk_startup_speaker)

        # --- Watchdog ---
        self.chk_watchdog = QCheckBox("Turn off speaker before PC shutdown/restart")
        self.chk_watchdog.setChecked(self.settings["watchdog_enabled"])
        self.chk_watchdog.stateChanged.connect(self._on_watchdog_changed)
        layout.addWidget(self.chk_watchdog)

        layout.addWidget(_make_separator())

        # --- Idle timer row ---
        idle_row = QHBoxLayout()
        idle_row.setSpacing(6)

        self.chk_idle = QCheckBox("Turn off speaker when idle:")
        self.chk_idle.setChecked(self.settings["idle_timer_enabled"])
        self.chk_idle.stateChanged.connect(self._on_idle_changed)
        idle_row.addWidget(self.chk_idle)

        self.combo_idle = QComboBox()
        self.combo_idle.addItems(IDLE_OPTIONS)
        current_val = self.settings["idle_timer_minutes"]
        if current_val in IDLE_VALUES:
            self.combo_idle.setCurrentIndex(IDLE_VALUES.index(current_val))
        self.combo_idle.setEnabled(self.settings["idle_timer_enabled"])
        self.combo_idle.setFixedWidth(70)
        self.combo_idle.setStyleSheet(COMBO_STYLE)
        self.combo_idle.currentIndexChanged.connect(self._on_idle_time_changed)
        idle_row.addWidget(self.combo_idle)
        idle_row.addStretch()
        layout.addLayout(idle_row)

        self.chk_idle_auto_on = QCheckBox("Auto turn on speaker when active")
        self.chk_idle_auto_on.setChecked(self.settings["idle_auto_on"])
        self.chk_idle_auto_on.setEnabled(self.settings["idle_timer_enabled"])
        self.chk_idle_auto_on.stateChanged.connect(self._on_idle_auto_on_changed)
        layout.addWidget(self.chk_idle_auto_on)

        # --- Audio-aware idle ---
        self.chk_audio_aware = QCheckBox("Keep speaker on while audio is playing")
        self.chk_audio_aware.setChecked(self.settings["idle_audio_aware"])
        self.chk_audio_aware.setEnabled(self.settings["idle_timer_enabled"])
        self.chk_audio_aware.stateChanged.connect(self._on_audio_aware_changed)
        if not PYCAW_AVAILABLE:
            self.chk_audio_aware.setEnabled(False)
            self.chk_audio_aware.setToolTip("Requires pycaw library (pip install pycaw)")
        layout.addWidget(self.chk_audio_aware)

        layout.addStretch()

        # --- Close button ---
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(80)
        close_btn.clicked.connect(self.close)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    # ----- Handlers -----

    def _on_startup_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["start_with_windows"] = enabled
        self.chk_startup_speaker.setEnabled(enabled)
        if not enabled:
            self.chk_startup_speaker.setChecked(False)
            self.settings["speaker_on_at_startup"] = False
            remove_startup_shortcut()
        else:
            create_startup_shortcut(speaker_on=self.settings["speaker_on_at_startup"])
        save_settings(self.settings)

    def _on_startup_speaker_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["speaker_on_at_startup"] = enabled
        if self.settings["start_with_windows"]:
            create_startup_shortcut(speaker_on=enabled)
        save_settings(self.settings)

    def _on_watchdog_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["watchdog_enabled"] = enabled
        save_settings(self.settings)
        logger.info(f"Watchdog {'enabled' if enabled else 'disabled'}")

    def _on_idle_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["idle_timer_enabled"] = enabled
        self.combo_idle.setEnabled(enabled)
        self.chk_idle_auto_on.setEnabled(enabled)
        # Audio-aware checkbox: enabled only when idle timer is on AND pycaw is available
        self.chk_audio_aware.setEnabled(enabled and PYCAW_AVAILABLE)
        if enabled:
            self._idle_timer.start(1_000)
        else:
            self._idle_timer.stop()
        save_settings(self.settings)
        logger.info(f"Idle timer {'enabled' if enabled else 'disabled'}")

    def _on_idle_time_changed(self, index):
        if 0 <= index < len(IDLE_VALUES):
            self.settings["idle_timer_minutes"] = IDLE_VALUES[index]
            save_settings(self.settings)
            logger.info(f"Idle timer set to {IDLE_VALUES[index]} minutes")

    def _on_idle_auto_on_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["idle_auto_on"] = enabled
        save_settings(self.settings)

    def _on_audio_aware_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["idle_audio_aware"] = enabled
        save_settings(self.settings)
        logger.info(f"Audio-aware idle {'enabled' if enabled else 'disabled'}")

    # ----- Close -----

    def closeEvent(self, event):
        self._parent._settings_window = None
        event.accept()


# ---------------------------------------------------------------------------
# Settings window
# ---------------------------------------------------------------------------
COMBO_STYLE = (
    "QComboBox { background-color: #2a2a4a; color: #cccccc; border: 1px solid #555577; "
    "border-radius: 3px; padding: 2px 6px; font-size: 11px; }"
    "QComboBox::drop-down { border: none; }"
    "QComboBox QAbstractItemView { background-color: #2a2a4a; color: #cccccc; "
    "selection-background-color: #2980b9; }"
)

IDLE_OPTIONS = ["1 min", "3 min", "5 min", "10 min", "15 min", "20 min", "25 min", "30 min", "45 min", "60 min"]
IDLE_VALUES  = [1, 3, 5, 10, 15, 20, 25, 30, 45, 60]


# ---------------------------------------------------------------------------
# Watchdog shutdown helper — runs turn-off in a background thread
# ---------------------------------------------------------------------------
WATCHDOG_TIMEOUT = 4  # seconds — allow headroom for Tapo auth + network variance


class _WatchdogThread(threading.Thread):
    """
    Fires turn-off in a daemon thread so that nativeEvent can return
    immediately to the OS.  Use the built-in Event to let WM_ENDSESSION
    know whether the work finished.
    """

    def __init__(self, controller: TapoController):
        super().__init__(daemon=True)
        self._controller = controller
        self.done = threading.Event()
        self.success = False

    def run(self):
        try:
            quick = TapoController(
                ip=self._controller.ip,
                email=self._controller.email,
                password=self._controller.password,
                timeout=WATCHDOG_TIMEOUT,
            )
            asyncio.run(quick.turn_off())
            self.success = True
            logger.info("Watchdog thread: speaker OFF succeeded.")
        except Exception as e:
            logger.error(f"Watchdog thread: speaker OFF failed — {e}")
        finally:
            self.done.set()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, controller: TapoController, config: dict,
                 start_hidden: bool = False, speaker_on_at_startup: bool = False):
        super().__init__()
        self.controller = controller
        self.config = config
        self.settings = load_settings()
        self.plug_state = PlugState.UNKNOWN
        self._workers: list[AsyncWorker] = []
        self._busy = False
        self._shutdown_intercepted = False
        self._idle_turned_off = False  # tracks if idle timer turned off the speaker
        self._audio_was_playing = False  # previous tick's audio state
        self._audio_stop_time: float | None = None  # monotonic time when audio stopped
        self._watchdog_thread: _WatchdogThread | None = None  # background watchdog

        self._setup_ui()
        self._setup_tray()
        self._setup_idle_timer()
        self._setup_show_event_listener()

        # Startup behavior
        if speaker_on_at_startup:
            logger.info("Startup flag: turning speaker ON")
            QTimer.singleShot(500, self._startup_speaker_on)
        else:
            QTimer.singleShot(500, self.refresh_status)

        if start_hidden:
            self.hide()
        else:
            self.show()

    # ----- UI Setup -----

    def _setup_ui(self):
        self.setWindowTitle("Did You Close the Speaker?")
        self.setFixedSize(400, 270)
        self.setStyleSheet(STYLE_SHEET)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        # --- Title ---
        title = QLabel("🔊 Did You Close the Speaker?")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # --- Status row ---
        status_row = QHBoxLayout()
        self.status_label = QLabel("● Status: checking...")
        status_row.addWidget(self.status_label)
        status_row.addStretch()
        self.refresh_btn = QPushButton("↻ Refresh")
        self.refresh_btn.setObjectName("refreshBtn")
        self.refresh_btn.clicked.connect(self.refresh_status)
        status_row.addWidget(self.refresh_btn)
        layout.addLayout(status_row)

        # --- Speaker toggle ---
        self.toggle_btn = QPushButton("Speaker ON / OFF")
        self.toggle_btn.setObjectName("toggleBtn")
        self.toggle_btn.clicked.connect(self.toggle_speaker)
        layout.addWidget(self.toggle_btn)

        layout.addWidget(_make_separator())

        # --- Power actions ---
        power_label = QLabel("Safe Power Actions")
        power_label.setObjectName("sectionLabel")
        power_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(power_label)

        power_row = QHBoxLayout()
        power_row.setSpacing(8)

        self.shutdown_btn = QPushButton("⏻ Shutdown")
        self.shutdown_btn.setObjectName("shutdownBtn")
        self.shutdown_btn.clicked.connect(lambda: self.safe_power_action("shutdown"))

        self.sleep_btn = QPushButton("🌙 Sleep")
        self.sleep_btn.setObjectName("sleepBtn")
        self.sleep_btn.clicked.connect(lambda: self.safe_power_action("sleep"))

        self.restart_btn = QPushButton("🔄 Restart")
        self.restart_btn.setObjectName("restartBtn")
        self.restart_btn.clicked.connect(lambda: self.safe_power_action("restart"))

        power_row.addWidget(self.shutdown_btn)
        power_row.addWidget(self.sleep_btn)
        power_row.addWidget(self.restart_btn)
        layout.addLayout(power_row)

        layout.addWidget(_make_separator())

        # --- Settings button ---
        self._settings_window = None
        settings_btn = QPushButton("⚙️  Settings")
        settings_btn.setObjectName("settingsBtn")
        settings_btn.clicked.connect(self._open_settings)
        layout.addWidget(settings_btn)

        layout.addStretch()

    # ----- Settings window -----

    def _open_settings(self):
        if self._settings_window is None or not self._settings_window.isVisible():
            self._settings_window = SettingsWindow(self, self.settings, self._idle_check_timer)
            self._settings_window.show()
        else:
            self._settings_window.activateWindow()
            self._settings_window.raise_()

    # ----- Idle timer -----

    def _setup_idle_timer(self):
        """
        Poll GetLastInputInfo every second.
        Only runs when idle timer is enabled — zero overhead when off.
        """
        self._idle_check_timer = QTimer(self)
        self._idle_check_timer.timeout.connect(self._check_idle)
        if self.settings.get("idle_timer_enabled", False):
            self._idle_check_timer.start(1_000)

    # ----- Single instance show-window listener -----

    def _setup_show_event_listener(self):
        """
        Create a Named Event and poll it every 500ms.
        When a second instance signals it, show this window.
        Polling cost: negligible (one WaitForSingleObject with 0 timeout).
        """
        if not IS_WINDOWS:
            return
        self._show_event_handle = ctypes.windll.kernel32.CreateEventW(
            None, True, False, EVENT_NAME  # manual reset, initially non-signaled
        )
        self._show_event_timer = QTimer(self)
        self._show_event_timer.timeout.connect(self._check_show_event)
        self._show_event_timer.start(500)

    def _check_show_event(self):
        """Check if another instance signaled us to show."""
        if not IS_WINDOWS or not self._show_event_handle:
            return
        # WaitForSingleObject with 0 timeout = instant check, no blocking
        result = ctypes.windll.kernel32.WaitForSingleObject(self._show_event_handle, 0)
        if result == 0:  # WAIT_OBJECT_0 = signaled
            ctypes.windll.kernel32.ResetEvent(self._show_event_handle)
            logger.info("Show signal received from another instance.")
            self._show_window()

    def _get_idle_seconds(self) -> float:
        """Return seconds since last keyboard/mouse input."""
        if not IS_WINDOWS:
            return 0.0
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
            return millis / 1000.0
        return 0.0

    def _check_idle(self):
        """Called every 1s by QTimer. Checks idle state and acts."""
        if not self.settings.get("idle_timer_enabled", False):
            return

        threshold_sec = self.settings.get("idle_timer_minutes", 15) * 60
        audio_aware = self.settings.get("idle_audio_aware", False)
        audio_playing = audio_aware and is_audio_playing()

        # Track audio state transitions
        if audio_aware:
            if audio_playing and not self._audio_was_playing:
                # Audio just started — clear stop time
                self._audio_stop_time = None
            elif not audio_playing and self._audio_was_playing:
                # Audio just stopped — record the moment
                self._audio_stop_time = time.monotonic()
                logger.info("Audio stopped. Idle timer starts from now.")
            self._audio_was_playing = audio_playing

        # Audio is playing — skip idle entirely
        if audio_playing:
            if self._idle_turned_off:
                # Audio started playing after idle turn-off — auto-on if enabled
                if self.settings.get("idle_auto_on", False):
                    if self.plug_state == PlugState.OFF and not self._busy:
                        logger.info("Audio playing after idle turn-off. Auto turning speaker ON.")
                        self._idle_turned_off = False
                        self._set_busy(True)
                        self.status_label.setText("● Audio detected — turning ON...")
                        self.status_label.setStyleSheet("color: #f39c12;")
                        self._run_async(
                            self.controller.turn_on,
                            on_done=lambda _: self._on_toggle_done(True),
                            on_error=self._on_toggle_error,
                        )
                else:
                    self._idle_turned_off = False
            return

        # Determine effective idle seconds
        input_idle_sec = self._get_idle_seconds()

        if audio_aware and self._audio_stop_time is not None:
            # Use the shorter of: time since last input, time since audio stopped
            audio_idle_sec = time.monotonic() - self._audio_stop_time
            idle_sec = min(input_idle_sec, audio_idle_sec)
        else:
            idle_sec = input_idle_sec

        if threshold_sec - 5 <= idle_sec < threshold_sec and not self._idle_turned_off:
            logger.info(f"Idle check: {idle_sec:.0f}s / {threshold_sec}s")

        if idle_sec >= threshold_sec:
            # User is idle — turn off speaker if it's on
            if self.plug_state == PlugState.ON and not self._idle_turned_off and not self._busy:
                logger.info(f"Idle for {idle_sec:.0f}s (threshold {threshold_sec}s). Turning off speaker.")
                self._idle_turned_off = True
                self._set_busy(True)
                self.status_label.setText("● Idle — turning OFF...")
                self.status_label.setStyleSheet("color: #f39c12;")
                self._run_async(
                    self.controller.turn_off,
                    on_done=lambda _: self._on_idle_off_done(),
                    on_error=lambda e: self._on_idle_off_error(e),
                )
        else:
            # User is active
            if self._idle_turned_off and self.settings.get("idle_auto_on", False):
                if self.plug_state == PlugState.OFF and not self._busy:
                    logger.info("User returned from idle. Auto turning speaker ON.")
                    self._idle_turned_off = False
                    self._set_busy(True)
                    self.status_label.setText("● Resuming — turning ON...")
                    self.status_label.setStyleSheet("color: #f39c12;")
                    self._run_async(
                        self.controller.turn_on,
                        on_done=lambda _: self._on_toggle_done(True),
                        on_error=self._on_toggle_error,
                    )
            elif self._idle_turned_off and not self.settings.get("idle_auto_on", False):
                self._idle_turned_off = False

    def _on_idle_off_done(self):
        self.plug_state = PlugState.OFF
        self.status_label.setText("● Speaker: OFF (idle)")
        self.status_label.setStyleSheet("color: #7f8c8d;")
        self.toggle_btn.setText("Turn Speaker ON")
        self.tray_toggle_action.setText("Speaker ON")
        self._update_tray_icon()
        self._set_busy(False)

    def _on_idle_off_error(self, err_msg):
        logger.error(f"Idle turn-off failed: {err_msg}")
        self._idle_turned_off = False
        self._set_busy(False)

    # ----- System tray -----

    def _setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        if TRAY_ICON_PATH.exists():
            self.tray_icon.setIcon(QIcon(str(TRAY_ICON_PATH)))
        elif ICON_PATH.exists():
            self.tray_icon.setIcon(QIcon(str(ICON_PATH)))
        else:
            self._update_tray_icon()

        tray_menu = QMenu()

        show_action = QAction("Show Window", self)
        show_action.triggered.connect(self._show_window)
        tray_menu.addAction(show_action)
        tray_menu.addSeparator()

        self.tray_toggle_action = QAction("Speaker OFF", self)
        self.tray_toggle_action.triggered.connect(self.toggle_speaker)
        tray_menu.addAction(self.tray_toggle_action)
        tray_menu.addSeparator()

        for label, action in [("Safe Shutdown", "shutdown"), ("Safe Sleep", "sleep"), ("Safe Restart", "restart")]:
            act = QAction(label, self)
            act.triggered.connect(lambda checked, a=action: self.safe_power_action(a))
            tray_menu.addAction(act)

        tray_menu.addSeparator()
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self._quit_app)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self._tray_activated)
        self.tray_icon.setToolTip("Did You Close the Speaker?")
        self.tray_icon.show()

    def _update_tray_icon(self):
        if TRAY_ICON_PATH.exists():
            self.tray_icon.setIcon(QIcon(str(TRAY_ICON_PATH)))
        elif ICON_PATH.exists():
            self.tray_icon.setIcon(QIcon(str(ICON_PATH)))
        else:
            color_map = {
                PlugState.ON: "#27ae60",
                PlugState.OFF: "#7f8c8d",
                PlugState.UNKNOWN: "#f39c12",
                PlugState.ERROR: "#c0392b",
            }
            self.tray_icon.setIcon(make_circle_icon(color_map[self.plug_state]))

    def _tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self._show_window()

    def _show_window(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def _quit_app(self):
        self.tray_icon.hide()
        QApplication.quit()

    # ----- Close to tray -----

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    # ----- Native event handling (wake + shutdown watchdog) -----

    def _start_watchdog_thread(self):
        """Launch a background thread to turn off the speaker."""
        if self._watchdog_thread is not None and self._watchdog_thread.is_alive():
            logger.info("Watchdog thread already running, skipping duplicate launch.")
            return
        self._watchdog_thread = _WatchdogThread(self.controller)
        self._watchdog_thread.start()
        logger.info("Watchdog background thread started.")

    def _wait_for_watchdog(self, timeout: float = 2.0):
        """Block until the watchdog thread finishes or timeout expires."""
        if self._watchdog_thread is not None:
            self._watchdog_thread.done.wait(timeout=timeout)
            if self._watchdog_thread.success:
                logger.info("Watchdog completed successfully before session end.")
            else:
                logger.warning("Watchdog did not complete successfully (timeout or error).")

    def nativeEvent(self, event_type, message):
        if IS_WINDOWS and event_type == b"windows_generic_MSG":
            msg = ctypes.wintypes.MSG.from_address(int(message))

            # ── Wake from sleep/hibernate ──
            if msg.message == WM_POWERBROADCAST:
                if msg.wParam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
                    logger.info("System resumed from sleep/hibernate.")
                    self._set_busy(False)
                    self.plug_state = PlugState.UNKNOWN
                    self._idle_turned_off = False
                    self._audio_was_playing = False
                    self._audio_stop_time = None
                    self._shutdown_intercepted = False
                    QTimer.singleShot(3000, self.refresh_status)

                # ── Entering sleep ──
                elif msg.wParam == PBT_APMSUSPEND:
                    if self.settings.get("watchdog_enabled", False):
                        logger.info("PBT_APMSUSPEND — watchdog turning off speaker.")
                        self._start_watchdog_thread()
                        if self._watchdog_thread is not None:
                            self._watchdog_thread.done.wait(timeout=WATCHDOG_TIMEOUT + 0.5)
                            if self._watchdog_thread.success:
                                logger.info("Speaker OFF before sleep.")
                            else:
                                logger.warning("Speaker OFF before sleep may have failed.")

            # ── Shutdown/restart: phase 1 (QUERYENDSESSION) ──
            # This is the OS's most generous window — block here to buy time.
            elif msg.message == WM_QUERYENDSESSION:
                if self.settings.get("watchdog_enabled", False):
                    logger.info("WM_QUERYENDSESSION — launching watchdog and BLOCKING.")
                    self._shutdown_intercepted = True

                    # Tell Windows we're still cleaning up so it shows
                    # "This app is preventing shutdown" instead of killing us.
                    hwnd = int(self.winId())
                    try:
                        ctypes.windll.user32.ShutdownBlockReasonCreate(
                            hwnd, "Turning off speakers safely..."
                        )
                    except Exception as e:
                        logger.warning(f"ShutdownBlockReasonCreate failed: {e}")

                    self._start_watchdog_thread()

                    # Block until watchdog finishes (or timeout).
                    # This is the ONLY reliable place to stall the shutdown sequence.
                    if self._watchdog_thread is not None:
                        self._watchdog_thread.done.wait(timeout=WATCHDOG_TIMEOUT + 1.0)

                    try:
                        ctypes.windll.user32.ShutdownBlockReasonDestroy(hwnd)
                    except Exception:
                        pass

                return True, 1

            # ── Shutdown/restart: phase 2 (ENDSESSION) ──
            # OS has already decided to shut down. If QUERYENDSESSION did its
            # job, the watchdog is already done. Just wait for stragglers.
            elif msg.message == WM_ENDSESSION:
                if msg.wParam and self.settings.get("watchdog_enabled", False):
                    if self._shutdown_intercepted:
                        # QUERYENDSESSION already started it — just wait for completion
                        if self._watchdog_thread is not None and self._watchdog_thread.is_alive():
                            logger.info("WM_ENDSESSION — watchdog still running, waiting briefly.")
                            self._watchdog_thread.done.wait(timeout=1.0)
                        else:
                            logger.info("WM_ENDSESSION — watchdog already completed.")
                    else:
                        # Edge case: QUERYENDSESSION was never received (rare)
                        logger.info("WM_ENDSESSION — last-chance watchdog launch.")
                        self._start_watchdog_thread()
                        self._wait_for_watchdog(timeout=WATCHDOG_TIMEOUT + 0.5)
                return True, 0

        return super().nativeEvent(event_type, message)

    # Legacy helper kept for reference
    async def _watchdog_turn_off(self):
        """Quick turn-off with aggressive timeout for shutdown scenarios."""
        quick = TapoController(
            ip=self.controller.ip,
            email=self.controller.email,
            password=self.controller.password,
            timeout=WATCHDOG_TIMEOUT,
        )
        await quick.turn_off()

    # ----- Startup speaker on -----

    def _startup_speaker_on(self):
        self._set_busy(True)
        self.status_label.setText("● Turning ON (startup)...")
        self.status_label.setStyleSheet("color: #f39c12;")
        self._run_async(
            self.controller.turn_on,
            on_done=lambda _: self._on_toggle_done(True),
            on_error=self._on_toggle_error,
        )

    # ----- Async helpers -----

    def _run_async(self, coro_func, on_done=None, on_error=None, *args, **kwargs):
        worker = AsyncWorker(coro_func, *args, **kwargs)
        self._workers.append(worker)

        def cleanup():
            if worker in self._workers:
                self._workers.remove(worker)

        if on_done:
            worker.finished.connect(on_done)
        if on_error:
            worker.error.connect(on_error)
        worker.finished.connect(cleanup)
        worker.error.connect(cleanup)
        worker.start()

    def _set_busy(self, busy: bool):
        self._busy = busy
        self.toggle_btn.setEnabled(not busy)
        self.shutdown_btn.setEnabled(not busy)
        self.sleep_btn.setEnabled(not busy)
        self.restart_btn.setEnabled(not busy)
        self.refresh_btn.setEnabled(not busy)

    # ----- Status -----

    def refresh_status(self):
        if self._busy:
            return
        self._set_busy(True)
        self.status_label.setText("● Checking...")
        self.status_label.setStyleSheet("color: #f39c12;")
        self._run_async(
            self.controller.get_status,
            on_done=self._on_status_result,
            on_error=self._on_status_error,
        )

    def _on_status_result(self, info):
        self._set_busy(False)
        if info and info.get("device_on") is not None:
            is_on = info["device_on"]
            self.plug_state = PlugState.ON if is_on else PlugState.OFF
            state_text = "ON" if is_on else "OFF"
            state_color = "#27ae60" if is_on else "#7f8c8d"
            self.status_label.setText(f"● Speaker: {state_text}")
            self.status_label.setStyleSheet(f"color: {state_color};")
            self.toggle_btn.setText("Turn Speaker OFF" if is_on else "Turn Speaker ON")
            self.tray_toggle_action.setText("Speaker OFF" if is_on else "Speaker ON")
        else:
            self._on_status_error("No response from device")
        self._update_tray_icon()

    def _on_status_error(self, err_msg):
        self._set_busy(False)
        self.plug_state = PlugState.ERROR
        self.status_label.setText("● Error: could not reach plug")
        self.status_label.setStyleSheet("color: #c0392b;")
        self._update_tray_icon()

    # ----- Toggle speaker -----

    def toggle_speaker(self):
        if self._busy:
            return
        self._set_busy(True)

        if self.plug_state == PlugState.ON:
            self.status_label.setText("● Turning OFF...")
            self.status_label.setStyleSheet("color: #f39c12;")
            self._run_async(
                self.controller.turn_off,
                on_done=lambda _: self._on_toggle_done(False),
                on_error=self._on_toggle_error,
            )
        else:
            self.status_label.setText("● Turning ON...")
            self.status_label.setStyleSheet("color: #f39c12;")
            self._run_async(
                self.controller.turn_on,
                on_done=lambda _: self._on_toggle_done(True),
                on_error=self._on_toggle_error,
            )

    def _on_toggle_done(self, turned_on: bool):
        self.plug_state = PlugState.ON if turned_on else PlugState.OFF
        state_text = "ON" if turned_on else "OFF"
        state_color = "#27ae60" if turned_on else "#7f8c8d"
        self.status_label.setText(f"● Speaker: {state_text}")
        self.status_label.setStyleSheet(f"color: {state_color};")
        self.toggle_btn.setText("Turn Speaker OFF" if turned_on else "Turn Speaker ON")
        self.tray_toggle_action.setText("Speaker OFF" if turned_on else "Speaker ON")
        self._update_tray_icon()
        self._set_busy(False)

    def _on_toggle_error(self, err_msg):
        self.plug_state = PlugState.ERROR
        self.status_label.setText("● Toggle failed")
        self.status_label.setStyleSheet("color: #c0392b;")
        self._update_tray_icon()
        self._set_busy(False)

    # ----- Safe power actions -----

    def safe_power_action(self, action: str):
        action_labels = {"shutdown": "Shutdown", "sleep": "Sleep", "restart": "Restart"}
        label = action_labels[action]

        reply = QMessageBox.question(
            self, f"Safe {label}",
            f"This will turn off the speaker and {label.lower()} your PC.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self._set_busy(True)
        self.status_label.setText(f"● Turning off speaker for {label.lower()}...")
        self.status_label.setStyleSheet("color: #f39c12;")

        self._run_async(
            self._do_safe_power,
            on_done=lambda _: None,
            on_error=lambda err: self._on_power_error(err, action),
            action=action,
        )

    async def _do_safe_power(self, action: str):
        delay = self.config.get("delay_after_power_off_sec", 2)

        try:
            await self.controller.turn_off()
            logger.info("Speaker plug turned OFF for safe power action.")
        except Exception as e:
            logger.error(f"Failed to turn off speaker: {e}")
            raise RuntimeError(
                f"Could not turn off speaker plug: {e}\n\n"
                f"Power action aborted. Turn off the speaker manually if needed."
            )

        if delay > 0:
            logger.info(f"Waiting {delay}s for speaker to safely power down...")
            await asyncio.sleep(delay)

        action_map = {"shutdown": shutdown_windows, "sleep": sleep_windows, "restart": restart_windows}
        logger.info(f"Executing: {action}")
        action_map[action]()

    def _on_power_error(self, err_msg: str, action: str):
        self._set_busy(False)
        self.plug_state = PlugState.ERROR
        self.status_label.setText("● Power action failed")
        self.status_label.setStyleSheet("color: #c0392b;")
        self._update_tray_icon()
        QMessageBox.warning(self, "Safe Power Action Failed", f"{err_msg}\n\nYour PC was NOT {action}.")


# ---------------------------------------------------------------------------
# Single instance enforcement
# ---------------------------------------------------------------------------
MUTEX_NAME = "Global\\DYCTS_DidYouCloseTheSpeaker"
EVENT_NAME = "Global\\DYCTS_ShowWindow"


def _signal_existing_instance():
    """Signal the existing instance to show its window via a Named Event."""
    if not IS_WINDOWS:
        return
    handle = ctypes.windll.kernel32.OpenEventW(0x0002, False, EVENT_NAME)  # EVENT_MODIFY_STATE
    if handle:
        ctypes.windll.kernel32.SetEvent(handle)
        ctypes.windll.kernel32.CloseHandle(handle)


def _acquire_mutex():
    """Try to acquire a named mutex. Returns handle if success, None if already running."""
    if not IS_WINDOWS:
        return True
    handle = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    last_err = ctypes.windll.kernel32.GetLastError()
    if last_err == 183:  # ERROR_ALREADY_EXISTS
        ctypes.windll.kernel32.CloseHandle(handle)
        return None
    return handle


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--startup", action="store_true", help="Launched from Windows startup")
    parser.add_argument("--speaker-on", action="store_true", help="Turn speaker on at launch")
    args = parser.parse_args()

    # Single instance check
    mutex = _acquire_mutex()
    if mutex is None:
        logger.info("Another instance is already running. Signaling it to show.")
        _signal_existing_instance()
        sys.exit(0)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    cfg = load_config()
    controller = TapoController(
        ip=cfg["plug_ip"],
        email=cfg["tapo_email"],
        password=cfg["tapo_password"],
        timeout=cfg.get("timeout_sec", 5),
    )

    window = MainWindow(
        controller, cfg,
        start_hidden=args.startup,
        speaker_on_at_startup=args.speaker_on,
    )

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()