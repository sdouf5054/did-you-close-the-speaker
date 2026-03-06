"""
Did You Close the Speaker? — GUI Application
System tray icon with a compact control window.
"""

import sys
import os
import asyncio
import logging
import json
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
# Paths
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).parent
LOG_DIR = APP_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
SETTINGS_PATH = APP_DIR / "settings.json"

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
QFrame#separator {
    background-color: #2a2a4a;
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

        self._setup_ui()
        self._setup_tray()
        self._setup_idle_timer()

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
        self.setFixedSize(400, 460)
        self.setStyleSheet(STYLE_SHEET)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)

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

        # --- Settings ---
        settings_label = QLabel("Settings")
        settings_label.setObjectName("sectionLabel")
        settings_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(settings_label)

        self.chk_startup = QCheckBox("컴퓨터 시작 시 실행")
        self.chk_startup.setChecked(self.settings["start_with_windows"])
        self.chk_startup.stateChanged.connect(self._on_startup_changed)
        layout.addWidget(self.chk_startup)

        self.chk_startup_speaker = QCheckBox("컴퓨터 시작 시 실행 및 Speaker ON")
        self.chk_startup_speaker.setChecked(self.settings["speaker_on_at_startup"])
        self.chk_startup_speaker.setEnabled(self.settings["start_with_windows"])
        self.chk_startup_speaker.stateChanged.connect(self._on_startup_speaker_changed)
        layout.addWidget(self.chk_startup_speaker)

        self.chk_watchdog = QCheckBox("PC 종료/다시 시작 전에 스피커 먼저 끄기")
        self.chk_watchdog.setChecked(self.settings["watchdog_enabled"])
        self.chk_watchdog.stateChanged.connect(self._on_watchdog_changed)
        layout.addWidget(self.chk_watchdog)

        # Idle timer row
        idle_row = QHBoxLayout()
        idle_row.setSpacing(6)

        self.chk_idle = QCheckBox("입력 없으면 스피커 끄기:")
        self.chk_idle.setChecked(self.settings["idle_timer_enabled"])
        self.chk_idle.stateChanged.connect(self._on_idle_changed)
        idle_row.addWidget(self.chk_idle)

        IDLE_OPTIONS = ["5분", "10분", "15분", "20분", "25분", "30분", "45분", "60분"]
        IDLE_VALUES = [5, 10, 15, 20, 25, 30, 45, 60]

        self.combo_idle = QComboBox()
        self.combo_idle.addItems(IDLE_OPTIONS)
        current_val = self.settings["idle_timer_minutes"]
        if current_val in IDLE_VALUES:
            self.combo_idle.setCurrentIndex(IDLE_VALUES.index(current_val))
        self.combo_idle.setEnabled(self.settings["idle_timer_enabled"])
        self.combo_idle.setFixedWidth(70)
        self.combo_idle.setStyleSheet(
            "QComboBox { background-color: #2a2a4a; color: #cccccc; border: 1px solid #555577; "
            "border-radius: 3px; padding: 2px 6px; font-size: 11px; }"
            "QComboBox::drop-down { border: none; }"
            "QComboBox QAbstractItemView { background-color: #2a2a4a; color: #cccccc; "
            "selection-background-color: #2980b9; }"
        )
        self.combo_idle.currentIndexChanged.connect(self._on_idle_time_changed)
        idle_row.addWidget(self.combo_idle)
        idle_row.addStretch()
        layout.addLayout(idle_row)

        self.chk_idle_auto_on = QCheckBox("입력 감지 시 스피커 자동 켜기")
        self.chk_idle_auto_on.setChecked(self.settings["idle_auto_on"])
        self.chk_idle_auto_on.setEnabled(self.settings["idle_timer_enabled"])
        self.chk_idle_auto_on.stateChanged.connect(self._on_idle_auto_on_changed)
        layout.addWidget(self.chk_idle_auto_on)

        layout.addStretch()

    # ----- Settings handlers -----

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
        if enabled:
            self._idle_check_timer.start(30_000)
        else:
            self._idle_check_timer.stop()
        save_settings(self.settings)
        logger.info(f"Idle timer {'enabled' if enabled else 'disabled'}")

    def _on_idle_time_changed(self, index):
        values = [5, 10, 15, 20, 25, 30, 45, 60]
        if 0 <= index < len(values):
            self.settings["idle_timer_minutes"] = values[index]
            save_settings(self.settings)
            logger.info(f"Idle timer set to {values[index]} minutes")

    def _on_idle_auto_on_changed(self, state):
        enabled = state == Qt.Checked
        self.settings["idle_auto_on"] = enabled
        save_settings(self.settings)

    # ----- Idle timer -----

    def _setup_idle_timer(self):
        """
        Poll GetLastInputInfo every 30 seconds.
        Only runs when idle timer is enabled — zero overhead when off.
        """
        self._idle_check_timer = QTimer(self)
        self._idle_check_timer.timeout.connect(self._check_idle)
        if self.settings.get("idle_timer_enabled", False):
            self._idle_check_timer.start(30_000)

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
        """Called every 30s by QTimer. Checks idle state and acts."""
        if not self.settings.get("idle_timer_enabled", False):
            return

        idle_sec = self._get_idle_seconds()
        threshold_sec = self.settings.get("idle_timer_minutes", 15) * 60

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
                # Auto turn on when user returns
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
                # Reset flag so idle can trigger again next time
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

    def nativeEvent(self, event_type, message):
        if IS_WINDOWS and event_type == b"windows_generic_MSG":
            msg = ctypes.wintypes.MSG.from_address(int(message))

            # Wake from sleep/hibernate
            if msg.message == WM_POWERBROADCAST:
                if msg.wParam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
                    logger.info("System resumed from sleep/hibernate.")
                    self._set_busy(False)
                    self.plug_state = PlugState.UNKNOWN
                    self._idle_turned_off = False
                    QTimer.singleShot(3000, self.refresh_status)

            # Shutdown/restart watchdog
            elif msg.message == WM_QUERYENDSESSION:
                if self.settings.get("watchdog_enabled", False):
                    logger.info("WM_QUERYENDSESSION — watchdog turning off speaker.")
                    self._shutdown_intercepted = True
                    try:
                        asyncio.run(self._watchdog_turn_off())
                        logger.info("Watchdog: speaker OFF before shutdown.")
                    except Exception as e:
                        logger.error(f"Watchdog failed: {e}")

            elif msg.message == WM_ENDSESSION:
                if msg.wParam and self.settings.get("watchdog_enabled", False):
                    if not self._shutdown_intercepted:
                        logger.info("WM_ENDSESSION — last-chance speaker off.")
                        try:
                            asyncio.run(self._watchdog_turn_off())
                        except Exception as e:
                            logger.error(f"Watchdog (ENDSESSION) failed: {e}")

        return super().nativeEvent(event_type, message)

    async def _watchdog_turn_off(self):
        """Quick turn-off with aggressive timeout for shutdown scenarios."""
        quick = TapoController(
            ip=self.controller.ip,
            email=self.controller.email,
            password=self.controller.password,
            timeout=3,
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
# Entry point
# ---------------------------------------------------------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--startup", action="store_true", help="Launched from Windows startup")
    parser.add_argument("--speaker-on", action="store_true", help="Turn speaker on at launch")
    args = parser.parse_args()

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