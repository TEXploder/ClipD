import json
import os
import re
import shutil
import string
import sys
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import ctypes
import ctypes.wintypes as wintypes
import winreg

from PySide6 import QtCore, QtGui, QtWidgets
from cryptography.fernet import Fernet


APP_NAME = "ClipD"
APP_DISPLAY_NAME = "ClipD"
HISTORY_FILE_NAME = "history.bin"
KEY_FILE_NAME = "key.bin"
SETTINGS_FILE_NAME = "settings.json"
MAX_HISTORY_ITEMS = 50

BASE_DIR = Path(sys.argv[0]).resolve().parent
ICON_ICO_PATH = BASE_DIR / "logo.ico"
ICON_PNG_PATH = BASE_DIR / "logo.png"


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


class Theme:
    PRIMARY_BG = QtGui.QColor("#12131c")
    CARD_BG = QtGui.QColor("#1e2130")
    ACCENT = QtGui.QColor("#7f5af0")
    ACCENT_GRADIENT_START = QtGui.QColor("#7f5af0")
    ACCENT_GRADIENT_END = QtGui.QColor("#2cb67d")
    TEXT_PRIMARY = QtGui.QColor("#f5f7ff")
    TEXT_MUTED = QtGui.QColor("#9aa3c0")
    BORDER = QtGui.QColor("#2b2f44")

    @classmethod
    def apply_settings(cls, settings: "AppSettings") -> None:
        cls.ACCENT_GRADIENT_START = QtGui.QColor(settings.accent_start)
        if not cls.ACCENT_GRADIENT_START.isValid():
            cls.ACCENT_GRADIENT_START = QtGui.QColor("#7f5af0")
        cls.ACCENT_GRADIENT_END = QtGui.QColor(settings.accent_end)
        if not cls.ACCENT_GRADIENT_END.isValid():
            cls.ACCENT_GRADIENT_END = QtGui.QColor("#2cb67d")
        cls.ACCENT = cls.ACCENT_GRADIENT_START


def color_to_rgba(color: QtGui.QColor, alpha: float) -> str:
    alpha = clamp(alpha, 0.0, 1.0)
    return f"rgba({color.red()},{color.green()},{color.blue()},{int(alpha * 255)})"


def apply_app_theme(app: QtWidgets.QApplication, settings: "AppSettings") -> None:
    Theme.apply_settings(settings)

    app.setStyle("Fusion")
    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, Theme.PRIMARY_BG)
    palette.setColor(QtGui.QPalette.Base, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.AlternateBase, Theme.PRIMARY_BG.darker(115))
    palette.setColor(QtGui.QPalette.ToolTipBase, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.ToolTipText, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Text, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Button, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.ButtonText, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Highlight, Theme.ACCENT)
    palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.WindowText, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.PlaceholderText, Theme.TEXT_MUTED)
    app.setPalette(palette)

    font = QtGui.QFont("Segoe UI", 10)
    app.setFont(font)

    accent = Theme.ACCENT
    accent_hover = QtGui.QColor(accent)
    accent_hover = accent_hover.lighter(125)
    accent_pressed = QtGui.QColor(accent)
    accent_pressed = accent_pressed.darker(120)

    global_stylesheet = f"""
    QWidget {{
        background-color: #12131c;
        color: #f5f7ff;
    }}
    QLineEdit, QTextEdit {{
        background-color: #1e2130;
        border: 1px solid #2b2f44;
        border-radius: 10px;
        padding: 8px 12px;
        selection-background-color: {accent.name()};
    }}
    QLineEdit:focus {{
        border: 1px solid {accent.name()};
    }}
    QListWidget {{
        background: transparent;
        border: none;
    }}
    QPushButton {{
        background-color: {color_to_rgba(accent, 0.18)};
        border: 1px solid {color_to_rgba(accent, 0.35)};
        border-radius: 12px;
        padding: 8px 16px;
        font-weight: 600;
    }}
    QPushButton:hover {{
        background-color: {color_to_rgba(accent_hover, 0.24)};
        border: 1px solid {color_to_rgba(accent_hover, 0.55)};
    }}
    QPushButton:pressed {{
        background-color: {color_to_rgba(accent_pressed, 0.5)};
    }}
    QToolButton {{
        color: rgba(245, 247, 255, 210);
        border-radius: 10px;
        padding: 6px 12px;
        background-color: {color_to_rgba(accent, 0.12)};
        border: 1px solid {color_to_rgba(accent, 0.28)};
    }}
    QToolButton:hover {{
        background-color: {color_to_rgba(accent_hover, 0.18)};
        border: 1px solid {color_to_rgba(accent_hover, 0.45)};
    }}
    QFrame#titleBar {{
        background-color: rgba(24, 26, 40, 235);
        border-radius: 14px;
        border: 1px solid {color_to_rgba(accent, 0.35)};
    }}
    QLabel#titleBarLabel {{
        color: rgba(245, 247, 255, 230);
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }}
    QPushButton[accent=\"true\"] {{
        background-color: {color_to_rgba(accent, 0.8)};
        border: 1px solid {color_to_rgba(accent, 0.95)};
        color: #f5f7ff;
    }}
    QPushButton[accent=\"true\"]:hover {{
        background-color: {color_to_rgba(accent_hover, 0.9)};
    }}
    QPushButton[destructive=\"true\"] {{
        background-color: rgba(255, 99, 71, 0.18);
        border: 1px solid rgba(255, 99, 71, 0.5);
        color: rgba(255, 143, 119, 0.95);
    }}
    QPushButton[destructive=\"true\"]:hover {{
        background-color: rgba(255, 99, 71, 0.28);
        border: 1px solid rgba(255, 99, 71, 0.65);
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 12px;
        margin: 8px;
    }}
    QScrollBar::handle:vertical {{
        background: {color_to_rgba(accent, 0.5)};
        border-radius: 6px;
        min-height: 20px;
    }}
    """
    app.setStyleSheet(global_stylesheet)


def app_data_dir() -> Path:
    appdata = os.getenv("APPDATA")
    if not appdata:
        raise RuntimeError("APPDATA environment variable is not set")
    target = Path(appdata) / APP_NAME
    target.mkdir(parents=True, exist_ok=True)
    return target


def settings_path() -> Path:
    return app_data_dir() / SETTINGS_FILE_NAME


@dataclass
class AppSettings:
    toast_duration_ms: int = 2600
    toast_scale: float = 1.0
    accent_start: str = "#7f5af0"
    accent_end: str = "#2cb67d"
    hotkey_next: str = "Ctrl+Alt+Down"
    hotkey_prev: str = "Ctrl+Alt+Up"
    hotkey_show_history: str = "Ctrl+Alt+V"
    show_preview_overlay: bool = True
    capture_protection_enabled: bool = True
    first_run: bool = True

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def load(cls, path: Path) -> "AppSettings":
        if not path.exists():
            return cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return cls(**data).sanitized()
        except Exception:
            return cls()

    def save(self, path: Path) -> None:
        path.write_text(json.dumps(self.to_dict(), indent=2), encoding="utf-8")

    def sanitized(self) -> "AppSettings":
        preview_value = self.show_preview_overlay
        if isinstance(preview_value, str):
            preview_value = preview_value.strip().lower() not in ("0", "false", "no", "off")
        else:
            preview_value = bool(preview_value)
        first_run_value = self.first_run
        if isinstance(first_run_value, str):
            first_run_value = first_run_value.strip().lower() in ("1", "true", "yes", "on")
        else:
            first_run_value = bool(first_run_value)
        capture_value = self.capture_protection_enabled
        if isinstance(capture_value, str):
            capture_value = capture_value.strip().lower() not in ("0", "false", "no", "off")
        else:
            capture_value = bool(capture_value)
        sanitized = AppSettings(
            toast_duration_ms=int(clamp(self.toast_duration_ms, 600, 10000)),
            toast_scale=float(clamp(self.toast_scale, 0.6, 1.6)),
            accent_start=QtGui.QColor(self.accent_start).name()
            if QtGui.QColor(self.accent_start).isValid()
            else "#7f5af0",
            accent_end=QtGui.QColor(self.accent_end).name()
            if QtGui.QColor(self.accent_end).isValid()
            else "#2cb67d",
            hotkey_next=self.hotkey_next or "Ctrl+Alt+Down",
            hotkey_prev=self.hotkey_prev or "Ctrl+Alt+Up",
            hotkey_show_history=self.hotkey_show_history or "Ctrl+Alt+V",
            show_preview_overlay=preview_value,
            capture_protection_enabled=capture_value,
            first_run=first_run_value,
        )
        return sanitized

    def copy(self) -> "AppSettings":
        return AppSettings(**self.to_dict())


DEFAULT_SETTINGS = AppSettings()


def display_hotkey(sequence: str) -> str:
    result = sequence or ""
    replacements = {
        "Ctrl": "Strg",
        "Control": "Strg",
        "Meta": "Win",
        "Super": "Win",
        "Return": "Enter",
    }
    for src, dst in replacements.items():
        result = result.replace(src, dst)
    return result
class EncryptedStorage:
    def __init__(self, storage_path: Path) -> None:
        self.storage_path = storage_path
        self.key_path = storage_path.with_name(KEY_FILE_NAME)
        self._fernet = Fernet(self._load_or_create_key())

    def _load_or_create_key(self) -> bytes:
        if self.key_path.exists():
            return self.key_path.read_bytes()
        key = Fernet.generate_key()
        self.key_path.write_bytes(key)
        return key

    def load(self) -> List[dict]:
        if not self.storage_path.exists():
            return []
        try:
            encrypted = self.storage_path.read_bytes()
            data = self._fernet.decrypt(encrypted)
            return json.loads(data.decode("utf-8"))
        except Exception:
            # corrupted history -> start fresh
            return []

    def save(self, items: List[dict]) -> None:
        payload = json.dumps(items, ensure_ascii=False).encode("utf-8")
        token = self._fernet.encrypt(payload)
        self.storage_path.write_bytes(token)


@dataclass
class ClipboardItem:
    content: str
    timestamp: float

    @classmethod
    def from_dict(cls, data: dict) -> "ClipboardItem":
        return cls(content=data["content"], timestamp=data["timestamp"])


class ClipboardHistory(QtCore.QObject):
    historyUpdated = QtCore.Signal(list)
    selectionChanged = QtCore.Signal(object)

    def __init__(self, clipboard: QtGui.QClipboard, storage: EncryptedStorage) -> None:
        super().__init__()
        self._clipboard = clipboard
        self._storage = storage
        self._history: List[ClipboardItem] = [
            ClipboardItem.from_dict(entry) for entry in self._storage.load()
        ]
        self._current_index: Optional[int] = 0 if self._history else None
        self._suspend_capture = False
        self._clipboard.dataChanged.connect(self._on_clipboard_change)

    def _on_clipboard_change(self) -> None:
        if self._suspend_capture:
            return
        mime = self._clipboard.mimeData()
        if not mime or not mime.hasText():
            return
        text = mime.text()
        if not text.strip():
            return

        if self._history and self._history[0].content == text:
            return

        item = ClipboardItem(content=text, timestamp=time.time())
        self._history.insert(0, item)
        del self._history[MAX_HISTORY_ITEMS:]
        self._current_index = 0
        self._persist()
        self.historyUpdated.emit(self._history.copy())
        self.selectionChanged.emit(self.current_item())

    def _persist(self) -> None:
        serializable = [asdict(entry) for entry in self._history]
        self._storage.save(serializable)

    def current_item(self) -> Optional[ClipboardItem]:
        if self._current_index is None or not self._history:
            return None
        return self._history[self._current_index]

    def select_next(self) -> Optional[ClipboardItem]:
        if not self._history:
            return None
        if self._current_index is None:
            self._current_index = 0
        else:
            self._current_index = (self._current_index + 1) % len(self._history)
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def select_previous(self) -> Optional[ClipboardItem]:
        if not self._history:
            return None
        if self._current_index is None:
            self._current_index = 0
        else:
            self._current_index = (self._current_index - 1) % len(self._history)
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def select_index(self, index: int) -> Optional[ClipboardItem]:
        if not self._history:
            return None
        index = max(0, min(index, len(self._history) - 1))
        self._current_index = index
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def remove_entry(self, entry: ClipboardItem) -> None:
        for idx, existing in enumerate(self._history):
            if existing == entry:
                del self._history[idx]
                if not self._history:
                    self._current_index = None
                else:
                    if self._current_index is None:
                        self._current_index = 0
                    else:
                        self._current_index = min(self._current_index, len(self._history) - 1)
                self._persist()
                self.historyUpdated.emit(self._history.copy())
                self.selectionChanged.emit(self.current_item())
                return

    def clear(self) -> None:
        if not self._history:
            return
        self._history.clear()
        self._current_index = None
        self._persist()
        self.historyUpdated.emit([])
        self.selectionChanged.emit(None)

    def push_to_clipboard(self, item: ClipboardItem) -> None:
        self._suspend_capture = True
        self._clipboard.setText(item.content)
        QtCore.QTimer.singleShot(100, self._resume_capture)

    def _resume_capture(self) -> None:
        self._suspend_capture = False

    def all_items(self) -> List[ClipboardItem]:
        return self._history.copy()


class PreviewToast(QtWidgets.QWidget):
    def __init__(self, settings: AppSettings) -> None:
        super().__init__()
        self._settings = settings
        self.setWindowFlags(
            QtCore.Qt.ToolTip
            | QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, True)

        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._start_fade_out)

        self._opacity_effect = QtWidgets.QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self._opacity_effect)
        self._opacity_effect.setOpacity(0.0)

        self._fade = QtCore.QPropertyAnimation(self._opacity_effect, b"opacity", self)
        self._fade.setDuration(160)
        self._fade.setEasingCurve(QtCore.QEasingCurve.InOutQuad)
        self._fade.finished.connect(self._on_fade_finished)
        self._fade_target = 0.0

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(0)

        container = QtWidgets.QFrame()
        container.setStyleSheet("QFrame { background-color: transparent; }")
        container_layout = QtWidgets.QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        self._card = QtWidgets.QFrame()
        self._card.setObjectName("toastCard")
        card_layout = QtWidgets.QVBoxLayout(self._card)
        card_layout.setContentsMargins(18, 16, 18, 18)
        card_layout.setSpacing(8)

        title_row = QtWidgets.QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(8)

        self._accent_icon = QtWidgets.QLabel()
        self._accent_icon.setFixedSize(10, 10)
        title_row.addWidget(self._accent_icon)

        title = QtWidgets.QLabel("Clipboard Preview")
        title.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        title.setStyleSheet(
            """
            QLabel {
                font-size: 11px;
                letter-spacing: 1px;
                text-transform: uppercase;
                color: rgba(245, 247, 255, 180);
                background-color: transparent;
            }
            """
        )
        title_row.addWidget(title)
        title_row.addStretch(1)
        card_layout.addLayout(title_row)

        self._label = QtWidgets.QLabel("")
        self._label.setWordWrap(True)
        self._label.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self._label.setStyleSheet(
            """
            QLabel {
                color: #f5f7ff;
                font-size: 15px;
                font-weight: 500;
                background-color: transparent;
            }
            """
        )
        card_layout.addWidget(self._label)

        self._hint = QtWidgets.QLabel("Strg+V fuegt ein | Esc blendet aus")
        self._hint.setStyleSheet(
            """
            QLabel {
                color: rgba(154, 163, 192, 200);
                font-size: 11px;
                background-color: transparent;
            }
            """
        )
        card_layout.addWidget(self._hint)
        card_layout.setAlignment(self._hint, QtCore.Qt.AlignLeft | QtCore.Qt.AlignBottom)

        container_layout.addWidget(self._card)
        layout.addWidget(container)

        close_shortcut = QtGui.QShortcut(QtGui.QKeySequence("Esc"), self)
        close_shortcut.setContext(QtCore.Qt.ApplicationShortcut)
        close_shortcut.activated.connect(self._start_fade_out)

        self._apply_styles()

    def apply_settings(self, settings: AppSettings) -> None:
        self._settings = settings
        self._apply_styles()

    def _apply_styles(self) -> None:
        if not self._settings.show_preview_overlay:
            self._timer.stop()
            super().hide()
        start = QtGui.QColor(self._settings.accent_start)
        if not start.isValid():
            start = QtGui.QColor("#7f5af0")
        end = QtGui.QColor(self._settings.accent_end)
        if not end.isValid():
            end = QtGui.QColor("#2cb67d")
        self._card.setStyleSheet(
            f"""
            QFrame#toastCard {{
                background-color: rgba(24, 26, 40, 230);
                border-radius: 16px;
                border: 1px solid {color_to_rgba(start, 0.45)};
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(24, 26, 40, 240),
                    stop:1 rgba(30, 33, 48, 240)
                );
            }}
            """
        )
        self._accent_icon.setStyleSheet(
            f"""
            QLabel {{
                border-radius: 5px;
                background-color: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 {start.name()}, stop:1 {end.name()}
                );
            }}
            """
        )

    def _start_fade_out(self) -> None:
        self._timer.stop()
        self._animate_opacity(0.0, hide_after=True)

    def _on_fade_finished(self) -> None:
        if self._fade_target == 0.0:
            super().hide()

    def show_preview(self, item: ClipboardItem) -> None:
        display_text = item.content.strip()
        if len(display_text) > 220:
            display_text = display_text[:220] + "..."
        self._label.setText(display_text if display_text else "<leer>")
        if not self._settings.show_preview_overlay:
            super().hide()
            return
        self.ensurePolished()
        size = self.sizeHint()
        scale = float(clamp(self._settings.toast_scale, 0.6, 1.6))
        base_width = max(size.width(), 260) if size.isValid() else 260
        base_height = max(size.height(), 140) if size.isValid() else 140
        width = int(clamp(base_width * scale, 160, 520))
        height = int(clamp(base_height * scale, 120, 320))
        self.resize(width, height)

        cursor_pos = QtGui.QCursor.pos()
        screen = QtGui.QGuiApplication.screenAt(cursor_pos)
        if screen:
            available = screen.availableGeometry()
        else:
            available = QtGui.QGuiApplication.primaryScreen().availableGeometry()
        target_x = cursor_pos.x() + 24
        target_y = cursor_pos.y() + 24
        target_x = max(available.left(), min(target_x, available.right() - self.width()))
        target_y = max(available.top(), min(target_y, available.bottom() - self.height()))
        self.move(target_x, target_y)

        if not self.isVisible():
            self._opacity_effect.setOpacity(0.0)
            super().show()
        else:
            super().show()
        self.raise_()
        self._animate_opacity(1.0, hide_after=False)
        self._timer.start(int(clamp(self._settings.toast_duration_ms, 600, 10000)))

    def hide(self) -> None:
        self._start_fade_out()

    def _animate_opacity(self, value: float, hide_after: bool) -> None:
        self._fade.stop()
        self._fade_target = value
        self._fade.setStartValue(self._opacity_effect.opacity())
        self._fade.setEndValue(value)
        self._fade.setDuration(200 if hide_after else 180)
        self._fade.start()


class HistoryDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, settings: AppSettings, parent: Optional[QtCore.QObject] = None) -> None:
        super().__init__(parent)
        self._settings = settings

    def update_settings(self, settings: AppSettings) -> None:
        self._settings = settings

    def paint(
        self,
        painter: QtGui.QPainter,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> None:
        entry = index.data(QtCore.Qt.UserRole)
        if not isinstance(entry, ClipboardItem):
            super().paint(painter, option, index)
            return

        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

        rect = option.rect.adjusted(8, 3, -8, -5)
        is_selected = option.state & QtWidgets.QStyle.State_Selected

        background_color = QtGui.QColor(30, 33, 48)
        start_color = QtGui.QColor(self._settings.accent_start)
        if not start_color.isValid():
            start_color = QtGui.QColor("#7f5af0")
        end_color = QtGui.QColor(self._settings.accent_end)
        if not end_color.isValid():
            end_color = QtGui.QColor("#2cb67d")
        if is_selected:
            gradient = QtGui.QLinearGradient(rect.topLeft(), rect.bottomRight())
            gradient.setColorAt(0.0, start_color.lighter(120))
            gradient.setColorAt(1.0, end_color)
            painter.setBrush(gradient)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 35), 1.0))
        else:
            painter.setBrush(background_color)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 20), 1.0))

        painter.drawRoundedRect(rect, 14, 14)

        accent_rect = QtCore.QRect(rect.left() + 14, rect.top() + 14, 6, max(36, rect.height() - 28))
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(start_color if not is_selected else QtCore.Qt.white)
        painter.drawRoundedRect(accent_rect, 3, 3)

        timestamp = time.strftime("%d.%m.%Y %H:%M:%S", time.localtime(entry.timestamp))
        text_rect = rect.adjusted(34, 14, -18, -18)

        timestamp_font = QtGui.QFont(option.font)
        timestamp_font.setPointSize(option.font.pointSize() - 1)
        painter.setFont(timestamp_font)
        painter.setPen(QtGui.QColor(154, 163, 192))
        painter.drawText(text_rect, QtCore.Qt.TextSingleLine, timestamp)

        content_rect = text_rect.adjusted(0, 20, 0, 0)
        content_font = QtGui.QFont(option.font)
        content_font.setPointSize(option.font.pointSize() + 1)
        content_font.setBold(True)
        painter.setFont(content_font)
        painter.setPen(QtGui.QColor(245, 247, 255))
        snippet = entry.content.replace("\r", " ").replace("\n", " ")
        snippet = re.sub(r"(\S{40})", r"\1" + "\u200b", snippet)
        if len(snippet) > 320:
            snippet = snippet[:320] + "..."
        text_option = QtGui.QTextOption()
        text_option.setWrapMode(QtGui.QTextOption.WrapAtWordBoundaryOrAnywhere)
        painter.drawText(QtCore.QRectF(content_rect), snippet or "<leer>", text_option)

        painter.restore()

    def sizeHint(
        self,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> QtCore.QSize:
        return QtCore.QSize(0, 88)


class HistoryWindow(QtWidgets.QMainWindow):
    itemActivated = QtCore.Signal(ClipboardItem)

    def __init__(
        self,
        history: ClipboardHistory,
        settings: AppSettings,
        settings_callback: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_DISPLAY_NAME} - Verlauf")
        self._history = history
        self._settings = settings
        self._settings_callback = settings_callback
        self.setWindowIcon(window_icon(self._settings))
        self.setMinimumSize(520, 600)
        self.setWindowFlags(
            QtCore.Qt.Window
            | QtCore.Qt.WindowCloseButtonHint
            | QtCore.Qt.WindowMinimizeButtonHint
        )
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)

        layout = QtWidgets.QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(18)

        self._header = QtWidgets.QFrame()
        self._header.setObjectName("historyHeader")
        self._header.setStyleSheet(
            """
            QFrame#historyHeader {
                border-radius: 20px;
                padding: 1px;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(127, 90, 240, 180),
                    stop:1 rgba(44, 182, 125, 160)
                );
            }
            """
        )
        self._header_inner = QtWidgets.QFrame()
        self._header_inner.setObjectName("historyHeaderInner")
        self._header_inner.setStyleSheet(
            """
            QFrame#historyHeaderInner {
                background-color: rgba(18, 19, 28, 220);
                border-radius: 18px;
            }
            """
        )

        header_layout_outer = QtWidgets.QVBoxLayout(self._header)
        header_layout_outer.setContentsMargins(0, 0, 0, 0)
        header_layout_outer.addWidget(self._header_inner)

        header_layout = QtWidgets.QVBoxLayout(self._header_inner)
        header_layout.setContentsMargins(24, 20, 24, 20)
        header_layout.setSpacing(4)

        self._title_label = QtWidgets.QLabel(f"{APP_DISPLAY_NAME} Verlauf")
        self._title_label.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self._title_label.setStyleSheet(
            """
            QLabel {
                font-size: 22px;
                font-weight: 700;
                letter-spacing: 0.5px;
                background-color: transparent;
            }
            """
        )
        header_layout.addWidget(self._title_label)

        self._stats_label = QtWidgets.QLabel("")
        self._stats_label.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self._stats_label.setStyleSheet(
            """
            QLabel {
                font-size: 12px;
                color: rgba(245, 247, 255, 160);
                background-color: transparent;
            }
            """
        )
        header_layout.addWidget(self._stats_label)
        layout.addWidget(self._header)

        controls = QtWidgets.QHBoxLayout()
        controls.setSpacing(12)

        self._search_box = QtWidgets.QLineEdit()
        self._search_box.setPlaceholderText("Verlauf durchsuchen...")
        self._search_box.setClearButtonEnabled(True)
        controls.addWidget(self._search_box, 1)

        self._clear_button = QtWidgets.QToolButton()
        self._clear_button.setText("Verlauf leeren")
        #self._clear_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TrashIcon))
        self._clear_button.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self._clear_button.setCheckable(False)
        self._clear_button.clicked.connect(self._confirm_clear)
        controls.addWidget(self._clear_button)

        layout.addLayout(controls)

        self._list_container = QtWidgets.QFrame()
        self._list_container.setObjectName("historyListContainer")
        self._list_container.setStyleSheet(
            """
            QFrame#historyListContainer {
                background-color: rgba(18, 19, 28, 210);
                border-radius: 18px;
                border: 1px solid rgba(255, 255, 255, 25);
            }
            """
        )
        list_layout = QtWidgets.QVBoxLayout(self._list_container)
        list_layout.setContentsMargins(4, 12, 4, 12)
        list_layout.setSpacing(0)

        self._list = QtWidgets.QListWidget()
        self._list.setUniformItemSizes(False)
        self._list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._list.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self._list.setSpacing(8)
        self._list.setWordWrap(True)
        self._list.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self._delegate = HistoryDelegate(self._settings, self._list)
        self._list.setItemDelegate(self._delegate)
        list_layout.addWidget(self._list)

        self._empty_state = QtWidgets.QLabel(
            "Noch keine Eintraege. Kopiere Text, um hier etwas zu sehen."
        )
        self._empty_state.setAlignment(QtCore.Qt.AlignCenter)
        self._empty_state.setStyleSheet(
            """
            QLabel {
                color: rgba(154, 163, 192, 180);
                font-size: 13px;
                padding: 40px 0;
            }
            """
        )
        list_layout.addWidget(self._empty_state)
        layout.addWidget(self._list_container, 1)

        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(10)

        self._copy_button = QtWidgets.QPushButton("In Zwischenablage")
        self._copy_button.setProperty("accent", True)
        self._copy_button.style().unpolish(self._copy_button)
        self._copy_button.style().polish(self._copy_button)
        actions.addWidget(self._copy_button)

        actions.addStretch(1)

        self._settings_button = QtWidgets.QPushButton("Einstellungen")
        self._settings_button.setProperty("accent", True)
        self._settings_button.style().unpolish(self._settings_button)
        self._settings_button.style().polish(self._settings_button)
        actions.addWidget(self._settings_button)

        self._close_button = QtWidgets.QPushButton("Schliessen")
        actions.addWidget(self._close_button)

        layout.addLayout(actions)

        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 18)
        shadow.setColor(QtGui.QColor(0, 0, 0, 160))
        self._list_container.setGraphicsEffect(shadow)

        self._copy_button.clicked.connect(self._activate_selected)
        self._close_button.clicked.connect(self.close)
        self._settings_button.clicked.connect(self._open_settings)
        self._search_box.textChanged.connect(self._on_search_changed)
        self._list.itemDoubleClicked.connect(lambda _: self._activate_selected())
        self._list.installEventFilter(self)

        history.historyUpdated.connect(self._refresh)
        self._items_cache: List[ClipboardItem] = history.all_items()
        self._refresh(self._items_cache)
        self._apply_styles()

    def _apply_styles(self) -> None:
        start = QtGui.QColor(self._settings.accent_start)
        if not start.isValid():
            start = QtGui.QColor("#7f5af0")
        end = QtGui.QColor(self._settings.accent_end)
        if not end.isValid():
            end = QtGui.QColor("#2cb67d")
        self._header.setStyleSheet(
            f"""
            QFrame#historyHeader {{
                border-radius: 20px;
                padding: 1px;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 {color_to_rgba(start, 0.75)},
                    stop:1 {color_to_rgba(end, 0.65)}
                );
            }}
            """
        )
        self._header_inner.setStyleSheet(
            """
            QFrame#historyHeaderInner {
                background-color: rgba(18, 19, 28, 220);
                border-radius: 18px;
            }
            """
        )
        self._list_container.setStyleSheet(
            f"""
            QFrame#historyListContainer {{
                background-color: rgba(18, 19, 28, 210);
                border-radius: 18px;
                border: 1px solid {color_to_rgba(start, 0.25)};
            }}
            """
        )
        for button in (self._copy_button, self._settings_button):
            button.style().unpolish(button)
            button.style().polish(button)
        self._delegate.update_settings(self._settings)
        self._list.viewport().update()

    def apply_settings(self, settings: AppSettings) -> None:
        self._settings = settings
        self._apply_styles()
        self._apply_current_filter()

    def _open_settings(self) -> None:
        if self._settings_callback:
            self._settings_callback()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if obj is self._list and event.type() == QtCore.QEvent.KeyPress:
            if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                self._activate_selected()
                return True
            if event.key() == QtCore.Qt.Key_Delete:
                self._delete_selected()
                return True
        return super().eventFilter(obj, event)

    def _refresh(self, items: List[ClipboardItem]) -> None:
        self._items_cache = items
        self._apply_current_filter()

    def _on_search_changed(self, _: str) -> None:
        self._apply_current_filter()

    def _apply_current_filter(self) -> None:
        query = self._search_box.text().strip().lower()
        if query:
            visible = [item for item in self._items_cache if query in item.content.lower()]
        else:
            visible = list(self._items_cache)
        self._populate_list(visible)
        self._update_stats(len(visible))

    def _populate_list(self, items: List[ClipboardItem]) -> None:
        self._list.clear()
        for entry in items:
            item = QtWidgets.QListWidgetItem()
            item.setData(QtCore.Qt.UserRole, entry)
            self._list.addItem(item)
        has_items = bool(items)
        self._empty_state.setVisible(not has_items)
        self._list.setVisible(has_items)
        if has_items:
            self._list.setCurrentRow(0)

    def _update_stats(self, visible_count: int) -> None:
        total = len(self._items_cache)
        query = self._search_box.text().strip()
        if query:
            self._stats_label.setText(
                f"{visible_count} Treffer · {total} Eintraege gesamt"
            )
        else:
            self._stats_label.setText(
                f"{total} Eintraege · {display_hotkey(self._settings.hotkey_prev)} / {display_hotkey(self._settings.hotkey_next)}"
            )

    def _activate_selected(self) -> None:
        selected = self._list.currentItem()
        if not selected:
            return
        entry = selected.data(QtCore.Qt.UserRole)
        if entry:
            self.itemActivated.emit(entry)
            self.close()

    def _delete_selected(self) -> None:
        selected = self._list.currentItem()
        if not selected:
            return
        entry = selected.data(QtCore.Qt.UserRole)
        if entry:
            self._history.remove_entry(entry)

    def _confirm_clear(self) -> None:
        if not self._items_cache:
            return
        reply = QtWidgets.QMessageBox.question(
            self,
            "Verlauf leeren",
            "Soll der gesamte Verlauf geloescht werden?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if reply == QtWidgets.QMessageBox.Yes:
            self._history.clear()


class HotkeyEditor(QtWidgets.QLineEdit):
    sequenceChanged = QtCore.Signal(str)

    def __init__(self, initial: str, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self._sequence = initial
        self._update_text()
        self.setPlaceholderText("Shortcut aufnehmen")

    def sequence(self) -> str:
        return self._sequence

    def setSequence(self, sequence: str) -> None:
        self._sequence = sequence
        self._update_text()

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        key = event.key()
        if key in (
            QtCore.Qt.Key_Control,
            QtCore.Qt.Key_Shift,
            QtCore.Qt.Key_Alt,
            QtCore.Qt.Key_Meta,
        ):
            return
        if key in (QtCore.Qt.Key_Backspace, QtCore.Qt.Key_Delete):
            self._sequence = ""
            self._update_text()
            self.sequenceChanged.emit(self._sequence)
            event.accept()
            return
        if key == QtCore.Qt.Key_Escape:
            self.clearFocus()
            event.accept()
            return

        sequence = QtGui.QKeySequence(event.modifiers() | key)
        text = sequence.toString(QtGui.QKeySequence.PortableText)
        if text:
            self._sequence = text
            self._update_text()
            self.sequenceChanged.emit(self._sequence)
            event.accept()
            return
        super().keyPressEvent(event)

    def _update_text(self) -> None:
        if self._sequence:
            self.setText(display_hotkey(self._sequence))
        else:
            self.setText("Nicht gesetzt")


class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget, settings: AppSettings) -> None:
        super().__init__(parent)
        self.setWindowTitle("Einstellungen")
        self.setModal(True)
        self._settings = settings.copy()
        self._accent_start_color = self._settings.accent_start
        self._accent_end_color = self._settings.accent_end

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        overlay_group = QtWidgets.QGroupBox("Overlay")
        overlay_layout = QtWidgets.QFormLayout(overlay_group)
        overlay_layout.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        overlay_layout.setFormAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)

        scale_row = QtWidgets.QHBoxLayout()
        self._scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self._scale_slider.setRange(60, 160)
        self._scale_slider.setTickInterval(5)
        self._scale_slider.setValue(int(self._settings.toast_scale * 100))
        self._scale_slider.setTickPosition(QtWidgets.QSlider.TicksBelow)
        self._scale_value_label = QtWidgets.QLabel(f"{self._scale_slider.value()} %")
        self._scale_value_label.setMinimumWidth(48)
        scale_row.addWidget(self._scale_slider, 1)
        scale_row.addWidget(self._scale_value_label, 0)
        overlay_layout.addRow("Größe:", scale_row)
        self._scale_slider.valueChanged.connect(
            lambda v: self._scale_value_label.setText(f"{v} %")
        )

        self._duration_spin = QtWidgets.QSpinBox()
        self._duration_spin.setRange(600, 10000)
        self._duration_spin.setSingleStep(100)
        self._duration_spin.setValue(self._settings.toast_duration_ms)
        overlay_layout.addRow("Anzeigezeit (ms):", self._duration_spin)

        color_row = QtWidgets.QHBoxLayout()
        self._accent_start_button = self._create_color_button(self._accent_start_color)
        self._accent_end_button = self._create_color_button(self._accent_end_color)
        start_label = QtWidgets.QLabel("Start")
        end_label = QtWidgets.QLabel("Ende")
        start_label.setStyleSheet("color: rgba(245,247,255,180); background: transparent;")
        end_label.setStyleSheet("color: rgba(245,247,255,180); background: transparent;")
        color_row.addWidget(start_label)
        color_row.addWidget(self._accent_start_button)
        color_row.addSpacing(16)
        color_row.addWidget(end_label)
        color_row.addWidget(self._accent_end_button)
        overlay_layout.addRow("Farbverlauf:", color_row)

        self._preview_checkbox = QtWidgets.QCheckBox("Overlay beim Durchscrollen anzeigen")
        self._preview_checkbox.setChecked(self._settings.show_preview_overlay)
        overlay_layout.addRow("Vorschau:", self._preview_checkbox)

        self._capture_checkbox = QtWidgets.QCheckBox("Fenster vor Bildschirmaufnahmen schuetzen")
        self._capture_checkbox.setChecked(self._settings.capture_protection_enabled)
        overlay_layout.addRow("Aufnahmeschutz:", self._capture_checkbox)

        layout.addWidget(overlay_group)

        hotkey_group = QtWidgets.QGroupBox("Hotkeys")
        hotkey_layout = QtWidgets.QFormLayout(hotkey_group)
        hotkey_layout.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        self._prev_editor = HotkeyEditor(self._settings.hotkey_prev)
        self._next_editor = HotkeyEditor(self._settings.hotkey_next)
        self._show_editor = HotkeyEditor(self._settings.hotkey_show_history)

        hotkey_layout.addRow("Vorheriger Eintrag:", self._prev_editor)
        hotkey_layout.addRow("Naechster Eintrag:", self._next_editor)
        hotkey_layout.addRow("Verlauf öffnen:", self._show_editor)

        layout.addWidget(hotkey_group)

        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel
        )
        self._reset_button = QtWidgets.QPushButton("Standardwerte")
        button_box.addButton(self._reset_button, QtWidgets.QDialogButtonBox.ActionRole)
        layout.addWidget(button_box)

        button_box.accepted.connect(self._accept)
        button_box.rejected.connect(self.reject)
        self._reset_button.clicked.connect(self._reset_defaults)

        self._accent_start_button.clicked.connect(
            lambda: self._pick_color("Startfarbe", True)
        )
        self._accent_end_button.clicked.connect(
            lambda: self._pick_color("Endfarbe", False)
        )

        self._result_settings: Optional[AppSettings] = None

    def result_settings(self) -> Optional[AppSettings]:
        return self._result_settings

    def _create_color_button(self, color: str) -> QtWidgets.QPushButton:
        btn = QtWidgets.QPushButton()
        btn.setFixedSize(52, 24)
        btn.setStyleSheet(self._color_stylesheet(color))
        return btn

    def _color_stylesheet(self, color: str) -> str:
        return (
            "QPushButton {"
            f"background-color: {color};"
            "border-radius: 12px;"
            "border: 1px solid rgba(255, 255, 255, 60);"
            "}"
            "QPushButton:hover {"
            "border: 1px solid rgba(255, 255, 255, 120);"
            "}"
        )

    def _pick_color(self, title: str, is_start: bool) -> None:
        current = QtGui.QColor(self._accent_start_color if is_start else self._accent_end_color)
        color = QtWidgets.QColorDialog.getColor(
            current, self, title, QtWidgets.QColorDialog.DontUseNativeDialog
        )
        if color.isValid():
            if is_start:
                self._accent_start_color = color.name()
                self._accent_start_button.setStyleSheet(
                    self._color_stylesheet(self._accent_start_color)
                )
            else:
                self._accent_end_color = color.name()
                self._accent_end_button.setStyleSheet(
                    self._color_stylesheet(self._accent_end_color)
                )

    def _reset_defaults(self) -> None:
        defaults = DEFAULT_SETTINGS.copy()
        self._scale_slider.setValue(int(defaults.toast_scale * 100))
        self._duration_spin.setValue(defaults.toast_duration_ms)
        self._accent_start_color = defaults.accent_start
        self._accent_end_color = defaults.accent_end
        self._accent_start_button.setStyleSheet(
            self._color_stylesheet(self._accent_start_color)
        )
        self._accent_end_button.setStyleSheet(
            self._color_stylesheet(self._accent_end_color)
        )
        self._prev_editor.setSequence(defaults.hotkey_prev)
        self._next_editor.setSequence(defaults.hotkey_next)
        self._show_editor.setSequence(defaults.hotkey_show_history)
        self._preview_checkbox.setChecked(defaults.show_preview_overlay)
        self._capture_checkbox.setChecked(defaults.capture_protection_enabled)

    def _accept(self) -> None:
        try:
            new_settings = AppSettings(
                toast_duration_ms=self._duration_spin.value(),
                toast_scale=self._scale_slider.value() / 100.0,
                accent_start=self._accent_start_color,
                accent_end=self._accent_end_color,
                hotkey_prev=self._prev_editor.sequence(),
                hotkey_next=self._next_editor.sequence(),
                hotkey_show_history=self._show_editor.sequence(),
                show_preview_overlay=self._preview_checkbox.isChecked(),
                capture_protection_enabled=self._capture_checkbox.isChecked(),
                first_run=self._settings.first_run,
            ).sanitized()
            for seq in (
                new_settings.hotkey_prev,
                new_settings.hotkey_next,
                new_settings.hotkey_show_history,
            ):
                if not seq:
                    raise ValueError("Hotkeys duerfen nicht leer sein.")
                parse_hotkey(seq)
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Ungueltige Eingabe", str(exc))
            return

        self._result_settings = new_settings
        self.accept()


class HotkeyManager(QtCore.QObject):
    hotkeyTriggered = QtCore.Signal(int)

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_WIN = 0x0008
    WM_HOTKEY = 0x0312

    def __init__(self, parent: Optional[QtCore.QObject] = None) -> None:
        super().__init__(parent)
        self._ids: List[int] = []
        self._event_filter = _NativeHotkeyEventFilter(self.hotkeyTriggered)
        QtWidgets.QApplication.instance().installNativeEventFilter(self._event_filter)

    def register_hotkey(self, hotkey_id: int, modifiers: int, key: int) -> None:
        if not ctypes.windll.user32.RegisterHotKey(None, hotkey_id, modifiers, key):
            raise RuntimeError(f"Hotkey {hotkey_id} konnte nicht registriert werden")
        self._ids.append(hotkey_id)

    def unregister_all(self) -> None:
        for hotkey_id in self._ids:
            ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
        self._ids.clear()

    def __del__(self) -> None:
        self.unregister_all()


class _NativeHotkeyEventFilter(QtCore.QAbstractNativeEventFilter):
    def __init__(self, signal) -> None:
        super().__init__()
        self._signal = signal

    def nativeEventFilter(self, eventType, message) -> tuple[bool, int]:
        if eventType not in (b"windows_generic_MSG", "windows_generic_MSG"):
            return False, 0
        try:
            pointer = int(message)
        except (TypeError, ValueError):
            return False, 0
        msg = ctypes.cast(pointer, ctypes.POINTER(wintypes.MSG)).contents
        if msg.message == HotkeyManager.WM_HOTKEY:
            self._signal.emit(int(msg.wParam))
            return True, 0
        return False, 0


VK_CODE_MAP = {
    "UP": 0x26,
    "DOWN": 0x28,
    "LEFT": 0x25,
    "RIGHT": 0x27,
    "PAGEUP": 0x21,
    "PAGEDOWN": 0x22,
    "HOME": 0x24,
    "END": 0x23,
    "INSERT": 0x2D,
    "DELETE": 0x2E,
    "DEL": 0x2E,
    "SPACE": 0x20,
    "TAB": 0x09,
    "ESC": 0x1B,
    "ESCAPE": 0x1B,
    "ENTER": 0x0D,
    "RETURN": 0x0D,
    "BACKSPACE": 0x08,
    "PLUS": 0xBB,
    "MINUS": 0xBD,
    "COMMA": 0xBC,
    "PERIOD": 0xBE,
    "POINT": 0xBE,
}


def parse_hotkey(sequence: str) -> Tuple[int, int]:
    if not sequence:
        raise ValueError("Hotkey darf nicht leer sein.")
    parts = [part.strip() for part in sequence.replace("-", "+").split("+") if part.strip()]
    modifiers = 0
    key_code: Optional[int] = None

    for part in parts:
        lower = part.lower()
        if lower in ("ctrl", "control", "strg"):
            modifiers |= HotkeyManager.MOD_CONTROL
            continue
        if lower in ("alt", "option"):
            modifiers |= HotkeyManager.MOD_ALT
            continue
        if lower in ("shift",):
            modifiers |= HotkeyManager.MOD_SHIFT
            continue
        if lower in ("win", "meta", "super"):
            modifiers |= HotkeyManager.MOD_WIN
            continue
        if key_code is not None:
            raise ValueError(f"Mehrere Tasten in Hotkey '{sequence}' erkannt.")
        key_code = _resolve_key_code(part)

    if key_code is None:
        raise ValueError(f"Es wurde keine Taste in '{sequence}' erkannt.")
    return modifiers, key_code


def _resolve_key_code(part: str) -> int:
    upper = part.upper()
    if len(upper) == 1 and upper in string.ascii_uppercase:
        return ord(upper)
    if len(upper) == 1 and upper in string.digits:
        return ord(upper)
    if upper in VK_CODE_MAP:
        return VK_CODE_MAP[upper]
    if upper.startswith("F") and upper[1:].isdigit():
        index = int(upper[1:])
        if 1 <= index <= 24:
            return 0x70 + (index - 1)
    raise ValueError(f"Unbekannte Taste '{part}' in Hotkey.")


WDA_NONE = 0x00000000
WDA_MONITOR = 0x00000001
WDA_EXCLUDEFROMCAPTURE = 0x00000011


def set_window_capture_protection(widget: QtWidgets.QWidget, enabled: bool) -> None:
    if widget is None or not sys.platform.startswith("win"):
        return
    try:
        hwnd = int(widget.winId())
    except Exception:
        return
    if not hwnd:
        return
    try:
        user32 = ctypes.windll.user32
    except AttributeError:
        return
    affinity = WDA_EXCLUDEFROMCAPTURE if enabled else WDA_NONE
    if not user32.SetWindowDisplayAffinity(hwnd, affinity):
        if enabled:
            user32.SetWindowDisplayAffinity(hwnd, WDA_MONITOR)


def resource_path(name: str) -> Path:
    return BASE_DIR / name


def perform_initial_install(settings: AppSettings) -> Path:
    current_path = Path(sys.argv[0]).resolve()
    if not settings.first_run or not sys.platform.startswith("win"):
        return current_path
    if current_path.suffix.lower() != ".exe":
        return current_path
    program_files = os.getenv("ProgramFiles")
    if not program_files:
        return current_path
    target_path = current_path
    try:
        target_dir = Path(program_files) / "TEX-Programme" / APP_NAME
        target_dir.mkdir(parents=True, exist_ok=True)
        candidate = target_dir / current_path.name
        if current_path.resolve() != candidate.resolve():
            shutil.copy2(current_path, candidate)
        if candidate.exists():
            target_path = candidate
        for asset in ("logo.png", "logo.ico"):
            asset_source = resource_path(asset)
            if asset_source.exists():
                try:
                    shutil.copy2(asset_source, target_dir / asset)
                except Exception:
                    pass
        start_menu_root = Path(os.getenv("APPDATA", "")) / "Microsoft/Windows/Start Menu/Programs" / APP_NAME
        try:
            start_menu_root.mkdir(parents=True, exist_ok=True)
            shutil.copy2(target_path, start_menu_root / target_path.name)
        except Exception:
            pass
    except Exception:
        return current_path
    return target_path


def ensure_autostart(script_path: Path) -> None:
    executable = Path(sys.executable)
    if executable.name.lower().startswith("python"):
        command = f'"{executable}" "{script_path}"'
    else:
        command = f'"{script_path}"'
    try:
        key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run"
        )
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        winreg.CloseKey(key)
    except OSError:
        pass


def load_app_icon() -> QtGui.QIcon:
    for path in (ICON_ICO_PATH, ICON_PNG_PATH):
        if path.exists():
            icon = QtGui.QIcon(str(path))
            if not icon.isNull():
                return icon
    return QtGui.QIcon()


def build_tray_icon(settings: AppSettings) -> QtGui.QIcon:
    icon = load_app_icon()
    if not icon.isNull():
        return icon
    pixmap = QtGui.QPixmap(64, 64)
    pixmap.fill(QtCore.Qt.transparent)

    painter = QtGui.QPainter(pixmap)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)
    gradient = QtGui.QLinearGradient(0, 0, 64, 64)
    start = QtGui.QColor(settings.accent_start)
    if not start.isValid():
        start = QtGui.QColor("#4cf0c7")
    end = QtGui.QColor(settings.accent_end)
    if not end.isValid():
        end = QtGui.QColor("#ef38ef")
    gradient.setColorAt(0, start)
    gradient.setColorAt(1, end)
    painter.setBrush(QtGui.QBrush(gradient))
    painter.setPen(QtGui.QPen(QtGui.QColor("#ffffff"), 2))
    painter.drawRoundedRect(4, 4, 56, 56, 16, 16)
    painter.setPen(QtGui.QPen(QtGui.QColor("#ffffff"), 4))
    painter.drawLine(22, 20, 42, 20)
    painter.drawLine(22, 32, 42, 32)
    painter.drawLine(22, 44, 38, 44)
    painter.end()

    return QtGui.QIcon(pixmap)


def window_icon(settings: AppSettings) -> QtGui.QIcon:
    icon = load_app_icon()
    if icon.isNull():
        icon = build_tray_icon(settings)
    return icon


class MainController(QtCore.QObject):
    HOTKEY_NEXT = 1001
    HOTKEY_PREV = 1002
    HOTKEY_SHOW = 1003

    def __init__(self, app: QtWidgets.QApplication, settings: AppSettings) -> None:
        super().__init__()
        self._app = app
        self._settings = settings.sanitized()
        self._main_window: Optional["MainWindow"] = None
        QtWidgets.QApplication.setQuitOnLastWindowClosed(False)

        self._storage = EncryptedStorage(app_data_dir() / HISTORY_FILE_NAME)
        self._clipboard_history = ClipboardHistory(
            clipboard=app.clipboard(), storage=self._storage
        )
        self._history_window: Optional[HistoryWindow] = None
        self._shortcut_host = QtWidgets.QWidget()
        self._shortcut_host.setAttribute(QtCore.Qt.WA_DontShowOnScreen, True)
        self._shortcut_host.hide()
        self._show_history_shortcut = QtGui.QShortcut(
            QtGui.QKeySequence(self._settings.hotkey_show_history), self._shortcut_host
        )
        self._show_history_shortcut.setContext(QtCore.Qt.ApplicationShortcut)
        self._show_history_shortcut.activated.connect(self._show_history)

        self._toast = PreviewToast(self._settings)

        self._tray = QtWidgets.QSystemTrayIcon(build_tray_icon(self._settings))
        self._tray.setToolTip(APP_DISPLAY_NAME)
        tray_menu = QtWidgets.QMenu()
        open_action = tray_menu.addAction("Verlauf anzeigen")
        open_action.triggered.connect(self._show_history)
        tray_menu.addSeparator()
        quit_action = tray_menu.addAction("Beenden")
        quit_action.triggered.connect(self._quit)
        self._tray.setContextMenu(tray_menu)
        self._tray.activated.connect(self._on_tray_activated)
        self._tray.show()

        self._hotkeys = HotkeyManager()
        self._hotkeys.hotkeyTriggered.connect(self._process_hotkey)
        self._register_hotkeys()

        self._clipboard_history.selectionChanged.connect(self._on_selection_change)
        current = self._clipboard_history.current_item()
        if current and self._settings.show_preview_overlay:
            self._toast.show_preview(current)

    def _register_hotkeys(self) -> None:
        self._hotkeys.unregister_all()
        mappings = [
            (self.HOTKEY_PREV, self._settings.hotkey_prev, DEFAULT_SETTINGS.hotkey_prev),
            (self.HOTKEY_NEXT, self._settings.hotkey_next, DEFAULT_SETTINGS.hotkey_next),
            (
                self.HOTKEY_SHOW,
                self._settings.hotkey_show_history,
                DEFAULT_SETTINGS.hotkey_show_history,
            ),
        ]
        changed = False
        for hotkey_id, sequence, fallback in mappings:
            try:
                modifiers, key = parse_hotkey(sequence)
            except ValueError:
                modifiers, key = parse_hotkey(fallback)
                if hotkey_id == self.HOTKEY_PREV:
                    self._settings.hotkey_prev = fallback
                elif hotkey_id == self.HOTKEY_NEXT:
                    self._settings.hotkey_next = fallback
                else:
                    self._settings.hotkey_show_history = fallback
                changed = True
            try:
                self._hotkeys.register_hotkey(hotkey_id, modifiers, key)
            except RuntimeError as exc:
                QtWidgets.QMessageBox.warning(
                    None,
                    "Hotkey-Fehler",
                    f"{exc}\nDer Hotkey wird auf den Standardwert zurueckgesetzt.",
                )
                modifiers, key = parse_hotkey(fallback)
                if hotkey_id == self.HOTKEY_PREV:
                    self._settings.hotkey_prev = fallback
                elif hotkey_id == self.HOTKEY_NEXT:
                    self._settings.hotkey_next = fallback
                else:
                    self._settings.hotkey_show_history = fallback
                changed = True
                self._hotkeys.register_hotkey(hotkey_id, modifiers, key)
        if changed:
            try:
                self._settings.save(settings_path())
            except Exception:
                pass

    def _apply_capture_protection(self) -> None:
        if self._history_window is not None:
            set_window_capture_protection(self._history_window, self._settings.capture_protection_enabled)
        if self._main_window:
            set_window_capture_protection(self._main_window, self._settings.capture_protection_enabled)

    def register_main_window(self, window: "MainWindow") -> None:
        self._main_window = window
        self._main_window.apply_settings(self._settings)
        self._apply_capture_protection()

    def current_settings(self) -> AppSettings:
        return self._settings.copy()

    def apply_settings(self, new_settings: AppSettings) -> None:
        self._settings = new_settings.sanitized()
        try:
            self._settings.save(settings_path())
        except Exception:
            pass
        apply_app_theme(self._app, self._settings)
        self._toast.apply_settings(self._settings)
        if not self._settings.show_preview_overlay:
            self._toast.hide()
        else:
            current = self._clipboard_history.current_item()
            if current:
                self._toast.show_preview(current)
        if self._history_window is not None:
            self._history_window.apply_settings(self._settings)
        if self._main_window:
            self._main_window.apply_settings(self._settings)
        self._tray.setIcon(build_tray_icon(self._settings))
        self._tray.setToolTip(APP_DISPLAY_NAME)
        self._show_history_shortcut.setKey(
            QtGui.QKeySequence(self._settings.hotkey_show_history)
        )
        self._apply_capture_protection()
        self._register_hotkeys()

    def _ensure_history_window(self) -> HistoryWindow:
        if self._history_window is None:
            window = HistoryWindow(
                self._clipboard_history, self._settings, self.open_settings_dialog
            )
            window.itemActivated.connect(self._on_item_activated)
            set_window_capture_protection(window, self._settings.capture_protection_enabled)
            window.hide()
            self._history_window = window
        return self._history_window

    def open_settings_dialog(self) -> None:
        history_window = self._history_window
        parent = history_window if history_window and history_window.isVisible() else self._main_window
        dialog = SettingsDialog(parent or self._ensure_history_window(), self._settings)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            result = dialog.result_settings()
            if result:
                self.apply_settings(result)

    def mark_first_run_completed(self) -> None:
        if self._settings.first_run:
            self._settings.first_run = False
            try:
                self._settings.save(settings_path())
            except Exception:
                pass

    def _process_hotkey(self, hotkey_id: int) -> None:
        if hotkey_id == self.HOTKEY_NEXT:
            item = self._clipboard_history.select_next()
        elif hotkey_id == self.HOTKEY_PREV:
            item = self._clipboard_history.select_previous()
        elif hotkey_id == self.HOTKEY_SHOW:
            self._show_history()
            return
        else:
            return

        if item:
            self._clipboard_history.push_to_clipboard(item)
            if self._settings.show_preview_overlay:
                self._toast.show_preview(item)
            else:
                self._toast.hide()

    def _on_selection_change(self, item: Optional[ClipboardItem]) -> None:
        if item and self._settings.show_preview_overlay:
            self._toast.show_preview(item)
        elif not self._settings.show_preview_overlay:
            self._toast.hide()

    def _on_item_activated(self, item: ClipboardItem) -> None:
        self._clipboard_history.push_to_clipboard(item)
        if self._settings.show_preview_overlay:
            self._toast.show_preview(item)
        else:
            self._toast.hide()

    def _show_history(self) -> None:
        window = self._ensure_history_window()
        window.show()
        window.raise_()
        window.activateWindow()

    def show_history(self) -> None:
        self._show_history()

    def _on_tray_activated(self, reason: QtWidgets.QSystemTrayIcon.ActivationReason) -> None:
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self._show_history()

    def _quit(self) -> None:
        self._hotkeys.unregister_all()
        self._tray.hide()
        self._app.quit()


class MainWindow(QtWidgets.QWidget):
    def __init__(self, controller: MainController) -> None:
        super().__init__()
        self._controller = controller
        self._settings = controller.current_settings()
        self.setWindowFlags(
            QtCore.Qt.Window
            | QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowMinimizeButtonHint
        )
        self.setWindowTitle(APP_DISPLAY_NAME)
        self.setWindowIcon(window_icon(self._settings))
        self.resize(360, 260)

        self._drag_active = False
        self._drag_offset = QtCore.QPoint()

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        self._title_bar = QtWidgets.QFrame()
        self._title_bar.setObjectName("titleBar")
        title_layout = QtWidgets.QHBoxLayout(self._title_bar)
        title_layout.setContentsMargins(14, 8, 14, 8)
        title_layout.setSpacing(10)

        self._title_icon = QtWidgets.QLabel()
        icon_pixmap = window_icon(self._settings).pixmap(20, 20)
        self._title_icon.setPixmap(icon_pixmap)
        title_layout.addWidget(self._title_icon)

        self._title_label = QtWidgets.QLabel(APP_DISPLAY_NAME)
        self._title_label.setObjectName("titleBarLabel")
        title_layout.addWidget(self._title_label)
        title_layout.addStretch(1)

        self._minimize_button = QtWidgets.QToolButton()
        self._minimize_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarMinButton))
        self._minimize_button.clicked.connect(self._on_minimize_clicked)
        title_layout.addWidget(self._minimize_button)

        self._close_button = QtWidgets.QToolButton()
        self._close_button.setIcon(self.style().standardIcon(QtWidgets.QStyle.SP_TitleBarCloseButton))
        self._close_button.clicked.connect(self._close_to_tray)
        title_layout.addWidget(self._close_button)

        layout.addWidget(self._title_bar)
        for w in (self._title_bar, self._title_icon, self._title_label):
            w.installEventFilter(self)

        self._hero = QtWidgets.QFrame()
        self._hero.setObjectName("welcomeCard")
        hero_layout = QtWidgets.QVBoxLayout(self._hero)
        hero_layout.setContentsMargins(22, 22, 22, 22)
        hero_layout.setSpacing(10)

        title = QtWidgets.QLabel(APP_DISPLAY_NAME)
        title.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        title.setStyleSheet(
            """
            QLabel {
                font-size: 20px;
                font-weight: 700;
                background-color: transparent;
            }
            """
        )
        hero_layout.addWidget(title)

        caption = QtWidgets.QLabel(
            f"{APP_DISPLAY_NAME} begleitet dich immer im Hintergrund und bleibt ueber das Tray erreichbar."
        )
        caption.setWordWrap(True)
        caption.setStyleSheet(
            """
            QLabel {
                font-size: 12px;
                color: rgba(245, 247, 255, 150);
                background-color: transparent;
            }
            """
        )
        hero_layout.addWidget(caption)

        self._feature_label = QtWidgets.QLabel()
        self._feature_label.setStyleSheet(
            """
            QLabel {
                font-size: 12px;
                color: rgba(245, 247, 255, 190);
                background-color: transparent;
            }
            """
        )
        self._feature_label.setAlignment(QtCore.Qt.AlignLeft)
        hero_layout.addWidget(self._feature_label)

        self._status_chip = QtWidgets.QLabel(f"{APP_DISPLAY_NAME} laeuft im Hintergrund")
        self._status_chip.setAlignment(QtCore.Qt.AlignCenter)
        self._status_chip.setStyleSheet(
            """
            QLabel {
                border-radius: 12px;
                padding: 6px 12px;
                background-color: rgba(44, 182, 125, 70);
                color: rgba(44, 182, 125, 210);
                font-weight: 600;
                letter-spacing: 0.5px;
            }
            """
        )
        hero_layout.addWidget(self._status_chip)
        layout.addWidget(self._hero)

        action_row = QtWidgets.QHBoxLayout()
        action_row.setSpacing(14)

        self._history_button = QtWidgets.QPushButton("Verlauf oeffnen")
        self._history_button.clicked.connect(self._controller.show_history)
        self._history_button.setProperty("accent", True)
        self._history_button.style().unpolish(self._history_button)
        self._history_button.style().polish(self._history_button)
        action_row.addWidget(self._history_button, 1)

        self._hide_button = QtWidgets.QPushButton("Schliessen")
        self._hide_button.clicked.connect(self._close_to_tray)
        action_row.addWidget(self._hide_button)

        layout.addLayout(action_row)
        self.apply_settings(self._settings)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        event.ignore()
        self.hide()

    def apply_settings(self, settings: AppSettings) -> None:
        self._settings = settings
        start = QtGui.QColor(settings.accent_start)
        if not start.isValid():
            start = QtGui.QColor("#7f5af0")
        border = color_to_rgba(start, 0.4)
        self._hero.setStyleSheet(
            f"""
            QFrame#welcomeCard {{
                background-color: rgba(24, 26, 40, 235);
                border: 1px solid {border};
                border-radius: 18px;
            }}
            """
        )
        self.setWindowIcon(window_icon(settings))
        self._title_icon.setPixmap(window_icon(settings).pixmap(20, 20))
        self._status_chip.setText(f"{APP_DISPLAY_NAME} laeuft im Hintergrund")
        preview_state = "aktiv" if settings.show_preview_overlay else "deaktiviert"
        capture_state = "aktiv" if settings.capture_protection_enabled else "deaktiviert"
        lines = [
            f"\u2022 {display_hotkey(settings.hotkey_prev)} / {display_hotkey(settings.hotkey_next)}: Verlauf wechseln",
            f"\u2022 {display_hotkey(settings.hotkey_show_history)}: Verlauf oeffnen",
            f"\u2022 Vorschau-Overlay: {preview_state}",
            f"\u2022 Aufnahmeschutz: {capture_state}",
            "\u2022 Enter in der Liste: Auswahl uebernehmen",
        ]
        self._feature_label.setText("\n".join(lines))
        for button in (self._history_button, self._hide_button):
            button.style().unpolish(button)
            button.style().polish(button)

    def _on_minimize_clicked(self) -> None:
        self.showMinimized()

    def _close_to_tray(self) -> None:
        self.hide()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if obj in (self._title_bar, self._title_label, self._title_icon):
            if event.type() == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
                self._drag_active = True
                self._drag_offset = self._event_global_pos(event) - self.frameGeometry().topLeft()
                return True
            if event.type() == QtCore.QEvent.MouseMove and self._drag_active and event.buttons() & QtCore.Qt.LeftButton:
                self.move(self._event_global_pos(event) - self._drag_offset)
                return True
            if event.type() == QtCore.QEvent.MouseButtonRelease and event.button() == QtCore.Qt.LeftButton:
                self._drag_active = False
                return True
        return super().eventFilter(obj, event)

    @staticmethod
    def _event_global_pos(event: QtCore.QEvent) -> QtCore.QPoint:
        if hasattr(event, "globalPosition"):
            return event.globalPosition().toPoint()
        return event.globalPos()


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    settings = AppSettings.load(settings_path()).sanitized()
    install_target = perform_initial_install(settings)
    apply_app_theme(app, settings)
    app.setWindowIcon(load_app_icon())
    controller = MainController(app, settings)
    ensure_autostart(install_target)
    window = MainWindow(controller)
    controller.register_main_window(window)
    if settings.first_run:
        window.show()
        controller.mark_first_run_completed()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
