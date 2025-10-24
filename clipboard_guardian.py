import base64
import math
import json
import os
import re
import shutil
import string
import sys
import time
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

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

OVERLAY_THEMES = ("classic", "glass", "minimal")

BASE_DIR = Path(sys.argv[0]).resolve().parent
ICON_ICO_PATH = BASE_DIR / "logo.ico"
ICON_PNG_PATH = BASE_DIR / "logo.png"


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def decode_bytes_to_text(payload: bytes) -> str:
    for encoding in ("utf-8", "utf-16le", "utf-16", "latin1"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


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
        cls.configure_palette(settings)

    @classmethod
    def configure_palette(cls, settings: "AppSettings") -> None:
        if getattr(settings, "theme_mode", "dark") == "light":
            cls.PRIMARY_BG = QtGui.QColor("#f4f6fb")
            cls.CARD_BG = QtGui.QColor("#ffffff")
            cls.TEXT_PRIMARY = QtGui.QColor("#222330")
            cls.TEXT_MUTED = QtGui.QColor("#5d6378")
            cls.BORDER = QtGui.QColor("#d5d9e8")
        else:
            cls.PRIMARY_BG = QtGui.QColor("#12131c")
            cls.CARD_BG = QtGui.QColor("#1e2130")
            cls.TEXT_PRIMARY = QtGui.QColor("#f5f7ff")
            cls.TEXT_MUTED = QtGui.QColor("#9aa3c0")
            cls.BORDER = QtGui.QColor("#2b2f44")


def color_to_rgba(color: QtGui.QColor, alpha: float) -> str:
    alpha = clamp(alpha, 0.0, 1.0)
    return f"rgba({color.red()},{color.green()},{color.blue()},{int(alpha * 255)})"


def apply_app_theme(app: QtWidgets.QApplication, settings: "AppSettings") -> None:
    Theme.apply_settings(settings)

    app.setStyle("Fusion")
    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, Theme.PRIMARY_BG)
    palette.setColor(QtGui.QPalette.Base, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.AlternateBase, Theme.CARD_BG.darker(102))
    palette.setColor(QtGui.QPalette.ToolTipBase, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.ToolTipText, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Text, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Button, Theme.CARD_BG)
    palette.setColor(QtGui.QPalette.ButtonText, Theme.TEXT_PRIMARY)
    palette.setColor(QtGui.QPalette.Highlight, Theme.ACCENT)
    palette.setColor(
        QtGui.QPalette.HighlightedText,
        QtCore.Qt.white if settings.theme_mode == "dark" else QtGui.QColor("#1e2131"),
    )
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

    base_hex = Theme.PRIMARY_BG.name()
    card_hex = Theme.CARD_BG.name()
    text_hex = Theme.TEXT_PRIMARY.name()
    border_hex = Theme.BORDER.name()
    muted_hex = Theme.TEXT_MUTED.name()
    title_bar_bg = color_to_rgba(QtGui.QColor(24, 26, 40) if settings.theme_mode == "dark" else QtGui.QColor(247, 248, 255), 0.92)
    title_text = "rgba(245, 247, 255, 230)" if settings.theme_mode == "dark" else "rgba(36, 38, 58, 230)"

    global_stylesheet = f"""
    QWidget {{
        background-color: {base_hex};
        color: {text_hex};
    }}
    QLineEdit, QTextEdit {{
        background-color: {card_hex};
        border: 1px solid {border_hex};
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
        color: {text_hex};
    }}
    QPushButton:hover {{
        background-color: {color_to_rgba(accent_hover, 0.24)};
        border: 1px solid {color_to_rgba(accent_hover, 0.55)};
    }}
    QPushButton:pressed {{
        background-color: {color_to_rgba(accent_pressed, 0.5)};
    }}
    QToolButton {{
        color: {text_hex};
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
        background-color: {title_bar_bg};
        border-radius: 14px;
        border: 1px solid {color_to_rgba(accent, 0.35)};
    }}
    QLabel#titleBarLabel {{
        color: {title_text};
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }}
    QPushButton[accent="true"] {{
        background-color: {color_to_rgba(accent, 0.8)};
        border: 1px solid {color_to_rgba(accent, 0.95)};
        color: #f5f7ff;
    }}
    QPushButton[accent="true"]:hover {{
        background-color: {color_to_rgba(accent_hover, 0.9)};
    }}
    QPushButton[destructive="true"] {{
        background-color: rgba(255, 99, 71, 0.18);
        border: 1px solid rgba(255, 99, 71, 0.5);
        color: rgba(255, 143, 119, 0.95);
    }}
    QPushButton[destructive="true"]:hover {{
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
    theme_mode: str = "dark"
    overlay_theme: str = "classic"
    overlay_opacity: int = 90
    overlay_follow_mouse: bool = False
    overlay_anchor: str = "bottom-right"
    overlay_offset_x: int = 24
    overlay_offset_y: int = 24
    animation_in_ms: int = 240
    animation_out_ms: int = 200
    auto_clear_enabled: bool = False
    auto_clear_interval_minutes: int = 1440
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
        interval_value = int(clamp(self.auto_clear_interval_minutes, 5, 10080))
        theme_value = self.theme_mode if self.theme_mode in ("dark", "light") else "dark"
        anchor_value = self.overlay_anchor if self.overlay_anchor in ("top-left", "top-right", "bottom-left", "bottom-right") else "bottom-right"
        opacity_value = int(clamp(self.overlay_opacity, 20, 100))
        offset_x_value = int(clamp(self.overlay_offset_x, -500, 500))
        offset_y_value = int(clamp(self.overlay_offset_y, -500, 500))
        anim_in_value = int(clamp(self.animation_in_ms, 50, 2000))
        anim_out_value = int(clamp(self.animation_out_ms, 50, 2000))
        overlay_theme_value = self.overlay_theme if self.overlay_theme in OVERLAY_THEMES else "classic"
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
            theme_mode=theme_value,
            overlay_theme=overlay_theme_value,
            overlay_opacity=opacity_value,
            overlay_follow_mouse=bool(self.overlay_follow_mouse),
            overlay_anchor=anchor_value,
            overlay_offset_x=offset_x_value,
            overlay_offset_y=offset_y_value,
            animation_in_ms=anim_in_value,
            animation_out_ms=anim_out_value,
            auto_clear_enabled=bool(self.auto_clear_enabled),
            auto_clear_interval_minutes=interval_value,
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
    html: Optional[str] = None
    image_data: Optional[str] = None
    urls: List[str] = field(default_factory=list)
    files: List[str] = field(default_factory=list)
    format: str = "text"
    pinned: bool = False
    rtf_data: Optional[str] = None
    csv_data: Optional[str] = None

    @classmethod
    def from_dict(cls, data: dict) -> "ClipboardItem":
        return cls(
            content=data.get("content", ""),
            timestamp=data.get("timestamp", time.time()),
            html=data.get("html"),
            image_data=data.get("image_data"),
            urls=data.get("urls", []) or [],
            files=data.get("files", []) or [],
            format=data.get("format", "text"),
            pinned=data.get("pinned", False),
            rtf_data=data.get("rtf_data"),
            csv_data=data.get("csv_data"),
        )


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
        self._trim_history()
        self._current_index: Optional[int] = 0 if self._history else None
        self._suspend_capture = False
        self._clipboard.dataChanged.connect(self._on_clipboard_change)

    def _on_clipboard_change(self) -> None:
        if self._suspend_capture:
            return
        mime = self._clipboard.mimeData()
        new_item = self._create_item_from_mime(mime)
        if not new_item:
            return

        if self._history and self._items_equal(self._ordered_items()[0], new_item):
            return

        self._history.insert(0, new_item)
        self._trim_history()
        self._current_index = 0
        self._persist()
        self.historyUpdated.emit(self.all_items())
        self.selectionChanged.emit(self.current_item())

    def _persist(self) -> None:
        serializable = [asdict(entry) for entry in self._history]
        self._storage.save(serializable)

    def current_item(self) -> Optional[ClipboardItem]:
        ordered = self._ordered_items()
        if not ordered:
            return None
        if self._current_index is None:
            self._current_index = 0
        self._current_index = max(0, min(self._current_index, len(ordered) - 1))
        return ordered[self._current_index]

    def select_next(self) -> Optional[ClipboardItem]:
        ordered = self._ordered_items()
        if not ordered:
            return None
        if self._current_index is None:
            self._current_index = 0
        else:
            self._current_index = (self._current_index + 1) % len(ordered)
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def select_previous(self) -> Optional[ClipboardItem]:
        ordered = self._ordered_items()
        if not ordered:
            return None
        if self._current_index is None:
            self._current_index = 0
        else:
            self._current_index = (self._current_index - 1) % len(ordered)
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def select_index(self, index: int) -> Optional[ClipboardItem]:
        ordered = self._ordered_items()
        if not ordered:
            return None
        index = max(0, min(index, len(ordered) - 1))
        self._current_index = index
        self.selectionChanged.emit(self.current_item())
        return self.current_item()

    def remove_entry(self, entry: ClipboardItem) -> None:
        if entry.pinned:
            return
        if entry not in self._history:
            return
        for idx, existing in enumerate(self._history):
            if existing == entry:
                del self._history[idx]
                self._persist()
                ordered = self._ordered_items()
                if not ordered:
                    self._current_index = None
                else:
                    self._current_index = min(self._current_index or 0, len(ordered) - 1)
                self.historyUpdated.emit(self.all_items())
                self.selectionChanged.emit(self.current_item())
                return

    def clear(self) -> None:
        if not self._history:
            return
        pinned_items = [item for item in self._history if item.pinned]
        if len(pinned_items) == len(self._history):
            return
        self._history = pinned_items
        self._current_index = None
        self._persist()
        self.historyUpdated.emit(self.all_items())
        self.selectionChanged.emit(None)

    def push_to_clipboard(self, item: ClipboardItem) -> None:
        self._suspend_capture = True
        mime = QtCore.QMimeData()
        if item.html:
            mime.setHtml(item.html)
        if item.content:
            mime.setText(item.content)
        if item.rtf_data:
            try:
                raw = base64.b64decode(item.rtf_data)
                mime.setData("text/rtf", raw)
                mime.setData('application/x-qt-windows-mime;value="Rich Text Format"', raw)
            except Exception:
                pass
        if item.csv_data:
            try:
                raw_csv = base64.b64decode(item.csv_data)
                mime.setData("text/csv", raw_csv)
                mime.setData("application/csv", raw_csv)
                mime.setData('application/x-qt-windows-mime;value="Csv"', raw_csv)
            except Exception:
                pass
        if item.image_data:
            try:
                image_bytes = base64.b64decode(item.image_data)
                image = QtGui.QImage()
                image.loadFromData(image_bytes, "PNG")
                mime.setImageData(image)
            except Exception:
                pass
        urls = []
        for path in item.files:
            urls.append(QtCore.QUrl.fromLocalFile(path))
        for url in item.urls:
            if url:
                urls.append(QtCore.QUrl(url))
        if urls:
            mime.setUrls(urls)
        if mime.formats():
            self._clipboard.setMimeData(mime)
        else:
            self._clipboard.setText(item.content)
        QtCore.QTimer.singleShot(100, self._resume_capture)

    def _resume_capture(self) -> None:
        self._suspend_capture = False

    def all_items(self) -> List[ClipboardItem]:
        return self._ordered_items().copy()

    def toggle_pin(self, entry: ClipboardItem) -> None:
        if entry not in self._history:
            return
        entry.pinned = not entry.pinned
        ordered = self._ordered_items()
        if entry in ordered:
            self._current_index = ordered.index(entry)
        self._persist()
        self.historyUpdated.emit(self.all_items())
        self.selectionChanged.emit(self.current_item())

    def _ordered_items(self) -> List[ClipboardItem]:
        pinned = [item for item in self._history if item.pinned]
        others = [item for item in self._history if not item.pinned]
        return pinned + others

    def _trim_history(self) -> None:
        unpinned = [item for item in self._history if not item.pinned]
        while len(unpinned) > MAX_HISTORY_ITEMS:
            to_remove = unpinned.pop()
            try:
                self._history.remove(to_remove)
            except ValueError:
                break

    def _create_item_from_mime(self, mime: Optional[QtCore.QMimeData]) -> Optional[ClipboardItem]:
        if not mime:
            return None
        text = mime.text() if mime.hasText() else ""
        html = mime.html() if mime.hasHtml() else None
        if (not html or not html.strip()) and mime.hasFormat("text/html"):
            try:
                raw_html = mime.data("text/html")
                if raw_html:
                    html = decode_bytes_to_text(bytes(raw_html))
            except Exception:
                html = html or None
        image_data = None
        if mime.hasImage():
            try:
                image = mime.imageData()
                if isinstance(image, QtGui.QPixmap):
                    qimage = image.toImage()
                elif isinstance(image, QtGui.QImage):
                    qimage = image
                else:
                    qimage = QtGui.QImage(image)
                buffer = QtCore.QBuffer()
                buffer.open(QtCore.QIODevice.WriteOnly)
                qimage.save(buffer, "PNG")
                image_data = base64.b64encode(bytes(buffer.data())).decode("ascii")
            except Exception:
                image_data = None
        urls: List[str] = []
        files: List[str] = []
        if mime.hasUrls():
            for qurl in mime.urls():
                if qurl.isLocalFile():
                    files.append(qurl.toLocalFile())
                urls.append(qurl.toString())
        rtf_data = None
        for fmt in (
            "text/rtf",
            'application/x-qt-windows-mime;value="Rich Text Format"',
            'application/x-qt-windows-mime;value="RTF"',
        ):
            if mime.hasFormat(fmt):
                try:
                    raw_rtf = mime.data(fmt)
                    if raw_rtf:
                        rtf_data = base64.b64encode(bytes(raw_rtf)).decode("ascii")
                        break
                except Exception:
                    continue
        csv_data = None
        for fmt in (
            "text/csv",
            "application/csv",
            "application/x-qt-windows-mime;value=\"Csv\"",
            "application/x-qt-windows-mime;value=\"CSV\"",
        ):
            if mime.hasFormat(fmt):
                try:
                    raw_csv = mime.data(fmt)
                    if raw_csv:
                        buffer = bytes(raw_csv)
                        csv_data = base64.b64encode(buffer).decode("ascii")
                        break
                except Exception:
                    continue
        format_type = "text"
        if image_data:
            format_type = "image"
        elif csv_data:
            format_type = "table"
        elif html and "<table" in html.lower():
            format_type = "table"
        elif html and html.strip():
            format_type = "html"
        elif rtf_data:
            format_type = "rich"
        elif files:
            format_type = "files"
        elif urls:
            format_type = "urls"
        elif text.strip():
            format_type = "text"
        else:
            return None
        return ClipboardItem(
            content=text,
            timestamp=time.time(),
            html=html,
            image_data=image_data,
            urls=urls,
            files=files,
            format=format_type,
            rtf_data=rtf_data,
            csv_data=csv_data,
        )

    def _items_equal(self, a: ClipboardItem, b: ClipboardItem) -> bool:
        if a.format != b.format:
            return False
        if a.format == "image":
            return a.image_data == b.image_data
        if a.format == "table":
            return (
                a.csv_data == b.csv_data
                and (a.html or "") == (b.html or "")
                and a.content == b.content
            )
        if a.format == "rich":
            return a.rtf_data == b.rtf_data and a.content == b.content
        if a.format == "html":
            return a.html == b.html and a.content == b.content
        if a.format == "files":
            return set(a.files) == set(b.files)
        if a.format == "urls":
            return set(a.urls) == set(b.urls)
        return a.content == b.content


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
        self._slide = QtCore.QPropertyAnimation(self, b"pos", self)
        self._slide.setEasingCurve(QtCore.QEasingCurve.OutCubic)

        self._follow_timer = QtCore.QTimer(self)
        self._follow_timer.setInterval(70)
        self._follow_timer.timeout.connect(self._update_follow_position)

        self._current_pos: Optional[QtCore.QPoint] = None
        self._current_geom: Optional[QtCore.QRect] = None
        self._active_item: Optional[ClipboardItem] = None
        self._pixmap_cache: Dict[str, QtGui.QPixmap] = {}

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(0)

        self._card_container = QtWidgets.QFrame()
        self._card_container.setObjectName("toastContainer")
        self._card_container.setStyleSheet("QFrame { background-color: transparent; }")
        container_layout = QtWidgets.QVBoxLayout(self._card_container)
        container_layout.setContentsMargins(8, 8, 8, 8)
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

        self._title_label = QtWidgets.QLabel("Clipboard Preview")
        self._title_label.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self._title_label.setStyleSheet(
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
        title_row.addWidget(self._title_label)
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

        self._image_label = QtWidgets.QLabel()
        self._image_label.setAlignment(QtCore.Qt.AlignCenter)
        self._image_label.setVisible(False)
        card_layout.addWidget(self._image_label)

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
        layout.addWidget(self._card_container)

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
            self._follow_timer.stop()
            super().hide()
            self._current_pos = None
            return
        start = QtGui.QColor(self._settings.accent_start)
        if not start.isValid():
            start = QtGui.QColor("#7f5af0")
        end = QtGui.QColor(self._settings.accent_end)
        if not end.isValid():
            end = QtGui.QColor("#2cb67d")
        dark_mode = self._settings.theme_mode == "dark"
        overlay_theme = (
            self._settings.overlay_theme
            if getattr(self._settings, "overlay_theme", "classic") in OVERLAY_THEMES
            else "classic"
        )
        base_bg = QtGui.QColor("#181a28") if dark_mode else QtGui.QColor("#ffffff")
        opacity = clamp(self._settings.overlay_opacity / 100.0, 0.2, 1.0)

        if overlay_theme == "glass":
            surface_bg = QtGui.QColor(35, 38, 54) if dark_mode else QtGui.QColor(255, 255, 255)
            background_color = color_to_rgba(surface_bg, max(0.45, opacity * 0.7))
            border_rule = f"border: 1px solid {color_to_rgba(QtGui.QColor(255, 255, 255), 0.26 if dark_mode else 0.22)};"
            halo_color = color_to_rgba(QtGui.QColor(18, 20, 32), 0.55 if dark_mode else 0.35)
            container_style = f"QFrame {{ background-color: {halo_color}; border-radius: 20px; }}"
        elif overlay_theme == "minimal":
            background_color = color_to_rgba(base_bg, max(0.5, opacity))
            border_rule = f"border: 1px solid {color_to_rgba(start, 0.2)};"
            container_style = "QFrame { background-color: transparent; border-radius: 20px; }"
        else:
            background_color = color_to_rgba(base_bg, opacity)
            border_rule = f"border: 1px solid {color_to_rgba(start, 0.45)};"
            halo_color = color_to_rgba(QtGui.QColor(11, 12, 20), 0.6 if dark_mode else 0.18)
            container_style = f"QFrame {{ background-color: {halo_color}; border-radius: 20px; }}"

        if isinstance(getattr(self, "_card_container", None), QtWidgets.QFrame):
            self._card_container.setStyleSheet(container_style)
        self._card.setStyleSheet(
            f"""
            QFrame#toastCard {{
                background-color: {background_color};
                border-radius: 16px;
                {border_rule}
            }}
            """
        )
        if overlay_theme == "minimal":
            accent_style = (
                f"""
                QLabel {{
                    border-radius: 3px;
                    background-color: {start.name()};
                }}
                """
            )
        else:
            accent_style = (
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
        self._accent_icon.setStyleSheet(accent_style)

        text_color = "#f5f7ff" if dark_mode else "#222330"
        hint_color = "rgba(154, 163, 192, 200)" if dark_mode else "rgba(72, 80, 98, 200)"
        title_color = "rgba(245, 247, 255, 180)" if dark_mode else "rgba(44, 48, 68, 200)"
        if overlay_theme == "glass" and not dark_mode:
            text_color = "#1f2338"
            hint_color = "rgba(62, 70, 96, 200)"
            title_color = "rgba(28, 30, 48, 220)"

        self._title_label.setStyleSheet(
            f"""
            QLabel {{
                font-size: 11px;
                letter-spacing: 1px;
                text-transform: uppercase;
                color: {title_color};
                background-color: transparent;
            }}
            """
        )
        self._label.setStyleSheet(
            f"""
            QLabel {{
                color: {text_color};
                font-size: 15px;
                font-weight: 500;
                background-color: transparent;
            }}
            """
        )
        self._hint.setStyleSheet(
            f"""
            QLabel {{
                color: {hint_color};
                font-size: 11px;
                background-color: transparent;
            }}
            """
        )

    def _start_fade_out(self) -> None:
        self._timer.stop()
        self._start_hide_animation()
        self._animate_opacity(0.0, hide_after=True, duration=self._settings.animation_out_ms)

    def _on_fade_finished(self) -> None:
        if self._fade_target == 0.0:
            super().hide()
            self._current_pos = None

    def show_preview(self, item: ClipboardItem) -> None:
        self._timer.stop()
        self._follow_timer.stop()
        self._fade.stop()
        self._slide.stop()
        self._active_item = item
        preview_text = self._format_preview_text(item)
        if item.image_data:
            pixmap = self._get_pixmap(item.image_data)
            if pixmap and not pixmap.isNull():
                scaled = pixmap.scaled(320, 220, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                self._image_label.setPixmap(scaled)
                self._image_label.setVisible(True)
            else:
                self._image_label.clear()
                self._image_label.setVisible(False)
        else:
            self._image_label.clear()
            self._image_label.setVisible(False)
        if preview_text:
            self._label.setText(preview_text)
            self._label.setVisible(True)
        else:
            self._label.clear()
            self._label.setVisible(False)
        cursor_pos = QtGui.QCursor.pos()
        screen = QtGui.QGuiApplication.screenAt(cursor_pos)
        if screen:
            available = screen.availableGeometry()
        else:
            available = QtGui.QGuiApplication.primaryScreen().availableGeometry()
        self._current_geom = available
        self.ensurePolished()
        size_hint = self.sizeHint()
        min_hint = self.minimumSizeHint()
        scale = float(clamp(self._settings.toast_scale, 0.6, 1.6))
        base_width = max(size_hint.width(), min_hint.width(), 260)
        base_height = max(size_hint.height(), min_hint.height(), 160)
        width = int(clamp(base_width * scale, base_width, 520))
        height = int(clamp(base_height * scale, base_height, 360))
        self.resize(width, height)
        target_pos = self._calculate_target_position(cursor_pos, available)
        self._current_pos = target_pos
        if self._settings.overlay_follow_mouse:
            start_pos = QtCore.QPoint(target_pos.x(), target_pos.y() + 18)
        else:
            start_pos = QtCore.QPoint(target_pos)
        if not self.isVisible():
            self.move(start_pos)
            self._opacity_effect.setOpacity(0.0)
            super().show()
            self.raise_()
            self._start_show_animation(target_pos, duration=self._settings.animation_in_ms, start_pos=start_pos)
        else:
            super().show()
            self.raise_()
            self._start_show_animation(target_pos, duration=self._settings.animation_in_ms)
        if self._settings.overlay_follow_mouse:
            self._follow_timer.start(70)
        else:
            self._follow_timer.stop()
        duration_ms = int(clamp(self._settings.toast_duration_ms, 600, 10000))
        self._timer.start(duration_ms)

    def hide(self) -> None:
        self._follow_timer.stop()
        self._start_fade_out()

    def _start_show_animation(self, target_pos: QtCore.QPoint, duration: int = 220, start_pos: Optional[QtCore.QPoint] = None) -> None:
        self._slide.stop()
        if start_pos is None:
            start_pos = self.pos()
        self._slide.setStartValue(start_pos)
        self._slide.setEndValue(target_pos)
        self._slide.setDuration(duration)
        self._slide.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        self._slide.start()
        self._animate_opacity(1.0, hide_after=False, duration=duration)

    def _start_hide_animation(self) -> None:
        self._follow_timer.stop()
        if self._current_pos is None:
            return
        end_pos = self._current_pos + QtCore.QPoint(0, 18)
        self._slide.stop()
        self._slide.setStartValue(self.pos())
        self._slide.setEndValue(end_pos)
        self._slide.setDuration(self._settings.animation_out_ms)
        self._slide.setEasingCurve(QtCore.QEasingCurve.InCubic)
        self._slide.start()

    def _animate_opacity(self, value: float, hide_after: bool, duration: Optional[int] = None) -> None:
        self._fade.stop()
        self._fade_target = value
        self._fade.setStartValue(self._opacity_effect.opacity())
        self._fade.setEndValue(value)
        if duration is None:
            duration = self._settings.animation_out_ms if hide_after else self._settings.animation_in_ms
        self._fade.setDuration(duration)
        self._fade.start()

    def _calculate_target_position(self, cursor_pos: QtCore.QPoint, available: QtCore.QRect) -> QtCore.QPoint:
        width = self.width()
        height = self.height()
        offset_x = self._settings.overlay_offset_x
        offset_y = self._settings.overlay_offset_y
        if self._settings.overlay_follow_mouse:
            x = cursor_pos.x() + offset_x
            y = cursor_pos.y() + offset_y
        else:
            anchor = self._settings.overlay_anchor
            if anchor == "top-left":
                base_x = available.left()
                base_y = available.top()
            elif anchor == "top-right":
                base_x = available.right() - width
                base_y = available.top()
            elif anchor == "bottom-left":
                base_x = available.left()
                base_y = available.bottom() - height
            else:
                base_x = available.right() - width
                base_y = available.bottom() - height
            x = base_x + offset_x
            y = base_y + offset_y
        x = max(available.left(), min(x, available.right() - width))
        y = max(available.top(), min(y, available.bottom() - height))
        return QtCore.QPoint(x, y)

    def _update_follow_position(self) -> None:
        if not self._settings.overlay_follow_mouse or self._current_geom is None:
            return
        cursor_pos = QtGui.QCursor.pos()
        screen = QtGui.QGuiApplication.screenAt(cursor_pos)
        if screen:
            self._current_geom = screen.availableGeometry()
        available = self._current_geom or QtGui.QGuiApplication.primaryScreen().availableGeometry()
        target_pos = self._calculate_target_position(cursor_pos, available)
        self._current_pos = target_pos
        self.move(target_pos)

    def _format_preview_text(self, item: ClipboardItem) -> str:
        if item.format == "image":
            return "Bildvorschau"
        if item.format == "files":
            return "\n".join(Path(path_value).name for path_value in item.files[:4])
        if item.format == "urls":
            return "\n".join(item.urls[:4])
        if item.format == "table":
            snippet = ""
            if item.csv_data:
                try:
                    raw_csv = base64.b64decode(item.csv_data)
                    snippet = decode_bytes_to_text(raw_csv)
                except Exception:
                    snippet = ""
            if not snippet and item.html:
                doc = QtGui.QTextDocument()
                doc.setHtml(item.html)
                snippet = doc.toPlainText()
            if not snippet:
                snippet = item.content
        elif item.format in ("html", "rich") and item.html:
            doc = QtGui.QTextDocument()
            doc.setHtml(item.html)
            snippet = doc.toPlainText()
        else:
            snippet = item.content
        snippet = snippet.replace("\r", " ").replace("\n", " ")
        snippet = re.sub(r'(\S{60})', r'\1' + '\u200b', snippet)
        if len(snippet) > 340:
            snippet = snippet[:340] + "..."
        return snippet or "<leer>"

    def _get_pixmap(self, data: str) -> Optional[QtGui.QPixmap]:
        if not data:
            return None
        pixmap = self._pixmap_cache.get(data)
        if pixmap is None:
            try:
                pixmap = QtGui.QPixmap()
                pixmap.loadFromData(base64.b64decode(data), "PNG")
                self._pixmap_cache[data] = pixmap
            except Exception:
                return None
        return pixmap


class HistoryDelegate(QtWidgets.QStyledItemDelegate):
    actionTriggered = QtCore.Signal(str, ClipboardItem)

    def __init__(self, settings: AppSettings, parent: Optional[QtCore.QObject] = None) -> None:
        super().__init__(parent)
        self._settings = settings
        self._pixmap_cache: Dict[str, QtGui.QPixmap] = {}

    def update_settings(self, settings: AppSettings) -> None:
        self._settings = settings
        self._pixmap_cache.clear()

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

        card_rect = option.rect.adjusted(10, 6, -10, -6)
        is_selected = option.state & QtWidgets.QStyle.State_Selected
        dark_mode = self._settings.theme_mode == "dark"

        accent_color = QtGui.QColor(self._settings.accent_start)
        if not accent_color.isValid():
            accent_color = QtGui.QColor("#7f5af0")
        alt_accent = QtGui.QColor(self._settings.accent_end)
        if not alt_accent.isValid():
            alt_accent = QtGui.QColor("#2cb67d")

        base_color = QtGui.QColor(34, 38, 54) if dark_mode else QtGui.QColor(248, 249, 254)
        border_color = QtGui.QColor(58, 63, 85) if dark_mode else QtGui.QColor(215, 220, 235)
        text_color = QtGui.QColor(244, 246, 255) if dark_mode else QtGui.QColor(31, 35, 55)
        muted_color = QtGui.QColor(158, 166, 190) if dark_mode else QtGui.QColor(117, 124, 146)

        fill_color = QtGui.QColor(base_color)
        fill_color.setAlpha(235 if dark_mode else 255)
        if entry.pinned and not is_selected:
            fill_color = QtGui.QColor(accent_color)
            fill_color.setAlpha(55 if dark_mode else 85)
            border_color = accent_color
        if is_selected:
            fill_color = QtGui.QColor(accent_color)
            fill_color.setAlpha(120 if dark_mode else 160)
            border_color = accent_color

        painter.setPen(QtGui.QPen(border_color, 1))
        painter.setBrush(fill_color)
        painter.drawRoundedRect(card_rect, 14, 14)

        star_rect = self._star_rect(card_rect)
        delete_rect = self._delete_rect(card_rect)

        timestamp = time.strftime("%d.%m.%Y %H:%M", time.localtime(entry.timestamp))
        timestamp_font = QtGui.QFont(option.font)
        timestamp_font.setPointSize(max(option.font.pointSize() - 1, 8))
        painter.setFont(timestamp_font)
        painter.setPen(muted_color)
        timestamp_metrics = QtGui.QFontMetrics(timestamp_font)
        timestamp_height = timestamp_metrics.height()
        timestamp_rect = QtCore.QRect(
            card_rect.left() + 24,
            card_rect.top() + 18,
            timestamp_metrics.horizontalAdvance(timestamp) + 4,
            timestamp_height,
        )
        painter.drawText(
            timestamp_rect,
            QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter,
            timestamp,
        )

        format_label = self._format_label(entry)
        format_font = QtGui.QFont(option.font)
        format_font.setPointSize(max(option.font.pointSize() - 1, 8))
        format_font.setBold(True)
        painter.setFont(format_font)
        metrics = QtGui.QFontMetrics(format_font)
        pill_width = metrics.horizontalAdvance(format_label) + 24
        pill_width = min(pill_width, max(card_rect.width() - 48, 60))
        pill_height = 26
        available_pill_x = star_rect.left() - pill_width - 12
        desired_pill_x = max(timestamp_rect.right() + 12, card_rect.left() + 24)
        pill_x = min(desired_pill_x, available_pill_x)
        pill_y = timestamp_rect.top() + (timestamp_rect.height() - pill_height) // 2
        pill_rect = QtCore.QRect(pill_x, pill_y, pill_width, pill_height)
        painter.setPen(QtCore.Qt.NoPen)
        pill_color = QtGui.QColor(accent_color)
        pill_color.setAlpha(200 if dark_mode else 220)
        painter.setBrush(pill_color if entry.pinned or is_selected else QtGui.QColor(accent_color.lighter(125)))
        painter.drawRoundedRect(pill_rect, 12, 12)
        painter.setPen(QtCore.Qt.white if entry.pinned or is_selected else QtGui.QColor("#1e2230"))
        painter.drawText(pill_rect, QtCore.Qt.AlignCenter, format_label)

        icon_rect = QtCore.QRect(card_rect.left() + 24, card_rect.top() + 56, 28, 28)

        has_image = entry.format == "image" and entry.image_data
        content_top = card_rect.top() + 52
        content_height = max(card_rect.height() - 72, 24)
        if has_image:
            available_width = max(card_rect.width() - 64 - 120, 140)
            content_rect = QtCore.QRect(
                card_rect.left() + 64,
                content_top,
                available_width,
                content_height,
            )
            thumb_rect = QtCore.QRect(card_rect.right() - 104, card_rect.top() + 48, 88, 88)
        else:
            available_width = max(card_rect.width() - 88, 160)
            content_rect = QtCore.QRect(
                card_rect.left() + 64,
                content_top,
                available_width,
                content_height,
            )
            thumb_rect = None

        body_font = QtGui.QFont(option.font)
        body_font.setPointSize(option.font.pointSize() + 1)
        snippet_text = self._preview_text(entry)
        text_option = QtGui.QTextOption()
        text_option.setWrapMode(QtGui.QTextOption.WordWrap)
        text_option.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        painter.save()
        painter.setClipRect(content_rect)
        painter.setPen(text_color)
        painter.setFont(body_font)
        painter.drawText(QtCore.QRectF(content_rect), snippet_text, text_option)
        painter.restore()

        self._draw_format_icon(painter, icon_rect, entry, text_color)
        self._draw_star_icon(painter, star_rect, entry.pinned)
        self._draw_delete_icon(painter, delete_rect)

        if has_image and thumb_rect is not None:
            painter.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)
            pixmap = self._get_pixmap(entry.image_data)
            if pixmap and not pixmap.isNull():
                scaled = pixmap.scaled(
                    thumb_rect.size(),
                    QtCore.Qt.KeepAspectRatio,
                    QtCore.Qt.SmoothTransformation,
                )
                target = QtCore.QRect(
                    thumb_rect.left() + (thumb_rect.width() - scaled.width()) // 2,
                    thumb_rect.top() + (thumb_rect.height() - scaled.height()) // 2,
                    scaled.width(),
                    scaled.height(),
                )
                painter.drawPixmap(target, scaled)
                preview_frame = QtGui.QColor(text_color)
                preview_frame.setAlpha(45 if dark_mode else 90)
                painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
                painter.setPen(QtGui.QPen(preview_frame, 1))
                painter.setBrush(QtCore.Qt.NoBrush)
                painter.drawRoundedRect(thumb_rect.adjusted(0, 0, -1, -1), 10, 10)

        painter.restore()

    def sizeHint(
        self,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> QtCore.QSize:
        entry = index.data(QtCore.Qt.UserRole)
        if not isinstance(entry, ClipboardItem):
            return super().sizeHint(option, index)

        view_width = 0
        if option.widget is not None and hasattr(option.widget, "viewport"):
            view_width = option.widget.viewport().width()
        if not view_width:
            view_width = option.rect.width()
        if not view_width:
            view_width = 480

        card_width = max(view_width - 20, 220)
        has_image = entry.format == "image" and entry.image_data
        if has_image:
            available_width = max(card_width - 64 - 120, 140)
        else:
            available_width = max(card_width - 88, 160)

        body_font = QtGui.QFont(option.font)
        body_font.setPointSize(option.font.pointSize() + 1)
        snippet_text = self._preview_text(entry)
        text_height = self._text_height(snippet_text, body_font, available_width)

        min_content_height = 68 if has_image else 24
        content_height = max(int(math.ceil(text_height)), min_content_height)
        total_height = content_height + 84
        return QtCore.QSize(0, total_height)

    def editorEvent(
        self,
        event: QtCore.QEvent,
        model: QtCore.QAbstractItemModel,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> bool:
        entry = index.data(QtCore.Qt.UserRole)
        if not isinstance(entry, ClipboardItem):
            return super().editorEvent(event, model, option, index)

        if event.type() == QtCore.QEvent.MouseButtonRelease and event.buttons() == QtCore.Qt.NoButton:
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            rect = option.rect.adjusted(8, 4, -8, -4)
            if self._star_rect(rect).contains(pos):
                self.actionTriggered.emit("toggle_pin", entry)
                return True
            if self._delete_rect(rect).contains(pos):
                self.actionTriggered.emit("delete", entry)
                return True
        return super().editorEvent(event, model, option, index)

    def _star_rect(self, rect: QtCore.QRect) -> QtCore.QRect:
        return QtCore.QRect(rect.right() - 34, rect.top() + 10, 22, 22)

    def _delete_rect(self, rect: QtCore.QRect) -> QtCore.QRect:
        return QtCore.QRect(rect.right() - 34, rect.bottom() - 34, 22, 22)

    def _draw_star_icon(self, painter: QtGui.QPainter, rect: QtCore.QRect, pinned: bool) -> None:
        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        fill = QtGui.QColor(255, 215, 0) if pinned else QtGui.QColor(200, 204, 220)
        outline = QtGui.QColor(255, 235, 120) if pinned else QtGui.QColor(140, 146, 170)
        center = QtCore.QPointF(rect.center())
        outer = max(4.0, min(rect.width(), rect.height()) / 2.0 - 1.0)
        inner = outer * 0.45
        points: List[QtCore.QPointF] = []
        for index in range(10):
            angle = (math.pi / 5.0) * index - math.pi / 2.0
            radius = outer if index % 2 == 0 else inner
            points.append(
                QtCore.QPointF(
                    center.x() + radius * math.cos(angle),
                    center.y() + radius * math.sin(angle),
                )
            )
        polygon = QtGui.QPolygonF(points)
        painter.setBrush(fill)
        painter.setPen(QtGui.QPen(outline, 1))
        painter.drawPolygon(polygon)
        painter.restore()

    def _draw_delete_icon(self, painter: QtGui.QPainter, rect: QtCore.QRect) -> None:
        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        dark_mode = self._settings.theme_mode == "dark"
        color = QtGui.QColor(214, 102, 102) if dark_mode else QtGui.QColor(205, 72, 72)
        painter.setPen(QtGui.QPen(color, 2))
        inset = 4
        painter.drawLine(
            rect.left() + inset,
            rect.top() + inset,
            rect.right() - inset,
            rect.bottom() - inset,
        )
        painter.drawLine(
            rect.left() + inset,
            rect.bottom() - inset,
            rect.right() - inset,
            rect.top() + inset,
        )
        painter.restore()

    def _draw_format_icon(self, painter: QtGui.QPainter, rect: QtCore.QRect, entry: ClipboardItem, text_color: QtGui.QColor) -> None:
        painter.save()
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        base_color = QtGui.QColor(self._settings.accent_start)
        if not base_color.isValid():
            base_color = QtGui.QColor("#7f5af0")
        highlight = QtGui.QColor(text_color)
        highlight.setAlpha(90 if self._settings.theme_mode == "dark" else 120)
        painter.setBrush(base_color)
        painter.setPen(QtGui.QPen(highlight, 1.2))
        painter.drawEllipse(rect)
        painter.setPen(QtGui.QPen(QtGui.QColor("#ffffff")))
        painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
        icons = {
            "image": "IMG",
            "html": "HTM",
            "rich": "RTF",
            "files": "FIL",
            "urls": "URL",
            "table": "TAB",
            "text": "TXT",
        }
        painter.drawText(rect, QtCore.Qt.AlignCenter, icons.get(entry.format, "TXT"))
        painter.restore()

    def _text_height(self, text: str, font: QtGui.QFont, width: int) -> float:
        if width <= 0:
            width = 160
        doc = QtGui.QTextDocument()
        doc.setDefaultFont(font)
        doc.setPlainText(text or "")
        doc.setTextWidth(float(width))
        height = doc.size().height()
        if height <= 0:
            metrics = QtGui.QFontMetrics(font)
            return float(metrics.lineSpacing())
        return height

    def _preview_text(self, entry: ClipboardItem) -> str:
        if entry.format == "image":
            return "Bildvorschau"
        if entry.format == "files":
            return "\n".join(Path(path).name for path in entry.files[:4])
        if entry.format == "urls":
            return "\n".join(entry.urls[:4])
        if entry.format == "table":
            snippet = ""
            if entry.csv_data:
                try:
                    raw_csv = base64.b64decode(entry.csv_data)
                    snippet = decode_bytes_to_text(raw_csv)
                except Exception:
                    snippet = ""
            if not snippet and entry.html:
                doc = QtGui.QTextDocument()
                doc.setHtml(entry.html)
                snippet = doc.toPlainText()
            if not snippet:
                snippet = entry.content
        elif entry.format in ("html", "rich") and entry.html:
            doc = QtGui.QTextDocument()
            doc.setHtml(entry.html)
            snippet = doc.toPlainText()
        else:
            snippet = entry.content
        snippet = snippet.replace("\r", " ").replace("\n", " ")
        snippet = re.sub(r"(\S{60})", r"\1" + "\u200b", snippet)
        if len(snippet) > 340:
            snippet = snippet[:340] + "..."
        return snippet or "<leer>"

    def _format_label(self, entry: ClipboardItem) -> str:
        labels = {
            "image": "Bild",
            "html": "Rich Text / HTML",
            "rich": "Rich Text",
            "files": "Dateien",
            "urls": "Links",
            "table": "Tabelle",
            "text": "Text",
        }
        return labels.get(entry.format, "Text")

    def _get_pixmap(self, data: str) -> Optional[QtGui.QPixmap]:
        if not data:
            return None
        pixmap = self._pixmap_cache.get(data)
        if pixmap is None:
            try:
                pixmap = QtGui.QPixmap()
                pixmap.loadFromData(base64.b64decode(data), "PNG")
                self._pixmap_cache[data] = pixmap
            except Exception:
                return None
        return pixmap


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
        self._delegate.actionTriggered.connect(self._handle_delegate_action)
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
        self._list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self._list.customContextMenuRequested.connect(self._show_context_menu)

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
        dark_mode = self._settings.theme_mode == "dark"
        header_start = color_to_rgba(start, 0.75 if dark_mode else 0.45)
        header_end = color_to_rgba(end, 0.65 if dark_mode else 0.35)
        self._header.setStyleSheet(
            f"""
            QFrame#historyHeader {{
                border-radius: 20px;
                padding: 1px;
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 {header_start},
                    stop:1 {header_end}
                );
            }}
            """
        )
        inner_bg = "rgba(18, 19, 28, 220)" if dark_mode else "rgba(247, 248, 255, 235)"
        self._header_inner.setStyleSheet(
            f"""
            QFrame#historyHeaderInner {{
                background-color: {inner_bg};
                border-radius: 18px;
            }}
            """
        )
        list_bg = "rgba(18, 19, 28, 210)" if dark_mode else "rgba(255, 255, 255, 235)"
        list_border = color_to_rgba(start, 0.25 if dark_mode else 0.3)
        self._list_container.setStyleSheet(
            f"""
            QFrame#historyListContainer {{
                background-color: {list_bg};
                border-radius: 18px;
                border: 1px solid {list_border};
            }}
            """
        )
        title_color = "#f5f7ff" if dark_mode else "#1f2338"
        stats_color = (
            "rgba(245, 247, 255, 160)" if dark_mode else "rgba(70, 75, 95, 200)"
        )
        empty_color = (
            "rgba(154, 163, 192, 180)" if dark_mode else "rgba(110, 118, 140, 200)"
        )
        self._title_label.setStyleSheet(
            f"""
            QLabel {{
                font-size: 22px;
                font-weight: 700;
                letter-spacing: 0.5px;
                color: {title_color};
                background-color: transparent;
            }}
            """
        )
        self._stats_label.setStyleSheet(
            f"""
            QLabel {{
                font-size: 12px;
                color: {stats_color};
                background-color: transparent;
            }}
            """
        )
        self._empty_state.setStyleSheet(
            f"""
            QLabel {{
                color: {empty_color};
                font-size: 13px;
                padding: 40px 0;
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
            visible = [item for item in self._items_cache if self._matches_query(item, query)]
        else:
            visible = list(self._items_cache)
        self._populate_list(visible)
        self._update_stats(len(visible))

    def _populate_list(self, items: List[ClipboardItem]) -> None:
        self._list.clear()
        for entry in items:
            item = QtWidgets.QListWidgetItem()
            item.setData(QtCore.Qt.UserRole, entry)
            height = 152 if entry.format == "image" else 108
            item.setSizeHint(QtCore.QSize(0, height))
            self._list.addItem(item)
        has_items = bool(items)
        self._empty_state.setVisible(not has_items)
        self._list.setVisible(has_items)
        if has_items:
            self._list.setCurrentRow(0)

    def _update_stats(self, visible_count: int) -> None:
        total = len(self._items_cache)
        query = self._search_box.text().strip()
        pinned = len([item for item in self._items_cache if item.pinned])
        if query:
            self._stats_label.setText(f"{visible_count} Treffer | {total} Eintraege gesamt")
        else:
            self._stats_label.setText(
                f"{total} Eintraege (davon {pinned} angepinnt) | "
                f"{display_hotkey(self._settings.hotkey_prev)} / {display_hotkey(self._settings.hotkey_next)}"
            )

    def _matches_query(self, item: ClipboardItem, query: str) -> bool:
        candidates: List[str] = []
        if item.content:
            candidates.append(item.content.lower())
        if item.html:
            doc = QtGui.QTextDocument()
            doc.setHtml(item.html)
            candidates.append(doc.toPlainText().lower())
        for path_value in item.files:
            candidates.append(Path(path_value).name.lower())
        for url in item.urls:
            candidates.append(url.lower())
        if item.format == "image":
            candidates.append("bild")
        if item.pinned:
            candidates.append("pinned")
        return any(query in candidate for candidate in candidates if candidate)

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
            if entry.pinned:
                QtWidgets.QMessageBox.information(self, "Hinweis", "Eintrag ist angepinnt. Bitte zuerst loesen.")
            else:
                self._history.remove_entry(entry)

    def _show_context_menu(self, point: QtCore.QPoint) -> None:
        item = self._list.itemAt(point)
        if item is None:
            return
        entry = item.data(QtCore.Qt.UserRole)
        if not isinstance(entry, ClipboardItem):
            return
        self._list.setCurrentItem(item)
        menu = QtWidgets.QMenu(self)
        pin_label = "Anheften" if not entry.pinned else "Loesen"
        pin_action = menu.addAction(pin_label)
        delete_action = menu.addAction("Loeschen")
        delete_action.setEnabled(not entry.pinned)
        chosen = menu.exec(self._list.mapToGlobal(point))
        if chosen == pin_action:
            self._history.toggle_pin(entry)
            self._apply_current_filter()
        elif chosen == delete_action and not entry.pinned:
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

    def _handle_delegate_action(self, action: str, entry: ClipboardItem) -> None:
        if action == "toggle_pin":
            self._history.toggle_pin(entry)
        elif action == "delete":
            if entry.pinned:
                QtWidgets.QMessageBox.information(self, "Hinweis", "Eintrag ist angepinnt. Bitte zuerst loesen.")
            else:
                self._history.remove_entry(entry)
        self._apply_current_filter()


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

        self._theme_combo = QtWidgets.QComboBox()
        self._theme_combo.addItem("Dunkel", "dark")
        self._theme_combo.addItem("Hell", "light")
        theme_index = self._theme_combo.findData(self._settings.theme_mode)
        if theme_index >= 0:
            self._theme_combo.setCurrentIndex(theme_index)
        overlay_layout.addRow("Theme:", self._theme_combo)

        self._overlay_theme_combo = QtWidgets.QComboBox()
        overlay_theme_options = [
            ("Klassisch", "classic"),
            ("Glas", "glass"),
            ("Minimal", "minimal"),
        ]
        for label, value in overlay_theme_options:
            self._overlay_theme_combo.addItem(label, value)
        overlay_theme_index = self._overlay_theme_combo.findData(self._settings.overlay_theme)
        if overlay_theme_index >= 0:
            self._overlay_theme_combo.setCurrentIndex(overlay_theme_index)
        overlay_layout.addRow("Overlay-Style:", self._overlay_theme_combo)

        opacity_container = QtWidgets.QWidget()
        opacity_layout = QtWidgets.QHBoxLayout(opacity_container)
        opacity_layout.setContentsMargins(0, 0, 0, 0)
        opacity_layout.setSpacing(6)
        self._opacity_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self._opacity_slider.setRange(20, 100)
        self._opacity_slider.setValue(self._settings.overlay_opacity)
        self._opacity_value_label = QtWidgets.QLabel(f"{self._settings.overlay_opacity}%")
        self._opacity_slider.valueChanged.connect(lambda v: self._opacity_value_label.setText(f"{v}%"))
        opacity_layout.addWidget(self._opacity_slider, 1)
        opacity_layout.addWidget(self._opacity_value_label)
        overlay_layout.addRow("Overlay Opazitaet:", opacity_container)

        self._follow_checkbox = QtWidgets.QCheckBox("Toast folgt dem Mauszeiger")
        self._follow_checkbox.setChecked(self._settings.overlay_follow_mouse)
        overlay_layout.addRow("Maus-Follow:", self._follow_checkbox)

        self._anchor_combo = QtWidgets.QComboBox()
        self._anchor_combo.addItem("Oben links", "top-left")
        self._anchor_combo.addItem("Oben rechts", "top-right")
        self._anchor_combo.addItem("Unten links", "bottom-left")
        self._anchor_combo.addItem("Unten rechts", "bottom-right")
        anchor_index = self._anchor_combo.findData(self._settings.overlay_anchor)
        if anchor_index >= 0:
            self._anchor_combo.setCurrentIndex(anchor_index)
        overlay_layout.addRow("Verankerung:", self._anchor_combo)
        self._anchor_combo.setEnabled(not self._follow_checkbox.isChecked())
        self._follow_checkbox.toggled.connect(lambda checked: self._anchor_combo.setEnabled(not checked))

        offset_container = QtWidgets.QWidget()
        offset_layout = QtWidgets.QHBoxLayout(offset_container)
        offset_layout.setContentsMargins(0, 0, 0, 0)
        offset_layout.setSpacing(6)
        self._offset_x_spin = QtWidgets.QSpinBox()
        self._offset_x_spin.setRange(-500, 500)
        self._offset_x_spin.setValue(self._settings.overlay_offset_x)
        self._offset_y_spin = QtWidgets.QSpinBox()
        self._offset_y_spin.setRange(-500, 500)
        self._offset_y_spin.setValue(self._settings.overlay_offset_y)
        offset_layout.addWidget(QtWidgets.QLabel("X:"))
        offset_layout.addWidget(self._offset_x_spin)
        offset_layout.addWidget(QtWidgets.QLabel("Y:"))
        offset_layout.addWidget(self._offset_y_spin)
        overlay_layout.addRow("Offset:", offset_container)

        animation_container = QtWidgets.QWidget()
        animation_layout = QtWidgets.QHBoxLayout(animation_container)
        animation_layout.setContentsMargins(0, 0, 0, 0)
        animation_layout.setSpacing(6)
        self._anim_in_spin = QtWidgets.QSpinBox()
        self._anim_in_spin.setRange(50, 2000)
        self._anim_in_spin.setValue(self._settings.animation_in_ms)
        self._anim_out_spin = QtWidgets.QSpinBox()
        self._anim_out_spin.setRange(50, 2000)
        self._anim_out_spin.setValue(self._settings.animation_out_ms)
        animation_layout.addWidget(QtWidgets.QLabel("Ein:"))
        animation_layout.addWidget(self._anim_in_spin)
        animation_layout.addWidget(QtWidgets.QLabel("Aus:"))
        animation_layout.addWidget(self._anim_out_spin)
        overlay_layout.addRow("Animation (ms):", animation_container)

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

        retention_group = QtWidgets.QGroupBox("Verlauf & Aufbewahrung")
        retention_layout = QtWidgets.QFormLayout(retention_group)
        retention_layout.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

        self._auto_clear_checkbox = QtWidgets.QCheckBox("Verlauf automatisch bereinigen")
        self._auto_clear_checkbox.setChecked(self._settings.auto_clear_enabled)
        retention_layout.addRow("Auto-Clear:", self._auto_clear_checkbox)

        self._auto_clear_spin = QtWidgets.QSpinBox()
        self._auto_clear_spin.setRange(5, 10080)
        self._auto_clear_spin.setSuffix(" Min")
        self._auto_clear_spin.setValue(self._settings.auto_clear_interval_minutes)
        retention_layout.addRow("Intervall:", self._auto_clear_spin)

        self._auto_clear_checkbox.toggled.connect(self._auto_clear_spin.setEnabled)
        self._auto_clear_spin.setEnabled(self._auto_clear_checkbox.isChecked())

        layout.addWidget(retention_group)

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
        self._auto_clear_checkbox.setChecked(defaults.auto_clear_enabled)
        self._auto_clear_spin.setValue(defaults.auto_clear_interval_minutes)
        self._auto_clear_spin.setEnabled(self._auto_clear_checkbox.isChecked())
        theme_index = self._theme_combo.findData(defaults.theme_mode)
        if theme_index >= 0:
            self._theme_combo.setCurrentIndex(theme_index)
        overlay_theme_index = self._overlay_theme_combo.findData(defaults.overlay_theme)
        if overlay_theme_index >= 0:
            self._overlay_theme_combo.setCurrentIndex(overlay_theme_index)
        self._opacity_slider.setValue(defaults.overlay_opacity)
        self._opacity_value_label.setText(f"{defaults.overlay_opacity}%")
        self._follow_checkbox.setChecked(defaults.overlay_follow_mouse)
        anchor_index = self._anchor_combo.findData(defaults.overlay_anchor)
        if anchor_index >= 0:
            self._anchor_combo.setCurrentIndex(anchor_index)
        self._anchor_combo.setEnabled(not self._follow_checkbox.isChecked())
        self._offset_x_spin.setValue(defaults.overlay_offset_x)
        self._offset_y_spin.setValue(defaults.overlay_offset_y)
        self._anim_in_spin.setValue(defaults.animation_in_ms)
        self._anim_out_spin.setValue(defaults.animation_out_ms)

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
                theme_mode=self._theme_combo.currentData() or "dark",
                overlay_theme=self._overlay_theme_combo.currentData() or "classic",
                overlay_opacity=self._opacity_slider.value(),
                overlay_follow_mouse=self._follow_checkbox.isChecked(),
                overlay_anchor=self._anchor_combo.currentData(),
                overlay_offset_x=self._offset_x_spin.value(),
                overlay_offset_y=self._offset_y_spin.value(),
                animation_in_ms=self._anim_in_spin.value(),
                animation_out_ms=self._anim_out_spin.value(),
                auto_clear_enabled=self._auto_clear_checkbox.isChecked(),
                auto_clear_interval_minutes=self._auto_clear_spin.value(),
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

        self._auto_clear_timer = QtCore.QTimer(self)
        self._auto_clear_timer.setSingleShot(False)
        self._auto_clear_timer.timeout.connect(self._auto_clear_history)
        self._configure_auto_clear_timer()

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

    def _configure_auto_clear_timer(self) -> None:
        if self._settings.auto_clear_enabled:
            interval_ms = max(5, int(self._settings.auto_clear_interval_minutes)) * 60 * 1000
            self._auto_clear_timer.start(interval_ms)
        else:
            self._auto_clear_timer.stop()

    def _auto_clear_history(self) -> None:
        if not self._settings.auto_clear_enabled:
            return
        self._clipboard_history.clear()

    def register_main_window(self, window: "MainWindow") -> None:
        self._main_window = window
        self._main_window.apply_settings(self._settings)
        self._apply_capture_protection()
        self._configure_auto_clear_timer()

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
        self._configure_auto_clear_timer()
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

        self._hero_title = QtWidgets.QLabel(APP_DISPLAY_NAME)
        self._hero_title.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self._hero_title.setStyleSheet(
            """
            QLabel {
                font-size: 20px;
                font-weight: 700;
                background-color: transparent;
            }
            """
        )
        hero_layout.addWidget(self._hero_title)

        self._hero_caption = QtWidgets.QLabel(
            f"{APP_DISPLAY_NAME} begleitet dich immer im Hintergrund und bleibt ueber das Tray erreichbar."
        )
        self._hero_caption.setWordWrap(True)
        self._hero_caption.setStyleSheet(
            """
            QLabel {
                font-size: 12px;
                color: rgba(245, 247, 255, 150);
                background-color: transparent;
            }
            """
        )
        hero_layout.addWidget(self._hero_caption)

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
        dark_mode = settings.theme_mode == "dark"
        border = color_to_rgba(start, 0.4 if dark_mode else 0.22)
        hero_bg = "rgba(24, 26, 40, 235)" if dark_mode else "rgba(255, 255, 255, 240)"
        self._hero.setStyleSheet(
            f"""
            QFrame#welcomeCard {{
                background-color: {hero_bg};
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
        title_color = "#f5f7ff" if dark_mode else "#1f2238"
        caption_color = (
            "rgba(245, 247, 255, 150)" if dark_mode else "rgba(70, 75, 95, 190)"
        )
        feature_color = (
            "rgba(245, 247, 255, 190)" if dark_mode else "rgba(54, 59, 80, 200)"
        )
        self._hero_title.setStyleSheet(
            f"""
            QLabel {{
                font-size: 20px;
                font-weight: 700;
                color: {title_color};
                background-color: transparent;
            }}
            """
        )
        self._hero_caption.setStyleSheet(
            f"""
            QLabel {{
                font-size: 12px;
                color: {caption_color};
                background-color: transparent;
            }}
            """
        )
        self._feature_label.setStyleSheet(
            f"""
            QLabel {{
                font-size: 12px;
                color: {feature_color};
                background-color: transparent;
            }}
            """
        )
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

