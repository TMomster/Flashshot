# Copyright (c) 2026 Momster
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Lesser General Public License for more details.
# 
# You should have received a copy of the GNU Lesser General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import sys
import os
import json
import time
import subprocess
import winreg
import ctypes
import ctypes.wintypes
import threading
import traceback
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import (
    Qt, QObject, QCoreApplication, QEvent, QThreadPool, QRunnable, Signal, QTimer, QSharedMemory
)
from PySide6.QtGui import (
    QAction, QIcon, QKeySequence, QScreen, QPixmap
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QSystemTrayIcon, QMenu, QMessageBox,
    QWizard, QWizardPage, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QComboBox, QCheckBox, QTextEdit, QKeySequenceEdit,
    QSpinBox, QHBoxLayout, QGroupBox, QDialog, QPlainTextEdit, QWidget
)

from PySide6.QtCore import QUrl
from PySide6.QtMultimedia import QSoundEffect

# ================== 版本信息 ==================
SOFTWARE_VERSION = "2.4.1"

# ================== 日志系统 ==================
_log_buffer = []
_log_lock = threading.Lock()
_log_start_time = None

def init_logging():
    global _log_start_time
    _log_start_time = datetime.now()
    log_message("INFO", f"Flashshot v{SOFTWARE_VERSION} 启动")
    log_message("INFO", f"启动时间: {_log_start_time.strftime('%Y-%m-%d %H:%M:%S')}")

def log_message(level: str, msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    with _log_lock:
        _log_buffer.append(f"[{timestamp}] [{level}] {msg}")

def flush_log_to_file(exception_info: str = None, manual: bool = False):
    logs_dir = os.path.join(get_app_data_dir(), "logs")
    ensure_dir(logs_dir)
    if manual:
        filename = f"manual_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    else:
        filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    filepath = os.path.join(logs_dir, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"=== Flashshot 日志 ===\n")
        f.write(f"程序版本: {SOFTWARE_VERSION}\n")
        f.write(f"启动时间: {_log_start_time.strftime('%Y-%m-%d %H:%M:%S') if _log_start_time else '未知'}\n")
        f.write(f"写入时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        if exception_info:
            f.write(f"\n=== 异常信息 ===\n{exception_info}\n")
        f.write("\n=== 运行日志 ===\n")
        with _log_lock:
            f.write("\n".join(_log_buffer))
        f.write(f"\n\n=== 日志写入完成 ===\n")
    return filepath

def global_exception_handler(exc_type, exc_value, exc_tb):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)
    exception_msg = "".join(tb_lines)
    log_message("CRITICAL", f"未捕获异常:\n{exception_msg}")
    log_file = flush_log_to_file(exception_msg)
    try:
        app = QApplication.instance()
        if not app:
            app = QApplication(sys.argv)
        dialog = QDialog()
        dialog.setWindowTitle("Flashshot 意外崩溃")
        dialog.setMinimumSize(700, 500)
        layout = QVBoxLayout(dialog)
        info_label = QLabel(f"程序发生未处理的异常，已保存日志文件：\n{log_file}\n\n您可以复制日志内容以便分析问题。")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        text_edit = QPlainTextEdit()
        text_edit.setReadOnly(True)
        with _log_lock:
            full_log = "\n".join(_log_buffer)
        text_edit.setPlainText(f"=== 异常信息 ===\n{exception_msg}\n\n=== 运行日志 ===\n{full_log}")
        layout.addWidget(text_edit)
        btn_layout = QHBoxLayout()
        copy_btn = QPushButton("复制日志")
        open_dir_btn = QPushButton("打开日志目录")
        close_btn = QPushButton("关闭")
        btn_layout.addWidget(copy_btn)
        btn_layout.addWidget(open_dir_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(text_edit.toPlainText()))
        open_dir_btn.clicked.connect(lambda: os.startfile(os.path.dirname(log_file)))
        close_btn.clicked.connect(dialog.accept)
        dialog.exec()
    except Exception as e:
        try:
            QMessageBox.critical(None, "Flashshot 崩溃", f"程序崩溃，日志已保存到：{log_file}")
        except:
            print(f"程序崩溃，日志已保存到：{log_file}")
    sys.exit(1)

sys.excepthook = global_exception_handler
threading.excepthook = lambda args: global_exception_handler(args.exc_type, args.exc_value, args.exc_tb)

# ================== 全局通知管理器（左侧图标 + 右侧文字垂直居中）==================
class NotificationManager:
    _instance = None
    _lock = threading.Lock()
    
    # ================== 资源文件配置 ==================
    ICON_FILENAME = "resources\\Flashshot.png"      # 图标文件名
    SOUND_FILENAME = "resources\\Flashshot_noti.wav" # 音效文件名（WAV格式）
    ICON_SIZE = 32                        # 图标显示尺寸
    NOTIFY_WIDTH = 340                    # 通知窗口宽度
    NOTIFY_HEIGHT = 70                    # 通知窗口高度
    SOUND_VOLUME = 0.8                    # 音量（0.0 - 1.0）
    # ================================================
    
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
            return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        
        self.pending_count = 0
        self.pending_message = ""
        self.timer = None
        self.hwnd = None
        self.enabled = True
        self.duration = 1000
        self.enable_sound = True
        self.icon_hbitmap = None
        
        # 初始化音效播放器
        self.sound_effect = None
        self._init_sound_effect()
        
        self._create_overlay_window()
        self._load_icon()
    
    def configure(self, enabled: bool, duration_ms: int, enable_sound: bool):
        self.enabled = enabled
        self.duration = duration_ms
        self.enable_sound = enable_sound
    
    def _get_resource_paths(self, filename):
        """获取资源文件的搜索路径列表"""
        paths = []
        # 1. exe 所在目录（打包后）
        if getattr(sys, 'frozen', False):
            paths.append(os.path.join(os.path.dirname(sys.executable), filename))
        else:
            # 开发模式：脚本所在目录
            paths.append(os.path.join(os.path.dirname(__file__), filename))
        
        # 2. 当前工作目录
        paths.append(os.path.join(os.getcwd(), filename))
        
        # 3. 用户数据目录
        paths.append(os.path.join(get_app_data_dir(), filename))
        
        return paths
    
    def _init_sound_effect(self):
        """初始化音效播放器"""
        try:
            from PySide6.QtCore import QUrl
            from PySide6.QtMultimedia import QSoundEffect
            
            for path in self._get_resource_paths(self.SOUND_FILENAME):
                if os.path.exists(path):
                    self.sound_effect = QSoundEffect()
                    self.sound_effect.setSource(QUrl.fromLocalFile(path))
                    self.sound_effect.setVolume(self.SOUND_VOLUME)
                    log_message("INFO", f"通知音效加载成功: {path}")
                    return
            
            log_message("WARNING", f"未找到 {self.SOUND_FILENAME}，将使用系统提示音")
            self.sound_effect = None
        except Exception as e:
            log_message("WARNING", f"初始化音效播放器失败: {e}")
            self.sound_effect = None
    
    def _load_icon(self):
        """加载通知图标"""
        try:
            import win32gui
            import win32con
            
            for path in self._get_resource_paths(self.ICON_FILENAME):
                if os.path.exists(path):
                    self.icon_hbitmap = ctypes.windll.user32.LoadImageW(
                        None, path, 0x00, self.ICON_SIZE, self.ICON_SIZE, 0x0010
                    )
                    if self.icon_hbitmap:
                        log_message("INFO", f"通知图标加载成功: {path}")
                        return
            
            self._load_icon_via_qt()
            
        except Exception as e:
            log_message("WARNING", f"加载通知图标失败: {e}")
            self.icon_hbitmap = None

    def _load_icon_via_qt(self):
        """备用方案：通过 Qt 加载图片并转换为 HBITMAP"""
        try:
            from PySide6.QtGui import QPixmap
            
            for path in self._get_resource_paths(self.ICON_FILENAME):
                if os.path.exists(path):
                    pixmap = QPixmap(path)
                    if not pixmap.isNull():
                        pixmap = pixmap.scaled(self.ICON_SIZE, self.ICON_SIZE, 
                                               Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        hbitmap = pixmap.toImage().toHBITMAP()
                        self.icon_hbitmap = int(hbitmap)
                        log_message("INFO", f"通知图标通过 Qt 加载成功: {path}")
                        return
        except Exception as e:
            log_message("WARNING", f"Qt 加载图标失败: {e}")
    
    def _create_overlay_window(self):
        try:
            import win32gui
            import win32con
            
            wc = win32gui.WNDCLASS()
            wc.lpszClassName = "FlashshotOverlay"
            wc.hbrBackground = win32con.COLOR_BTNFACE + 1
            wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
            win32gui.RegisterClass(wc)
            
            self.hwnd = win32gui.CreateWindowEx(
                win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED,
                "FlashshotOverlay",
                "",
                win32con.WS_POPUP,
                0, 0, self.NOTIFY_WIDTH, self.NOTIFY_HEIGHT,
                None, None, None, None
            )
            win32gui.SetLayeredWindowAttributes(self.hwnd, 0, 230, win32con.LWA_ALPHA)
            win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
            log_message("INFO", "通知窗口创建成功")
        except Exception as e:
            log_message("WARNING", f"创建顶层窗口失败: {e}")
            self.hwnd = None
    
    def show(self, message: str):
        if not self.enabled:
            return
        
        if message.startswith("Flashshot "):
            message = message[10:]
        
        self.pending_count += 1
        self.pending_message = message
        
        if self.timer and self.timer.isActive():
            self.timer.stop()
        
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._flush)
        self.timer.start(300)
    
    def _flush(self):
        if self.pending_count == 0:
            return
        
        if self.pending_count == 1:
            final_msg = self.pending_message
        else:
            final_msg = f"已保存 {self.pending_count} 张截图"
        
        self.pending_count = 0
        
        if self.enable_sound:
            self._play_sound()
        
        self._show_window(final_msg)
    
    def _play_sound(self):
        """播放通知音效"""
        if not self.enable_sound:
            return
        
        try:
            if self.sound_effect is not None:
                self.sound_effect.play()
                log_message("INFO", "播放自定义通知音效")
            else:
                ctypes.windll.user32.MessageBeep(0xFFFFFFFF)
                log_message("INFO", "播放系统提示音")
        except Exception as e:
            log_message("WARNING", f"播放音效失败: {e}")
            try:
                ctypes.windll.user32.MessageBeep(0xFFFFFFFF)
            except:
                pass
    
    def _show_window(self, message: str):
        """显示通知 - 直接使用 Qt 备用方案"""
        self._show_qt_fallback(message)
    
    def _show_win32_window(self, message: str):
        """Win32 通知已禁用"""
        self._show_qt_fallback(message)
    
    def _hide_window(self):
        if self.hwnd:
            try:
                import win32gui
                import win32con
                win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
            except:
                pass
    
    def _show_qt_fallback(self, message: str):
        try:
            from PySide6.QtWidgets import QLabel, QVBoxLayout, QHBoxLayout, QWidget
            from PySide6.QtCore import QTimer, Qt
            from PySide6.QtGui import QPixmap, QFont
            
            # 创建窗口
            self.fallback_window = QWidget()
            self.fallback_window.setWindowFlags(
                Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnTopHint
            )
            
            # 主布局
            main_layout = QHBoxLayout(self.fallback_window)
            main_layout.setContentsMargins(12, 10, 12, 10)
            main_layout.setSpacing(10)
            
            # 图标
            icon_label = QLabel()
            icon_label.setFixedSize(self.ICON_SIZE, self.ICON_SIZE)
            pixmap = None
            
            for path in self._get_resource_paths(self.ICON_FILENAME):
                if os.path.exists(path):
                    pixmap = QPixmap(path)
                    if not pixmap.isNull():
                        pixmap = pixmap.scaled(self.ICON_SIZE, self.ICON_SIZE, 
                                               Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        icon_label.setPixmap(pixmap)
                        break
            
            if pixmap is None or pixmap.isNull():
                icon_label.hide()
            
            # 文字区域
            text_layout = QVBoxLayout()
            text_layout.setContentsMargins(0, 0, 0, 0)
            text_layout.setSpacing(3)
            
            title_label = QLabel("Flashshot")
            content_label = QLabel(message)
            content_label.setWordWrap(True)
            
            text_layout.addWidget(title_label)
            text_layout.addWidget(content_label)
            
            main_layout.addWidget(icon_label)
            main_layout.addLayout(text_layout, 1)
            
            # 设置窗口样式
            self.fallback_window.setStyleSheet("""
                QWidget {
                    background-color: #1a1a1a;
                    border-radius: 12px;
                    border: 1px solid #3a3a3a;
                }
                QLabel {
                    background-color: transparent;
                    border: none;
                    outline: none;
                }
            """)
            
            title_label.setStyleSheet("font-weight: bold; font-size: 13px; color: white;")
            title_label.setFont(QFont("Microsoft YaHei", 13, QFont.Bold))
            
            content_label.setStyleSheet("font-size: 11px; color: #CCCCCC;")
            content_label.setFont(QFont("Microsoft YaHei", 11))
            
            # 自适应宽度
            content_label.adjustSize()
            text_width = max(180, min(380, content_label.width() + 80))
            self.fallback_window.setFixedSize(text_width + 60, self.NOTIFY_HEIGHT)
            
            # 定位到右下角
            screen = QApplication.primaryScreen().availableGeometry()
            x = screen.right() - self.fallback_window.width() - 20
            y = screen.bottom() - self.fallback_window.height() - 20
            self.fallback_window.move(x, y)
            
            self.fallback_window.show()
            QTimer.singleShot(self.duration, self.fallback_window.close)
            
        except Exception as e:
            log_message("ERROR", f"Qt 备用通知失败: {e}")
# ================== 单例运行管理 ==================
class SingleInstance:
    def __init__(self, app_name="Flashshot"):
        self.app_name = app_name
        self.shared_memory = None
    def try_acquire(self):
        self.shared_memory = QSharedMemory(self.app_name)
        if self.shared_memory.attach():
            return False
        if self.shared_memory.create(1):
            return True
        return False

# ================== Windows API 键盘钩子 ==================
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
WM_QUIT = 0x0012

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", ctypes.c_uint),
        ("scanCode", ctypes.c_uint),
        ("flags", ctypes.c_uint),
        ("time", ctypes.c_uint),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]

HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)
user32 = ctypes.windll.user32

user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]
user32.CallNextHookEx.restype = ctypes.c_int

SetWindowsHookExW = user32.SetWindowsHookExW
CallNextHookEx = user32.CallNextHookEx
UnhookWindowsHookEx = user32.UnhookWindowsHookEx
GetMessage = user32.GetMessageW
TranslateMessage = user32.TranslateMessage
DispatchMessage = user32.DispatchMessageW
PostQuitMessage = user32.PostQuitMessage
PeekMessage = user32.PeekMessageW

class HotkeyEvent(QEvent):
    EVENT_TYPE = QEvent.Type(QEvent.registerEventType())
    def __init__(self, callback_name: str):
        super().__init__(HotkeyEvent.EVENT_TYPE)
        self.callback_name = callback_name

class LowLevelKeyboardHook(QObject):
    screenshot_triggered = Signal(str)
    def __init__(self):
        super().__init__()
        self.hook_id = None
        self.hook_proc = None
        self.modifier_state = {'ctrl': False, 'alt': False, 'shift': False, 'win': False}
        self.hotkeys = {}
        self.running = False
        self.hotkey_lock = threading.Lock()
        self._thread = None
        self._stop_event = threading.Event()

    def add_hotkey(self, hotkey_str: str, callback_name: str) -> bool:
        result = self._parse_hotkey(hotkey_str)
        if result:
            vk, mods = result
            with self.hotkey_lock:
                self.hotkeys[vk] = (mods, callback_name)
            log_message("INFO", f"注册热键: {hotkey_str} -> {callback_name}")
            return True
        log_message("ERROR", f"热键解析失败: {hotkey_str}")
        return False

    def remove_hotkey(self, callback_name: str):
        with self.hotkey_lock:
            to_remove = [vk for vk, (_, cb) in self.hotkeys.items() if cb == callback_name]
            for vk in to_remove:
                del self.hotkeys[vk]
                log_message("INFO", f"移除热键: {callback_name}")

    def update_hotkey(self, hotkey_str: str, callback_name: str) -> bool:
        self.remove_hotkey(callback_name)
        return self.add_hotkey(hotkey_str, callback_name)

    def _parse_hotkey(self, hotkey_str):
        if not hotkey_str:
            return None
        
        parts = hotkey_str.lower().replace(" ", "").split('+')
        
        mod_map = {
            'ctrl': 'ctrl', 'control': 'ctrl',
            'alt': 'alt',
            'shift': 'shift',
            'win': 'win', 'windows': 'win'
        }
        
        required_mods = set()
        main_key = None
        
        for part in parts:
            if part in mod_map:
                required_mods.add(mod_map[part])
            else:
                main_key = part
        
        if not main_key:
            return None
        
        vk = self._key_name_to_vk(main_key)
        if vk is None:
            log_message("WARNING", f"无法识别的按键: {main_key}")
            return None
        
        return (vk, required_mods)

    def _key_name_to_vk(self, key_name):
        key_name = key_name.lower().strip()
        
        if key_name.startswith('f'):
            try:
                num = int(key_name[1:])
                if 1 <= num <= 24:
                    return 0x70 + num - 1
            except:
                pass
        
        if len(key_name) == 1 and 'a' <= key_name <= 'z':
            return ord(key_name.upper())
        
        if len(key_name) == 1 and key_name.isdigit():
            return ord(key_name)
        
        symbol_map = {
            '`': 0xC0, '-': 0xBD, '=': 0xBB, '[': 0xDB, ']': 0xDD,
            '\\': 0xDC, ';': 0xBA, "'": 0xDE, ',': 0xBC, '.': 0xBE,
            '/': 0xBF, '+': 0xBB, '_': 0xBD, ':': 0xBA, '"': 0xDE,
            '<': 0xBC, '>': 0xBE, '?': 0xBF, '~': 0xC0, '!': 0x31,
            '@': 0x32, '#': 0x33, '$': 0x34, '%': 0x35, '^': 0x36,
            '&': 0x37, '*': 0x38, '(': 0x39, ')': 0x30,
        }
        
        if key_name in symbol_map:
            return symbol_map[key_name]
        
        special_map = {
            'backspace': 0x08, 'tab': 0x09, 'enter': 0x0D, 'return': 0x0D,
            'space': 0x20, 'spacebar': 0x20, 'escape': 0x1B, 'esc': 0x1B,
            'home': 0x24, 'end': 0x23, 'pageup': 0x21, 'pgup': 0x21,
            'pagedown': 0x22, 'pgdn': 0x22, 'pgdown': 0x22, 'insert': 0x2D,
            'ins': 0x2D, 'delete': 0x2E, 'del': 0x2E, 'up': 0x26,
            'down': 0x28, 'left': 0x25, 'right': 0x27, 'printscreen': 0x2C,
            'prtsc': 0x2C, 'scrolllock': 0x91, 'pause': 0x13, 'capslock': 0x14,
        }
        
        return special_map.get(key_name, None)

    def _update_modifier_state(self, vkCode, key_down):
        if vkCode in (0xA2, 0xA3):
            self.modifier_state['ctrl'] = key_down
        elif vkCode in (0xA4, 0xA5):
            self.modifier_state['alt'] = key_down
        elif vkCode in (0xA0, 0xA1):
            self.modifier_state['shift'] = key_down
        elif vkCode in (0x5B, 0x5C):
            self.modifier_state['win'] = key_down

    def _check_hotkey(self, vkCode):
        with self.hotkey_lock:
            if vkCode not in self.hotkeys:
                return None
            required_mods, callback = self.hotkeys[vkCode]
        for mod in required_mods:
            if not self.modifier_state.get(mod, False):
                return None
        return callback

    def _hook_callback(self, nCode, wParam, lParam):
        try:
            if nCode >= 0:
                kb = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vkCode = kb.vkCode
                key_down = (wParam in (WM_KEYDOWN, WM_SYSKEYDOWN))
                self._update_modifier_state(vkCode, key_down)
                if key_down:
                    cb = self._check_hotkey(vkCode)
                    if cb:
                        QCoreApplication.postEvent(self, HotkeyEvent(cb))
        except Exception as e:
            log_message("ERROR", f"键盘钩子回调异常: {e}")
        if self.hook_id:
            return CallNextHookEx(self.hook_id, nCode, wParam, lParam)
        return 0

    def start_hook(self):
        if self.running:
            return
        self.running = True
        self._stop_event.clear()
        self.hook_proc = HOOKPROC(self._hook_callback)
        self.hook_id = SetWindowsHookExW(WH_KEYBOARD_LL, self.hook_proc, None, 0)
        if not self.hook_id:
            log_message("ERROR", "安装键盘钩子失败，请以管理员权限运行")
            self.running = False
            return
        log_message("INFO", "键盘钩子已安装")
        
        msg = ctypes.wintypes.MSG()
        # 使用 GetMessage 阻塞循环，通过 MsgWaitForMultipleObjects 支持停止信号
        while self.running and not self._stop_event.is_set():
            # 使用 PeekMessage 检查是否有消息，同时可以检查停止标志
            ret = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE = 1
            if ret:
                if msg.message == WM_QUIT:
                    log_message("INFO", "收到 WM_QUIT 消息")
                    break
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                # 没有消息时，短暂休眠让出 CPU，但仍能快速响应停止信号
                # 使用 1ms 休眠，热键响应延迟可接受
                time.sleep(0.001)
        
        # 清理钩子
        if self.hook_id:
            UnhookWindowsHookEx(self.hook_id)
            self.hook_id = None
        self.running = False
        log_message("INFO", "键盘钩子已停止")

    def stop_hook(self):
        """停止键盘钩子"""
        if not self.running:
            log_message("INFO", "键盘钩子未运行，无需停止")
            return
        log_message("INFO", "正在停止键盘钩子...")
        self._stop_event.set()
        self.running = False
        # 发送退出消息以打破消息循环
        PostQuitMessage(0)
        log_message("INFO", "键盘钩子停止请求已发送")

    def is_running(self):
        return self.running

    def customEvent(self, event):
        if event.type() == HotkeyEvent.EVENT_TYPE:
            self.screenshot_triggered.emit(event.callback_name)
        super().customEvent(event)

# ================== 异步保存任务 ==================
class SaveScreenshotTask(QRunnable):
    def __init__(self, pixmap: QPixmap, filepath: str, quality: int):
        super().__init__()
        self.pixmap = pixmap
        self.filepath = filepath
        self.quality = quality
    def run(self):
        try:
            self.pixmap.save(self.filepath, "JPEG", quality=self.quality)
            log_message("INFO", f"截图已保存: {self.filepath}")
        except Exception as e:
            log_message("ERROR", f"保存截图失败: {e}")

class SaveSequenceTask(QRunnable):
    def __init__(self, frames, dir_path, quality):
        super().__init__()
        self.frames = frames
        self.dir_path = dir_path
        self.quality = quality
    def run(self):
        try:
            os.makedirs(self.dir_path, exist_ok=True)
            for i, (ts, pixmap) in enumerate(self.frames, start=1):
                filename = f"Backshot_{i}.jpg"
                filepath = os.path.join(self.dir_path, filename)
                pixmap.save(filepath, "JPEG", quality=self.quality)
            log_message("INFO", f"回放已保存: {len(self.frames)} 帧 -> {self.dir_path}")
        except Exception as e:
            log_message("ERROR", f"保存回放失败: {e}")

# ================== 回放缓冲区 ==================
class ReplayBuffer(QObject):
    def __init__(self, max_duration: int = 10, interval_ms: int = 500, scale_percent: int = 50):
        super().__init__()
        self.max_duration = max_duration
        self.interval_ms = interval_ms
        self.scale_percent = scale_percent
        self.buffer = []
        self.timer = None
        self.running = False
        self.lock = threading.Lock()
        self.config_lock = threading.Lock()

    def update_config(self, max_duration: int, interval_ms: int, scale_percent: int):
        with self.config_lock:
            self.max_duration = max_duration
            self.interval_ms = interval_ms
            self.scale_percent = scale_percent
            
            if self.running and self.timer:
                self.timer.setInterval(interval_ms)
            log_message("INFO", f"回放配置已更新: 时长={max_duration}s, 间隔={interval_ms}ms, 采样={scale_percent}%")

    def start(self):
        if self.running:
            return
        self.running = True
        self.timer = QTimer()
        self.timer.timeout.connect(self._capture_frame)
        self.timer.start(self.interval_ms)
        log_message("INFO", f"回放缓冲区启动")

    def stop(self):
        self.running = False
        if self.timer:
            self.timer.stop()
            self.timer = None
        with self.lock:
            self.buffer.clear()
        log_message("INFO", "回放缓冲区停止")

    def _capture_frame(self):
        try:
            screen = QApplication.primaryScreen()
            if not screen:
                return
            pixmap = screen.grabWindow(0)
            with self.config_lock:
                scale = self.scale_percent
            if scale < 100:
                factor = scale / 100.0
                new_size = pixmap.size() * factor
                pixmap = pixmap.scaled(new_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            ts = time.time()
            with self.lock:
                self.buffer.append((ts, pixmap))
                with self.config_lock:
                    cutoff = ts - self.max_duration
                while self.buffer and self.buffer[0][0] < cutoff:
                    self.buffer.pop(0)
        except Exception as e:
            log_message("ERROR", f"回放抓帧异常: {e}")

    def get_all_frames(self):
        with self.lock:
            return list(self.buffer)

# ================== 工具函数 ==================
def get_user_dir():
    return str(Path.home())

def get_app_data_dir():
    return os.path.join(get_user_dir(), "MomsterTech", "Flashshot")

def get_config_path():
    return os.path.join(get_app_data_dir(), "config.json")

def ensure_dir(path):
    Path(path).mkdir(parents=True, exist_ok=True)

# ================== 开机自启动（新：任务计划 + 启动文件夹）==================
def set_autostart(enabled: bool) -> bool:
    """
    设置开机自启动（任务计划程序 + 启动文件夹双重保障）
    返回 True 表示操作成功
    """
    task_name = "Flashshot"
    username = os.environ.get('USERNAME', '')
    
    # 获取程序真实启动路径
    if getattr(sys, 'frozen', False):
        # 打包后的 exe
        exe_path = sys.executable
        working_dir = os.path.dirname(exe_path)
        arguments = ""
    else:
        # 开发环境：使用 Python 解释器运行脚本
        exe_path = sys.executable
        script_path = os.path.abspath(__file__)
        working_dir = os.path.dirname(script_path)
        arguments = f'"{script_path}"'
    
    success = False
    
    # 1. 任务计划程序方式（首选）
    if enabled:
        # 创建任务 XML
        task_xml = f'''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>{username}</Author>
    <Description>Flashshot 开机自启动任务</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{username}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{username}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{exe_path}</Command>
      <Arguments>{arguments}</Arguments>
      <WorkingDirectory>{working_dir}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>'''
        
        temp_xml = os.path.join(os.environ['TEMP'], f'{task_name}_autostart.xml')
        try:
            with open(temp_xml, 'w', encoding='utf-16') as f:
                f.write(task_xml)
            
            # 创建任务（如果存在则强制覆盖）
            cmd = f'schtasks /create /tn "{task_name}" /xml "{temp_xml}" /f'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                log_message("INFO", f"任务计划程序自启动已创建: {task_name}")
                success = True
            else:
                log_message("ERROR", f"任务计划创建失败: {result.stderr}")
        except Exception as e:
            log_message("ERROR", f"创建任务计划异常: {e}")
        finally:
            if os.path.exists(temp_xml):
                os.remove(temp_xml)
    else:
        # 删除任务
        cmd = f'schtasks /delete /tn "{task_name}" /f'
        subprocess.run(cmd, shell=True, capture_output=True)
        log_message("INFO", f"已删除任务计划自启动: {task_name}")
        success = True
    
    # 2. 启动文件夹方式（作为补充，确保万一任务计划失效时仍能启动）
    startup_folder = os.path.join(os.environ['APPDATA'], 
                                  r'Microsoft\Windows\Start Menu\Programs\Startup')
    shortcut_path = os.path.join(startup_folder, "Flashshot.lnk")
    
    if enabled:
        # 创建快捷方式（需要 PowerShell）
        if getattr(sys, 'frozen', False):
            target = exe_path
            args = ""
        else:
            target = sys.executable
            args = f'"{__file__}"'
        
        # 转义路径中的反斜杠和引号
        target_escaped = target.replace('\\', '\\\\').replace('"', '\\"')
        args_escaped = args.replace('\\', '\\\\').replace('"', '\\"')
        working_dir_escaped = working_dir.replace('\\', '\\\\').replace('"', '\\"')
        
        ps_script = f'''
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut("{shortcut_path}")
        $Shortcut.TargetPath = "{target_escaped}"
        $Shortcut.Arguments = "{args_escaped}"
        $Shortcut.WorkingDirectory = "{working_dir_escaped}"
        $Shortcut.Save()
        '''
        try:
            subprocess.run(['powershell', '-Command', ps_script], capture_output=True, check=False)
            log_message("INFO", f"启动文件夹快捷方式已创建: {shortcut_path}")
            success = True
        except Exception as e:
            log_message("WARNING", f"创建启动文件夹快捷方式失败: {e}")
    else:
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            log_message("INFO", "已删除启动文件夹快捷方式")
    
    return success

def is_autostart_enabled() -> bool:
    """检查自启动是否已配置"""
    task_name = "Flashshot"
    result = subprocess.run(f'schtasks /query /tn "{task_name}"', shell=True, capture_output=True)
    if result.returncode == 0:
        return True
    startup_lnk = os.path.join(os.environ['APPDATA'], 
                               r'Microsoft\Windows\Start Menu\Programs\Startup\Flashshot.lnk')
    return os.path.exists(startup_lnk)

def is_admin():
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevate():
    """请求管理员权限并重启"""
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit(0)

def create_desktop_shortcut():
    desktop = os.path.join(get_user_dir(), "Desktop")
    shortcut_path = os.path.join(desktop, "Flashshot.lnk")
    target = sys.executable if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{__file__}"'
    ps = f'''
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{target}"
    $Shortcut.Save()
    '''
    subprocess.run(["powershell", "-Command", ps], capture_output=True)
    log_message("INFO", "已创建桌面快捷方式")

def remove_desktop_shortcut():
    desktop = os.path.join(get_user_dir(), "Desktop")
    shortcut_path = os.path.join(desktop, "Flashshot.lnk")
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        log_message("INFO", "已删除桌面快捷方式")

def load_config():
    path = get_config_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
            if "notification_duration_ms" not in cfg:
                cfg["notification_duration_ms"] = 1000
            if "notifications_enabled" not in cfg:
                cfg["notifications_enabled"] = True
            if "enable_sound" not in cfg:
                cfg["enable_sound"] = True
            return cfg
    except Exception as e:
        log_message("ERROR", f"加载配置失败: {e}")
        return None

def save_config(config):
    ensure_dir(get_app_data_dir())
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)
    log_message("INFO", "配置已保存")

def is_config_outdated():
    cfg = load_config()
    if not cfg:
        return True
    return cfg.get("version") != SOFTWARE_VERSION

# ================== 设置向导 ==================
class SingleKeySequenceEdit(QKeySequenceEdit):
    FORBIDDEN_KEYS = {
        Qt.Key_Exclam, Qt.Key_QuoteDbl, Qt.Key_NumberSign, Qt.Key_Dollar,
        Qt.Key_Percent, Qt.Key_Ampersand, Qt.Key_Apostrophe, Qt.Key_ParenLeft,
        Qt.Key_ParenRight, Qt.Key_Asterisk, Qt.Key_Plus, Qt.Key_Comma,
        Qt.Key_Minus, Qt.Key_Period, Qt.Key_Slash, Qt.Key_Colon,
        Qt.Key_Semicolon, Qt.Key_Less, Qt.Key_Equal, Qt.Key_Greater,
        Qt.Key_Question, Qt.Key_At, Qt.Key_BracketLeft, Qt.Key_Backslash,
        Qt.Key_BracketRight, Qt.Key_AsciiCircum, Qt.Key_Underscore, Qt.Key_QuoteLeft,
        Qt.Key_BraceLeft, Qt.Key_Bar, Qt.Key_BraceRight, Qt.Key_AsciiTilde,
        Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right,
        # Qt.Key_Delete, Qt.Key_Insert, Qt.Key_Home, Qt.Key_End,
        # Qt.Key_PageUp, Qt.Key_PageDown, Qt.Key_Print, Qt.Key_ScrollLock, Qt.Key_Pause,
    }
    
    ALLOWED_KEY_TYPES = set([getattr(Qt, f'Key_{c}') for c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ']) | \
                        set([getattr(Qt, f'Key_{i}') for i in range(10)]) | \
                        {Qt.Key_F1, Qt.Key_F2, Qt.Key_F3, Qt.Key_F4, Qt.Key_F5, Qt.Key_F6,
                         Qt.Key_F7, Qt.Key_F8, Qt.Key_F9, Qt.Key_F10, Qt.Key_F11, Qt.Key_F12,
                         Qt.Key_F13, Qt.Key_F14, Qt.Key_F15, Qt.Key_F16, Qt.Key_F17, Qt.Key_F18,
                         Qt.Key_F19, Qt.Key_F20, Qt.Key_F21, Qt.Key_F22, Qt.Key_F23, Qt.Key_F24,
                         Qt.Key_Space, Qt.Key_Tab, Qt.Key_Return, Qt.Key_Enter,
                         Qt.Key_Backspace, Qt.Key_Escape, Qt.Key_CapsLock}
    
    MODIFIER_KEYS = {Qt.Key_Control, Qt.Key_Alt, Qt.Key_Shift, Qt.Key_Meta}
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._suppress_validation = True
        self.keySequenceChanged.connect(self._on_changed)
        self._clearing = False
        self._last_valid_sequence = None
        self._parent_notification = None
        self._suppress_validation = False
    
    def setParentNotification(self, notification_manager):
        self._parent_notification = notification_manager
    
    def _get_key_parts(self, key_combination):
        if hasattr(key_combination, 'key'):
            main_key = key_combination.key()
            modifiers = key_combination.keyboardModifiers()
        else:
            main_key = int(key_combination) & 0xFFFFFF
            modifiers = int(key_combination) & ~0xFFFFFF
        return main_key, modifiers
    
    def _on_changed(self, seq):
        if self._suppress_validation:
            return
        if seq.isEmpty():
            return
        
        if seq.count() > 1:
            self.blockSignals(True)
            self.setKeySequence(QKeySequence(seq[0]))
            self.blockSignals(False)
            seq = self.keySequence()
        
        if not self._validate_hotkey(seq):
            if self._last_valid_sequence:
                self.blockSignals(True)
                self.setKeySequence(self._last_valid_sequence)
                self.blockSignals(False)
            else:
                self.clear()
            return
        
        self._last_valid_sequence = self.keySequence()
    
    def _validate_hotkey(self, seq):
        if seq.isEmpty():
            return False
        
        key_combo = seq[0]
        main_key, modifiers = self._get_key_parts(key_combo)
        
        if main_key == 0 or main_key in self.MODIFIER_KEYS:
            self._show_error("快捷键必须包含一个主键（如字母、数字、F1-F12等）")
            return False
        
        if main_key in self.FORBIDDEN_KEYS:
            self._show_error("禁止使用符号键和方向键作为快捷键")
            return False
        
        return True
    
    def _show_error(self, message):
        log_message("WARNING", f"快捷键设置失败: {message}")
        if self._parent_notification:
            self._parent_notification.show(f"快捷键无效: {message}")
    
    def keyPressEvent(self, event):
        key = event.key()
        
        if key == Qt.Key_Backspace or key == Qt.Key_Delete:
            if not self.keySequence().isEmpty() and not self._clearing:
                self._clearing = True
                self.clear()
                self._clearing = False
                self._last_valid_sequence = None
                event.accept()
                return
        
        super().keyPressEvent(event)
        
        current_seq = self.keySequence()
        if current_seq.isEmpty():
            self._last_valid_sequence = None
        elif self._validate_hotkey(current_seq):
            self._last_valid_sequence = current_seq

class SetupWizard(QWizard):
    def __init__(self, existing_config=None):
        super().__init__()
        self.setWindowTitle(f"Flashshot v{SOFTWARE_VERSION} 设置向导")
        self.setWizardStyle(QWizard.ModernStyle)
        self.existing_config = existing_config
        
        self.page_hotkey = QWizardPage()
        self.page_dir = QWizardPage()
        self.page_quality = QWizardPage()
        self.page_replay = QWizardPage()
        self.page_notification = QWizardPage()
        self.page_other = QWizardPage()
        self.page_confirm = QWizardPage()
        
        self.init_hotkey_page()
        self.init_dir_page()
        self.init_quality_page()
        self.init_replay_page()
        self.init_notification_page()
        self.init_other_page()
        self.init_confirm_page()
        
        self.addPage(self.page_hotkey)
        self.addPage(self.page_dir)
        self.addPage(self.page_quality)
        self.addPage(self.page_replay)
        self.addPage(self.page_notification)
        self.addPage(self.page_other)
        self.confirm_page_id = self.addPage(self.page_confirm)
        
        self.prefill()
    
    def prefill(self):
        if not self.existing_config:
            return
        self.hotkey_edit.setKeySequence(QKeySequence(self.existing_config.get("hotkey", "f12")))
        self.replay_hotkey_edit.setKeySequence(QKeySequence(self.existing_config.get("replay_hotkey", "ctrl+shift+pageup")))
        d = self.existing_config.get("save_dir", "")
        if d:
            self.dir_display.setText(d)
        q = self.existing_config.get("quality", "high")
        idx = {"high": 0, "medium": 1, "low": 2}.get(q, 0)
        self.quality_combo.setCurrentIndex(idx)
        self.replay_enable_check.setChecked(self.existing_config.get("replay_enabled", False))
        self.replay_seconds.setValue(self.existing_config.get("replay_duration", 10))
        self.replay_interval.setValue(self.existing_config.get("replay_interval_ms", 500))
        scale = self.existing_config.get("replay_scale", 50)
        scale_idx = {100:0, 75:1, 50:2, 25:3}.get(scale, 2)
        self.replay_scale_combo.setCurrentIndex(scale_idx)
        self.autostart_check.setChecked(self.existing_config.get("autostart", False))
        self.desktop_check.setChecked(self.existing_config.get("desktop_shortcut", False))
        self.notification_enable_check.setChecked(self.existing_config.get("notifications_enabled", True))
        self.sound_check.setChecked(self.existing_config.get("enable_sound", True))
        dur_ms = self.existing_config.get("notification_duration_ms", 1000)
        dur_sec = dur_ms / 1000.0
        dur_idx = {0.5:0, 1.0:1, 1.5:2, 2.0:3}.get(dur_sec, 1)
        self.duration_combo.setCurrentIndex(dur_idx)
    
    def init_hotkey_page(self):
        self.page_hotkey.setTitle("快捷键设置")
        layout = QVBoxLayout()
        self.hotkey_edit = SingleKeySequenceEdit()
        self.hotkey_edit.setKeySequence(QKeySequence("F12"))
        layout.addWidget(QLabel("普通截图快捷键:"))
        layout.addWidget(self.hotkey_edit)
        
        tip_label = QLabel("提示：支持字母、数字、F1-F12等键，可配合Ctrl/Alt/Shift/Win组合")
        tip_label.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(tip_label)
        
        layout.addStretch()
        self.page_hotkey.setLayout(layout)

    def init_replay_page(self):
        self.page_replay.setTitle("回放截屏设置")
        layout = QVBoxLayout()
        self.replay_enable_check = QCheckBox("启用回放截屏功能")
        layout.addWidget(self.replay_enable_check)
        
        gb = QGroupBox("回放参数")
        gb_layout = QVBoxLayout()
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("回溯时长(秒):"))
        self.replay_seconds = QSpinBox()
        self.replay_seconds.setRange(1, 60)
        self.replay_seconds.setValue(10)
        h1.addWidget(self.replay_seconds)
        h1.addStretch()
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("抓帧间隔(毫秒):"))
        self.replay_interval = QSpinBox()
        self.replay_interval.setRange(100, 2000)
        self.replay_interval.setSingleStep(100)
        self.replay_interval.setValue(500)
        h2.addWidget(self.replay_interval)
        h2.addStretch()
        h3 = QHBoxLayout()
        h3.addWidget(QLabel("采样比例:"))
        self.replay_scale_combo = QComboBox()
        self.replay_scale_combo.addItems(["1.0x", "0.75x", "0.5x", "0.25x"])
        self.replay_scale_combo.setCurrentIndex(2)
        h3.addWidget(self.replay_scale_combo)
        h3.addStretch()
        gb_layout.addLayout(h1)
        gb_layout.addLayout(h2)
        gb_layout.addLayout(h3)
        gb.setLayout(gb_layout)
        layout.addWidget(gb)
        
        layout.addWidget(QLabel("回放截屏快捷键:"))
        self.replay_hotkey_edit = SingleKeySequenceEdit()
        self.replay_hotkey_edit.setKeySequence(QKeySequence("Ctrl+Shift+PageUp"))
        layout.addWidget(self.replay_hotkey_edit)
        
        tip_label = QLabel("提示：支持字母、数字、F1-F12等键，可配合Ctrl/Alt/Shift/Win组合")
        tip_label.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(tip_label)
        
        layout.addStretch()
        self.page_replay.setLayout(layout)
        
        self.replay_enable_check.toggled.connect(self.replay_seconds.setEnabled)
        self.replay_enable_check.toggled.connect(self.replay_interval.setEnabled)
        self.replay_enable_check.toggled.connect(self.replay_scale_combo.setEnabled)
        self.replay_enable_check.toggled.connect(self.replay_hotkey_edit.setEnabled)
    
    def init_dir_page(self):
        self.page_dir.setTitle("保存目录")
        layout = QVBoxLayout()
        self.dir_display = QLineEdit()
        self.dir_display.setReadOnly(True)
        default_dir = os.path.join(get_user_dir(), "Pictures", "Flashshot")
        self.dir_display.setText(default_dir)
        btn = QPushButton("浏览...")
        btn.clicked.connect(self._browse_dir)
        layout.addWidget(QLabel("截图保存目录:"))
        layout.addWidget(self.dir_display)
        layout.addWidget(btn)
        layout.addStretch()
        self.page_dir.setLayout(layout)
    
    def _browse_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择目录", self.dir_display.text())
        if d:
            self.dir_display.setText(d)
    
    def init_quality_page(self):
        self.page_quality.setTitle("图片质量")
        layout = QVBoxLayout()
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["高", "中", "低"])
        layout.addWidget(QLabel("截图质量:"))
        layout.addWidget(self.quality_combo)
        layout.addStretch()
        self.page_quality.setLayout(layout)
    
    def init_notification_page(self):
        self.page_notification.setTitle("通知设置")
        layout = QVBoxLayout()
        self.notification_enable_check = QCheckBox("启用通知（截图时右下角提示）")
        self.notification_enable_check.setChecked(True)
        layout.addWidget(self.notification_enable_check)
        
        h = QHBoxLayout()
        h.addWidget(QLabel("通知显示时长:"))
        self.duration_combo = QComboBox()
        self.duration_combo.addItems(["0.5 秒", "1.0 秒", "1.5 秒", "2.0 秒"])
        self.duration_combo.setCurrentIndex(1)
        h.addWidget(self.duration_combo)
        h.addStretch()
        layout.addLayout(h)
        
        self.sound_check = QCheckBox("截图时播放提示音")
        self.sound_check.setChecked(True)
        layout.addWidget(self.sound_check)
        layout.addStretch()
        self.page_notification.setLayout(layout)
    
    def init_other_page(self):
        self.page_other.setTitle("其他选项")
        layout = QVBoxLayout()
        self.autostart_check = QCheckBox("开机自启动")
        self.desktop_check = QCheckBox("创建桌面快捷方式")
        layout.addWidget(self.autostart_check)
        layout.addWidget(self.desktop_check)
        layout.addStretch()
        self.page_other.setLayout(layout)
    
    def init_confirm_page(self):
        self.page_confirm.setTitle("确认设置")
        layout = QVBoxLayout()
        self.confirm_text = QTextEdit()
        self.confirm_text.setReadOnly(True)
        layout.addWidget(self.confirm_text)
        layout.addStretch()
        self.page_confirm.setLayout(layout)
    
    def initializePage(self, id):
        if id == self.confirm_page_id:
            hotkey = self.hotkey_edit.keySequence().toString().lower()
            replay_hk = self.replay_hotkey_edit.keySequence().toString().lower()
            save_dir = self.dir_display.text().strip()
            quality = self.quality_combo.currentText()
            replay_en = self.replay_enable_check.isChecked()
            replay_dur = self.replay_seconds.value()
            replay_int = self.replay_interval.value()
            scale_text = self.replay_scale_combo.currentText()
            if "1.0x" in scale_text:
                replay_scale = 100
            elif "0.75x" in scale_text:
                replay_scale = 75
            elif "0.5x" in scale_text:
                replay_scale = 50
            else:
                replay_scale = 25
            autostart = self.autostart_check.isChecked()
            desktop = self.desktop_check.isChecked()
            notif_en = self.notification_enable_check.isChecked()
            sound_en = self.sound_check.isChecked()
            dur_text = self.duration_combo.currentText()
            if "0.5" in dur_text:
                notif_duration_ms = 500
            elif "1.0" in dur_text:
                notif_duration_ms = 1000
            elif "1.5" in dur_text:
                notif_duration_ms = 1500
            else:
                notif_duration_ms = 2000
            
            self.temp = {
                "hotkey": hotkey, "replay_hotkey": replay_hk, "save_dir": save_dir,
                "quality": quality, "replay_enabled": replay_en,
                "replay_duration": replay_dur, "replay_interval_ms": replay_int,
                "replay_scale": replay_scale, "autostart": autostart,
                "desktop_shortcut": desktop, "notifications_enabled": notif_en,
                "enable_sound": sound_en, "notification_duration_ms": notif_duration_ms
            }
            
            text = f"""
            <b>普通截图快捷键:</b> {hotkey}<br>
            <b>回放截屏快捷键:</b> {replay_hk if replay_en else "(未启用)"}<br>
            <b>保存目录:</b> {save_dir}<br>
            <b>图片质量:</b> {quality}<br>
            <b>回放功能:</b> {'启用' if replay_en else '禁用'}<br>
            <b>回放时长/间隔/采样:</b> {replay_dur}秒 / {replay_int}毫秒 / {replay_scale}%<br>
            <b>开机自启动:</b> {'是' if autostart else '否'}<br>
            <b>桌面快捷方式:</b> {'是' if desktop else '否'}<br>
            <b>通知:</b> {'启用' if notif_en else '禁用'}<br>
            <b>通知时长:</b> {notif_duration_ms/1000}秒<br>
            <b>提示音:</b> {'启用' if sound_en else '禁用'}<br>
            """
            self.confirm_text.setHtml(text)
    
    def accept(self):
        cfg = {
            "version": SOFTWARE_VERSION,
            "hotkey": self.temp["hotkey"],
            "replay_hotkey": self.temp["replay_hotkey"],
            "save_dir": self.temp["save_dir"],
            "quality": {"高": "high", "中": "medium", "低": "low"}[self.temp["quality"]],
            "replay_enabled": self.temp["replay_enabled"],
            "replay_duration": self.temp["replay_duration"],
            "replay_interval_ms": self.temp["replay_interval_ms"],
            "replay_scale": self.temp["replay_scale"],
            "autostart": self.temp["autostart"],
            "desktop_shortcut": self.temp["desktop_shortcut"],
            "notifications_enabled": self.temp["notifications_enabled"],
            "enable_sound": self.temp["enable_sound"],
            "notification_duration_ms": self.temp["notification_duration_ms"],
            "first_run_done": True
        }
        ensure_dir(cfg["save_dir"])
        save_config(cfg)
        set_autostart(cfg["autostart"])
        if cfg["desktop_shortcut"]:
            create_desktop_shortcut()
        else:
            remove_desktop_shortcut()
        
        super().accept()

# ================== 主应用程序 ==================
class FlashshotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flashshot")
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.resize(300, 200)
        
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon.fromTheme("camera-photo", QIcon()))
        self.setup_tray_menu()
        self.tray_icon.show()
        
        self.hook = LowLevelKeyboardHook()
        self.hook.screenshot_triggered.connect(self._on_hotkey)
        self.hook_thread = None
        self._hook_thread_stopped = False
        
        self.replay_buffer = None
        self.replay_enabled = False
        
        self.notification_mgr = NotificationManager()
        
        self.last_timestamp = None
        self.sequence_counter = 1
        
        self.config = load_config()
        need_wizard = (not self.config or not self.config.get("first_run_done") or is_config_outdated())
        
        if need_wizard:
            self.run_setup_wizard()
        else:
            self.apply_config()
        
        if self.config and not self.config.get("license_acknowledged", False):
            self.show_license_notice()
        
        if self.config and self.config.get("notifications_enabled", True):
            self.notification_mgr.show("Flashshot 已启动")

        QApplication.instance().aboutToQuit.connect(self.on_about_to_quit)


    
    def on_about_to_quit(self):
        """应用程序即将退出时的清理"""
        log_message("INFO", "应用程序即将退出，执行清理...")
        self.cleanup_and_exit()
    
    def cleanup_and_exit(self):
        """彻底清理资源并退出"""
        log_message("INFO", "开始清理资源...")
        
        # 1. 停止回放缓冲区
        if self.replay_buffer:
            self.replay_buffer.stop()
            self.replay_buffer = None
        
        # 2. 停止键盘钩子
        if self.hook:
            self.hook.stop_hook()
        
        # 3. 等待钩子线程结束（最多等待2秒）
        if self.hook_thread and self.hook_thread.is_alive():
            log_message("INFO", "等待键盘钩子线程结束...")
            self.hook_thread.join(timeout=2.0)
            if self.hook_thread.is_alive():
                log_message("WARNING", "键盘钩子线程未能在超时内结束，将强制退出")
        
        # 4. 隐藏托盘图标
        if self.tray_icon:
            self.tray_icon.hide()
        
        log_message("INFO", "资源清理完成")
    
    def setup_tray_menu(self):
        self.tray_menu = QMenu()
        self.replay_toggle_action = QAction("回放功能: 关闭", self)
        self.replay_toggle_action.triggered.connect(self.toggle_replay_from_tray)
        self.tray_menu.addAction(self.replay_toggle_action)
        self.tray_menu.addSeparator()
        
        open_dir_action = QAction("打开截图保存目录", self)
        open_dir_action.triggered.connect(self.open_screenshot_dir)
        self.tray_menu.addAction(open_dir_action)
        
        export_log_action = QAction("导出日志", self)
        export_log_action.triggered.connect(self.export_manual_log)
        self.tray_menu.addAction(export_log_action)
        
        open_log_action = QAction("打开日志目录", self)
        open_log_action.triggered.connect(self.open_log_dir)
        self.tray_menu.addAction(open_log_action)
        
        self.tray_menu.addSeparator()
        about_action = QAction("关于 Flashshot", self)
        about_action.triggered.connect(self.show_about_dialog)
        self.tray_menu.addAction(about_action)
        
        settings_action = QAction("重新设置向导", self)
        settings_action.triggered.connect(self.run_setup_wizard)
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_app)
        self.tray_menu.addAction(settings_action)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
    
    def open_screenshot_dir(self):
	    if self.config and "save_dir" in self.config:
	        dir_path = self.config["save_dir"]
	        # 转换为绝对路径，避免相对路径导致打开错误位置
	        abs_path = os.path.abspath(dir_path)
	        log_message("INFO", f"配置中的保存目录: {dir_path}")
	        log_message("INFO", f"绝对路径: {abs_path}")
	        
	        # 确保目录存在
	        ensure_dir(abs_path)
	        
	        # 如果路径实际指向文档目录，给出警告
	        docs_path = os.path.join(get_user_dir(), "Documents")
	        if abs_path == docs_path:
	            self.notification_mgr.show("警告：当前保存目录为「文档」目录，建议重新设置")
	        
	        try:
	            # 使用 explorer 直接打开，避免 ShellExecute 的安全拦截
	            subprocess.Popen(['explorer', abs_path], shell=False)
	            log_message("INFO", f"已打开截图目录: {abs_path}")
	        except Exception as e:
	            log_message("ERROR", f"打开截图目录失败: {e}")
	            self.notification_mgr.show(f"无法打开目录: {str(e)}")
	    else:
	        self.notification_mgr.show("未配置保存目录")
    
    def export_manual_log(self):
        try:
            filepath = flush_log_to_file(manual=True)
            clipboard = QApplication.clipboard()
            clipboard.setText(filepath)
            self.notification_mgr.show(f"日志已导出 (路径已复制)")
            log_message("INFO", f"手动导出日志: {filepath}")
        except Exception as e:
            self.notification_mgr.show(f"导出日志失败: {str(e)}")
            log_message("ERROR", f"导出日志失败: {e}")
    
    def open_log_dir(self):
        logs_dir = os.path.join(get_app_data_dir(), "logs")
        ensure_dir(logs_dir)
        try:
            os.startfile(logs_dir)
            log_message("INFO", f"已打开日志目录: {logs_dir}")
        except Exception as e:
            log_message("ERROR", f"打开日志目录失败: {e}")
            self.notification_mgr.show(f"无法打开日志目录: {str(e)}")
    
    def show_about_dialog(self):
        self.notification_mgr.show(f"Flashshot v{SOFTWARE_VERSION}")
        log_message("INFO", f"显示关于信息: v{SOFTWARE_VERSION}")
    
    def show_license_notice(self):
        self.notification_mgr.show("本软件使用了 PySide6 (LGPLv3 许可证)")
        log_message("INFO", "已显示开源许可证通知")
        if self.config:
            self.config["license_acknowledged"] = True
            save_config(self.config)
    
    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.do_screenshot()
    
    def toggle_replay_from_tray(self):
        if not self.config:
            return
        new_state = not self.replay_enabled
        self.config["replay_enabled"] = new_state
        save_config(self.config)
        self.apply_replay_state(new_state)
        self.notification_mgr.show(f"回放功能已{'启用' if new_state else '禁用'}")
    
    def apply_replay_state(self, enabled):
        self.replay_enabled = enabled
        if enabled:
            duration = self.config.get("replay_duration", 10)
            interval = self.config.get("replay_interval_ms", 500)
            scale = self.config.get("replay_scale", 50)
            if self.replay_buffer is None:
                self.replay_buffer = ReplayBuffer(duration, interval, scale)
            else:
                self.replay_buffer.update_config(duration, interval, scale)
            self.replay_buffer.start()
            replay_hk = self.config.get("replay_hotkey", "")
            if replay_hk:
                self.hook.update_hotkey(replay_hk, "replay")
        else:
            if self.replay_buffer:
                self.replay_buffer.stop()
                self.replay_buffer = None
            self.hook.remove_hotkey("replay")
        self.replay_toggle_action.setText(f"回放功能: {'开启' if enabled else '关闭'}")
    
    def run_setup_wizard(self):
        wizard = SetupWizard(self.config)
        if hasattr(wizard, 'hotkey_edit'):
            wizard.hotkey_edit.setParentNotification(self.notification_mgr)
        if hasattr(wizard, 'replay_hotkey_edit'):
            wizard.replay_hotkey_edit.setParentNotification(self.notification_mgr)
        if wizard.exec():
            self.config = load_config()
            self.apply_config_dynamic()
            if self.config.get("hotkey"):
                self.notification_mgr.show("配置已更新")
    
    def apply_config_dynamic(self):
        if not self.config:
            return
        
        self.notification_mgr.configure(
            enabled=self.config.get("notifications_enabled", True),
            duration_ms=self.config.get("notification_duration_ms", 1000),
            enable_sound=self.config.get("enable_sound", True)
        )
        
        hotkey = self.config.get("hotkey", "f12")
        self.hook.update_hotkey(hotkey, "screenshot")
        
        replay_en = self.config.get("replay_enabled", False)
        if replay_en:
            duration = self.config.get("replay_duration", 10)
            interval = self.config.get("replay_interval_ms", 500)
            scale = self.config.get("replay_scale", 50)
            if self.replay_buffer is None:
                self.replay_buffer = ReplayBuffer(duration, interval, scale)
                self.replay_buffer.start()
            else:
                self.replay_buffer.update_config(duration, interval, scale)
            replay_hk = self.config.get("replay_hotkey", "")
            if replay_hk:
                self.hook.update_hotkey(replay_hk, "replay")
        else:
            if self.replay_buffer:
                self.replay_buffer.stop()
                self.replay_buffer = None
            self.hook.remove_hotkey("replay")
        
        self.replay_enabled = replay_en
        self.replay_toggle_action.setText(f"回放功能: {'开启' if replay_en else '关闭'}")
        
        ensure_dir(self.config.get("save_dir", ""))
        set_autostart(self.config.get("autostart", False))
        if self.config.get("desktop_shortcut", False):
            create_desktop_shortcut()
        else:
            remove_desktop_shortcut()
        
        log_message("INFO", "配置动态应用完成")
    
    def apply_config(self):
        """首次启动时应用配置（启动键盘钩子线程）"""
        if not self.config:
            return
        
        self.notification_mgr.configure(
            enabled=self.config.get("notifications_enabled", True),
            duration_ms=self.config.get("notification_duration_ms", 1000),
            enable_sound=self.config.get("enable_sound", True)
        )
        
        hotkey = self.config.get("hotkey", "f12")
        
        # 注册热键
        self.hook.hotkeys.clear()
        if not self.hook.add_hotkey(hotkey, "screenshot"):
            log_message("ERROR", "普通截图热键注册失败")
        
        # 应用回放状态
        replay_en = self.config.get("replay_enabled", False)
        self.apply_replay_state(replay_en)
        
        # 启动键盘钩子线程（如果未启动）
        if self.hook_thread is None or not self.hook_thread.is_alive():
            self.hook_thread = threading.Thread(target=self.hook.start_hook, daemon=False)
            self.hook_thread.start()
            log_message("INFO", "键盘钩子线程已启动")
        
        # 其他配置
        ensure_dir(self.config.get("save_dir", ""))
        set_autostart(self.config.get("autostart", False))
        if self.config.get("desktop_shortcut", False):
            create_desktop_shortcut()
        else:
            remove_desktop_shortcut()
        
        log_message("INFO", "配置应用完成")

    def _run_hook_with_cleanup(self):
        """运行钩子并在结束时记录日志"""
        try:
            self.hook.start_hook()
        except Exception as e:
            log_message("ERROR", f"键盘钩子线程异常: {e}")
        finally:
            log_message("INFO", "键盘钩子线程已退出")
            self._hook_thread_stopped = True
    
    def _on_hotkey(self, callback):
        if callback == "screenshot":
            self.do_screenshot()
        elif callback == "replay":
            self.do_replay_capture()
    
    def _get_unique_filename(self, save_dir, base_name):
        ts = time.strftime("%Y%m%d_%H%M%S")
        if base_name == ts:
            if ts == self.last_timestamp:
                self.sequence_counter += 1
            else:
                self.last_timestamp = ts
                self.sequence_counter = 1
            filename = f"Flashshot_{ts}_{self.sequence_counter}.jpg"
        else:
            filename = base_name
        return os.path.join(save_dir, filename)
    
    def do_screenshot(self):
        if not self.config:
            return
        save_dir = self.config["save_dir"]
        ensure_dir(save_dir)
        screen = QApplication.primaryScreen()
        if not screen:
            log_message("ERROR", "无法获取主屏幕")
            return
        pixmap = screen.grabWindow(0)
        ts = time.strftime("%Y%m%d_%H%M%S")
        filepath = self._get_unique_filename(save_dir, ts)
        quality = {"high": 95, "medium": 75, "low": 50}.get(self.config["quality"], 95)
        task = SaveScreenshotTask(pixmap, filepath, quality)
        QThreadPool.globalInstance().start(task)
        self.notification_mgr.show("截图已保存")
        log_message("INFO", f"触发截图: {filepath}")
    
    def do_replay_capture(self):
        if not self.replay_enabled or not self.replay_buffer:
            self.notification_mgr.show("回放功能未启用")
            return
        frames = self.replay_buffer.get_all_frames()
        if not frames:
            self.notification_mgr.show("回放缓冲区为空")
            return
        now = datetime.now()
        base_dir = self.config["save_dir"]
        seq = 1
        while True:
            dir_name = f"Backshot_{now.strftime('%Y%m%d_%H%M%S')}_{seq}"
            dir_path = os.path.join(base_dir, dir_name)
            if not os.path.exists(dir_path):
                break
            seq += 1
        quality = {"high": 95, "medium": 75, "low": 50}.get(self.config["quality"], 95)
        task = SaveSequenceTask(frames, dir_path, quality)
        QThreadPool.globalInstance().start(task)
        self.notification_mgr.show(f"回放已保存 ({len(frames)} 帧)")
        log_message("INFO", f"触发回放保存: {dir_path}")
    
    def quit_app(self):
        """彻底退出程序"""
        log_message("INFO", "用户主动退出")
        
        # 执行清理
        self.cleanup_and_exit()
        
        # 退出应用程序
        QApplication.quit()
    
    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.notification_mgr.show("程序已最小化到托盘")

# ================== 主入口 ==================
def main():
    if sys.platform != "win32":
        QMessageBox.critical(None, "错误", "仅支持 Windows")
        sys.exit(1)
    
    # 检查管理员权限，如果未以管理员运行则自动提权（键盘钩子需要管理员权限）
    if not is_admin():
        log_message("INFO", "当前不是管理员权限，正在请求提权...")
        elevate()
        # elevate 会重启程序，不会执行后续代码
    
    init_logging()
    
    single = SingleInstance("Flashshot")
    if not single.try_acquire():
        temp_app = QApplication(sys.argv)
        tray = QSystemTrayIcon()
        pix = QPixmap(16, 16)
        pix.fill(Qt.transparent)
        tray.setIcon(QIcon(pix))
        tray.show()
        tray.showMessage("Flashshot", "程序正在运行中", QSystemTrayIcon.Information, 1000)
        QTimer.singleShot(1500, temp_app.quit)
        temp_app.exec()
        sys.exit(0)
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    QThreadPool.globalInstance().setMaxThreadCount(4)
    window = FlashshotApp()
    
    exit_code = app.exec()
    
    # 额外清理：确保钩子线程已结束
    log_message("INFO", f"应用程序退出，退出码: {exit_code}")
    
    sys.exit(exit_code)

if __name__ == "__main__":
    main()