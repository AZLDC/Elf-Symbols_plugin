# -*- coding: utf-8 -*-
"""
功能：Alt-Alt-W 觸發新注音輸入法「繁體/簡體輸出」切換
      依輸入法狀態更新托盤圖示並可由托盤切換
"""
import importlib
import os
import subprocess
import sys
import time
import threading
import ctypes
import winreg
from ctypes import wintypes
from dataclasses import dataclass
from typing import Optional

# 舊版 Python 可能沒有 wintypes.LRESULT，需自行定義
if not hasattr(wintypes, "LRESULT"):
    wintypes.LRESULT = ctypes.c_long
for _name in ("HCURSOR", "HICON", "HBRUSH", "HMENU"):
    if not hasattr(wintypes, _name):
        setattr(wintypes, _name, wintypes.HANDLE)

# 64 位元系統上的 LONG_PTR/ULONG_PTR 需擴充成 64 位，避免 Win32 回呼參數 overflow。
if ctypes.sizeof(ctypes.c_void_p) == 8:
    wintypes.LPARAM = ctypes.c_longlong
    wintypes.WPARAM = ctypes.c_ulonglong
    wintypes.LRESULT = ctypes.c_longlong

# 預先檢查依賴並在命令提示字元輸出錯誤
PROGRAM_START_TIME = time.time()
DEPENDENCY_READY_TIME = PROGRAM_START_TIME

def _require_plugin(module_path: str, friendly_name: str):
    try:
        return importlib.import_module(module_path)
    except ImportError as exc:
        print(f"缺少 {friendly_name}，程式無法繼續執行。請先安裝對應套件。")
        print(f"模組路徑：{module_path}")
        print(f"詳細錯誤：{exc}")
        raise SystemExit(1)

keyboard = _require_plugin("keyboard", "Keyboard 全域快捷鍵模組")

# 全域參數集中在這裡：包含註冊 AUMID 供系統識別、自訂提示文字、快速鍵時序與相關等待時間。
# 這些值被設為常數是為了方便調整與維護，避免魔法數散落各處。
MODEL_ID = "com.AZLDC.TCSCTRN"
TITLE_HINT = "模式切換 Alt、Alt、W"
SEQ_TIMEOUT = 0.5  # Alt-Alt 時間窗口（秒）
READ_RETRY_MAX = 5  # 重試次數
INPUT_LANG_EN = "00000409"  # 英文輸入法
INPUT_LANG_BOPOMO = "00000404"  # 新注音/注音輸入法
VALUE_NAME = "Enable Simplified Chinese Output"
# 用來偵測新注音的版本，會從下面註冊檔資訊找對應的機碼
# 依序羅列常見 Office/IME 版本，若未找到會讓後續流程決定要建立哪個節點。
CANDIDATE_PATHS = [
    r"SOFTWARE\Microsoft\IME\16.0\IMETC",
    r"SOFTWARE\Microsoft\IME\15.0\IMETC",
    r"SOFTWARE\Microsoft\IME\14.0\IMETC",
    r"SOFTWARE\Microsoft\IME\13.0\IMETC",
]
# 註冊 AppUserModelID 讓托盤圖示與通知可被系統正確歸屬，避免被視為未知程式。
def set_app_user_model_id(app_id: str) -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass

# 根據目前腳本所在資料夾準備圖示與資源，避免使用者移動程式後造成檔案遺失。
# PyInstaller 打包後資源會解壓到 sys._MEIPASS，需特別處理路徑。
def _resource_path(relative_path: str) -> str:
    """取得資源檔案的絕對路徑，相容 PyInstaller 打包環境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包後的臨時資源目錄
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

ICON_TRAD = _resource_path("繁.png")
ICON_SIMP = _resource_path("簡.png")
ICON_DEFAULT = _resource_path("轉.png")
# 啟動後保留預設圖示的秒數，避免托盤頻繁閃爍，也提供系統時間完成初始化。
INITIAL_ICON_HOLD_SECONDS = 3.0
CTFMON_RESTART_DELAY = 0.1
SYSTEM_ROOT = os.environ.get("SystemRoot", r"C:\\Windows")
CTFMON_EXE_PATH = os.path.join(SYSTEM_ROOT, "System32", "ctfmon.exe")
CREATE_NO_WINDOW = 0x08000000

# Win32/GDI+ 輔助定義，取代第三方托盤套件並讓 PNG 圖示能被載入
user32 = ctypes.windll.user32
shell32 = ctypes.windll.shell32
gdi32 = ctypes.windll.gdi32
gdiplus = ctypes.windll.gdiplus
kernel32 = ctypes.windll.kernel32

user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = wintypes.LRESULT

IDI_APPLICATION = 0x7F00
IDC_ARROW = 0x7F00
WM_DESTROY = 0x0002
WM_COMMAND = 0x0111
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205
WM_CONTEXTMENU = 0x007B
WM_APP = 0x8000
NIF_MESSAGE = 0x0001
NIF_ICON = 0x0002
NIF_TIP = 0x0004
NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
TPM_LEFTALIGN = 0x0000
TPM_BOTTOMALIGN = 0x0020
TPM_RETURNCMD = 0x0100
MF_STRING = 0x0000

def MAKEINTRESOURCE(value: int):
    return ctypes.cast(ctypes.c_void_p(value), wintypes.LPWSTR)

class GdiplusStartupInput(ctypes.Structure):
    _fields_ = [
        ("GdiplusVersion", ctypes.c_uint32),
        ("DebugEventCallback", ctypes.c_void_p),
        ("SuppressBackgroundThread", wintypes.BOOL),
        ("SuppressExternalCodecs", wintypes.BOOL),
    ]

class ICONINFO(ctypes.Structure):
    _fields_ = [
        ("fIcon", wintypes.BOOL),
        ("xHotspot", wintypes.DWORD),
        ("yHotspot", wintypes.DWORD),
        ("hbmMask", wintypes.HBITMAP),
        ("hbmColor", wintypes.HBITMAP),
    ]

class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeoutOrVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wintypes.HICON),
    ]

WNDPROC = ctypes.WINFUNCTYPE(wintypes.LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

class WNDCLASS(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HCURSOR),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]

gdiplus_token = ctypes.c_ulong(0)

def _startup_gdiplus() -> None:
    if gdiplus_token.value:
        return
    startup_input = GdiplusStartupInput(1, None, False, False)
    status = gdiplus.GdiplusStartup(ctypes.byref(gdiplus_token), ctypes.byref(startup_input), None)
    if status != 0:
        raise RuntimeError(f"GdiplusStartup failed with status {status}")

def _shutdown_gdiplus() -> None:
    if gdiplus_token.value:
        gdiplus.GdiplusShutdown(gdiplus_token)
        gdiplus_token.value = 0

def _create_icon_from_png(path: str) -> Optional[wintypes.HICON]:
    if not os.path.exists(path):
        return None
    try:
        _startup_gdiplus()
    except Exception as exc:
        print(f"GDI+ 初始化失敗：{exc}")
        return None

    bitmap = ctypes.c_void_p()
    status = gdiplus.GdipCreateBitmapFromFile(ctypes.c_wchar_p(path), ctypes.byref(bitmap))
    if status != 0 or not bitmap:
        return None

    width = ctypes.c_uint()
    height = ctypes.c_uint()
    gdiplus.GdipGetImageWidth(bitmap, ctypes.byref(width))
    gdiplus.GdipGetImageHeight(bitmap, ctypes.byref(height))

    hbm_color = wintypes.HBITMAP()
    status = gdiplus.GdipCreateHBITMAPFromBitmap(bitmap, ctypes.byref(hbm_color), ctypes.c_uint32(0x00FFFFFF))
    if status != 0 or not hbm_color:
        gdiplus.GdipDisposeImage(bitmap)
        return None

    hbm_mask = gdi32.CreateBitmap(width.value, height.value, 1, 1, None)
    icon_info = ICONINFO()
    icon_info.fIcon = True
    icon_info.xHotspot = 0
    icon_info.yHotspot = 0
    icon_info.hbmColor = hbm_color
    icon_info.hbmMask = hbm_mask
    hicon = user32.CreateIconIndirect(ctypes.byref(icon_info))

    if hbm_mask:
        gdi32.DeleteObject(hbm_mask)
    if hbm_color:
        gdi32.DeleteObject(hbm_color)
    gdiplus.GdipDisposeImage(bitmap)

    return hicon if hicon else None


def _restart_ctfmon() -> None:
    """重新啟動 ctfmon.exe 以刷新輸入法喵"""
    if os.name != "nt":
        return

    try:
        subprocess.run(
            ["taskkill", "/IM", "ctfmon.exe", "/F"],
            check=False,
            capture_output=True,
            creationflags=CREATE_NO_WINDOW,
        )
    except FileNotFoundError:
        print("找不到 taskkill，無法先關閉 ctfmon.exe")
    except Exception as exc:
        print(f"關閉 ctfmon.exe 失敗：{exc}")

    if not os.path.exists(CTFMON_EXE_PATH):
        print("找不到 ctfmon.exe，請檢查系統檔案")
        return

    try:
        subprocess.Popen(CTFMON_EXE_PATH, creationflags=CREATE_NO_WINDOW)
    except Exception as exc:
        print(f"啟動 ctfmon.exe 失敗：{exc}")

_icon_handle_cache: dict[str, wintypes.HICON] = {}

def _get_icon_handle(path: str) -> wintypes.HICON:
    handle = _icon_handle_cache.get(path)
    if handle:
        return handle
    handle = _create_icon_from_png(path)
    if not handle:
        handle = user32.LoadIconW(None, MAKEINTRESOURCE(IDI_APPLICATION))
    _icon_handle_cache[path] = handle
    return handle

def _cleanup_icon_cache() -> None:
    for handle in _icon_handle_cache.values():
        if handle:
            user32.DestroyIcon(handle)
    _icon_handle_cache.clear()

# 將註冊值資訊包成 dataclass，方便一次帶出路徑、鍵名、原始值與型別，減少傳遞多個參數的錯誤。
@dataclass(frozen=True)
class RegValueInfo:
    path: str
    value_name: str
    raw: object
    reg_type: int

# 統一管理開啟註冊機碼的方式，集中在 HKEY_CURRENT_USER，方便日後調整或補強權限處理。
def _open_key(path: str, access: int):
    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, access)

# 嘗試讀取指定路徑的值，失敗時回傳 None 避免中斷流程，讓外層可繼續嘗試下一個路徑。
def _try_read_value(path: str, value_name: str):
    try:
        with _open_key(path, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, value_name)
    except OSError:
        return None

# 逐一檢查可能的 IME 設定路徑，找到後即回傳詳細資訊，若都沒有就交給後續建立預設值。
def find_value() -> Optional[RegValueInfo]:
    for candidate in CANDIDATE_PATHS:
        got = _try_read_value(candidate, VALUE_NAME)
        if got:
            raw, reg_type = got
            return RegValueInfo(candidate, VALUE_NAME, raw, reg_type)
    return None

# 將註冊表中的任何字串或整數結果整理成 0/1，確保後續邏輯不用處理多種格式。
def _coerce_to_bit(raw: object) -> Optional[int]:
    if raw is None:
        return None
    if isinstance(raw, int):
        return raw if raw in (0, 1) else None
    if isinstance(raw, str):
        text = raw.strip().lower()
        if text in ("0", "1"):
            return int(text)
        if text.startswith("0x"):
            try:
                # 有些版本會以 0x00000001 這種格式儲存，需轉回十進位再判斷。
                val = int(text, 16)
                if val in (0, 1):
                    return val
                if val in (0x00000000, 0x00000001):
                    return 1 if val == 0x00000001 else 0
            except ValueError:
                return None
    return None

# 根據位元值換成人類可懂的狀態文字，讓托盤訊息與印出內容友善好讀。
def get_mode_text(bit: Optional[int]) -> str:
    if bit == 0:
        return "繁體輸出"
    if bit == 1:
        return "簡體輸出"
    return "狀態更新中"

# 先確保路徑存在，後續才寫得進去；這裡用 CreateKey 可以同時作為存在檢查。
def _ensure_key(path: str):
    winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)

# 將新的繁/簡模式寫回註冊表，並試著沿用既有型別，避免破壞原設定。
def set_output_mode(target_bit: int) -> tuple[bool, str]:
    info = find_value()
    target_path = info.path if info else CANDIDATE_PATHS[0]
    _ensure_key(target_path)

    current_type = info.reg_type if info else winreg.REG_DWORD

    try:
        with _open_key(target_path, winreg.KEY_SET_VALUE) as key:
            if current_type == winreg.REG_SZ:
                # 原本是文字格式就維持文字格式，使其他工具仍能讀懂。
                data = "0x00000001" if target_bit == 1 else "0x00000000"
                winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, data)
            else:
                # 預設走 REG_DWORD，Windows 也能直接識別。
                winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_DWORD, int(target_bit))
        return True, target_path
    except OSError as exc:
        return False, f"{target_path} ({exc})"

# 組合托盤提示文字：用兩行呈現目前輸出狀態與快捷提示，減少使用者忘記操作方式。
def build_tray_title(mode_text: str) -> str:
    return f"輸入法輸出 - {mode_text}\n{TITLE_HINT}"

ime_mode_bit: Optional[int] = None
_alt_timestamps = []
_waiting_for_w = False
_w_hook = None
_timeout_timer = None
_tray_updates_enabled = False
_initial_icon_timer: Optional[threading.Timer] = None

WM_TRAYICON = WM_APP + 1
IDM_TOGGLE = 1001
IDM_EXIT = 1002
_tray_hwnd: Optional[wintypes.HWND] = None
_tray_wndproc: Optional[WNDPROC] = None
_tray_icon_added = False

# 根據目前狀態挑選對應的圖示與說明，若初始化尚未完成就先顯示預設圖示。
def _resolve_icon_assets() -> tuple[str, str]:
    if not _tray_updates_enabled:
        return ICON_DEFAULT, build_tray_title("正在初始化")
    if ime_mode_bit == 1:
        return ICON_SIMP, build_tray_title("簡體輸出")
    if ime_mode_bit == 0:
        return ICON_TRAD, build_tray_title("繁體輸出")
    return ICON_DEFAULT, build_tray_title("狀態更新中")


# 延遲載入的圖示確保在資源準備好後再打開，即便啟動時檔案仍在讀取中也能平順過渡。
def _enable_tray_updates() -> None:
    global _tray_updates_enabled, _initial_icon_timer
    _tray_updates_enabled = True
    _initial_icon_timer = None
    update_tray_icon()


# 啟動時先顯示預設圖示一段時間，避免瞬間切來切去造成閃爍，也留下緩衝更新狀態。
def _compute_initial_hold_delay() -> float:
    install_duration = max(0.0, DEPENDENCY_READY_TIME - PROGRAM_START_TIME)
    target_hold = max(INITIAL_ICON_HOLD_SECONDS, install_duration)
    elapsed = time.time() - PROGRAM_START_TIME
    remaining = target_hold - elapsed
    return remaining if remaining > 0 else 0.0


def _schedule_initial_icon_release(delay: float) -> None:
    global _initial_icon_timer, _tray_updates_enabled
    _tray_updates_enabled = False
    if _initial_icon_timer is not None:
        _initial_icon_timer.cancel()
    # 以背景計時器控制何時允許托盤更新，避免主緒阻塞。
    timer = threading.Timer(delay, _enable_tray_updates)
    timer.daemon = True
    _initial_icon_timer = timer
    timer.start()

_tray_class_name = "OfficeHealthIMETrayWindow"
_tray_nid: Optional[NOTIFYICONDATAW] = None
_tray_menu: Optional[wintypes.HMENU] = None


def _tray_window_proc(hwnd: wintypes.HWND, msg: wintypes.UINT, wparam: wintypes.WPARAM, lparam: wintypes.LPARAM):
    if msg == WM_TRAYICON:
        if lparam == WM_LBUTTONUP:
            toggle_ime_mode()
        elif lparam in (WM_RBUTTONUP, WM_CONTEXTMENU):
            _show_context_menu(hwnd)
        return 0
    if msg == WM_COMMAND:
        command_id = wparam & 0xFFFF
        if command_id == IDM_TOGGLE:
            toggle_ime_mode()
        elif command_id == IDM_EXIT:
            user32.PostMessageW(hwnd, WM_DESTROY, 0, 0)
        return 0
    if msg == WM_DESTROY:
        _release_tray_resources()
        user32.PostQuitMessage(0)
        return 0
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


def _register_tray_window_class() -> None:
    global _tray_wndproc
    if _tray_wndproc is not None:
        return
    hinstance = kernel32.GetModuleHandleW(None)
    wndclass = WNDCLASS()
    _tray_wndproc = WNDPROC(_tray_window_proc)
    wndclass.lpfnWndProc = _tray_wndproc
    wndclass.hInstance = hinstance
    wndclass.lpszClassName = _tray_class_name
    wndclass.hCursor = user32.LoadCursorW(None, MAKEINTRESOURCE(IDC_ARROW))
    wndclass.hIcon = user32.LoadIconW(None, MAKEINTRESOURCE(IDI_APPLICATION))
    wndclass.hbrBackground = 0
    atom = user32.RegisterClassW(ctypes.byref(wndclass))
    if not atom:
        err = kernel32.GetLastError()
        if err != 1410:  # ERROR_CLASS_ALREADY_EXISTS
            raise ctypes.WinError(err)


def _create_tray_window() -> wintypes.HWND:
    global _tray_hwnd
    if _tray_hwnd:
        return _tray_hwnd
    _register_tray_window_class()
    hinstance = kernel32.GetModuleHandleW(None)
    hwnd = user32.CreateWindowExW(
        0,
        _tray_class_name,
        "OfficeHealthIMETray",
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        hinstance,
        None,
    )
    if not hwnd:
        raise ctypes.WinError(ctypes.get_last_error())
    _tray_hwnd = hwnd
    return hwnd


def _ensure_tray_menu() -> wintypes.HMENU:
    global _tray_menu
    if _tray_menu:
        return _tray_menu
    menu = user32.CreatePopupMenu()
    if not menu:
        raise ctypes.WinError(ctypes.get_last_error())
    # 
    # 關閉
    user32.AppendMenuW(menu, MF_STRING, IDM_EXIT, "關閉APP")
    _tray_menu = menu
    return menu


def _destroy_tray_menu() -> None:
    global _tray_menu
    if _tray_menu:
        user32.DestroyMenu(_tray_menu)
        _tray_menu = None


def _show_context_menu(hwnd: wintypes.HWND) -> None:
    menu = _ensure_tray_menu()
    point = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(point))
    user32.SetForegroundWindow(hwnd)
    cmd = user32.TrackPopupMenu(
        menu,
        TPM_LEFTALIGN | TPM_BOTTOMALIGN | TPM_RETURNCMD,
        point.x,
        point.y,
        0,
        hwnd,
        None,
    )
    if cmd:
        user32.PostMessageW(hwnd, WM_COMMAND, cmd, 0)


def _ensure_notify_icon() -> None:
    global _tray_nid, _tray_icon_added
    if _tray_icon_added and _tray_nid:
        return
    hwnd = _create_tray_window()
    icon_path, title = _resolve_icon_assets()
    nid = NOTIFYICONDATAW()
    nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
    nid.hWnd = hwnd
    nid.uID = 1
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
    nid.uCallbackMessage = WM_TRAYICON
    nid.hIcon = _get_icon_handle(icon_path)
    nid.szTip = title[:127]
    if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
        raise ctypes.WinError(ctypes.get_last_error())
    _tray_nid = nid
    _tray_icon_added = True


def _remove_notify_icon() -> None:
    global _tray_nid, _tray_icon_added
    if _tray_icon_added and _tray_nid:
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(_tray_nid))
    _tray_nid = None
    _tray_icon_added = False


def _release_tray_resources() -> None:
    global _tray_hwnd
    _remove_notify_icon()
    _destroy_tray_menu()
    _cleanup_icon_cache()
    _shutdown_gdiplus()
    if _tray_hwnd:
        user32.DestroyWindow(_tray_hwnd)
        _tray_hwnd = None


def update_tray_icon() -> None:
    """更新托盤圖示"""
    global _tray_nid
    try:
        _ensure_notify_icon()
    except Exception as exc:
        print(f"托盤初始化失敗：{exc}")
        return

    if not _tray_nid:
        return

    icon_path, title = _resolve_icon_assets()
    _tray_nid.hIcon = _get_icon_handle(icon_path)
    _tray_nid.szTip = title[:127]
    shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(_tray_nid))

# 週期性讀取輸出模式；若剛寫入會先等一下再讀，確保系統真的套用到期望值。
def refresh_ime_state(wait_for: Optional[int] = None) -> None:
    """讀取目前 IME 輸出模式，可等待指定結果"""
    global ime_mode_bit
    attempts = 0
    bit: Optional[int] = None
    info = None
    while attempts < READ_RETRY_MAX:
        # 每次讀取最新的註冊表內容，找到就轉換成 0/1。
        info = find_value()
        if not info:
            bit = None
            break
        bit = _coerce_to_bit(info.raw)
        if wait_for is None or bit == wait_for:
            # 若呼叫方期待特定值就持續重試，避免讀到舊狀態。
            break
        attempts += 1

    ime_mode_bit = bit
    if not info:
        print("找不到 Enable Simplified Chinese Output，可能尚未建立或版本路徑不同")
    else:
        mode_text = get_mode_text(ime_mode_bit)
        print(f"目前輸入法模式：{mode_text} (值={info.raw})")

    update_tray_icon()

# 封裝向前景視窗送出輸入法切換訊號的細節，避免主流程需要處理 HWND 與訊息代碼。
def _post_input_lang_request(hwnd: ctypes.c_void_p, layout_id: str) -> bool:
    user32 = ctypes.windll.user32
    hkl = user32.LoadKeyboardLayoutW(layout_id, 1)
    if not hkl:
        return False
    # 送出 WM_INPUTLANGCHANGEREQUEST，請目前視窗切換成指定的輸入法配置。
    result = user32.PostMessageW(hwnd, 0x0050, 0, hkl)  # WM_INPUTLANGCHANGEREQUEST
    return bool(result)

# 切換至英文再切回注音，藉由「跳一次」讓輸出模式改變被系統刷新。
def _refresh_input_language() -> None:
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return

    # 先切到英文再切回注音，透過「重新激活」迫使 IME 寫入的新模式立即生效。
    switched_en = _post_input_lang_request(hwnd, INPUT_LANG_EN)
    if switched_en:
        _restart_ctfmon()
    _post_input_lang_request(hwnd, INPUT_LANG_BOPOMO)

# 讀取當前狀態、寫入對側模式，再觸發輸入法刷新，確保使用者立即感受到變更。
def toggle_ime_mode() -> None:
    """切換 IME 繁/簡輸出"""
    global ime_mode_bit

    refresh_ime_state()
    current = ime_mode_bit if ime_mode_bit in (0, 1) else 0
    # 目標值永遠是目前狀態的反面，這樣可確保每次切換都能往另一路徑。
    target = 0 if current == 1 else 1

    ok, detail = set_output_mode(target)
    if not ok:
        print(f"寫入失敗：{detail}")
        return
    _refresh_input_language()
    refresh_ime_state(wait_for=target)

# 中斷 Alt-Alt 序列後的 W 鍵等待，並釋放 hook/timer，避免掛鉤一直存在造成資源佔用。
def _stop_waiting_for_w() -> None:
    global _waiting_for_w, _w_hook, _timeout_timer
    _waiting_for_w = False
    if _w_hook is not None:
        keyboard.unhook(_w_hook)
        _w_hook = None
    if _timeout_timer is not None:
        _timeout_timer.cancel()
        _timeout_timer = None

# Alt-Alt 成功後監聽 W 鍵，確保只有按下 W 才觸發切換，並在執行完畢後解除 hook。
def _on_w_hook(event: keyboard.KeyboardEvent) -> bool:
    global _waiting_for_w, _w_hook, _timeout_timer

    if event.event_type != "down":
        return True

    key_name = event.name.lower() if event.name else ""
    if key_name == "w" and _waiting_for_w:
        _waiting_for_w = False
        if _timeout_timer is not None:
            _timeout_timer.cancel()
            _timeout_timer = None

        def delayed_toggle():
            global _w_hook
            time.sleep(0.01)
            if _w_hook is not None:
                keyboard.unhook(_w_hook)
                _w_hook = None
            # 以小延遲避免鍵盤事件還沒清除就切換，降低干擾。
            toggle_ime_mode()

        threading.Thread(target=delayed_toggle, daemon=True).start()
        return False

    return True

# 設定 W 鍵監聽與逾時倒數，讓使用者在短時間內能完成 Alt-Alt-W，不成功就自動恢復。
def _start_waiting_for_w() -> None:
    global _waiting_for_w, _w_hook, _timeout_timer

    _stop_waiting_for_w()
    _waiting_for_w = True
    _w_hook = keyboard.hook(_on_w_hook, suppress=True)
    # 超過時間沒按 W 就取消等待，避免長時間鎖住鍵盤事件。
    _timeout_timer = threading.Timer(SEQ_TIMEOUT, _stop_waiting_for_w)
    _timeout_timer.daemon = True
    _timeout_timer.start()
    print("[等待 W 鍵以切換輸出模式...]")

def on_key_event(event: keyboard.KeyboardEvent) -> None:
    """偵測 Alt-Alt-W"""
    global _alt_timestamps

    key_name = event.name.lower() if event.name else ""

    if event.event_type == "down":
        now = time.time()
        if key_name == "alt":
            # 保留時間視窗內的 Alt 按下時間點，確保判斷只依近期操作。
            _alt_timestamps = [t for t in _alt_timestamps if now - t < SEQ_TIMEOUT]
            _alt_timestamps.append(now)
            if len(_alt_timestamps) >= 2 and now - _alt_timestamps[-2] < SEQ_TIMEOUT:
                # 兩次 Alt 夠接近就觸發等待 W，其他按鍵則視為失敗並重置。
                _alt_timestamps.clear()
                _start_waiting_for_w()
        else:
            _alt_timestamps.clear()


# Win32 訊息迴圈，維持托盤事件
def _message_loop() -> None:
    msg = wintypes.MSG()
    while True:
        result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if result == 0:  # WM_QUIT
            break
        if result == -1:
            break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))


def run_tray() -> None:
    set_app_user_model_id(MODEL_ID)
    _schedule_initial_icon_release(_compute_initial_hold_delay())
    _create_tray_window()
    update_tray_icon()
    _message_loop()


# 程式進入點：先同步狀態、掛上鍵盤全域監聽，再提示使用方式，最後進入托盤主迴圈。
def main() -> None:
    refresh_ime_state()
    keyboard.hook(on_key_event)

    # 以文字提示提醒使用者操作方式，當托盤圖示不可見時仍能得知快捷鍵。
    print("Alt-Alt-W = 切換輸入法繁/簡輸出")
    print("左鍵點擊托盤圖示 = 切換輸入法繁/簡輸出")
    print("托盤圖示執行中...")

    run_tray()

if __name__ == "__main__":
    main()
