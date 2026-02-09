# -*- coding: utf-8 -*-
"""
功能 : Alt-Alt-W 觸發新注音輸入法「繁體/簡體輸出」切換
      依輸入法狀態更新托盤圖示並可由托盤切換
"""
# ============================================================================
# 匯入區塊
# ============================================================================
# 匯入 importlib 以便動態載入可選插件，避免缺模組時提早失敗
import importlib
import importlib.util
# 匯入 os 處理檔案、目錄與環境變數操作
import os
# 匯入 re 供註冊表版本字串比對使用
import re
# 匯入 subprocess 以呼叫外部程式（例如 ctfmon、taskkill）
import subprocess
# 匯入 sys 取得 PyInstaller 解壓目錄與腳本資訊
import sys
# 匯入 time 供排程、計時與重試控制
import time
# 匯入 threading 建立背景計時器與非同步操作
import threading
# 匯入 ctypes 直接呼叫 Win32 API
import ctypes
# 匯入 winreg 讀寫註冊表中的 IME 設定
import winreg
# 從 ctypes 匯入 wintypes 以取得 Windows 特有型別定義
from ctypes import wintypes
# 匯入 ModuleType 讓動態載入的模組能有型別提示
from types import ModuleType
# 匯入 dataclass 讓註冊表資料封裝得更結構化
from dataclasses import dataclass
# 匯入 Optional 用於型別提示表示「可能為 None」
from typing import Optional

# 控制是否在切換輸出模式時強行關閉 ctfmon，再重新啟動
KILL_CTFMON = False

# ============================================================================
# Win32 型別補丁（須在常數定義前完成）
# ============================================================================
# 為舊版 Python 補齊缺失的 Win32 型別，避免 API 綁定時噴錯
if not hasattr(wintypes, "LRESULT"):
    # 以 c_long 模擬遺失的 LRESULT 型別
    wintypes.LRESULT = ctypes.c_long
# 逐一檢查常見的 handle 型別，若缺少就使用 HANDLE 取代
for _name in ("HCURSOR", "HICON", "HBRUSH", "HMENU"):
    if not hasattr(wintypes, _name):
        setattr(wintypes, _name, wintypes.HANDLE)

# 若執行於 64 位元環境，將 LPARAM/WPARAM/LRESULT 調整成 64 位以防溢位
if ctypes.sizeof(ctypes.c_void_p) == 8:
    wintypes.LPARAM = ctypes.c_longlong
    wintypes.WPARAM = ctypes.c_ulonglong
    wintypes.LRESULT = ctypes.c_longlong

# ============================================================================
# 時間與依賴追蹤
# ============================================================================
# 紀錄程式啟動時間，稍後計算托盤圖示延遲
PROGRAM_START_TIME = time.time()
# 在依賴確認完畢後更新，若未使用額外依賴則維持相同值
DEPENDENCY_READY_TIME = PROGRAM_START_TIME

def _require_plugin(module_path: str, friendly_name: str):
    """載入可選套件並在缺少時給出具體指引

    參數:
        module_path (str): 可由 importlib 導入的模組路徑
        friendly_name (str): 給使用者看的易讀名稱，便於辨識缺少哪個套件

    回傳:
        ModuleType: 成功載入後的模組物件，讓後續直接使用其中 API

    說明:
        - 此函式集中處理 ImportError，避免在多處重複 try/except
        - 當缺少依賴時輸出指示並立即結束，以免程式處於半初始化狀態
    """
    try:
        return importlib.import_module(module_path)
    except ImportError as exc:
        print(f"缺少 {friendly_name}，程式無法繼續執行。請先安裝對應套件。")
        print(f"模組路徑 : {module_path}")
        print(f"詳細錯誤 : {exc}")
        raise SystemExit(1)

keyboard = _require_plugin("keyboard", "Keyboard 全域快捷鍵模組")

# ============================================================================
# 應用程式常數定義
# ============================================================================
# 全域參數集中在這裡 : 包含註冊 AUMID 供系統識別、自訂提示文字、快速鍵時序與相關等待時間。
# 這些值被設為常數是為了方便調整與維護，避免魔法數散落各處。
MODEL_ID = "com.AZLDC.TCSCTRN"
TITLE_HINT = "模式切換 Alt、Alt、W"
SEQ_TIMEOUT = 0.5  # Alt-Alt 時間窗口（秒）
READ_RETRY_MAX = 5  # 重試次數
INPUT_LANG_EN = "00000409"  # 英文輸入法
INPUT_LANG_BOPOMO = "00000404"  # 新注音/注音輸入法
INPUT_LANG_BOPOMO_ID = int(INPUT_LANG_BOPOMO, 16)
INPUT_LANG_LABELS = {
    INPUT_LANG_EN.upper(): "英文鍵盤",
    INPUT_LANG_BOPOMO.upper(): "新注音鍵盤",
    "00000411": "日文鍵盤",
}
ALT_KEY_NAMES = ("alt", "left alt", "right alt")
SHIFT_KEY_NAMES = ("shift", "left shift", "right shift")
VALUE_NAME = "Enable Simplified Chinese Output"
# 針對 IME 設定的登錄路徑，改以「基底路徑 + 動態列舉 + 常見預設值」三層策略，
# 讓未知版本也能被自動納入，減少每次升級都得修改常數的情況。
IME_BASE_PATH = r"SOFTWARE\Microsoft\IME"
IME_VERSION_HINTS = ("16.0", "15.0", "14.0", "13.0")
CANDIDATE_PATHS: list[str] = []

def _build_candidate_paths() -> None:
    """依目前註冊表自動推導可用路徑，並保留常見預設值"""

    if CANDIDATE_PATHS:
        return

    seen_paths = set()

    def _append(path: str) -> None:
        normalized = path.replace("/", "\\")
        if not normalized or normalized in seen_paths:
            return
        seen_paths.add(normalized)
        CANDIDATE_PATHS.append(normalized)

    version_names: list[str] = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, IME_BASE_PATH, 0, winreg.KEY_READ) as base_key:
            index = 0
            while True:
                try:
                    sub_name = winreg.EnumKey(base_key, index)
                except OSError:
                    break
                if re.match(r"^\d+\.\d+$", sub_name):
                    version_names.append(sub_name)
                index += 1
    except OSError:
        pass

    for sub_name in sorted(version_names, reverse=True):
        _append(fr"{IME_BASE_PATH}\\{sub_name}\\IMETC")

    for version in IME_VERSION_HINTS:
        _append(fr"{IME_BASE_PATH}\\{version}\\IMETC")

    _append(fr"{IME_BASE_PATH}\\IMETC")

    if not CANDIDATE_PATHS:
        _append(fr"{IME_BASE_PATH}\\16.0\\IMETC")

_build_candidate_paths()

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

# --- 系統路徑與外部程式 ---
SYSTEM_ROOT = os.environ.get("SystemRoot", r"C:\\Windows")
SYSTEM32_PATH = os.path.join(SYSTEM_ROOT, "System32")
CTFMON_EXE_PATH = os.path.join(SYSTEM32_PATH, "ctfmon.exe")
TASKKILL_EXE_PATH = os.path.join(SYSTEM32_PATH, "taskkill.exe")

# ============================================================================
# Cursors_FIX 整合（選用）
# ============================================================================
_CURSOR_FIX_FILENAME = "Cursors_FIX.py"

def _cursor_fix_script_path() -> str:
    if hasattr(sys, "_MEIPASS"):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, _CURSOR_FIX_FILENAME)

def _load_cursor_fix_module() -> Optional[ModuleType]:
    global _cursor_fix_module
    if _cursor_fix_module is not None:
        return _cursor_fix_module
    script_path = _cursor_fix_script_path()
    if not os.path.exists(script_path):
        return None
    spec = importlib.util.spec_from_file_location("Cursors_FIX_embed", script_path)
    if spec is None or spec.loader is None:
        print("Cursors_FIX 解析失敗，略過啟動")
        return None
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        print(f"載入 Cursors_FIX 失敗 : {exc}")
        return None
    _cursor_fix_module = module
    return module

def _cursor_fix_worker(module: ModuleType) -> None:
    get_cursor_handle = getattr(module, "get_cursor_handle", None)
    force_reload = getattr(module, "force_reload_cursors", None)
    if not callable(get_cursor_handle) or not callable(force_reload):
        print("Cursors_FIX 模組缺少必要函式，無法啟動")
        return

    poll_interval = getattr(module, "POLL_INTERVAL_SEC", getattr(module, "poll_interval_sec", 0.02))
    cooldown = getattr(module, "COOLDOWN_SEC", getattr(module, "cooldown_sec", 0.20))
    last_handle = None
    cooldown_until = 0.0
    print("Cursors_FIX 游標監控執行緒啟動")

    while not _cursor_fix_stop_event.wait(poll_interval):
        handle = get_cursor_handle()
        if not handle:
            continue
        if last_handle is None:
            last_handle = handle
            print(f"[Cursors_FIX] 初始游標 : {handle:#010x}")
            continue
        if handle != last_handle:
            now = time.monotonic()
            if now >= cooldown_until:
                print(f"[Cursors_FIX] 游標切換 {last_handle:#010x} -> {handle:#010x}，觸發重載")
                try:
                    force_reload()
                except Exception as exc:
                    print(f"[Cursors_FIX] 重載失敗 : {exc}")
                cooldown_until = now + cooldown
            last_handle = handle

    print("Cursors_FIX 游標監控執行緒結束")

def _start_cursor_fix_monitor() -> None:
    global _cursor_fix_thread
    if _cursor_fix_thread is not None:
        return
    module = _load_cursor_fix_module()
    if module is None:
        return
    _cursor_fix_stop_event.clear()
    thread = threading.Thread(target=_cursor_fix_worker, args=(module,), daemon=True)
    _cursor_fix_thread = thread
    thread.start()

def _stop_cursor_fix_monitor() -> None:
    global _cursor_fix_thread
    if _cursor_fix_thread is None:
        return
    _cursor_fix_stop_event.set()
    _cursor_fix_thread.join(timeout=2.0)
    _cursor_fix_thread = None
    _cursor_fix_stop_event.clear()
SW_HIDE = 0

# ============================================================================
# Win32 API 綁定與常數
# ============================================================================
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
EVENT_SYSTEM_FOREGROUND = 0x0003
EVENT_SYSTEM_SWITCHSTART = 0x0014
EVENT_SYSTEM_SWITCHEND = 0x0015
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNPROCESS = 0x0002
HSHELL_LANGUAGE = 0x002A
HSHELL_RUDEAPPACTIVATED = 0x8000
WM_SHELLHOOKMESSAGE = user32.RegisterWindowMessageW("SHELLHOOK")

def MAKEINTRESOURCE(value: int):
    return ctypes.cast(ctypes.c_void_p(value), wintypes.LPWSTR)

# ============================================================================
# Win32 / ctypes 結構定義
# ============================================================================
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
WINEVENTPROC = ctypes.WINFUNCTYPE(
    None,
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.HWND,
    wintypes.LONG,
    wintypes.LONG,
    wintypes.DWORD,
    wintypes.DWORD,
)

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

# Win32 GUI thread 相關結構：供 GetGUIThreadInfo 回填焦點視窗與插入號資訊。
# RECT 結構定義矩形區域的座標，分別為左、上、右、下邊界
class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),  # 矩形左邊界
        ("top", wintypes.LONG),   # 矩形上邊界
        ("right", wintypes.LONG), # 矩形右邊界
        ("bottom", wintypes.LONG),# 矩形下邊界
    ]

class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", RECT),
    ]

# ============================================================================
# 全域狀態變數
# ============================================================================
# --- GDI+ 狀態 ---
gdiplus_token = ctypes.c_ulong(0)

# --- 前景視窗監控狀態 ---
_win_event_hook: Optional[wintypes.HANDLE] = None
_win_event_proc: Optional[WINEVENTPROC] = None
_last_foreground_hwnd: Optional[wintypes.HWND] = None

# --- IME 狀態 ---
ime_mode_bit: Optional[int] = None
_current_input_layout: Optional[int] = None

# --- 快捷鍵狀態 ---
_alt_timestamps: list = []
_waiting_for_w: bool = False
_w_hook = None
_timeout_timer = None
_alt_shift_probe_ts: float = 0.0
_alt_down: bool = False
_shift_down: bool = False

# --- 托盤狀態 ---
_tray_updates_enabled: bool = False
_initial_icon_timer: Optional[threading.Timer] = None
WM_TRAYICON = WM_APP + 1
IDM_TOGGLE = 1001
IDM_EXIT = 1002
_tray_hwnd: Optional[wintypes.HWND] = None
_tray_wndproc: Optional[WNDPROC] = None
_shell_hook_registered: bool = False
_tray_icon_added: bool = False
_tray_class_name = "OfficeHealthIMETrayWindow"
_tray_nid: Optional[NOTIFYICONDATAW] = None
_tray_menu: Optional[wintypes.HMENU] = None
_icon_handle_cache: dict[str, wintypes.HICON] = {}

# --- Cursors_FIX 狀態 ---
_cursor_fix_module: Optional[ModuleType] = None
_cursor_fix_thread: Optional[threading.Thread] = None
_cursor_fix_stop_event = threading.Event()

# ============================================================================
# GDI+ 圖示處理函數
# ============================================================================
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
        print(f"GDI+ 初始化失敗 : {exc}")
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

# ============================================================================
# ctfmon 控制函數
# ============================================================================
def _start_ctfmon_process() -> bool:
    """以多種方式啟動 ctfmon，避免權限需求"""
    if not os.path.exists(CTFMON_EXE_PATH):
        print("找不到 ctfmon.exe，請檢查系統檔案")
        return False

    try:
        subprocess.Popen(CTFMON_EXE_PATH, creationflags=CREATE_NO_WINDOW)
        return True
    except OSError as exc:
        if getattr(exc, "winerror", None) != 740:
            print(f"以 Popen 啟動 ctfmon.exe 失敗 : {exc}")
    except Exception as exc:
        print(f"以 Popen 啟動 ctfmon.exe 失敗 : {exc}")

    try:
        result = shell32.ShellExecuteW(None, "open", CTFMON_EXE_PATH, None, SYSTEM_ROOT, SW_HIDE)
        if result > 32:
            return True
        print(f"ShellExecute 啟動 ctfmon.exe 失敗（代碼 {result}）")
    except Exception as exc:
        print(f"ShellExecute 啟動 ctfmon.exe 失敗 : {exc}")
    return False

def _kill_ctfmon_elevated() -> bool:
    """使用提升權限的 taskkill 終止 ctfmon"""
    executable = TASKKILL_EXE_PATH if os.path.exists(TASKKILL_EXE_PATH) else "taskkill.exe"
    args = "/IM ctfmon.exe /F"
    try:
        result = shell32.ShellExecuteW(None, "runas", executable, args, None, SW_HIDE)
        if result > 32:
            print("已透過提升權限的 taskkill 終止 ctfmon.exe")
            return True
        print(f"提升權限的 taskkill 失敗（代碼 {result}）")
    except Exception as exc:
        print(f"提升權限的 taskkill 失敗 : {exc}")
    return False

def _kill_ctfmon_process() -> bool:
    """嘗試終止既有的 ctfmon.exe，必要時提升權限"""
    escalation_needed = False
    try:
        result = subprocess.run(
            ["taskkill", "/IM", "ctfmon.exe", "/F"],
            check=False,
            capture_output=True,
            creationflags=CREATE_NO_WINDOW,
            text=True,
        )
    except FileNotFoundError:
        print("找不到 taskkill，無法關閉 ctfmon.exe")
        return False
    except PermissionError:
        print("taskkill 需要提升權限，嘗試升級")
        escalation_needed = True
    except Exception as exc:
        print(f"關閉 ctfmon.exe 失敗 : {exc}")
        return False
    else:
        if result.returncode == 0:
            return True
        output = ((result.stdout or "") + (result.stderr or "")).strip()
        if "存取被拒" in output:
            print("taskkill 無權終止 ctfmon.exe（存取被拒），嘗試升級")
            escalation_needed = True
        else:
            detail = output or f"exit code {result.returncode}"
            print(f"taskkill ctfmon.exe 失敗 : {detail}")
            return False

    if not escalation_needed:
        return False
    return _kill_ctfmon_elevated()

def _restart_ctfmon() -> None:
    """重新啟動 ctfmon.exe 以刷新輸入法"""
    if os.name != "nt":
        return

    if KILL_CTFMON:
        if not _kill_ctfmon_process():
            print("無法終止既有的 ctfmon.exe，請以系統管理員身分手動關閉後重試")
            return

    if not _start_ctfmon_process():
        print("無法重新啟動 ctfmon.exe，輸入法可能無法立即刷洗")

# ============================================================================
# 圖示快取管理
# ============================================================================
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

# ============================================================================
# 註冊表操作函數
# ============================================================================
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

# ============================================================================
# 托盤視窗與圖示管理
# ============================================================================
# 組合托盤提示文字 : 用兩行呈現目前輸出狀態與快捷提示，減少使用者忘記操作方式。
def build_tray_title(mode_text: str) -> str:
    return f"輸入法輸出 - {mode_text}"

# 根據目前狀態挑選對應的圖示與說明，若初始化尚未完成就先顯示預設圖示。
def _resolve_icon_assets() -> tuple[str, str]:
    if not _tray_updates_enabled:
        return ICON_DEFAULT, build_tray_title("正在初始化")
    if not _is_bopomo_active():
        if _current_input_layout:
            label, _ = _describe_layout(_current_input_layout)
            status = f"目前為 {label}"
        else:
            status = "請切換至注音輸入法"
        return ICON_DEFAULT, build_tray_title(status)
    if ime_mode_bit == 1:
        return ICON_SIMP, build_tray_title(f"簡體輸出\n{TITLE_HINT}")
    if ime_mode_bit == 0:
        return ICON_TRAD, build_tray_title(f"繁體輸出\n{TITLE_HINT}")
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

def _tray_window_proc(hwnd: wintypes.HWND, msg: wintypes.UINT, wparam: wintypes.WPARAM, lparam: wintypes.LPARAM):
    if msg == WM_TRAYICON:
        if lparam == WM_LBUTTONUP:
            toggle_ime_mode()
        elif lparam in (WM_RBUTTONUP, WM_CONTEXTMENU):
            _show_context_menu(hwnd)
        return 0
    if msg == WM_SHELLHOOKMESSAGE:
        _handle_shell_hook_event(int(wparam), lparam)
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
    _register_shell_hook_window(hwnd)
    return hwnd

def _register_shell_hook_window(hwnd: wintypes.HWND) -> None:
    global _shell_hook_registered
    if _shell_hook_registered or not WM_SHELLHOOKMESSAGE:
        return
    if not user32.RegisterShellHookWindow(hwnd):
        err = ctypes.get_last_error()
        print(f"RegisterShellHookWindow 失敗 : {err}")
        return
    _shell_hook_registered = True
    print("Shell hook 已註冊")

def _unregister_shell_hook_window() -> None:
    global _shell_hook_registered
    if not _shell_hook_registered or not _tray_hwnd:
        return
    if not user32.DeregisterShellHookWindow(_tray_hwnd):
        err = ctypes.get_last_error()
        print(f"DeregisterShellHookWindow 失敗 : {err}")
    _shell_hook_registered = False

def _handle_shell_hook_event(code: int, detail: wintypes.LPARAM) -> None:
    if code == HSHELL_LANGUAGE:
        print("偵測到輸入法切換事件，重新讀取 HKL")
        _schedule_async_layout_refresh(
            "ShellHook 輸入法切換",
            attempts=4,
            delay=0.12,
            initial_delay=0.05,
        )
    elif code == HSHELL_RUDEAPPACTIVATED:
        print("偵測到前景應用程式切換事件")
        _schedule_async_layout_refresh(
            "ShellHook 前景切換",
            attempts=3,
            delay=0.12,
            initial_delay=0.08,
        )

def _ensure_tray_menu() -> wintypes.HMENU:
    global _tray_menu
    if _tray_menu:
        return _tray_menu
    menu = user32.CreatePopupMenu()
    if not menu:
        raise ctypes.WinError(ctypes.get_last_error())
    
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
    _unregister_shell_hook_window()
    _stop_foreground_monitor()
    _stop_cursor_fix_monitor()
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
        print(f"托盤初始化失敗 : {exc}")
        return

    if not _tray_nid:
        return

    icon_path, title = _resolve_icon_assets()
    _tray_nid.hIcon = _get_icon_handle(icon_path)
    _tray_nid.szTip = title[:127]
    shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(_tray_nid))

# ============================================================================
# 輸入法狀態偵測函數
# ============================================================================
def _describe_layout(hkl: int) -> tuple[str, str]:
    """依 HKL 轉換成可辨識的語系文字與顯示用十六進位碼"""
    layout_hex = f"{hkl & 0xFFFFFFFF:08X}".upper()
    lang_hex = f"{hkl & 0xFFFF:04X}".zfill(8).upper()
    label = INPUT_LANG_LABELS.get(layout_hex, INPUT_LANG_LABELS.get(lang_hex, "其他鍵盤"))
    return label, layout_hex


def _is_bopomo_layout(hkl: Optional[int]) -> bool:
    if hkl is None:
        return False
    lang_id = hkl & 0xFFFF
    if lang_id == INPUT_LANG_BOPOMO_ID:
        return True
    layout_hex = f"{hkl & 0xFFFFFFFF:08X}".upper()
    return layout_hex.endswith(INPUT_LANG_BOPOMO.upper())


def _is_bopomo_active() -> bool:
    global _current_input_layout
    if _current_input_layout is None:
        _current_input_layout = _read_foreground_layout()
    return _is_bopomo_layout(_current_input_layout)


def _format_current_layout_detail() -> str:
    if _current_input_layout is None:
        return "無法偵測目前輸入法"
    label, layout_hex = _describe_layout(_current_input_layout)
    return f"目前輸入法為 {label} (HKL={layout_hex})"


def _ensure_bopomo_input(reason: str) -> bool:
    if _is_bopomo_active():
        return True
    print(f"{reason} : 僅在注音輸入法中啟用，{_format_current_layout_detail()}")
    return False

def _read_foreground_layout() -> Optional[int]:
    """讀取目前前景視窗（或其焦點子視窗）的 HKL"""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None
    process_id = wintypes.DWORD()
    thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    if not thread_id:
        return None
    gui_info = GUITHREADINFO(cbSize=ctypes.sizeof(GUITHREADINFO))
    if user32.GetGUIThreadInfo(thread_id, ctypes.byref(gui_info)):
        focus_hwnd = gui_info.hwndFocus or gui_info.hwndActive or gui_info.hwndCaret
        if focus_hwnd:
            focus_thread_id = user32.GetWindowThreadProcessId(focus_hwnd, ctypes.byref(process_id))
            if focus_thread_id:
                thread_id = focus_thread_id

    hkl = user32.GetKeyboardLayout(thread_id)
    if hkl:
        return hkl

    current_thread_id = kernel32.GetCurrentThreadId()
    attached = False
    if current_thread_id != thread_id:
        attached = bool(user32.AttachThreadInput(current_thread_id, thread_id, True))
    try:
        hkl = user32.GetKeyboardLayout(0)
    finally:
        if attached:
            user32.AttachThreadInput(current_thread_id, thread_id, False)
    return hkl if hkl else None

def _log_current_layout(hkl: Optional[int]) -> None:
    """將目前偵測到的輸入法狀態輸出到主控台"""
    if hkl is None:
        print("目前輸入法 : 無法取得（前景視窗不存在）")
        return
    label, layout_hex = _describe_layout(hkl)
    print(f"目前輸入法 : {label} (HKL={layout_hex})")

def _refresh_foreground_layout(force_log: bool = False, source: Optional[str] = None) -> None:
    """更新目前輸入法快照，必要時輸出訊息"""
    global _current_input_layout
    hkl = _read_foreground_layout()
    changed = hkl != _current_input_layout
    _current_input_layout = hkl
    if force_log or changed:
        if source:
            print(f"{source} : 重新讀取輸入法狀態")
        _log_current_layout(hkl)
    elif source:
        if hkl is None:
            print(f"{source} : 仍無法取得輸入法狀態")
        else:
            label, layout_hex = _describe_layout(hkl)
            # print(f"{source} : 輸入法維持 {label} (HKL={layout_hex})")
    if changed:
        update_tray_icon()

def _schedule_async_layout_refresh(
    source: str,
    attempts: int = 3,
    delay: float = 0.15,
    initial_delay: float = 0.0,
) -> None:
    """以背景緒多次刷新輸入法狀態，避開 WinEvent 尚未完成切換的時間差"""

    def worker() -> None:
        if initial_delay > 0:
            time.sleep(initial_delay)
        for attempt in range(attempts):
            force_log = attempt == 0
            suffix = "" if attempt == 0 else f" (重試 {attempt + 1})"
            _refresh_foreground_layout(force_log=force_log, source=f"{source}{suffix}")
            if attempt < attempts - 1:
                time.sleep(delay)

    threading.Thread(target=worker, daemon=True).start()

# ============================================================================
# 前景視窗監控（WinEvent Hook）
# ============================================================================
def _on_foreground_event(
    hook_handle: wintypes.HANDLE,
    event: wintypes.DWORD,
    hwnd: wintypes.HWND,
    id_object: wintypes.LONG,
    id_child: wintypes.LONG,
    event_thread: wintypes.DWORD,
    event_time: wintypes.DWORD,
) -> None:
    del hook_handle, id_object, id_child, event_thread, event_time
    global _last_foreground_hwnd
    if event == EVENT_SYSTEM_FOREGROUND:
        if not hwnd:
            return
        if hwnd == _last_foreground_hwnd:
            return
        _last_foreground_hwnd = hwnd
        print(f"WinEvent: 前景視窗切換到 HWND={hwnd}")
        _schedule_async_layout_refresh(
            "WinEvent 前景變更",
            attempts=3,
            delay=0.12,
            initial_delay=0.08,
        )
        return
    if event in (EVENT_SYSTEM_SWITCHSTART, EVENT_SYSTEM_SWITCHEND):
        phase = "開始" if event == EVENT_SYSTEM_SWITCHSTART else "結束"
        print(f"WinEvent: 輸入法切換{phase} (HWND={hwnd})")
        initial_delay = 0.0 if event == EVENT_SYSTEM_SWITCHEND else 0.05
        _schedule_async_layout_refresh(
            "WinEvent 輸入法切換",
            attempts=4,
            delay=0.12,
            initial_delay=initial_delay,
        )

def _start_foreground_monitor() -> None:
    """啟動 WinEvent hook 以監聽前景視窗變更"""
    global _win_event_hook, _win_event_proc, _last_foreground_hwnd
    if _win_event_hook:
        return
    callback = WINEVENTPROC(_on_foreground_event)
    hook = user32.SetWinEventHook(
        EVENT_SYSTEM_FOREGROUND,
        EVENT_SYSTEM_SWITCHEND,
        0,
        callback,
        0,
        0,
        WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
    )
    if not hook:
        raise ctypes.WinError(ctypes.get_last_error())
    _win_event_proc = callback
    _win_event_hook = hook
    _last_foreground_hwnd = user32.GetForegroundWindow()
    print("WinEvent 前景監聽已啟動")
    _refresh_foreground_layout(force_log=True, source="初始化")

def _stop_foreground_monitor() -> None:
    """解除 WinEvent hook"""
    global _win_event_hook, _win_event_proc
    if _win_event_hook:
        user32.UnhookWinEvent(_win_event_hook)
        _win_event_hook = None
    _win_event_proc = None

# ============================================================================
# IME 模式切換函數
# ============================================================================
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
        print(f"目前輸入法模式 : {mode_text} (值={info.raw})")

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

    if not _ensure_bopomo_input("切換繁/簡輸出"):
        return

    refresh_ime_state()
    current = ime_mode_bit if ime_mode_bit in (0, 1) else 0
    # 目標值永遠是目前狀態的反面，這樣可確保每次切換都能往另一路徑。
    target = 0 if current == 1 else 1

    ok, detail = set_output_mode(target)
    if not ok:
        print(f"寫入失敗 : {detail}")
        return
    _refresh_input_language()
    refresh_ime_state(wait_for=target)

# ============================================================================
# 快捷鍵處理函數（Alt-Alt-W）
# ============================================================================
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
    global _alt_down, _shift_down

    key_name = event.name.lower() if event.name else ""
    is_alt_key = key_name in ALT_KEY_NAMES
    is_shift_key = key_name in SHIFT_KEY_NAMES

    if event.event_type == "down":
        if is_shift_key:
            _shift_down = True
            if _alt_down:
                _trigger_alt_shift_probe()
        elif is_alt_key:
            was_alt_down = _alt_down
            _alt_down = True
            if not was_alt_down:
                if _shift_down:
                    _trigger_alt_shift_probe()
                now = time.time()
                # 保留時間視窗內的 Alt 按下時間點，確保判斷只依近期操作。
                _alt_timestamps = [t for t in _alt_timestamps if now - t < SEQ_TIMEOUT]
                _alt_timestamps.append(now)
                if len(_alt_timestamps) >= 2 and now - _alt_timestamps[-2] < SEQ_TIMEOUT:
                    # 兩次 Alt 夠接近就觸發等待 W，其他按鍵則視為失敗並重置。
                    _alt_timestamps.clear()
                    if _ensure_bopomo_input("Alt-Alt 快捷"):
                        _start_waiting_for_w()
        else:
            _alt_timestamps.clear()
    elif event.event_type == "up":
        if is_shift_key:
            _shift_down = False
        elif is_alt_key:
            _alt_down = False

def _trigger_alt_shift_probe() -> None:
    global _alt_shift_probe_ts
    now = time.time()
    if now - _alt_shift_probe_ts < 0.2:
        return
    _alt_shift_probe_ts = now
    print("偵測到 Alt+Shift 輸入法切換操作")
    _schedule_async_layout_refresh(
        "Alt+Shift 偵測",
        attempts=4,
        delay=0.1,
        initial_delay=0.02,
    )

# ============================================================================
# 訊息迴圈與主程式入口
# ============================================================================
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

# 程式進入點 : 先同步狀態、掛上鍵盤全域監聽，再提示使用方式，最後進入托盤主迴圈。
def main() -> None:
    refresh_ime_state()
    keyboard.hook(on_key_event)
    _start_foreground_monitor()
    _start_cursor_fix_monitor()

    # 以文字提示提醒使用者操作方式，當托盤圖示不可見時仍能得知快捷鍵。
    print("Alt-Alt-W = 切換輸入法繁/簡輸出")
    print("左鍵點擊托盤圖示 = 切換輸入法繁/簡輸出")
    print("托盤圖示執行中...")

    run_tray()

if __name__ == "__main__":
    main()
