import os
import time
import ctypes
import subprocess
import threading
from typing import Optional
from ctypes import wintypes

TARGET_EXE = r"C:\Program Files\ShareMouse\ShareMouse.exe"
POLL_INTERVAL_SEC = 1.0      # 功能：輪詢 ShareMouse 是否存活；單位：秒
RESTART_GRACE_SEC = 2.0      # 功能：重啟後緩衝時間，避免瞬間重複啟動；單位：秒
IDLE_TIMEOUT_SEC = 300.0     # 功能：閒置多久後關閉 ShareMouse；單位：秒

TH32CS_SNAPPROCESS = 0x00000002
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
ERROR_INSUFFICIENT_BUFFER = 122
DETACHED_PROCESS = 0x00000008
ERROR_ELEVATION_REQUIRED = 740
SW_SHOWNORMAL = 1

if not hasattr(wintypes, "ULONG_PTR"):
    pointer_size = ctypes.sizeof(ctypes.c_void_p)
    wintypes.ULONG_PTR = ctypes.c_ulonglong if pointer_size == ctypes.sizeof(ctypes.c_ulonglong) else ctypes.c_ulong

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
shell32 = ctypes.WinDLL('shell32', use_last_error=True)
user32 = ctypes.WinDLL('user32', use_last_error=True)
_watchdog_thread: Optional[threading.Thread] = None


class PROCESSENTRY32(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("cntUsage", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("th32DefaultHeapID", wintypes.ULONG_PTR),
        ("th32ModuleID", wintypes.DWORD),
        ("cntThreads", wintypes.DWORD),
        ("th32ParentProcessID", wintypes.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wintypes.DWORD),
        ("szExeFile", wintypes.WCHAR * 260),
    ]


kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
kernel32.Process32FirstW.restype = wintypes.BOOL
kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
kernel32.Process32NextW.restype = wintypes.BOOL
kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.GetTickCount.restype = wintypes.DWORD
kernel32.GetTickCount.argtypes = []


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("dwTime", wintypes.DWORD),
    ]


user32.GetLastInputInfo.restype = wintypes.BOOL
user32.GetLastInputInfo.argtypes = [ctypes.POINTER(LASTINPUTINFO)]


def _idle_seconds() -> float:
    """功能：取得使用者最後一次滑鼠/鍵盤輸入後經過的時間；單位：秒"""
    info = LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if not user32.GetLastInputInfo(ctypes.byref(info)):
        return 0.0
    tick_now = kernel32.GetTickCount()
    delta = (tick_now - info.dwTime) & 0xFFFFFFFF  # 處理 GetTickCount 49.7 天 wrap-around
    return delta / 1000.0


def _terminate_target(target_path: str) -> bool:
    """功能：強制結束 ShareMouse 程序；單位：布林值"""
    image_name = os.path.basename(target_path)
    try:
        completed = subprocess.run(
            ["taskkill", "/F", "/IM", image_name],
            creationflags=subprocess.CREATE_NO_WINDOW,
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        print(f"[ERROR] 呼叫 taskkill 失敗：{exc}")
        return False

    if completed.returncode == 0:
        print(f"[INFO] 已關閉 ShareMouse：{image_name}")
        return True

    print(
        f"[WARN] taskkill 結束 {image_name} 失敗，returncode={completed.returncode}，"
        f"stderr={(completed.stderr or '').strip()}"
    )
    return False


def _iter_process_ids():
    """功能：列舉系統中所有行程 ID；單位：無"""
    snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snapshot == INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        print(f"[DEBUG] CreateToolhelp32Snapshot 失敗，錯誤碼: {err}")
        return

    try:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)

        if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
            err = ctypes.get_last_error()
            print(f"[DEBUG] Process32FirstW 失敗，錯誤碼: {err}")
            return

        yield entry.th32ProcessID

        while kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
            yield entry.th32ProcessID
    finally:
        kernel32.CloseHandle(snapshot)


def _get_process_image(pid):
    """功能：取得指定 PID 對應的完整路徑；單位：字串"""
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None

    try:
        buf_len = wintypes.DWORD(512)
        while True:
            buf = ctypes.create_unicode_buffer(buf_len.value)
            size = wintypes.DWORD(buf_len.value)
            if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                return buf.value

            err = ctypes.get_last_error()
            if err == ERROR_INSUFFICIENT_BUFFER:
                buf_len.value *= 2
                continue

            return None
    finally:
        kernel32.CloseHandle(handle)


def _is_target_alive(target_path):
    """功能：判斷 ShareMouse 目前是否仍在執行；單位：布林值"""
    normalized_target = os.path.normcase(os.path.normpath(target_path))
    for pid in _iter_process_ids() or []:
        image = _get_process_image(pid)
        if not image:
            continue

        normalized_image = os.path.normcase(os.path.normpath(image))
        if normalized_image == normalized_target:
            return True
    return False


def _launch_with_shell_runas(target_path: str) -> bool:
    """功能：透過 ShellExecute 觸發 UAC 提升後啟動 ShareMouse；單位：布林值"""
    directory = os.path.dirname(target_path) or None
    result = shell32.ShellExecuteW(None, "runas", target_path, None, directory, SW_SHOWNORMAL)
    if result <= 32:
        print(f"[ERROR] ShellExecuteW runas 失敗，回傳值: {result}")
        return False
    print("[INFO] 已透過 UAC 提示要求提升權限後啟動 ShareMouse")
    return True


def _launch_target(target_path):
    """功能：啟動 ShareMouse；單位：布林值"""
    if not os.path.exists(target_path):
        print(f"[WARN] 找不到 ShareMouse 執行檔：{target_path}")
        return False

    try:
        subprocess.Popen(
            [target_path],
            creationflags=DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=False,
        )
        print(f"[INFO] 已啟動 ShareMouse：{target_path}")
        return True
    except OSError as exc:
        if getattr(exc, "winerror", None) == ERROR_ELEVATION_REQUIRED:
            print("[WARN] ShareMouse 需要提升權限，嘗試透過 UAC 重新啟動")
            return _launch_with_shell_runas(target_path)

        print(f"[ERROR] ShareMouse 啟動失敗：{exc}")
        return False


def main():
    """功能：每秒巡檢 ShareMouse，閒置 5 分鐘關閉、恢復輸入再啟動；單位：無"""
    normalized_target = os.path.normpath(TARGET_EXE)
    print(
        f"[INFO] ShareMouse Watchdog 啟動，目標：{normalized_target}，"
        f"閒置門檻：{IDLE_TIMEOUT_SEC:.0f}s"
    )

    idle_mode = False
    while True:
        idle = _idle_seconds()

        if idle >= IDLE_TIMEOUT_SEC:
            if not idle_mode:
                print(
                    f"[INFO] 偵測到閒置 {idle:.0f}s 超過 {IDLE_TIMEOUT_SEC:.0f}s，"
                    f"暫停 Watchdog 並關閉 ShareMouse ({time.strftime('%H:%M:%S')})"
                )
                if _is_target_alive(normalized_target):
                    _terminate_target(normalized_target)
                idle_mode = True
            time.sleep(POLL_INTERVAL_SEC)
            continue

        if idle_mode:
            print(
                f"[INFO] 偵測到輸入恢復（閒置 {idle:.1f}s），"
                f"恢復 Watchdog 並準備啟動 ShareMouse ({time.strftime('%H:%M:%S')})"
            )
            idle_mode = False

        alive = _is_target_alive(normalized_target)
        if not alive:
            print(f"[WARN] 偵測到 ShareMouse 未執行，準備重啟 ({time.strftime('%H:%M:%S')})")
            if _launch_target(normalized_target):
                time.sleep(RESTART_GRACE_SEC)
        time.sleep(POLL_INTERVAL_SEC)


def _ensure_background_watchdog() -> bool:
    """功能：確保背景 watchdog 執行緒啟動；單位：布林值"""
    global _watchdog_thread
    if _watchdog_thread and _watchdog_thread.is_alive():
        return False

    thread = threading.Thread(target=main, name="ShareMouseWatchdog", daemon=True)
    thread.start()
    _watchdog_thread = thread
    print("[INFO] ShareMouse Watchdog 背景執行緒啟動")
    return True


if __name__ == "__main__":
    main()
else:
    _ensure_background_watchdog()
