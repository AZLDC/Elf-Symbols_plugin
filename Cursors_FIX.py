import time
import ctypes
from ctypes import wintypes

SPI_SETCURSORS = 0x0057
CURSOR_SHOWING = 0x00000001
IDC_ARROW = 32512

# 所有系統游標類型 ID（供 SetSystemCursor 使用）
ALL_OCR_IDS = (
    32512,  # OCR_NORMAL
    32513,  # OCR_IBEAM
    32514,  # OCR_WAIT
    32515,  # OCR_CROSS
    32516,  # OCR_UP
    32642,  # OCR_SIZENWSE
    32643,  # OCR_SIZENESW
    32644,  # OCR_SIZEWE
    32645,  # OCR_SIZENS
    32646,  # OCR_SIZEALL
    32648,  # OCR_NO
    32649,  # OCR_HAND
    32650,  # OCR_APPSTARTING
)

HCURSOR = wintypes.HANDLE

class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hCursor", HCURSOR),
        ("ptScreenPos", wintypes.POINT),
    ]

user32 = ctypes.WinDLL('user32', use_last_error=True)

# 確保 handle 相關函式回傳完整指標寬度
user32.LoadCursorW.restype = wintypes.HANDLE
user32.CopyIcon.restype = wintypes.HANDLE
user32.SetSystemCursor.argtypes = [wintypes.HANDLE, wintypes.DWORD]
user32.SetSystemCursor.restype = wintypes.BOOL

def get_cursor_handle():
    """功能：取得目前游標控制碼；單位：無"""
    ci = CURSORINFO()
    ci.cbSize = ctypes.sizeof(CURSORINFO)

    if not user32.GetCursorInfo(ctypes.byref(ci)):
        return None

    if not (ci.flags & CURSOR_SHOWING):
        return None

    return int(ctypes.cast(ci.hCursor, ctypes.c_void_p).value or 0)

def force_reload_cursors():
    """功能：強制重置動畫游標（Win10 1809 凍幀修正）；單位：無

    原理：先用靜態游標取代所有系統游標類型，觸發動畫引擎重置，
    再從登錄檔重載游標方案以恢復正確的動畫游標。
    """
    # 取得內建靜態箭頭游標
    arrow = user32.LoadCursorW(0, IDC_ARROW)
    if not arrow:
        print(f"[DEBUG] LoadCursorW 失敗，錯誤碼: {ctypes.get_last_error()}")
        return False

    # 用靜態游標暫時取代所有系統游標類型，重置動畫引擎
    for ocr_id in ALL_OCR_IDS:
        copy = user32.CopyIcon(arrow)
        if copy:
            user32.SetSystemCursor(copy, ocr_id)

    # 從登錄檔重載游標方案，恢復動畫游標
    result = user32.SystemParametersInfoW(SPI_SETCURSORS, 0, None, 0)
    if not result:
        err = ctypes.get_last_error()
        print(f"[DEBUG] SystemParametersInfoW 失敗，錯誤碼: {err}")
        return False

    # print(f"[DEBUG] 游標重置完成（靜態取代 → 方案重載）")
    return True

def main():
    poll_interval_sec = 0.02  # 功能：輪詢間隔；單位：秒
    cooldown_sec = 0.10       # 功能：去抖動時間；單位：秒

    last_h = None
    cooldown_until = 0.0
    print("[DEBUG] 游標監控已啟動")

    while True:
        time.sleep(poll_interval_sec)
        now = time.monotonic()
        h = get_cursor_handle()

        if not h:
            continue

        # 首次取得有效控制碼，僅記錄不觸發重載
        if last_h is None:
            last_h = h
            print(f"[DEBUG] 初始游標控制碼: {h:#010x}")
            continue

        # 偵測到游標切換
        if h != last_h:
            if now >= cooldown_until:
                print(f"[DEBUG] {time.strftime('%H:%M:%S')} 游標切換: {last_h:#010x} -> {h:#010x}，重載游標方案")
                force_reload_cursors()
                cooldown_until = now + cooldown_sec
            else:
                print(f"[DEBUG] {time.strftime('%H:%M:%S')} 游標切換: {last_h:#010x} -> {h:#010x}（冷卻中，略過重載）")

        # 無論是否冷卻，都追蹤最新控制碼，避免冷卻結束後誤觸發
        last_h = h

if __name__ == "__main__":
    main()
