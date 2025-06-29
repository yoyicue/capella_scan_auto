#!/usr/bin/env python3
"""bulk_img_to_csc.py

批量把 ./img_in/*.png → ./csc_out/*.csc
目标应用：C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe
Qt 5.15.2 / pywinauto 0.6.8 适配
"""
from __future__ import annotations

from pathlib import Path
from time import sleep
import ctypes
import sys

# 3rd-party
from pywinauto import Application
from pywinauto.keyboard import send_keys

try:
    import win32con  # type: ignore
    import win32gui  # type: ignore
except ImportError as exc:
    sys.exit("[ERR] 需要先安装 pywin32： pip install pywin32 — " + str(exc))

# -----------------------------------------------------------------------------
# 环境准备：DPI 感知确保坐标点击在高 DPI 下正确
# -----------------------------------------------------------------------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor v2
except Exception:
    pass  # 旧系统可忽略
try:
    _hdc = win32gui.GetDC(0)
    SCALING = win32gui.GetDeviceCaps(_hdc, win32con.LOGPIXELSX) / 96.0
    win32gui.ReleaseDC(0, _hdc)
except Exception:
    SCALING = 1.0  # 兜底，不影响 UIA 操作

# -----------------------------------------------------------------------------
# 路径 & 常量
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "img_in"
OUTPUT_DIR = BASE_DIR / "csc_out"
CAPSCAN_EXE = r"C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe"

# Qt QAction objectName → AutomationId（Qt 5.15.2）
START_BTN_ID = "actionStartRecognition"
SAVE_LEVEL_ID = "actionSave_Level_of_Recognition"

# -----------------------------------------------------------------------------
# 帮助函数
# -----------------------------------------------------------------------------

def connect_or_start() -> Application:
    """启动或连接单实例的 capscan.exe"""
    try:
        return Application(backend="uia").connect(path=CAPSCAN_EXE, timeout=4)
    except Exception:
        return Application(backend="uia").start(CAPSCAN_EXE, timeout=30)


def wait_recognition_finished(main_win, timeout: int = 120) -> bool:
    """轮询"开始识别"按钮重新可用 → 识别完成"""
    btn = main_win.child_window(auto_id=START_BTN_ID, control_type="Button")
    for _ in range(timeout):
        if btn.exists() and btn.is_enabled():
            return True
        sleep(1)
    return False

# -----------------------------------------------------------------------------
# 主逻辑
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    if not INPUT_DIR.exists():
        sys.exit(f"[ERR] 输入目录不存在: {INPUT_DIR}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    app = connect_or_start()
    main = app.window(title_re=".*capella-scan.*")
    main.wait("visible", 20)

    png_files = sorted(INPUT_DIR.glob("*.png"))
    if not png_files:
        sys.exit(f"[INFO] {INPUT_DIR} 下未找到 *.png 文件")

    for img_path in png_files:
        try:
            # Step 3 打开文件对话框
            send_keys("^o")  # Ctrl+O
            open_dlg = app.window(class_name="#32770")  # Windows CommonItemDialog
            open_dlg.wait("visible", 10)

            # Step 4 输入路径 & Open
            open_dlg.child_window(auto_id="1148", control_type="Edit").set_edit_text(str(img_path))
            open_dlg.child_window(title="Open", control_type="Button").click()

            # Step 5 等图像加载
            sleep(1)

            # Step 6 启动识别
            main.child_window(auto_id=START_BTN_ID, control_type="Button").click_input()
            if not wait_recognition_finished(main):
                print(f"[WARN] 识别 {img_path.name} 超时，跳过")
                send_keys("^w")  # 关闭标签
                continue

            # Step 7 导出 csc
            out_file = OUTPUT_DIR / f"{img_path.stem}.csc"
            send_keys("+^m")  # Shift+Ctrl+M
            save_dlg = app.window(class_name="#32770")
            save_dlg.wait("visible", 10)
            save_dlg.child_window(auto_id="1001", control_type="Edit").set_edit_text(str(out_file))
            save_dlg.child_window(title="Save", control_type="Button").click()

            # Step 8 收尾
            send_keys("^w")
            print(f"[OK] {img_path.name} → {out_file.name}")

        except Exception as exc:
            print(f"[ERR] 处理 {img_path.name} 失败: {exc}")
            send_keys("^w")

    print("[DONE] 全部转换完成") 