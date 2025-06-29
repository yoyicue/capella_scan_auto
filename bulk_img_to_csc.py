#!/usr/bin/env python3
"""bulk_img_to_csc.py

批量把 ./img_in/*.png → ./csc_out/*.csc
目标应用：C:\\Program Files (x86)\\capella-software\\capella-scan 9\\bin\\capscan.exe
Qt 5.15.2 / pywinauto 0.6.8 适配
"""
from __future__ import annotations

from pathlib import Path
from time import sleep
import ctypes
import sys
import subprocess

# 3rd-party
from pywinauto import Application, Desktop
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

# 全局等待时间配置（秒）
WAIT_SHORT = 0.5    # 短等待：UI 响应、焦点切换
WAIT_MEDIUM = 1.0   # 中等待：对话框出现、文件加载
WAIT_LONG = 2.0     # 长等待：目录导航、程序启动
WAIT_PROCESS = 5.0  # 进程等待：程序启动、识别完成

# -----------------------------------------------------------------------------
# 帮助函数
# -----------------------------------------------------------------------------

def connect_or_start() -> Application:
    """启动或连接单实例的 capscan.exe"""
    # 直接启动新实例（因为在调用前已经清理了旧进程）
    print("[INFO] 启动新的 capscan.exe 实例...")
    try:
        exe_path = Path(CAPSCAN_EXE)
        exe_dir = exe_path.parent
        subprocess.Popen(str(exe_path), cwd=str(exe_dir))
        # 等待程序初始化
        sleep(WAIT_PROCESS)
        app = Application(backend="uia").connect(path=CAPSCAN_EXE, timeout=20)
        print(f"[INFO] 已启动并连接到新实例 (PID: {app.process})。")
        return app
    except Exception as e:
        sys.exit(f"[ERR] 启动并连接 capscan.exe 失败: {e}")


def is_file_dialog(win, dialog_type: str) -> bool:
    """通用函数，判断窗口是否为指定类型的文件对话框 (open/save)"""
    try:
        # 1. 检查是否为 Windows 通用对话框
        if win.class_name() != "#32770":
            return False
        # 2. 检查标题是否匹配
        if dialog_type.lower() not in win.window_text().lower():
            return False
        # 3. 检查是否有对应的文件名编辑框
        if dialog_type == 'open' and win.child_window(auto_id="1148", control_type="Edit").exists():
            return True
        if dialog_type == 'save' and win.child_window(auto_id="1001", control_type="Edit").exists():
            return True
    except Exception:
        return False
    return False


def wait_for_state(app: Application, state: str, timeout: int = 20) -> 'WindowSpecification':
    """
    轮询等待应用进入指定状态并返回该状态的窗口。
    状态: 'main', 'open', 'save'
    """
    print(f"[STATE] 等待 '{state}' 窗口变为活动状态...")
    for _ in range(timeout * 10):
        # 获取 capscan 进程的所有窗口
        try:
            windows = app.windows()
            for win in windows:
                # 检查窗口是否可见
                if not win.is_visible():
                    continue
                    
                title = win.window_text()
                class_name = win.class_name()
                
                # 根据 state 判断窗口类型
                if state == 'main':
                    # 主窗口：标题为 "capella-scan 9" 或包含文件名，类名为 "MainWindow"
                    if (title == 'capella-scan 9' or 
                        'capella-scan' in title.lower() or 
                        '.png' in title.lower() or 
                        '.csc' in title.lower()) and class_name == 'MainWindow':
                        print(f"[STATE] 已检测到 'main' 窗口: '{title}'")
                        return win
                elif state == 'open' or state == 'save':
                    # 检查当前窗口是否为对话框
                    if is_file_dialog(win, state):
                        print(f"[STATE] 已检测到 '{state}' 窗口: '{title}'")
                        return win
                    
                    # 检查子窗口中是否有对话框
                    try:
                        children = win.children()
                        for child in children:
                            child_title = child.window_text()
                            child_class = child.class_name()
                            
                            # 查找 Open File 或 Save 对话框
                            if (child_class == "#32770" and 
                                ((state == 'open' and 'open' in child_title.lower()) or
                                 (state == 'save' and 'save' in child_title.lower()))):
                                print(f"[STATE] 已检测到 '{state}' 子窗口: '{child_title}'")
                                return child
                    except Exception:
                        pass
        except Exception as e:
            print(f"[DEBUG] 窗口检测异常: {e}")
            pass

        sleep(WAIT_SHORT / 5)  # 更频繁的检测
    raise TimeoutError(f"等待 '{state}' 状态超时（{timeout}秒）")


def wait_recognition_finished(main_win, timeout: int = 120) -> bool:
    """轮询"开始识别"按钮重新可用 → 识别完成"""
    for _ in range(timeout):
        try:
            # 使用 descendants 查找按钮
            buttons = main_win.descendants(control_type="Button")
            for btn in buttons:
                try:
                    if hasattr(btn, 'automation_id') and btn.automation_id == START_BTN_ID:
                        if btn.exists() and btn.is_enabled():
                            return True
                        break
                except:
                    continue
        except Exception:
            pass
        sleep(WAIT_MEDIUM)
    return False

# -----------------------------------------------------------------------------
# 主逻辑
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    print("[INFO] 准备环境：强制关闭所有旧的 capscan.exe 实例...")
    subprocess.run("taskkill /F /IM capscan.exe", capture_output=True, check=False)
    # 等待进程完全终止 - 增加等待时间并验证
    for _ in range(3):  # 最多等待3秒
        result = subprocess.run("tasklist /FI \"IMAGENAME eq capscan.exe\"", 
                              capture_output=True, text=True, check=False)
        if "capscan.exe" not in result.stdout:
            print("[INFO] 旧进程已完全清理。")
            break
        sleep(WAIT_MEDIUM)
    else:
        print("[WARN] 旧进程可能未完全清理，继续执行...")

    if not INPUT_DIR.exists():
        sys.exit(f"[ERR] 输入目录不存在: {INPUT_DIR}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    app = connect_or_start()
    
    try:
        main = wait_for_state(app, 'main')
    except TimeoutError as e:
        sys.exit(f"[ERR] 启动后无法定位主窗口: {e}")

    png_files = sorted(INPUT_DIR.glob("*.png"))
    if not png_files:
        sys.exit(f"[INFO] {INPUT_DIR} 下未找到 *.png 文件")

    for img_path in png_files:
        try:
            # 确保当前是主窗口状态
            main = wait_for_state(app, 'main')

            # Step 3: 触发打开文件
            send_keys("^o")

            # Step 4: 等待并操作 Open 对话框
            open_dlg = wait_for_state(app, 'open')
            # 查找文件名编辑框并填写路径
            try:
                # 方法1: 先导航到目录，再选择文件
                # 1.1 在地址栏输入目录路径
                send_keys("^l")  # Ctrl+L 聚焦地址栏
                sleep(WAIT_SHORT)
                send_keys(str(img_path.parent), with_spaces=True)  # 输入目录路径（不加引号）
                send_keys("{ENTER}")  # 确认导航
                sleep(WAIT_LONG)  # 等待目录加载
                
                # 1.2 验证是否成功跳转到目标目录
                # 通过检查当前路径或直接尝试输入文件名
                
                # 1.3 在文件列表中选择文件
                # 方式1: 直接输入文件名首字母快速定位
                print(f"[INFO] 寻找文件: {img_path.name}")
                send_keys(img_path.name[0])  # 输入文件名首字母，快速定位
                sleep(WAIT_SHORT)
                
                # 方式2: 使用 Ctrl+F 搜索文件
                # send_keys("^f")  # Ctrl+F 打开搜索
                # sleep(WAIT_SHORT)
                # send_keys(img_path.name, with_spaces=True)
                # send_keys("{ENTER}")
                
                # 方式3: 直接在文件名框输入完整文件名
                send_keys("{F4}")  # F4 通常会定位到文件名框
                sleep(WAIT_SHORT)
                send_keys("^a")  # 全选
                send_keys(img_path.name, with_spaces=True)
                
                print(f"[INFO] 已定位文件: {img_path.name}")
            except Exception as e:
                print(f"[WARN] 填写文件路径失败: {e}")
                # 兜底方案: 尝试直接输入文件名
                send_keys("{F4}")
                sleep(WAIT_SHORT)
                send_keys("^a")
                send_keys(img_path.name, with_spaces=True)
                
            # 确保文件被打开 - 多种尝试方式
            print(f"[INFO] 尝试打开文件...")
            # 方式1: 直接回车（如果文件已选中）
            send_keys("{ENTER}")
            sleep(WAIT_SHORT)
            
            # 方式2: 双击文件（如果文件在列表中可见）
            # send_keys("{ENTER}")  # 如果文件被选中，回车应该能打开
            
            # 方式3: 点击"打开"按钮
            # 可以尝试 Alt+O 快捷键
            send_keys("%o")  # Alt+O，通常是"打开"按钮的快捷键
            sleep(WAIT_MEDIUM)
            
            # Step 5 & 6: 等待主窗口恢复并启动识别
            main = wait_for_state(app, 'main') # 等待图像加载完毕
            # 使用 descendants 查找开始识别按钮
            try:
                buttons = main.descendants(control_type="Button")
                start_btn = None
                for btn in buttons:
                    try:
                        if hasattr(btn, 'automation_id') and btn.automation_id == START_BTN_ID:
                            start_btn = btn
                            break
                    except:
                        continue
                
                if start_btn:
                    start_btn.click_input()
                else:
                    # 兜底：使用快捷键
                    send_keys("{F5}")
            except Exception as e:
                print(f"[WARN] 点击开始识别按钮失败: {e}")
                send_keys("{F5}")  # 兜底
                
            if not wait_recognition_finished(main):
                print(f"[WARN] 识别 {img_path.name} 超时，跳过")
                send_keys("^w")  # 关闭标签
                continue

            # Step 7: 导出 csc
            out_file = OUTPUT_DIR / f"{img_path.stem}.csc"
            send_keys("+^m")  # Shift+Ctrl+M
            save_dlg = wait_for_state(app, 'save')
            # 查找保存文件名编辑框并填写路径
            try:
                # 导航到输出目录并输入文件名
                send_keys("^l")  # 聚焦地址栏
                sleep(WAIT_SHORT)
                send_keys(str(out_file.parent), with_spaces=True)  # 输入目录路径（不加引号）
                send_keys("{ENTER}")
                sleep(WAIT_LONG)
                
                # 输入文件名
                send_keys("{TAB}")  # Tab 切换到文件名框
                sleep(WAIT_SHORT)
                send_keys("^a")  # 全选当前内容
                send_keys(out_file.name, with_spaces=True)  # 只输入文件名
                print(f"[INFO] 已输入保存文件名: {out_file.name}")
                     
                # 查找保存按钮并点击
                print(f"[INFO] 尝试保存文件...")
                send_keys("{ENTER}")  # 直接回车保存
                sleep(WAIT_MEDIUM)
                     
            except Exception as e:
                print(f"[WARN] 保存文件失败: {e}")
                send_keys("{ENTER}")  # 兜底

            # Step 8: 收尾
            main = wait_for_state(app, 'main') # 等待保存完成，焦点回到主窗口
            send_keys("^w")
            print(f"[OK] {img_path.name} → {out_file.name}")

        except Exception as exc:
            print(f"[ERR] 处理 {img_path.name} 失败: {exc}")
            # 尝试回到主窗口，然后关闭标签页
            try:
                main = wait_for_state(app, 'main', timeout=5)
                main.send_keys("^w")
            except TimeoutError:
                pass  # 如果连主窗口都找不到，就放弃

    print("[DONE] 全部转换完成") 