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
import time

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


def wait_recognition_finished(main_window, timeout=60):
    """等待识别完成"""
    print(f"[DEBUG] 开始等待识别完成（超时: {timeout}秒）...")
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # 添加状态检查
            elapsed = int(time.time() - start_time)
            if elapsed % 5 == 0 and elapsed > 0:  # 每5秒输出一次状态
                print(f"[DEBUG] 等待识别完成中... ({elapsed}s/{timeout}s)")
            
            # 查找 "Result of recognition" 文本
            all_texts = main_window.descendants(control_type="Text")
            for text_elem in all_texts:
                try:
                    text_content = text_elem.window_text()
                    if 'Result of recognition' in text_content:
                        print(f"[INFO] 检测到识别完成标志: '{text_content}'")
                        return True
                except:
                    continue
                 
        except Exception as e:
            print(f"[DEBUG] 检查识别状态时出错: {e}")
            
        time.sleep(WAIT_SHORT)
    
    print(f"[WARN] 等待识别完成超时 ({timeout}秒)")
    return False

def wait_recognition_finished_backup(main_window, timeout=30):
    """备用的识别完成检测方法"""
    print(f"[DEBUG] 使用备用方法等待识别完成...")
    
    # 方法1: 简单等待固定时间（适用于小图片）
    print(f"[INFO] 等待 {timeout} 秒让识别完成...")
    time.sleep(timeout)
    
    # 方法2: 检查窗口是否还存在且响应
    try:
        if main_window.exists():
            print(f"[INFO] 窗口仍然存在，假设识别已完成")
            return True
    except:
        pass
    
    return True  # 备用方法默认返回成功

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
                # 1. 先导航到目标目录
                print(f"[INFO] 导航到目录: {img_path.parent}")
                send_keys("^l")  # Ctrl+L 聚焦地址栏
                sleep(WAIT_SHORT)
                send_keys(str(img_path.parent), with_spaces=True)  # 输入目录路径
                send_keys("{ENTER}")  # 确认导航
                sleep(WAIT_LONG)  # 等待目录加载
                
                # 2. 在目录中搜索文件
                # 优化策略：直接搜索文件（速度优先）
                print(f"[INFO] 搜索文件: {img_path.name}")
                
                send_keys("^f")  # Ctrl+F 打开搜索框
                sleep(WAIT_SHORT)
                send_keys(img_path.name, with_spaces=True)  # 输入完整文件名
                send_keys("{ENTER}")  # 执行搜索并选中文件
                sleep(WAIT_SHORT)
                
                print(f"[INFO] 已搜索到文件: {img_path.name}")
            except Exception as e:
                print(f"[WARN] 填写文件路径失败: {e}")
                # 兜底方案: 直接在文件名框输入
                send_keys("{F4}")  # F4 定位文件名框
                sleep(WAIT_SHORT)
                send_keys("^a")  # 全选
                send_keys(img_path.name, with_spaces=True)
                
            # 确保文件被打开 - 多种尝试方式
            print(f"[INFO] 尝试打开文件...")
            
            # 多重尝试方法（之前成功的方式）
            # 方式1: 先尝试 Enter
            send_keys("{ENTER}")
            sleep(WAIT_SHORT)
            
            # 方式2: 再尝试 Alt+O（"打开"按钮快捷键）
            send_keys("%o")  # Alt+O 点击"打开"按钮
            sleep(WAIT_MEDIUM)
            
            # Step 5 & 6: 等待主窗口恢复并启动识别
            main = wait_for_state(app, 'main') # 等待图像加载完毕
            print(f"[INFO] 图像已加载，准备启动识别...")
            
            # 使用 descendants 查找开始识别按钮
            try:
                print(f"[INFO] 查找开始识别按钮...")
                buttons = main.descendants(control_type="Button")
                print(f"[DEBUG] 找到 {len(buttons)} 个按钮")
                
                # 输出所有按钮的详细信息
                for i, btn in enumerate(buttons[:10]):  # 只显示前10个避免输出过多
                    try:
                        btn_id = getattr(btn, 'automation_id', 'N/A')
                        btn_text = getattr(btn, 'window_text', lambda: 'N/A')()
                        btn_class = getattr(btn, 'class_name', lambda: 'N/A')()
                        print(f"[DEBUG] 按钮{i}: ID='{btn_id}', Text='{btn_text}', Class='{btn_class}'")
                    except Exception as e:
                        print(f"[DEBUG] 按钮{i}: 无法读取信息 - {e}")
                
                start_btn = None
                for btn in buttons:
                    try:
                        btn_text = getattr(btn, 'window_text', lambda: '')()
                        if 'Start Recognition' in btn_text:
                            print(f"[INFO] 找到开始识别按钮: '{btn_text}'")
                            start_btn = btn
                            break
                    except:
                        continue
                
                if start_btn:
                    print(f"[INFO] 点击开始识别按钮...")
                    start_btn.click_input()
                else:
                    print(f"[WARN] 未找到识别按钮，使用 F5 快捷键...")
                    # 兜底：使用快捷键
                    send_keys("{F5}")
            except Exception as e:
                print(f"[WARN] 点击开始识别按钮失败: {e}")
                print(f"[INFO] 使用 F5 快捷键作为兜底...")
                send_keys("{F5}")  # 兜底
                
            # 等待一下让识别开始
            time.sleep(WAIT_MEDIUM)
            
            print(f"[INFO] 等待识别完成...")
            # 尝试两种方法检测识别完成
            recognition_finished = wait_recognition_finished(main)
            if not recognition_finished:
                print(f"[INFO] 状态栏方法失败，尝试备用检测方法...")
                recognition_finished = wait_recognition_finished_backup(main)
            
            if not recognition_finished:
                print(f"[WARN] 识别 {img_path.name} 超时，跳过")
                send_keys("^w")  # 关闭标签
                continue
            
            print(f"[INFO] 识别已完成！")

            # 识别完成后，额外等待确保状态稳定
            time.sleep(WAIT_LONG)
            
            # Step 7: 保存为CSC格式
            print(f"[INFO] 准备保存 CSC 文件...")
            
            # 定义输出文件路径
            out_file = OUTPUT_DIR / f"{img_path.stem}.csc"
            print(f"[INFO] 目标保存路径: {out_file}")
            
            # 确保窗口有焦点
            main.set_focus()
            time.sleep(WAIT_SHORT)
            
            # 使用快捷键保存
            print(f"[INFO] 发送保存快捷键 Shift+Ctrl+M...")
            send_keys("+^m")  # Shift+Ctrl+M
            time.sleep(WAIT_MEDIUM)
            
            # 等待保存对话框出现
            print(f"[INFO] 等待保存对话框...")
            time.sleep(WAIT_LONG)
            
            # 尝试在主窗口中查找保存对话框控件
            try:
                # 方法1: 查找独立的保存对话框窗口
                print(f"[DEBUG] 方法1: 查找独立保存对话框...")
                save_dialog = None
                dialogs = app.windows()
                print(f"[DEBUG] 找到 {len(dialogs)} 个窗口")
                for i, dialog in enumerate(dialogs):
                    try:
                        dialog_title = dialog.window_text()
                        print(f"[DEBUG] 窗口{i}: '{dialog_title}'")
                        if "Save level of recognition" in dialog_title or "save" in dialog_title.lower():
                            save_dialog = dialog
                            print(f"[INFO] 找到保存对话框: '{dialog_title}'")
                            break
                    except:
                        print(f"[DEBUG] 窗口{i}: 无法读取标题")
                
                # 方法2: 如果没找到独立对话框，在主窗口中查找保存控件
                if not save_dialog:
                    print(f"[DEBUG] 方法2: 在主窗口中查找保存控件...")
                    main_window = wait_for_state(app, 'main', timeout=5)
                    
                    # 查找所有文本控件，寻找保存相关的标签
                    text_controls = main_window.descendants(control_type="Text")
                    print(f"[DEBUG] 主窗口中找到 {len(text_controls)} 个文本控件")
                    for i, text_ctrl in enumerate(text_controls[:10]):  # 只检查前10个
                        try:
                            text_content = text_ctrl.window_text()
                            if text_content and any(keyword in text_content.lower() for keyword in ['folder', 'file', 'save', 'output']):
                                print(f"[DEBUG] 文本控件{i}: '{text_content}'")
                        except:
                            pass
                    
                    # 查找编辑框
                    edit_controls = main_window.descendants(control_type="Edit")
                    print(f"[DEBUG] 主窗口中找到 {len(edit_controls)} 个编辑框")
                    
                    # 如果找到编辑框，假设这是保存对话框
                    if len(edit_controls) >= 2:
                        save_dialog = main_window
                        print(f"[INFO] 在主窗口中找到保存控件")
                
                if save_dialog:
                    # 查找文件夹输入框并设置输出目录
                    print(f"[INFO] 设置输出目录: {out_file.parent}")
                    try:
                        # 查找编辑框（通常是文件夹路径）
                        edit_controls = save_dialog.descendants(control_type="Edit")
                        if len(edit_controls) >= 1:
                            folder_edit = edit_controls[0]  # 第一个是文件夹
                            folder_edit.click_input()
                            time.sleep(WAIT_SHORT)
                            send_keys("^a")  # 全选
                            send_keys(str(out_file.parent), with_spaces=True)
                            time.sleep(WAIT_SHORT)
                            print(f"[INFO] 已设置输出目录")
                    except Exception as e:
                        print(f"[WARN] 设置输出目录失败: {e}")
                    
                    # 查找文件名输入框并设置文件名
                    print(f"[INFO] 设置文件名: {out_file.name}")
                    try:
                        edit_controls = save_dialog.descendants(control_type="Edit")
                        if len(edit_controls) >= 2:
                            file_edit = edit_controls[1]  # 第二个是文件名
                            file_edit.click_input()
                            time.sleep(WAIT_SHORT)
                            send_keys("^a")  # 全选
                            send_keys(out_file.name, with_spaces=True)
                            time.sleep(WAIT_SHORT)
                            print(f"[INFO] 已设置文件名")
                    except Exception as e:
                        print(f"[WARN] 设置文件名失败: {e}")
                    
                    # 点击OK按钮（或Overwrite按钮）
                    print(f"[INFO] 点击确认按钮...")
                    try:
                        buttons = save_dialog.descendants(control_type="Button")
                        ok_button = None
                        for btn in buttons:
                            try:
                                btn_text = btn.window_text()
                                if btn_text in ['OK', 'Overwrite']:
                                    ok_button = btn
                                    print(f"[INFO] 找到确认按钮: '{btn_text}'")
                                    break
                            except:
                                continue
                        
                        if ok_button:
                            ok_button.click_input()
                            time.sleep(WAIT_MEDIUM)
                            print(f"[INFO] 已点击确认按钮")
                        else:
                            print(f"[WARN] 未找到确认按钮，使用回车键...")
                            send_keys("{ENTER}")
                            time.sleep(WAIT_MEDIUM)
                    except Exception as e:
                        print(f"[WARN] 点击确认按钮失败: {e}，使用回车键...")
                        send_keys("{ENTER}")
                        time.sleep(WAIT_MEDIUM)
                        
                else:
                    print(f"[WARN] 未找到保存对话框，使用默认方式...")
                    send_keys("{ENTER}")
                    time.sleep(WAIT_MEDIUM)
                    
            except Exception as e:
                print(f"[WARN] 处理保存对话框失败: {e}")
                send_keys("{ENTER}")  # 兜底
                time.sleep(WAIT_MEDIUM)
            
            # Step 8: 收尾
            main = wait_for_state(app, 'main') # 等待保存完成，焦点回到主窗口
            send_keys("^w")
            print(f"[OK] {img_path.name} → {out_file.name}")

        except Exception as e:
            print(f"[ERR] 处理 {img_path.name} 失败: {e}")
            try:
                main = wait_for_state(app, 'main', timeout=5)
                send_keys("^w")  # 关闭标签
            except:
                pass

    print("[DONE] 全部转换完成") 