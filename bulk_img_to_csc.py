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

# 全局等待时间配置（秒）- 优化后的配置
WAIT_SHORT = 0.3    # 短等待：UI 响应、焦点切换 (从0.5秒优化为0.3秒)
WAIT_MEDIUM = 0.8   # 中等待：对话框出现、文件加载 (从1.0秒优化为0.8秒)
WAIT_LONG = 1.5     # 长等待：目录导航、程序启动 (从2.0秒优化为1.5秒)
WAIT_PROCESS = 5.0  # 进程等待：程序启动、识别完成

# 全局状态跟踪
_dialog_directory_set = False  # FileDialog 是否已设置正确目录
_start_time = time.time()  # 程序启动时间基准

def tprint(msg: str, level: str = "INFO") -> None:
    """带时间戳的打印函数"""
    elapsed = time.time() - _start_time
    print(f"[{level}] [{elapsed:6.1f}s] {msg}")

# -----------------------------------------------------------------------------
# 帮助函数
# -----------------------------------------------------------------------------

def wait_until(predicate, timeout: float = 10.0, interval: float = 0.2, desc: str = "condition") -> bool:
    """轮询 predicate 直到返回 True 或超时。

    Args:
        predicate: 可调用对象，返回 True 表示条件满足。
        timeout: 最大等待时间（秒）。
        interval: 轮询间隔（秒）。
        desc: 日志描述。

    Returns:
        bool: 是否在超时前满足条件。
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            if predicate():
                return True
        except Exception:
            pass
        sleep(interval)
    tprint(f"等待 {desc} 超时({timeout}s)", "WARN")
    return False

def connect_or_start() -> Application:
    """启动或连接单实例的 capscan.exe"""
    # 直接启动新实例（因为在调用前已经清理了旧进程）
    tprint("启动新的 capscan.exe 实例...")
    try:
        exe_path = Path(CAPSCAN_EXE)
        exe_dir = exe_path.parent
        subprocess.Popen(str(exe_path), cwd=str(exe_dir))

        # 轮询等待进程就绪，而非固定休眠
        def _can_connect():
            try:
                Application(backend="uia").connect(path=CAPSCAN_EXE, timeout=1)
                return True
            except Exception:
                return False

        if not wait_until(_can_connect, timeout=WAIT_PROCESS * 2, interval=0.5, desc="capscan.exe 启动"):
            raise RuntimeError("capscan.exe 启动超时")

        app = Application(backend="uia").connect(path=CAPSCAN_EXE, timeout=5)
        tprint(f"已启动并连接到新实例 (PID: {app.process})")
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
    tprint(f"等待 '{state}' 窗口变为活动状态...")
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
                        tprint(f"已检测到 'main' 窗口: '{title}'")
                        return win
                elif state == 'open' or state == 'save':
                    # 检查当前窗口是否为对话框
                    if is_file_dialog(win, state):
                        tprint(f"已检测到 '{state}' 窗口: '{title}'")
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
                                tprint(f"已检测到 '{state}' 子窗口: '{child_title}'")
                                return child
                    except Exception:
                        pass
        except Exception as e:
            tprint(f"窗口检测异常: {e}", "DEBUG")
            pass

        sleep(0.2)  # 优化：统一轮询间隔为0.2秒
    raise TimeoutError(f"等待 '{state}' 状态超时（{timeout}秒）")


def wait_recognition_finished(main_window, timeout=60):
    """等待识别完成"""
    tprint(f"开始等待识别完成（超时: {timeout}秒）...")
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # 添加状态检查
            elapsed = int(time.time() - start_time)
            if elapsed % 5 == 0 and elapsed > 0:  # 每5秒输出一次状态
                tprint(f"等待识别完成中... ({elapsed}s/{timeout}s)", "DEBUG")
            
            # 查找 "Result of recognition" 文本
            all_texts = main_window.descendants(control_type="Text")
            for text_elem in all_texts:
                try:
                    text_content = text_elem.window_text()
                    if 'Result of recognition' in text_content:
                        tprint(f"检测到识别完成标志: '{text_content}'")
                        return True
                except:
                    continue
                 
        except Exception as e:
            tprint(f"检查识别状态时出错: {e}", "DEBUG")
            
        time.sleep(WAIT_SHORT)
    
    tprint(f"等待识别完成超时 ({timeout}秒)", "WARN")
    return False

def wait_recognition_finished_backup(main_window, timeout=30):
    """备用的识别完成检测方法 - 优化版本"""
    tprint(f"使用备用方法检测识别完成...", "DEBUG")
    
    # 改为轮询检测而非固定等待
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 检查窗口状态和响应性
            if main_window.exists():
                # 尝试获取窗口状态信息
                try:
                    texts = main_window.descendants(control_type="Text")
                    # 如果能正常获取控件，认为界面稳定
                    if len(texts) > 0:
                        tprint(f"界面状态稳定，假设识别已完成")
                        return True
                except:
                    pass
        except:
            pass
        
        time.sleep(2.0)  # 每2秒检查一次
    
    tprint(f"备用检测超时，假设识别完成", "WARN")
    return True  # 备用方法默认返回成功

def wait_for_save_dialog(main_window, timeout: float = 10.0, interval: float = 0.2) -> bool:
    """等待保存对话框相关控件（Edit / OK 按钮）出现在主窗口。

    由于 Capella-scan 的保存对话框是嵌入在主窗口中的 Qt 子对话框，
    无法通过独立窗口检测，因此采用控件特征轮询。
    """
    def _has_save_controls() -> bool:
        try:
            if not main_window.exists():
                return False
            
            # 方法1：检测特定的保存按钮文本
            buttons = main_window.descendants(control_type="Button")
            for btn in buttons:
                try:
                    btn_text = btn.window_text()
                    if btn_text in ("OK", "Overwrite", "保存", "Save"):
                        tprint(f"检测到保存按钮: '{btn_text}'", "DEBUG")
                        return True
                except Exception:
                    continue
            
            # 方法2：检测编辑框数量变化（作为辅助判断）
            edits = main_window.descendants(control_type="Edit")
            if len(edits) >= 8:  # 根据日志，保存对话框出现时有8个编辑框
                tprint(f"检测到保存编辑框: {len(edits)}个", "DEBUG")
                return True
                
        except Exception:
            pass
        return False

    return wait_until(_has_save_controls, timeout=timeout, interval=interval, desc="保存对话框出现")

def try_command_line_open(img_path: Path) -> bool:
    """尝试通过命令行参数直接打开文件（最高效方式）"""
    try:
        tprint(f"尝试命令行直接打开: {img_path}")
        # 构造命令行
        cmd = f'"{CAPSCAN_EXE}" "{img_path}"'
        
        # 执行命令（非阻塞）
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            tprint("命令行打开成功")
            return True
        else:
            tprint(f"命令行打开失败: {result.stderr}", "DEBUG")
            return False
            
    except Exception as e:
        tprint(f"命令行方式异常: {e}", "DEBUG")
        return False

def smart_open_file(open_dlg, img_path: Path) -> bool:
    """智能打开文件：先检查当前目录，避免不必要的导航 - 优化版本"""
    global _dialog_directory_set
    try:
        if _dialog_directory_set:
            # 目录已设置，直接输入文件名
            tprint(f"复用已设置目录，直接选择: {img_path.name}")
            send_keys("{F4}")  # F4 定位文件名框
            sleep(0.2)  # 优化：缩短等待时间
            send_keys("^a")  # 全选
            send_keys(img_path.name, with_spaces=True)
            sleep(0.2)  # 优化：缩短等待时间
        else:
            # 首次设置目录
            tprint(f"首次设置目录: {img_path.parent}")
            send_keys("^l")  # Ctrl+L 聚焦地址栏
            sleep(0.2)  # 优化：缩短等待时间
            send_keys("^a")  # 全选地址栏
            send_keys(str(img_path.parent), with_spaces=True)
            send_keys("{ENTER}")
            sleep(1.0)  # 优化：目录加载等待时间从2秒缩短为1秒
            
            # 设置文件名
            send_keys("{F4}")  # F4 定位文件名框
            sleep(0.2)  # 优化：缩短等待时间
            send_keys("^a")
            send_keys(img_path.name, with_spaces=True)
            _dialog_directory_set = True  # 标记目录已设置
        
        # 确认打开文件
        send_keys("{ENTER}")
        sleep(0.2)  # 优化：缩短等待时间
        send_keys("%o")  # Alt+O 作为兜底
        sleep(0.6)  # 优化：缩短等待时间
        return True
        
    except Exception as e:
        tprint(f"智能文件打开失败: {e}", "WARN")
        return False

def process_single_file(app: Application, img_path: Path) -> bool:
    """处理单个文件的完整流程"""
    file_start_time = time.time()
    tprint(f"=== 开始处理文件: {img_path.name} ===")
    
    try:
        # 确保当前是主窗口状态
        main = wait_for_state(app, 'main')

        # Step 1: 尝试高效方式打开文件
        open_start_time = time.time()
        file_opened = False
        
        # 方式1: 命令行参数（最高效，但可能不支持）
        if try_command_line_open(img_path):
            file_opened = True
            tprint(f"使用命令行方式成功打开文件")
        else:
            # 方式2: UI 对话框（兜底方式）
            tprint(f"使用 UI 对话框方式打开文件")
            send_keys("^o")
            open_dlg = wait_for_state(app, 'open')
            
            if not smart_open_file(open_dlg, img_path):
                tprint(f"UI 文件选择失败，跳过该文件", "WARN")
                return False
            file_opened = True
        
        if not file_opened:
            tprint(f"所有文件打开方式均失败", "ERR")
            return False
        
        open_elapsed = time.time() - open_start_time
        tprint(f"文件打开耗时: {open_elapsed:.1f}s")
        
        # Step 2: 等待主窗口恢复并启动识别
        main = wait_for_state(app, 'main') # 等待图像加载完毕
        tprint(f"图像已加载，准备启动识别...")
        
        # 识别阶段计时
        recognition_start_time = time.time()
        
        # 使用 descendants 查找开始识别按钮
        try:
            tprint(f"查找开始识别按钮...")
            buttons = main.descendants(control_type="Button")
            tprint(f"找到 {len(buttons)} 个按钮", "DEBUG")
            
            start_btn = None
            for btn in buttons:
                try:
                    btn_text = getattr(btn, 'window_text', lambda: '')()
                    if 'Start Recognition' in btn_text:
                        tprint(f"找到开始识别按钮: '{btn_text}'")
                        start_btn = btn
                        break
                except:
                    continue
            
            if start_btn:
                tprint(f"点击开始识别按钮...")
                start_btn.click_input()
            else:
                tprint(f"未找到识别按钮，使用 F5 快捷键...", "WARN")
                send_keys("{F5}")
        except Exception as e:
            tprint(f"点击开始识别按钮失败: {e}", "WARN")
            tprint(f"使用 F5 快捷键作为兜底...")
            send_keys("{F5}")  # 兜底
            
        # 不再固定等待，直接进入识别完成检测
        tprint(f"等待识别完成...")
        # 尝试两种方法检测识别完成
        recognition_finished = wait_recognition_finished(main)
        if not recognition_finished:
            tprint(f"状态栏方法失败，尝试备用检测方法...")
            recognition_finished = wait_recognition_finished_backup(main)
        
        if not recognition_finished:
            tprint(f"识别 {img_path.name} 超时，跳过", "WARN")
            send_keys("^w")  # 关闭标签
            return False
        
        recognition_elapsed = time.time() - recognition_start_time
        tprint(f"识别已完成！")
        tprint(f"识别耗时: {recognition_elapsed:.1f}s")

        # Step 4: 保存为CSC格式
        save_start_time = time.time()
        tprint(f"准备保存 CSC 文件...")
        
        # 定义输出文件路径
        out_file = OUTPUT_DIR / f"{img_path.stem}.csc"
        tprint(f"目标保存路径: {out_file}")
        
        # 优化：尝试确保窗口有焦点，但不强制等待
        try:
            main.set_focus()
            time.sleep(0.1)  # 最小化焦点设置等待时间
        except:
            pass  # 如果焦点设置失败，继续执行
        
        # 发送保存快捷键
        tprint(f"发送保存快捷键 Shift+Ctrl+M...")
        send_keys("+^m")  # Shift+Ctrl+M

        # 优化：缩短保存对话框检测超时时间
        if not wait_for_save_dialog(main, timeout=3):
            tprint(f"3秒内未检测到保存控件，继续兜底处理", "WARN")
        
        # 调用原有保存逻辑以实际设置目录和文件名
        if handle_save_dialog(app, out_file):
            tprint("保存成功")
        else:
            tprint("保存可能失败", "WARN")
        
        # Step 5: 收尾
        main = wait_for_state(app, 'main') # 等待保存完成，焦点回到主窗口
        send_keys("^w")  # 关闭当前标签
        
        total_elapsed = time.time() - file_start_time
        tprint(f"文件 {img_path.name} 处理完成！总耗时: {total_elapsed:.1f}s (打开:{open_elapsed:.1f}s, 识别:{recognition_elapsed:.1f}s, 保存:{save_elapsed:.1f}s)")
        return True

    except Exception as e:
        tprint(f"处理 {img_path.name} 失败: {e}", "ERR")
        try:
            main = wait_for_state(app, 'main', timeout=5)
            send_keys("^w")  # 关闭标签
        except:
            pass
        return False

def handle_save_dialog(app: Application, out_file: Path) -> bool:
    """处理保存对话框的简化版本 - 优化版本"""
    try:
        # 在主窗口中查找保存控件
        main_window = wait_for_state(app, 'main', timeout=5)
        
        # 查找编辑框
        edit_controls = main_window.descendants(control_type="Edit")
        tprint(f"找到 {len(edit_controls)} 个编辑框", "DEBUG")
        
        # 如果找到编辑框，设置保存路径
        if len(edit_controls) >= 2:
            # 正确的方式：分别设置目录和文件名
            try:
                # 设置输出目录
                folder_edit = edit_controls[0]  # 第一个是文件夹
                folder_edit.click_input()
                time.sleep(0.1)
                send_keys("^a")  # 全选
                send_keys(str(out_file.parent), with_spaces=True)
                time.sleep(0.1)
                tprint(f"已设置输出目录: {out_file.parent}")
                
                # 设置文件名
                file_edit = edit_controls[1]  # 第二个是文件名
                file_edit.click_input()
                time.sleep(0.1)
                send_keys("^a")  # 全选
                send_keys(out_file.name, with_spaces=True)
                time.sleep(0.1)
                tprint(f"已设置文件名: {out_file.name}")
            except Exception as e:
                tprint(f"设置路径失败: {e}", "WARN")
            
            # 点击OK按钮
            try:
                buttons = main_window.descendants(control_type="Button")
                for btn in buttons:
                    try:
                        btn_text = btn.window_text()
                        if btn_text in ['OK', 'Overwrite']:
                            tprint(f"点击确认按钮: '{btn_text}'")
                            btn.click_input()
                            time.sleep(0.3)  # 进一步优化等待时间
                            return True
                    except:
                        continue
            except Exception as e:
                tprint(f"点击确认按钮失败: {e}", "WARN")
        
        # 兜底方案
        tprint(f"使用回车键确认保存...")
        send_keys("{ENTER}")
        time.sleep(0.3)  # 进一步优化等待时间
        return True
        
    except Exception as e:
        tprint(f"处理保存对话框失败: {e}", "WARN")
        send_keys("{ENTER}")  # 兜底
        return False

# -----------------------------------------------------------------------------
# 主逻辑
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    # 调试阶段：清理所有旧的 capscan 进程
    tprint("清理环境：强制关闭所有旧的 capscan.exe 实例...")
    
    # 方法1: 使用taskkill强制终止
    subprocess.run("taskkill /F /IM capscan.exe /T", capture_output=True, check=False)
    
    # 方法2: 使用PowerShell强制停止
    subprocess.run(["powershell", "-Command", "Get-Process -Name 'capscan' -ErrorAction SilentlyContinue | Stop-Process -Force"], 
                  capture_output=True, check=False)
    
    # 等待进程完全终止
    for attempt in range(5):  # 最多等待5秒，增加重试次数
        result = subprocess.run("tasklist /FI \"IMAGENAME eq capscan.exe\"", 
                              capture_output=True, text=True, check=False)
        if "capscan.exe" not in result.stdout:
            tprint("旧进程已完全清理。")
            break
        else:
            tprint(f"第{attempt+1}次清理尝试，仍有capscan进程运行...", "DEBUG")
            # 再次尝试强制清理
            subprocess.run("taskkill /F /IM capscan.exe /T", capture_output=True, check=False)
        sleep(WAIT_MEDIUM)
    else:
        tprint("经过多次尝试仍有旧进程存在，可能需要手动清理或重启系统...", "WARN")
        # 显示剩余进程信息
        result = subprocess.run(["powershell", "-Command", "Get-Process -Name 'capscan' -ErrorAction SilentlyContinue | Select-Object Id, ProcessName"], 
                              capture_output=True, text=True, check=False)
        if result.stdout.strip():
            tprint(f"剩余进程信息:\n{result.stdout}", "DEBUG")
        tprint("继续执行，但可能会有冲突...", "WARN")
    
    # 环境检查
    if not INPUT_DIR.exists():
        sys.exit(f"[ERR] 输入目录不存在: {INPUT_DIR}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 获取待处理文件列表
    png_files = sorted(INPUT_DIR.glob("*.png"))
    if not png_files:
        sys.exit(f"[INFO] {INPUT_DIR} 下未找到 *.png 文件")

    tprint(f"找到 {len(png_files)} 个PNG文件待处理")
    
    # 启动程序（只启动一次）
    tprint("启动 Capella-scan 程序...")
    app = None
    main = None
    
    try:
        # 尝试连接已存在的实例，如果失败则启动新实例
        try:
            app = Application(backend="uia").connect(path=CAPSCAN_EXE, timeout=5)
            tprint(f"连接到现有实例 (PID: {app.process})")
        except:
            tprint("未找到现有实例，启动新程序...")
            app = connect_or_start()
        
        # 等待主窗口就绪
        main = wait_for_state(app, 'main')
        tprint(f"程序已就绪，开始批量处理...")
        
        # 内层循环：逐个处理文件
        batch_start_time = time.time()
        success_count = 0
        
        for i, img_path in enumerate(png_files, 1):
            tprint(f"=== 处理第 {i}/{len(png_files)} 个文件: {img_path.name} ===")
            
            if process_single_file(app, img_path):
                success_count += 1
                tprint(f"{img_path.name} 处理成功", "OK")
            else:
                tprint(f"{img_path.name} 处理失败", "ERR")
        
        batch_elapsed = time.time() - batch_start_time
        tprint(f"批量处理完成！成功: {success_count}/{len(png_files)}", "DONE")
        tprint(f"批量处理总耗时: {batch_elapsed:.1f}s")
        if success_count > 0:
            avg_time = batch_elapsed / len(png_files)
            tprint(f"平均每文件耗时: {avg_time:.1f}s")
            tprint(f"处理效率: {success_count/batch_elapsed*60:.1f} 文件/分钟")
        
    except Exception as e:
        tprint(f"程序启动失败: {e}", "ERR")
        sys.exit(1)
    finally:
        # 清理：关闭程序
        if app:
            try:
                tprint("关闭程序...")
                main = wait_for_state(app, 'main', timeout=5)
                send_keys("%{F4}")  # Alt+F4 关闭程序
            except:
                pass 