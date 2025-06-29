# Capella Scan Auto

## 项目简介
本仓库通过 `bulk_img_to_csc.py` 脚本批量调用 **capella-scan 9** 将位于 `img_in/` 目录下的 PNG 图片自动识别并保存为 `csc_out/` 目录中的 `.csc` 乐谱文件，旨在彻底解放重复点击劳动力。

* **适用版本**：capella-scan 9.x（默认安装路径 `C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe`）
* **运行环境**：Windows 10/11 + Python 3.9+（x86/32-bit）
* **自动化框架**：pywinauto（UIA backend）

## 目录结构
```
capella_scan_auto/
├─ bulk_img_to_csc.py   # 主脚本（可直接运行）
├─ requirements.txt     # Python 依赖
└─ README.md            # 说明文档
```

## 依赖安装
1. **创建并激活虚拟环境（推荐）**
   ```powershell
   py -3 -m venv venv
   .\venv\Scripts\Activate.ps1  # PowerShell
   ```
2. **安装依赖**
   ```powershell
   pip install -r requirements.txt
   # 若提示缺失 pywin32 可单独安装
   pip install pywin32
   ```
3. **capella-scan**：需已正确安装并能手动运行

## 快速开始
1. 将待识别的 PNG 图像放入 `img_in/` 目录。
2. 打开 **管理员 x86 Native Tools Command Prompt for VS 2022**（或其他**管理员 32-bit 终端**），确保脚本与目标进程位宽一致并具备强杀进程权限。
3. 运行脚本：
   ```powershell
   python bulk_img_to_csc.py
   ```4. 处理完成后，在 `csc_out/` 目录获取生成的 `.csc` 文件。

## 可调整参数
脚本顶部的 **全局等待时间配置** 与 capella-scan 安装路径均可按需修改。
```python
WAIT_SHORT = 0.5      # UI 短等待
WAIT_SAVE_DIALOG = 0.3  # 保存对话框渲染等待
CAPSCAN_EXE = r"C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe"
```

## 性能表现
经过多轮优化（包括减少固定等待、采用更快的 UIA 调用、并行化文件对话框操作等），脚本在处理单个标准 A4 琴谱 PNG 时，除核心识别（OCR）耗时外，额外开销可控制在 **2-3秒** 内。

**日志输出示例:**
脚本自带 `tprint()` 统一带时戳打印，方便观察各阶段耗时。
```
[INFO] [  0.0s] === 开始处理文件: score_page_1.png ===
[INFO] [  1.2s] 文件打开耗时: 1.1s
[INFO] [ 11.5s] 识别已完成！
[INFO] [ 11.5s] 识别耗时: 10.3s
[INFO] [ 12.8s] 保存耗时: 1.3s
[OK]   [ 13.1s] 文件 score_page_1.png 处理成功
...
[DONE] [135.2s] 批量处理完成！成功: 10/10
[INFO] [135.2s] 平均每文件耗时: 13.5s
[INFO] [135.2s] 处理效率: 4.4 文件/分钟
```

## 常见问题
| 现象 | 解决办法 |
| --- | --- |
| `PermissionError`, 无法杀掉 capscan 进程 | 确保在**管理员权限的 32-bit 终端**中运行，参考“快速开始”指引。 |
| `pywinauto.uia_element_info.UIAElementInfoError` | 检查 capella-scan 版本；若为旧版请自行调整 UIA 控件匹配逻辑 |

## 贡献
欢迎提交 PR / Issue 共同完善！

## License
MIT 
