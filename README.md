# Capella Scan Auto

## 项目简介
本仓库通过 `bulk_img_to_csc.py` 脚本批量调用 **capella-scan 9** 将指定目录下的 PNG 图片自动识别并保存为 `.csc` 乐谱文件，旨在彻底解放重复点击劳动力。

* **适用版本**：capella-scan 9.x（默认安装路径 `C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe`，可自定义）
* **运行环境**：Windows 10/11 + Python 3.9+（x86/32-bit）
* **自动化框架**：pywinauto（UIA backend）

## 目录结构
```
capella_scan_auto/
├─ bulk_img_to_csc.py   # 主脚本
├─ config.ini.template  # 配置文件模板
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
1. **首次运行：** 从 `config.ini.template` 复制一份并重命名为 `config.ini`。
2. **配置路径：** 打开 `config.ini`，根据你的环境修改 `input_dir`（输入目录）、`output_dir`（输出目录）和 `capella_scan_exe`（程序路径）。
   - 所有输入输出路径均通过 `config.ini` 配置，无需修改脚本。
   - `img_in/` 和 `csc_out/` 仅为默认值，可自定义为任意有效目录。
3. 将待识别的 PNG 图像放入你配置的 `input_dir` 目录。
4. 打开 **管理员终端**（PowerShell 或 CMD），确保脚本具备强杀进程权限。
5. 运行脚本：
   ```powershell
   python bulk_img_to_csc.py
   ```
6. 处理完成后，在你配置的 `output_dir` 目录获取生成的 `.csc` 文件。

## 可调整参数
- 所有路径配置均已移至 `config.ini` 文件，方便修改且不会被 git 提交覆盖。
- 脚本顶部的 **全局等待时间配置** 仍可按需调整以适应不同机器性能。
```python
WAIT_SHORT = 0.5      # UI 短等待
WAIT_SAVE_DIALOG = 0.3  # 保存对话框渲染等待
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
| `PermissionError`, 无法杀掉 capscan 进程 | 确保在**管理员权限的 32-bit 终端**中运行，参考"快速开始"指引。 |
| `pywinauto.uia_element_info.UIAElementInfoError` | 检查 capella-scan 版本；若为旧版请自行调整 UIA 控件匹配逻辑 |

## 贡献
欢迎提交 PR / Issue 共同完善！

## License
MIT 
