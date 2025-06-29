# Capella Scan Auto

## 项目简介
本仓库通过 `bulk_img_to_csc.py` 脚本批量调用 **capella-scan 9**（Capella Software 出品）将位于 `img_in/` 目录下的 PNG 图片自动识别并保存为 `csc_out/` 目录中的 `.csc` 乐谱文件，旨在彻底解放重复点击劳动力。

* **适用版本**：capella-scan 9.x（默认安装路径 `C:\Program Files (x86)\capella-software\capella-scan 9\bin\capscan.exe`）
* **运行环境**：Windows 10/11 + Python 3.9+（x86/32-bit）
* **自动化框架**：pywinauto（UIA backend）

## 目录结构
```
capella_scan_auto/
├─ bulk_img_to_csc.py   # 主脚本（可直接运行）
├─ requirements.txt     # Python 依赖
├─ img_in/              # 放待识别的 .png 图像
├─ csc_out/             # 生成的 .csc 文件将存于此
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

## 日志输出
脚本自带 `tprint()` 统一带时戳打印，方便观察各阶段耗时：
```
[INFO] [  0.0s] 启动新的 capscan.exe 实例...
[INFO] [ 12.4s] 识别已完成！
[INFO] [ 15.8s] 文件 xxx.png 处理完成！总耗时: 15.8s
```

## 常见问题
| 现象 | 解决办法 |
| --- | --- |
| `PermissionError`, 无法杀掉 capscan 进程 | 以**管理员** PowerShell 运行脚本 |
| `pywinauto.uia_element_info.UIAElementInfoError` | 检查 capella-scan 版本；若为旧版请自行调整 UIA 控件匹配逻辑 |
| Git 推送失败 `Permission denied (publickey)` | 生成并添加 SSH 公钥，或切换 https 方式 |

## 贡献
欢迎提交 PR / Issue 共同完善！

## License
MIT 
