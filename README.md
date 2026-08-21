﻿﻿﻿﻿﻿﻿﻿# AutoExtract 智能解压工具

输入密码后自动解压压缩包；不是压缩包则按内容识别/改名重试；支持分卷、多层嵌套、批量解压。

## 功能

- **密码自动解压**：提供 `-Password` 直接用；留空则先试空密码（未加密包直接解），加密包才提示输入密码；密码正确后自动记住复用
- **非压缩包处理**：7-Zip/WinRAR 按文件内容识别真实格式（扩展名写错也能解）；仅 .NET 引擎时自动改扩展名重试
- **分卷压缩包**：自动识别 `.001/.z01/.part1` 分卷；卷1 文件名被资源站改坏（如 `xxx.7z(删掉` 缺后缀、或前缀与兄弟卷不一致）时自动改名重试，**绝不删除原始分卷文件**
- **多层嵌套**：解压成功后自动递归解压内层压缩包（默认最深 5 层，`-MaxDepth` 可调），中间压缩包自动删除，最终文件提升到原始压缩包同目录
- **每层不同密码**：嵌套层密码与之前不同时，提示重新输入
- **批量解压**：目标为文件夹时自动扫描所有压缩包（`-Recurse` 递归子文件夹）；`-All` 内容探测伪装扩展名的压缩包
- **完成提示**：成功/失败播放提示音，GUI 版弹完成提示框

## 文件

| 文件 | 说明 |
|---|---|
| `AutoExtract.ps1` | 命令行版主脚本 |
| `AutoExtract.bat` | 命令行入口（拖放文件/文件夹即可） |
| `AutoExtractGUI.ps1` | GUI 版源码（WinForms，拖放自动解压） |
| `AutoExtractGUI.bat` | GUI 启动入口 |
| `AutoExtractGUI.exe` | GUI 版打包 exe（ps2exe，拖放到 exe 上自动解压） |
| `tests/acceptance-test.ps1` | **验收测试套件**（见下） |

## 验收测试

改代码后跑一遍验收，确认所有场景仍正常工作（防止回归）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\acceptance-test.ps1
```

覆盖 10 组场景共 23 项检查：
- 普通 zip / 加密 zip（正确密码 / 错误密码）
- 标准分卷 / 卷1 被改名破坏 / 卷1 前缀不一致
- 三层嵌套套娃（中间包清理）
- 伪装扩展名（.mp4 实为 zip）
- 外层含分卷（嵌套分卷 + 中间卷清理）
- 源文件保护（解压后原始文件必须保留）
- 无临时目录残留

## 使用方法

### 命令行版

```powershell
# 拖放文件到 AutoExtract.bat，或：
powershell -ExecutionPolicy Bypass -File AutoExtract.ps1 "D:\下载\神秘文件.bin"
powershell -ExecutionPolicy Bypass -File AutoExtract.ps1 "D:\下载\a.zip" "D:\下载\b.7z" -Password 123456
powershell -ExecutionPolicy Bypass -File AutoExtract.ps1 "D:\下载\文件夹" -Recurse -All
powershell -ExecutionPolicy Bypass -File AutoExtract.ps1 "D:\下载\套娃.7z" -MaxDepth 10
```

### GUI 版

- 双击 `AutoExtractGUI.exe` 打开，拖文件进列表点"开始解压"
- 或直接把压缩包拖到 `AutoExtractGUI.exe` 图标上，自动加入并开始解压

## 依赖

- **推荐安装 [7-Zip](https://www.7-zip.org/)**：按内容识别格式、支持 zip/7z/rar/tar/gz/bz2/xz 全格式 + 密码 + 分卷
- WinRAR/UnRAR 可选（rar 兜底）
- 都没有时仅能用 .NET 内置解 zip（无密码支持）
- 验收测试需要 7-Zip 和 Python（用于构造测试压缩包）

## 参数

| 参数 | 说明 |
|---|---|
| `-Password <pw>` | 解压密码（未加密包会被忽略） |
| `-MaxDepth <n>` | 嵌套解压最大层数，默认 5，0 = 只解一层 |
| `-Recurse` | 文件夹模式下递归子文件夹 |
| `-All` | 批量时用内容探测识别伪装扩展名的压缩包 |

## 说明

- 解压位置：原始压缩包所在目录（最终文件与压缩包同一层）
- 中间压缩包解压后自动删除（不可恢复）
- 嵌套层密码不同时提示重新输入
- **安全保证**：脚本只处理临时工作目录中的文件副本，绝不删除原始压缩包文件








