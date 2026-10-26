# 轻图 · ClipSlim

<img src="assets/clipslim.png" width="96" height="96" alt="轻图图标">

Windows 上的轻量图片压缩工具：读取剪贴板，压缩图片，复制回剪贴板，或放进 WPS / 金山在线表格的单元格。

**[下载 Windows 版本](https://github.com/UNO1742/clipslim/releases/latest)** · [配套无缝拼图](https://github.com/UNO1742/seamless-puzzle-compact)

## 功能

- 中文紧凑界面，可置顶；支持剪贴板、文件和拖放。
- JPEG 质量、PNG 透明背景、最长边缩放；压缩前后大小对比。
- 复制真实图片、复制压缩文件、另存为。
- 桌面 WPS：按当前选中单元格或合并区域等比居中插入；保留锚点，随单元格调整。
- 金山在线表格：记住 Edge 中最后选中的文档，点击后核对网址、切回页面并自动粘贴。
- 接收无缝拼图的本地请求，压缩后使用同一目标模式处理。

## 使用

解压发布 ZIP，运行 `ClipSlim-v1.7.1.exe`。可运行 `install-shortcut.ps1` 创建桌面“轻图”快捷方式。

### 桌面 WPS

1. 在轻图底部选择“桌面 WPS”。
2. 在 WPS 表格中选中目标单元格，退出文字编辑状态。
3. 读取图片并压缩，点击“插入 WPS 单元格”。

### 金山在线表格（Edge）

1. 首次在自己的 Edge 中设置：**插入 → 图片 → 从外部粘贴图片时 → 默认为单元格图片**。
2. 轻图底部选择“金山在线表格”。
3. 在网页单击目标格并停留约一秒，再回轻图点击“自动插入在线表格”。

程序只对已识别的 HTTPS 金山文档页面自动发送粘贴请求；不能确认网址或焦点时停止并提示。发送后在网页确认结果，避免重复点击。Edge 最大化、全屏状态会保留，最小化时才恢复窗口。不需要浏览器扩展。

在线粘贴使用兼容 PNG 图片，保留处理后的尺寸与画面。网页最终存储体积由金山文档决定，不保证等于 JPEG 压缩结果的字节数；需要保留 JPEG 文件体积时，另存为后通过网页图片入口上传。

## 环境

- Windows 10/11、.NET Framework 4.8。
- 桌面插入需要已安装且可通过 COM 连接的 WPS 表格；在线自动粘贴需要桌面版 Microsoft Edge。
- 不包含账号、Cookie、用户图片或具体工作表链接。配置与临时图片写入当前用户的本地应用数据目录。

## 从源码构建

Windows PowerShell 下运行，无需另装 .NET SDK：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

输出到 `dist/clipslim-win-x64/`。也可以用 .NET SDK / Visual Studio 构建 `ClipSlim.csproj`，需要 .NET Framework 4.8 参考程序集。

纯图片处理检查：

```powershell
$test = Start-Process .\dist\clipslim-win-x64\ClipSlim-v1.7.1.exe -ArgumentList '--self-test', "$PWD\artifacts\self-test" -Wait -PassThru
if ($test.ExitCode -ne 0) { throw 'Self-test failed' }
```

`--integration-test <目录>` 检查剪贴板与在线模式，**会改变系统剪贴板**。`--layout-test <报告文件>` 检查紧凑布局。WPS 插入与浏览器上传需要在自己的测试表中验证。

## 项目结构

- `src/ClipSlim.cs`：界面、压缩、剪贴板、WPS COM 和拼图命名管道。
- `src/OnlineBrowserBridge.cs`：前台 Edge 目标识别、网址和焦点检查、自动粘贴。
- `assets/`：本项目生成的图标；PNG 用于预览，ICO 用于程序和任务栏。

## 开源许可

MIT，见 [LICENSE](LICENSE)。轻图为独立实现。欢迎提交问题和改进；报告问题时请使用测试图片和测试表格复现。
