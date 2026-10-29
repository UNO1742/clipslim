$ErrorActionPreference = 'Stop'
$exe = Join-Path $PSScriptRoot 'ClipSlim-v1.7.2.exe'
if (!(Test-Path -LiteralPath $exe)) { throw 'Extract the release ZIP before creating the shortcut.' }
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('DesktopDirectory')) '轻图.lnk'))
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.IconLocation = $exe + ',0'
$shortcut.Description = '轻图：剪贴板图片压缩、WPS 单元格插入与金山在线表格自动粘贴'
$shortcut.Save()
Write-Output 'Created desktop shortcut: 轻图'
