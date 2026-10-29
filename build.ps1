param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'dist/clipslim-win-x64'))
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe' }
if (!(Test-Path -LiteralPath $compiler)) { throw 'Install .NET Framework 4.8 before building ClipSlim.' }
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$references = foreach ($assembly in @('UIAutomationClient', 'UIAutomationTypes', 'WindowsBase')) {
    $directory = Join-Path $env:WINDIR "Microsoft.NET/assembly/GAC_MSIL/$assembly"
    $reference = Get-ChildItem -LiteralPath $directory -Filter "$assembly.dll" -Recurse | Select-Object -First 1
    if (!$reference) { throw "Missing framework assembly: $assembly" }
    '/r:' + $reference.FullName
}
$executable = Join-Path $OutputDirectory 'ClipSlim-v1.7.2.exe'
& $compiler /nologo /target:winexe /optimize+ /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll @references "/out:$executable" "/win32icon:$(Join-Path $PSScriptRoot 'assets/clipslim.ico')" (Join-Path $PSScriptRoot 'src/ClipSlim.cs') (Join-Path $PSScriptRoot 'src/OnlineBrowserBridge.cs')
if ($LASTEXITCODE -ne 0) { throw 'ClipSlim compilation failed.' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'LICENSE'),(Join-Path $PSScriptRoot 'README.md'),(Join-Path $PSScriptRoot 'CHANGELOG.md'),(Join-Path $PSScriptRoot 'install-shortcut.ps1') -Destination $OutputDirectory -Force
Write-Output "Built: $executable"
