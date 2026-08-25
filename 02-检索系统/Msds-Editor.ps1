<#
  Msds Editor 统一指令
  ============================================================
  MSDS 结构读取程序封装 (结构读取\)
  流程: 导入 → 读取 → 显示 → 覆写指向分析

  用法:
    .\Msds-Editor.ps1                 启动可视化界面 (GUI，默认)
    .\Msds-Editor.ps1 gui             启动可视化界面
    .\Msds-Editor.ps1 cli <docx>      命令行解析 MSDS 16 节并打印
    .\Msds-Editor.ps1 extract <docx...>  分层检索提取 (section→大标题→小标题→字段)
        [--query 词] [--scope label|value|all|section] [--json|--tsv] [--out 文件] [--sections 1,3,9]
    .\Msds-Editor.ps1 test            运行回归测试 (pytest)
    .\Msds-Editor.ps1 doctor          环境自检 (python / python-docx / tkinter)

  示例:
    .\Msds-Editor.ps1 cli "F:\...\PU-1034 msds_CN 国彩.docx"
    .\Msds-Editor.ps1 extract "F:\...\PU-1034 msds_CN 国彩.docx" --sections 3
    .\Msds-Editor.ps1 extract "F:\模板.docx" "F:\PU-1034.docx" --query "供应商" --scope label
    .\Msds-Editor.ps1 test
  ============================================================
#>

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# 路径配置
# ============================================================
$Script:EditorRoot = $PSScriptRoot
$Script:MainPy      = Join-Path $Script:EditorRoot 'main.py'

# Python 候选 (按优先级)
$Script:PythonCandidates = @(
    'E:\MorenAnzhuangLujing\Anaconda\python.exe',
    'C:\Python*\python.exe',
    'C:\ProgramData\anaconda3\python.exe',
    'python'
)

# ============================================================
# 辅助函数
# ============================================================

function Resolve-Python {
    <# 定位可用的 Python 解释器 #>
    foreach ($cand in $Script:PythonCandidates) {
        $exact = Get-ChildItem -Path $cand -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exact) { return $exact.FullName }
        $cmd = Get-Command $cand -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Invoke-Doctor {
    <# 环境自检: python / python-docx / tkinter #>
    Write-Host ''
    Write-Host '  ---------- Msds Editor 环境自检 ----------' -ForegroundColor Cyan
    $py = Resolve-Python
    if (-not $py) {
        Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red
        return $false
    }
    Write-Host "  [ok] Python     : $py" -ForegroundColor Green
    Write-Host '  [check] python-docx / tkinter ...' -ForegroundColor Gray
    $docx = & $py -c "import docx; print('ok')" 2>&1
    if ($docx -match 'ok') { Write-Host '  [ok] python-docx' -ForegroundColor Green }
    else { Write-Host "  [!] python-docx 缺失: $docx" -ForegroundColor Red }
    $tk = & $py -c "import tkinter; print('ok')" 2>&1
    if ($tk -match 'ok') { Write-Host '  [ok] tkinter' -ForegroundColor Green }
    else { Write-Host "  [!] tkinter 缺失: $tk" -ForegroundColor Red }
    Write-Host '  -----------------------------------------' -ForegroundColor Cyan
    Write-Host ''
    return $true
}

function Invoke-Gui {
    <# 启动可视化界面 (分离进程, 不阻塞终端) #>
    $py = Resolve-Python
    if (-not $py) { Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red; return $false }
    if (-not (Test-Path $Script:MainPy)) {
        Write-Host "  [!] 未找到入口: $Script:MainPy" -ForegroundColor Red
        return $false
    }
    Write-Host "  正在启动 Msds Editor 界面 ..." -ForegroundColor Green
    Start-Process -FilePath $py -ArgumentList "`"$Script:MainPy`"" -WorkingDirectory $Script:EditorRoot
    Write-Host '  Msds Editor 界面已启动。' -ForegroundColor Green
    return $true
}

function Invoke-Cli {
    <# 命令行解析 MSDS docx #>
    param([string]$Docx)
    $py = Resolve-Python
    if (-not $py) { Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red; return $false }
    if (-not $Docx -or -not (Test-Path $Docx)) {
        Write-Host '  用法: .\Msds-Editor.ps1 cli <docx路径>' -ForegroundColor Yellow
        return $false
    }
    Write-Host "  正在解析: $Docx" -ForegroundColor Cyan
    Push-Location $Script:EditorRoot
    & $py $Script:MainPy --cli $Docx | Out-Host
    $code = $LASTEXITCODE
    Pop-Location
    return ($code -eq 0)
}

function Invoke-Extract {
    <# 分层检索提取: section→大标题→小标题→字段+内容 (支持批量/JSON/TSV) #>
    param([string[]]$Args2)
    $py = Resolve-Python
    if (-not $py) { Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red; return $false }
    if ($Args2.Count -lt 1) {
        Write-Host '  用法: .\Msds-Editor.ps1 extract <docx...> [--query 词] [--scope label|value|all|section] [--json|--tsv] [--out 文件] [--sections 1,3,9]' -ForegroundColor Yellow
        return $false
    }
    Write-Host '  分层检索提取 ...' -ForegroundColor Cyan
    Push-Location $Script:EditorRoot
    & $py $Script:MainPy --extract @Args2 | Out-Host
    $code = $LASTEXITCODE
    Pop-Location
    return ($code -eq 0)
}

function Invoke-Test {
    <# 运行回归测试 #>
    $py = Resolve-Python
    if (-not $py) { Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red; return $false }
    Write-Host '  正在运行回归测试 (pytest) ...' -ForegroundColor Cyan
    Push-Location $Script:EditorRoot
    & $py -m pytest tests -q | Out-Host
    $code = $LASTEXITCODE
    Pop-Location
    return ($code -eq 0)
}

# ============================================================
# 主流程
# ============================================================

function Main {
    param([string[]]$CommandArgs)
    $mode = if ($CommandArgs.Count -ge 1) { $CommandArgs[0].ToLower() } else { 'gui' }

    switch ($mode) {
        'gui'     { return (Invoke-Gui) }
        'cli'     { return (Invoke-Cli -Docx ($CommandArgs[1] -join ' ')) }
        'extract' { return (Invoke-Extract -Args2 $CommandArgs[1..($CommandArgs.Count-1)]) }
        'test'    { return (Invoke-Test) }
        'doctor'  { return (Invoke-Doctor) }
        default {
            Write-Host "未知指令: $mode" -ForegroundColor Red
            Write-Host '  可用指令: gui | cli <docx> | extract <docx...> [--query|--scope|--json|--tsv|--sections] | test | doctor' -ForegroundColor Yellow
            return $false
        }
    }
}

# 入口
$ok = Main -CommandArgs $args
exit $(if ($ok) { 0 } else { 1 })