[CmdletBinding()]
param(
    [string]$Template = 'F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\推导方案\PEA-4139 MSDS_CN 冠志 模板.docx',
    [string]$Out = '',
    [string]$WriteItems = '',
    [string]$InputJson = '',
    [string]$Sections = 'all',
    [string]$FieldMap = '',
    [switch]$Yes,
    [string]$Python = 'python'
)

$ErrorActionPreference = 'Stop'
$Script = Join-Path $PSScriptRoot 'msds_form_overwrite.py'

if (-not (Test-Path -LiteralPath $Script)) {
    throw "交互脚本不存在：$Script"
}
if (-not (Test-Path -LiteralPath $Template)) {
    throw "模板不存在：$Template"
}

Write-Host '=== MSDS 17节中文交互表单 → Word覆写 ===' -ForegroundColor Cyan
Write-Host "模板：$Template"
Write-Host '说明：直接回车保留模板值；多行字段单独输入“结束”；输入 !清空 可清空字段。' -ForegroundColor Yellow

$Args = @($Script, '--template', $Template, '--sections', $Sections)
if ($Out) { $Args += @('--out', $Out) }
if ($WriteItems) { $Args += @('--write-items', $WriteItems) }
if ($InputJson) { $Args += @('--input-json', $InputJson) }
if ($FieldMap) { $Args += @('--field-map', $FieldMap) }
if ($Yes) { $Args += '--yes' }

& $Python @Args
$ExitCode = $LASTEXITCODE
if ($ExitCode -eq 0) {
    Write-Host '=== 产出流程结束：成功 ===' -ForegroundColor Green
} else {
    Write-Host "=== 产出流程结束：失败，退出码 $ExitCode ===" -ForegroundColor Red
}
exit $ExitCode
