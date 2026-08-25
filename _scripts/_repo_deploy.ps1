# -*- coding: utf-8 -*-
# 归类部署脚本: 冠志\MSDS (01-07 系统化工作区) → MSDS--TDS-repo (归档区)
# 原则: 只推 代码/程序/技能/插件/文档/模板/结构化库表; 不推 MSDS 产品文档与法规 PDF 原文.
# 工作区结构 (2026-08-18 重组): 01-表单系统 / 02-检索系统 / 03-数据库 / 04-推断引擎 /
#                               05-覆写模块 / 06-联网搜证 / 07-多格式输出 / docs / _scripts
$ErrorActionPreference = 'Stop'

$src = 'F:\正式项目与模块化内容\冠志\MSDS'
$dst = 'F:\正式项目与模块化内容\冠志\MSDS--TDS-repo'

function Copy-Tree([string]$from, [string]$to, [string[]]$excludePat = @()) {
    $from = Join-Path $src $from
    $to = Join-Path $dst $to
    New-Item -ItemType Directory -Path $to -Force | Out-Null
    $items = Get-ChildItem -LiteralPath $from -Force -ErrorAction SilentlyContinue
    foreach ($it in $items) {
        $skip = $false
        foreach ($p in $excludePat) {
            if ($it.Name -like $p) { $skip = $true; break }
        }
        if ($skip) { continue }
        Copy-Item -LiteralPath $it.FullName -Destination $to -Recurse -Force
    }
    Write-Host ("  ✓ {0}" -f $to)
}

function Copy-One([string]$from, [string]$to) {
    $f = Join-Path $src $from
    $t = Join-Path $dst $to
    New-Item -ItemType Directory -Path (Split-Path $t -Parent) -Force | Out-Null
    if (Test-Path -LiteralPath $f) { Copy-Item -LiteralPath $f -Destination $t -Force }
    Write-Host ("  ✓ {0}" -f $to)
}

Write-Host '=== 01-表单系统 ==='
Copy-Tree '01-表单系统' '01-表单系统' @('__pycache__')
# 表单代码在 02-检索系统 内运行, 归档时复制到 01 使归档区独立完整
Copy-One '02-检索系统\gui\form_window.py' '01-表单系统\gui\form_window.py'
Copy-One '02-检索系统\core\form_schema.py' '01-表单系统\core\form_schema.py'

Write-Host '=== 02-检索系统 (结构读取 完整可运行包) ==='
Copy-Tree '02-检索系统' '02-检索系统' @('.git', '.pytest_cache', '__pycache__', '_scratch')

Write-Host '=== 03-数据库 ==='
Copy-Tree '03-数据库\正式库\Data Base' '03-数据库\Data Base'
Copy-Tree '03-数据库\正式库\标准字段数据库Excel' '03-数据库\标准字段数据库Excel' @('~$*')
Copy-One '03-数据库\正式库\DB\标准库_入库总表_20260813_v17.xlsx' '03-数据库\DB\标准库_入库总表_20260813_v17.xlsx'
foreach ($f in 'msds_db.py', 'database.py', 'pivot_table.py') {
    Copy-One "02-检索系统\core\$f" "03-数据库\core\$f"
}
foreach ($f in 'build_msds_db.py', 'build_database.py', 'build_pivot_table.py',
              'build_standard_table.py', 'build_std_matrix.py',
              'fill_std_library.py', 'extract_std_structure.py') {
    Copy-One "02-检索系统\tools\$f" "03-数据库\tools\$f"
}
foreach ($f in '模板17节结构清单.md', '数据库检索需求.md', 'PRD_数据库检索.md') {
    Copy-One "02-检索系统\docs\$f" "03-数据库\docs\$f"
}
Copy-Tree '03-数据库\数据清理模块' '03-数据库\数据清理模块' @('__pycache__')
Copy-Tree '03-数据库\数据清理模块\标准模板\标准模板\定稿模板' '03-数据库\定稿模板' @('备份*', '模板草稿')

Write-Host '=== 04-推断引擎 ==='
Copy-Tree '04-推断引擎\判断skill' '04-推断引擎\判断skill' @('__pycache__')
Copy-Tree '04-推断引擎\法规匹配库\推断引擎数据' '04-推断引擎\推断引擎数据'
foreach ($f in '国内外法规归档规划.md', '国内外法规归档清单.md', 'SDS最新版框架参考.md') {
    Copy-One "04-推断引擎\法规匹配库\$f" "04-推断引擎\$f"
}
Copy-One '04-推断引擎\_scan_scale.py' '04-推断引擎\_scan_scale.py'

Write-Host '=== 05-覆写模块 ==='
Copy-Tree '05-覆写模块' '05-覆写模块' @('__pycache__', '_scratch', '_archive')

Write-Host '=== 06-联网搜证 ==='
Copy-Tree '06-联网搜证' '06-联网搜证' @('__pycache__')
# 清理工具目录内的一级 _scratch (Copy-Tree 的 exclude 只过滤顶层, 子目录需显式删除)
$t06Scratch = Join-Path $dst '06-联网搜证\tools\_scratch'
if (Test-Path -LiteralPath $t06Scratch) { Remove-Item -LiteralPath $t06Scratch -Recurse -Force }
foreach ($f in '标准原文归档\型号与标准\型号与标准法规清单.md',
              '标准原文归档\型号与标准\EN71-3_19元素迁移限量表.csv',
              '标准原文归档\型号与标准\RoHS2.0_十项限值表.csv') {
    Copy-One "04-推断引擎\法规匹配库\$f" "06-联网搜证\$f"
}

Write-Host '=== 07-多格式输出 ==='
Copy-Tree '07-多格式输出' '07-多格式输出' @('~$*')
Copy-One '03-数据库\正式库\标准字段数据库Excel\导出表\PEA-4139 MSDS_CN 冠志 模板_信息_20260817_085127.xlsx' '07-多格式输出\导出表样例\PEA-4139_透视总表.xlsx'

Write-Host '=== 根: PRD ==='
Copy-One 'docs\PRD_MSDS推导产出系统_全功能需求文档.md' 'docs\PRD_MSDS推导产出系统_全功能需求文档.md'

Write-Host '`n=== 完成 ==='
