<#
  MSDS 批量读取快速入口 (批量化读取)
  ============================================================
  基于 batch_read.py 的 PowerShell 封装: 自动定位 Python,
  支持 文件/目录/通配符 输入, 汇总统计 / 节过滤 / 检索 / 导出 / 报告.

  用法:
    .\batch-read.ps1 <docx或目录或通配符...> [选项]
    .\batch-read.ps1 doctor              # 环境自检

  选项 (透传给 batch_read.py):
    --sections 1,3,9   仅提取指定节
    --query 词         关键词检索 (多词空格分隔; 默认 AND)
    --scope label|value|all|section   检索范围 (section=精确节号)
    --any              多关键词 OR 匹配 (默认 AND)
    --name-filter 子串1,子串2   按文件名子串过滤 (逗号分隔或连续多值, OR,
                        大小写不敏感; 只查某批次型号如 BL,OS)
    --hits             文件命中清单: 每命中文件一行 "命中条目数+文件名",
                        0 命中不列出, 末尾合计 (批量审计"谁命中")
    --show-empty       检索时保留 0 命中文件空块 (默认过滤, 减噪音)
    --json | --tsv | --matrix   输出格式 (默认 text)
                        --matrix 宽表矩阵: 行=文档, 列=字段, 列顺序=
                        (节号, reader 节内顺序); 列头=Section{n} 序号范围 标题;
                        .xlsx 直接写 Excel
    --states           仅 --matrix: 每字段列后加值状态三态列
                        (有值/无数据/无此字段)
    --comp-cols        成分分列输出: 每文档一行, 成分1|CAS1|含量1|
                        成分2|CAS2|含量2 交替平铺 (text/TSV/JSON);
                        方便数据库入库; GUI 显示与核心解析不变
    --out 文件         导出到文件 (TSV 带 BOM, Excel 可开)
    --summary          打印汇总统计
    --report 文件      生成 JSON 报告 (默认不含条目全文)
    --with-entries     报告含条目全文
    --verify           token 级完整性校验
    --fail-fast        首个读取失败立即退出
    --skip-empty       跳过无内容文件
    --quiet            只输出汇总/报告
    --verbose          打印每文件进度

  示例:
    .\batch-read.ps1 "F:\...\PU-1034 msds_CN 国彩.docx" --sections 9
    .\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --tsv --out s3.tsv --summary
    .\batch-read.ps1 "F:\数据库\MSDS\英文\*.docx" --query "供应商" --scope label
    .\batch-read.ps1 "F:\数据库\MSDS" --report scan.json --summary --quiet
    .\batch-read.ps1 "F:\数据库\MSDS" --verify --report integrity.json
    .\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --out 对比表.xlsx
    .\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --states --out 对比表_带状态.xlsx
    .\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --comp-cols --tsv --out 成分分列.tsv
    # ---- 批量化检索增强 ----
    .\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇" --hits
    .\batch-read.ps1 "F:\数据库\MSDS" --query "供应商" --scope label --name-filter BL,OS --tsv
    .\batch-read.ps1 "F:\数据库\MSDS" --query "危险 警示" --any --hits --summary
  ============================================================
#>

#Requires -Version 5.1
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Script:BatchDir = $PSScriptRoot
$Script:BatchPy  = Join-Path $Script:BatchDir 'batch_read.py'

$Script:PythonCandidates = @(
    'E:\MorenAnzhuangLujing\Anaconda\python.exe',
    'C:\Python*\python.exe',
    'C:\ProgramData\anaconda3\python.exe',
    'python'
)

function Resolve-Python {
    foreach ($cand in $Script:PythonCandidates) {
        $exact = Get-ChildItem -Path $cand -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exact) { return $exact.FullName }
        $cmd = Get-Command $cand -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

# 注意: 不要用 "$x = function" 或 "$x = Main" 捕获输出 — 会把 python 的
# stdout 吞进变量. 本脚本在顶层直接调用, 退出码通过 $LASTEXITCODE 传递.
$py = Resolve-Python
if (-not $py) {
    Write-Host '  [!] 未找到 Python 解释器' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $Script:BatchPy)) {
    Write-Host "  [!] 未找到 batch_read.py: $Script:BatchPy" -ForegroundColor Red
    exit 1
}

if ($args.Count -ge 1 -and $args[0] -eq 'doctor') {
    Write-Host "  [ok] Python: $py" -ForegroundColor Green
    & $py -c "import docx; print('  [ok] python-docx')"
    exit 0
}

if ($args.Count -lt 1) {
    Write-Host '  用法: .\batch-read.ps1 <docx或目录或通配符...> [--sections] [--query] [--tsv|--json] [--out] [--summary] [--report] [--verify] ...' -ForegroundColor Yellow
    Write-Host '  提示: .\batch-read.ps1 --help 查看完整帮助' -ForegroundColor Gray
    exit 2
}

Push-Location $Script:BatchDir
# 先捕获输出并立即保存 $LASTEXITCODE (管道后会被覆盖), 再逐行透传
$outLines = & $py $Script:BatchPy @args 2>&1
$code = $LASTEXITCODE
$outLines | ForEach-Object { Write-Output $_ }
Pop-Location
exit $code