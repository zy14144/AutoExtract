#requires -Version 3.0
<#
.SYNOPSIS
    AutoExtract 验收测试套件：从头到尾跑一遍，验证是否达到预期效果。

.DESCRIPTION
    自动构造各种压缩包场景 → 调用 AutoExtract.ps1 处理 → 验证最终产物：
      - 解压出的文件存在且大小精确匹配
      - 原始输入文件必须保留（源文件保护）
      - 无 AutoExtract_work_* 残留、无中间包残留
    全部通过才算验收成功。以后改代码后跑一遍，防止回归。

.USAGE
    powershell -NoProfile -ExecutionPolicy Bypass -File tests\acceptance-test.ps1
    可选参数：
      -ScriptPath  指定 AutoExtract.ps1 路径（默认 scripts\AutoExtract.ps1）
      -Verbose     显示每个场景的详细日志
#>
param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot '..\scripts\AutoExtract.ps1'),
    [switch]$Verbose
)

$ErrorActionPreference = 'Continue'
$passCount = 0
$failCount = 0
$results = @()

# 测试根目录（临时）
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoExtract_Accept_{0}" -f ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

function Write-Result {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $script:results += [pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail }
    if ($Ok) { $script:passCount++ } else { $script:failCount++ }
    if ($Ok) {
        Write-Host ("  [PASS] {0}" -f $Name) -ForegroundColor Green
    } else {
        Write-Host ("  [FAIL] {0}  ← {1}" -f $Name, $Detail) -ForegroundColor Red
    }
}

function Invoke-AutoExtract {
    param([string]$Target, [string]$Password = '', [int]$MaxDepth = 5)
    if ($Verbose) {
        Write-Host ("    >>> AutoExtract: {0} (pw='{1}', depth={2})" -f (Split-Path $Target -Leaf), $Password, $MaxDepth) -ForegroundColor DarkGray
    }
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
    $ps = @('-Paths', "`"$Target`"")
    if ($Password) { $ps += '-Password'; $ps += "`"$Password`"" }
    $ps += '-MaxDepth'; $ps += "$MaxDepth"
    $output = & powershell @args @ps 2>&1 | Out-String
    if ($Verbose) { Write-Host ($output -replace "`n", "`n    ") -ForegroundColor DarkGray }
    return $output
}

# 断言辅助
function Test-FileExists {
    param([string]$Path, [long]$ExpectedSize = -1)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    if ($ExpectedSize -ge 0) {
        return ((Get-Item -LiteralPath $Path).Length -eq $ExpectedSize)
    }
    return $true
}

function Test-NoWorkDirs {
    param([string]$Dir)
    return @(Get-ChildItem -LiteralPath $Dir -Directory -Filter 'AutoExtract_work_*' -ErrorAction SilentlyContinue).Count -eq 0
}

# ---------- 场景准备辅助 ----------
function New-RandomFile {
    param([string]$Path, [int]$SizeMB)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    python -c "import os,random; random.seed(42); open(r'$Path','wb').write(bytes(random.getrandbits(8) for _ in range($SizeMB*1024*1024)))"
}

function New-SimpleZip {
    param([string]$ZipPath, [string]$FileContent, [string]$InnerName = 'content.txt', [string]$Password = '')
    $tmp = Join-Path (Split-Path -Parent $ZipPath) ('_z_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    Set-Content -Path (Join-Path $tmp $InnerName) -Value $FileContent -Encoding UTF8
    if ($Password) {
        & 'E:\Software\7-Zip\7z.exe' a -y -tzip ("-p{0}" -f $Password) $ZipPath (Join-Path $tmp $InnerName) | Out-Null
    } else {
        & 'E:\Software\7-Zip\7z.exe' a -y -tzip $ZipPath (Join-Path $tmp $InnerName) | Out-Null
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# =========================================================
# 场景 1：普通 zip
# =========================================================
Write-Host "`n=== 场景 1: 普通 zip ===" -ForegroundColor Cyan
$s1 = Join-Path $testRoot 's1'; New-Item -ItemType Directory -Path $s1 -Force | Out-Null
New-SimpleZip -ZipPath (Join-Path $s1 'plain.zip') -FileContent 'PLAIN CONTENT'
$out1 = Invoke-AutoExtract -Target (Join-Path $s1 'plain.zip') -Password 'x'
Write-Result '解压出 content.txt' (Test-FileExists (Join-Path $s1 'content.txt'))
Write-Result '源文件保留' (Test-FileExists (Join-Path $s1 'plain.zip'))
Write-Result '无 work 残留' (Test-NoWorkDirs $s1)

# =========================================================
# 场景 2：加密 zip，正确密码
# =========================================================
Write-Host "`n=== 场景 2: 加密 zip（正确密码）===" -ForegroundColor Cyan
$s2 = Join-Path $testRoot 's2'; New-Item -ItemType Directory -Path $s2 -Force | Out-Null
New-SimpleZip -ZipPath (Join-Path $s2 'enc.zip') -FileContent 'SECRET DATA 12345' -Password 'pw123'
$out2 = Invoke-AutoExtract -Target (Join-Path $s2 'enc.zip') -Password 'pw123'
Write-Result '解压出 content.txt' (Test-FileExists (Join-Path $s2 'content.txt'))
Write-Result '源文件保留' (Test-FileExists (Join-Path $s2 'enc.zip'))

# =========================================================
# 场景 3：加密 zip，错误密码 → 必须明确失败
# =========================================================
Write-Host "`n=== 场景 3: 加密 zip（错误密码 → 应失败）===" -ForegroundColor Cyan
$s3 = Join-Path $testRoot 's3'; New-Item -ItemType Directory -Path $s3 -Force | Out-Null
New-SimpleZip -ZipPath (Join-Path $s3 'enc2.zip') -FileContent 'SHOULD NOT EXTRACT' -Password 'realpw'
$out3 = Invoke-AutoExtract -Target (Join-Path $s3 'enc2.zip') -Password 'wrongpw'
# 错误密码的判定：不依赖中文输出匹配（编码可能乱码），直接检查产物不存在 + 源保留
Write-Result '错误密码未解出文件' (-not (Test-FileExists (Join-Path $s3 'content.txt')))
Write-Result '源文件保留' (Test-FileExists (Join-Path $s3 'enc2.zip'))
# 无 work 残留（失败路径也会清理临时目录）
Write-Result '无 work 残留' (Test-NoWorkDirs $s3)

# =========================================================
# 场景 4：标准分卷（.001/.002 命名）
# =========================================================
Write-Host "`n=== 场景 4: 标准分卷 ===" -ForegroundColor Cyan
$s4 = Join-Path $testRoot 's4'; New-Item -ItemType Directory -Path $s4 -Force | Out-Null
New-RandomFile -Path (Join-Path $s4 'src\data.bin') -SizeMB 5
& 'E:\Software\7-Zip\7z.exe' a -v3m -tzip (Join-Path $s4 'vol.zip') (Join-Path $s4 'src\data.bin') | Out-Null
$out4 = Invoke-AutoExtract -Target (Join-Path $s4 'vol.zip.001') -Password 'x'
Write-Result '解出 data.bin (5MB)' (Test-FileExists (Join-Path $s4 'data.bin') 5242880)
Write-Result '分卷保留' ((Test-FileExists (Join-Path $s4 'vol.zip.001')) -and (Test-FileExists (Join-Path $s4 'vol.zip.002')))

# =========================================================
# 场景 5：分卷卷1 被改名破坏（去掉 .001 后缀）
# =========================================================
Write-Host "`n=== 场景 5: 分卷卷1 缺 .001 后缀 ===" -ForegroundColor Cyan
$s5 = Join-Path $testRoot 's5'; New-Item -ItemType Directory -Path $s5 -Force | Out-Null
New-RandomFile -Path (Join-Path $s5 'src\data.bin') -SizeMB 5
& 'E:\Software\7-Zip\7z.exe' a -v3m -tzip (Join-Path $s5 'vol.zip') (Join-Path $s5 'src\data.bin') | Out-Null
Rename-Item (Join-Path $s5 'vol.zip.001') -NewName 'vol.zip(删掉'
$out5 = Invoke-AutoExtract -Target (Join-Path $s5 'vol.zip(删掉') -Password 'x'
Write-Result '解出 data.bin (5MB)' (Test-FileExists (Join-Path $s5 'data.bin') 5242880)
Write-Result '卷1源文件保留' (Test-FileExists (Join-Path $s5 'vol.zip(删掉'))

# =========================================================
# 场景 6：分卷卷1 前缀不一致（code.zip(删掉.001 vs code.zip.002）
# =========================================================
Write-Host "`n=== 场景 6: 分卷卷1 前缀不一致 ===" -ForegroundColor Cyan
$s6 = Join-Path $testRoot 's6'; New-Item -ItemType Directory -Path $s6 -Force | Out-Null
New-RandomFile -Path (Join-Path $s6 'src\data.bin') -SizeMB 5
& 'E:\Software\7-Zip\7z.exe' a -v3m -tzip (Join-Path $s6 'code.zip') (Join-Path $s6 'src\data.bin') | Out-Null
Rename-Item (Join-Path $s6 'code.zip.001') -NewName 'code.zip(删掉.001'
$out6 = Invoke-AutoExtract -Target (Join-Path $s6 'code.zip(删掉.001') -Password 'x'
Write-Result '解出 data.bin (5MB)' (Test-FileExists (Join-Path $s6 'data.bin') 5242880)

# =========================================================
# 场景 7：多层嵌套（3 层套娃）
# 注意：素材放在独立子目录 _mat，不在 s7 根目录留多余输入文件
# =========================================================
Write-Host "`n=== 场景 7: 三层嵌套套娃 ===" -ForegroundColor Cyan
$s7 = Join-Path $testRoot 's7'; New-Item -ItemType Directory -Path $s7 -Force | Out-Null
$s7mat = Join-Path $s7 '_mat'; New-Item -ItemType Directory -Path $s7mat -Force | Out-Null
New-SimpleZip -ZipPath (Join-Path $s7mat 'inner.zip') -FileContent 'NESTED LEVEL 3' -InnerName 'deep.txt'
New-SimpleZip -ZipPath (Join-Path $s7mat 'mid.zip') -FileContent 'PLACEHOLDER' -InnerName 'inner.zip'
# 用真实 inner.zip 替换 mid 里的占位
$tmp7 = Join-Path $s7mat '_m'; New-Item -ItemType Directory -Path $tmp7 -Force | Out-Null
Copy-Item (Join-Path $s7mat 'inner.zip') (Join-Path $tmp7 'inner.zip')
Remove-Item (Join-Path $s7mat 'mid.zip') -Force
& 'E:\Software\7-Zip\7z.exe' a -y -tzip (Join-Path $s7mat 'mid.zip') (Join-Path $tmp7 'inner.zip') | Out-Null
Remove-Item $tmp7 -Recurse -Force
New-SimpleZip -ZipPath (Join-Path $s7mat 'outer.zip') -FileContent 'PLACEHOLDER' -InnerName 'mid.zip'
$tmp7b = Join-Path $s7mat '_o'; New-Item -ItemType Directory -Path $tmp7b -Force | Out-Null
Copy-Item (Join-Path $s7mat 'mid.zip') (Join-Path $tmp7b 'mid.zip')
Remove-Item (Join-Path $s7mat 'outer.zip') -Force
& 'E:\Software\7-Zip\7z.exe' a -y -tzip (Join-Path $s7mat 'outer.zip') (Join-Path $tmp7b 'mid.zip') | Out-Null
Remove-Item $tmp7b -Recurse -Force
# 只把 outer.zip 复制到 s7 根目录作为唯一输入
Copy-Item (Join-Path $s7mat 'outer.zip') (Join-Path $s7 'outer.zip') -Force
$out7 = Invoke-AutoExtract -Target (Join-Path $s7 'outer.zip') -Password 'x'
Write-Result '解出 deep.txt' (Test-FileExists (Join-Path $s7 'deep.txt'))
# 中间包检查：s7 根目录应只有 outer.zip（源）和 deep.txt（产物）
$s7left = @(Get-ChildItem -LiteralPath $s7 -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('outer.zip','deep.txt') })
Write-Result '中间包删除' ($s7left.Count -eq 0)
Write-Result '外层源文件保留' (Test-FileExists (Join-Path $s7 'outer.zip'))

# =========================================================
# 场景 8：伪装扩展名（.mp4 实为 zip）
# =========================================================
Write-Host "`n=== 场景 8: 伪装扩展名 (.mp4 实为 zip) ===" -ForegroundColor Cyan
$s8 = Join-Path $testRoot 's8'; New-Item -ItemType Directory -Path $s8 -Force | Out-Null
New-SimpleZip -ZipPath (Join-Path $s8 'fake.mp4') -FileContent 'DISGUISED ZIP CONTENT'
$out8 = Invoke-AutoExtract -Target (Join-Path $s8 'fake.mp4') -Password 'x'
Write-Result '解出 content.txt' (Test-FileExists (Join-Path $s8 'content.txt'))
Write-Result '伪装源文件保留' (Test-FileExists (Join-Path $s8 'fake.mp4'))

# =========================================================
# 场景 9：嵌套分卷（外层含分卷，处理后中间卷清理）
# =========================================================
Write-Host "`n=== 场景 9: 外层含分卷（嵌套）===" -ForegroundColor Cyan
$s9 = Join-Path $testRoot 's9'; New-Item -ItemType Directory -Path $s9 -Force | Out-Null
New-RandomFile -Path (Join-Path $s9 'src\vid.mp4') -SizeMB 4
& 'E:\Software\7-Zip\7z.exe' a -v3m -tzip -psamepass (Join-Path $s9 'code.zip') (Join-Path $s9 'src\vid.mp4') | Out-Null
Rename-Item (Join-Path $s9 'code.zip.001') -NewName 'code.zip(删掉.001'
# 动态收集实际分卷文件（避免 .003 硬编码）
$s9vols = @(Get-ChildItem -LiteralPath $s9 -File -Filter 'code.zip*' | ForEach-Object { $_.FullName })
& 'E:\Software\7-Zip\7z.exe' a -y -t7z -psamepass (Join-Path $s9 'outer.7z') $s9vols | Out-Null
Remove-Item $s9vols -Force -ErrorAction SilentlyContinue
$out9 = Invoke-AutoExtract -Target (Join-Path $s9 'outer.7z') -Password 'samepass'
Write-Result '解出 vid.mp4 (4MB)' (Test-FileExists (Join-Path $s9 'vid.mp4') 4194304)
Write-Result '分卷中间包清理' (@(Get-ChildItem $s9 -Filter 'code.zip*' -ErrorAction SilentlyContinue).Count -eq 0)
Write-Result '外层源文件保留' (Test-FileExists (Join-Path $s9 'outer.7z'))
Write-Result '无 work 残留' (Test-NoWorkDirs $s9)

# =========================================================
# 场景 9.5：嵌入式 zip（mp4 头 + 尾部 zip 数据 + 内层 7z，模拟用户真实案例）
# =========================================================
Write-Host "`n=== 场景 9.5: 嵌入式 zip（mp4 头+zip 尾）===" -ForegroundColor Cyan
$s95 = Join-Path $testRoot 's95'; New-Item -ItemType Directory -Path $s95 -Force | Out-Null
# 1) 造一个真实 7z 文件（语文818.7z）内含视频（模拟外层 zip 解出的嵌套包）
New-RandomFile -Path (Join-Path $s95 'inner\1.mp4') -SizeMB 2
& 'E:\Software\7-Zip\7z.exe' a -y -t7z -p112233 (Join-Path $s95 '语文818.7z') (Join-Path $s95 'inner\1.mp4') | Out-Null
# 2) 把 7z 打包进一个加密 zip（外层 zip，密码 112233）
$tmpZip = Join-Path $s95 'outer_tmp.zip'
& 'E:\Software\7-Zip\7z.exe' a -y -tzip -p112233 $tmpZip (Join-Path $s95 '语文818.7z') | Out-Null
# 3) 造一个 mp4 头（3MB 随机数据 + mp4 魔数）
$mp4Head = Join-Path $s95 'head.bin'
$fs = [System.IO.File]::Create($mp4Head)
$fs.Write([byte[]]@(0x00,0x00,0x00,0x20,0x66,0x74,0x79,0x70,0x69,0x73,0x6F,0x6D),0,12)
$rnd = New-Object Random 42
$buf = New-Object byte[] (3MB)
$rnd.NextBytes($buf)
$fs.Write($buf, 0, $buf.Length)
$fs.Close()
# 4) 拼接：mp4 头 + zip 数据
$fakeFile = Join-Path $s95 '视频.mp4'
Copy-Item $mp4Head $fakeFile
$fs = [System.IO.File]::OpenWrite($fakeFile)
$fs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
$zipBytes = [System.IO.File]::ReadAllBytes($tmpZip)
$fs.Write($zipBytes, 0, $zipBytes.Length)
$fs.Close()
Remove-Item $mp4Head, $tmpZip, (Join-Path $s95 '语文818.7z') -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $s95 'inner') -Recurse -Force -ErrorAction SilentlyContinue
$out95 = Invoke-AutoExtract -Target $fakeFile -Password '112233'
Write-Result '嵌入式 zip 解出内层视频' (Test-FileExists (Join-Path $s95 '1.mp4') 2097152)
Write-Result '源文件保留' (Test-FileExists $fakeFile)
Write-Result '中间包清理' (@(Get-ChildItem $s95 -Filter '*.7z' -ErrorAction SilentlyContinue).Count -eq 0 -and @(Get-ChildItem $s95 -Filter '*embedded*' -ErrorAction SilentlyContinue).Count -eq 0)
Write-Result '无 work 残留' (Test-NoWorkDirs $s95)

# =========================================================
# 场景 9.6：伪装分卷（分卷文件加 .pdf 伪装后缀，资源站常见）
# =========================================================
Write-Host "`n=== 场景 9.6: 伪装分卷（.001.pdf/.002.pdf）===" -ForegroundColor Cyan
$s96 = Join-Path $testRoot 's96'; New-Item -ItemType Directory -Path $s96 -Force | Out-Null
New-RandomFile -Path (Join-Path $s96 'big.bin') -SizeMB 3
& 'E:\Software\7-Zip\7z.exe' a -v2m -tzip -psamepass (Join-Path $s96 'Fayrrara.7z') (Join-Path $s96 'big.bin') | Out-Null
Get-ChildItem $s96 -Filter 'Fayrrara.7z.*' | Where-Object { $_.Name -match '\.0\d{2}$' } | ForEach-Object {
    $newname = $_.Name + '.pdf'
    Rename-Item -LiteralPath $_.FullName -NewName $newname
}
$out96 = Invoke-AutoExtract -Target $s96 -Password 'samepass'
Write-Result '伪装分卷解出 big.bin (3MB)' (Test-FileExists (Join-Path $s96 'big.bin') 3145728)
Write-Result '源分卷文件保留' (@(Get-ChildItem $s96 -Filter 'Fayrrara.7z.*.pdf').Count -eq 2)
# 不依赖中文计数文本（GBK 管道会破坏中文，中文匹配不可靠），用产物 + 残留断言覆盖核心
Write-Result '无 work 残留' (Test-NoWorkDirs $s96)

# =========================================================
# 场景 10.5：GUI 版验证（语法、启动、exe 打包时效）
# =========================================================
Write-Host "`n=== 场景 10.5: GUI 版验证 ===" -ForegroundColor Cyan
$guiSrc = Join-Path $PSScriptRoot '..\scripts\AutoExtractGUI.ps1'
$guiExe = Join-Path $PSScriptRoot '..\AutoExtractGUI.exe'

# 10.5a: GUI 脚本语法检查
if (Test-Path $guiSrc) {
    $gTokens = $null; $gErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($guiSrc, [ref]$gTokens, [ref]$gErrors) | Out-Null
    Write-Result 'GUI 脚本语法正确' ($gErrors.Count -eq 0)
} else {
    Write-Result 'GUI 脚本语法正确' $false "找不到 $guiSrc"
}

# 10.5b: exe 存在
Write-Result 'GUI exe 存在' (Test-FileExists $guiExe)

# 10.5c: exe 打包时效（exe 时间戳必须晚于 GUI 源码）
if ((Test-Path $guiSrc) -and (Test-Path $guiExe)) {
    $srcT = (Get-Item $guiSrc).LastWriteTime
    $exeT = (Get-Item $guiExe).LastWriteTime
    Write-Result 'exe 比源码新（已重新打包）' ($exeT -gt $srcT) "exe=$exeT src=$srcT"
}

# 10.5d: GUI 启动不崩溃（启动 exe，等 5 秒，进程应存活；再关闭）
if (Test-Path $guiExe) {
    $gProc = $null
    try {
        $gProc = Start-Process $guiExe -PassThru -ErrorAction Stop
        Start-Sleep -Seconds 6
        $alive = -not $gProc.HasExited
        Write-Result 'GUI 启动不崩溃' $alive
        if ($alive) { Stop-Process -Id $gProc.Id -Force -ErrorAction SilentlyContinue }
    } catch {
        Write-Result 'GUI 启动不崩溃' $false $_.Exception.Message
        if ($gProc) { Stop-Process -Id $gProc.Id -Force -ErrorAction SilentlyContinue }
    }
}

# =========================================================
# 场景 10：无残留验证（全部场景的 work 目录）
# =========================================================
Write-Host "`n=== 场景 10: 全局无残留 ===" -ForegroundColor Cyan
$allWork = @(Get-ChildItem -Path $testRoot -Directory -Recurse -Filter 'AutoExtract_work_*' -ErrorAction SilentlyContinue)
Write-Result '所有场景无 work 残留' ($allWork.Count -eq 0)

# =========================================================
# 汇总
# =========================================================
Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ("验收结果: 通过 {0} | 失败 {1}" -f $passCount, $failCount) -ForegroundColor $(if ($failCount -eq 0) { 'Green' } else { 'Red' })
Write-Host ("=" * 60) -ForegroundColor Cyan
if ($failCount -gt 0) {
    Write-Host "`n失败的检查项:" -ForegroundColor Red
    $results | Where-Object { -not $_.Ok } | ForEach-Object { Write-Host ("  - {0}: {1}" -f $_.Name, $_.Detail) -ForegroundColor Red }
    Write-Host "`n验收未通过！脚本需要修复。" -ForegroundColor Red
} else {
    Write-Host "`n✅ 验收全部通过！脚本可以交付。" -ForegroundColor Green
}

# 清理测试数据（保留失败现场便于排查）
if ($failCount -eq 0) {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
} else {
    Write-Host ("测试数据保留在: {0}" -f $testRoot) -ForegroundColor Yellow
}
