#requires -Version 3.0
<#
.SYNOPSIS
    智能解压脚本：输入密码后自动解压压缩包；如果目标文件不是压缩包，会尝试改成常见压缩包格式再解压。

.DESCRIPTION
    引擎优先级：
      1. 7-Zip (7z.exe)   —— 按文件内容识别真实格式，扩展名写错也能解；支持 zip/7z/rar/tar/gz/bz2/xz 等 + 密码
      2. WinRAR / UnRAR   —— 支持 rar/zip 等 + 密码
      3. .NET ZipFile     —— 仅 zip 内容（无密码支持，PowerShell 5.1 限制）

    密码策略：
      - 提供 -Password 参数：直接使用该密码尝试
      - 未提供：先尝试空密码（未加密的包直接成功，不打扰你）
      - 遇到加密包时才提示输入密码；密码正确后会被记住，后续文件自动复用

    解压位置：原文件同目录下，以"原文件名(去扩展名)"命名的文件夹。目录已存在则自动加序号。

    多层嵌套压缩包：解压成功后会自动扫描产物，发现压缩包继续递归解压（默认最深 5 层，
    可用 -MaxDepth 调整；设为 0 则禁用嵌套解压）。嵌套包复用已记住的密码。
    嵌套产物中的任何文件都会被尝试——伪装扩展名（如 .mp4 实为 zip）也能逐层解开。
    嵌套解压完成后：中间层压缩包自动删除，最终文件全部提升到【原始压缩包所在目录】，
    与原始压缩包保持同一层（不再层层嵌套子目录）。

    批量解压：目标为文件夹时，自动扫描其中所有压缩包逐个处理（配合 -Recurse 递归子文件夹）。
    默认只按压缩扩展名识别；加 -All 会用 7-Zip 按内容探测所有文件（能认出伪装扩展名的压缩包）。

    改名重试说明：仅当系统只有 .NET 内置引擎（无 7-Zip/WinRAR）时才会执行改扩展名重试；
    有 7-Zip/WinRAR 时它们按文件内容识别真实格式，扩展名写错也能直接解，无需改名。

    用法（支持拖放，可同时给多个文件/文件夹）：
      .\AutoExtract.ps1 "D:\下载\神秘文件.bin"
      .\AutoExtract.ps1 "D:\下载\a.zip" "D:\下载\b.7z" -Password 123456
      .\AutoExtract.ps1 "D:\下载\某个文件夹" -Recurse -All
      .\AutoExtract.ps1 "D:\下载\套娃.7z" -MaxDepth 10
#>
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Paths,

    [string]$Password,

    [switch]$Recurse,

    [int]$MaxDepth = 5,

    [switch]$All
)

# 外部命令(7z/rar)输出 stderr 时不触发终止错误；错误统一由退出码判断
$ErrorActionPreference = 'Continue'

# 控制台输出编码：保证中文正常显示（配合 chcp 65001）
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# ---------- 全局状态 ----------
$script:SevenZip = $null        # 7z.exe 路径
$script:RarTool   = $null        # WinRAR/UnRAR 路径
$script:AutoPass  = ''           # 当前使用的密码（参数或用户输入的）
$script:PassFromParam = -not [string]::IsNullOrEmpty($Password)
$script:ArchiveExts = @('.zip','.7z','.rar','.tar','.gz','.tgz','.bz2','.tbz2','.xz','.txz','.lzma','.zst','.zstd','.001','.cab','.iso','.wim','.exe')

# ---------- 工具查找 ----------
function Find-Tool {
    param([string[]]$Names, [string[]]$KnownPaths)
    foreach ($n in $Names) {
        $c = Get-Command $n -ErrorAction SilentlyContinue
        if ($c) { return $c.Source }
    }
    foreach ($p in $KnownPaths) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

function Init-Engines {
    # 7-Zip：注册表 + PATH + 常见路径
    $reg7z = @()
    foreach ($root in 'HKLM:\SOFTWARE\7-Zip','HKLM:\SOFTWARE\WOW6432Node\7-Zip','HKCU:\SOFTWARE\7-Zip') {
        try {
            $v = (Get-ItemProperty -Path $root -Name 'Path' -ErrorAction Stop).Path
            if ($v) { $reg7z += (Join-Path $v '7z.exe') }
        } catch {}
    }
    $script:SevenZip = Find-Tool -Names @('7z.exe','7za.exe') -KnownPaths ($reg7z + @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        "$env:LOCALAPPDATA\Programs\7-Zip\7z.exe"
    ))

    # WinRAR / UnRAR：优先命令行版 UnRAR（避免 GUI 弹窗）
    $script:RarTool = Find-Tool -Names @('UnRAR.exe','WinRAR.exe','rar.exe') -KnownPaths @(
        "$env:ProgramFiles\WinRAR\UnRAR.exe",
        "${env:ProgramFiles(x86)}\WinRAR\UnRAR.exe",
        "$env:ProgramFiles\WinRAR\WinRAR.exe",
        "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe"
    )
}

# ---------- 安全读取密码（隐藏输入） ----------
function Read-Secret {
    param([string]$Prompt)
    $ss = Read-Host -Prompt $Prompt -AsSecureString
    if ($null -eq $ss) { return '' }
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# ---------- 输出目录是否有实际文件（非空目录壳） ----------
function Test-DirHasContent {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    # 必须存在"文件"才算有内容：7z 解压失败时可能只创建空目录结构，不能算成功
    return @(Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue).Count -gt 0
}

# ---------- 引擎：7-Zip ----------
# 返回: ok | wrongpass | notarchive | error
function Test-ArchiveWith7z {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $argList = @('x','-y',"-o$OutDir")
    if ([string]::IsNullOrEmpty($Pass)) { $argList += '-p' } else { $argList += "-p$Pass" }
    $argList += $File
    $output = & $script:SevenZip @argList 2>&1
    $code = $LASTEXITCODE
    $text = ($output | Out-String)
    if ($code -eq 0 -or $code -eq 1) {
        if (Test-DirHasContent -Path $OutDir) { return 'ok' }
        return 'error'
    }
    # 先判"不是压缩包"（Cannot open 优先于 password 提示，避免误判）
    if ($text -match 'Cannot open the file as archive|Can not open|unsupported archive|not supported archive|No archives to extract|not recognized as archive|Unexpected end of archive') { return 'notarchive' }
    # 再判密码错误：精确匹配 7z 的密码错误特征（"Wrong password?" / 加密数据错误）
    if ($text -match 'Wrong password|Can not open encrypted archive|data error in encrypted|Cannot open encrypted') { return 'wrongpass' }
    return 'error'
}

# ---------- 引擎：WinRAR / UnRAR ----------
# 返回: ok | wrongpass | error
function Test-ArchiveWithRar {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $argList = @('x','-y','-o+')
    if ([string]::IsNullOrEmpty($Pass)) { $argList += '-p' } else { $argList += "-p$Pass" }
    $argList += $File
    $argList += "$OutDir\"
    $output = & $script:RarTool @argList 2>&1
    $code = $LASTEXITCODE
    $text = ($output | Out-String)
    if ($code -eq 0) {
        if (Test-DirHasContent -Path $OutDir) { return 'ok' }
        return 'error'
    }
    if ($text -match 'Incorrect password|Bad password|CRC failed in encrypted') { return 'wrongpass' }
    return 'error'
}

# ---------- 引擎：.NET ZipFile（仅 zip 内容，无密码） ----------
# 返回: ok | unsupported | notarchive
function Test-ArchiveWithDotNet {
    param([string]$File, [string]$OutDir, [string]$Pass)
    if (-not [string]::IsNullOrEmpty($Pass)) { return 'unsupported' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($File, $OutDir)
        return 'ok'
    } catch {
        if (Test-Path -LiteralPath $OutDir) {
            Remove-Item -LiteralPath $OutDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        return 'notarchive'
    }
}

# ---------- 解压进度显示 ----------
# 预估解压总量（7z l -slt 的 Size 键和），供进度条使用
function Get-EstimateSize {
    param([string]$File, [string]$Pass)
    if (-not $script:SevenZip) { return 0 }
    $out = if ([string]::IsNullOrEmpty($Pass)) { & $script:SevenZip l -slt $File -p 2>&1 | Out-String } else { & $script:SevenZip l -slt $File "-p$Pass" 2>&1 | Out-String }
    if ($LASTEXITCODE -ne 0) { return 0 }
    # -slt 格式每项含 "Size = N" 行，求和所有 Size（排除目录 0 值）
    $total = 0L
    foreach ($line in ($out -split [char]10)) {
        if ($line -match '^\s*Size\s*=\s*(\d+)\s*$') { $total += [long]$matches[1] }
    }
    return $total
}

# 启动进度监控：runspace 独立线程每 500ms 统计工作目录已写字节，刷新进度行（\r 覆盖）
# 返回 watcher 对象（含可调用 Stop 委托）
function Start-ProgressWatcher {
    param([string]$OutDir, [long]$TotalBytes)
    if ($TotalBytes -le 0) { return $null }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($dir, $total)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($true) {
            $bytes = 0L
            try {
                $di = New-Object System.IO.DirectoryInfo($dir)
                if ($di.Exists) {
                    foreach ($f in $di.GetFiles('*', [System.IO.SearchOption]::AllDirectories)) {
                        try { $bytes += $f.Length } catch { }
                    }
                }
            } catch { }
            if ($total -gt 0) {
                $pct = [Math]::Min(100.0, [Math]::Round($bytes * 100.0 / $total, 1))
                $el = $sw.Elapsed.TotalSeconds
                $speed = if ($el -gt 2) { [Math]::Round($bytes / $el / 1024 / 1024) } else { 0 }
                [System.Console]::Write(("`r  解压中: {0}%  ({1:0.00} GB / {2:0.00} GB, {3} MB/s)    " -f $pct, ($bytes/1GB), ($total/1GB), $speed))
            }
            Start-Sleep -Milliseconds 500
        }
    }).AddArgument($OutDir).AddArgument($TotalBytes)
    $async = $ps.BeginInvoke()
    return @{ Ps = $ps; Async = $async; Rs = $rs }
}

function Stop-ProgressWatcher {
    param($Watcher)
    if ($null -eq $Watcher) { return }
    try {
        $Watcher.Ps.Stop() | Out-Null
        [System.Console]::Write("`r" + (' ' * 80) + "`r")
    } catch { }
    try { $Watcher.Ps.Dispose() } catch { }
    try { $Watcher.Rs.Close() } catch { }
}

# ---------- 解压尝试链：按引擎优先级依次尝试 ----------
function Invoke-Extract {
    param([string]$File, [string]$OutDir, [string]$Pass)
    # 大文件解压显示进度：先预估总量，解压期间后台线程刷新百分比
    $watcher = $null
    if ([System.IO.Path]::GetExtension($File).ToLowerInvariant() -in @('.zip','.7z','.rar','.tar','.gz','.bz2','.xz','.001','.z01')) {
        $est = Get-EstimateSize -File $File -Pass $Pass
        if ($est -gt 100MB) { $watcher = Start-ProgressWatcher -OutDir $OutDir -TotalBytes $est }
    }
    try {
        $all = @()
        if ($script:SevenZip) {
            $r = Test-ArchiveWith7z -File $File -OutDir $OutDir -Pass $Pass
            $all += $r
            if ($r -eq 'ok') { return 'ok' }
        }
        if ($script:RarTool) {
            $r = Test-ArchiveWithRar -File $File -OutDir $OutDir -Pass $Pass
            $all += $r
            if ($r -eq 'ok') { return 'ok' }
        }
        $r = Test-ArchiveWithDotNet -File $File -OutDir $OutDir -Pass $Pass
        $all += $r
        if ($r -eq 'ok') { return 'ok' }
        # wrongpass 只在真实引擎（7z/rar）明确报密码错误时才算
        if ($all -contains 'wrongpass') { return 'wrongpass' }
        if ($all -contains 'notarchive') { return 'notarchive' }
        return 'error'
    } finally {
        Stop-ProgressWatcher $watcher
    }
}

# ---------- 工作目录：与源文件同盘（避免 18GB 大包跨盘复制到 C 盘 temp） ----------
function Get-WorkDir {
    param([string]$File)
    # 放在源文件所在目录的隐藏子目录里，解压后 robocopy 提升到目标目录（同盘，快）
    $base = Split-Path -Parent $File
    $dir = Join-Path $base ("AutoExtract_work_{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

# ---------- 改名重试：依次改成常见压缩包扩展名再解压 ----------
# 仅在没有真实引擎（7-Zip/WinRAR）时使用 —— 真实引擎按内容识别，扩展名写错也能直接解，
# 改名重试纯属多余；只有 .NET 内置引擎（部分依赖扩展名）才需要它兜底。
function Invoke-RenameRetry {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $exts = @('.zip','.7z','.rar','.tar','.gz','.bz2','.xz')
    $passFail = $false
    foreach ($ext in $exts) {
        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoExtract_{0}{1}" -f ([guid]::NewGuid().ToString('N')), $ext)
        try {
            Copy-Item -LiteralPath $File -Destination $tmp -Force
            $r = Invoke-Extract -File $tmp -OutDir $OutDir -Pass $Pass
        } finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
        if ($r -eq 'ok') { return @{ Result = 'ok'; Ext = $ext } }
        if ($r -eq 'wrongpass') { $passFail = $true }
    }
    return @{ Result = if ($passFail) { 'wrongpass' } else { 'notarchive' }; Ext = '' }
}

# ---------- 内容探测：用 7z 判断文件是否真的是压缩包（不依赖扩展名） ----------
# 返回: $true = 可被 7z 识别为压缩包；$false = 不是（或没有 7z）
function Test-IsArchiveByContent {
    param([string]$File)
    if (-not $script:SevenZip) { return $false }
    # 探测必须带密码参数：7z 对加密包不带 -p 会交互等待输入而卡住
    # 空密码用 -p，有密码用 -p{pass}
    if ([string]::IsNullOrEmpty($script:AutoPass)) {
        $output = & $script:SevenZip l $File -p 2>&1 | Out-String
    } else {
        $output = & $script:SevenZip l $File "-p$($script:AutoPass)" 2>&1 | Out-String
    }
    return ($LASTEXITCODE -eq 0)
}

# ---------- 嵌入式 zip 检测与截取解压 ----------
# 场景：文件头是媒体格式（mp4/jpg 等），但尾部嵌入了 zip 数据（资源站拼接包）。
# 7z 按文件头识别失败（Cannot open the file as archive），但 zip 以 EOCD (PK\x05\x06)
# 结尾、一定在文件尾部 128KB 内——找到 EOCD 反推近似起点，再在窗口内精确定位
# PK\x03\x04 本地文件头，从该偏移截取到文件尾生成副本，再用 7z 解压副本。
# 返回: @{Result=ok;Ext=...} | @{Result=wrongpass} | 'none'(未命中嵌入式场景)
function Try-EmbeddedZip {
    param([string]$File, [string]$OutDir, [string]$Pass)
    if (-not (Test-Path -LiteralPath $File)) { return 'none' }   # 文件已不存在（快照失效），不尝试
    $len = (Get-Item -LiteralPath $File).Length
    if ($len -lt 64) { return 'none' }  # 太小不可能包含 zip

    # 1) 读文件尾部 128KB 找 EOCD（从尾部往前找最后一个 PK\x05\x06）
    $tailLen = [Math]::Min(131072, $len)
    $fs = [System.IO.File]::OpenRead($File)
    try {
        $fs.Seek($len - $tailLen, [System.IO.SeekOrigin]::Begin) | Out-Null
        $tail = New-Object byte[] $tailLen
        $fs.Read($tail, 0, $tailLen) | Out-Null
    } finally { $fs.Close() }

    $eocdRel = -1
    for ($i = $tailLen - 22; $i -ge 0; $i--) {
        if ($tail[$i] -eq 0x50 -and $tail[$i+1] -eq 0x4B -and $tail[$i+2] -eq 0x05 -and $tail[$i+3] -eq 0x06) {
            $eocdRel = $i
            break
        }
    }
    if ($eocdRel -lt 0) { return 'none' }

    $eocdAbs = $len - $tailLen + $eocdRel
    $cdOffset   = [System.BitConverter]::ToUInt32($tail, $eocdRel + 16)
    $commentLen = [System.BitConverter]::ToUInt16($tail, $eocdRel + 20)
    $approx = $eocdAbs - $commentLen - $cdOffset   # 反推近似 zip 起点

    # 2) 在近似起点 ±128KB 窗口内精确搜索 PK\x03\x04（zip 本地文件头）
    $winStart = [Math]::Max(0, $approx - 131072)
    $winEnd   = [Math]::Min($len, $approx + 131072 + 4)
    $winLen   = $winEnd - $winStart
    if ($winLen -le 4) { return 'none' }
    $fs = [System.IO.File]::OpenRead($File)
    try {
        $fs.Seek($winStart, [System.IO.SeekOrigin]::Begin) | Out-Null
        $win = New-Object byte[] $winLen
        $fs.Read($win, 0, $winLen) | Out-Null
    } finally { $fs.Close() }

    $pkAbs = -1
    for ($i = 0; $i -le $winLen - 4; $i++) {
        if ($win[$i] -eq 0x50 -and $win[$i+1] -eq 0x4B -and $win[$i+2] -eq 0x03 -and $win[$i+3] -eq 0x04) {
            $pkAbs = $winStart + $i
            break
        }
    }
    if ($pkAbs -lt 0) { return 'none' }

    # 3) 从精确起点截取到文件尾生成副本，再用 7z 解压副本
    $tmpBase = Split-Path -Parent $File
    $tmpName = [System.IO.Path]::GetFileNameWithoutExtension($File) + '_embedded.zip'
    $tmpPath = Join-Path $tmpBase $tmpName
    $created = $false
    try {
        $fs = [System.IO.File]::OpenRead($File)
        $out = [System.IO.File]::Create($tmpPath)
        try {
            $fs.Seek($pkAbs, [System.IO.SeekOrigin]::Begin) | Out-Null
            $fs.CopyTo($out)
            $out.Flush()
            $created = $true
        } finally { $out.Close(); $fs.Close() }

        if (-not $script:SevenZip) { return 'none' }
        $argList = @('x','-y',"-o$OutDir")
        if ([string]::IsNullOrEmpty($Pass)) { $argList += '-p' } else { $argList += "-p$Pass" }
        $argList += $tmpPath
        $output = & $script:SevenZip @argList 2>&1
        $code = $LASTEXITCODE
        $text = ($output | Out-String)
        if ($code -eq 0 -or $code -eq 1) {
            if (Test-DirHasContent -Path $OutDir) { return @{ Result = 'ok'; Ext = '.zip' } }
            return 'none'
        }
        if ($text -match 'Wrong password|Can not open encrypted archive|data error in encrypted|Cannot open encrypted') {
            return @{ Result = 'wrongpass'; Ext = '.zip' }
        }
        return 'none'
    } finally {
        if ($created) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }
    }
}

# ---------- 分卷改名重试 ----------
# 场景：分卷压缩包的卷1 被资源站改坏名字（如 "xxx.7z(删掉" 丢了 .001 后缀，
# 或 "xxx.7z(删掉.001" 前缀与兄弟卷 "xxx.7z.002" 不一致），7z 无法识别为同一分卷。
# 检测到同目录存在后续卷（.002/.003/.z02/.part2 等）时，以兄弟卷名为标准前缀，
# 把卷1 复制成"标准前缀 + .001"再解压（7z 自动找同目录的后续卷组成完整分卷）。
# 返回: @{Result=ok;Ext=...} | @{Result=wrongpass} | 'none'(未命中分卷场景)
function Try-VolumeRename {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $baseDir = Split-Path -Parent $File
    $fileName = Split-Path -Leaf $File
    $ext = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()

    # 自身已是标准分卷名（.001 且能匹配兄弟卷）则不处理——但前缀不一致时仍需处理，
    # 所以这里只跳过完全无分卷特征的常规压缩包
    if ($ext -in @('.rar','.zip','.7z','.tar','.gz','.bz2','.xz')) { return 'none' }

    # 检查同目录是否存在后续分卷（排除自身：卷号大于当前文件的卷）
    # 分卷号允许带伪装后缀（如 .001.pdf / .002.mp3 / .part1.rar.jpg）：
    # 资源站常给每个分卷追加媒体扩展名，7z 按内容识别不出，必须去伪装后缀才能组卷
    $curNum = $null
    if ($fileName -match '\.(0\d{2})(\.[^.]+)?$') { $curNum = [int]$matches[1] }
    elseif ($fileName -match '\.z(\d{2})(\.[^.]+)?$') { $curNum = [int]$matches[1] }
    elseif ($fileName -match '\.part(\d+)\.rar(\.[^.]+)?$') { $curNum = [int]$matches[1] }

    $siblings = @(Get-ChildItem -LiteralPath $baseDir -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -ne $fileName -and $_.Name -match '(\.0\d{2}|\.z\d{2}|\.part\d+\.rar)(\.[^.]+)?$'
        })
    if ($siblings.Count -eq 0) { return 'none' }
    $sibling = $siblings[0]

    # 若当前文件本身已是 .001 且前缀与兄弟卷一致（标准命名），无需处理
    $sibNum = $null
    if ($sibling.Name -match '\.(0\d{2})(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    elseif ($sibling.Name -match '\.z(\d{2})(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    elseif ($sibling.Name -match '\.part(\d+)\.rar(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    $curPrefix = if ($null -ne $curNum) { $fileName -replace '\.(0\d{2}|z\d{2}|part\d+\.rar)(\.[^.]+)?$','' } else { $fileName }
    $sibPrefix = if ($null -ne $sibNum) { $sibling.Name -replace '\.(0\d{2}|z\d{2}|part\d+\.rar)(\.[^.]+)?$','' } else { $sibling.Name }

    # 当前文件是否带伪装后缀（分卷号之后还有额外扩展名）
    $curCamou = ($fileName -match '\.(0\d{2}|z\d{2}|part\d+\.rar)\.[^.]+$')

    # 卷1（.001）且前缀一致且无伪装后缀 → 标准分卷，7z 自己会处理，不需要改名
    if ($curNum -eq 1 -and $curPrefix -eq $sibPrefix -and -not $curCamou) { return 'none' }

    Write-Host "  检测到分卷压缩包（卷1 命名与后续卷不一致或带伪装后缀），尝试改名组卷重试…" -ForegroundColor DarkGray
    # 以兄弟卷名为标准，推导卷1 应叫的名字（同时去掉伪装后缀）：.002.pdf->.001 / .z02->.z01 / .part2.rar->.part1.rar
    $copyName = $sibling.Name -replace '\.(0\d{2})(\.[^.]+)?$', '.001' -replace '\.(z\d{2})(\.[^.]+)?$', '.z01' -replace '\.part\d+\.rar(\.[^.]+)?$', '.part1.rar'
    $copyPath = Join-Path $baseDir $copyName
    # 卷1 已叫标准名则无需处理
    if ($copyPath -eq $File) { return 'none' }
    # 避免覆盖同目录已有的其他文件
    if (Test-Path -LiteralPath $copyPath) {
        $copyPath = Join-Path $baseDir ("AutoExtract_vol1_" + $copyName)
    }

    $copied = @($copyPath)
    try {
        Copy-Item -LiteralPath $File -Destination $copyPath -Force
        # 兄弟卷带伪装后缀时也复制成标准名，7z 才能找到后续卷组成完整分卷
        foreach ($sv in $siblings) {
            if ($sv.FullName -eq $File) { continue }
            if ($sv.Name -notmatch '\.(0\d{2}|z\d{2}|part\d+\.rar)\.[^.]+$') { continue }  # 已是标准名
            $svCopyName = $sv.Name -replace '\.(0\d{2})(\.[^.]+)?$', '.$1' -replace '\.(z\d{2})(\.[^.]+)?$', '.z$1' -replace '\.part(\d+)\.rar(\.[^.]+)?$', '.part$1.rar'
            $svCopyPath = Join-Path $baseDir $svCopyName
            if (Test-Path -LiteralPath $svCopyPath) { continue }  # 已存在标准名（7z 能直接用）
            Copy-Item -LiteralPath $sv.FullName -Destination $svCopyPath -Force
            $copied += $svCopyPath
        }
        $r = Invoke-Extract -File $copyPath -OutDir $OutDir -Pass $Pass
    } finally {
        foreach ($cp in $copied) {
            Remove-Item -LiteralPath $cp -Force -ErrorAction SilentlyContinue
        }
    }
    if ($r -eq 'ok') { return @{ Result = 'ok'; Ext = [System.IO.Path]::GetExtension($copyName) } }
    if ($r -eq 'wrongpass') { return @{ Result = 'wrongpass'; Ext = '' } }
    return 'none'
}

# ---------- 批量：判断文件是否值得尝试解压 ----------
# 返回: $true = 尝试；$false = 跳过
function Should-TryFile {
    param([string]$Name, [bool]$ForceAll)
    $ext = [System.IO.Path]::GetExtension($Name).ToLowerInvariant()
    if ($ext -in $script:ArchiveExts -or $ext -eq '') { return $true }   # 已知压缩格式 / 无扩展名
    # 文件名含分卷特征（.001/.z01/.part1 等，允许带伪装后缀如 .001.pdf）→ 强制尝试：
    # 资源站常给分卷追加媒体扩展名，不能因扩展名伪装而跳过
    # 但分卷"后续卷"（.002+）跳过——卷1 处理时会同目录组卷，无需单独处理（避免误报失败）
    if ($Name -match '(\.0\d{2}|\.z\d{2}|\.part\d+\.rar)(\.[^.]+)?$') {
        $volNum = 0
        if ($Name -match '\.0(\d{2})(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        elseif ($Name -match '\.z(\d{2})(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        elseif ($Name -match '\.part(\d+)\.rar(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        if ($volNum -gt 1) { return $false }   # 后续卷，交给卷1 组卷处理
        return $true
    }
    if ($ForceAll) {
        # -All：对"可疑"扩展名做内容探测（能认出伪装扩展名的压缩包），
        # 但明确非压缩的媒体/文本/图片/代码/字体等类型不浪费探测
        $definitelyNot = @(
            '.txt','.md','.log','.json','.xml','.ini','.cfg','.csv','.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx',
            '.jpg','.jpeg','.png','.gif','.webp','.bmp','.ico','.svg','.heic',
            '.mp3','.wav','.flac','.m4a','.aac','.ogg',
            '.mp4','.mkv','.avi','.mov','.flv','.webm','.wmv','.ts','.m4v',
            '.exe','.dll','.msi','.sys',
            '.py','.js','.ts','.html','.css','.java','.c','.cpp','.go','.rs','.sh','.bat','.ps1',
            '.ttf','.otf','.woff','.woff2'
        )
        if ($ext -in $definitelyNot) { return $false }
        return $true
    }
    return $false
}

# ---------- 提升目录内容到目标目录（合并，含子目录） ----------
# 用 robocopy 的 /MOVE 合并移动，避免 Move-Item 目录合并问题；robocopy 退出码 0-7 均视为成功
function Move-DirContents {
    param([string]$Source, [string]$Dest)
    if (-not (Test-Path -LiteralPath $Source)) { return }
    New-Item -ItemType Directory -Path $Dest -Force | Out-Null
    $null = robocopy $Source $Dest /E /MOVE /NFL /NDL /NJH /NJS /NC /NS /NP
    # robocopy: 0-7 正常（0无文件,1成功复制,2多余,3,4不匹配,5,6,7），8+ 才是错误
    $rc = $LASTEXITCODE
    if ($rc -ge 8) {
        Write-Host "警告: robocopy 提升内容时出错 (代码 $rc)" -ForegroundColor Yellow
    }
    # 清理源和目标两侧的空目录（中间压缩包被删除后留下的空壳）
    Remove-EmptyDirs -Path $Source
    Remove-EmptyDirs -Path $Dest
}

# ---------- 递归删除空目录（不删根） ----------
function Remove-EmptyDirs {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    # 自底向上删除空目录
    $dirs = @(Get-ChildItem -LiteralPath $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object { $_.FullName.Length } -Descending)
    foreach ($d in $dirs) {
        if (@(Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------- 复制分卷兄弟卷到目标目录 ----------
# 分卷压缩包的所有卷必须在同一目录，7z 才能组成完整分卷解压。
# 嵌套场景：卷1 在临时 workDir，兄弟卷（.002/.003/.z02/part2）可能在原始目录（RootDir），
# 需要把兄弟卷复制到卷1 所在目录，否则 7z 把卷1 当独立包只解出残缺数据。
# 返回: 复制的兄弟卷数量
function Copy-SiblingVolumes {
    param([string]$File, [string]$RootDir)
    $fileDir = Split-Path -Parent $File
    $fileName = Split-Path -Leaf $File
    if (-not $RootDir -or $RootDir -eq $fileDir) { return 0 }

    # 识别分卷模式并构造"前缀"（不含卷号部分；分卷号允许带伪装后缀如 .001.pdf）
    $prefix = $null
    $volNum = $null
    if ($fileName -match '^(.*)\.(\d{3})(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    elseif ($fileName -match '^(.*)\.z(\d{2})(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    elseif ($fileName -match '^(.*)\.part(\d+)\.rar(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    if ($null -eq $prefix) { return 0 }

    # 从 RootDir 找同名前缀的后续卷并复制（避免动态正则，用简单字符串判断）
    $copied = 0
    $rootVols = @(Get-ChildItem -LiteralPath $RootDir -File -Force -ErrorAction SilentlyContinue)
    foreach ($rv in $rootVols) {
        if ($rv.Name -eq $fileName) { continue }
        if (-not $rv.Name.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        # 判断是否为后续分卷（前缀后紧跟卷号且卷号更大，允许伪装后缀）
        $suffix = $rv.Name.Substring($prefix.Length)
        $isSibling = $false
        if ($suffix -match '^\.(\d{3})(\.[^.]+)?$' -and [int]$matches[1] -gt $volNum) { $isSibling = $true }
        elseif ($suffix -match '^\.z(\d{2})(\.[^.]+)?$' -and [int]$matches[1] -gt $volNum) { $isSibling = $true }
        elseif ($suffix -match '^\.part(\d+)\.rar(\.[^.]+)?$' -and [int]$matches[1] -gt $volNum) { $isSibling = $true }
        if ($isSibling) {
            $dest = Join-Path $fileDir $rv.Name
            if (-not (Test-Path -LiteralPath $dest)) {
                Copy-Item -LiteralPath $rv.FullName -Destination $dest -Force
                $copied++
            }
        }
    }
    return $copied
}

# ---------- 嵌套解压：扫描目录中解压产物里的压缩包并递归解压 ----------
# 产物中的中间压缩包解压后自动删除；最终内容由调用方提升到 RootDir
# 返回: 成功处理的压缩包数量
function Expand-Nested {
    param([string]$Dir, [int]$Depth, [string]$RootDir)
    $found = 0
    # 扫描根 $Dir 本身可能是 AutoExtract_work_xxx（上层工作目录），必须扫描它的文件；
    # 但 $Dir 内部的嵌套层临时目录（子目录名为 AutoExtract_work_*）会被排除——
    # 它们在处理过程中被创建/删除，扫描到会对已删除路径报错刷屏，且把临时产物当嵌套包误处理。
    $subDirs = @(Get-ChildItem -LiteralPath $Dir -Directory -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '^AutoExtract_work_' })
    $items = @(Get-ChildItem -LiteralPath $Dir -File -Force -ErrorAction SilentlyContinue)
    foreach ($sd in $subDirs) {
        $items += @(Get-ChildItem -LiteralPath $sd.FullName -File -Force -ErrorAction SilentlyContinue)
    }
    foreach ($it in $items) {
        if ($script:StopRequested) { break }
        # 快照里的文件可能已被前一次嵌套处理删除（临时目录随处理完成清理），跳过即可
        if (-not (Test-Path -LiteralPath $it.FullName)) { continue }
        $ext = [System.IO.Path]::GetExtension($it.Name).ToLowerInvariant()

        # 分卷卷1：先把兄弟卷复制到同一目录（workDir 缺卷时从 RootDir 复制），
        # 保证 7z 能组成完整分卷。注意：只复制到 workDir，绝不删除原始文件。
        # $null = 捕获返回值，避免泄漏进函数输出流（否则 $found 变数组）
        $null = Copy-SiblingVolumes -File $it.FullName -RootDir $RootDir
        # 分卷卷1（.001 等）即使 Test-IsArchiveByContent 失败也要尝试：
        # 前缀被改坏时 7z 探测会失败，但 Try-VolumeRename 能通过改名组卷
        $extL = [System.IO.Path]::GetExtension($it.Name).ToLowerInvariant()
        $isVolFirst = ($extL -in @('.001','.z01','.part1.rar')) -or ($it.Name -match '\.(001|z01|part1\.rar)(\.[^.]+)?$')
        if ($isVolFirst -or (Test-IsArchiveByContent -File $it.FullName)) {
            # 超大嵌套包提示（>2GB），避免无感知长时间等待
            if ($it.Length -gt 2GB) {
                Write-Host ("  嵌套包较大 ({0} GB)，解压可能较慢，请耐心等待…" -f [math]::Round($it.Length/1GB,1)) -ForegroundColor Yellow
            }
            $subOk = Expand-OneFile -File $it.FullName -Depth ($Depth + 1) -RootDir $RootDir
            if ($subOk) {
                $found++
                # 删除中间压缩包（嵌套层解压成功后立即清理）——只删 workDir/临时目录里的副本
                # 分卷：卷1 处理成功后，同目录的所有分卷卷都是中间产物，一并清理。
                # 注意：卷1 前缀可能被改坏（code.zip(删掉.001 vs code.zip.002），
                # 所以清理时匹配"同根分卷"（任一 .00N/.zNN/.partN 后缀的相邻文件）
                $itDir = Split-Path -Parent $it.FullName
                $removed = @($it.FullName)
                $allVols = @(Get-ChildItem -LiteralPath $itDir -File -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '(\.0\d{2}|\.z\d{2}|\.part\d+\.rar)(\.[^.]+)?$' })
                if ($allVols.Count -gt 0) {
                    foreach ($sv in $allVols) {
                        if ($sv.FullName -notin $removed) { $removed += $sv.FullName }
                    }
                }
                foreach ($rf in $removed) {
                    Remove-Item -LiteralPath $rf -Force -ErrorAction SilentlyContinue
                    Write-Host ("  已删除中间压缩包: {0}" -f (Split-Path -Leaf $rf)) -ForegroundColor DarkGray
                }
            }
        }
    }
    return $found
}

# ---------- 处理单个文件 ----------
function Expand-OneFile {
    param([string]$File, [int]$Depth = 1, [string]$RootDir = '')
    if ($script:StopRequested) { return $false }
    $indent = '  ' * ([Math]::Max(0, $Depth - 1))
    $fileName = Split-Path -Leaf $File
    $baseDir  = Split-Path -Parent $File

    # RootDir = 最终输出目录（原始压缩包所在层）。顶层未指定时取文件所在目录
    if ([string]::IsNullOrEmpty($RootDir)) { $RootDir = $baseDir }

    # 解压到工作目录（与源文件同盘），成功后内容提升到 RootDir（保证中间失败不污染目标目录）
    $workDir = Get-WorkDir -File $File

    Write-Host ("{0}处理: {1}" -f $indent, $fileName) -ForegroundColor Cyan

    # 1) 用当前密码直接解压（7z 按内容识别格式，扩展名错也能解）
    # 1) 分卷卷1 特殊处理：直接解压会被 7z 当独立包（只解出残缺数据）而误判成功。
    #    必须先 Try-VolumeRename（重命名使前缀与兄弟卷一致），7z 才能组成完整分卷。
    $ext0 = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()
    $isVolFirst = ($ext0 -in @('.001','.z01','.part1.rar')) -or ($fileName -match '\.(001|z01|part1\.rar)(\.[^.]+)?$')
    $r = $null
    $volHandled = $false
    if ($isVolFirst) {
        $sibChk = Get-ChildItem -LiteralPath $baseDir -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne $fileName -and $_.Name -match '(\.0\d{2}|\.z\d{2}|\.part\d+\.rar)(\.[^.]+)?$' } |
            Select-Object -First 1
        if ($sibChk) {
            Write-Host ("{0}  分卷卷1，尝试组卷解压…" -f $indent) -ForegroundColor DarkGray
            $vr = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($vr -is [hashtable] -and $vr.Result -eq 'ok') {
                $nested = 0
                if ($Depth -lt $MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                if ($Depth -eq 1) {
                    Write-Host ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) -ForegroundColor Green
                } else {
                    Write-Host ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) -ForegroundColor Green
                }
                if ($nested -gt 0) {
                    Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
                }
                return $true
            }
            if ($vr -is [hashtable] -and $vr.Result -eq 'wrongpass') {
                $r = 'wrongpass'
                $volHandled = $true
            }
        }
        if (-not $volHandled) {
            $r = Invoke-Extract -File $File -OutDir $workDir -Pass $script:AutoPass
        }
    } else {
        $r = Invoke-Extract -File $File -OutDir $workDir -Pass $script:AutoPass
    }

    # 2) 常规解压成功
    if ($r -eq 'ok') {
        $nested = 0
        if ($Depth -lt $MaxDepth) {
            $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
        }
        # 3) 剩余内容（含普通文件）提升到 RootDir（原始压缩包所在层）
        Move-DirContents -Source $workDir -Dest $RootDir
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue

        if ($Depth -eq 1) {
            Write-Host ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) -ForegroundColor Green
        } else {
            Write-Host ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) -ForegroundColor Green
        }
        if ($nested -gt 0) {
            Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
        }
        return $true
    }

    # 3) 密码问题 → 提示输入
    if ($r -eq 'wrongpass') {
        # 顶层且给了死密码参数 → 直接跳过；嵌套层（Depth>1）密码可能不同，始终给重输机会
        if ($script:PassFromParam -and $Depth -eq 1) {
            Write-Host ("{0}  密码错误（参数密码无效），跳过" -f $indent) -ForegroundColor Red
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
        if ($Depth -gt 1) {
            Write-Host ("{0}  该层密码与之前不同，请重新输入" -f $indent) -ForegroundColor Yellow
        }
        $script:AutoPass = Read-Secret -Prompt "  该文件需要密码"
        $r = Invoke-Extract -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($r -eq 'ok') {
            $nested = 0
            if ($Depth -lt $MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            if ($Depth -eq 1) {
                Write-Host ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) -ForegroundColor Green
            } else {
                Write-Host ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) -ForegroundColor Green
            }
            if ($nested -gt 0) {
                Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
            }
            return $true
        }
    }

    # 4) 不是标准压缩包 / 无法识别
    #    有 7z/rar 真实引擎时：7z 已按内容识别过，改扩展名无意义，直接判定失败
    #    仅有 .NET 引擎时：改名重试兜底（.NET 依赖扩展名判断）
    if ($r -eq 'notarchive' -or $r -eq 'error') {
        # 4a) 分卷改名重试：卷1 缺 .001 后缀的分卷包（同目录有 .002 等）
        $volR = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($volR -eq 'ok' -or ($volR -is [hashtable] -and $volR.Result -eq 'ok')) {
            $volExt = if ($volR -is [hashtable]) { $volR.Ext } else { '.001' }
            $nested = 0
            if ($Depth -lt $MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host ("{0}  OK 识别为分卷压缩包（{1}），解压完成 → {2}" -f $indent, $volExt, $RootDir) -ForegroundColor Green
            if ($nested -gt 0) {
                Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
            }
            return $true
        }
        if ($volR -is [hashtable] -and $volR.Result -eq 'wrongpass' -and (-not $script:PassFromParam -or $Depth -gt 1)) {
            if ($Depth -gt 1) {
                Write-Host ("{0}  该层密码与之前不同，请重新输入" -f $indent) -ForegroundColor Yellow
            }
            $script:AutoPass = Read-Secret -Prompt "  该文件需要密码"
            $volR2 = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($volR2 -is [hashtable] -and $volR2.Result -eq 'ok') {
                $nested = 0
                if ($Depth -lt $MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host ("{0}  OK 识别为分卷压缩包（{1}），解压完成 → {2}" -f $indent, $volR2.Ext, $RootDir) -ForegroundColor Green
                return $true
            }
        }

        $needRenameRetry = (-not $script:SevenZip -and -not $script:RarTool)

        # 4b) 嵌入式 zip 检测：文件头是媒体格式但尾部嵌入了 zip 数据（资源站拼接包）
        #     7z 按文件头识别失败，但 WinRAR 能通过尾部 EOCD 定位打开——我们截取 zip 段再用 7z 解压
        $embR = Try-EmbeddedZip -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($embR -is [hashtable] -and $embR.Result -eq 'ok') {
            Write-Host ("{0}  检测到文件尾部嵌入压缩数据，已截取解压…" -f $indent) -ForegroundColor DarkGray
            $nested = 0
            if ($Depth -lt $MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host ("{0}  OK 识别为嵌入式压缩包，解压完成 → {1}" -f $indent, $RootDir) -ForegroundColor Green
            if ($nested -gt 0) {
                Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
            }
            return $true
        }
        if ($embR -is [hashtable] -and $embR.Result -eq 'wrongpass') {
            if ($script:PassFromParam -and $Depth -eq 1) {
                Write-Host ("{0}  密码错误（参数密码无效），跳过" -f $indent) -ForegroundColor Red
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                return $false
            }
            if ($Depth -gt 1) {
                Write-Host ("{0}  该层密码与之前不同，请重新输入" -f $indent) -ForegroundColor Yellow
            }
            $script:AutoPass = Read-Secret -Prompt "  该文件需要密码"
            $embR2 = Try-EmbeddedZip -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($embR2 -is [hashtable] -and $embR2.Result -eq 'ok') {
                Write-Host ("{0}  检测到文件尾部嵌入压缩数据，已截取解压…" -f $indent) -ForegroundColor DarkGray
                $nested = 0
                if ($Depth -lt $MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host ("{0}  OK 识别为嵌入式压缩包，解压完成 → {1}" -f $indent, $RootDir) -ForegroundColor Green
                if ($nested -gt 0) {
                    Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
                }
                return $true
            }
        }

        # 若分卷识别出是加密包但密码不对（且不给重输），明确提示密码错误而非"非压缩包"
        $volWasWrongPass = ($volR -is [hashtable] -and $volR.Result -eq 'wrongpass')
        if ($volWasWrongPass -and $script:PassFromParam) {
            Write-Host ("{0}  密码错误（参数密码无效），跳过" -f $indent) -ForegroundColor Red
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
        if ($needRenameRetry) {
            Write-Host ("{0}  不是标准压缩包，尝试改扩展名重试…" -f $indent) -ForegroundColor DarkGray
            $rr = Invoke-RenameRetry -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($rr.Result -eq 'ok') {
                $nested = 0
                if ($Depth -lt $MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host ("{0}  OK 识别为 {1} 格式，解压完成 → {2}" -f $indent, $rr.Ext, $RootDir) -ForegroundColor Green
                if ($nested -gt 0) {
                    Write-Host ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) -ForegroundColor DarkGray
                }
                return $true
            }
            if ($rr.Result -eq 'wrongpass' -and (-not $script:PassFromParam -or $Depth -gt 1)) {
                if ($Depth -gt 1) {
                    Write-Host ("{0}  该层密码与之前不同，请重新输入" -f $indent) -ForegroundColor Yellow
                }
                $script:AutoPass = Read-Secret -Prompt "  该文件需要密码"
                $rr2 = Invoke-RenameRetry -File $File -OutDir $workDir -Pass $script:AutoPass
                if ($rr2.Result -eq 'ok') {
                    $nested = 0
                    if ($Depth -lt $MaxDepth) {
                        $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                    }
                    Move-DirContents -Source $workDir -Dest $RootDir
                    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host ("{0}  OK 识别为 {1} 格式，解压完成 → {2}" -f $indent, $rr2.Ext, $RootDir) -ForegroundColor Green
                    return $true
                }
            }
        } else {
            Write-Host ("{0}  非压缩包文件（已按内容识别确认）" -f $indent) -ForegroundColor DarkGray
        }
        Write-Host ("{0}  无法解压：不是可识别的压缩包，或密码错误" -f $indent) -ForegroundColor Red
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    # 其他错误（含密码仍然错误）
    Write-Host ("{0}  解压失败：{1}" -f $indent, $r) -ForegroundColor Red
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    return $false
}

# ---------- 主流程 ----------
Init-Engines

if ($script:SevenZip) {
    Write-Host "解压引擎: 7-Zip ($script:SevenZip)" -ForegroundColor DarkGray
} elseif ($script:RarTool) {
    Write-Host "解压引擎: $script:RarTool" -ForegroundColor DarkGray
} else {
    Write-Host "警告: 未找到 7-Zip / WinRAR，仅能用 .NET 内置解 zip（不支持密码）" -ForegroundColor Yellow
}

$script:AutoPass = $Password

if (-not $Paths -or $Paths.Count -eq 0) {
    $input2 = Read-Host "请把文件/文件夹拖入此窗口，或输入路径"
    if ([string]::IsNullOrWhiteSpace($input2)) {
        Write-Host "未提供任何目标，退出。" -ForegroundColor Yellow
        exit 1
    }
    $Paths = @($input2)
}

$okCount = 0; $failCount = 0
foreach ($p in $Paths) {
    $p = $p.Trim('"')
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host "路径不存在: $p" -ForegroundColor Red
        $failCount++
        continue
    }
    $item = Get-Item -LiteralPath $p
    if ($item.PSIsContainer) {
        # 批量模式：扫描文件夹中的压缩包（顶层或递归子文件夹）
        if ($Recurse) {
            $files = @(Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue)
        } else {
            $files = @(Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue)
        }
        foreach ($f in $files) {
            if ($script:StopRequested) { break }
            # 默认按压缩扩展名/无扩展名筛选；-All 时对所有文件内容探测
            if (Should-TryFile -Name $f.Name -ForceAll $All) {
                if (Expand-OneFile -File $f.FullName) { $okCount++ } else { $failCount++ }
            }
        }
    } else {
        if (Expand-OneFile -File $item.FullName) { $okCount++ } else { $failCount++ }
    }
}

Write-Host ""
Write-Host ("完成：成功 {0} 个，失败 {1} 个" -f $okCount, $failCount) -ForegroundColor Green

# 完成提示音：成功播放上行双音，全部失败播放低音
try {
    if ($okCount -gt 0 -and $failCount -eq 0) {
        [Console]::Beep(880, 150); Start-Sleep -Milliseconds 80; [Console]::Beep(1320, 300)   # 成功上行双音
    } elseif ($okCount -gt 0) {
        [Console]::Beep(660, 150); Start-Sleep -Milliseconds 80; [Console]::Beep(990, 250)     # 部分成功
    } else {
        [Console]::Beep(220, 400)                                                              # 全部失败低音
    }
} catch {}
