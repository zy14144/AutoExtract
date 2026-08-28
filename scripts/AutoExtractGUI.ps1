#requires -Version 3.0
<#
.SYNOPSIS
    智能解压工具（图形界面版）：输入密码后自动解压压缩包；不是压缩包则按内容识别/改名重试；支持多层嵌套。

.DESCRIPTION
    与命令行版 AutoExtract.ps1 功能一致，提供 WinForms 图形界面：
      - 拖放文件/文件夹或按钮添加，支持批量
      - 密码输入框（可选，留空则先试空密码，加密包才弹窗输入）
      - 多层嵌套解压（可设最大层数，0 = 禁用）
      - 彩色日志输出，后台线程处理不卡界面

    引擎：7-Zip > WinRAR/UnRAR > .NET 内置(仅zip)。7-Zip/WinRAR 按内容识别格式。

    用法：直接运行 AutoExtractGUI.ps1（或双击 AutoExtractGUI.bat）。
    也可把文件/文件夹拖放到 exe 上：自动加入列表并开始解压。
#>

# ---------- DPI 感知（2K/4K 高分屏字体清晰） ----------
# WinForms 默认不支持高分屏 DPI 缩放，在 2K/4K 屏上会模糊。
# PS 5.1 用 Win32 API 声明进程 DPI 感知，并在窗体上启用 Dpi 缩放。
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DpiAware {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
[void][DpiAware]::SetProcessDPIAware()

# ---------- 启动路径：拖放到 exe/脚本上的文件 ----------
# 不用 param 块：ps2exe 打包的 exe 对 param 传递兼容性差，用 $args 最可靠
$StartupPaths = @()
foreach ($a in $args) {
    $s = [string]$a
    if ($s -and $s -notmatch '^-') { $StartupPaths += $s.Trim('"') }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- 全局状态 ----------
$script:SevenZip = $null
$script:RarTool   = $null
$script:AutoPass  = ''
$script:PassFromParam = $false
$script:StopRequested = $false
$script:Running  = $false
$script:OkCount  = 0
$script:FailCount = 0
$script:MaxDepth = 5
$script:Recurse  = $false
$script:ForceAll = $false
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
    $script:RarTool = Find-Tool -Names @('UnRAR.exe','WinRAR.exe','rar.exe') -KnownPaths @(
        "$env:ProgramFiles\WinRAR\UnRAR.exe",
        "${env:ProgramFiles(x86)}\WinRAR\UnRAR.exe",
        "$env:ProgramFiles\WinRAR\WinRAR.exe",
        "${env:ProgramFiles(x86)}\WinRAR\WinRAR.exe"
    )
}

# ---------- 日志（跨线程安全） ----------
function Add-Log {
    param([string]$Text, [string]$Color = 'Black')
    if ($script:Form.InvokeRequired) {
        $script:Form.Invoke([Action[string,string]]{ param($t,$c) Add-Log -Text $t -Color $c }, @($Text, $Color)) | Out-Null
        return
    }
    try {
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        # Log colors tuned for the light background (light tones are darkened a notch)
        $LogColor = switch ($Color) {
            'Yellow' { [System.Drawing.Color]::FromArgb(176, 137, 0) }
            'Orange' { [System.Drawing.Color]::FromArgb(198, 102, 0) }
            'Cyan'   { [System.Drawing.Color]::FromArgb(0, 130, 150) }
            'White'  { [System.Drawing.Color]::FromArgb(90, 90, 90) }
            default  { [System.Drawing.Color]::FromName($Color) }
        }
        $script:LogBox.SelectionColor = $LogColor
        $script:LogBox.AppendText($Text + "`r`n")
        $script:LogBox.ScrollToCaret()
    } catch {}
}

function Update-Status {
    param([string]$Text)
    if ($script:Form.InvokeRequired) {
        $script:Form.Invoke([Action[string]]{ param($t) Update-Status -Text $t }, @($Text)) | Out-Null
        return
    }
    $script:LblStatus.Text = $Text
}

# ---------- 密码弹窗（在后台线程中显示，独立消息循环） ----------
function Show-PasswordDialog {
    param([string]$Prompt)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '需要解压密码'
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.ClientSize = New-Object System.Drawing.Size(360, 120)
    $f.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    # DPI 缩放（配合进程级 DPI 感知，高分屏清晰）
    $f.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $f.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Prompt
    $lbl.SetBounds(15, 12, 330, 22)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.SetBounds(15, 42, 330, 24)
    $tb.UseSystemPasswordChar = $true

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '确定'
    $btnOk.SetBounds(100, 78, 75, 28)
    $btnOk.DialogResult = 'OK'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.SetBounds(185, 78, 75, 28)
    $btnCancel.DialogResult = 'Cancel'

    $f.Controls.AddRange(@($lbl, $tb, $btnOk, $btnCancel))
    $f.AcceptButton = $btnOk
    $f.CancelButton = $btnCancel
    $f.Add_Shown({ $tb.Focus() })
    # 关键：先泵一次消息循环，确保父窗体完全就绪再弹模态框，
    # 否则在 Add_Shown 自动启动场景下模态对话框可能卡死
    [System.Windows.Forms.Application]::DoEvents()
    $result = $f.ShowDialog()
    if ($result -eq 'OK') { return $tb.Text }
    return $null
}

# ---------- 输出目录是否有实际文件（非空目录壳） ----------
function Test-DirHasContent {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    return @(Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue).Count -gt 0
}

# ---------- 预估解压总量（7z l -slt 的 Size 键和），供进度显示 ----------
function Get-EstimateSize {
    param([string]$File, [string]$Pass)
    if (-not $script:SevenZip) { return 0 }
    $out = if ([string]::IsNullOrEmpty($Pass)) { & $script:SevenZip l -slt $File -p 2>&1 | Out-String } else { & $script:SevenZip l -slt $File "-p$Pass" 2>&1 | Out-String }
    if ($LASTEXITCODE -ne 0) { return 0 }
    $total = 0L
    foreach ($line in ($out -split [char]10)) {
        if ($line -match '^\s*Size\s*=\s*(\d+)\s*$') { $total += [long]$matches[1] }
    }
    return $total
}

# 统计目录已写字节（供进度计算）
function Get-DirBytes {
    param([string]$Dir)
    $bytes = 0L
    try {
        $di = New-Object System.IO.DirectoryInfo($Dir)
        if ($di.Exists) {
            foreach ($f in $di.GetFiles('*', [System.IO.SearchOption]::AllDirectories)) {
                try { $bytes += $f.Length } catch { }
            }
        }
    } catch { }
    return $bytes
}

# 启动进度监控：WinForms Timer（UI 线程回调，无 runspace/Invoke 死锁风险），
# 每 500ms 统计已写字节更新状态栏。ps2exe 下 runspace + $form.Invoke 会死锁（已实测）。
# 返回 watcher 对象（含可调用 Stop 委托）
function Start-ProgressWatcher {
    param([string]$OutDir, [long]$TotalBytes)
    if ($TotalBytes -le 0) { return $null }
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $timer.Add_Tick({
        try {
            $bytes = Get-DirBytes -Dir $OutDir
            if ($TotalBytes -gt 0) {
                $pct = [Math]::Min(100.0, [Math]::Round($bytes * 100.0 / $TotalBytes, 1))
                $el = $sw.Elapsed.TotalSeconds
                $speed = if ($el -gt 2) { [Math]::Round($bytes / $el / 1024 / 1024) } else { 0 }
                Update-Status ("解压中: {0}%  ({1:0.00} GB / {2:0.00} GB, {3} MB/s)" -f $pct, ($bytes/1GB), ($TotalBytes/1GB), $speed)
            }
        } catch { }
    })
    $timer.Start()
    return @{ Timer = $timer; Sw = $sw }
}

function Stop-ProgressWatcher {
    param($Watcher)
    if ($null -eq $Watcher) { return }
    try { $Watcher.Timer.Stop() } catch { }
    try { $Watcher.Timer.Dispose() } catch { }
    Update-Status '就绪'
}

# ---------- 引擎：7-Zip（同步 + 后台进度轮询） ----------
# 处理链本身运行在 UI 线程（BeginInvoke 派发），7z 必须同步执行——
# 异步 + DoEvents 会在处理链中途重入消息泵导致卡死（与 ShowDialog 嵌套冲突同类问题）。
# 进度由独立 runspace 线程轮询目录字节数，通过 Invoke 更新状态栏（线程安全）。
function Test-ArchiveWith7z {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $argList = @('x','-y',"-o$OutDir")
    if ([string]::IsNullOrEmpty($Pass)) { $argList += '-p' } else { $argList += "-p$Pass" }
    $argList += $File
    # 预估总量：仅在文件较大时显示进度（7z l 开销小，可接受）
    $est = Get-EstimateSize -File $File -Pass $Pass
    $watcher = $null
    if ($est -gt 50MB) { $watcher = Start-ProgressWatcher -OutDir $OutDir -TotalBytes $est }
    try {
        $output = & $script:SevenZip @argList 2>&1
        $code = $LASTEXITCODE
        $text = ($output | Out-String)
        if ($code -eq 0 -or $code -eq 1) {
            if (Test-DirHasContent -Path $OutDir) { return 'ok' }
            return 'error'
        }
        # 先判"不是压缩包"（Cannot open / Unexpected end 优先，避免 CRC 误判为密码错误）
        if ($text -match 'Cannot open the file as archive|Can not open|unsupported archive|not supported archive|No archives to extract|not recognized as archive|Unexpected end of archive') { return 'notarchive' }
        # 再判密码错误：精确匹配 7z 的密码错误特征
        if ($text -match 'Wrong password|Can not open encrypted archive|data error in encrypted|Cannot open encrypted') { return 'wrongpass' }
        return 'error'
    } finally {
        Stop-ProgressWatcher $watcher
    }
}

# ---------- 引擎：WinRAR / UnRAR（同步） ----------
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

# ---------- 引擎：.NET ZipFile ----------
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

# ---------- 解压尝试链 ----------
function Invoke-Extract {
    param([string]$File, [string]$OutDir, [string]$Pass)
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
    if ($all -contains 'wrongpass') { return 'wrongpass' }
    if ($all -contains 'notarchive') { return 'notarchive' }
    return 'error'
}

# ---------- 改名重试（仅无 7z/rar 时） ----------
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

# ---------- 内容探测 ----------
function Test-IsArchiveByContent {
    param([string]$File)
    if (-not $script:SevenZip) { return $false }
    # 探测必须带密码参数：7z 对加密包不带 -p 会交互等待输入而卡住
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
# 场景：分卷压缩包的卷1 被资源站改坏名字（如 "xxx.7z(删掉" 丢了 .001 后缀），
# 7z 无法直接识别。检测到同目录存在后续卷（.002/.003/.z02/.part2 等）时，
# 从后续卷名推导卷1 的正确名字并复制过去再解压（7z 会自动找同目录的后续卷）。
# 返回: @{Result=ok;Ext=...} | @{Result=wrongpass} | 'none'(未命中分卷场景)
# ---------- 复制分卷兄弟卷到目标目录 ----------
function Copy-SiblingVolumes {
    param([string]$File, [string]$RootDir)
    $fileDir = Split-Path -Parent $File
    $fileName = Split-Path -Leaf $File
    if (-not $RootDir -or $RootDir -eq $fileDir) { return 0 }

    $prefix = $null
    $volNum = $null
    if ($fileName -match '^(.*)\.(\d{3})(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    elseif ($fileName -match '^(.*)\.z(\d{2})(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    elseif ($fileName -match '^(.*)\.part(\d+)\.rar(\.[^.]+)?$') { $prefix = $matches[1]; $volNum = [int]$matches[2] }
    if ($null -eq $prefix) { return 0 }

    $copied = 0
    $rootVols = @(Get-ChildItem -LiteralPath $RootDir -File -Force -ErrorAction SilentlyContinue)
    foreach ($rv in $rootVols) {
        if ($rv.Name -eq $fileName) { continue }
        if (-not $rv.Name.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
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

function Try-VolumeRename {
    param([string]$File, [string]$OutDir, [string]$Pass)
    $baseDir = Split-Path -Parent $File
    $fileName = Split-Path -Leaf $File
    $ext = [System.IO.Path]::GetExtension($fileName).ToLowerInvariant()

    if ($ext -in @('.rar','.zip','.7z','.tar','.gz','.bz2','.xz')) { return 'none' }

    # 检查同目录是否存在后续分卷（排除自身；分卷号允许带伪装后缀如 .001.pdf）
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

    $sibNum = $null
    if ($sibling.Name -match '\.(0\d{2})(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    elseif ($sibling.Name -match '\.z(\d{2})(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    elseif ($sibling.Name -match '\.part(\d+)\.rar(\.[^.]+)?$') { $sibNum = [int]$matches[1] }
    $curPrefix = if ($null -ne $curNum) { $fileName -replace '\.(0\d{2}|z\d{2}|part\d+\.rar)(\.[^.]+)?$','' } else { $fileName }
    $sibPrefix = if ($null -ne $sibNum) { $sibling.Name -replace '\.(0\d{2}|z\d{2}|part\d+\.rar)(\.[^.]+)?$','' } else { $sibling.Name }

    # 当前文件是否带伪装后缀（分卷号之后还有额外扩展名）
    $curCamou = ($fileName -match '\.(0\d{2}|z\d{2}|part\d+\.rar)\.[^.]+$')

    # 卷1 且前缀一致且无伪装 → 标准分卷，无需改名
    if ($curNum -eq 1 -and $curPrefix -eq $sibPrefix -and -not $curCamou) { return 'none' }

    Add-Log '  检测到分卷压缩包（卷1 命名与后续卷不一致或带伪装后缀），尝试改名组卷重试…' 'DarkGray'
    $sibName = $sibling.Name
    $copyName = $sibName -replace '\.(0\d{2})(\.[^.]+)?$', '.001' -replace '\.(z\d{2})(\.[^.]+)?$', '.z01' -replace '\.part\d+\.rar(\.[^.]+)?$', '.part1.rar'
    $copyPath = Join-Path $baseDir $copyName
    if ($copyPath -eq $File) { return 'none' }
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
function Should-TryFile {
    param([string]$Name, [bool]$ForceAll)
    $ext = [System.IO.Path]::GetExtension($Name).ToLowerInvariant()
    if ($ext -in $script:ArchiveExts -or $ext -eq '') { return $true }
    # 文件名含分卷特征（.001/.z01/.part1 等，允许带伪装后缀如 .001.pdf）→ 强制尝试：
    # 资源站常给分卷追加媒体扩展名，不能因扩展名伪装而跳过。
    # 分卷"后续卷"（.002+）跳过——卷1 处理时会组卷，无需单独处理（避免误报失败）
    if ($Name -match '(\.0\d{2}|\.z\d{2}|\.part\d+\.rar)(\.[^.]+)?$') {
        $volNum = 0
        if ($Name -match '\.0(\d{2})(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        elseif ($Name -match '\.z(\d{2})(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        elseif ($Name -match '\.part(\d+)\.rar(\.[^.]+)?$') { $volNum = [int]$matches[1] }
        if ($volNum -gt 1) { return $false }   # 后续卷，交给卷1 组卷处理
        return $true
    }
    if ($ForceAll) {
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

# ---------- 工作目录：与源文件同盘（避免大包跨盘复制到 C 盘 temp） ----------
function Get-WorkDir {
    param([string]$File)
    $base = Split-Path -Parent $File
    $dir = Join-Path $base ("AutoExtract_work_{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

# ---------- 提升目录内容到目标目录（合并，含子目录） ----------
function Move-DirContents {
    param([string]$Source, [string]$Dest)
    if (-not (Test-Path -LiteralPath $Source)) { return }
    New-Item -ItemType Directory -Path $Dest -Force | Out-Null
    $null = robocopy $Source $Dest /E /MOVE /NFL /NDL /NJH /NJS /NC /NS /NP
    $rc = $LASTEXITCODE
    if ($rc -ge 8) {
        Add-Log ("警告: robocopy 提升内容时出错 (代码 {0})" -f $rc) 'Orange'
    }
    Remove-EmptyDirs -Path $Source
    Remove-EmptyDirs -Path $Dest
}

# ---------- 递归删除空目录（不删根） ----------
function Remove-EmptyDirs {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $dirs = @(Get-ChildItem -LiteralPath $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object { $_.FullName.Length } -Descending)
    foreach ($d in $dirs) {
        if (@(Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------- 嵌套解压 ----------
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
        # 分卷卷1：先把兄弟卷复制到同一目录（workDir 缺卷时从 RootDir 复制），
        # 保证 7z 能组成完整分卷。注意：只复制到 workDir，绝不删除原始文件。
        $null = Copy-SiblingVolumes -File $it.FullName -RootDir $RootDir
        # 分卷卷1（.001 等）即使 Test-IsArchiveByContent 失败也要尝试：
        # 前缀被改坏时 7z 探测会失败，但 Try-VolumeRename 能通过改名组卷
        $extL = [System.IO.Path]::GetExtension($it.Name).ToLowerInvariant()
        $isVolFirst = ($extL -in @('.001','.z01','.part1.rar')) -or ($it.Name -match '\.(001|z01|part1\.rar)(\.[^.]+)?$')
        if ($isVolFirst -or (Test-IsArchiveByContent -File $it.FullName)) {
            if ($it.Length -gt 2GB) {
                Add-Log ("  嵌套包较大 ({0} GB)，解压可能较慢，请耐心等待…" -f [math]::Round($it.Length/1GB,1)) 'Yellow'
            }
            $subOk = Expand-OneFile -File $it.FullName -Depth ($Depth + 1) -RootDir $RootDir
            if ($subOk) {
                $found++
                # 分卷：卷1 处理成功后，同目录的所有分卷卷（.002/.003 等）都是中间产物，一并清理。
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
                    Add-Log ("  已删除中间压缩包: {0}" -f (Split-Path -Leaf $rf)) 'DarkGray'
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

    if ([string]::IsNullOrEmpty($RootDir)) { $RootDir = $baseDir }

    $workDir = Get-WorkDir -File $File

    Add-Log ("{0}处理: {1}" -f $indent, $fileName) 'Cyan'

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
            Add-Log ("{0}  分卷卷1，尝试组卷解压…" -f $indent) 'DarkGray'
            $vr = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($vr -is [hashtable] -and $vr.Result -eq 'ok') {
                $nested = 0
                if ($Depth -lt $script:MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                if ($Depth -eq 1) {
                    Add-Log ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) 'Green'
                } else {
                    Add-Log ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) 'Green'
                }
                if ($nested -gt 0) {
                    Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
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

    # 2) 用当前密码直接解压
    if ($r -eq 'ok') {
        $nested = 0
        if ($Depth -lt $script:MaxDepth) {
            $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
        }
        Move-DirContents -Source $workDir -Dest $RootDir
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        if ($Depth -eq 1) {
            Add-Log ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) 'Green'
        } else {
            Add-Log ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) 'Green'
        }
        if ($nested -gt 0) {
            Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
        }
        return $true
    }

    # 2) 密码问题 → 弹窗输入（嵌套层密码可能不同，始终给重输机会）
    if ($r -eq 'wrongpass') {
        if ($script:PassFromParam -and $Depth -eq 1) {
            Add-Log ("{0}  密码错误（已填写的密码无效），跳过" -f $indent) 'Red'
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
        if ($Depth -gt 1) {
            Add-Log ("{0}  该层密码与之前不同，请重新输入" -f $indent) 'Yellow'
        }
        $pw = Show-PasswordDialog -Prompt ("该文件需要密码：{0}" -f $fileName)
        if ($null -eq $pw) {
            Add-Log ("{0}  已取消输入密码，跳过" -f $indent) 'Yellow'
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
        $script:AutoPass = $pw
        $r = Invoke-Extract -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($r -eq 'ok') {
            $nested = 0
            if ($Depth -lt $script:MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            if ($Depth -eq 1) {
                Add-Log ("{0}  OK 解压完成 → {1}" -f $indent, $RootDir) 'Green'
            } else {
                Add-Log ("{0}  OK 解压完成（已并入 {1}）" -f $indent, $RootDir) 'Green'
            }
            if ($nested -gt 0) {
                Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
            }
            return $true
        }
    }

    # 3) 不是标准压缩包 → 有真实引擎则内容识别已确认失败；仅 .NET 时改名重试
    if ($r -eq 'notarchive' -or $r -eq 'error') {
        # 3a) 分卷改名重试：卷1 缺 .001 后缀的分卷包（同目录有 .002 等）
        $volR = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($volR -is [hashtable] -and $volR.Result -eq 'ok') {
            $nested = 0
            if ($Depth -lt $script:MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Add-Log ("{0}  OK 识别为分卷压缩包（{1}），解压完成 → {2}" -f $indent, $volR.Ext, $RootDir) 'Green'
            if ($nested -gt 0) {
                Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
            }
            return $true
        }
        # 分卷包密码错误：参数密码无效时明确提示；否则弹窗重输
        if ($volR -is [hashtable] -and $volR.Result -eq 'wrongpass') {
            if ($script:PassFromParam -and $Depth -eq 1) {
                Add-Log ("{0}  密码错误（参数密码无效），跳过" -f $indent) 'Red'
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                return $false
            }
            if ($Depth -gt 1) {
                Add-Log ("{0}  该层密码与之前不同，请重新输入" -f $indent) 'Yellow'
            }
            $pw = Show-PasswordDialog -Prompt ("该文件需要密码：{0}" -f $fileName)
            if ($null -ne $pw) {
                $script:AutoPass = $pw
                $volR2 = Try-VolumeRename -File $File -OutDir $workDir -Pass $script:AutoPass
                if ($volR2 -is [hashtable] -and $volR2.Result -eq 'ok') {
                    $nested = 0
                    if ($Depth -lt $script:MaxDepth) {
                        $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                    }
                    Move-DirContents -Source $workDir -Dest $RootDir
                    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                    Add-Log ("{0}  OK 识别为分卷压缩包（{1}），解压完成 → {2}" -f $indent, $volR2.Ext, $RootDir) 'Green'
                    return $true
                }
            }
        }

        $needRenameRetry = (-not $script:SevenZip -and -not $script:RarTool)

        # 4b) 嵌入式 zip 检测：文件头是媒体格式但尾部嵌入了 zip 数据（资源站拼接包）
        #     7z 按文件头识别失败，但 WinRAR 能通过尾部 EOCD 定位打开——我们截取 zip 段再用 7z 解压
        $embR = Try-EmbeddedZip -File $File -OutDir $workDir -Pass $script:AutoPass
        if ($embR -is [hashtable] -and $embR.Result -eq 'ok') {
            Add-Log ("{0}  检测到文件尾部嵌入压缩数据，已截取解压…" -f $indent) 'DarkGray'
            $nested = 0
            if ($Depth -lt $script:MaxDepth) {
                $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
            }
            Move-DirContents -Source $workDir -Dest $RootDir
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Add-Log ("{0}  OK 识别为嵌入式压缩包，解压完成 → {1}" -f $indent, $RootDir) 'Green'
            if ($nested -gt 0) {
                Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
            }
            return $true
        }
        if ($embR -is [hashtable] -and $embR.Result -eq 'wrongpass') {
            if ($script:PassFromParam -and $Depth -eq 1) {
                Add-Log ("{0}  密码错误（参数密码无效），跳过" -f $indent) 'Red'
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                return $false
            }
            if ($Depth -gt 1) {
                Add-Log ("{0}  该层密码与之前不同，请重新输入" -f $indent) 'Yellow'
            }
            $pw = Show-PasswordDialog -Prompt ("该文件需要密码：{0}" -f $fileName)
            if ($null -ne $pw) {
                $script:AutoPass = $pw
                $embR2 = Try-EmbeddedZip -File $File -OutDir $workDir -Pass $script:AutoPass
                if ($embR2 -is [hashtable] -and $embR2.Result -eq 'ok') {
                    Add-Log ("{0}  检测到文件尾部嵌入压缩数据，已截取解压…" -f $indent) 'DarkGray'
                    $nested = 0
                    if ($Depth -lt $script:MaxDepth) {
                        $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                    }
                    Move-DirContents -Source $workDir -Dest $RootDir
                    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                    Add-Log ("{0}  OK 识别为嵌入式压缩包，解压完成 → {1}" -f $indent, $RootDir) 'Green'
                    if ($nested -gt 0) {
                        Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
                    }
                    return $true
                }
            }
        }

        if ($needRenameRetry) {
            Add-Log ("{0}  不是标准压缩包，尝试改扩展名重试…" -f $indent) 'DarkGray'
            $rr = Invoke-RenameRetry -File $File -OutDir $workDir -Pass $script:AutoPass
            if ($rr.Result -eq 'ok') {
                $nested = 0
                if ($Depth -lt $script:MaxDepth) {
                    $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                }
                Move-DirContents -Source $workDir -Dest $RootDir
                Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                Add-Log ("{0}  OK 识别为 {1} 格式，解压完成 → {2}" -f $indent, $rr.Ext, $RootDir) 'Green'
                if ($nested -gt 0) {
                    Add-Log ("{0}  嵌套: 继续解压了 {1} 个压缩包" -f $indent, $nested) 'DarkGray'
                }
                return $true
            }
            if ($rr.Result -eq 'wrongpass' -and (-not $script:PassFromParam -or $Depth -gt 1)) {
                if ($Depth -gt 1) {
                    Add-Log ("{0}  该层密码与之前不同，请重新输入" -f $indent) 'Yellow'
                }
                $pw = Show-PasswordDialog -Prompt ("该文件需要密码：{0}" -f $fileName)
                if ($null -ne $pw) {
                    $script:AutoPass = $pw
                    $rr2 = Invoke-RenameRetry -File $File -OutDir $workDir -Pass $script:AutoPass
                    if ($rr2.Result -eq 'ok') {
                        $nested = 0
                        if ($Depth -lt $script:MaxDepth) {
                            $nested = Expand-Nested -Dir $workDir -Depth $Depth -RootDir $RootDir
                        }
                        Move-DirContents -Source $workDir -Dest $RootDir
                        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                        Add-Log ("{0}  OK 识别为 {1} 格式，解压完成 → {2}" -f $indent, $rr2.Ext, $RootDir) 'Green'
                        return $true
                    }
                }
            }
        } else {
            Add-Log ("{0}  非压缩包文件（已按内容识别确认）" -f $indent) 'DarkGray'
        }
        Add-Log ("{0}  无法解压：不是可识别的压缩包，或密码错误" -f $indent) 'Red'
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    # 其他错误
    Add-Log ("{0}  解压失败：{1}" -f $indent, $r) 'Red'
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    return $false
}

# ---------- 后台处理入口 ----------
# 注意：同步运行于 UI 线程（ps2exe 打包后 .NET Thread 无 PowerShell runspace 会崩溃），
# 通过 DoEvents 泵消息保持界面响应。
function Start-Processing {
    try {
        Init-Engines
        if ($script:SevenZip) {
            Add-Log ("解压引擎: 7-Zip ({0})" -f $script:SevenZip) 'DarkGray'
        } elseif ($script:RarTool) {
            Add-Log ("解压引擎: {0}" -f $script:RarTool) 'DarkGray'
        } else {
            Add-Log '警告: 未找到 7-Zip / WinRAR，仅能用 .NET 内置解 zip（不支持密码）' 'Orange'
        }
        Add-Log ("多层嵌套: {0} 层" -f $script:MaxDepth) 'DarkGray'
        Add-Log '' 

        $script:OkCount = 0
        $script:FailCount = 0
        $targets = @($script:ListBox.Items | ForEach-Object { [string]$_ })

        foreach ($p in $targets) {
            if ($script:StopRequested) {
                Add-Log '已停止。' 'Yellow'
                break
            }
            $p = $p.Trim('"')
            if (-not (Test-Path -LiteralPath $p)) {
                Add-Log ("路径不存在: {0}" -f $p) 'Red'
                $script:FailCount++
                continue
            }
            $item = Get-Item -LiteralPath $p
            if ($item.PSIsContainer) {
                if ($script:Recurse) {
                    $files = @(Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue)
                } else {
                    $files = @(Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue)
                }
                foreach ($f in $files) {
                    if ($script:StopRequested) { break }
                    # 批量：默认按压缩扩展名/无扩展名筛选；勾选"扫描伪装扩展名"时内容探测
                    if (Should-TryFile -Name $f.Name -ForceAll $script:ForceAll) {
                        if (Expand-OneFile -File $f.FullName -Depth 1) { $script:OkCount++ } else { $script:FailCount++ }
                        Update-Status ("正在处理…  成功 {0} | 失败 {1}" -f $script:OkCount, $script:FailCount)
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
            } else {
                if (Expand-OneFile -File $item.FullName -Depth 1) { $script:OkCount++ } else { $script:FailCount++ }
                Update-Status ("正在处理…  成功 {0} | 失败 {1}" -f $script:OkCount, $script:FailCount)
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $script:StopRequested) {
            Add-Log ''
            Add-Log ("完成：成功 {0} 个，失败 {1} 个" -f $script:OkCount, $script:FailCount) 'Green'
            Update-Status ("完成   成功 {0} | 失败 {1}" -f $script:OkCount, $script:FailCount)
            # 完成提示音 + 弹窗
            if ($script:OkCount -gt 0) {
                [System.Media.SystemSounds]::Asterisk.Play()
                $msg = "解压完成：成功 {0} 个，失败 {1} 个" -f $script:OkCount, $script:FailCount
                if ($script:FailCount -gt 0) { $msg += "`n（有 {0} 个失败，详见日志）" -f $script:FailCount }
                [System.Windows.Forms.MessageBox]::Show($msg, '完成', 'OK', 'Information') | Out-Null
            } else {
                [System.Media.SystemSounds]::Exclamation.Play()
                [System.Windows.Forms.MessageBox]::Show('解压失败：没有成功解压任何文件，详见日志。', '完成', 'OK', 'Exclamation') | Out-Null
            }
        } else {
            Update-Status ("已停止   成功 {0} | 失败 {1}" -f $script:OkCount, $script:FailCount)
        }
    } catch {
        Add-Log ("发生错误: {0}" -f $_.Exception.Message) 'Red'
        Update-Status '发生错误'
    } finally {
        $script:Running = $false
        Set-RunningUi -IsRunning $false
    }
}

# ---------- 运行状态 UI ----------
function Set-RunningUi {
    param([bool]$IsRunning)
    $script:BtnStart.Enabled = -not $IsRunning
    $script:BtnStop.Enabled = $IsRunning
    $script:ListBox.Enabled = -not $IsRunning
    $script:BtnAddFile.Enabled = -not $IsRunning
    $script:BtnAddDir.Enabled = -not $IsRunning
    $script:BtnRemove.Enabled = -not $IsRunning
    $script:BtnClear.Enabled = -not $IsRunning
    $script:ChkRecurse.Enabled = -not $IsRunning
    $script:ChkForceAll.Enabled = -not $IsRunning
    $script:NumDepth.Enabled = -not $IsRunning
    $script:TxtPassword.Enabled = -not $IsRunning
    $script:ChkShowPass.Enabled = -not $IsRunning
}

# ================= 构建界面（浅色简洁风） =================
# iOS/macOS style palette: light gray-window background + white card sections
$BgColor      = [System.Drawing.Color]::FromArgb(245, 246, 248)   # window background
$CardColor    = [System.Drawing.Color]::White                     # card / content area
$BorderColor  = [System.Drawing.Color]::FromArgb(214, 218, 224)   # control thin border
$HoverColor   = [System.Drawing.Color]::FromArgb(240, 242, 245)   # button hover
$PressedColor = [System.Drawing.Color]::FromArgb(228, 231, 235)   # button pressed
$TextDark     = [System.Drawing.Color]::FromArgb(50, 50, 50)      # primary text
$TextGray     = [System.Drawing.Color]::FromArgb(120, 120, 120)   # description text
$AccentColor  = [System.Drawing.Color]::FromArgb(0, 122, 255)     # iOS blue (primary button)

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = '智能解压工具 AutoExtract'
$script:Form.StartPosition = 'CenterScreen'
$script:Form.ClientSize = New-Object System.Drawing.Size(720, 580)
$script:Form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$script:Form.MinimumSize = New-Object System.Drawing.Size(640, 500)
# DPI 缩放：控件布局按 DPI 比例自动缩放（配合 SetProcessDPIAware 在高分屏清晰）
$script:Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$script:Form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$script:Form.BackColor = $BgColor

# 窗口图标：从当前可执行文件提取（ps2exe 打包后指向 exe，自动获得自定义图标；单文件分发也生效）
try {
    $script:Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath)
} catch {
    # 提取失败（如脚本直跑时）忽略，不影响功能
}

# --- 卡片 1：文件列表区 ---
$panelFiles = New-Object System.Windows.Forms.Panel
$panelFiles.SetBounds(12, 8, 696, 212)
$panelFiles.BackColor = $CardColor

$lblTargets = New-Object System.Windows.Forms.Label
$lblTargets.Text = '要解压的文件/文件夹（可拖放）:'
$lblTargets.SetBounds(14, 10, 320, 20)
$lblTargets.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5, [System.Drawing.FontStyle]::Bold)
$lblTargets.ForeColor = $TextDark

$script:ListBox = New-Object System.Windows.Forms.ListBox
$script:ListBox.SetBounds(14, 34, 668, 132)
$script:ListBox.SelectionMode = 'MultiExtended'
$script:ListBox.AllowDrop = $true
$script:ListBox.HorizontalScrollbar = $true
$script:ListBox.BackColor = $CardColor
$script:ListBox.BorderStyle = 'FixedSingle'

# 文件按钮行（扁平样式）
$script:BtnAddFile = New-Object System.Windows.Forms.Button
$script:BtnAddFile.Text = '添加文件'
$script:BtnAddFile.SetBounds(14, 172, 90, 28)
$script:BtnAddFile.FlatStyle = 'Flat'
$script:BtnAddFile.FlatAppearance.BorderSize = 1
$script:BtnAddFile.FlatAppearance.BorderColor = $BorderColor
$script:BtnAddFile.FlatAppearance.MouseOverBackColor = $HoverColor
$script:BtnAddFile.FlatAppearance.MouseDownBackColor = $PressedColor
$script:BtnAddFile.BackColor = $CardColor
$script:BtnAddFile.ForeColor = $TextDark

$script:BtnAddDir = New-Object System.Windows.Forms.Button
$script:BtnAddDir.Text = '添加文件夹'
$script:BtnAddDir.SetBounds(110, 172, 100, 28)
$script:BtnAddDir.FlatStyle = 'Flat'
$script:BtnAddDir.FlatAppearance.BorderSize = 1
$script:BtnAddDir.FlatAppearance.BorderColor = $BorderColor
$script:BtnAddDir.FlatAppearance.MouseOverBackColor = $HoverColor
$script:BtnAddDir.FlatAppearance.MouseDownBackColor = $PressedColor
$script:BtnAddDir.BackColor = $CardColor
$script:BtnAddDir.ForeColor = $TextDark

$script:BtnRemove = New-Object System.Windows.Forms.Button
$script:BtnRemove.Text = '移除选中'
$script:BtnRemove.SetBounds(216, 172, 90, 28)
$script:BtnRemove.FlatStyle = 'Flat'
$script:BtnRemove.FlatAppearance.BorderSize = 1
$script:BtnRemove.FlatAppearance.BorderColor = $BorderColor
$script:BtnRemove.FlatAppearance.MouseOverBackColor = $HoverColor
$script:BtnRemove.FlatAppearance.MouseDownBackColor = $PressedColor
$script:BtnRemove.BackColor = $CardColor
$script:BtnRemove.ForeColor = $TextDark

$script:BtnClear = New-Object System.Windows.Forms.Button
$script:BtnClear.Text = '清空列表'
$script:BtnClear.SetBounds(312, 172, 90, 28)
$script:BtnClear.FlatStyle = 'Flat'
$script:BtnClear.FlatAppearance.BorderSize = 1
$script:BtnClear.FlatAppearance.BorderColor = $BorderColor
$script:BtnClear.FlatAppearance.MouseOverBackColor = $HoverColor
$script:BtnClear.FlatAppearance.MouseDownBackColor = $PressedColor
$script:BtnClear.BackColor = $CardColor
$script:BtnClear.ForeColor = $TextDark

# --- 卡片 2：选项区 ---
$panelOptions = New-Object System.Windows.Forms.Panel
$panelOptions.SetBounds(12, 224, 696, 78)
$panelOptions.BackColor = $CardColor

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = '密码（留空=自动检测，加密包才弹窗）:'
$lblPass.SetBounds(14, 8, 280, 20)
$lblPass.ForeColor = $TextGray

$script:TxtPassword = New-Object System.Windows.Forms.TextBox
$script:TxtPassword.SetBounds(14, 30, 240, 26)
$script:TxtPassword.UseSystemPasswordChar = $true
$script:TxtPassword.BackColor = $CardColor
$script:TxtPassword.BorderStyle = 'FixedSingle'

$script:ChkShowPass = New-Object System.Windows.Forms.CheckBox
$script:ChkShowPass.Text = '显示'
$script:ChkShowPass.SetBounds(260, 30, 55, 24)
$script:ChkShowPass.BackColor = $CardColor
$script:ChkShowPass.ForeColor = $TextDark

$script:ChkRecurse = New-Object System.Windows.Forms.CheckBox
$script:ChkRecurse.Text = '递归处理子文件夹'
$script:ChkRecurse.SetBounds(330, 32, 140, 20)
$script:ChkRecurse.BackColor = $CardColor
$script:ChkRecurse.ForeColor = $TextDark

$script:ChkForceAll = New-Object System.Windows.Forms.CheckBox
$script:ChkForceAll.Text = '扫描伪装扩展名'
$script:ChkForceAll.SetBounds(330, 52, 145, 20)
$script:ChkForceAll.BackColor = $CardColor
$script:ChkForceAll.ForeColor = $TextDark

$lblDepth = New-Object System.Windows.Forms.Label
$lblDepth.Text = '嵌套层数:'
$lblDepth.SetBounds(480, 8, 70, 20)
$lblDepth.ForeColor = $TextDark

$script:NumDepth = New-Object System.Windows.Forms.NumericUpDown
$script:NumDepth.SetBounds(480, 30, 60, 24)
$script:NumDepth.Minimum = 0
$script:NumDepth.Maximum = 99
$script:NumDepth.Value = 5

$lblDepthTip = New-Object System.Windows.Forms.Label
$lblDepthTip.Text = '0 = 只解一层'
$lblDepthTip.SetBounds(548, 32, 90, 20)
$lblDepthTip.ForeColor = $TextGray

# --- 卡片 3：操作区 ---
$panelActions = New-Object System.Windows.Forms.Panel
$panelActions.SetBounds(12, 306, 696, 50)
$panelActions.BackColor = $CardColor

$script:BtnStart = New-Object System.Windows.Forms.Button
$script:BtnStart.Text = '开始解压'
$script:BtnStart.SetBounds(14, 7, 120, 36)
$script:BtnStart.FlatStyle = 'Flat'
$script:BtnStart.FlatAppearance.BorderSize = 0
$script:BtnStart.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(10, 132, 255)
$script:BtnStart.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 102, 220)
$script:BtnStart.BackColor = $AccentColor
$script:BtnStart.ForeColor = [System.Drawing.Color]::White
$script:BtnStart.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)

$script:BtnStop = New-Object System.Windows.Forms.Button
$script:BtnStop.Text = '停止'
$script:BtnStop.SetBounds(142, 7, 90, 36)
$script:BtnStop.FlatStyle = 'Flat'
$script:BtnStop.FlatAppearance.BorderSize = 1
$script:BtnStop.FlatAppearance.BorderColor = $BorderColor
$script:BtnStop.FlatAppearance.MouseOverBackColor = $HoverColor
$script:BtnStop.FlatAppearance.MouseDownBackColor = $PressedColor
$script:BtnStop.BackColor = $CardColor
$script:BtnStop.ForeColor = $TextDark
$script:BtnStop.Enabled = $false

# --- 卡片 4：日志区 ---
$panelLog = New-Object System.Windows.Forms.Panel
$panelLog.SetBounds(12, 360, 696, 176)
$panelLog.BackColor = $CardColor

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = '日志:'
$lblLog.SetBounds(14, 8, 60, 20)
$lblLog.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5, [System.Drawing.FontStyle]::Bold)
$lblLog.ForeColor = $TextDark

$script:LogBox = New-Object System.Windows.Forms.RichTextBox
$script:LogBox.SetBounds(14, 30, 668, 136)
$script:LogBox.ReadOnly = $true
$script:LogBox.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
$script:LogBox.ForeColor = $TextDark
$script:LogBox.Font = New-Object System.Drawing.Font('Consolas', 9)
$script:LogBox.BorderStyle = 'FixedSingle'

# --- 状态栏 ---
$script:LblStatus = New-Object System.Windows.Forms.Label
$script:LblStatus.SetBounds(12, 544, 696, 24)
$script:LblStatus.Text = '就绪'
$script:LblStatus.ForeColor = $TextGray

# --- 添加卡片到窗体 ---
$panelFiles.Controls.AddRange(@($lblTargets, $script:ListBox, $script:BtnAddFile, $script:BtnAddDir, $script:BtnRemove, $script:BtnClear))
$panelOptions.Controls.AddRange(@($lblPass, $script:TxtPassword, $script:ChkShowPass, $script:ChkRecurse, $script:ChkForceAll, $lblDepth, $script:NumDepth, $lblDepthTip))
$panelActions.Controls.AddRange(@($script:BtnStart, $script:BtnStop))
$panelLog.Controls.AddRange(@($lblLog, $script:LogBox))
$script:Form.Controls.AddRange(@($panelFiles, $panelOptions, $panelActions, $panelLog, $script:LblStatus))

# ================= 事件 =================
# 拖放
$script:ListBox.Add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})
$script:ListBox.Add_DragDrop({
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($f in $files) { [void]$script:ListBox.Items.Add([string]$f) }
})

# 添加文件
$script:BtnAddFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = '选择要解压的文件'
    if ($dlg.ShowDialog() -eq 'OK') {
        foreach ($f in $dlg.FileNames) { [void]$script:ListBox.Items.Add($f) }
    }
})

# 添加文件夹
$script:BtnAddDir.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择要解压的文件夹'
    if ($dlg.ShowDialog() -eq 'OK') { [void]$script:ListBox.Items.Add($dlg.SelectedPath) }
})

# 移除选中
$script:BtnRemove.Add_Click({
    $sel = @($script:ListBox.SelectedItems | ForEach-Object { [string]$_ })
    foreach ($s in $sel) { [void]$script:ListBox.Items.Remove($s) }
})

# 清空
$script:BtnClear.Add_Click({ $script:ListBox.Items.Clear() })

# 显示密码
$script:ChkShowPass.Add_CheckedChanged({
    $script:TxtPassword.UseSystemPasswordChar = -not $script:ChkShowPass.Checked
})

# 开始解压（同步执行，DoEvents 泵消息保持 UI 响应；ps2exe 兼容，不用 .NET Thread）
function Start-AutoExtract {
    if ($script:Running) { return }
    if ($script:ListBox.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('请先添加要解压的文件或文件夹。', '提示', 'OK', 'Information') | Out-Null
        return
    }
    $script:Running = $true
    $script:StopRequested = $false
    $script:MaxDepth = [int]$script:NumDepth.Value
    $script:Recurse = $script:ChkRecurse.Checked
    $script:ForceAll = $script:ChkForceAll.Checked
    $script:AutoPass = $script:TxtPassword.Text
    $script:PassFromParam = -not [string]::IsNullOrEmpty($script:TxtPassword.Text)
    Set-RunningUi -IsRunning $true
    $script:LogBox.Clear()
    Add-Log ('开始处理 {0} 个目标…' -f $script:ListBox.Items.Count) 'Blue'
    [System.Windows.Forms.Application]::DoEvents()

    # 关键：通过 BeginInvoke 让 Start-Processing 在消息循环的独立迭代中运行，
    # 而不是嵌套在当前事件（Click/Shown）内部。否则处理中弹出密码对话框
    # （模态 ShowDialog）会与父窗体的消息循环冲突，导致界面卡死。
    $script:Form.BeginInvoke([Action]{ Start-Processing }) | Out-Null
}

# 开始
$script:BtnStart.Add_Click({ Start-AutoExtract })

# 停止
$script:BtnStop.Add_Click({
    $script:StopRequested = $true
    $script:BtnStop.Enabled = $false
    Add-Log '正在停止（处理完当前文件后退出）…' 'Yellow'
})

# 关闭窗口时若在运行，直接退出（后台线程为后台线程）
$script:Form.Add_FormClosing({
    if ($script:Running) {
        $script:StopRequested = $true
    }
})

# ================= 启动 =================
# 拖放到 exe/脚本上的文件路径：窗体显示后自动加入列表并开始解压
$script:Form.Add_Shown({
    try {
        if ($StartupPaths -and $StartupPaths.Count -gt 0) {
            foreach ($sp in $StartupPaths) {
                $sp = $sp.Trim('"')
                if (Test-Path -LiteralPath $sp) { [void]$script:ListBox.Items.Add([string]$sp) }
            }
            if ($script:ListBox.Items.Count -gt 0) {
                # Start-AutoExtract 内部已用 BeginInvoke 延迟处理，
                # 这里直接调用即可（消息循环就绪后自动启动）
                Start-AutoExtract
            }
        }
    } catch {
        Add-Log ("自动启动失败: {0}" -f $_.Exception.Message) 'Red'
    }
})
[void]$script:Form.ShowDialog()
