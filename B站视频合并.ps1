<#
.SYNOPSIS
    哔哩哔哩缓存视频合并工具（图形界面版）
.DESCRIPTION
    把 B 站 Windows 客户端下载/缓存的 m4s 音视频分片，无损合并成可直接播放的 mp4。

    · 支持把文件夹 / 视频子文件夹 / m4s 文件直接拖进窗口
    · 自动识别哪条是视频流、哪条是音频流，并剥离 B 站特有的 9 字节前缀
    · 使用 ffmpeg -c copy 无损封装，不重新编码，画质无损失、速度极快
    · 输出目录可自定义，默认与本工具同级的 outputs 文件夹
    · 找不到 ffmpeg 时可手动指定，位置会被记住

.NOTES
    启动方式：双击同目录的「启动B站视频合并.cmd」
    自检模式：powershell -STA -File "本脚本" -SelfTest
#>
[CmdletBinding()]
param([switch]$SelfTest)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if (-not $SelfTest) { [System.Windows.Forms.Application]::EnableVisualStyles() }

trap { [void][System.Windows.Forms.MessageBox]::Show(($_ | Out-String), '出错了'); exit 1 }

# ==================== 全局 ====================
$script:AppDir      = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:OutDir      = Join-Path $script:AppDir 'outputs'
$script:TempDir     = Join-Path $script:AppDir 'work'
$script:CfgFile     = Join-Path $script:AppDir 'settings.json'
$script:DefaultScan = Join-Path $env:USERPROFILE 'Videos\bilibili'
$script:CfgFfmpeg   = $null
$script:Ffmpeg      = $null
$script:Tasks       = New-Object System.Collections.ArrayList
$script:IsRunning   = $false

foreach ($d in @($script:OutDir, $script:TempDir)) {
    if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
}

# 读取上次保存的设置（输出目录 / ffmpeg 位置）
try {
    if (Test-Path -LiteralPath $script:CfgFile) {
        $cfg = Get-Content -LiteralPath $script:CfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($cfg.OutDir)     { $script:OutDir    = $cfg.OutDir }
        if ($cfg.FfmpegPath) { $script:CfgFfmpeg = $cfg.FfmpegPath }
    }
} catch { }

# ==================== 工具函数 ====================
# 按优先级查找 ffmpeg：配置记忆 -> 本程序同目录 -> B站客户端自带 -> 常见安装位置 -> PATH
function Resolve-Ffmpeg {
    $cands = New-Object System.Collections.ArrayList
    if ($script:CfgFfmpeg) { [void]$cands.Add($script:CfgFfmpeg) }
    if ($script:AppDir)    { [void]$cands.Add((Join-Path $script:AppDir 'ffmpeg.exe')) }

    # B 站 Windows 客户端自带一份 ffmpeg，B 站用户通常不必额外安装
    if ($env:APPDATA)      { [void]$cands.Add((Join-Path $env:APPDATA 'bilibili\ffmpeg\ffmpeg.exe')) }

    # 常见的包管理器 / 手动安装位置
    if ($env:LOCALAPPDATA) { [void]$cands.Add((Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links\ffmpeg.exe')) }
    if ($env:ProgramData)  { [void]$cands.Add((Join-Path $env:ProgramData 'chocolatey\bin\ffmpeg.exe')) }
    [void]$cands.Add('C:\ffmpeg\bin\ffmpeg.exe')
    [void]$cands.Add('C:\Program Files\ffmpeg\bin\ffmpeg.exe')
    [void]$cands.Add('C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe')

    # 最后看 PATH
    $cmd = Get-Command ffmpeg.exe -ErrorAction SilentlyContinue
    if ($cmd) { [void]$cands.Add($cmd.Source) }

    foreach ($x in $cands) {
        if ($x -and (Test-Path -LiteralPath $x)) { return $x }
    }
    return $null
}
# 读文件头部，判断这条 m4s 是视频流还是音频流
function Get-StreamKind {
    param([string]$File)
    $fs = [System.IO.File]::OpenRead($File)
    try {
        $len = [Math]::Min(262144, [int]$fs.Length)
        $buf = New-Object byte[] $len
        [void]$fs.Read($buf, 0, $len)
    } finally { $fs.Dispose() }
    $txt = [System.Text.Encoding]::ASCII.GetString($buf)
    if ($txt.Contains('vide')) { return 'video' }
    if ($txt.Contains('soun')) { return 'audio' }
    return 'unknown'
}

# 剥离 B 站 m4s 开头的 9 个 "0" 字节（确认存在才剥，避免误伤正常 mp4）
function Copy-WithoutBiliPrefix {
    param([string]$InFile, [string]$OutFile)
    $src = [System.IO.File]::OpenRead($InFile)
    try {
        $head = New-Object byte[] 9
        $n = $src.Read($head, 0, 9)
        $skip = 0
        if ($n -eq 9) {
            $allZero = $true
            foreach ($b in $head) { if ($b -ne 0x30) { $allZero = $false; break } }
            if ($allZero) { $skip = 9 }
        }
        $src.Position = $skip
        $dst = [System.IO.File]::Create($OutFile)
        try { $src.CopyTo($dst) } finally { $dst.Dispose() }
    } finally { $src.Dispose() }
}

# 文件名清洗：替换掉 Windows 不允许的字符
function Get-SafeName {
    param([string]$Name)
    $bad = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($bad -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    $r = $sb.ToString().Trim().TrimEnd('.')
    if (-not $r) { $r = '未命名视频' }
    if ($r.Length -gt 120) { $r = $r.Substring(0, 120) }
    return $r
}

# 找一个不冲突的输出文件名；同名且大小一致 => 判定为已处理过，跳过
function Resolve-OutFile {
    param([string]$Dir, [string]$Base, [long]$ExpectSize)
    $target = Join-Path $Dir ($Base + '.mp4')
    if (-not (Test-Path -LiteralPath $target)) { return @{ Path = $target; Skip = $false } }
    $exist = Get-Item -LiteralPath $target
    $tol = [Math]::Max(8192, $ExpectSize * 0.02)
    if ([Math]::Abs($exist.Length - $ExpectSize) -le $tol) { return @{ Path = $target; Skip = $true } }
    for ($i = 2; $i -lt 100; $i++) {
        $p = Join-Path $Dir ("$Base ($i).mp4")
        if (-not (Test-Path -LiteralPath $p)) { return @{ Path = $p; Skip = $false } }
    }
    throw "输出目录里同名文件太多：$Base"
}

# 从一个缓存目录构建任务对象（识别视频流 / 音频流 / 标题）
function New-TaskFromDir {
    param([string]$Dir)
    $files = @(Get-ChildItem -LiteralPath $Dir -Filter '*.m4s' -File -ErrorAction SilentlyContinue | Sort-Object Length -Descending)
    if ($files.Count -eq 0) { return $null }
    $v = $null; $a = $null
    foreach ($f in $files) {
        $kind = Get-StreamKind -File $f.FullName
        if ($kind -eq 'unknown' -and $f.Name -match '-(\d{5})\.m4s$') {
            # 兜底：按 B 站清晰度编号判断，302xx 为音频，其余为视频
            $qn = [int]$Matches[1]
            if ($qn -ge 30200 -and $qn -lt 30300) { $kind = 'audio' } else { $kind = 'video' }
        }
        if ($kind -eq 'video' -and -not $v) { $v = $f; continue }
        if ($kind -eq 'audio' -and -not $a) { $a = $f; continue }
    }
    if (-not $v) { return $null }

    $title = $null
    foreach ($metaName in @('videoInfo.json', '.videoInfo')) {
        if ($title) { break }
        $mp = Join-Path $Dir $metaName
        if (Test-Path -LiteralPath $mp) {
            try {
                $j = Get-Content -LiteralPath $mp -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($j.title) { $title = $j.title }
            } catch { }
        }
    }
    if (-not $title) { $title = Split-Path -Leaf $Dir }

    $aPath = $null; $aName = $null; $aSize = 0
    if ($a) { $aPath = $a.FullName; $aName = $a.Name; $aSize = [long]$a.Length }
    return [pscustomobject]@{
        Dir       = $Dir
        Title     = $title
        VideoPath = $v.FullName
        VideoName = $v.Name
        AudioPath = $aPath
        AudioName = $aName
        Size      = [long]$v.Length + $aSize
        Status    = '待处理'
    }
}

# 把用户给的路径展开成任务列表
function ConvertTo-Tasks {
    param([string]$Path)
    $result = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $Path)) { return $result }
    $item = Get-Item -LiteralPath $Path
    if (-not $item.PSIsContainer) {
        $d = [System.IO.Path]::GetDirectoryName($item.FullName)
        if ($d -and (Test-Path -LiteralPath $d)) { $item = Get-Item -LiteralPath $d } else { return $result }
    }
    $dirs = New-Object System.Collections.ArrayList
    if (@(Get-ChildItem -LiteralPath $item.FullName -Filter '*.m4s' -File -ErrorAction SilentlyContinue).Count -gt 0) {
        [void]$dirs.Add($item.FullName)
    } else {
        foreach ($sub in Get-ChildItem -LiteralPath $item.FullName -Directory -ErrorAction SilentlyContinue) {
            if (@(Get-ChildItem -LiteralPath $sub.FullName -Filter '*.m4s' -File -ErrorAction SilentlyContinue).Count -gt 0) {
                [void]$dirs.Add($sub.FullName)
            }
        }
    }
    foreach ($d in $dirs) {
        $t = New-TaskFromDir -Dir $d
        if ($t) { [void]$result.Add($t) }
    }
    return $result
}

function Save-Config {
    try {
        $o = [pscustomobject]@{ OutDir = $script:OutDir; FfmpegPath = $script:Ffmpeg }
        [System.IO.File]::WriteAllText($script:CfgFile, ($o | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
}

# 合并单个任务：剥离前缀 -> ffmpeg 无损封装 -> 落盘
function Invoke-MergeTask {
    param($T)
    $base = Get-SafeName -Name $T.Title
    $r = Resolve-OutFile -Dir $script:OutDir -Base $base -ExpectSize $T.Size
    if ($r.Skip) {
        Write-Log ("  已存在且大小一致，跳过：{0}" -f (Split-Path -Leaf $r.Path))
        return 'skip'
    }
    $outFile = $r.Path
    if (Test-Path -LiteralPath $outFile) { [System.IO.File]::Delete($outFile) }

    $tmpV = Join-Path $script:TempDir 'gui_v.mp4'
    $tmpA = Join-Path $script:TempDir 'gui_a.m4a'
    $tmpO = Join-Path $script:TempDir 'gui_out.mp4'
    foreach ($x in @($tmpV, $tmpA, $tmpO)) { if (Test-Path -LiteralPath $x) { [System.IO.File]::Delete($x) } }
    try {
        Write-Log ("  视频流：{0}（{1} MB）" -f $T.VideoName, [math]::Round((Get-Item -LiteralPath $T.VideoPath).Length / 1MB, 1))
        if ($T.AudioPath) { Write-Log ("  音频流：{0}" -f $T.AudioName) } else { Write-Log '  音频流：无（输出无声视频）' }
        Copy-WithoutBiliPrefix -InFile $T.VideoPath -OutFile $tmpV
        $ffArgs = @('-hide_banner', '-loglevel', 'error', '-y', '-i', $tmpV)
        if ($T.AudioPath) {
            Copy-WithoutBiliPrefix -InFile $T.AudioPath -OutFile $tmpA
            $ffArgs += @('-i', $tmpA)
        }
        $ffArgs += @('-c', 'copy', '-movflags', '+faststart', $tmpO)
        & $script:Ffmpeg @ffArgs
        if ($LASTEXITCODE -ne 0) { throw "ffmpeg 退出码 $LASTEXITCODE" }
        [System.IO.File]::Move($tmpO, $outFile)
        Write-Log ("  已输出：{0}（{1} MB）" -f (Split-Path -Leaf $outFile), [math]::Round((Get-Item -LiteralPath $outFile).Length / 1MB, 1))
        return 'ok'
    } finally {
        foreach ($x in @($tmpV, $tmpA, $tmpO)) { if (Test-Path -LiteralPath $x) { [System.IO.File]::Delete($x) } }
    }
}
# ==================== 界面 ====================
$fontUI    = New-Object System.Drawing.Font('微软雅黑', 9)
$fontTitle = New-Object System.Drawing.Font('微软雅黑', 12, [System.Drawing.FontStyle]::Bold)

$form = New-Object System.Windows.Forms.Form
$form.Text = '哔哩哔哩缓存视频合并工具'
$form.ClientSize = New-Object System.Drawing.Size(624, 552)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.Font = $fontUI
$form.AllowDrop = $true

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = '哔哩哔哩缓存视频合并工具'
$lblTitle.Font = $fontTitle
$lblTitle.Location = New-Object System.Drawing.Point(12, 10)
$lblTitle.Size = New-Object System.Drawing.Size(600, 26)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = '把 bilibili 文件夹 / 视频子文件夹 / m4s 文件拖进窗口，或点下面的按钮添加'
$lblHint.ForeColor = [System.Drawing.Color]::Gray
$lblHint.Location = New-Object System.Drawing.Point(14, 38)
$lblHint.Size = New-Object System.Drawing.Size(600, 18)

$panelDrop = New-Object System.Windows.Forms.Panel
$panelDrop.Location = New-Object System.Drawing.Point(12, 60)
$panelDrop.Size = New-Object System.Drawing.Size(600, 46)
$panelDrop.BorderStyle = 'FixedSingle'
$panelDrop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$panelDrop.AllowDrop = $true
$lblDrop = New-Object System.Windows.Forms.Label
$lblDrop.Text = '拖拽到此处添加（文件夹或文件都可以）'
$lblDrop.TextAlign = 'MiddleCenter'
$lblDrop.ForeColor = [System.Drawing.Color]::FromArgb(120, 130, 145)
$lblDrop.Dock = 'Fill'
$panelDrop.Controls.Add($lblDrop)

$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = '添加文件夹'
$btnAddFolder.Location = New-Object System.Drawing.Point(12, 112)
$btnAddFolder.Size = New-Object System.Drawing.Size(104, 28)

$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = '添加 m4s 文件'
$btnAddFile.Location = New-Object System.Drawing.Point(124, 112)
$btnAddFile.Size = New-Object System.Drawing.Size(112, 28)

$btnRescan = New-Object System.Windows.Forms.Button
$btnRescan.Text = '扫描默认下载目录'
$btnRescan.Location = New-Object System.Drawing.Point(244, 112)
$btnRescan.Size = New-Object System.Drawing.Size(136, 28)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = '清空列表'
$btnClear.Location = New-Object System.Drawing.Point(388, 112)
$btnClear.Size = New-Object System.Drawing.Size(88, 28)

$lv = New-Object System.Windows.Forms.ListView
$lv.Location = New-Object System.Drawing.Point(12, 146)
$lv.Size = New-Object System.Drawing.Size(600, 172)
$lv.View = 'Details'
$lv.CheckBoxes = $true
$lv.FullRowSelect = $true
$lv.MultiSelect = $true
$lv.AllowDrop = $true
$lv.HideSelection = $false
[void]$lv.Columns.Add('视频标题', 330)
[void]$lv.Columns.Add('大小', 90)
[void]$lv.Columns.Add('状态', 155)

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = '输出到：'
$lblOut.Location = New-Object System.Drawing.Point(12, 330)
$lblOut.Size = New-Object System.Drawing.Size(56, 22)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(68, 328)
$txtOut.Size = New-Object System.Drawing.Size(448, 23)
$txtOut.Text = $script:OutDir

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '浏览'
$btnBrowse.Location = New-Object System.Drawing.Point(522, 327)
$btnBrowse.Size = New-Object System.Drawing.Size(90, 26)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(12, 362)
$progress.Size = New-Object System.Drawing.Size(600, 18)
$progress.Minimum = 0
$progress.Maximum = 100

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 388)
$logBox.Size = New-Object System.Drawing.Size(600, 96)
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$logBox.BackColor = [System.Drawing.Color]::White

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = '开始合并'
$btnStart.Location = New-Object System.Drawing.Point(12, 494)
$btnStart.Size = New-Object System.Drawing.Size(176, 42)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(0, 161, 214)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = '打开输出目录'
$btnOpen.Location = New-Object System.Drawing.Point(198, 494)
$btnOpen.Size = New-Object System.Drawing.Size(140, 42)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = '关闭'
$btnExit.Location = New-Object System.Drawing.Point(534, 494)
$btnExit.Size = New-Object System.Drawing.Size(78, 42)

foreach ($c in @($lblTitle, $lblHint, $panelDrop, $btnAddFolder, $btnAddFile, $btnRescan, $btnClear, $lv, $lblOut, $txtOut, $btnBrowse, $progress, $logBox, $btnStart, $btnOpen, $btnExit)) {
    [void]$form.Controls.Add($c)
}

$script:Form     = $form
$script:ListView = $lv
$script:LogBox   = $logBox
$script:Progress = $progress
$script:BtnStart = $btnStart
$script:TxtOut   = $txtOut

function Write-Log {
    param([string]$Msg)
    $script:LogBox.AppendText(('[{0}] {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $Msg, [Environment]::NewLine))
    $script:LogBox.SelectionStart = $script:LogBox.TextLength
    $script:LogBox.ScrollToCaret()
    [void]$script:Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-TaskRow {
    param($T)
    $row = New-Object System.Windows.Forms.ListViewItem($T.Title)
    [void]$row.SubItems.Add(('{0} MB' -f [math]::Round($T.Size / 1MB, 1)))
    [void]$row.SubItems.Add('待处理')
    $row.Checked = $true
    $row.Tag = $T
    [void]$script:ListView.Items.Add($row)
}

function Set-RowStatus {
    param($Row, [string]$Text)
    $Row.SubItems[2].Text = $Text
    [void]$script:Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-Paths {
    param([string[]]$Paths)
    $added = 0; $dup = 0; $bad = 0
    foreach ($p in $Paths) {
        $ts = @(ConvertTo-Tasks -Path $p)
        if ($ts.Count -eq 0) { $bad++ }
        foreach ($t in $ts) {
            $exists = $false
            foreach ($x in $script:Tasks) { if ($x.Dir -eq $t.Dir) { $exists = $true; break } }
            if ($exists) { $dup++; continue }
            [void]$script:Tasks.Add($t)
            Add-TaskRow -T $t
            $added++
        }
    }
    $msg = "已添加 $added 个任务"
    if ($dup -gt 0) { $msg += "，跳过 $dup 个重复项" }
    if ($bad -gt 0) { $msg += "，$bad 个路径里没找到 m4s 分片" }
    Write-Log $msg
}

function Start-Merge {
    if ($script:IsRunning) { return }
    if (-not $script:Ffmpeg) {
        [void][System.Windows.Forms.MessageBox]::Show('还没有可用的 ffmpeg，无法合并。请重启本程序并按提示指定 ffmpeg.exe 的位置。', '缺少 ffmpeg', 'OK', 'Warning')
        return
    }
    $rows = @()
    foreach ($row in $script:ListView.Items) { if ($row.Checked) { $rows += $row } }
    if ($rows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('请先勾选要合并的视频。', '提示', 'OK', 'Information')
        return
    }
    $script:OutDir = $script:TxtOut.Text.Trim()
    try {
        if (-not (Test-Path -LiteralPath $script:OutDir)) { [void](New-Item -ItemType Directory -Path $script:OutDir -Force) }
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("无法创建输出目录：$($script:OutDir)", '错误', 'OK', 'Error')
        return
    }
    Save-Config

    $script:IsRunning = $true
    $script:BtnStart.Enabled = $false
    $script:Progress.Maximum = $rows.Count
    $script:Progress.Value = 0
    $ok = 0; $skip = 0; $fail = 0; $n = 0
    Write-Log "开始合并，共 $($rows.Count) 个视频"
    try {
        foreach ($row in $rows) {
            $n++
            $t = $row.Tag
            Set-RowStatus -Row $row -Text '处理中...'
            Write-Log "处理：$($t.Title)"
            try {
                $r = Invoke-MergeTask -T $t
                if ($r -eq 'skip') { $skip++; Set-RowStatus -Row $row -Text '已存在，跳过' }
                else { $ok++; Set-RowStatus -Row $row -Text '完成' }
            } catch {
                $fail++
                Set-RowStatus -Row $row -Text '失败'
                Write-Log ("  失败：$($_.Exception.Message)")
            }
            $script:Progress.Value = $n
            [void]$script:Form.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
    } finally {
        $script:IsRunning = $false
        $script:BtnStart.Enabled = $true
    }
    Write-Log "全部结束：成功 $ok，跳过 $skip，失败 $fail"
}
# ==================== 事件 ====================
$onDrop = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        Add-Paths -Paths @($e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
    }
}
$onDragEnter = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } else { $e.Effect = 'None' }
}
$form.Add_DragEnter($onDragEnter)
$form.Add_DragDrop($onDrop)
$panelDrop.Add_DragEnter($onDragEnter)
$panelDrop.Add_DragDrop($onDrop)
$lv.Add_DragEnter($onDragEnter)
$lv.Add_DragDrop($onDrop)

$btnAddFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择 B 站下载/缓存目录（整个 bilibili 文件夹或单个视频子文件夹都可以）'
    if ($dlg.ShowDialog() -eq 'OK') { Add-Paths -Paths @($dlg.SelectedPath) }
})
$btnAddFile.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'B站分片 (*.m4s)|*.m4s|所有文件 (*.*)|*.*'
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq 'OK') { Add-Paths -Paths $ofd.FileNames }
})
$btnRescan.Add_Click({
    Write-Log "扫描默认下载目录：$($script:DefaultScan)"
    Add-Paths -Paths @($script:DefaultScan)
})
$btnClear.Add_Click({
    $script:Tasks.Clear()
    $script:ListView.Items.Clear()
    $script:Progress.Value = 0
    Write-Log '已清空列表'
})
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择输出目录'
    if ($dlg.ShowDialog() -eq 'OK') { $script:TxtOut.Text = $dlg.SelectedPath }
})
$btnStart.Add_Click({ Start-Merge })
$btnOpen.Add_Click({
    $d = $script:TxtOut.Text.Trim()
    if (Test-Path -LiteralPath $d) { Start-Process explorer.exe -ArgumentList "`"$d`"" } else { [void][System.Windows.Forms.MessageBox]::Show('输出目录还不存在，先合并一次吧。', '提示') }
})
$btnExit.Add_Click({ $script:Form.Close() })

# ==================== 启动 ====================
$script:Ffmpeg = Resolve-Ffmpeg

if ($SelfTest) {
    Write-Output 'SelfTest: 窗体与控件构建成功'
    Write-Output ("  控件数量: {0}" -f $form.Controls.Count)
    Write-Output ("  ffmpeg  : {0}" -f $script:Ffmpeg)
    Write-Output ("  输出目录: {0}" -f $script:OutDir)
    $ts = @(ConvertTo-Tasks -Path $script:DefaultScan)
    Write-Output ("  默认目录扫描到 {0} 个视频：" -f $ts.Count)
    foreach ($t in $ts) {
        $hasAudio = '无'
        if ($t.AudioPath) { $hasAudio = '有' }
        Write-Output ("    - {0}  [{1} MB, 音频:{2}]" -f $t.Title, [math]::Round($t.Size / 1MB, 1), $hasAudio)
    }
    $form.Dispose()
    exit 0
}

# 没找到 ffmpeg 就请用户手动指定，并记住位置
if (-not $script:Ffmpeg) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = '没有找到 ffmpeg.exe，请手动指定（取消则本次无法合并）'
    $ofd.Filter = 'ffmpeg (ffmpeg.exe)|ffmpeg.exe|所有文件 (*.*)|*.*'
    if ($ofd.ShowDialog() -eq 'OK') {
        $script:Ffmpeg = $ofd.FileName
        Save-Config
    }
}

if (-not $script:Ffmpeg) {
    Write-Log '警告：没有可用的 ffmpeg，合并功能不可用。'
} else {
    Write-Log "就绪，ffmpeg：$($script:Ffmpeg)"
    Write-Log "输出目录：$($script:OutDir)"
}
if (Test-Path -LiteralPath $script:DefaultScan) {
    Add-Paths -Paths @($script:DefaultScan)
} else {
    Write-Log "默认下载目录不存在：$($script:DefaultScan)，请手动拖入文件夹"
}

[void]$form.ShowDialog()
$form.Dispose()