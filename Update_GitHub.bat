@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
title RPG Collector - GitHub 更新
set "RPG_GITHUB_BAT=%~f0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:RPG_GITHUB_BAT,[Text.Encoding]::UTF8);$m='# __RPG_GITHUB_POWERSHELL__';& ([ScriptBlock]::Create($s.Substring($s.LastIndexOf($m)+$m.Length)))"
set "RPG_GITHUB_EXIT=%errorlevel%"
echo.
pause
exit /b %RPG_GITHUB_EXIT%

# __RPG_GITHUB_POWERSHELL__
# Windows PowerShell is built into Windows. Python is not required.
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$phase = '检查项目'
$messageFile = $null
$gitExe = $null
$repoReady = $false
$createdCommit = $false

function Invoke-Git {
    param([string[]] $GitArgs)
    & $gitExe --no-pager -c core.quotepath=false @GitArgs
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        throw "Git 命令失败，退出码 $code。请查看上方的具体错误。"
    }
}

function Read-Git {
    param([string[]] $GitArgs)
    $lines = @(& $gitExe --no-pager -c core.quotepath=false @GitArgs)
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        throw "无法读取 Git 状态，退出码 $code。请查看上方的具体错误。"
    }
    return $lines
}

function Assert-NoPendingOperation {
    foreach ($marker in @('rebase-merge', 'rebase-apply', 'MERGE_HEAD', 'CHERRY_PICK_HEAD', 'REVERT_HEAD', 'sequencer')) {
        $markerPath = [string](Read-Git -GitArgs @('rev-parse', '--git-path', $marker))
        if (Test-Path -LiteralPath $markerPath) {
            throw '检测到尚未完成的同步、合并或提交操作。请先处理现有操作，再运行本脚本；不会自动选择冲突版本。'
        }
    }
    $unmerged = @(Read-Git -GitArgs @('ls-files', '--unmerged'))
    if ($unmerged.Count -gt 0) {
        throw '检测到尚未解决的文件冲突。请先解决冲突，本次不会继续提交或上传。'
    }
}

function Test-ExcludedPath {
    param([string] $Path)
    $normalized = $Path.Replace('\', '/')
    return ($normalized -match '(^|/)(venv|\.venv|output|__pycache__)(/|$)' -or
            $normalized -match '\.pyc$' -or
            $normalized -match '(^|/)~\$[^/]*\.xlsx$')
}

try {
    $folder = [IO.Path]::GetDirectoryName($env:RPG_GITHUB_BAT)
    Set-Location -LiteralPath $folder
    $gitExe = (Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $top = [string](Read-Git -GitArgs @('rev-parse', '--show-toplevel'))
    if ([IO.Path]::GetFullPath($top).TrimEnd('\', '/') -ne [IO.Path]::GetFullPath($folder).TrimEnd('\', '/')) {
        throw '请把 Update_GitHub.bat 放在项目根目录，与 RPGcollector.py 放在一起。'
    }
    $repoReady = $true
    Assert-NoPendingOperation
    $branch = [string](Read-Git -GitArgs @('branch', '--show-current'))
    if ($branch -ne 'main') {
        throw "当前分支为 '$branch'，本脚本只处理 main 分支。请先确认分支。"
    }
    $remote = [string](Read-Git -GitArgs @('remote', 'get-url', '--push', 'origin'))
    Write-Host "上传目标：$remote"
    Write-Host '提交范围：项目内的改动；排除运行环境、output、缓存和 Excel 临时文件。'
    Write-Host ''

    # Detect the file occupation that caused the previous Excel unlink error.
    foreach ($book in @(Get-ChildItem -LiteralPath $folder -Filter '*.xlsx' -File)) {
        if ($book.Name.StartsWith('~$')) { continue }
        try {
            $stream = [IO.File]::Open($book.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
            $stream.Dispose()
        } catch {
            throw "无法独占读取 $($book.Name)。请保存并关闭 Excel/WPS，或检查文件访问权限后重试。"
        }
    }

    # Never silently include an excluded file that was staged before this run.
    foreach ($path in @(Read-Git -GitArgs @('diff', '--cached', '--name-only'))) {
        if (Test-ExcludedPath -Path $path) {
            throw "暂存区已有应排除的文件：$path。请先将它移出暂存区，再重试。"
        }
    }

    $phase = '选择本次提交文件'
    $addArgs = @('add', '--all', '--', '.',
        ':(glob,exclude)**/venv/**',
        ':(glob,exclude)**/.venv/**',
        ':(glob,exclude)**/output/**',
        ':(glob,exclude)**/__pycache__/**',
        ':(glob,exclude)**/*.pyc',
        ':(glob,exclude)**/~$*.xlsx')
    Invoke-Git -GitArgs $addArgs
    & $gitExe --no-pager diff --cached --quiet --exit-code
    $diffCode = $LASTEXITCODE
    if ($diffCode -gt 1) { throw '无法检查本次提交清单。' }

    if ($diffCode -eq 1) {
        Write-Host '本次准备提交的文件：' -ForegroundColor Cyan
        Invoke-Git -GitArgs @('diff', '--cached', '--stat')
        Write-Host ''
        $message = Read-Host '请输入本次修改说明（直接回车取消）'
        if ([string]::IsNullOrWhiteSpace($message)) {
            Write-Host '已取消：没有创建提交，也没有同步或上传。文件仍在暂存区，可修改后重新运行。'
            exit 0
        }
        $phase = '创建本地版本'
        $messageFile = Join-Path ([IO.Path]::GetTempPath()) ('RPG_GitHub_' + [Guid]::NewGuid().ToString('N') + '.txt')
        [IO.File]::WriteAllText($messageFile, $message.Trim() + [Environment]::NewLine, $utf8)
        # A message file preserves Chinese, quotes, percent signs and shell symbols.
        Invoke-Git -GitArgs @('commit', '--file', $messageFile)
        $createdCommit = $true
    } else {
        Write-Host '没有新的文件改动，跳过创建提交，继续同步并上传已有版本。'
    }

    $phase = '同步 GitHub 最新版本'
    Write-Host ''
    Write-Host '正在同步 GitHub 最新版本……' -ForegroundColor Cyan
    Invoke-Git -GitArgs @('pull', '--rebase', '--autostash', 'origin', 'main')
    # A successful command may still leave conflicts when restoring an autostash.
    Assert-NoPendingOperation

    $phase = '上传到 GitHub'
    Write-Host ''
    Write-Host '正在上传；如弹出浏览器，请完成 GitHub 登录授权……' -ForegroundColor Cyan
    Invoke-Git -GitArgs @('push', 'origin', 'main')
    Write-Host ''
    Write-Host '更新完成：main 分支已上传。' -ForegroundColor Green
    Invoke-Git -GitArgs @('status', '--short', '--branch')
} catch {
    Write-Host ''
    Write-Host "已停止，未执行后续步骤。停止位置：$phase" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    if ($createdCommit) {
        Write-Host '本次本地提交已保留。处理错误后重试，无需重复创建相同的提交。'
    }
    Write-Host '若出现 CONFLICT：先选定并保存正确版本，再 git add 对应文件，然后执行 git rebase --continue。'
    Write-Host '请确认冲突处理完成，再重新运行本脚本；不会自动跳过提交或强制推送。'
    if ($repoReady -and $gitExe) {
        & $gitExe --no-pager -c core.quotepath=false status --short --branch
    }
    exit 1
} finally {
    if ($messageFile -and (Test-Path -LiteralPath $messageFile)) {
        Remove-Item -LiteralPath $messageFile -Force -ErrorAction SilentlyContinue
    }
}
exit 0
