Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$null = Add-Type -Path (Join-Path $PSScriptRoot 'WindowsNative.cs')
$null = Add-Type -AssemblyName System.IO.Compression.FileSystem
$script:MutexName = 'Local\Elwina.WordPdf.Worker.v1'

function Read-JsonFile([string] $Path) {
    try { return (ConvertFrom-Json ([IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8))) }
    catch { throw "recovery_required: Cannot read recovery record ${Path}: $($_.Exception.Message)" }
}
function Write-JsonFile([string] $Path, $Value) {
    $temporary = $Path + '.tmp'
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -Depth 30 -Compress -InputObject $Value))
    $stream = [IO.File]::Open($temporary, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
    [Cirya.WordPdf.WindowsNative]::Publish($temporary, $Path, $true)
}
function Get-Identity([int] $ProcessId) {
    $process = [Diagnostics.Process]::GetProcessById($ProcessId)
    try { return @{ pid = $ProcessId; created = [Cirya.WordPdf.WindowsNative]::Created($process); name = $process.ProcessName + '.exe' } }
    finally { $process.Dispose() }
}
function Stop-Owned($Record) {
    if ($null -ne $Record) {
        [Cirya.WordPdf.WindowsNative]::StopOwned([int]$Record.pid, [double]$Record.created, [string]$Record.name)
    }
}
function Get-Sha256([string] $Path) {
    $algorithm = [Security.Cryptography.SHA256]::Create()
    $stream = [IO.File]::OpenRead($Path)
    try { return ([BitConverter]::ToString($algorithm.ComputeHash($stream))).Replace('-', '').ToLowerInvariant() }
    finally { $stream.Dispose(); $algorithm.Dispose() }
}
function Get-StateRoot {
    $sessionId = [Diagnostics.Process]::GetCurrentProcess().SessionId
    return (Join-Path $env:LOCALAPPDATA "WordPdfMcp\session-$sessionId")
}
function Clear-Worker($Pending) {
    $statePath = Join-Path $Pending.task_dir 'state.json'
    $workerProperty = $Pending.PSObject.Properties['worker']
    if ($null -ne $workerProperty) { Stop-Owned $Pending.worker }
    if (-not [IO.File]::Exists($statePath)) {
        if ($null -ne $workerProperty) { return }
        throw 'recovery_required: Worker startup ownership is unknown; inspect the pending record.'
    }
    $state = Read-JsonFile $statePath
    if ($null -ne $state.PSObject.Properties['worker']) { Stop-Owned $state.worker }
    if ($null -ne $state.PSObject.Properties['word']) {
        if ($state.word.name -ine 'WINWORD.EXE') { throw 'recovery_required: Invalid Word process identity.' }
        Stop-Owned $state.word
    } elseif ($state.phase -ne 'before_word') {
        throw 'recovery_required: Word may have started before its identity was recorded. Follow docs/RECOVERY.md.'
    }
}
function Remove-TaskData($Pending, [string] $Root) {
    # Only remove a validated direct task-* child of this session's state directory.
    $taskPath = [IO.Path]::GetFullPath([string]$Pending.task_dir)
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($taskPath) -ine $rootPath -or [IO.Path]::GetFileName($taskPath) -notlike 'task-*') {
        throw 'recovery_required: Task directory is outside the session state directory.'
    }
    if ([IO.Directory]::Exists($taskPath)) {
        $taskItem = Get-Item -LiteralPath $taskPath -Force
        if (($taskItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'recovery_required: Refusing to remove a task directory that is a reparse point.'
        }
        Remove-Item -LiteralPath $taskPath -Recurse -Force
    }
    # Only remove the uniquely named stage from this record, never the published output.
    if ($null -ne $Pending.PSObject.Properties['stage']) {
        $stagePath = [IO.Path]::GetFullPath([string]$Pending.stage)
        if ([IO.Path]::GetFileName($stagePath) -eq ('.wordpdf-' + [IO.Path]::GetFileName($taskPath) + '.pdf')) {
            if ([IO.File]::Exists($stagePath)) { [IO.File]::Delete($stagePath) }
        }
    }
}
function Get-PackageInventory([string] $Path) {
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq '.doc') { return @{ format = 'doc'; warnings = @('Legacy DOC inventory is not implemented.') } }
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        $count = @($archive.Entries | Where-Object { $_.FullName.StartsWith('word/embeddings/') }).Count
        $warnings = @()
        if ($count -gt 0) { $warnings = @('Embedded objects require visual acceptance testing.') }
        return @{ format = $extension.TrimStart('.'); embedded_parts = $count; warnings = $warnings }
    } catch { return @{ warnings = @('Package inventory unavailable (encrypted or invalid OOXML file).') } }
    finally { if ($null -ne $archive) { $archive.Dispose() } }
}
