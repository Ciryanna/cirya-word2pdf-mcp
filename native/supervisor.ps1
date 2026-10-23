param([Parameter(Mandatory = $true)][string] $RequestPath)
. (Join-Path $PSScriptRoot 'common.ps1')

function Invoke-Conversion($Request) {
    $timeout = 180
    $queueTimeout = 600
    if ($null -ne $Request.PSObject.Properties['timeout_seconds']) { $timeout = [int]$Request.timeout_seconds }
    if ($null -ne $Request.PSObject.Properties['queue_timeout_seconds']) { $queueTimeout = [int]$Request.queue_timeout_seconds }
    if ($timeout -lt 10 -or $timeout -gt 1800 -or $queueTimeout -lt 1 -or $queueTimeout -gt 3600) {
        throw 'invalid_timeout: Conversion must be 10-1800s; queue must be 1-3600s.'
    }
    $overwrite = $false
    $includeMarkup = $false
    if ($null -ne $Request.PSObject.Properties['overwrite']) { $overwrite = [bool]$Request.overwrite }
    if ($null -ne $Request.PSObject.Properties['include_markup']) { $includeMarkup = [bool]$Request.include_markup }
    $root = Get-StateRoot
    $lease = [Cirya.WordPdf.MutexLease]::new($script:MutexName, ($queueTimeout * 1000))
    try {
        [void][IO.Directory]::CreateDirectory($root)
        $journal = Join-Path $root 'pending.json'
        if ([IO.File]::Exists($journal)) {
            $previous = Read-JsonFile $journal
            Clear-Worker $previous
            Remove-TaskData $previous $root
            [IO.File]::Delete($journal)
        }
        $source = [IO.Path]::GetFullPath([string]$Request.input_path)
        $destination = [IO.Path]::GetFullPath([string]$Request.output_path)
        $extension = [IO.Path]::GetExtension($source).ToLowerInvariant()
        if (-not [IO.File]::Exists($source) -or $extension -notin @('.doc', '.docx', '.docm')) {
            throw 'invalid_input: Input must be an existing local .doc/.docx/.docm file.'
        }
        if ([IO.Path]::GetExtension($destination) -ine '.pdf' -or
            -not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($destination))) {
            throw 'invalid_output: Output must end in .pdf and its folder must exist.'
        }
        if ([IO.File]::Exists($destination) -and -not $overwrite) { throw 'output_exists: Output already exists.' }
        if ([IO.Directory]::Exists($destination)) { throw 'invalid_output: Output names a directory.' }
        $taskDirectory = Join-Path $root ('task-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($taskDirectory)
        $stage = Join-Path ([IO.Path]::GetDirectoryName($destination)) ('.wordpdf-' + [IO.Path]::GetFileName($taskDirectory) + '.pdf')
        $pending = [PSCustomObject]@{ task_dir = $taskDirectory; stage = $stage; supervisor = (Get-Identity $PID) }
        $process = $null
        $journalWritten = $false
        $clean = $false
        try {
            $originalHash = Get-Sha256 $source
            $copied = Join-Path $taskDirectory ('source' + $extension)
            [IO.File]::Copy($source, $copied)
            if ((Get-Sha256 $copied) -ne $originalHash) { throw 'source_changed: Source changed while copying; retry.' }
            $inventory = Get-PackageInventory $copied
            $workerRequestPath = Join-Path $taskDirectory 'request.json'
            Write-JsonFile $workerRequestPath @{ input_path = $copied; output_path = $stage;
                include_markup = $includeMarkup; parent_pid = $PID }
            Write-JsonFile $journal $pending
            $journalWritten = $true
            $startInfo = [Diagnostics.ProcessStartInfo]::new()
            $startInfo.FileName = Join-Path $PSHOME 'powershell.exe'
            $workerScript = Join-Path $PSScriptRoot 'worker.ps1'
            $startInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "' + $workerScript + '" -RequestPath "' + $workerRequestPath + '"'
            $startInfo.UseShellExecute = $false
            $startInfo.CreateNoWindow = $true
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            try { $process = [Diagnostics.Process]::Start($startInfo) }
            catch { $clean = $true; throw "worker_start_failed: $($_.Exception.Message)" }
            $pending | Add-Member -NotePropertyName worker -NotePropertyValue (Get-Identity $process.Id)
            Write-JsonFile $journal $pending
            $workerOutput = $process.StandardOutput.ReadToEndAsync()
            $workerErrors = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit($timeout * 1000)) {
                Stop-Owned $pending.worker
                throw "conversion_timeout: Word conversion exceeded ${timeout}s."
            }
            [IO.File]::WriteAllText((Join-Path $taskDirectory 'worker.log'), ($workerOutput.Result + $workerErrors.Result), [Text.Encoding]::UTF8)
            $resultPath = Join-Path $taskDirectory 'result.json'
            if (-not [IO.File]::Exists($resultPath)) {
                throw "word_failed: Worker exited without a result. $($workerErrors.Result)"
            }
            $result = Read-JsonFile $resultPath
            if ($process.ExitCode -ne 0 -or $result.status -ne 'success') { throw "word_failed: $($result.message)" }
            Clear-Worker $pending
            $clean = $true
            # This Node process only parses PDF structure; Word has already exited.
            # Keep the mutex held until validation and atomic publication are complete.
            $verification = & $Request.node_path $Request.validator_path $stage ([string]$result.word_pages)
            if ($LASTEXITCODE -ne 0) { throw 'invalid_pdf: PDF validation or Word/PDF page count check failed.' }
            $pdf = ConvertFrom-Json ($verification -join "`n")
            if ((Get-Sha256 $source) -ne $originalHash) { throw 'source_changed: Source changed during conversion; retry.' }
            [Cirya.WordPdf.WindowsNative]::Publish($stage, $destination, $overwrite)
            return @{
                status = 'success'; api = $result.api; word_pages = $result.word_pages;
                word_version = $result.word_version; word_build = $result.word_build; warnings = @($result.warnings);
                pages = $pdf.pages; page_sizes_pt = $pdf.page_sizes_pt; bytes = $pdf.bytes;
                input_sha256 = $originalHash; output_sha256 = (Get-Sha256 $destination);
                output_path = $destination; inventory = $inventory;
                recovered_abandoned_mutex = $lease.Abandoned; visual_fidelity_verified = $false
            }
        } finally {
            try {
                if ($journalWritten -and -not $clean) { Clear-Worker $pending; $clean = $true }
            } catch { throw "recovery_required: Cleanup could not be verified: $($_.Exception.Message). State: $taskDirectory" }
            finally {
                if ($null -ne $process) { $process.Dispose() }
                if ($clean -or -not $journalWritten) {
                    Remove-TaskData $pending $root
                    if ($journalWritten) { [IO.File]::Delete($journal) }
                }
            }
        }
    } finally { $lease.Dispose() }
}

try {
    $request = Read-JsonFile $RequestPath
    [Cirya.WordPdf.WindowsNative]::WatchParent([int]$request.parent_pid)
    if ($request.action -eq 'status') {
        $root = Get-StateRoot
        $pendingRecord = $null
        if ([IO.File]::Exists((Join-Path $root 'pending.json'))) { $pendingRecord = Join-Path $root 'pending.json' }
        $result = @{ platform = 'win32'; mutex = $script:MutexName; state_directory = $root;
            pending_record = $pendingRecord; runtime = 'node-powershell-com';
            note = 'Pending may mean a running task; Word and fidelity are not probed.' }
    } elseif ($request.action -eq 'convert') { $result = Invoke-Conversion $request }
    else { throw 'invalid_action: Unknown internal action.' }
} catch {
    $message = $_.Exception.Message
    $code = 'word_failed'
    if ($message -match '([a-z_]+):\s*(.*)') { $code = $Matches[1]; $message = $Matches[2] }
    $result = @{ status = 'error'; code = $code; message = $message }
}
[Console]::WriteLine((ConvertTo-Json -Depth 30 -Compress -InputObject $result))
