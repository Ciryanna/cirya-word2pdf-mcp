param([Parameter(Mandatory = $true)][string] $RequestPath)
. (Join-Path $PSScriptRoot 'common.ps1')
$request = Read-JsonFile $RequestPath
$statePath = Join-Path ([IO.Path]::GetDirectoryName($RequestPath)) 'state.json'
$resultPath = Join-Path ([IO.Path]::GetDirectoryName($RequestPath)) 'result.json'
$state = @{ phase = 'before_word'; worker = (Get-Identity $PID) }
Write-JsonFile $statePath $state
[Cirya.WordPdf.WindowsNative]::WatchParent([int]$request.parent_pid)
$word = $null
$document = $null
$documents = $null
$options = $null
$job = $null
$owned = $false
$savedOptions = @{}
$result = $null
try {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Word COM requires STA.' }
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $state.phase = 'starting_word'
    Write-JsonFile $statePath $state
    $launchStarted = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() / 1000.0
    try { $word = New-Object -ComObject Word.Application }
    catch {
        $cause = $_.Exception
        while ($null -ne $cause) {
            if ($cause.HResult -in @(-2147221164, -2147221005)) {
                $state.phase = 'before_word'; Write-JsonFile $statePath $state; break
            }
            $cause = $cause.InnerException
        }
        throw
    }
    $documents = $word.Documents
    if ([int]$documents.Count -ne 0) { throw 'Word already contains documents; refusing to use it.' }
    $caption = $word.Caption
    $marker = 'WordPdfMcp-' + [Guid]::NewGuid().ToString('N')
    try {
        $word.Caption = $marker
        $wordPid = [Cirya.WordPdf.WindowsNative]::TaggedWordPid($marker)
    } finally { $word.Caption = $caption }
    $wordIdentity = Get-Identity $wordPid
    if ($wordPid -in $before -or $wordIdentity.name -ine 'WINWORD.EXE' -or
        $wordIdentity.created -lt ($launchStarted - 0.1) -or [int]$documents.Count -ne 0) {
        throw 'Word did not provide a verifiably new process; refusing to use it.'
    }
    $owned = $true
    $state.phase = 'word_identified'
    $state.word = $wordIdentity
    Write-JsonFile $statePath $state
    $job = [Cirya.WordPdf.OwnedJob]::new($wordPid)
    $state.phase = 'job_assigned'
    Write-JsonFile $statePath $state
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $options = $word.Options
    foreach ($name in @('UpdateLinksAtOpen', 'UpdateFieldsAtPrint', 'UpdateLinksAtPrint', 'SaveNormalPrompt')) {
        $savedOptions[$name] = $options.$name
        $options.$name = $false
    }
    $missing = [Type]::Missing
    # Explicit sentinels avoid password prompts; macros/links are disabled before opening.
    $noPassword = [Guid]::NewGuid().ToString('N')
    $document = [Cirya.WordPdf.WindowsNative]::Invoke($documents, 'Open', [object[]]@($request.input_path, $false, $true, $false,
        $noPassword, $noPassword, $false, $noPassword, $noPassword,
        $missing, $missing, $false, $false, $missing, $true, $missing))
    $null = [Cirya.WordPdf.WindowsNative]::Invoke($document, 'Repaginate', @())
    $pages = [int][Cirya.WordPdf.WindowsNative]::Invoke($document, 'ComputeStatistics', [object[]]@(2))
    $item = 0
    if ($request.include_markup) { $item = 7 }
    $api = 'ExportAsFixedFormat'
    # Check method availability; export failures never trigger an engine/API retry.
    if ([Cirya.WordPdf.WindowsNative]::HasMethod($document, 'ExportAsFixedFormat2')) {
        $api = 'ExportAsFixedFormat2'
        $null = [Cirya.WordPdf.WindowsNative]::Invoke($document, $api, [object[]]@($request.output_path, 17, $false, 0, 0, 1, 1,
            $item, $false, $true, 0, $true, $true, $false, $true, $missing))
    } else {
        $null = [Cirya.WordPdf.WindowsNative]::Invoke($document, $api, [object[]]@($request.output_path, 17, $false, 0, 0, 1, 1,
            $item, $false, $true, 0, $true, $true, $false, $missing))
    }
    $warnings = @('Fonts and equation/object appearance require visual acceptance testing.')
    if ($api -ne 'ExportAsFixedFormat2') { $warnings += 'Word lacks ExportAsFixedFormat2; original-image-quality option was unavailable.' }
    $result = @{ status = 'success'; api = $api; word_pages = $pages;
        word_version = [string]$word.Version; word_build = [string]$word.Build; warnings = $warnings }
} catch {
    $result = @{ status = 'error'; message = $_.Exception.Message }
} finally {
    try {
        if ($null -ne $document) { $null = [Cirya.WordPdf.WindowsNative]::Invoke($document, 'Close', [object[]]@(0)) }
    } catch { $result = @{ status = 'error'; message = $_.Exception.Message } }
    finally {
        try {
            if ($null -ne $word -and $owned) {
                try { foreach ($name in $savedOptions.Keys) { $options.$name = $savedOptions[$name] } }
                finally { $null = [Cirya.WordPdf.WindowsNative]::Invoke($word, 'Quit', [object[]]@(0)) }
            }
        } catch { $result = @{ status = 'error'; message = $_.Exception.Message } }
        finally {
            foreach ($comObject in @($document, $documents, $options, $word)) {
                if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
                }
            }
            $document = $documents = $options = $word = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            if ($null -ne $job) { $job.Dispose() }
        }
    }
}
Write-JsonFile $resultPath $result
if ($result.status -eq 'success') { exit 0 } else { exit 1 }
