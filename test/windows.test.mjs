import { test } from 'node:test';
import assert from 'node:assert/strict';
import { runPowerShell, common, psQuote } from './helpers.mjs';

test('Windows mutex serializes independent processes and reports abandonment', { skip: process.platform !== 'win32' }, async () => {
  const output = await runPowerShell(`
$name = 'Local\\Cirya.WordPdf.Test.' + [Guid]::NewGuid().ToString('N')
$lease = [Cirya.WordPdf.MutexLease]::new($name, 5000)
$lease.Dispose()
$scriptPath = Join-Path $env:TEMP ('mutex-' + [Guid]::NewGuid().ToString('N') + '.ps1')
$source = '. ' + ${psQuote(psQuote(common))} + @'

$lease = [Cirya.WordPdf.MutexLease]::new($args[0], 10000)
[Console]::WriteLine('acquired')
[Console]::Out.Flush()
Start-Sleep -Seconds 30
'@
[IO.File]::WriteAllText($scriptPath, $source)
$info = [Diagnostics.ProcessStartInfo]::new()
$info.FileName = Join-Path $PSHOME 'powershell.exe'
$info.Arguments = '-NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "' + $scriptPath + '" ' + $name
$info.UseShellExecute = $false
$info.CreateNoWindow = $true
$info.RedirectStandardOutput = $true
$child = [Diagnostics.Process]::Start($info)
try {
    if ($child.StandardOutput.ReadLine() -ne 'acquired') { throw 'Child did not acquire mutex.' }
    $handle = [Threading.Mutex]::new($false, $name)
    try {
        $blocked = $false
        try { $other = [Cirya.WordPdf.MutexLease]::new($name, 100); $other.Dispose() }
        catch { $blocked = $true }
        if (-not $blocked) { throw 'Mutex did not block.' }
        $child.Kill(); $child.WaitForExit()
        $recovered = [Cirya.WordPdf.MutexLease]::new($name, 5000)
        try { if (-not $recovered.Abandoned) { throw 'Abandonment was not reported.' } }
        finally { $recovered.Dispose() }
    } finally { $handle.Dispose() }
    [Console]::WriteLine('passed')
} finally {
    if (-not $child.HasExited) { $child.Kill(); $child.WaitForExit() }
    $child.Dispose(); [IO.File]::Delete($scriptPath)
}`);
  assert.equal(output, 'passed');
});

test('Windows Job closes only its owned child; PID reuse refuses termination', { skip: process.platform !== 'win32' }, async () => {
  const output = await runPowerShell(`
function New-Sleeper {
    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = Join-Path $PSHOME 'powershell.exe'
    $info.Arguments = '-NoProfile -NonInteractive -Command "Start-Sleep -Seconds 30"'
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true
    return [Diagnostics.Process]::Start($info)
}
$child = New-Sleeper; $other = New-Sleeper
$job = $null
try {
    $otherIdentity = Get-Identity $other.Id
    [Cirya.WordPdf.WindowsNative]::StopOwned($other.Id, ($otherIdentity.created - 100), 'powershell.exe')
    if ($other.HasExited) { throw 'PID reuse guard failed.' }
    $job = [Cirya.WordPdf.OwnedJob]::new($child.Id)
    $job.Dispose(); $job = $null
    if (-not $child.WaitForExit(5000)) { throw 'Assigned child survived job close.' }
    if ($other.HasExited) { throw 'Unassigned child was stopped.' }
    [Console]::WriteLine('passed')
} finally {
    if ($null -ne $job) { $job.Dispose() }
    foreach ($process in @($child, $other)) {
        if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit() }
        $process.Dispose()
    }
}`);
  assert.equal(output, 'passed');
});

test('unknown Word ownership keeps recovery fail-closed', { skip: process.platform !== 'win32' }, async () => {
  const output = await runPowerShell(`
$folder = Join-Path $env:TEMP ('recovery-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($folder)
try {
    Write-JsonFile (Join-Path $folder 'state.json') @{ phase = 'starting_word' }
    $blocked = $false
    try { Clear-Worker ([PSCustomObject]@{ task_dir = $folder }) }
    catch { if ($_.Exception.Message -match 'recovery_required') { $blocked = $true } else { throw } }
    if (-not $blocked) { throw 'Unknown ownership was not blocked.' }
    [Console]::WriteLine('passed')
} finally { [IO.File]::Delete((Join-Path $folder 'state.json')); [IO.Directory]::Delete($folder) }
`);
  assert.equal(output, 'passed');
});

test('atomic publication preserves an existing PDF unless overwrite is explicit', { skip: process.platform !== 'win32' }, async () => {
  const output = await runPowerShell(`
$folder = Join-Path $env:TEMP ('publish-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($folder)
$source = Join-Path $folder 'stage.pdf'; $destination = Join-Path $folder 'existing.pdf'
try {
    [IO.File]::WriteAllText($source, 'new'); [IO.File]::WriteAllText($destination, 'old')
    $blocked = $false
    try { [Cirya.WordPdf.WindowsNative]::Publish($source, $destination, $false) }
    catch { $blocked = $true }
    if (-not $blocked -or [IO.File]::ReadAllText($destination) -ne 'old') { throw 'Existing PDF was modified.' }
    [Cirya.WordPdf.WindowsNative]::Publish($source, $destination, $true)
    if ([IO.File]::ReadAllText($destination) -ne 'new') { throw 'Explicit overwrite failed.' }
    [Console]::WriteLine('passed')
} finally {
    if ([IO.File]::Exists($source)) { [IO.File]::Delete($source) }
    [IO.File]::Delete($destination); [IO.Directory]::Delete($folder)
}
`);
  assert.equal(output, 'passed');
});
