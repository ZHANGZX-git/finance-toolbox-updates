// Formal-release validation on an ephemeral Windows runner. No installer is launched here:
// the installed 0.1.4 application's own installUpdate() must perform the upgrade.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawn, spawnSync } from 'node:child_process';
import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';

const expectedVersion = process.env.EXPECTED_VERSION;
const expectedHash = process.env.EXPECTED_SHA256?.toLowerCase();
const repository = process.env.GH_REPO;
const work = process.env.VALIDATION_WORK;
const executable = process.env.VALIDATION_APP_EXE;
const userData = process.env.VALIDATION_USER_DATA;
assert.equal(process.platform, 'win32');
assert.match(expectedVersion ?? '', /^\d+\.\d+\.\d+$/);
assert.match(expectedHash ?? '', /^[0-9a-f]{64}$/);
assert.match(repository ?? '', /^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/);
for (const value of [work, executable, userData]) assert.ok(value && path.isAbsolute(value));
assert.equal(path.resolve(userData).toLowerCase(), path.join(process.env.APPDATA, '财务工具箱').toLowerCase(), 'Validation must use the production default userData identity.');
const expectedData = path.join(userData, 'engine');
const deadline = Date.now() + 9 * 60_000;
const reportFile = path.join(work, 'app-update-result.json');
const diagnosticsDirectory = path.join(work, 'diagnostics');
const observerScript = path.join(diagnosticsDirectory, 'observe-update.ps1');
const observerStop = path.join(diagnosticsDirectory, 'observer.stop');
const report = { status: 'RUNNING', runner: 'windows-2022', oldVersion: '0.1.4', expectedVersion, checks: {}, diagnostics: [] };
const applicationLogs = new Map();
let observer;
let launchSequence = 0;
let stage = 'initializing';
let client;
let launched;
let timedOut = false;
const timeout = setTimeout(() => { timedOut = true; }, 9 * 60_000);

function remaining(maximum = 30_000) {
  if (timedOut || Date.now() >= deadline) throw new Error('Overall validation timeout.');
  return Math.max(1, Math.min(maximum, deadline - Date.now()));
}
async function pause(ms = 500) { await new Promise(resolve => setTimeout(resolve, Math.min(ms, remaining(ms)))); }
async function waitFor(label, action, milliseconds) {
  const end = Math.min(deadline, Date.now() + milliseconds);
  let lastError;
  while (Date.now() < end) {
    remaining();
    try { const result = await action(); if (result) return result; }
    catch (error) { lastError = error; }
    await pause();
  }
  throw new Error(`${label} timed out${lastError ? `; last probe: ${lastError.message}` : ''}.`);
}
function psQuote(value) { return `'${String(value).replaceAll("'", "''")}'`; }
function powershell(script) {
  const result = spawnSync('pwsh', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', `$ErrorActionPreference='Stop'; [Console]::OutputEncoding=[System.Text.UTF8Encoding]::new(); ${script}`], { encoding: 'utf8', windowsHide: true, timeout: remaining(15_000) });
  if (result.status !== 0) throw new Error(`Windows process/registry probe failed (exit ${result.status}): ${(result.stderr || result.error?.message || '').slice(-6000)}`);
  const output = result.stdout.trim();
  return output ? JSON.parse(output) : null;
}
function appProcesses() {
  const engineExe = path.join(path.dirname(executable), 'resources', 'engine', 'python.exe');
  return powershell(`ConvertTo-Json -InputObject @(@(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq ${psQuote(executable)} -or $_.ExecutablePath -eq ${psQuote(engineExe)} }) | ForEach-Object { @{pid=[int]$_.ProcessId; parentPid=[int]$_.ParentProcessId; path=$_.ExecutablePath; commandLine=$_.CommandLine} }) -Depth 4 -Compress`) ?? [];
}
function mainProcesses() { return appProcesses().filter(item => item.path.toLowerCase() === executable.toLowerCase() && !/\s--type=/.test(item.commandLine ?? '')); }
function installedVersion() {
  // InstallLocation lives in NSIS's separate application key, not its Uninstall key.
  const uninstallString = `"${path.join(path.dirname(executable), 'Uninstall 财务工具箱.exe')}" /currentuser`;
  const result = powershell(`$items=@(Get-ItemProperty 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*' -ErrorAction SilentlyContinue | Where-Object { $_.UninstallString -eq ${psQuote(uninstallString)} }); @{version=$(if ($items.Count -eq 1) {$items[0].DisplayVersion} else {''});count=$items.Count;entries=@($items | Select-Object PSChildName,DisplayName,DisplayVersion,UninstallString,QuietUninstallString,InstallLocation)} | ConvertTo-Json -Depth 6 -Compress`);
  report.lastInstalledVersion = result;
  return result;
}
function safeAppEnvironment() {
  // Do not pass GitHub/API credentials into the application under test.
  const allowed = new Set(['systemroot', 'windir', 'comspec', 'path', 'pathext', 'temp', 'tmp', 'userprofile', 'username', 'userdomain', 'appdata', 'localappdata', 'programdata', 'allusersprofile', 'programfiles', 'programfiles(x86)', 'commonprogramfiles', 'commonprogramfiles(x86)', 'homedrive', 'homepath', 'number_of_processors', 'processor_architecture', 'os', 'public']);
  const environment = Object.fromEntries(Object.entries(process.env).filter(([key]) => allowed.has(key.toLowerCase())));
  return { ...environment, ELECTRON_ENABLE_LOGGING: '1' };
}
function startApp() {
  const label = `application-${++launchSequence}`;
  const processHandle = spawn(executable, ['--remote-debugging-address=127.0.0.1', '--remote-debugging-port=9222'], { env: safeAppEnvironment(), stdio: ['ignore', 'pipe', 'pipe'], windowsHide: false });
  for (const [name, stream] of [['stdout', processHandle.stdout], ['stderr', processHandle.stderr]]) {
    stream.setEncoding('utf8');
    stream.on('data', chunk => {
      const key = `${label}-${name}.log`;
      applicationLogs.set(key, ((applicationLogs.get(key) ?? '') + chunk).slice(-2 * 1024 * 1024));
    });
  }
  processHandle.on('error', error => { report.diagnostics.push(`Application launch failed: ${error.message}`); });
  processHandle.on('exit', (exitCode, signal) => { (report.applicationExits ??= []).push({ label, pid: processHandle.pid, exitCode, signal, at: new Date().toISOString() }); });
  return processHandle;
}
async function closeApp(pid) {
  powershell(`$p=Get-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue; if ($p) { if (-not $p.CloseMainWindow()) { throw 'No main window available for graceful close' } }; 'true'`);
  await waitFor('graceful application close', () => !mainProcesses().some(item => item.pid === pid), 30_000);
}

// Runs only on the disposable synthetic Windows runner. The observer does not
// launch, terminate, repair or click an installer; it records the real updater.
const observerSource = String.raw`param([ValidateSet('Observe','Snapshot')][string]$Mode='Observe')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new()
$diag=$env:VALIDATION_DIAGNOSTICS
$app=$env:VALIDATION_APP_EXE
$work=$env:VALIDATION_WORK
$cache=Join-Path $env:LOCALAPPDATA 'finance-toolbox-updater'
$helperPid=[int]$env:VALIDATION_HELPER_PID
$tracked=[System.Collections.Generic.HashSet[int]]::new()
$handles=@{}
$eventsFile=Join-Path $diag 'process-events.jsonl'
function Emit([string]$kind,$data) {
  $line=@{at=[DateTime]::UtcNow.ToString('o');kind=$kind;data=$data}|ConvertTo-Json -Depth 8 -Compress
  [IO.File]::AppendAllText($eventsFile,$line+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
}
function Related($p) {
  if ([int]$p.ProcessId -eq $PID -or [int]$p.ProcessId -eq $helperPid) {return $false}
  if ($tracked.Contains([int]$p.ParentProcessId)) {return $true}
  if ($p.Name -match '^(FinanceToolbox.*\.exe|财务工具箱\.exe|Uninstall 财务工具箱\.exe|old-uninstaller\.exe|elevate\.exe|Au_.*\.exe|Un_.*\.exe|ns[^\\]*\.(tmp|exe))$') {return $true}
  if ($p.ExecutablePath -and ($p.ExecutablePath.StartsWith($work,[StringComparison]::OrdinalIgnoreCase) -or $p.ExecutablePath.StartsWith($cache,[StringComparison]::OrdinalIgnoreCase))) {return $true}
  return $false
}
function ProcessRecord($p) {
  $result=@{pid=[int]$p.ProcessId;parentPid=[int]$p.ParentProcessId;name=$p.Name;path=$p.ExecutablePath;commandLine=$p.CommandLine;sessionId=$p.SessionId;created=$p.CreationDate}
  try {$live=Get-Process -Id $p.ProcessId -ErrorAction Stop;$result.mainWindowTitle=$live.MainWindowTitle;$result.mainWindowHandle=$live.MainWindowHandle.ToInt64()} catch {$result.probeError=$_.Exception.Message}
  return $result
}
function Track($p) {
  $id=[int]$p.ProcessId
  if ($tracked.Add($id)) {
    $record=ProcessRecord $p
    try {$live=Get-Process -Id $id -ErrorAction Stop;$null=$live.Handle;$handles[$id]=$live;$record.handleAcquired=$true} catch {$record.handleAcquired=$false;$record.handleError=$_.Exception.Message}
    Emit 'process-observed' $record
  }
}
if ($Mode -eq 'Snapshot') {
  $snapshot=@{at=[DateTime]::UtcNow.ToString('o');processes=@();registries=@();files=@();windows=@()}
  $all=@(Get-CimInstance Win32_Process)
  $related=@($all|Where-Object {Related $_})
  foreach ($p in $related) {$null=$tracked.Add([int]$p.ProcessId)}
  foreach ($p in $all) {if (Related $p) {$snapshot.processes+=ProcessRecord $p}}
  foreach ($root in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')) {
    foreach ($entry in @(Get-ItemProperty $root -ErrorAction SilentlyContinue|Where-Object {$_.DisplayName -like '*财务工具箱*' -or $_.UninstallString -like '*财务工具箱*'})) {
      $item=@{root=$root;key=$entry.PSChildName;displayName=$entry.DisplayName;version=$entry.DisplayVersion;uninstallString=$entry.UninstallString;quietUninstallString=$entry.QuietUninstallString;installLocation=$entry.InstallLocation;applicationKeys=@()}
      foreach ($hive in @('HKCU','HKLM')) {
        $appKey=$hive+':\Software\'+$entry.PSChildName
        $application=Get-ItemProperty $appKey -ErrorAction SilentlyContinue
        if ($application) {$item.applicationKeys+=@{key=$appKey;installLocation=$application.InstallLocation;keepShortcuts=$application.KeepShortcuts}}
      }
      $snapshot.registries+=$item
    }
  }
  $installDir=Split-Path -Parent $app
  foreach ($file in @(Get-ChildItem -LiteralPath $installDir -File -ErrorAction SilentlyContinue|Where-Object {$_.Name -like '*Uninstall*' -or $_.Name -like '*.bak' -or $_.Name -like '*.tmp' -or $_.Name -eq '财务工具箱.exe'})+@(Get-ChildItem -LiteralPath (Join-Path $cache 'pending') -File -ErrorAction SilentlyContinue)) {
    $item=@{path=$file.FullName;bytes=$file.Length;modifiedUtc=$file.LastWriteTimeUtc.ToString('o')}
    try {$item.sha256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()} catch {$item.hashError=$_.Exception.Message}
    $snapshot.files+=$item
  }
  try {
    Add-Type -TypeDefinition @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public static class ValidationWindows {
 public delegate bool Callback(IntPtr hwnd,IntPtr arg);
 [DllImport("user32.dll")]static extern bool EnumWindows(Callback fn,IntPtr arg);
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr hwnd,Callback fn,IntPtr arg);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int count);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr hwnd,StringBuilder text,int count);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr hwnd);
 static object Describe(IntPtr hwnd){uint pid;GetWindowThreadProcessId(hwnd,out pid);var text=new StringBuilder(4096);var cls=new StringBuilder(256);GetWindowText(hwnd,text,text.Capacity);GetClassName(hwnd,cls,cls.Capacity);return new{pid=pid,handle=hwnd.ToInt64(),text=text.ToString(),className=cls.ToString(),visible=IsWindowVisible(hwnd)};}
 public static object[] Collect(int[] wanted){var ids=new HashSet<int>(wanted);var output=new List<object>();EnumWindows((hwnd,arg)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(ids.Contains((int)pid)){var children=new List<object>();EnumChildWindows(hwnd,(child,x)=>{if(children.Count<120)children.Add(Describe(child));return children.Count<120;},IntPtr.Zero);output.Add(new{window=Describe(hwnd),children=children.ToArray()});}return true;},IntPtr.Zero);return output.ToArray();}
}
'@
    $snapshot.windows=@([ValidationWindows]::Collect([int[]]@($snapshot.processes|ForEach-Object {$_.pid})))
  } catch {$snapshot.windowError=$_.Exception.Message}
  try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $bounds=[Windows.Forms.SystemInformation]::VirtualScreen
    if ($bounds.Width -le 0 -or $bounds.Height -le 0 -or ([long]$bounds.Width*$bounds.Height) -gt 50000000) {throw 'Desktop screenshot bounds unavailable'}
    $bitmap=[Drawing.Bitmap]::new($bounds.Width,$bounds.Height)
    $graphics=[Drawing.Graphics]::FromImage($bitmap)
    try {$graphics.CopyFromScreen($bounds.Left,$bounds.Top,0,0,$bitmap.Size);$bitmap.Save((Join-Path $diag 'synthetic-runner-desktop.png'),[Drawing.Imaging.ImageFormat]::Png)} finally {$graphics.Dispose();$bitmap.Dispose()}
    $snapshot.screenshot='synthetic-runner-desktop.png'
  } catch {$snapshot.screenshotError=$_.Exception.Message}
  $snapshot|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $diag 'failure-windows-snapshot.json') -Encoding utf8
  exit 0
}
$prefix='finance-update-'+$PID
try {
  try {
    Register-CimIndicationEvent -Query 'SELECT * FROM Win32_ProcessStartTrace' -SourceIdentifier ($prefix+'-start')|Out-Null
    Register-CimIndicationEvent -Query 'SELECT * FROM Win32_ProcessStopTrace' -SourceIdentifier ($prefix+'-stop')|Out-Null
    Emit 'trace-ready' @{}
  } catch {Emit 'trace-subscription-error' @{message=$_.Exception.Message}}
  [IO.File]::WriteAllText((Join-Path $diag 'observer.ready'),'ready')
  $end=[DateTime]::UtcNow.AddMinutes(11)
  $lastSnapshot=''
  while (-not (Test-Path (Join-Path $diag 'observer.stop')) -and [DateTime]::UtcNow -lt $end) {
    foreach ($event in @(Get-Event -ErrorAction SilentlyContinue|Where-Object {$_.SourceIdentifier -like ($prefix+'-*')})) {
      $e=$event.SourceEventArgs.NewEvent
      $id=[int]$e.ProcessID
      if ($event.SourceIdentifier -eq ($prefix+'-start')) {
        $probe=@{ProcessId=$id;ParentProcessId=[int]$e.ParentProcessID;Name=$e.ProcessName;SessionId=$e.SessionID}
        if (Related $probe) {
          $null=$tracked.Add($id)
          Emit 'process-start-trace' $probe
          try {$live=Get-Process -Id $id -ErrorAction Stop;$null=$live.Handle;$handles[$id]=$live;Emit 'process-handle-acquired' @{pid=$id}} catch {Emit 'process-handle-unavailable' @{pid=$id;message=$_.Exception.Message}}
        }
      } elseif ($tracked.Contains($id)) {Emit 'process-stop-trace' @{pid=$id;parentPid=[int]$e.ParentProcessID;name=$e.ProcessName;exitStatus=$e.ExitStatus;sessionId=$e.SessionID}}
      Remove-Event -EventIdentifier $event.EventIdentifier -ErrorAction SilentlyContinue
    }
    $records=@()
    foreach ($p in @(Get-CimInstance Win32_Process)) {if (Related $p) {Track $p;$records+=ProcessRecord $p}}
    $current=$records|ConvertTo-Json -Depth 5 -Compress
    if ($current -ne $lastSnapshot) {Emit 'related-process-snapshot' $records;$lastSnapshot=$current}
    foreach ($id in @($handles.Keys)) {
      try {if ($handles[$id].HasExited) {Emit 'process-handle-exit' @{pid=$id;exitCode=$handles[$id].ExitCode};$handles[$id].Dispose();$handles.Remove($id)}} catch {Emit 'process-handle-error' @{pid=$id;message=$_.Exception.Message};$handles.Remove($id)}
    }
    Start-Sleep -Milliseconds 350
  }
} catch {Emit 'observer-error' @{message=$_.Exception.Message}}
finally {
  Get-EventSubscriber|Where-Object {$_.SourceIdentifier -like ($prefix+'-*')}|Unregister-Event -ErrorAction SilentlyContinue
  foreach ($handle in $handles.Values) {$handle.Dispose()}
  Emit 'observer-stopped' @{}
}
`;

function redactLog(value) {
  return String(value).replace(/(https?:\/\/[^\s?]+)\?[^\s]*/g, '$1?[query redacted]').replace(/(?:gh[pousr]_[A-Za-z0-9_]+|github_pat_[A-Za-z0-9_]+)/g, '[token redacted]');
}
async function startObserver() {
  await mkdir(diagnosticsDirectory, { recursive: true });
  await writeFile(observerScript, observerSource, 'utf8');
  observer = spawn('pwsh', ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', observerScript, '-Mode', 'Observe'], {
    env: { ...safeAppEnvironment(), VALIDATION_DIAGNOSTICS: diagnosticsDirectory, VALIDATION_WORK: work, VALIDATION_APP_EXE: executable, VALIDATION_HELPER_PID: String(process.pid) }, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true,
  });
  for (const [name, stream] of [['stdout', observer.stdout], ['stderr', observer.stderr]]) {
    stream.setEncoding('utf8'); stream.on('data', chunk => applicationLogs.set(`observer-${name}.log`, ((applicationLogs.get(`observer-${name}.log`) ?? '') + chunk).slice(-1024 * 1024)));
  }
  observer.on('error', error => report.diagnostics.push(`Process observer could not launch: ${error.message}`));
  try { await waitFor('diagnostic observer readiness', async () => { try { return (await stat(path.join(diagnosticsDirectory, 'observer.ready'))).isFile(); } catch { return false; } }, 15_000); }
  catch (error) { report.diagnostics.push(error.message); }
}
async function collectDiagnostics() {
  try { report.finalInstalledVersion = installedVersion(); } catch (error) { report.diagnostics.push(error.message); }
  const snapshot = spawnSync('pwsh', ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', observerScript, '-Mode', 'Snapshot'], {
    env: { ...safeAppEnvironment(), VALIDATION_DIAGNOSTICS: diagnosticsDirectory, VALIDATION_WORK: work, VALIDATION_APP_EXE: executable, VALIDATION_HELPER_PID: String(process.pid) }, encoding: 'utf8', windowsHide: true, timeout: 25_000,
  });
  report.snapshotExitCode = snapshot.status;
  if (snapshot.stderr || snapshot.error) await writeFile(path.join(diagnosticsDirectory, 'snapshot-errors.log'), redactLog(snapshot.stderr || snapshot.error.message));
  const logs = [];
  async function collect(directory, depth = 0) {
    if (depth > 2 || logs.length >= 20) return;
    let entries; try { entries = await readdir(directory, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      const filename = path.join(directory, entry.name);
      if (entry.isDirectory() && /^(logs?|crashpad)$/i.test(entry.name)) { await collect(filename, depth + 1); continue; }
      if (!entry.isFile() || !/\.log$|^chrome_debug\.log$|^debug\.log$|^update-settings\.json$|^update-info\.json$/i.test(entry.name)) continue;
      try {
        const info = await stat(filename);
        if (info.size > 2 * 1024 * 1024) continue;
        const destination = `captured-log-${logs.length}-${entry.name}`;
        await writeFile(path.join(diagnosticsDirectory, destination), redactLog(await readFile(filename, 'utf8')));
        logs.push({ source: filename, file: destination, bytes: info.size });
      } catch (error) { report.diagnostics.push(`Log capture failed: ${error.message}`); }
    }
  }
  await collect(userData);
  await collect(path.join(process.env.LOCALAPPDATA, 'finance-toolbox-updater', 'pending'));
  report.capturedLogs = logs;
}
async function finishObserver() {
  await writeFile(observerStop, 'stop\n');
  if (observer) {
    const end = Date.now() + 8000;
    while (observer.exitCode === null && Date.now() < end) await new Promise(resolve => setTimeout(resolve, 200));
    observer.stdout?.destroy(); observer.stderr?.destroy(); observer.unref();
  }
  for (const [filename, content] of applicationLogs) await writeFile(path.join(diagnosticsDirectory, filename), redactLog(content));
}

class CDP {
  constructor(socket) {
    this.socket = socket;
    this.sequence = 0;
    this.pending = new Map();
    socket.addEventListener('message', event => {
      const message = JSON.parse(event.data);
      if (message.method === 'Runtime.consoleAPICalled') {
        const text = message.params.args.map(item => item.value ?? '').join(' ');
        if (text.startsWith('VALIDATION_UPDATE_EVENT')) applicationLogs.set('renderer-update-events.log', ((applicationLogs.get('renderer-update-events.log') ?? '') + text + '\n').slice(-1024 * 1024));
      }
      if (!message.id) return;
      const entry = this.pending.get(message.id);
      if (!entry) return;
      clearTimeout(entry.timer); this.pending.delete(message.id);
      if (message.error) entry.reject(new Error(`CDP request rejected: ${JSON.stringify(message.error).slice(0, 8000)}`));
      else entry.resolve(message.result);
    });
    socket.addEventListener('close', () => this.rejectAll());
    socket.addEventListener('error', () => this.rejectAll());
    socket.send(JSON.stringify({ id: ++this.sequence, method: 'Runtime.enable', params: {} }));
  }
  rejectAll() {
    for (const entry of this.pending.values()) { clearTimeout(entry.timer); entry.reject(new Error('CDP connection closed.')); }
    this.pending.clear();
  }
  async evaluate(expression, maximum = 30_000) {
    const id = ++this.sequence;
    const result = await new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error('CDP request timed out.')); }, remaining(maximum));
      this.pending.set(id, { resolve, reject, timer });
      try { this.socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true } })); }
      catch { clearTimeout(timer); this.pending.delete(id); reject(new Error('CDP connection closed.')); }
    });
    if (result.exceptionDetails) throw new Error(`Renderer application API rejected: ${String(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text).slice(0, 12_000)}`);
    return result.result?.value;
  }
  close() { this.socket.close(); this.rejectAll(); }
}
async function connect() {
  return waitFor('renderer DevTools connection', async () => {
    const response = await fetch('http://127.0.0.1:9222/json/list', { signal: AbortSignal.timeout(1500) });
    if (!response.ok) return false;
    const target = (await response.json()).find(item => item.type === 'page' && item.url.startsWith('toolbox://app/') && item.webSocketDebuggerUrl);
    if (!target) return false;
    const socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => { socket.close(); reject(new Error('CDP connect timeout.')); }, remaining(4000));
      socket.addEventListener('open', () => { clearTimeout(timer); resolve(); }, { once: true });
      socket.addEventListener('error', () => { clearTimeout(timer); reject(new Error('CDP connect failed.')); }, { once: true });
    });
    const connected = new CDP(socket);
    try {
      if (!await connected.evaluate('Boolean(window.toolbox)', 4000)) { connected.close(); return false; }
      return connected;
    } catch { connected.close(); return false; }
  }, 45_000);
}
async function bootstrap(version) {
  const data = await client.evaluate('window.toolbox.getBootstrap()', 45_000);
  assert.equal(data.info.version, version, 'Installed renderer version differs.');
  assert.equal(data.info.platform, 'win32');
  assert.equal(data.info.packaged, true);
  assert.equal(data.info.pythonMode, 'bundled');
  assert.ok(data.info.engineVersion);
  assert.equal(path.resolve(data.info.dataDirectory).toLowerCase(), expectedData.toLowerCase(), 'Application changed its synthetic data-directory identity.');
  assert.ok(data.tools.some(item => item.id === 'voucher-import'));
  return data;
}
const supplier = { code: 'CI-UPGRADE-SUP-001', name: '合成升级验收供应商', category: 'TEST_ONLY' };
const inventory = { code: 'CI-UPGRADE-INV-001', name: '合成升级验收存货类别', category: 'TEST_ONLY' };

try {
  await mkdir(userData, { recursive: true });
  assert.equal((await readdir(userData)).length, 0, 'Production userData must be empty on this synthetic runner.');
  await writeFile(path.join(userData, 'synthetic-preservation.txt'), 'SYNTHETIC UPGRADE VALIDATION ONLY\n');
  assert.equal(installedVersion().version, '0.1.4');
  assert.equal(mainProcesses().length, 0, 'Unexpected existing app on disposable runner.');
  await startObserver();
  stage = 'launch old app and connect to bundled engine';
  launched = startApp();
  client = await connect();
  await bootstrap('0.1.4');
  await client.evaluate(`window.toolbox.onUpdateChanged(state => console.info('VALIDATION_UPDATE_EVENT', JSON.stringify({at:new Date().toISOString(),state}))); true`);
  report.checks.oldAppEngineConnected = true;
  const oldMain = await waitFor('old main process', () => mainProcesses().find(item => item.pid === launched.pid), 20_000);

  stage = 'create synthetic master-data sentinels';
  const saved = await client.evaluate(`(async () => ({supplier:await window.toolbox.saveMasterData('suppliers',${JSON.stringify(supplier)}),inventory:await window.toolbox.saveMasterData('inventoryCategories',${JSON.stringify(inventory)})}))()`);
  const beforeSummary = await client.evaluate('window.toolbox.getMasterDataSummary()');
  assert.equal(beforeSummary.suppliers.total, 1);
  assert.equal(beforeSummary.inventoryCategories.total, 1);
  report.checks.syntheticMasterDataSeeded = true;

  stage = 'check public update through old application';
  await client.evaluate(`window.toolbox.configureUpdates(${JSON.stringify({ repositoryUrl: `https://github.com/${repository}`, autoCheck: false })})`);
  const offered = await client.evaluate('window.toolbox.checkForUpdates()', 90_000);
  report.publicOffer = { status: offered.status, version: offered.version };
  assert.equal(offered.status, 'available', 'Published release was not offered to the real old application.');
  assert.equal(offered.version, expectedVersion);
  assert.equal(offered.canDownload, true);
  report.checks.publicReleaseOffered = true;

  stage = 'download update through old application';
  const downloaded = await client.evaluate('window.toolbox.downloadUpdate()', 5 * 60_000);
  report.applicationDownloadStatus = downloaded.status;
  assert.equal(downloaded.status, 'downloaded');
  assert.equal(downloaded.version, expectedVersion);
  assert.equal(downloaded.canInstall, true);
  const pendingDirectory = path.join(process.env.LOCALAPPDATA, 'finance-toolbox-updater', 'pending');
  const pendingFiles = (await readdir(pendingDirectory)).filter(name => name.toLowerCase().endsWith('.exe'));
  assert.equal(pendingFiles.length, 1, 'Expected a single downloaded installer in the fresh runner cache.');
  const actualHash = createHash('sha256').update(await readFile(path.join(pendingDirectory, pendingFiles[0]))).digest('hex');
  assert.equal(actualHash, expectedHash, 'Actual application-downloaded release asset hash mismatch.');
  report.downloadedInstallerSha256 = actualHash;
  report.cacheInstaller = { path: path.join(pendingDirectory, pendingFiles[0]), sha256: actualHash, bytes: (await stat(path.join(pendingDirectory, pendingFiles[0]))).size };
  report.updateStateBeforeInstall = await client.evaluate('window.toolbox.getUpdateState()');
  report.checks.applicationDownloadVerified = true;

  stage = 'invoke one-click application update';
  let installState;
  try { installState = await client.evaluate('window.toolbox.installUpdate()', 45_000); }
  catch (error) { if (error.message !== 'CDP connection closed.') throw error; }
  report.installUpdateReply = installState ?? 'CDP connection closed during quit';
  if (installState) assert.equal(installState.status, 'installing', 'Application refused or failed to begin installation.');
  client.close(); client = null;
  await waitFor('old application exit', () => launched.exitCode !== null || launched.signalCode !== null, 45_000);
  report.checks.oldApplicationExited = true;

  stage = 'wait for installer registration and automatic application restart';
  await waitFor('upgraded registry version', () => installedVersion().version === expectedVersion, 150_000);
  report.checks.registryVersionUpdated = true;
  const automaticallyRestarted = await waitFor('automatic application restart', () => mainProcesses().find(item => item.pid !== oldMain.pid), 90_000);
  // NSIS re-launches with --updated, not the old DevTools flags. Check its real
  // engine command line before a controlled CDP restart, confirming the same
  // production default userData identity without relying on dev environment propagation.
  await waitFor('automatically restarted engine uses the original synthetic data directory', () => appProcesses().some(item => item.parentPid === automaticallyRestarted.pid && item.path.toLowerCase().endsWith('\\python.exe') && item.commandLine?.toLowerCase().includes(expectedData.toLowerCase())), 45_000);
  report.checks.automaticRestartObserved = true;
  report.checks.automaticRestartDataIdentityPreserved = true;

  stage = 'gracefully reopen updated app for renderer inspection';
  await closeApp(automaticallyRestarted.pid);
  launched = startApp();
  client = await connect();
  await bootstrap(expectedVersion);
  const after = await client.evaluate(`(async () => ({supplier:await window.toolbox.listMasterData('suppliers',${JSON.stringify(supplier.code)},0),inventory:await window.toolbox.listMasterData('inventoryCategories',${JSON.stringify(inventory.code)},0),summary:await window.toolbox.getMasterDataSummary()}))()`);
  for (const [key, expected] of [['supplier', supplier], ['inventory', inventory]]) {
    assert.equal(after[key].total, 1);
    const record = after[key].records[0];
    assert.equal(record.id, saved[key].id, 'Master-data identity changed across upgrade.');
    for (const field of ['code', 'name', 'category']) assert.equal(record[field], expected[field]);
  }
  assert.equal(after.summary.suppliers.total, 1);
  assert.equal(after.summary.inventoryCategories.total, 1);
  assert.equal(await readFile(path.join(userData, 'synthetic-preservation.txt'), 'utf8'), 'SYNTHETIC UPGRADE VALIDATION ONLY\n');
  report.checks.updatedAppEngineConnected = true;
  report.checks.masterDataIdentityAndValuesPreserved = true;
  report.checks.userDataSentinelPreserved = true;
  report.status = 'PASS';
} catch (error) {
  report.status = 'FAIL';
  report.failedStage = stage;
  report.failureClass = error.code ?? error.name ?? 'Error';
  report.failureMessage = redactLog(error.message).slice(0, 16_000);
  report.failureStack = redactLog(error.stack).slice(0, 20_000);
  report.diagnostics.push('Synthetic runner only; detailed installer diagnostics captured before cleanup. No financial workbooks or SQLite contents are included.');
  process.exitCode = 1;
} finally {
  clearTimeout(timeout);
  await mkdir(diagnosticsDirectory, { recursive: true });
  if (client) {
    try { report.lastReadableUpdateState = await client.evaluate('window.toolbox.getUpdateState()', 4000); }
    catch (error) { report.diagnostics.push(`Final update-state read: ${error.message}`); }
  }
  try { await collectDiagnostics(); } catch (error) { report.diagnostics.push(`Diagnostics capture: ${error.message}`); }
  client?.close();
  try {
    for (const item of mainProcesses()) await closeApp(item.pid);
  } catch {
    report.diagnostics.push('Graceful cleanup did not complete on disposable runner.');
    if (report.status === 'PASS') { report.status = 'FAIL'; report.failedStage = 'graceful final cleanup'; process.exitCode = 1; }
  }
  try { await finishObserver(); } catch (error) { report.diagnostics.push(`Observer cleanup: ${error.message}`); }
  launched?.unref();
  launched?.stdout?.destroy(); launched?.stderr?.destroy();
  await writeFile(reportFile, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
}
