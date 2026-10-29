// Formal-release validation on an ephemeral Windows runner. No installer is launched here:
// the installed 0.1.4 application's own installUpdate() must perform the upgrade.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawn, spawnSync } from 'node:child_process';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
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
const report = { status: 'RUNNING', runner: 'windows-2022', oldVersion: '0.1.4', expectedVersion, checks: {}, diagnostics: [] };
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
  throw new Error(`${label} timed out${lastError ? ' after a transient probe error' : ''}.`);
}
function psQuote(value) { return `'${String(value).replaceAll("'", "''")}'`; }
function powershell(script) {
  const result = spawnSync('pwsh', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', `$ErrorActionPreference='Stop'; [Console]::OutputEncoding=[System.Text.UTF8Encoding]::new(); ${script}`], { encoding: 'utf8', windowsHide: true, timeout: remaining(15_000) });
  if (result.status !== 0) throw new Error('Windows process/registry probe failed.');
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
  return powershell(`$items=@(Get-ItemProperty 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*' -ErrorAction SilentlyContinue | Where-Object { $_.UninstallString -eq ${psQuote(uninstallString)} }); if ($items.Count -eq 1) { @{version=$items[0].DisplayVersion; count=1} | ConvertTo-Json -Compress } else { @{version='';count=$items.Count} | ConvertTo-Json -Compress }`);
}
function safeAppEnvironment() {
  // Do not pass GitHub/API credentials into the application under test.
  const allowed = new Set(['systemroot', 'windir', 'comspec', 'path', 'pathext', 'temp', 'tmp', 'userprofile', 'username', 'userdomain', 'appdata', 'localappdata', 'programdata', 'allusersprofile', 'programfiles', 'programfiles(x86)', 'commonprogramfiles', 'commonprogramfiles(x86)', 'homedrive', 'homepath', 'number_of_processors', 'processor_architecture', 'os', 'public']);
  const environment = Object.fromEntries(Object.entries(process.env).filter(([key]) => allowed.has(key.toLowerCase())));
  return { ...environment, ELECTRON_ENABLE_LOGGING: '1' };
}
function startApp() {
  const processHandle = spawn(executable, ['--remote-debugging-address=127.0.0.1', '--remote-debugging-port=9222'], { env: safeAppEnvironment(), stdio: 'ignore', windowsHide: false });
  processHandle.on('error', () => { report.diagnostics.push('Application process launch failed.'); });
  return processHandle;
}
async function closeApp(pid) {
  powershell(`$p=Get-Process -Id ${Number(pid)} -ErrorAction SilentlyContinue; if ($p) { if (-not $p.CloseMainWindow()) { throw 'No main window available for graceful close' } }; 'true'`);
  await waitFor('graceful application close', () => !mainProcesses().some(item => item.pid === pid), 30_000);
}

class CDP {
  constructor(socket) {
    this.socket = socket;
    this.sequence = 0;
    this.pending = new Map();
    socket.addEventListener('message', event => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const entry = this.pending.get(message.id);
      if (!entry) return;
      clearTimeout(entry.timer); this.pending.delete(message.id);
      if (message.error) entry.reject(new Error('CDP request rejected.'));
      else entry.resolve(message.result);
    });
    socket.addEventListener('close', () => this.rejectAll());
    socket.addEventListener('error', () => this.rejectAll());
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
    if (result.exceptionDetails) throw new Error('Renderer application API rejected the validation request.');
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
  stage = 'launch old app and connect to bundled engine';
  launched = startApp();
  client = await connect();
  await bootstrap('0.1.4');
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
  report.checks.applicationDownloadVerified = true;

  stage = 'invoke one-click application update';
  let installState;
  try { installState = await client.evaluate('window.toolbox.installUpdate()', 45_000); }
  catch (error) { if (error.message !== 'CDP connection closed.') throw error; }
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
  report.diagnostics.push('Validation failed at the recorded stage. No financial data or database contents are included.');
  process.exitCode = 1;
} finally {
  clearTimeout(timeout);
  client?.close();
  try {
    for (const item of mainProcesses()) await closeApp(item.pid);
  } catch {
    report.diagnostics.push('Graceful cleanup did not complete on disposable runner.');
    if (report.status === 'PASS') { report.status = 'FAIL'; report.failedStage = 'graceful final cleanup'; process.exitCode = 1; }
  }
  launched?.unref();
  await writeFile(reportFile, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
}
