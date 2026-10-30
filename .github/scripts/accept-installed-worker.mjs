// Disposable Windows CI only. Runs the installed engine, without changing it.
// All task files, suppliers, databases and uploaded logs are synthetic fixtures.
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { createWriteStream } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

assert.equal(process.platform, 'win32', 'This diagnostic must run on native Windows.');
const engineDirectory = process.env.ACCEPTANCE_ENGINE;
const work = process.env.ACCEPTANCE_WORK;
assert.ok(engineDirectory && path.isAbsolute(engineDirectory));
assert.ok(work && path.isAbsolute(work));
const python = path.join(engineDirectory, 'python.exe');
const originalMain = path.join(engineDirectory, 'app', 'main.py');
const logs = path.join(work, 'diagnostics');
await mkdir(logs, { recursive: true });
const reportPath = path.join(logs, 'worker-acceptance-result.json');
const report = { status: 'RUNNING', installerVersion: process.env.EXPECTED_VERSION || '0.1.9', syntheticOnly: true, cases: [] };
const environment = { ...process.env, PYTHONUTF8: '1', PYTHONUNBUFFERED: '1' };
// Download authentication is scoped to an earlier workflow step. Do not forward
// a token if this helper is reused manually with one present.
for (const key of Object.keys(environment)) {
  if (['gh_token', 'github_token', 'electron_run_as_node', 'toolbox_data_dir', 'toolbox_python'].includes(key.toLowerCase())) delete environment[key];
}
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const saveReport = () => writeFile(reportPath, JSON.stringify(report, null, 2) + '\n');

class Engine {
  constructor(main, dataDirectory, label) {
    this.pending = new Map();
    this.events = [];
    this.buffer = '';
    this.stderr = '';
    this.closed = false;
    this.out = createWriteStream(path.join(logs, `${label}-stdout.jsonl`));
    this.err = createWriteStream(path.join(logs, `${label}-stderr.log`));
    this.command = { executable: python, args: ['-I', '-X', 'utf8', '-u', '-B', main, '--data-dir', dataDirectory], cwd: dataDirectory };
    this.child = spawn(python, this.command.args, { cwd: dataDirectory, env: environment, stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
    this.child.stdout.setEncoding('utf8');
    this.child.stderr.setEncoding('utf8');
    this.child.stdout.on('data', chunk => {
      this.out.write(chunk);
      this.buffer += chunk;
      let end;
      while ((end = this.buffer.indexOf('\n')) >= 0) {
        const line = this.buffer.slice(0, end).trim();
        this.buffer = this.buffer.slice(end + 1);
        if (!line) continue;
        let value;
        try { value = JSON.parse(line); }
        catch { this.protocolError = 'Engine stdout contained a non-JSON line.'; continue; }
        if (value.event) { this.events.push(value); continue; }
        const pending = this.pending.get(value.id);
        if (!pending) continue;
        this.pending.delete(value.id);
        clearTimeout(pending.timer);
        if (value.error) pending.reject(new Error(`${value.error.code}: ${value.error.message}`));
        else pending.resolve(value.result);
      }
    });
    this.child.stderr.on('data', chunk => { this.err.write(chunk); this.stderr = (this.stderr + chunk).slice(-40_000); });
    this.child.stdin.on('error', error => { this.inputError = error.message; });
    this.child.on('error', error => { this.launchError = error.message; this.rejectAll(error); });
    this.done = new Promise(resolve => this.child.once('close', (code, signal) => {
      this.closed = true;
      this.exit = { code, signal };
      this.rejectAll(new Error(`Engine closed: ${code ?? signal}`));
      this.out.end(); this.err.end(); resolve();
    }));
  }
  rejectAll(error) {
    for (const item of this.pending.values()) { clearTimeout(item.timer); item.reject(error); }
    this.pending.clear();
  }
  call(method, params = {}, timeout = 15_000) {
    if (this.closed || this.launchError) return Promise.reject(new Error(this.launchError || 'Engine is closed.'));
    const id = randomUUID();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error(`${method} timed out.`)); }, timeout);
      this.pending.set(id, { resolve, reject, timer });
      this.child.stdin.write(JSON.stringify({ id, method, params }) + '\n', error => {
        if (!error) return;
        const pending = this.pending.get(id);
        if (pending) { clearTimeout(pending.timer); this.pending.delete(id); pending.reject(error); }
      });
    });
  }
  async stop() {
    if (!this.closed) {
      try { await this.call('engine.shutdown', {}, 8_000); }
      catch (error) { this.shutdownError = error.message; }
    }
    await Promise.race([this.done, delay(5_000)]);
    if (!this.closed && this.child.pid) {
      // Only this diagnostic's owned child tree, never a global python kill.
      this.forcedCleanup = true;
      spawnSync(path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'taskkill.exe'), ['/PID', String(this.child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore', timeout: 10_000 });
      await Promise.race([this.done, delay(3_000)]);
    }
    if (!this.closed) { this.child.unref(); this.child.stdin.destroy(); this.child.stdout.destroy(); this.child.stderr.destroy(); this.out.end(); this.err.end(); }
  }
}

import { fileURLToPath } from 'node:url';

const fixtureScript = path.join(path.dirname(fileURLToPath(import.meta.url)), 'acceptance-workbook-fixtures.py');
const fixtureDirectory = path.join(work, '合成输入');
const dataDirectory = path.join(work, 'synthetic-data');
const outputDirectory = path.join(work, 'synthetic-output');
const fileHash = async filename => createHash('sha256').update(await readFile(filename)).digest('hex');
let engine;

function workbookAction(action, ...args) {
  const child = spawnSync(python, ['-I', '-X', 'utf8', '-u', '-B', fixtureScript, action, fixtureDirectory, ...args], {
    cwd: work, env: environment, encoding: 'utf8', windowsHide: true, timeout: 30_000,
  });
  assert.equal(child.status, 0, child.stderr || child.error?.message || `Workbook ${action} failed.`);
  return JSON.parse(child.stdout);
}

async function completedTask(label, request) {
  report.stage = label;
  const submitted = await engine.call('jobs.submit', request);
  const deadline = Date.now() + 45_000;
  let last;
  while (Date.now() < deadline) {
    const records = await engine.call('jobs.list', {}, 8_000);
    last = records.find(item => item.id === submitted.id);
    if (last && !['queued', 'running'].includes(last.status)) break;
    await delay(100);
  }
  report.cases.push({ label, job: last ?? submitted, events: engine.events.filter(event => event.event === 'job.changed' && event.data?.id === submitted.id) });
  await saveReport();
  assert.equal(last?.status, 'completed', `${label}: ${last?.error || last?.message || 'task timeout'}`);
  assert.equal(last.progress, 100);
  assert.ok(last.finishedAt);
  assert.equal((await engine.call('engine.info')).version, report.engineInfo.version);
  return last;
}

try {
  await mkdir(dataDirectory, { recursive: true });
  await mkdir(outputDirectory, { recursive: true });
  report.stage = 'generate-synthetic-workbooks-with-installed-python';
  const fixture = workbookAction('generate');
  assert.equal(fixture.suppliers.length, 60);
  const originalHashes = { source: await fileHash(fixture.source), template: await fileHash(fixture.template) };
  report.fixture = { source: fixture.source, template: fixture.template, supplierCount: 60, expected: fixture.expected, originalHashes };
  report.pathFile = await readFile(path.join(engineDirectory, 'python314._pth'), 'utf8');
  engine = new Engine(originalMain, dataDirectory, 'installed-acceptance');
  report.command = engine.command;
  report.enginePid = engine.child.pid;
  report.engineInfo = await engine.call('engine.info');
  report.stage = 'seed-sixty-synthetic-suppliers';
  for (const record of fixture.suppliers) await engine.call('reference.save', { kind: 'suppliers', record });
  const reference = await engine.call('reference.summary');
  assert.equal(reference.suppliers.total, 60);
  report.referenceSummary = reference;
  const input = path.join(work, '合成文件清单.txt');
  const text = 'SYNTHETIC WINDOWS INSTALLED WORKER ACCEPTANCE ONLY\n';
  await writeFile(input, text);
  const expectedHash = createHash('sha256').update(text).digest('hex');
  const manifestRequest = { toolId: 'file-manifest', inputPaths: [input], outputDirectory, parameters: { includeHash: true } };
  const outputs = new Set();
  async function manifest(label) {
    const job = await completedTask(label, manifestRequest);
    assert.equal(job.summary.fileCount, 1);
    assert.equal(job.summary.totalBytes, Buffer.byteLength(text));
    assert.equal(job.outputs.length, 1);
    assert.ok((await readFile(job.outputs[0], 'utf8')).includes(expectedHash));
    assert.equal(await readFile(input, 'utf8'), text);
    assert.ok(!outputs.has(job.outputs[0]), 'Repeated tasks must have independent output paths.');
    outputs.add(job.outputs[0]);
  }
  // Each task closes a real Windows worker pipe. Run enough times in one engine
  // lifetime to catch successful-result -> PeekNamedPipe(109) regressions.
  for (let index = 1; index <= 3; index++) await manifest(`file-manifest-${index}`);
  let firstWorkbook;
  let firstWorkbookHash;
  const voucherRequest = { toolId: 'voucher-import', inputPaths: [fixture.source, fixture.template], outputDirectory, parameters: fixture.parameters };
  for (let index = 1; index <= 2; index++) {
    const job = await completedTask(`invoice-voucher-repeat-${index}`, voucherRequest);
    for (const [key, value] of Object.entries(fixture.expected)) assert.equal(job.summary[key], value, `Unexpected voucher summary: ${key}`);
    assert.equal(job.summary['供应商资料版本'], reference.suppliers.revision);
    const workbook = job.outputs.find(filename => filename.toLowerCase().endsWith('.xlsx'));
    const audit = job.outputs.find(filename => filename.endsWith('凭证核对清单.csv'));
    assert.ok(workbook && audit);
    assert.ok(!outputs.has(workbook), 'Repeated voucher outputs must have independent paths.');
    outputs.add(workbook);
    report.cases.at(-1).workbookVerification = workbookAction('verify', workbook);
    assert.equal(await fileHash(fixture.source), originalHashes.source, 'Source workbook changed.');
    assert.equal(await fileHash(fixture.template), originalHashes.template, 'Template workbook changed.');
    const hash = await fileHash(workbook);
    if (!firstWorkbook) { firstWorkbook = workbook; firstWorkbookHash = hash; }
    else {
      assert.equal(await fileHash(firstWorkbook), firstWorkbookHash, 'Repeat submission modified the first task output.');
      assert.equal(hash, firstWorkbookHash, 'Identical inputs produced a different voucher workbook.');
    }
    await saveReport();
  }
  await manifest('file-manifest-after-repeated-vouchers');
  assert.equal(engine.protocolError ?? null, null);
  assert.equal(report.cases.length, 6);
  assert.equal((await engine.call('reference.summary')).suppliers.total, 60);
  const finalJobs = await engine.call('jobs.list');
  for (const result of report.cases) assert.equal(finalJobs.find(job => job.id === result.job.id)?.status, 'completed');
  report.stage = 'all-tasks-completed';
  report.status = 'PASS';
} catch (error) {
  report.status = 'FAIL';
  report.failure = error.stack || String(error);
  process.exitCode = 1;
} finally {
  if (engine) {
    await engine.stop();
    report.engineExit = engine.exit ?? null;
    report.engineShutdownError = engine.shutdownError ?? null;
    report.forcedCleanup = engine.forcedCleanup ?? false;
    report.stderrTail = engine.stderr;
    if (report.status === 'PASS' && (engine.exit?.code !== 0 || engine.forcedCleanup || engine.shutdownError)) {
      report.status = 'FAIL';
      report.failure = 'Tasks passed but the installed engine did not shut down cleanly.';
      process.exitCode = 1;
    }
  }
  await saveReport();
  console.log(JSON.stringify(report, null, 2));
}
