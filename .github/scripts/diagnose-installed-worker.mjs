// Disposable Windows CI only. Runs the installed engine, without changing it.
// All task files, suppliers, databases and uploaded logs are synthetic fixtures.
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { createWriteStream } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

assert.equal(process.platform, 'win32', 'This diagnostic must run on native Windows.');
const engineDirectory = process.env.DIAGNOSTIC_ENGINE;
const work = process.env.DIAGNOSTIC_WORK;
assert.ok(engineDirectory && path.isAbsolute(engineDirectory));
assert.ok(work && path.isAbsolute(work));
const python = path.join(engineDirectory, 'python.exe');
const originalMain = path.join(engineDirectory, 'app', 'main.py');
const logs = path.join(work, 'diagnostics');
await mkdir(logs, { recursive: true });
const reportPath = path.join(logs, 'worker-result.json');
const report = { status: 'RUNNING', installerVersion: '0.1.8', syntheticOnly: true, cases: [] };
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

async function runCase(mode, main, supplierCount) {
  const label = `${mode}-${supplierCount}-suppliers`;
  const started = Date.now();
  const caseRoot = path.join(work, label);
  const dataDirectory = path.join(caseRoot, 'data');
  const outputDirectory = path.join(caseRoot, 'output');
  await mkdir(dataDirectory, { recursive: true });
  await mkdir(outputDirectory, { recursive: true });
  const input = path.join(caseRoot, '合成文件 worker hash.txt');
  const contents = 'SYNTHETIC WINDOWS WORKER DIAGNOSTIC ONLY\n';
  await writeFile(input, contents);
  const expectedHash = createHash('sha256').update(contents).digest('hex');
  const result = { mode, supplierCount, label, stage: 'start', status: 'RUNNING' };
  report.cases.push(result);
  await saveReport();
  const engine = new Engine(main, dataDirectory, label);
  result.command = engine.command;
  result.enginePid = engine.child.pid;
  try {
    result.stage = 'bootstrap';
    result.info = await engine.call('engine.info');
    const tools = await engine.call('tools.list');
    assert.ok(tools.some(tool => tool.id === 'file-manifest'));
    result.stage = 'seed-synthetic-suppliers';
    for (let index = 1; index <= supplierCount; index++) {
      const number = String(index).padStart(4, '0');
      await engine.call('reference.save', { kind: 'suppliers', record: { code: `CI-SUP-${number}`, name: `合成诊断供应商${number}（仅测试，不是真实单位）`, category: 'SYNTHETIC_DIAGNOSTIC_ONLY' } });
    }
    const summary = await engine.call('reference.summary');
    assert.equal(summary.suppliers.total, supplierCount);
    result.seededSummary = summary;
    result.stage = 'submit-real-file-manifest-task';
    const submitted = await engine.call('jobs.submit', { toolId: 'file-manifest', inputPaths: [input], outputDirectory, parameters: { includeHash: true } });
    result.jobId = submitted.id;
    result.stage = 'wait-for-worker-result';
    const deadline = Date.now() + 35_000;
    let terminal;
    while (Date.now() < deadline) {
      const records = await engine.call('jobs.list', {}, 8_000);
      const record = records.find(item => item.id === submitted.id);
      if (record) result.lastJob = record;
      if (record && !['queued', 'running'].includes(record.status)) { terminal = record; break; }
      await delay(150);
    }
    assert.ok(terminal, 'Real worker task did not reach a terminal state.');
    result.jobEvents = engine.events.filter(event => event.event === 'job.changed' && event.data?.id === submitted.id);
    result.engineStillResponds = !!(await engine.call('engine.info'));
    if (terminal.status === 'completed') {
      assert.equal(terminal.summary.fileCount, 1);
      assert.equal(terminal.summary.totalBytes, Buffer.byteLength(contents));
      assert.equal(terminal.outputs.length, 1);
      const csv = await readFile(terminal.outputs[0], 'utf8');
      assert.ok(csv.includes(expectedHash), 'Worker result does not contain the expected input SHA256.');
      assert.equal(await readFile(input, 'utf8'), contents);
      result.inputAndOutputHashVerified = true;
      result.status = 'TASK_COMPLETED';
    } else {
      result.status = 'TASK_FAILED';
      result.failure = terminal.error || terminal.message;
    }
    result.stage = 'done';
  } catch (error) {
    result.status = 'HARNESS_OR_ENGINE_ERROR';
    result.failure = error.stack || String(error);
  } finally {
    await engine.stop();
    result.engineExit = engine.exit ?? null;
    result.engineShutdownError = engine.shutdownError ?? null;
    result.forcedCleanup = engine.forcedCleanup ?? false;
    result.protocolError = engine.protocolError ?? null;
    result.stderrTail = engine.stderr;
    result.seconds = Math.round((Date.now() - started) / 10) / 100;
    await saveReport();
  }
}

try {
  const source = await readFile(originalMain, 'utf8');
  const needle = 'multiprocessing.set_executable(str(windowless))';
  assert.equal(source.split(needle).length, 2, 'Installed entry point differs; do not silently alter another code path.');
  const consoleMain = path.join(work, 'console-worker-main.py');
  // This is the only diagnostic source change. The installed runtime and its
  // toolbox modules remain byte-for-byte unchanged; only the temporary main
  // selects the parent python.exe for multiprocessing bootstrap.
  await writeFile(consoleMain, source.replace(needle, 'multiprocessing.set_executable(sys.executable)'));
  report.installedMainSha256 = createHash('sha256').update(source).digest('hex');
  report.consoleVariant = { path: consoleMain, replacement: 'multiprocessing.set_executable(sys.executable)' };
  report.pathFile = await readFile(path.join(engineDirectory, 'python314._pth'), 'utf8');
  const probe = spawnSync(python, ['-I', '-X', 'utf8', '-B', '-c', 'import sys,json,multiprocessing.spawn; print(json.dumps({"version":sys.version,"executable":sys.executable,"baseExecutable":sys._base_executable,"path":sys.path,"isolated":sys.flags.isolated,"noSite":sys.flags.no_site,"spawnExecutable":multiprocessing.spawn.get_executable()}))'], { env: environment, cwd: work, encoding: 'utf8', windowsHide: true, timeout: 15_000 });
  report.runtimeProbe = { exitCode: probe.status, stdout: probe.stdout, stderr: probe.stderr, error: probe.error?.message ?? null };
  for (const [mode, main] of [['installed-pythonw', originalMain], ['temporary-console-worker', consoleMain]]) {
    for (const count of [0, 60]) await runCase(mode, main, count);
  }
  report.status = 'DIAGNOSTIC_RESULTS';
  report.completedCases = report.cases.filter(item => item.status === 'TASK_COMPLETED').length;
  report.failedCases = report.cases.filter(item => item.status === 'TASK_FAILED').length;
  report.harnessErrors = report.cases.filter(item => item.status === 'HARNESS_OR_ENGINE_ERROR').length;
  // A reproduced worker failure is a useful diagnostic result, never a claim
  // that task execution passed. Fail the workflow only if the harness itself
  // could not complete one or more of its four probes.
  if (report.harnessErrors) process.exitCode = 1;
} catch (error) {
  report.status = 'DIAGNOSTIC_SETUP_FAILED';
  report.failure = error.stack || String(error);
  process.exitCode = 1;
} finally {
  await saveReport();
  console.log(JSON.stringify(report, null, 2));
}
