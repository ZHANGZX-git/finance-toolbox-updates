// Native Windows CI only. All inputs and databases below are synthetic.
// Exercises the reviewed installed release without changing its engine files.
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { createWriteStream } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

assert.equal(process.platform, 'win32', 'Native Windows is required for installed release acceptance.');
const engineDirectory = process.env.ACCEPTANCE_ENGINE;
const work = process.env.ACCEPTANCE_WORK;
assert.ok(engineDirectory && path.isAbsolute(engineDirectory));
assert.ok(work && path.isAbsolute(work));
assert.equal(process.env.EXPECTED_VERSION, '0.1.13', 'This acceptance targets the reviewed 0.1.13 release.');
const python = path.join(engineDirectory, 'python.exe');
const originalMain = path.join(engineDirectory, 'app', 'main.py');
const logs = path.join(work, 'diagnostics');
await mkdir(logs, { recursive: true });
const reportPath = path.join(logs, 'v013-voucher-acceptance-result.json');
const report = { status: 'RUNNING', installerVersion: '0.1.13', nativePlatform: process.platform, syntheticOnly: true, cases: [], engines: [] };
const environment = { ...process.env, PYTHONUTF8: '1', PYTHONUNBUFFERED: '1' };
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

const fixtureScript = path.join(path.dirname(fileURLToPath(import.meta.url)), 'accept-v013-fixtures.py');
const fixtureDirectory = path.join(work, 'v013-synthetic-inputs');
const dataDirectory = path.join(work, 'v013-synthetic-data');
const outputDirectory = path.join(work, 'v013-synthetic-output');
const fileHash = async filename => createHash('sha256').update(await readFile(filename)).digest('hex');
let engine;

function pythonAction(action, ...args) {
  const child = spawnSync(python, ['-I', '-X', 'utf8', '-u', '-B', fixtureScript, action, ...args], {
    cwd: work, env: environment, encoding: 'utf8', windowsHide: true, timeout: 60_000,
  });
  assert.equal(child.status, 0, child.stderr || child.error?.message || `Synthetic ${action} failed.`);
  return JSON.parse(child.stdout);
}

async function allReferences(kind) {
  const result = await engine.call('reference.list', { kind, limit: 100 });
  assert.equal(result.records.length, result.total);
  return result;
}

async function task(label, request, status = 'completed') {
  report.stage = label;
  const submitted = await engine.call('jobs.submit', request);
  const deadline = Date.now() + 90_000;
  let last;
  while (Date.now() < deadline) {
    last = (await engine.call('jobs.list')).find(item => item.id === submitted.id);
    if (last && !['queued', 'running'].includes(last.status)) break;
    await delay(70);
  }
  const item = { label, job: last || submitted, events: engine.events.filter(event => event.event === 'job.changed' && event.data?.id === submitted.id) };
  report.cases.push(item);
  if (last?.status !== 'completed') {
    item.issueReports = [];
    for (const filename of last?.outputs || []) {
      if (filename.endsWith('.csv')) item.issueReports.push({ path: filename, text: (await readFile(filename, 'utf8')).slice(0, 64 * 1024) });
    }
  }
  await saveReport();
  assert.equal(last?.status, status, `${label}: ${last?.error || last?.message || 'timeout'}`);
  assert.equal(last.progress, 100);
  assert.ok(last.finishedAt);
  assert.equal(last.toolVersion, '1.10.0');
  assert.equal(engine.protocolError ?? null, null);
  return last;
}

async function stopEngine() {
  if (!engine) return;
  const current = engine;
  await current.stop();
  report.engines.push({ command: current.command, exit: current.exit, protocolError: current.protocolError || null,
    forcedCleanup: current.forcedCleanup || false, shutdownError: current.shutdownError || null, stderrTail: current.stderr });
  engine = null;
  assert.equal(current.exit?.code, 0, current.stderr);
  assert.equal(current.forcedCleanup || false, false);
  assert.equal(current.shutdownError || null, null);
}

try {
  await mkdir(outputDirectory, { recursive: true });
  report.stage = 'audit-installed-version-and-packaged-bytes';
  report.installation = pythonAction('install', engineDirectory, process.env.EXPECTED_VERSION);
  report.stage = 'generate-synthetic-workbooks-and-legacy-databases';
  const fixture = pythonAction('generate', fixtureDirectory, dataDirectory);
  report.syntheticFixture = { sourceCategories: Object.keys(fixture.sources), supplierCount: fixture.suppliers.length, legacyCustomer: fixture.legacyCustomer, preservedHistoryId: fixture.history.id };
  const sourceHashes = {};
  for (const [category, source] of Object.entries(fixture.sources)) sourceHashes[category] = await fileHash(source);
  engine = new Engine(originalMain, dataDirectory, 'v013-first-start');
  report.engineInfo = await engine.call('engine.info');
  const tools = await engine.call('tools.list');
  const voucher = tools.find(item => item.id === 'voucher-import');
  assert.equal(voucher.version, '1.10.0');
  const templates = voucher.input.slots.find(item => item.key === 'template').fixedPaths;
  assert.deepEqual(Object.keys(templates).sort(), ['invoice', 'purchase', 'sales-cost', 'sales-revenue']);
  const templateHashes = {};
  for (const [category, template] of Object.entries(templates)) {
    assert.ok(path.resolve(template).startsWith(path.resolve(engineDirectory) + path.sep));
    templateHashes[category] = await fileHash(template);
  }
  report.fixedTemplates = templateHashes;
  const old = (await allReferences('customers')).records;
  assert.equal(old.length, 1);
  for (const [field, value] of Object.entries(fixture.legacyCustomer)) assert.equal(old[0][field], value);
  assert.equal(old[0].sourceName || '', '');
  assert.deepEqual((await engine.call('jobs.list')).find(job => job.id === fixture.history.id), fixture.history);
  assert.equal((await engine.call('reference.summary')).customers.revision, 7, 'A schema-only migration must not alter reference revisions.');
  report.aliasSchemaMigration = { preservedId: old[0].id, preservedCode: old[0].code, preservedLegalName: old[0].name, preservedTimestamp: old[0].updatedAt, unchangedRevision: 7, historyPreserved: true };
  const preview = await engine.call('reference.preview', { kind: 'customers', path: fixture.customerImport });
  assert.equal(preview.issueCount, 0);
  assert.equal(preview.addCount, 2);
  assert.equal(preview.updateCount, 1);
  await engine.call('reference.commit', { token: preview.token });
  const customerRecords = await allReferences('customers');
  const editedCustomer = customerRecords.records.find(record => record.id === fixture.legacyCustomer.id);
  assert.equal(editedCustomer.code, fixture.legacyCustomer.code);
  assert.equal(editedCustomer.name, fixture.legacyCustomer.name);
  assert.equal(editedCustomer.sourceName, 'CI合成门店A');
  const supplierRecords = await allReferences('suppliers');
  for (const record of fixture.suppliers) {
    const existing = supplierRecords.records.find(item => item.code === record.code);
    await engine.call('reference.save', { kind: 'suppliers', record: { ...record, ...(existing ? { id: existing.id } : {}) } });
  }
  assert.equal((await engine.call('reference.summary')).suppliers.total, 60);
  const request = (category, extra = {}) => ({ toolId: 'voucher-import', inputPaths: [fixture.sources[category], templates[category]], outputDirectory,
    parameters: { ...(category === 'invoice' ? fixture.invoice.parameters : { category }), ...extra } });
  const noOutputTax = await task('revenue-output-tax-missing-does-not-use-input', request('sales-revenue'), 'needs_attention');
  assert.ok(noOutputTax.attentionItems.some(item => item.kind === 'taxRates' && item.taxDirection === 'output'));
  assert.ok(!noOutputTax.outputs.some(filename => filename.endsWith('.xlsx')));
  for (const record of fixture.outputRates) await engine.call('reference.save', { kind: 'taxRates', record });
  const strictPurchase = await task('purchase-default-strict-blocks-missing-supplier', request('purchase'), 'needs_attention');
  assert.ok(strictPurchase.attentionItems.some(item => item.kind === 'suppliers' && item.reason === 'missing'));
  assert.ok(!strictPurchase.outputs.some(filename => filename.endsWith('.xlsx')));
  const referenceBefore = await engine.call('reference.summary');
  report.referenceSummaryBeforeTasks = referenceBefore;
  const firstOutputs = {};
  async function checked(category, label, extra = {}, supplierCode = 'CI-SUP-0001') {
    const job = await task(label, request(category, extra));
    const xlsx = job.outputs.find(filename => filename.endsWith('.xlsx'));
    const audit = job.outputs.find(filename => filename.endsWith('核对清单.csv'));
    const skipped = job.outputs.find(filename => filename.endsWith('采购跳过清单.csv'));
    assert.ok(xlsx && audit);
    report.cases.at(-1).workbookVerification = pythonAction('verify', category, templates[category], xlsx, audit, supplierCode, ...(skipped ? [skipped] : []));
    assert.equal(await fileHash(fixture.sources[category]), sourceHashes[category]);
    assert.equal(await fileHash(templates[category]), templateHashes[category]);
    for (const [kind, summaryKey] of [['suppliers','供应商资料版本'],['inventoryCategories','存货类别资料版本'],['customers','客户资料版本'],['taxRates','税率资料版本']]) {
      if (summaryKey in job.summary) assert.equal(job.summary[summaryKey], (await engine.call('reference.summary'))[kind].revision);
    }
    if (category === 'purchase') {
      assert.equal(job.summary['采购明细数'], 5);
      assert.equal(job.summary['实际处理明细数'], 4);
      assert.equal(job.summary['跳过明细数'], 1);
    }
    if (!firstOutputs[category]) firstOutputs[category] = { path: xlsx, hash: await fileHash(xlsx), jobId: job.id };
    await saveReport();
    return job;
  }
  await checked('invoice', 'invoice-55-first');
  await checked('purchase', 'purchase-skip-missing-supplier-only', { skipMissingSuppliers: true });
  await checked('sales-cost', 'sales-cost-four-groups-and-duplicates');
  await checked('sales-revenue', 'sales-revenue-aliases-and-raw-group-rounding');
  await checked('invoice', 'invoice-55-repeat');
  const catalog = await allReferences('suppliers');
  const mainSupplier = catalog.records.find(item => item.code === 'CI-SUP-0001');
  const invoiceBeforeEdit = firstOutputs.invoice;
  const updatedSupplier = await engine.call('reference.save', { kind: 'suppliers', record: { id: mainSupplier.id, code: 'CI-SUP-0001-UPDATED', name: mainSupplier.name } });
  const changed = await checked('invoice', 'invoice-sees-new-reference-snapshot', {}, updatedSupplier.code);
  assert.ok(changed.summary['供应商资料版本'] > referenceBefore.suppliers.revision);
  assert.equal(await fileHash(invoiceBeforeEdit.path), invoiceBeforeEdit.hash, 'Editing saved references must not rewrite a completed job.');
  report.referenceSnapshot = { olderJobId: invoiceBeforeEdit.jobId, newerJobId: changed.id, olderRevision: referenceBefore.suppliers.revision,
    newerRevision: changed.summary['供应商资料版本'], previousOutputBytesUnchanged: true, workerObservedUpdatedSupplierCode: true };
  const savedSummary = await engine.call('reference.summary');
  const savedCustomers = await allReferences('customers');
  const savedJobs = await engine.call('jobs.list');
  assert.equal(engine.protocolError || null, null);
  await stopEngine();
  engine = new Engine(originalMain, dataDirectory, 'v013-restart');
  await engine.call('engine.info');
  assert.deepEqual(await engine.call('reference.summary'), savedSummary);
  assert.deepEqual(await allReferences('customers'), savedCustomers);
  const restoredJobs = await engine.call('jobs.list');
  for (const oldJob of savedJobs) assert.deepEqual(restoredJobs.find(job => job.id === oldJob.id), oldJob);
  await checked('invoice', 'invoice-after-engine-restart', {}, updatedSupplier.code);
  await checked('purchase', 'purchase-after-engine-restart', { skipMissingSuppliers: true }, updatedSupplier.code);
  await checked('sales-cost', 'sales-cost-after-engine-restart');
  await checked('sales-revenue', 'sales-revenue-after-engine-restart');
  for (const previous of Object.values(firstOutputs)) assert.equal(await fileHash(previous.path), previous.hash);
  report.restart = { savedReferencesAndAliasesPreserved: true, previousJobsPreserved: savedJobs.length, allFourFlowsSucceededAfterRestart: true };
  report.installationAfter = pythonAction('install', engineDirectory, process.env.EXPECTED_VERSION);
  await stopEngine();
  report.stage = 'all-four-installed-windows-flows-passed';
  report.status = 'PASS';
} catch (error) {
  report.status = 'FAIL';
  report.failure = error.stack || String(error);
  process.exitCode = 1;
} finally {
  if (engine) {
    try { await stopEngine(); }
    catch (error) { report.status = 'FAIL'; report.shutdownFailure = error.stack || String(error); process.exitCode = 1; }
  }
  await saveReport();
  console.log(JSON.stringify(report, null, 2));
}

