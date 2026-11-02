"""Disposable, synthetic-only v0.1.17 Windows installed-engine acceptance data.

All finance inputs are created here with stdlib ZIP/XML. Verification imports
only the installed release's XLSX reader, never a checkout's finance handlers.
"""
from collections import defaultdict
from decimal import Decimal
import csv
import hashlib
import importlib.util
import json
from pathlib import Path
import sqlite3
import struct
import sys
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from toolbox.xlsx_io import Workbook, NS

here = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("synthetic_baseline", here / "acceptance-workbook-fixtures.py")
baseline = importlib.util.module_from_spec(spec)
spec.loader.exec_module(baseline)
write_workbook, sheet_xml = baseline.write_workbook, baseline.sheet_xml
LEGACY_ID = "11111111-1111-4111-8111-111111111111"
HISTORY_ID = "synthetic-history-before-v017"
CAT1, CAT2 = "CI合成存货甲", "CI合成存货乙"
CUSTOMER_NAMES = {letter: f"CI合成客户公司{letter}" for letter in "ABC"}
ALIASES = {letter: f"CI合成门店{letter}" for letter in "ABC"}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def purchase_footer(start=40):
    return {
        start: {"E": "9月采购入库-某某供应商（月份取文件名称里的月份）", "F": "借", "G": "库存商品_库存商品", "H": "金额取含税", "I": "项目类别固定取存货分类（管理）"},
        start+1: {"G": "库存商品_冲减进项税", "H": "金额取（不含税-含税）", "I": "项目类别固定取存货分类（管理）"},
        start+2: {"G": "其他应付款_暂估进项税", "H": "金额取（含税-不含税）"},
        start+3: {"F": "贷", "G": "应付账款_应付账款-暂估", "H": "金额取含税", "I": "供应商根据供应商业务信息列表匹配得到"},
    }


def cost_footer(start=40):
    result = {}
    for block, text in enumerate(("销售出库-结转成本", "销售退回-冲成本", "期末成本调整出库", "期末成本调整入库")):
        for offset, subject in enumerate(("主营业务成本_供应链成本_销售成本", "主营业务成本_冲减进项税", "库存商品_库存商品", "库存商品_冲减进项税")):
            result[start+block*4+offset] = {"H": subject, "I": "金额取含税" if offset in (0, 2) else "金额取（不含税-含税）",
                "J": "项目类别固定取成本类\n项目固定取销售成本" if offset < 2 else "项目类别固定取存货分类（管理）\n项目根据存货类别的表匹配得到",
                **({"F": f"9月{text}"} if offset == 0 else {}), **({"G": "借" if offset == 0 else "贷"} if offset in (0, 2) else {})}
    result[start+16] = {"F": "月份取文件名称里的月份\n摘要：月份+出入库类型（销售出库后缀-结转成本，销售退回后缀-冲成本）"}
    return result


def revenue_footer(start=40):
    aux = "项目类别固定取收入类项目固定取销售收入\n客户根据客户业务信息列表匹配得到"
    return {
        start: {"F": "9月销售出库-确认收入", "G": "借", "H": "其他应收款_内部往来_成本", "I": "金额取应收", "J": "客户根据客户业务信息列表匹配得到"},
        start+1: {"G": "贷", "H": "主营业务收入_供应链收入_销售收入", "I": "金额取应收", "J": aux},
        start+2: {"H": "主营业务收入_冲减销项税_冲减销项税（内部）", "I": "金额取（收入-应收）", "J": aux},
        start+3: {"H": "应交税费_应交增值税_销项税", "I": "金额取销项，不同税率分开汇总金额", "J": "项目根据税率表匹配得到"},
        start+4: {"F": "月份取文件名称里的月份\n摘要：月份+销售出库-确认收入"},
    }


def source(path, rows):
    write_workbook(path, [("合成验收数据", sheet_xml(rows), "visible")])
    return str(path)


def formula_source(path, rows, formulas, merges=()):
    """Author normal, shared and uncached formulas as actual OOXML cells."""
    tree = ET.fromstring(sheet_xml(rows, merges))
    for address, (expression, cached, attributes) in formulas.items():
        cell = tree.find(f"{{{NS}}}sheetData/{{{NS}}}row/{{{NS}}}c[@r='{address}']")
        assert cell is not None, address
        cell.attrib.pop("t", None)
        for child in list(cell):
            cell.remove(child)
        element = ET.SubElement(cell, f"{{{NS}}}f", attributes)
        element.text = expression
        if cached is not None:
            ET.SubElement(cell, f"{{{NS}}}v").text = cached
    write_workbook(path, [("合成验收数据", ET.tostring(tree, encoding="utf-8", xml_declaration=True), "visible")])
    return str(path)


def generate(directory, data_directory):
    directory.mkdir(parents=True, exist_ok=True)
    data_directory.mkdir(parents=True, exist_ok=True)
    invoice = baseline.generate(directory)
    continuation_rows = {
        1: {"A": "数电发票号码", "B": "销售方纳税人名称", "C": "销售方纳税人识别号", "D": "金额*", "E": "有效抵扣税额*", "F": "美团税额", "G": "美团付款", "H": "美团供应商名称", "I": "税率", "J": "购买方识别号*", "K": "发票来源"},
        2: baseline.invoice("10000000000000000101", D="100", E="9", F="9", G="109", I="9%", J="SYNTHETIC-BUYER", K="电子发票数电"),
        # Screenshot-style continuations have blank A/B/C/H, their own D/E/I,
        # and may state a separate payment in G. All are the same invoice.
        3: {"D": "200", "E": "26", "F": "26", "G": "226", "I": "13%", "J": "SYNTHETIC-BUYER", "K": "电子发票数电"},
        4: {"D": "50", "E": "3", "F": "3", "G": "53", "I": "6%", "J": "SYNTHETIC-BUYER", "K": "电子发票数电"},
        5: baseline.invoice("10000000000000000102", D="40", E="5.20", F="5.20", G="45.20", I="13%", J="SYNTHETIC-BUYER", K="电子发票数电"),
        6: baseline.invoice("10000000000000000103", D="80", E="7.20", F="9.80", G="109.80", I="9%", J="SYNTHETIC-BUYER", K="电子发票数电"),
        # A whole payment can appear once in a merged G cell.
        7: {"D": "20", "E": "2.60", "I": "13%", "J": "SYNTHETIC-BUYER", "K": "电子发票数电"},
        # The last invoice uses an unmerged whole payment on the main row.
        8: baseline.invoice("10000000000000000104", D="300", E="27", F="53", G="553", I="9%", J="SYNTHETIC-BUYER", K="电子发票数电"),
        9: {"D": "200", "E": "26", "I": "13%", "J": "SYNTHETIC-BUYER", "K": "电子发票数电"},
        10: {"B": "汇总", "D": ("SUM(D2:D9)", "990.00"), "E": ("SUM(E2:E9)", "106.00")},
        # A multi-row personal register below the total is outside the invoice
        # section, even when rows resemble splits or entirely new invoices.
        11: {"D": "222", "E": "3", "G": "22", "I": "9%"},
        12: baseline.invoice("10000000000000000999", D="777", E="13", H="CI个人登记非供应商", I="13%"),
        13: {"E": "-0.25"},
    }
    continuation_source = directory/"2026年9月CI空票号分税率续行.xlsx"
    write_workbook(continuation_source, [("合成验收数据", sheet_xml(continuation_rows, ("F6:F7", "G6:G7", "H6:H7")), "visible")])
    rate_headers = continuation_rows[1]
    rate_rows = {
        1: rate_headers,
        # Only I is merged. Each populated A is still a separate invoice and
        # G remains its source payment, independently of D/E and the I merge.
        2: baseline.invoice("10000000000000000201", D="172.62", E="22.45", F="1629", G="14157.93", I="FORMULA", J="SYNTHETIC-BUYER"),
        3: baseline.invoice("10000000000000000202", D="17.62", E="2.30", F="2.30", G="19.92", I="", J="SYNTHETIC-BUYER"),
        4: baseline.invoice("10000000000000000203", D="172.62", E="22.43", F="22.43", G="195.05", I="", J="SYNTHETIC-BUYER"),
        5: baseline.invoice("10000000000000000204", D="100", E="9", F="9", G="109", I="9%", J="SYNTHETIC-BUYER"),
        # One split invoice uses a shared formula. The follower has <f .../>
        # and its own cached value, exactly as Excel writes shared formulas.
        6: baseline.invoice("10000000000000000205", D="199232.11", E="17930.89", F="17930.89", G="217163", I="FORMULA", J="SYNTHETIC-BUYER"),
        7: {"D":"9200.88", "E":"1196.12", "F":"1196.12", "G":"10397", "I":"FORMULA", "J":"SYNTHETIC-BUYER"},
        8: {"B":"汇总", "D":("SUM(D2:D7)","208895.85"), "E":("SUM(E2:E7)","19183.19")},
    }
    rate_source = formula_source(directory/"2026年9月CI合并税率及共享公式.xlsx", rate_rows, {
        "I2": ("F2/(G2-F2)", "0.130019083832378", {}),
        "I6": ("E6/D6", "0.0900000005019271", {"t":"shared", "si":"0", "ref":"I6:I7"}),
        "I7": (None, "0.130000608637435", {"t":"shared", "si":"0"}),
    }, ("I2:I4",))
    unsupported_rate = formula_source(directory/"2026年9月CI实际混合税率10.xlsx", {
        1:rate_headers, 2:baseline.invoice("10000000000000000301", D="100", E="10", F="10", G="110", I="FORMULA", J="SYNTHETIC-BUYER"),
    }, {"I2":("E2/D2", "0.1", {})})
    missing_rate_cache = formula_source(directory/"2026年9月CI合并税率无缓存.xlsx", {
        1:rate_headers,
        2:baseline.invoice("10000000000000000302", D="100", E="13", F="13", G="113", I="FORMULA", J="SYNTHETIC-BUYER"),
        3:baseline.invoice("10000000000000000303", D="17.62", E="2.30", F="2.30", G="19.92", I="", J="SYNTHETIC-BUYER"),
    }, {"I2":("F2/(G2-F2)", None, {})}, ("I2:I3",))
    conflicting_rate_cache = formula_source(directory/"2026年9月CI税率缓存与金额冲突.xlsx", {
        1:rate_headers, 2:baseline.invoice("10000000000000000304", D="100", E="9", F="9", G="109", I="FORMULA", J="SYNTHETIC-BUYER"),
    }, {"I2":("E2/D2", "0.13", {})})
    orphan_payment_source = source(directory/"2026年9月CI付款无所属发票.xlsx", {
        1:rate_headers,
        2:baseline.invoice("10000000000000000305", D="100", E="9", F="9", G="109", I="9%", J="SYNTHETIC-BUYER"),
        3:{"G":"50.25", "H":baseline.SUPPLIER_NAME},
        4:{"B":"汇总", "D":"100", "E":"9"},
        5:{"G":"999", "H":"CI个人登记非供应商"},
    })
    purchase = formula_source(directory/"2026年9月CI采购.xlsx", {
        1: dict(zip("ABCDEF", ("单据类型", "对方机构", "物品类别", "入库成本金额(含税)", "入库成本金额（不含税）", "税额"))),
        2: dict(zip("ABCDEF", ("采购入库", baseline.SUPPLIER_NAME, CAT1, "113", "100", "13"))),
        3: {"C": CAT2, "D": "55", "E": "50", "F": "5"},
        4: {"B": baseline.SUPPLIER_NAME, "C": CAT1, "D": "-10", "E": "-9", "F": "-1"},
        5: dict(zip("ABCDEF", ("成本调整入库", baseline.SUPPLIER_NAME, CAT1, "0", "-2", "2"))),
        6: {"B": "CI待补充的缺码供应商", "C": CAT1, "D": "11", "E": "10", "F": "1"},
        # The first summary row and all later content are intentionally invalid.
        # Neither these totals nor the former voucher footer may be validated.
        30: {"D": "-999999.99", "E": "CI错误汇总不是金额", "F": "FORMULA"},
        40: {"A":"CI未知采购类型", "B":"CI不存在供应商", "C":"CI不存在类别", "D":"非法金额", "E":"FORMULA", "F":"错误税额", "G":"无凭证规则"},
        41: {"H":"旧要求已移除", "J":"CI私人登记"},
    }, {"F30":("SUM(F2:F6)", None, {}), "E40":("1/0", None, {})})
    cost = formula_source(directory/"2026年9月CI销售结转成本.xlsx", {
        1: dict(zip("ABCDE", ("出入库类型", "物品类别", "含税", "不含税", "税额"))),
        2: dict(zip("ABCDE", ("销售出库", CAT1, "100", "90", "10"))),
        3: {"B": CAT2, "C": "200", "D": "180", "E": "20"},
        4: {"B": CAT1, "C": "50", "D": "45", "E": "5"},
        5: dict(zip("ABCDE", ("销售退回", CAT1, "-30", "-27", "-3"))),
        6: {"A": "期末成本调整出库", "B": CAT2, "D": "-5", "E": "5"},
        7: dict(zip("ABCDE", ("期末成本调整入库", CAT1, "-20", "-18", "-2"))),
        30: {"C": "-888888.88", "D": "CI错误汇总不是不含税", "E": "FORMULA"},
        40: {"A":"CI未知出库类型", "B":"CI不存在存货", "C":"非法含税", "D":"FORMULA", "E":"错误税额", "J":"底部不提供凭证规则"},
        41: {"F":"仅个人登记", "K":"CI不支持的额外列"},
    }, {"E30":("SUM(E2:E7)", None, {}), "D40":("1/0", None, {})})
    frac_net = Decimal("0.036697247706422018349")
    frac_tax = Decimal("0.003302752293577981651")
    revenue = formula_source(directory/"2026年9月CI销售确认收入.xlsx", {
        1: dict(zip("ABCDE", ("对方机构", "税率", "求和项:应收", "求和项:收入", "求和项:销项"))),
        2: dict(zip("ABCDE", (ALIASES["A"], "9%", "109", "100", "9"))),
        3: {"B": "13%", "C": "113", "D": "100", "E": "13"},
        4: dict(zip("ABCDE", (ALIASES["B"], "9%", "-54.50", "-50", "-4.50"))),
        5: dict(zip("ABCDE", (ALIASES["C"], "9%", "0.04", str(frac_net), str(frac_tax)))),
        6: {"B": "9%", "C": "0.04", "D": str(frac_net), "E": str(frac_tax)},
        30: {"A": "(空白)", "B": "(空白)", "C": "999999.99", "D": "CI错误汇总不是收入", "E": "FORMULA"},
        40: {"A":"CI未知客户", "B":"10%", "C":"非法应收", "D":"FORMULA", "E":"NaN", "F":"缺少所有原凭证规则"},
        41: {"J":"CI私人登记不参与校验"},
    }, {"E30":("SUM(E2:E6)", None, {}), "D40":("1/0", None, {})})
    customer_import = source(directory/"CI客户名称映射.xlsx", {
        1: {"A": "合成客户映射"}, 2: {"A": "原始数据", "B": "客户编码", "C": "客户名称"},
        **{i+3: {"A": ALIASES[letter], "B": f"CI-CUSTOMER-{letter}", "C": CUSTOMER_NAMES[letter]} for i, letter in enumerate("ABC")}})
    paged_rows = {1: dict(zip("ABCDEF", ("单据类型", "对方机构", "物品类别", "入库成本金额(含税)", "入库成本金额（不含税）", "税额")))}
    for index in range(220):
        paged_rows[index + 2] = dict(zip("ABCDEF", ("采购入库", f"CI分页待补供应商{index:03}", CAT1, "113", "100", "13")))
    paged_rows[230] = {"D": "24860", "E": "22000", "F": "2860"}
    paged_rows.update(purchase_footer(240))
    paged_purchase = source(directory/"2026年9月CI分页采购.xlsx", paged_rows)
    stamp = "2026-01-01T00:00:00.000Z"
    db = sqlite3.connect(data_directory/"master-data.sqlite3")
    db.execute("CREATE TABLE master_kinds(kind TEXT PRIMARY KEY,revision INTEGER NOT NULL DEFAULT 0,imported_at TEXT,source_name TEXT NOT NULL DEFAULT '')")
    db.execute("CREATE TABLE master_records(id TEXT PRIMARY KEY,kind TEXT NOT NULL REFERENCES master_kinds(kind),code TEXT NOT NULL,name TEXT NOT NULL,category TEXT NOT NULL DEFAULT '',updated_at TEXT NOT NULL,tax_rate TEXT NOT NULL DEFAULT '',tax_direction TEXT NOT NULL DEFAULT '',UNIQUE(kind,code))")
    for kind in ("suppliers", "inventoryCategories", "customers", "taxRates"):
        db.execute("INSERT INTO master_kinds(kind,revision) VALUES (?,7)", (kind,))
    old_records = [(str(uuid4()), "suppliers", "CI-SUP-0001", baseline.SUPPLIER_NAME, "SYNTHETIC_ACCEPTANCE_ONLY", stamp, "", ""),
                   (str(uuid4()), "inventoryCategories", "CI-INV-A", CAT1, "存货分类（管理）", stamp, "", ""),
                   (str(uuid4()), "inventoryCategories", "CI-INV-B", CAT2, "存货分类（管理）", stamp, "", ""),
                   (LEGACY_ID, "customers", "CI-CUSTOMER-A", CUSTOMER_NAMES["A"], "SYNTHETIC", stamp, "", "")]
    old_records += [(str(uuid4()), "taxRates", f"CC1700{i}", f"进项税{percent}%", "", stamp, f"{percent}%", "input") for i,percent in enumerate((1,3,6,9,13),1)]
    db.executemany("INSERT INTO master_records VALUES (?,?,?,?,?,?,?,?)", old_records)
    db.commit(); db.close()
    history = {"id": HISTORY_ID, "createdAt": stamp, "status": "completed", "toolId": "file-manifest", "toolVersion": "1.0.0", "toolName": "合成迁移前历史", "inputPaths": [], "outputDirectory": "", "parameters": {}, "outputs": [], "summary": {"synthetic": True}, "progress": 100, "startedAt": stamp, "finishedAt": stamp, "message": "synthetic preserved history", "error": None}
    jobs = sqlite3.connect(data_directory/"jobs.sqlite3")
    jobs.execute("CREATE TABLE jobs(id TEXT PRIMARY KEY,created_at TEXT NOT NULL,status TEXT NOT NULL,record TEXT NOT NULL,work_directory TEXT)")
    jobs.execute("INSERT INTO jobs VALUES (?,?,?,?,NULL)", (HISTORY_ID, stamp, "completed", json.dumps(history, ensure_ascii=False)))
    jobs.commit(); jobs.close()
    return {"sources": {"invoice": invoice["source"], "purchase": purchase, "sales-cost": cost, "sales-revenue": revenue},
            "invoice": invoice, "suppliers": invoice["suppliers"], "customerImport": customer_import,
            "invoiceContinuations": str(continuation_source),
            "invoiceFormulaRates": rate_source,
            "invalidInvoiceRates": {"actual-mixed-10-percent":unsupported_rate, "merged-formula-without-cache":missing_rate_cache, "formula-cache-conflicts-with-amounts":conflicting_rate_cache},
            "invoiceOrphanPayment": orphan_payment_source,
            "pagedPurchase": paged_purchase,
            "missingSupplier": {"code": "CI-SUP-FILLED", "name": "CI待补充的缺码供应商"},
            "legacyCustomer": {"id": LEGACY_ID, "code": "CI-CUSTOMER-A", "name": CUSTOMER_NAMES["A"], "updatedAt": stamp},
            "history": history, "outputRates": [{"code": f"CI-OUTPUT-{rate}", "name": f"CI销项{rate}%", "taxRate": f"{rate}%", "taxDirection": "output"} for rate in (9,13)]}


def audit_install(engine, expected):
    engine = engine.resolve()
    manifest = json.loads((engine/"runtime-manifest.json").read_text(encoding="utf-8"))
    for field in ("sourceFiles", "templateFiles"):
        for item in manifest[field]:
            assert digest(engine/item["path"]) == item["sha256"], item["path"]
    assert (engine/"python314._pth").read_text() == "python314.zip\n.\napp\n"
    package = engine.parent/"app.asar"
    with package.open("rb") as stream:
        size, header_size, payload, json_size = struct.unpack("<4I", stream.read(16))
        assert size == 4 and payload+4 == header_size
        tree = json.loads(stream.read(json_size))
        entry = tree["files"]["package.json"]
        stream.seek(8+header_size+int(entry["offset"]))
        version = json.loads(stream.read(entry["size"]))["version"]
    assert version == expected, (version,expected)
    loose = [str(p.relative_to(engine)) for p in engine.rglob("*") if p.is_file() and p.suffix.lower() in {".db", ".sqlite", ".sqlite3", ".csv", ".xls"}]
    assert not loose, loose
    actual_templates = {p.relative_to(engine/"app/toolbox/templates").as_posix() for p in (engine/"app/toolbox/templates").rglob("*.xlsx")}
    assert actual_templates == {f"voucher-import/{name}.xlsx" for name in ("invoice","purchase","sales-cost","sales-revenue")}
    assert {p.relative_to(engine).as_posix() for p in engine.rglob("*.xlsx")} == {f"app/toolbox/templates/{name}" for name in actual_templates}
    return {"version": version, "allRuntimeManifestHashesMatch": True, "fixedTemplateCount": len(actual_templates), "noBundledDatabaseOrBusinessCsv": True}


def verify(category, template_path, output, audit_path, supplier_code):
    before, after = Workbook(template_path), Workbook(output)
    original, sheet = before.sheet("凭证查询卡片主表"), after.sheet("凭证查询卡片主表")
    invoice_category = category in {"invoice", "invoice-continuations", "invoice-formula-rates"}
    start = 9 if invoice_category else 10
    rows = [r for r in sheet.rows if r >= start]
    expected_count = {"invoice":109,"invoice-continuations":20,"invoice-formula-rates":21,"purchase":16,"sales-cost":18,"sales-revenue":11}[category]
    assert len(rows) == expected_count, (category,len(rows),expected_count)
    groups = defaultdict(list)
    for index,row in enumerate(rows,1):
        assert sheet.value(row,"R") == str(index)
        assert sheet.value(row,"D") == sheet.value(row,"E") == "CYFW"
        assert sheet.value(row,"G") == "2026-09" and sheet.value(row,"H") == "2026-09-30"
        assert all(sheet.cell(row,col).kind == "inlineStr" for col in range(1,73))
        assert bool(sheet.value(row,"X")) != bool(sheet.value(row,"Y"))
        assert sheet.value(row,"X") == sheet.value(row,"AC") and sheet.value(row,"Y") == sheet.value(row,"AD")
        groups[sheet.value(row,"C")].append(row)
        if sheet.value(row,"BC"):
            assert sheet.value(row,"BC") in ({supplier_code, "CI-SUP-FILLED"} if category == "purchase" else {supplier_code})
        if not invoice_category:
            assert not any(sheet.value(row,col) for col in ("L","M","AJ"))
    for group in groups.values():
        assert sum((Decimal(sheet.value(r,"X") or "0") for r in group),Decimal(0)) == sum((Decimal(sheet.value(r,"Y") or "0") for r in group),Decimal(0))
        assert [sheet.value(r,"S") for r in group] == [str(i) for i in range(1,len(group)+1)]
    assert len(groups) == {"invoice":26,"invoice-continuations":4,"invoice-formula-rates":5,"purchase":4,"sales-cost":4,"sales-revenue":1}[category]
    with ZipFile(template_path) as old, ZipFile(output) as new:
        assert old.namelist() == new.namelist() and old.comment == new.comment
        for part in old.namelist():
            if part != original.part:
                assert old.read(part) == new.read(part), part
        old_tree, new_tree = ET.fromstring(original.raw), ET.fromstring(sheet.raw)
        for number in range(1,start):
            selector = f"{{{NS}}}sheetData/{{{NS}}}row[@r='{number}']"
            a,b = old_tree.find(selector), new_tree.find(selector)
            assert (ET.tostring(a) if a is not None else None) == (ET.tostring(b) if b is not None else None)
        assert ET.tostring(old_tree.find(f"{{{NS}}}cols")) == ET.tostring(new_tree.find(f"{{{NS}}}cols"))
    with open(audit_path,encoding="utf-8-sig",newline="") as stream:
        audit = list(csv.DictReader(stream))
    details = {}
    if category in {"invoice-continuations", "invoice-formula-rates"}:
        assert len(audit) == (8 if category == "invoice-continuations" else 6)
        assert {r["供应商编码"] for r in audit} == {supplier_code}
        by_invoice = defaultdict(list)
        for item in audit:
            by_invoice[item["发票号码"]].append(item)
        expected = {
            "10000000000000000101": ("350.00", "38.00", "388.00", {"6%":"3.00", "9%":"9.00", "13%":"26.00"}),
            "10000000000000000102": ("40.00", "5.20", "45.20", {"13%":"5.20"}),
            "10000000000000000103": ("100.00", "9.80", "109.80", {"9%":"7.20", "13%":"2.60"}),
            "10000000000000000104": ("500.00", "53.00", "553.00", {"9%":"27.00", "13%":"26.00"}),
        }
        if category == "invoice-formula-rates":
            expected = {
                "10000000000000000201": ("172.62", "22.45", "14157.93", {"13%":"22.45"}),
                "10000000000000000202": ("17.62", "2.30", "19.92", {"13%":"2.30"}),
                "10000000000000000203": ("172.62", "22.43", "195.05", {"13%":"22.43"}),
                "10000000000000000204": ("100.00", "9.00", "109.00", {"9%":"9.00"}),
                "10000000000000000205": ("208432.99", "19127.01", "227560.00", {"9%":"17930.89", "13%":"1196.12"}),
            }
        assert set(by_invoice) == set(expected)
        for identifier, (net, tax, payment, tax_by_rate) in expected.items():
            data = by_invoice[identifier]
            assert {item["不含税金额"] for item in data} == {net}, (identifier, data)
            assert {item["有效抵扣税额"] for item in data} == {tax}, (identifier, data)
            assert {item["付款组金额（组内勿重复求和）"] for item in data} == {payment}, (identifier, data)
            actual_rates = {Decimal(item["税率"].removesuffix("%")):Decimal(item["该税率税额"]) for item in data}
            expected_rates = {Decimal(key.removesuffix("%")):Decimal(value) for key,value in tax_by_rate.items()}
            assert actual_rates == expected_rates, (identifier, actual_rates, expected_rates)
            assert len({item["凭证系统码"] for item in data}) == 1
        payment_rows = [r for r in rows if sheet.value(r,"T") == "220202"]
        expected_payments = ["388.00", "45.20", "109.80", "553.00"] if category == "invoice-continuations" else ["14157.93", "19.92", "195.05", "109.00", "227560.00"]
        assert [sheet.value(r,"X") for r in payment_rows] == expected_payments
        assert all(sheet.value(r,"BC") == supplier_code for r in payment_rows)
        tax_projects = defaultdict(Decimal)
        for r in rows:
            if sheet.value(r,"T") == "22210101": tax_projects[sheet.value(r,"AY")] += Decimal(sheet.value(r,"X"))
        expected_tax_projects = {"CC17003":Decimal("3.00"), "CC17004":Decimal("43.20"), "CC17005":Decimal("59.80")} if category == "invoice-continuations" else {"CC17004":Decimal("17939.89"), "CC17005":Decimal("1243.30")}
        assert dict(tax_projects) == expected_tax_projects
        total = sum((Decimal(sheet.value(r,"X") or "0") for r in rows), Decimal(0))
        expected_total = "1202.00" if category == "invoice-continuations" else "261225.09"
        assert total == Decimal(expected_total), total
        details = ({"invoiceCount":4,"sourceDetailRows":8,"taxRowCount":8,"netTotal":"990.00","taxTotal":"106.00","paymentTotal":"1096.00","debitTotal":"1202.00","creditTotal":"1202.00","consecutiveBlankIdentitySplitsPreserved":True,"lastInvoiceSplitPreserved":True,"separateMergedAndUnmergedWholePaymentsCountedOnce":True,"multiRowRegisterBelowTotalExcluded":True,"nextInvoiceStartsNewVoucher":True}
            if category == "invoice-continuations" else {"invoiceCount":5,"sourceDetailRows":6,"taxRowCount":6,"netTotal":"208895.85","taxTotal":"19183.19","paymentTotal":"242041.90","debitTotal":"261225.09","creditTotal":"261225.09","verticalMergedRatePreservedAcrossThreeInvoices":True,"plusAndMinusOneCentTaxesPreserved":True,"smallInvoiceMergedRateMatched":True,"nearStandardFormulaCachesValidatedAgainstAmounts":True,"sharedFormulaFollowerOwnCacheUsed":True,"rateMergeDoesNotMergeVouchersOrPayments":True,"sourcePaymentsUnchanged":True})
    elif category == "invoice":
        assert len(audit) == 59
        assert {r["供应商编码"] for r in audit} == {supplier_code}
        tax = defaultdict(Decimal)
        for r in rows:
            if sheet.value(r,"T") == "22210101": tax[sheet.value(r,"AY")] += Decimal(sheet.value(r,"X"))
        assert dict(tax) == {"CC17001":Decimal("1"),"CC17004":Decimal("207"),"CC17005":Decimal("403")}
    elif category == "purchase":
        assert {int(r["来源行"]) for r in audit} == {2,3,4,5,6}
        assert {r["供应商编码"] for r in audit} == {supplier_code, "CI-SUP-FILLED"}
        assert {r["T"] for r in ({"T":sheet.value(n,"T")} for n in rows)} == {"140501","140502","224101","220202"}
    elif category == "sales-cost":
        assert {int(r["来源行"]) for r in audit} == set(range(2,8))
        duplicates = [r for r in rows if sheet.value(r,"C")=="1" and sheet.value(r,"T")=="140501" and sheet.value(r,"AY")=="CI-INV-A"]
        assert [sheet.value(r,"Y") for r in duplicates] == ["100.00","50.00"]
        assert all(sheet.value(r,"AY")=="CC13002" for r in rows if sheet.value(r,"T").startswith("640"))
    else:
        assert {int(r["来源行"]) for r in audit} == set(range(2,7))
        assert any(r["原始销项（未按行舍入）"] == "0.003302752293577981651" for r in audit)
        customers = {sheet.value(r,"BE") for r in rows if sheet.value(r,"BE")}
        assert customers == {f"CI-CUSTOMER-{letter}" for letter in "ABC"}
        adjustments = {sheet.value(r,"BE"):sheet.value(r,"Y") for r in rows if sheet.value(r,"T")=="60010501"}
        assert adjustments["CI-CUSTOMER-C"] == "-0.01", adjustments
        tax = {sheet.value(r,"AY"):sheet.value(r,"Y") for r in rows if sheet.value(r,"T")=="22210102"}
        assert tax == {"CI-OUTPUT-9":"4.51","CI-OUTPUT-13":"13.00"}
        assert all(not sheet.value(r,"BE") and not sheet.value(r,"AW") for r in rows if sheet.value(r,"T")=="22210102")
    if not invoice_category:
        details["invalidSummaryTotalsAndUncachedFormulasIgnored"] = True
        details["invalidUnknownTypesAndMissingVoucherRulesBelowSummaryIgnored"] = True
    return {"verified": True,"category":category,"voucherCount":len(groups),"entryCount":len(rows),"auditRowCount":len(audit),"balanced":True,"all72ColumnsText":True,"allOtherZipMembersPreserved":True,"headersAndHiddenColumnsPreserved":True,**details}


if __name__ == "__main__":
    action = sys.argv[1]
    if action == "generate": result=generate(Path(sys.argv[2]).resolve(),Path(sys.argv[3]).resolve())
    elif action == "install": result=audit_install(Path(sys.argv[2]),sys.argv[3])
    elif action == "verify": result=verify(*sys.argv[2:])
    else: raise ValueError("Unknown synthetic acceptance action")
    print(json.dumps(result,ensure_ascii=False))
