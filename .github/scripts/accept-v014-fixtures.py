"""Disposable, synthetic-only v0.1.14 Windows installed-engine acceptance data.

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
HISTORY_ID = "synthetic-history-before-v014"
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


def generate(directory, data_directory):
    directory.mkdir(parents=True, exist_ok=True)
    data_directory.mkdir(parents=True, exist_ok=True)
    invoice = baseline.generate(directory)
    purchase = source(directory/"2026年9月CI采购.xlsx", {
        1: dict(zip("ABCDEF", ("单据类型", "对方机构", "物品类别", "入库成本金额(含税)", "入库成本金额（不含税）", "税额"))),
        2: dict(zip("ABCDEF", ("采购入库", baseline.SUPPLIER_NAME, CAT1, "113", "100", "13"))),
        3: {"C": CAT2, "D": "55", "E": "50", "F": "5"},
        4: {"B": baseline.SUPPLIER_NAME, "C": CAT1, "D": "-10", "E": "-9", "F": "-1"},
        5: dict(zip("ABCDEF", ("成本调整入库", baseline.SUPPLIER_NAME, CAT1, "0", "-2", "2"))),
        6: {"B": "CI待补充的缺码供应商", "C": CAT1, "D": "11", "E": "10", "F": "1"},
        30: {"D": "169", "E": "149", "F": "20"}, **purchase_footer()})
    cost = source(directory/"2026年9月CI销售结转成本.xlsx", {
        1: dict(zip("ABCDE", ("出入库类型", "物品类别", "含税", "不含税", "税额"))),
        2: dict(zip("ABCDE", ("销售出库", CAT1, "100", "90", "10"))),
        3: {"B": CAT2, "C": "200", "D": "180", "E": "20"},
        4: {"B": CAT1, "C": "50", "D": "45", "E": "5"},
        5: dict(zip("ABCDE", ("销售退回", CAT1, "-30", "-27", "-3"))),
        6: {"A": "期末成本调整出库", "B": CAT2, "D": "-5", "E": "5"},
        7: dict(zip("ABCDE", ("期末成本调整入库", CAT1, "-20", "-18", "-2"))),
        30: {"C": "300", "D": "265", "E": "35"}, **cost_footer()})
    frac_net = Decimal("0.036697247706422018349")
    frac_tax = Decimal("0.003302752293577981651")
    revenue = source(directory/"2026年9月CI销售确认收入.xlsx", {
        1: dict(zip("ABCDE", ("对方机构", "税率", "求和项:应收", "求和项:收入", "求和项:销项"))),
        2: dict(zip("ABCDE", (ALIASES["A"], "9%", "109", "100", "9"))),
        3: {"B": "13%", "C": "113", "D": "100", "E": "13"},
        4: dict(zip("ABCDE", (ALIASES["B"], "9%", "-54.50", "-50", "-4.50"))),
        5: dict(zip("ABCDE", (ALIASES["C"], "9%", "0.04", str(frac_net), str(frac_tax)))),
        6: {"B": "9%", "C": "0.04", "D": str(frac_net), "E": str(frac_tax)},
        30: {"A": "(空白)", "B": "(空白)", "C": "167.58", "D": str(Decimal("150")+2*frac_net), "E": str(Decimal("17.5")+2*frac_tax)}, **revenue_footer()})
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
    start = 9 if category == "invoice" else 10
    rows = [r for r in sheet.rows if r >= start]
    expected_count = {"invoice":109,"purchase":16,"sales-cost":18,"sales-revenue":11}[category]
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
        if category != "invoice":
            assert not any(sheet.value(row,col) for col in ("L","M","AJ"))
    for group in groups.values():
        assert sum((Decimal(sheet.value(r,"X") or "0") for r in group),Decimal(0)) == sum((Decimal(sheet.value(r,"Y") or "0") for r in group),Decimal(0))
        assert [sheet.value(r,"S") for r in group] == [str(i) for i in range(1,len(group)+1)]
    assert len(groups) == {"invoice":26,"purchase":4,"sales-cost":4,"sales-revenue":1}[category]
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
    if category == "invoice":
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
    return {"verified": True,"category":category,"voucherCount":len(groups),"entryCount":len(rows),"auditRowCount":len(audit),"balanced":True,"all72ColumnsText":True,"allOtherZipMembersPreserved":True,"headersAndHiddenColumnsPreserved":True}


if __name__ == "__main__":
    action = sys.argv[1]
    if action == "generate": result=generate(Path(sys.argv[2]).resolve(),Path(sys.argv[3]).resolve())
    elif action == "install": result=audit_install(Path(sys.argv[2]),sys.argv[3])
    elif action == "verify": result=verify(*sys.argv[2:])
    else: raise ValueError("Unknown synthetic acceptance action")
    print(json.dumps(result,ensure_ascii=False))
