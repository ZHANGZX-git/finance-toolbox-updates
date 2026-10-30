"""Synthetic-only installed-engine acceptance fixtures; no third-party imports.

Run with the installed resources/engine/python.exe. Reading and verification use
that release's toolbox.xlsx_io; workbook authoring follows the stdlib ZIP/XML
fixture pattern in this project's tests. No production source is copied here.
"""
from collections import defaultdict
from decimal import Decimal
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile

from toolbox.tools.voucher_import import TEMPLATE_FIELDS
from toolbox.xlsx_io import NS, Workbook, column_name

SUPPLIER_NAME = "合成验收供应商0001（仅测试）"
TEMPLATE_CODE = "T-CI-0001"


def sheet_xml(rows, merges=(), *, template=False):
    content = []
    for number, values in sorted(rows.items()):
        cells = []
        for column, value in values.items():
            index = 0
            for char in column:
                index = index * 26 + ord(char) - 64
            style = f' s="{index % 3}"' if template and number == 9 else ""
            if isinstance(value, tuple):
                formula, cached = value
                cell = f"<f>{escape(formula)}</f><v>{escape(str(cached))}</v>"
                cells.append(f'<c r="{column}{number}"{style}>{cell}</c>')
            else:
                cells.append(f'<c r="{column}{number}"{style} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
        content.append(f'<row r="{number}"' + (' ht="18" customHeight="1"' if template and number == 9 else "") + ">" + "".join(cells) + "</row>")
    return (f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{NS}"><dimension ref="A1:BT{max(rows, default=1)}"/>'
            '<cols><col min="1" max="2" hidden="1" width="20"/><col min="57" max="72" hidden="1"/></cols><sheetData>'
            + "".join(content) + "</sheetData>"
            + (f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{value}"/>' for value in merges) + "</mergeCells>" if merges else "")
            + "</worksheet>").encode("utf-8")


def write_workbook(destination, sheets):
    workbook = f'<workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    relations = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    with ZipFile(destination, "w") as archive:
        for index, (name, data, state) in enumerate(sheets, 1):
            workbook += f'<sheet name={quoteattr(name)} sheetId="{index}" state={quoteattr(state)} r:id="rId{index}"/>'
            relations += f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
            types += f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            archive.writestr(f"xl/worksheets/sheet{index}.xml", data)
        archive.writestr("xl/workbook.xml", workbook + "</sheets></workbook>")
        archive.writestr("xl/_rels/workbook.xml.rels", relations + "</Relationships>")
        archive.writestr("[Content_Types].xml", types + "</Types>")
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/styles.xml", f'<styleSheet xmlns="{NS}"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf numFmtId="49"/><xf numFmtId="49"/><xf numFmtId="49"/></cellXfs></styleSheet>')
        archive.writestr("custom/synthetic-preserve.bin", b"SYNTHETIC PRESERVE ALL OTHER ZIP MEMBERS\x00")
        archive.comment = b"synthetic-windows-worker-acceptance"


def invoice(identifier, **changes):
    record = {"A": str(identifier), "B": "合成测试销售方", "C": "SYNTHETIC-TAX", "D": "100.00", "E": "9.00", "F": "9.00", "G": "109.00", "H": SUPPLIER_NAME, "I": "9%", "L": f"9月收{SUPPLIER_NAME}发票-{identifier}"}
    record.update(changes)
    return record


def generate(directory):
    directory.mkdir(parents=True, exist_ok=True)
    source = directory / "2026年9月合成采购发票.xlsx"
    template = directory / "合成凭证模板.xlsx"
    headers = {"A": "数电发票号码", "B": "销售方纳税人名称", "C": "销售方纳税人识别号", "D": "金额*", "E": "有效抵扣税额*", "F": "美团税额", "G": "美团付款", "H": "美团供应商名称", "I": "税率", "J": "税额"}
    rows, row = {1: headers}, 2
    for index in range(55):
        identifier = 10000000000000000001 + index
        if index < 4:
            rows[row] = invoice(identifier, E="4.50", I="9%", J="4.50", G="111.00")
            rows[row + 1] = {"E": "6.50", "F": "6.50", "I": "13%", "J": "6.50"}
            row += 2
        elif index < 34:
            first = index == 4
            rows[row] = invoice(identifier, E="1.00" if first else "13.00", I="1%" if first else "13%", F="377.99" if first else "", G="3377.75" if first else "", H=SUPPLIER_NAME if first else "")
            row += 1
        else:
            rows[row] = invoice(identifier)
            row += 1
    rows[row] = {"D": ("SUM(D2:D60)", "5500.00"), "E": ("SUM(E2:E60)", "611.00")}
    rows[row + 1] = {"E": "-0.25"}
    write_workbook(source, [("Sheet1", sheet_xml(rows, ("F10:F39", "G10:G39", "H10:H39")), "visible")])
    sample = {column_name(index): "" for index in range(1, 73)}
    sample.update({"D": "SYNTHETIC-ORG", "E": "SYNTHETIC-BOOK", "F": "合成验收账簿", "G": "2026-08", "H": "2026-08-31", "I": "记", "L": "synthetic-maker", "W": "CNY", "T": "22210101", "V": "旧样例必须移除", "AY": "CC17004", "AZ": "进项税9%", "X": "9", "AC": "9"})
    template_rows = {1: {"C": "主表"}, 2: {"A": "实体编码"}, 3: {"A": "父实体编码"}, 4: {column_name(index): value for index, value in enumerate(TEMPLATE_FIELDS, 1)}, 5: {"C": "PrimaryKey"}, 6: {"C": "系统码"}, 7: {"A": "√"}, 8: {"A": "√"}, 9: sample, 10: {"T": "OLD-SAMPLE", "V": "旧样例分录必须移除"}}
    write_workbook(template, [
        ("凭证查询卡片主表", sheet_xml(template_rows, template=True), "visible"),
        ("参照", sheet_xml({1: {"D": "供应商编码", "E": "供应商名称"}, 2: {"D": TEMPLATE_CODE, "E": SUPPLIER_NAME}}), "hidden"),
        ("importTemplateInfo", sheet_xml({1: {"A": '{"synthetic": "preserve"}'}}), "veryHidden"),
    ])
    # Prove the same installed release can read both synthetic workbooks.
    assert len(Workbook(source).sheet("Sheet1").merges) == 3
    assert Workbook(template).sheet("参照").state == "hidden"
    suppliers = [{"code": f"CI-SUP-{index:04}", "name": SUPPLIER_NAME if index == 1 else f"合成验收供应商{index:04}（仅测试）", "category": "SYNTHETIC_ACCEPTANCE_ONLY"} for index in range(1, 61)]
    return {"source": str(source), "template": str(template), "suppliers": suppliers, "expected": {"发票数": 55, "凭证数": 26, "分录数": 109, "核对清单行数": 59, "会计期间": "2026-09", "制单日期": "2026-09-30"}, "parameters": {"category": "invoice", "taxSplitMode": "source", "mergedPaymentMode": "merged", "splitPaymentMode": "source", "includeTaxDifference": "no"}}


def verify(directory, output):
    template = Workbook(directory / "合成凭证模板.xlsx")
    original = template.sheet("凭证查询卡片主表")
    book = Workbook(output)
    sheet = book.sheet("凭证查询卡片主表")
    rows = [row for row in sheet.rows if row >= 9]
    assert len(rows) == 109, len(rows)
    assert [sheet.value(row, "R") for row in rows] == [str(index) for index in range(1, 110)]
    vouchers, rates = defaultdict(list), defaultdict(Decimal)
    for row in rows:
        vouchers[sheet.value(row, "C")].append(row)
        assert sheet.value(row, "D") == sheet.value(row, "E") == "CYFW"
        assert sheet.value(row, "G") == "2026-09"
        assert sheet.value(row, "H") == "2026-09-30"
        assert sheet.value(row, "J") == ""
        assert sheet.value(row, "T") != "OLD-SAMPLE"
        assert all(sheet.cell(row, column).kind == "inlineStr" for column in range(1, 73))
        if sheet.value(row, "T") == "22210101":
            rates[sheet.value(row, "AY")] += Decimal(sheet.value(row, "X"))
        if sheet.value(row, "BC"):
            assert sheet.value(row, "BC") == TEMPLATE_CODE, "Template supplier must take precedence over the conflicting saved code"
    assert list(vouchers) == [str(index) for index in range(1, 27)]
    for voucher_rows in vouchers.values():
        assert [sheet.value(row, "S") for row in voucher_rows] == [str(index) for index in range(1, len(voucher_rows) + 1)]
        debit = sum((Decimal(sheet.value(row, "X") or "0") for row in voucher_rows), Decimal("0"))
        credit = sum((Decimal(sheet.value(row, "Y") or "0") for row in voucher_rows), Decimal("0"))
        assert debit == credit, (voucher_rows, debit, credit)
    assert dict(rates) == {"CC17001": Decimal("1.00"), "CC17004": Decimal("207.00"), "CC17005": Decimal("403.00")}, rates
    merged = vouchers["5"]
    assert len(merged) == 5
    assert [sheet.value(row, "X") or sheet.value(row, "Y") for row in merged] == ["1.00", "377.00", "378.00", "3377.75", "3377.75"]
    assert "共30张" in sheet.value(merged[0], "V")
    total = sum((Decimal(sheet.value(row, "X") or "0") for row in rows), Decimal("0"))
    assert total == Decimal("6721.75"), total
    with ZipFile(template.path) as before, ZipFile(output) as after:
        assert set(before.namelist()) == set(after.namelist())
        for name in before.namelist():
            if name != original.part:
                assert before.read(name) == after.read(name), name
        prefix = lambda raw: raw.split(b"<sheetData>")[1].split(b'<row r="9"')[0]
        assert prefix(before.read(original.part)) == prefix(after.read(original.part))
        cols = lambda raw: ET.tostring(ET.fromstring(raw).find(f"{{{NS}}}cols"))
        assert cols(before.read(original.part)) == cols(after.read(original.part))
        assert before.comment == after.comment
    assert book.sheet("参照").state == "hidden"
    assert book.sheet("importTemplateInfo").state == "veryHidden"
    return {"verified": True, "rows": 109, "vouchers": 26, "totalDebit": str(total), "totalCredit": str(total), "taxByProject": {key: str(value) for key, value in rates.items()}, "templateSupplierPrecedence": True, "hiddenColumnsAndSheetsPreserved": True, "allOtherZipMembersPreserved": True}


if __name__ == "__main__":
    action = sys.argv[1]
    directory = Path(sys.argv[2]).resolve()
    result = generate(directory) if action == "generate" else verify(directory, Path(sys.argv[3]).resolve()) if action == "verify" else None
    assert result is not None, "Expected generate or verify"
    print(json.dumps(result, ensure_ascii=False))
