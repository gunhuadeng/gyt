# 住宅市场比较法测算与报告工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the generic prototype calculation with the user's exact residential tax-valuation workbook rules, expose editable P1/P2 adjustments, and export a filled copy of the original XLSX plus a consistent Word draft.

**Architecture:** Keep the local Python web app. Add a deterministic residential calculation engine and rule configuration, then patch only the input cells of a user-supplied XLSX template while preserving every other workbook part. The frontend submits one project JSON object to calculation and export endpoints.

**Tech Stack:** Python 3.10+ standard library, modular browser HTML/CSS/JavaScript, `unittest`, direct OOXML package manipulation.

**Spec:** `docs/superpowers/specs/2026-11-05-residential-comparison-engine-design.md`

## Global Constraints

- Use exactly three active comparable cases in the first version.
- Preserve certificate floor display and calculation floor basis separately: certificate `18/19`, above-ground `18/18`, correction coefficient `99.4`.
- P1 transaction-condition and P2 market-condition factors are independently editable per case, default `1.000`, and must only be greater than zero.
- Use arithmetic mean of the three comparable unit prices, not weighted average.
- Never modify the user's source XLSX.
- Do not send property data to external services.
- Use decimal arithmetic and Excel-compatible rounding.
- All output is a professional-review draft, not an automatically signed valuation report.
- This workspace is not a Git repository, so implementation tasks end with tests instead of commits.

## Review Focus

- Certificate floors and above-ground floors differ: correction must use above-ground floors while display retains both.
- P1 or P2 is non-neutral and later saved/reopened/exported: the exact factor and explanation must survive every path.
- A comparable is incomplete or has an invalid regional grade: calculation must stop with a field-specific message.
- The supplied XLSX lacks an expected sheet or cell: export must fail safely without producing a misleading file.
- Web calculation and exported workbook disagree: golden values `2,779` and `337,454` must fail the build before delivery.

---

### Task 1: Deterministic Rules and Golden Calculation

**Files:**
- Create: `residential_rules.py`
- Create: `residential_calculator.py`
- Create: `tests/test_residential_calculator.py`
- Test: `tests/test_residential_calculator.py`

**Interfaces:**
- Consumes: project JSON with `subject`, `comparables`, and `regional_grades`.
- Produces: `ResidentialCalculationError`; `calculate_residential_comparison(project: dict) -> dict`; `rmb_uppercase(value) -> str`.

- [ ] **Step 1: Write failing golden and floor tests**

```python
def test_reference_project_matches_2779_and_337454(self):
    result = calculate_residential_comparison(reference_project())
    self.assertEqual(result["estimated_unit_price"], "2779")
    self.assertEqual(result["estimated_total"], "337454")
    self.assertEqual(result["total_uppercase"], "人民币叁拾叁万柒仟肆佰伍拾肆元整")
    self.assertEqual([x["comparable_unit_price"] for x in result["comparables"]], ["2824", "2734", "2778"])

def test_certificate_and_correction_floors_are_separate(self):
    result = calculate_residential_comparison(reference_project())
    subject = result["subject"]
    self.assertEqual(subject["certificate_floor_display"], "18/19")
    self.assertEqual(subject["above_ground_total_floors"], 18)
    self.assertEqual(subject["floor_correction_display"], "18/18")
    self.assertEqual(subject["level_factor_pct"], "99.40")
```

- [ ] **Step 2: Run tests and verify failure**

Run: `python -m unittest tests.test_residential_calculator -v`

Expected: import failure because `residential_calculator` does not exist.

- [ ] **Step 3: Create the rule tables and calculation interface**

In `residential_rules.py`, encode the 15 regional factors, their first-to-fifth grade scores, structure factors, orientation factors, townhouse east/west factors, and high-rise floor rules from the design document.

In `residential_calculator.py`, define:

```python
class ResidentialCalculationError(ValueError): ...

def calculate_residential_comparison(project: dict) -> dict: ...

def rmb_uppercase(value: Any) -> str: ...
```

Compute `P3`, `P4`, `Va`, `Vb`, each `Vn`, the arithmetic mean, total value, uppercase amount, and spread. Return strings for displayed decimals.

- [ ] **Step 4: Run the focused tests**

Run: `python -m unittest tests.test_residential_calculator -v`

Expected: all tests pass, including golden `2779` and `337454`.

- [ ] **Step 5: Add and pass P1/P2 persistence tests**

```python
def test_non_neutral_p1_and_p2_affect_result(self):
    project = reference_project()
    project["comparables"][0]["transaction_condition_factor"] = 0.98
    project["comparables"][0]["market_condition_factor"] = 1.03
    result = calculate_residential_comparison(project)
    self.assertEqual(result["comparables"][0]["p1"], "0.980")
    self.assertEqual(result["comparables"][0]["p2"], "1.030")
    self.assertNotEqual(result["estimated_unit_price"], "2779")
```

Run: `python -m unittest tests.test_residential_calculator -v`

Expected: pass.

### Task 2: Validation and Input Normalization

**Files:**
- Modify: `residential_calculator.py`
- Modify: `tests/test_residential_calculator.py`

**Interfaces:**
- Consumes: raw frontend dictionaries and JSON-imported dictionaries.
- Produces: normalized typed values and path-specific validation errors.

- [ ] **Step 1: Write failing validation tests**

```python
def test_rejects_missing_third_case(self):
    project = reference_project()
    project["comparables"] = project["comparables"][:2]
    with self.assertRaisesRegex(ResidentialCalculationError, "必须填写 3 个"):
        calculate_residential_comparison(project)

def test_rejects_non_positive_p1(self):
    project = reference_project()
    project["comparables"][0]["transaction_condition_factor"] = 0
    with self.assertRaisesRegex(ResidentialCalculationError, "交易情况修正系数"):
        calculate_residential_comparison(project)

def test_rejects_invalid_regional_grade(self):
    project = reference_project()
    project["regional_grades"]["subject"]["natural_landscape"] = 6
    with self.assertRaisesRegex(ResidentialCalculationError, "自然景观"):
        calculate_residential_comparison(project)
```

- [ ] **Step 2: Run tests and verify failure**

Run: `python -m unittest tests.test_residential_calculator -v`

Expected: validation assertions fail because checks are absent.

- [ ] **Step 3: Implement validation**

Validate required fields, positive areas and prices, exactly three cases, grades 1–5, positive P1/P2, floor not above above-ground total, and complete regional-grade dictionaries. Error messages include the case and field label.

- [ ] **Step 4: Run the full calculator tests**

Run: `python -m unittest tests.test_residential_calculator -v`

Expected: all calculator tests pass.

### Task 3: Browser Workbench and Local Persistence

**Files:**
- Modify: `static/index.html`
- Modify: `static/app.js`
- Modify: `static/styles.css`
- Create: `sample-project.json`

**Interfaces:**
- Consumes: `/api/calculate` returning the Task 1 result shape.
- Produces: complete `project` object with `subject`, three `comparables`, `regional_grades`, and `adjustment_notes`.

- [ ] **Step 1: Add a browser smoke test script**

Create `tests/browser_smoke.mjs` using Playwright Core in the existing development cache. It checks page load, three case cards, the calculated unit price, P1/P2 inputs, JSON persistence, and no console errors.

- [ ] **Step 2: Run the smoke test and verify failure**

Run: `node tests/browser_smoke.mjs`

Expected: failure because the new fields and endpoints are absent.

- [ ] **Step 3: Rebuild the form**

Provide project/object fields, three case cards, 15 regional grade selectors, editable P1/P2 with explanations, correction traces, warnings, reset, JSON save/import, and Word/Excel export actions.

- [ ] **Step 4: Run browser smoke test**

Run: `node tests/browser_smoke.mjs`

Expected: pass with sample unit price `2,779` and no console errors.

### Task 4: Original XLSX Template Export

**Files:**
- Create: `xlsx_template.py`
- Modify: `app.py`
- Create: `tests/test_xlsx_template.py`

**Interfaces:**
- Consumes: `build_filled_workbook(template_bytes: bytes, project: dict, result: dict) -> bytes`.
- Produces: filled XLSX bytes with original non-input package parts preserved; `TemplateFormatError`.

- [ ] **Step 1: Write failing template tests**

```python
def test_preserves_hidden_sheets_and_images(self):
    output = build_filled_workbook(REFERENCE_XLSX.read_bytes(), reference_project(), reference_result())
    before = set(zipfile.ZipFile(io.BytesIO(REFERENCE_XLSX.read_bytes())).namelist())
    after = set(zipfile.ZipFile(io.BytesIO(output)).namelist())
    self.assertTrue(after.issuperset(before - {"xl/workbook.xml", "xl/worksheets/sheet5.xml", "xl/calcChain.xml"}))

def test_writes_non_neutral_p1_and_p2(self):
    project = reference_project()
    project["comparables"][0]["transaction_condition_factor"] = 0.98
    project["comparables"][0]["market_condition_factor"] = 1.03
    output = build_filled_workbook(REFERENCE_XLSX.read_bytes(), project, reference_result())
    sheet = read_sheet_xml(output, "1-1 测算表")
    self.assertEqual(cell_value(sheet, "C54"), "0.98")
    self.assertEqual(cell_value(sheet, "C55"), "1.03")

def test_rejects_template_without_calculation_sheet(self):
    broken = remove_sheet_from_xlsx(REFERENCE_XLSX.read_bytes(), "1-1 测算表")
    with self.assertRaisesRegex(TemplateFormatError, "1-1 测算表"):
        build_filled_workbook(broken, reference_project(), reference_result())
```

- [ ] **Step 2: Run tests and verify failure**

Run: `python -m unittest tests.test_xlsx_template -v`

Expected: import failure because `xlsx_template` does not exist.

- [ ] **Step 3: Implement safe template patching**

Map logical fields to these `1-1 测算表` input cells:

```text
D6:G6 地址
D7:G7 建筑面积
D8:G8 结构
D9:G9 楼层修正显示
D10:G10 建成年代
D11:G11 朝向
D12:G12 东西至
D14:G14 附属物说明
D15:G15 房屋性质
D16:G30 区域因素等级
D31:G31 交易日期
D32:G32 成交单价
B43:E43 建筑面积修正
B45:E45 层次系数
B48:E48 东西至系数
B51:E51 装修重置单价
B52:E52 装修成新率
B53:E53 P4 区域状况修正
B54:E54 P1 交易情况修正
B55:E55 P2 市场状况修正
B58:E58 Vb 附属物扣减
```

Preserve formula-only cells such as D13:G13, B46:E46, B47:E47, B49:E49, B50:E50, B56:E56, B57:E57, B60:E60, B72:E72, and B75:E77 unless a future explicit override feature is approved.

Preserve all ZIP entries not intentionally replaced, patch only `<c>` values, and set `fullCalcOnLoad="1"`. Raise `TemplateFormatError` when the template lacks the expected sheet or required cells.

- [ ] **Step 4: Run template tests**

Run: `python -m unittest tests.test_xlsx_template -v`

Expected: pass; hidden sheets and images count remains unchanged.

### Task 5: Word Draft and End-to-End Export

**Files:**
- Modify: `office_export.py`
- Modify: `app.py`
- Modify: `tests/test_exports.py`
- Modify: `README.md`

**Interfaces:**
- Consumes: the Task 1 result and Task 3 project shape.
- Produces: `build_residential_docx(project: dict, result: dict) -> bytes`; `/api/export/residential-xlsx` and `/api/export/residential-docx`.

- [ ] **Step 1: Write failing consistency tests**

```python
def test_docx_contains_same_result_as_calculator(self):
    project = reference_project()
    result = calculate_residential_comparison(project)
    package = open_docx(build_residential_docx(project, result))
    self.assertIn("2,779", package.document_text)
    self.assertIn("337,454", package.document_text)
    self.assertIn("18/19（地下一层、地上18层、顶层）", package.document_text)
```

- [ ] **Step 2: Run tests and verify failure**

Run: `python -m unittest tests.test_exports -v`

Expected: failure because the residential Word builder is absent.

- [ ] **Step 3: Implement Word mapping**

Generate a consistent draft with project details, floor-basis explanation, comparable and correction tables, P1/P2 explanations, result, and professional-review notice. Update README with the exact workflow and the `.docx` template limitation.

- [ ] **Step 4: Run all tests**

Run: `python -m unittest discover -s tests -p 'test_*.py' -v`

Expected: all tests pass.

### Task 6: Final Regression and Delivery

**Files:**
- Modify: `README.md`
- Create: `docs/verification/2026-11-05-residential-engine.md`

**Interfaces:**
- Consumes: all completed components.
- Produces: a tested runnable package and evidence report.

- [ ] **Step 1: Run complete automated verification**

Run: `python -m unittest discover -s tests -p 'test_*.py' -v`

Expected: zero failures and zero errors.

- [ ] **Step 2: Run browser and export smoke tests**

Run: `node tests/browser_smoke.mjs`

Expected: HTTP 200, sample `2,779`, non-neutral P1/P2 persistence, both Office downloads, and no browser errors.

- [ ] **Step 3: Verify Office package integrity**

Run: `python tests/verify_office_packages.py`

Expected: ZIP integrity valid, all XML parts parse, required workbook sheets and document text present.

- [ ] **Step 4: Record evidence and package**

Record exact commands and observed output in `docs/verification/2026-11-05-residential-engine.md`; create a new ZIP without overwriting the prior prototype.
