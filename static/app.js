const STORAGE_KEY = "residential-comparison-workbench-v1";

const structureOptions = ["钢混结构", "砖混结构", "混合结构", "砖木结构"];
const orientationOptions = ["北", "西", "东", "一室朝南", "二室朝南", "三室朝南"];
const eastWestOptions = ["东边户", "中间户", "西边户"];
const regionalFactors = [
  ["natural_landscape", "自然景观"],
  ["environmental_pollution", "环境污染"],
  ["community_layout", "建筑或小区布局"],
  ["building_density_appearance", "建筑密度、外型等"],
  ["outdoor_space_landscaping", "室外公共活动空间与绿化景观"],
  ["repair_management", "房屋、设施维修与管理"],
  ["security_parking", "小区治安、道路停车管理"],
  ["sanitation_greening", "环卫与绿化管理"],
  ["main_road_distance", "离市区主干道的距离"],
  ["public_transport", "公交线路情况"],
  ["school_district", "所在学区学校情况"],
  ["medical_facilities", "医院及医疗机构分布情况"],
  ["cultural_recreation", "文体娱乐设施"],
  ["major_commercial", "重要商业配套设施"],
  ["market_supermarket", "菜场超市"],
];
const gradeLabels = ["一等", "二等", "三等", "四等", "五等"];
const regionalDefault = Object.fromEntries(
  regionalFactors.map(([key], index) => [key, index < 7 ? 2 : 3]),
);

const exampleState = {
  subject: {
    report_title: "示例花园5幢1单元1801室住宅房地产市场价值评估",
    report_number: "示例住宅估字[2026]001号",
    valuation_date: "2026-08-28",
    address: "示例市滨江区示范路88号示例花园5幢1单元1801室",
    owner: "测试产权人",
    certificate_no: "苏（2026）示例不动产权第TEST001号",
    co_ownership: "单独所有",
    real_estate_unit_no: "TEST-UNIT-001",
    land_area: 32543.2,
    land_use: "城镇住宅用地",
    rights_type: "国有建设用地使用权/房屋所有权",
    rights_nature: "出让/安置房",
    rights_term: "2018年04月16日起2088年04月15日止",
    parcel_no: "TEST-PARCEL-001",
    use_type: "成套住宅",
    location_boundaries: "东临测试路，西临示例路，南临规划路，北临公园路。",
    location_description: "估价对象位于测试居住片区，周边交通和生活配套完善。",
    location_analysis: "测试区域居住环境较好，区位条件对估价对象价值产生积极影响。",
    area: 121.43,
    certificate_floor_display: "18/19",
    below_ground_floors: 1,
    above_ground_total_floors: 18,
    floor: 18,
    build_year: 2019,
    structure: "钢混结构",
    building_type: "high_rise",
    orientation: "二室朝南",
    east_west: "东边户",
    property_nature: "安置房",
    area_factor_pct: 100,
    decoration_unit_price: 0,
    decoration_newness: 0,
    appurtenance_description: "/",
    appurtenance_unit_price: 0,
  },
  comparables: [
    {
      name: "可比实例A",
      source: "贝壳房产网",
      address: "示例市滨江区示范路88号示例花园1幢2单元1606室",
      area: 62.85,
      structure: "钢混结构",
      building_type: "high_rise",
      above_ground_total_floors: 18,
      floor: 16,
      build_year: 2019,
      orientation: "一室朝南",
      east_west: "中间户",
      property_nature: "安置房",
      area_factor_pct: 102,
      decoration_unit_price: 500,
      decoration_newness: 0.8,
      appurtenance_description: "/",
      appurtenance_unit_price: 0,
      transaction_date: "2026-08-01",
      transaction_unit_price: 3341,
      transaction_condition_factor: 1,
      market_condition_factor: 1,
      transaction_condition_note: "",
      market_condition_note: "",
    },
    {
      name: "可比实例B",
      source: "E路平安信息系统",
      address: "示例市滨江区示范路88号示例花园4幢1单元1601室",
      area: 102.04,
      structure: "钢混结构",
      building_type: "high_rise",
      above_ground_total_floors: 18,
      floor: 16,
      build_year: 2019,
      orientation: "二室朝南",
      east_west: "东边户",
      property_nature: "安置房",
      area_factor_pct: 100,
      decoration_unit_price: 500,
      decoration_newness: 0.8,
      appurtenance_description: "/",
      appurtenance_unit_price: 0,
      transaction_date: "2026-06-02",
      transaction_unit_price: 3214,
      transaction_condition_factor: 1,
      market_condition_factor: 1,
      transaction_condition_note: "",
      market_condition_note: "",
    },
    {
      name: "可比实例C",
      source: "E路平安信息系统",
      address: "示例市滨江区示范路88号示例花园3幢1单元903室",
      area: 121.22,
      structure: "钢混结构",
      building_type: "high_rise",
      above_ground_total_floors: 18,
      floor: 9,
      build_year: 2019,
      orientation: "二室朝南",
      east_west: "中间户",
      property_nature: "安置房",
      area_factor_pct: 100,
      decoration_unit_price: 500,
      decoration_newness: 0.8,
      appurtenance_description: "/",
      appurtenance_unit_price: 0,
      transaction_date: "2026-05-18",
      transaction_unit_price: 3217,
      transaction_condition_factor: 1,
      market_condition_factor: 1,
      transaction_condition_note: "",
      market_condition_note: "",
    },
  ],
  regional_grades: {
    subject: { ...regionalDefault },
    comparables: [{ ...regionalDefault }, { ...regionalDefault }, { ...regionalDefault }],
  },
};

let state = loadState();
let calculationTimer;
let latestResult = null;
let xlsxTemplate = null;
let toastTimer;

const $ = (selector) => document.querySelector(selector);
const comparablesElement = $("#comparables");
const templateLabel = $("#templateButtonLabel");
const templateStatus = $("#templateStatus");
const exportExcelButton = $("#exportExcel");

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed?.subject && parsed?.comparables?.length === 3) return parsed;
    }
  } catch (error) {
    console.warn("读取本地数据失败", error);
  }
  return clone(exampleState);
}

function saveLocal() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function numberValue(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formattedNumber(value, places = 0) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: places,
    maximumFractionDigits: places,
  }).format(Number(value));
}

function options(items, selected) {
  return items
    .map(
      (item) => {
        const [value, label] = Array.isArray(item) ? item : [item, item];
        return `<option value="${escapeHtml(value)}" ${value === selected ? "selected" : ""}>${escapeHtml(label)}</option>`;
      },
    )
    .join("");
}

function fillSubjectFields() {
  document.querySelector('[data-subject="structure"]').innerHTML = options(
    structureOptions,
    state.subject.structure,
  );
  document.querySelector('[data-subject="orientation"]').innerHTML = options(
    orientationOptions,
    state.subject.orientation,
  );
  document.querySelector('[data-subject="east_west"]').innerHTML = options(
    eastWestOptions,
    state.subject.east_west,
  );
  document.querySelectorAll("[data-subject]").forEach((input) => {
    input.value = state.subject[input.dataset.subject] ?? "";
  });
  renderFloorSummary();
}

function field(label, fieldName, value, type = "text", optionsText = "", extra = "") {
  const input =
    type === "select"
      ? `<select data-field="${fieldName}" ${extra}>${optionsText}</select>`
      : `<input data-field="${fieldName}" type="${type}" value="${escapeHtml(value)}" ${extra} />`;
  return `<label class="field"><span>${label}</span>${input}</label>`;
}

function comparableTemplate(item, index) {
  const letter = String.fromCharCode(65 + index);
  return `
    <article class="comparable-card" data-case-index="${index}">
      <div class="comparable-head">
        <div class="comparable-title">
          <span class="comparable-index">${letter}</span>
          <strong>${escapeHtml(item.name || `可比实例${letter}`)}</strong>
        </div>
        <span class="case-badge">固定三个案例</span>
      </div>
      <div class="comparable-grid">
        ${field("案例名称", "name", item.name)}
        ${field("数据来源", "source", item.source)}
        <div class="field field-wide"><span>坐落</span><input data-field="address" type="text" value="${escapeHtml(item.address)}" /></div>
        ${field("建筑面积（㎡）", "area", item.area, "number", "", 'min="0" step="0.01"')}
        ${field("建筑结构", "structure", item.structure, "select", options(structureOptions, item.structure))}
        ${field("建筑类型", "building_type", item.building_type, "select", options([
          ["multi_story", "普通多层"],
          ["elevator_multi_story", "电梯多层"],
          ["high_rise", "小高层 / 高层"],
        ], item.building_type))}
        ${field("所在层", "floor", item.floor, "number", "", 'min="1" step="1"')}
        ${field("地上总层数", "above_ground_total_floors", item.above_ground_total_floors, "number", "", 'min="1" step="1"')}
        ${field("建成年代", "build_year", item.build_year, "number", "", 'min="1900" max="2100" step="1"')}
        ${field("朝向", "orientation", item.orientation, "select", options(orientationOptions, item.orientation))}
        ${field("东西至", "east_west", item.east_west, "select", options(eastWestOptions, item.east_west))}
        ${field("房屋性质", "property_nature", item.property_nature)}
        ${field("建筑面积修正", "area_factor_pct", item.area_factor_pct, "number", "", 'step="0.1"')}
        ${field("装修重置单价", "decoration_unit_price", item.decoration_unit_price, "number", "", 'min="0" step="10"')}
        ${field("装修成新率", "decoration_newness", item.decoration_newness, "number", "", 'min="0" max="1" step="0.01"')}
        ${field("附属物情况", "appurtenance_description", item.appurtenance_description)}
        ${field("附属物扣减", "appurtenance_unit_price", item.appurtenance_unit_price, "number", "", 'min="0" step="1"')}
        ${field("成交日期", "transaction_date", item.transaction_date, "date")}
        ${field("成交单价（元/㎡）", "transaction_unit_price", item.transaction_unit_price, "number", "", 'min="0" step="1"')}
      </div>
      <div class="factor-pair">
        <label class="factor-box">
          <span>P1 交易情况修正</span>
          <div>
            <input data-field="transaction_condition_factor" type="number" min="0.001" step="0.001" value="${numberValue(item.transaction_condition_factor || 1)}" />
            <small>${(numberValue(item.transaction_condition_factor || 1) * 100).toFixed(1)}%</small>
          </div>
          <input class="note-input" data-field="transaction_condition_note" type="text" value="${escapeHtml(item.transaction_condition_note || "")}" placeholder="如需修正，填写依据" />
        </label>
        <label class="factor-box">
          <span>P2 市场状况修正</span>
          <div>
            <input data-field="market_condition_factor" type="number" min="0.001" step="0.001" value="${numberValue(item.market_condition_factor || 1)}" />
            <small>${(numberValue(item.market_condition_factor || 1) * 100).toFixed(1)}%</small>
          </div>
          <input class="note-input" data-field="market_condition_note" type="text" value="${escapeHtml(item.market_condition_note || "")}" placeholder="如需修正，填写依据" />
        </label>
      </div>
    </article>
  `;
}

function renderComparables() {
  comparablesElement.innerHTML = state.comparables
    .map((item, index) => comparableTemplate(item, index))
    .join("");
}

function renderRegionalFactors() {
  $("#regionalFactors").innerHTML = regionalFactors
    .map(
      ([key, label]) => `
        <tr>
          <th>${escapeHtml(label)}</th>
          ${["subject", 0, 1, 2]
            .map((scope) => {
              const value =
                scope === "subject"
                  ? state.regional_grades.subject[key]
                  : state.regional_grades.comparables[scope][key];
              return `<td><select class="regional-grade-select" data-scope="${scope}" data-factor="${key}">
                ${gradeLabels
                  .map(
                    (grade, index) =>
                      `<option value="${index + 1}" ${Number(value) === index + 1 ? "selected" : ""}>${grade}</option>`,
                  )
                  .join("")}
              </select></td>`;
            })
            .join("")}
        </tr>
      `,
    )
    .join("");
}

function renderFloorSummary() {
  const floor = numberValue(state.subject.floor);
  const total = numberValue(state.subject.above_ground_total_floors);
  $("#floorCorrectionSummary").textContent = `${floor}/${total}${floor === total ? "（顶层）" : ""}`;
}

function projectPayload() {
  return clone(state);
}

function scheduleCalculation(delay = 160) {
  saveLocal();
  renderFloorSummary();
  clearTimeout(calculationTimer);
  calculationTimer = setTimeout(calculate, delay);
}

async function calculate() {
  $("#calcState").textContent = "计算中";
  try {
    const response = await fetch("/api/calculate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(projectPayload()),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "测算失败");
    latestResult = data;
    renderResult(data);
    $("#calcState").textContent = "已同步";
  } catch (error) {
    latestResult = null;
    $("#calcState").textContent = "需要检查";
    $("#resultContent").classList.add("hidden");
    $("#resultError").classList.remove("hidden");
    $("#resultError").textContent = error.message;
  }
}

function renderResult(result) {
  $("#resultError").classList.add("hidden");
  $("#resultContent").classList.remove("hidden");
  $("#estimatedTotal").textContent = formattedNumber(result.estimated_total, 0);
  $("#estimatedUnitPrice").textContent = formattedNumber(result.estimated_unit_price, 0);
  $("#totalUppercase").textContent = result.total_uppercase;
  $("#spreadPct").textContent = `${result.spread_pct}%`;
  $("#regionalScore").textContent = result.subject.regional_score;
  const spread = Number(result.spread_pct);
  $("#qualityBlock").className = `quality-block${spread > 10 ? " caution" : ""}`;
  $("#qualityBlock").innerHTML = `<span class="quality-dot"></span><span>${
    spread > 10 ? "比准价格离散度超过10%，建议复核案例与修正系数。" : "比准价格差异处于复核提示范围内。"
  }</span>`;
  $("#warnings").innerHTML = (result.warnings || [])
    .map((warning) => `<li>${escapeHtml(warning)}</li>`)
    .join("");
  $("#comparisonTable").innerHTML = `
    <table>
      <thead><tr><th>案例</th><th>P1</th><th>P2</th><th>P3</th><th>P4</th><th>比准单价</th></tr></thead>
      <tbody>${result.comparables
        .map(
          (item) => `<tr><td>${escapeHtml(item.name)}</td><td>${item.p1}</td><td>${item.p2}</td><td>${item.p3}</td><td>${item.p4}</td><td>${formattedNumber(item.comparable_unit_price)}</td></tr>`,
        )
        .join("")}</tbody>
    </table>
    <div class="factor-details">
      ${result.comparables
        .map(
          (item) => `
            <details class="factor-detail" data-factor-details>
              <summary>${escapeHtml(item.name)} · 实体因素明细</summary>
              <dl>
                <div><dt>建筑面积</dt><dd>${item.factor_breakdown.area}</dd></div>
                <div><dt>结构</dt><dd>${item.factor_breakdown.structure}</dd></div>
                <div><dt>层次</dt><dd>${item.factor_breakdown.level}</dd></div>
                <div><dt>成新</dt><dd>${item.factor_breakdown.age}</dd></div>
                <div><dt>朝向</dt><dd>${item.factor_breakdown.orientation}</dd></div>
                <div><dt>东西至</dt><dd>${item.factor_breakdown.east_west}</dd></div>
                <div><dt>实体综合</dt><dd>${item.factor_breakdown.physical}%</dd></div>
              </dl>
            </details>
          `,
        )
        .join("")}
    </div>`;
}

function updateComparable(target) {
  const card = target.closest("[data-case-index]");
  if (!card || !target.dataset.field) return;
  const index = Number(card.dataset.caseIndex);
  const field = target.dataset.field;
  if (!state.comparables[index]) return;
  state.comparables[index][field] =
    target.type === "number" ? numberValue(target.value) : target.value;
  const factorBox = target.closest(".factor-box");
  if (factorBox && target.type === "number") {
    factorBox.querySelector("small").textContent = `${(numberValue(target.value) * 100).toFixed(1)}%`;
  }
  scheduleCalculation();
}

function showToast(message, error = false) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.className = `toast show${error ? " error" : ""}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.className = "toast";
  }, 2800);
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function fileBaseName() {
  return (state.subject.report_number || state.subject.report_title || "住宅估价")
    .replaceAll(/[<>:"/\\|?*]/g, "_")
    .slice(0, 80);
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

async function exportFile(kind, button) {
  if (!latestResult) {
    showToast("当前数据不能导出，请先修正测算错误", true);
    return;
  }
  if (kind === "residential-xlsx" && !xlsxTemplate) {
    showToast("请先选择 Excel 原表", true);
    return;
  }
  const original = button.innerHTML;
  button.disabled = true;
  button.innerHTML = "<span class='file-icon'>…</span><span><strong>生成中</strong><small>请稍候</small></span>";
  try {
    const body = {
      ...projectPayload(),
      template_base64: xlsxTemplate ? bytesToBase64(new Uint8Array(await xlsxTemplate.arrayBuffer())) : null,
      template_filename: xlsxTemplate?.name || null,
    };
    const response = await fetch(`/api/export/${kind}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || "导出失败");
    }
    const extension = kind === "residential-xlsx" ? "xlsx" : "docx";
    const suffix = kind === "residential-xlsx" ? "住宅测算表-已填写" : "正式估价报告";
    downloadBlob(await response.blob(), `${fileBaseName()}-${suffix}.${extension}`);
    showToast(kind === "residential-xlsx" ? "原表已填写并下载" : "Word 正式报告已下载");
  } catch (error) {
    showToast(error.message, true);
  } finally {
    button.disabled = kind === "residential-xlsx" && !xlsxTemplate;
    button.innerHTML = original;
  }
}

document.querySelectorAll("[data-subject]").forEach((input) => {
  input.addEventListener("input", () => {
    const field = input.dataset.subject;
    state.subject[field] = input.type === "number" ? numberValue(input.value) : input.value;
    scheduleCalculation();
  });
});

comparablesElement.addEventListener("input", (event) => {
  updateComparable(event.target);
});

$("#regionalFactors").addEventListener("change", (event) => {
  const select = event.target.closest(".regional-grade-select");
  if (!select) return;
  const scope = select.dataset.scope;
  const target =
    scope === "subject"
      ? state.regional_grades.subject
      : state.regional_grades.comparables[Number(scope)];
  target[select.dataset.factor] = Number(select.value);
  scheduleCalculation(0);
});

$("#xlsxTemplateFile").addEventListener("change", (event) => {
  xlsxTemplate = event.target.files?.[0] || null;
  templateLabel.textContent = xlsxTemplate ? xlsxTemplate.name.slice(0, 24) : "选择 Excel 原表";
  templateStatus.textContent = xlsxTemplate ? xlsxTemplate.name : "尚未选择 Excel 原表";
  exportExcelButton.disabled = !xlsxTemplate;
});

$("#saveProject").addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(projectPayload(), null, 2)], {
    type: "application/json;charset=utf-8",
  });
  downloadBlob(blob, `${fileBaseName()}-项目数据.json`);
});

$("#loadExample").addEventListener("click", () => {
  state = clone(exampleState);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  fillSubjectFields();
  renderComparables();
  renderRegionalFactors();
  calculate();
  showToast("样例数据已恢复");
});

$("#importFile").addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const imported = JSON.parse(await file.text());
    if (!imported.subject || imported.comparables?.length !== 3 || !imported.regional_grades) {
      throw new Error("文件中缺少完整的住宅测算数据");
    }
    state = imported;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    fillSubjectFields();
    renderComparables();
    renderRegionalFactors();
    calculate();
    showToast("项目数据已导入");
  } catch (error) {
    showToast(`导入失败：${error.message}`, true);
  } finally {
    event.target.value = "";
  }
});

$("#exportWord").addEventListener("click", (event) => {
  exportFile("residential-docx", event.currentTarget);
});

$("#exportExcel").addEventListener("click", (event) => {
  exportFile("residential-xlsx", event.currentTarget);
});

fillSubjectFields();
renderComparables();
renderRegionalFactors();
calculate();
