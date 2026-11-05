import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const port = 8876;
const baseUrl = `http://127.0.0.1:${port}`;
const playwrightPath =
  process.env.PLAYWRIGHT_CORE_PATH ||
  "C:/Users/Administrator/Documents/Codex/2026-11-05/new-chat-2/work/npm-cache/_npx/423231821c231c73/node_modules/playwright-core";
const pythonExecutable =
  process.env.PYTHON_EXE ||
  "C:/Users/Administrator/.pyenv/pyenv-win/versions/3.14.1/python.exe";
const referenceWorkbook = process.env.REFERENCE_XLSX;
const require = createRequire(import.meta.url);
const { chromium } = require(playwrightPath);

const server = spawn(pythonExecutable, ["app.py", "--no-browser", "--port", String(port)], {
  cwd: projectDir,
  stdio: ["ignore", "pipe", "pipe"],
  windowsHide: true,
});

let serverOutput = "";
server.stdout.on("data", (chunk) => {
  serverOutput += chunk.toString();
});
server.stderr.on("data", (chunk) => {
  serverOutput += chunk.toString();
});

async function waitForServer() {
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/api/health`);
      if (response.ok) return;
    } catch {
      // The server is still starting.
    }
    await new Promise((resolve) => setTimeout(resolve, 150));
  }
  throw new Error(`Server did not start.\n${serverOutput}`);
}

let browser;
try {
  await waitForServer();
  if (!referenceWorkbook) throw new Error("Set REFERENCE_XLSX to the source workbook path");
  browser = await chromium.launch({ channel: "msedge", headless: true });
  const page = await browser.newPage({
    viewport: { width: 1440, height: 1100 },
    acceptDownloads: true,
  });
  const browserErrors = [];
  page.on("console", (message) => {
    if (message.type() === "error") browserErrors.push(message.text());
  });
  page.on("pageerror", (error) => browserErrors.push(error.message));

  const response = await page.goto(baseUrl, {
    waitUntil: "networkidle",
    timeout: 30000,
  });
  if (response?.status() !== 200) throw new Error(`Unexpected HTTP ${response?.status()}`);
  if (!(await page.title()).includes("住宅市场比较法")) {
    throw new Error(`Unexpected title: ${await page.title()}`);
  }

  await page.waitForSelector('[data-testid="estimated-unit-price"]', { timeout: 15000 });
  await page.waitForFunction(
    () => document.querySelector('[data-testid="estimated-unit-price"]')?.textContent?.includes("2,779"),
    null,
    { timeout: 15000 },
  );

  if ((await page.locator("[data-case-index]").count()) !== 3) {
    throw new Error("Expected exactly three comparable case cards");
  }
  if ((await page.locator(".regional-grade-select").count()) !== 60) {
    throw new Error("Expected 60 regional grade selectors");
  }
  if ((await page.locator("[data-factor-details]").count()) !== 3) {
    throw new Error("Expected three comparable factor-detail panels");
  }

  await page.locator('[data-case-index="0"] [data-field="transaction_condition_factor"]').fill("0.98");
  await page.locator('[data-case-index="0"] [data-field="market_condition_factor"]').fill("1.03");
  await page.waitForFunction(
    () => !document.querySelector('[data-testid="estimated-unit-price"]')?.textContent?.includes("2,779"),
    null,
    { timeout: 15000 },
  );

  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 15000 }),
    page.locator('[data-testid="save-project"]').click(),
  ]);
  const saved = JSON.parse(await fs.readFile(await download.path(), "utf8"));
  if (saved.comparables[0].transaction_condition_factor !== 0.98) {
    throw new Error("P1 did not survive JSON save");
  }
  if (saved.comparables[0].market_condition_factor !== 1.03) {
    throw new Error("P2 did not survive JSON save");
  }

  await page.locator("#xlsxTemplateFile").setInputFiles(referenceWorkbook);
  const exports = [];
  for (const [selector, type] of [
    ["[data-testid='export-word']", "docx"],
    ["[data-testid='export-excel']", "xlsx"],
  ]) {
    const [download] = await Promise.all([
      page.waitForEvent("download", { timeout: 20000 }),
      page.locator(selector).click(),
    ]);
    const bytes = await fs.readFile(await download.path());
    exports.push({
      type,
      name: download.suggestedFilename(),
      size: bytes.length,
      magic: bytes.subarray(0, 2).toString(),
    });
  }
  if (exports.some((item) => item.magic !== "PK" || item.size < 3000)) {
    throw new Error(`Office exports failed: ${JSON.stringify(exports)}`);
  }

  await page.locator('[data-testid="load-example"]').click();
  await page.waitForFunction(
    () => document.querySelector('[data-testid="estimated-unit-price"]')?.textContent?.includes("2,779"),
    null,
    { timeout: 15000 },
  );
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: path.join(projectDir, "preview.png"),
    fullPage: true,
  });
  if (browserErrors.length) {
    throw new Error(`Browser errors:\n${browserErrors.join("\n")}`);
  }
  console.log("browser smoke: OK");
  console.log("estimated unit price: 2,779");
  console.log("P1/P2 persistence: OK");
  console.log(`Office exports: ${JSON.stringify(exports)}`);
} finally {
  await browser?.close();
  server.kill();
}
