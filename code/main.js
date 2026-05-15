/**
 * ERP 刊登管理：从 Excel 读取 SKU，在「产品列表」中逐条查询，
 * 按规则写入第二列「判断结果」，并保存回原 Excel。
 *
 * 调试：
 * - set JINPU_DEBUG=1 && node code/main.js          （Windows CMD）
 * - $env:JINPU_DEBUG=1; node code/main.js           （PowerShell）
 * - node code/main.js --debug                       （等同 JINPU_DEBUG=1）
 * - node code/main.js --trace                       （结束生成 trace.zip，用 npx playwright show-trace trace.zip）
 * - node code/main.js --pause                       （进入「可刊登」变体流程前 page.pause()，接 Playwright Inspector）
 * - node code/main.js --slow                       （每步成功后多等 ~800ms，网卡时更稳）
 * - $env:JINPU_STEP_GAP_MS="500"                     （自定义每步间隔毫秒；设为 0 则不加间隔）
 *
 * 选图检查结果：终端搜 `[jinpu] 选图检查` 可看每行「有效配图数」与 PASS/FAIL；汇总行含 imagesCheck=pass|fail|skipped。
 *
 * 若出现 net::ERR_CONNECTION_RESET：多为网络/VPN/防火墙中断，脚本会对登录页 goto 自动重试；仍失败请浏览器手动打开同一网址排查。
 *
 * 变体流程步骤代号（卡住或报错时，在终端会看到「请搜索: Vxx_…」，在 main.js 里 Ctrl+F 该字符串即可跳到对应 runPlaywrightStep）：
 *   V01_PAUSE          page.pause()（仅 --pause）
 *   V03_FINE_CLICK     点击「精细刊登」
 *   V04_NEXT_CLICK     点击「下一步」
 *   V05_POST_NEXT_WAIT 下一步后等待 2s
 *   V07_SHOP_OPEN      点开店铺「请选择」
 *   V08_J6_CLICK       点选「J6」
 *   V09_TBODY_WAIT     在「产品信息」标题并列的 div.body 内等待变体行（header/body 并行结构）
 *   V10_TABLE_STABLE   再等 3s
 *   V11_SKU_NodeID  列表页「产品信息」区块抓 SKU 对应的 Node ID
 *   V12_SKU_Title 列表页「产品信息」区块抓 SKU 对应的标题
 *   V13_SCRAPE         抓取变体产品状态
 *   V13B_VARIANT_IMAGES 状态均通过后逐行点「选择图片」，检查弹层 div.box 内是否有配图
 *   V14_CLOSE_DIALOG   关刊登向导：取消 → 确定关闭（与手动两步一致）
 */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const readline = require('readline');
const XLSX = require('xlsx');

const CLI_ARGS = process.argv.slice(2).filter((a) => !String(a).startsWith('-'));
const DEBUG =
  process.env.JINPU_DEBUG === '1' ||
  process.argv.includes('--debug') ||
  process.argv.includes('-d');
const TRACE = process.argv.includes('--trace');
const PAUSE_BEFORE_VARIANT = process.argv.includes('--pause');
const CONFIRM_MODE = process.argv.includes('--confirm');

/** 每步成功后的短暂停顿，给接口/DOM 一点时间（主要靠 waitFor；此项兜底弱网）。 */
function resolveStepGapMs() {
  if (process.argv.includes('--slow')) return 800;
  const raw = process.env.JINPU_STEP_GAP_MS;
  if (raw === undefined || raw === '') return 350;
  const n = Number(raw);
  return Number.isFinite(n) && n >= 0 ? n : 350;
}
const STEP_GAP_MS = resolveStepGapMs();
let confirmEnabled = CONFIRM_MODE;

async function waitForConfirm(sku, label) {
  console.error('');
  console.error('╔════════════════════════════════════════════════════════════════╗');
  console.error(`║ [确认模式] SKU=${sku}  判断结果: ${label}`);
  console.error('║ 按 Enter 继续下一个 SKU；输入 off 关闭确认模式自动运行');
  console.error('╚════════════════════════════════════════════════════════════════╝');
  const rl = readline.createInterface({ input: process.stdin, output: process.stderr });
  const answer = await new Promise((resolve) => rl.question('', resolve));
  rl.close();
  const cmd = answer.trim().toLowerCase();
  if (cmd === 'off' || cmd === 'auto' || cmd === 'close') {
    confirmEnabled = false;
    console.error('[jinpu] 确认模式已关闭，后续 SKU 自动运行');
  }
}

async function stepBreath(page) {
  if (STEP_GAP_MS > 0 && page) await page.waitForTimeout(STEP_GAP_MS);
}

function debugLog(message, extra) {
  if (!DEBUG) return;
  const ts = new Date().toISOString();
  if (extra !== undefined) console.error(`[jinpu ${ts}]`, message, extra);
  else console.error(`[jinpu ${ts}]`, message);
}

async function debugScreenshot(page, basename) {
  const dir = path.join(__dirname, '..', 'debug-screenshots');
  fs.mkdirSync(dir, { recursive: true });
  const safe = String(basename).replace(/[^\w.-]+/g, '_');
  const file = path.join(dir, `${safe}-${Date.now()}.png`);
  await page.screenshot({ path: file, fullPage: true }).catch(() => {});
  console.error(`[jinpu] 已保存截图: ${file}`);
}

/**
 * 包一层步骤：终端始终打印「進行到哪一步」；失败时打印可在 main.js 里搜索的代号，并保存含代号的截图。
 * @param {string} stepCode 与文件头注释一致，便于 Ctrl+F
 */
async function runPlaywrightStep(stepCode, description, { page, sku }, fn) {
  console.error(`[jinpu] ▶ ${stepCode} | ${description} | SKU=${sku || '-'}`);
  debugLog(`STEP ${stepCode}`, description);
  try {
    const ret = await fn();
    console.error(`[jinpu] ✓ ${stepCode}`);
    await stepBreath(page);
    return ret;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (err instanceof Error) err.jinpuFailedStep = stepCode;
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════════╗');
    console.error(`║ [jinpu 失败] 在文件 code/main.js 里按 Ctrl+F 搜索:  ${stepCode}`);
    console.error('╠════════════════════════════════════════════════════════════════╣');
    console.error(`║ 步骤说明: ${description}`);
    console.error(`║ SKU: ${sku || '(无)'}`);
    console.error(`║ 错误信息: ${msg}`);
    console.error('╚════════════════════════════════════════════════════════════════╝');
    console.error('');
    await debugScreenshot(page, `FAIL-${stepCode}-SKU-${sku || 'unknown'}`);
    throw err;
  }
}

// 列表页表格列顺序（与页面 DOM 表头一致，用于 evaluate 里组装行对象）
const LIST_TABLE_COLUMN_KEYS = [
  '序号',
  '操作',
  'SKU',
  '已刊登店铺',
  '缩略图',
  '标题',
  '类目',
  '产品状态',
  '是否侵权违禁',
  '来源',
  '是否为代销',
];

/** 解析 SKU Excel：支持 node code/main.js "D:\\path\\你的表.xlsx"（会与 --debug/--trace 等开关共存） */
function resolveSkuExcelPath() {
  const fromArg = CLI_ARGS[0];
  if (fromArg) {
    const resolved = path.resolve(fromArg);
    if (!fs.existsSync(resolved)) {
      throw new Error(`指定的 Excel 不存在: ${resolved}`);
    }
    return resolved;
  }
  // 默认查找顺序：上级目录 skus → 当前工作目录 skus → data/skus
  const candidates = [
    path.join(__dirname, '..', 'skus.xlsx'),
    path.join(process.cwd(), 'skus.xlsx'),
    path.join(__dirname, '..', 'data', 'skus.xlsx'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return path.resolve(p);
  }
  // 都不存在则生成示例文件（仅演示 SKU，需自行改成真实列表）
  const out = path.join(__dirname, '..', 'data', 'skus.xlsx');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([['SKU'], ['3117200063911']]);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, out);
  console.error(
    `未找到 skus.xlsx，已生成示例（仅示例 SKU，请自行改成你的列表）：\n  ${out}\n也可把 Excel 命名为 skus.xlsx 放在项目根目录 ${path.join(__dirname, '..')} ，或运行时传入路径。`
  );
  return out;
}

/** 读取工作簿，返回首表 A 列为纯数字的 SKU 所在行（0-based 行号）；B 列已有判断结果的行跳过并记入 labelsByRowIndex */
function loadSkuRowsFromWorkbook(filePath, labelsByRowIndex) {
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
  const skuRows = [];
  matrix.forEach((row, rowIndex) => {
    const s = String(row[0] ?? '').trim();
    if (!s || !/^\d+$/.test(s)) return; // 跳过表头或非数字 SKU
    const existingLabel = String(row[1] ?? '').trim();
    if (existingLabel) {
      labelsByRowIndex.set(rowIndex, existingLabel);
      return;
    }
    skuRows.push({ sku: s, rowIndex });
  });
  return { wb, sheetName, matrix, skuRows };
}

/** 已刊登店铺中是否包含 J6 系店铺（如 J6CA、J6MX）；仅有 A8、J5 等不算 */
function hasJ6PublishedStore(shopsStr) {
  const raw = String(shopsStr ?? '').trim();
  if (!raw) return false;
  return raw
    .split(/[,，]/)
    .map((s) => s.trim())
    .filter(Boolean)
    .some((code) => /^J6/i.test(code));
}

/**
 * 第二列文案规则（列表页首行）：
 * - 侵权违禁 ≠「否」时拼接刊登店与违禁说明；若仅非 J6 店铺有刊登则加「其他店铺已刊登（非J6）」；
 * - 无侵权时：J6 已刊登 → 产品已刊登；否则列表「产品状态」非在售中 → 直接写该状态；否则先记为「可刊登」，再由刊登页变体检核覆盖。
 */
function deriveSecondColumnLabel(row) {
  const shops = String(row['已刊登店铺'] ?? '').trim();
  const status = String(row['产品状态'] ?? '').trim();
  const infringement = String(row['是否侵权违禁'] ?? '').trim();
  const hasAnyCell = [shops, status, infringement, String(row['SKU'] ?? '').trim()].some(Boolean);
  if (!hasAnyCell) return '';
  const hasJ6 = hasJ6PublishedStore(shops);
  const isInfringe = infringement !== '' && infringement !== '否';
  if (isInfringe) {
    const parts = [];
    if (hasJ6) {
      parts.push('产品已刊登');
    } else if (shops) {
      parts.push('其他店铺已刊登（非J6）');
    }
    parts.push(`${infringement}`);
    if (status && status !== '在售中') parts.push(status);
    return parts.join('；');
  }
  if (hasJ6) return '产品已刊登';
  if (status !== '在售中') return status;
  return '可刊登';
}

/**
 * 精细刊登向导可视容器：必须含「产品信息」文案，避免匹配列表页其它 vxe 表格。
 */
function listingWizardPanel(page) {
  const byRole = page.getByRole('dialog').filter({ hasText: '产品信息' });
  const elDlg = page.locator('.el-dialog__wrapper:visible .el-dialog').filter({ hasText: '产品信息' });
  const ivu = page.locator('.ivu-modal-wrap:visible .ivu-modal').filter({ hasText: '产品信息' });
  return byRole.or(elDlg).or(ivu).first();
}

/**
 * 「产品信息」区块的 .body（与 .header 并列：header 内 h5 标题，body 内才是 vxe 表）。
 * 先在向导弹层内找；XPath：含「产品信息」的 div.header → 父级 → 紧随其后的 div.body。
 */
function listingProductInfoBody(page) {
  const wizard = listingWizardPanel(page);
  const underWizard = wizard.locator(
    'xpath=.//div[contains(@class,"header")][.//h5[normalize-space()="产品信息"]]/following-sibling::div[contains(concat(" ", normalize-space(@class), " "), " body ")][1]'
  );
  const globalFallback = page.locator(
    'xpath=//div[contains(@class,"header")][.//h5[normalize-space()="产品信息"]]/following-sibling::div[contains(concat(" ", normalize-space(@class), " "), " body ")][1]'
  );
  return underWizard.or(globalFallback).first();
}

/**
 * iView：点向导「取消」后出现二次确认，必须再点「确定关闭」，否则会挡住列表「查询」。
 * 与手动一致：直接 getByRole('确定关闭')，不设 .v-transfer-dom 嵌套定位（易 wait 不到）。
 */
async function clickConfirmCloseIfPresent(page) {
  const ok = page.getByRole('button', { name: '确定关闭' });
  if (await ok.count() === 0) return;
  try {
    await ok.waitFor({ state: 'visible', timeout: 3_000 });
    await ok.click({ timeout: 15_000 });
    await page.waitForTimeout(400);
  } catch {
    /* 当前没有残留确认框 */
  }
}

/**
 * 刊登向导「产品信息」表格：读取各变体行的「产品状态」。
 * 仅在「产品信息」并列的 .body 内查找：优先 .tid_172，否则用含 seller_sku: 的行定位 .vxe-table。
 */
async function scrapeTid172VariantStatuses(page) {
  return page.evaluate(() => {
    function isHidden(el) {
      if (!(el instanceof HTMLElement)) return true;
      const st = window.getComputedStyle(el);
      return st.display === 'none' || st.visibility === 'hidden' || st.opacity === '0';
    }

    function findProductInfoBodyUnder(root) {
      if (!root) return null;
      const headers = root.querySelectorAll('div.header');
      for (const hb of headers) {
        const h5 = hb.querySelector('h5');
        if (!h5 || (h5.textContent || '').trim() !== '产品信息') continue;
        const wrap = hb.parentElement;
        const bodyEl = wrap?.querySelector(':scope > div.body');
        if (bodyEl instanceof HTMLElement && !isHidden(bodyEl)) return bodyEl;
      }
      return null;
    }

    function findProductInfoBodyEl() {
      const tryRoots = (selector) => Array.from(document.querySelectorAll(selector));
      const wizardCandidates = [
        ...tryRoots('[role="dialog"]'),
        ...tryRoots('.el-dialog__wrapper .el-dialog'),
        ...tryRoots('.ivu-modal-wrap .ivu-modal'),
      ];
      for (const root of wizardCandidates) {
        if (!(root instanceof HTMLElement) || isHidden(root)) continue;
        const inner = root.innerText || '';
        if (!inner.includes('产品信息')) continue;
        const b = findProductInfoBodyUnder(root);
        if (b) return b;
      }
      return findProductInfoBodyUnder(document.body);
    }

    function findVariantTableRoot(scopeEl) {
      if (!scopeEl) return null;
      const tid = scopeEl.querySelector('.tid_172');
      if (tid?.querySelector?.('.body--wrapper tbody tr.vxe-body--row')) return tid;
      const variantRow = Array.from(scopeEl.querySelectorAll('tbody tr.vxe-body--row')).find((tr) =>
        /seller_sku:/i.test(tr.textContent || '')
      );
      if (!variantRow) return null;
      return variantRow.closest('.vxe-table') || variantRow.closest('.tid_172');
    }

    const bodyScope = findProductInfoBodyEl();
    if (!bodyScope) return { ok: false, statuses: [], reason: 'no_product_info_body' };
    const block = findVariantTableRoot(bodyScope);
    if (!block) return { ok: false, statuses: [], reason: 'no_variant_table_in_product_body' };

    let idx = -1;
    const headerCols = Array.from(block.querySelectorAll('.vxe-header--column'));
    if (headerCols.length) {
      idx = headerCols.findIndex((col) =>
        (col.textContent || '').replace(/\s+/g, '').includes('产品状态')
      );
    }
    if (idx < 0) {
      const ths = Array.from(block.querySelectorAll('thead th'));
      idx = ths.findIndex((th) => (th.textContent || '').includes('产品状态'));
    }

    const tbody =
      block.querySelector('.body--wrapper tbody') || block.querySelector('tbody');
    if (!tbody) return { ok: false, statuses: [], reason: 'no_tbody' };

    const trs = Array.from(tbody.querySelectorAll('tr.vxe-body--row')).filter(
      (tr) => tr.querySelectorAll('td').length > 0
    );
    const statuses = [];
    for (const tr of trs) {
      const cells = Array.from(tr.querySelectorAll('td'));
      let text = '';
      if (idx >= 0 && idx < cells.length) {
        const cell = cells[idx];
        text = (cell.innerText || cell.textContent || '').replace(/\s+/g, ' ').trim();
      } else {
        for (const td of cells) {
          const t = (td.textContent || '').replace(/\s+/g, ' ').trim();
          if (/(在售中|停售)/.test(t)) {
            text = t;
            break;
          }
        }
      }
      statuses.push(text);
    }
    return { ok: true, statuses, rowCount: statuses.length };
  });
}

/**
 * 变体规则（按顺序）：
 * - 无行 → 变体销售状态异常；
 * - 变体行数 > 3（即 4 个及以上）→ 变体数超过3个，不刊登；
 * - 否则逐行：空白或非「在售中」→ 停售文案含「停售」→ 变体停售，其余 → 变体销售状态异常。
 */
function classifyVariantStatuses(statuses) {
  if (!statuses.length) return '变体销售状态异常';
  if (statuses.length > 3) return '变体数超过3个';
  for (const raw of statuses) {
    const s = String(raw ?? '').replace(/\s+/g, ' ').trim();
    if (!s) return '变体销售状态异常';
    if (s === '在售中') continue;
    if (/停售/.test(s)) return '变体停售';
    return '变体销售状态异常';
  }
  return null;
}

/**
 * 变体状态均合格后：对每个变体行点「选择图片」，看当前顶层弹层里 div.box 是否有 li.image_box 且 img 带有效 src。
 * 不用 rowid/col_*（会变）；任一无图则整体不可刊登。
 * 日志：terminal 搜 `[jinpu] 选图检查` 可看每行有效配图数与汇总。
 */
async function checkAllVariantsHaveListingImages(page, meta = {}) {
  const sku = meta.sku ?? '';
  try {
    const bodyScope = listingProductInfoBody(page);
    const rows = bodyScope.locator('tbody tr.vxe-body--row');
    const n = await rows.count();
    if (n === 0) {
      console.error(`[jinpu] 选图检查 SKU=${sku} 无变体行 => FAIL`);
      return false;
    }

    const boxSelector =
      '.ivu-modal-wrap:visible div.box, .el-dialog__wrapper:visible div.box, [role="dialog"]:visible div.box';

    let openedPicker = 0;
    for (let i = 0; i < n; i++) {
      const row = rows.nth(i);
      const btn = row.getByRole('button', { name: '选择图片' });
      try {
        await btn.click({ timeout: 12_000 });
      } catch {
        console.error(`[jinpu] 选图检查 SKU=${sku} 变体行#${i + 1}(index=${i}) 无「选择图片」按钮 => 跳过该行`);
        continue;
      }

      openedPicker += 1;
      await page.waitForTimeout(700);

      const boxLoc = page.locator(boxSelector).first();
      try {
        await boxLoc.waitFor({ state: 'visible', timeout: 12_000 });
      } catch {
        console.error(`[jinpu] 选图检查 SKU=${sku} 变体行#${i + 1} 弹层内 div.box 未出现 => FAIL`);
        await page.keyboard.press('Escape');
        await page.waitForTimeout(400);
        return false;
      }

      const imgCount = await boxLoc.evaluate((box) => {
        return Array.from(box.querySelectorAll('li.image_box img[src]')).filter((img) => {
          const s = (img.getAttribute('src') || '').trim();
          return s.length > 4 && !/^data:image\/svg/i.test(s);
        }).length;
      });

      /** 选图弹窗关闭方式：仅按 Esc（未点「取消/关闭」）；与 iView 顶层 Modal 行为一致 */
      await page.keyboard.press('Escape');
      await page.waitForTimeout(500);

      if (imgCount === 0) {
        console.error(
          `[jinpu] 选图检查 SKU=${sku} 变体行#${i + 1}：该变体没有图片（.box 内无有效配图）=> FAIL`
        );
        return false;
      }
      console.error(
        `[jinpu] 选图检查 SKU=${sku} 变体行#${i + 1}(index=${i})：该变体有图片（弹层 .box 内有效配图 ${imgCount} 张）`
      );
    }

    if (openedPicker === 0 && n > 0) {
      console.error(
        `[jinpu] 选图检查 SKU=${sku} 共${n}行但未点开任何选图窗(openedPicker=0) => FAIL`
      );
      return false;
    }
    console.error(
      `[jinpu] 选图检查 SKU=${sku} 汇总 变体行=${n} 点开选图=${openedPicker}次 => PASS`
    );
    return true;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(`[jinpu] 选图检查 SKU=${sku} 异常 => FAIL (${msg})`);
    try {
      await page.keyboard.press('Escape');
      await page.waitForTimeout(400);
    } catch {
      /* ignore */
    }
    return false;
  }
}


async function getValue(el) {
try {
  const val = await el.inputValue();
  if (val) return val;
} catch (e) {}
  return (await el.textContent()) || (await el.innerText()) || '';
}

/**
 * 列表判定为「可刊登」时：点「精细刊登」→「下一步」→ 选店铺 J6，根据产品信息表里各变体「产品状态」覆盖结论。
 * 同页弹层、无整页跳转；结束后尝试关闭弹层以便继续查下一 SKU。
 */
async function refineListingLabelWithVariantFlow(listingPage, baseLabel, meta = {}) {
  const sku = meta.sku ?? '';
  const stepCtx = { page: listingPage, sku };
  if (baseLabel !== '可刊登') {
    debugLog('skip variant-flow（列表结论非「可刊登」）', { sku, baseLabel });
    return baseLabel;
  }
  debugLog('variant-flow 开始', { sku, baseLabel });
  try {
    if (PAUSE_BEFORE_VARIANT) {
      await runPlaywrightStep(
        'V01_PAUSE',
        '--pause：page.pause()，用 Inspector 单步',
        stepCtx,
        async () => {
          await listingPage.pause();
        }
      );
    }

    const fine = listingPage
      .getByRole('button', { name: '精细刊登' })
      .or(listingPage.getByRole('link', { name: '精细刊登' }));
    await runPlaywrightStep(
      'V03_FINE_CLICK',
      '点击「精细刊登」（button 或 link）',
      stepCtx,
      async () => {
        await fine.first().click({ timeout: 20_000 });
      }
    );

    await runPlaywrightStep(
      'V04_NEXT_CLICK',
      '点击「下一步」',
      stepCtx,
      async () => {
        await listingPage.getByRole('button', { name: '下一步' }).click({ timeout: 20_000 });
      }
    );

    await runPlaywrightStep(
      'V05_POST_NEXT_WAIT',
      '下一步后固定等待 2s（接口/渲染）',
      stepCtx,
      async () => {
        await listingPage.waitForTimeout(2000);
      }
    );

    const shopNarrow = listingPage
      .locator('form')
      .filter({ hasText: '店铺 无匹配数据 J6' })
      .getByPlaceholder('请选择');
    const shopWide = listingPage
      .locator('form')
      .filter({ hasText: /店铺/ })
      .getByPlaceholder('请选择')
      .first();
    await runPlaywrightStep(
      'V07_SHOP_OPEN',
      '点开店铺下拉「请选择」（narrow 或 wide form）',
      stepCtx,
      async () => {
        const narrowCount = await shopNarrow.count();
        debugLog('shop dropdown', { narrowCount });
        if (narrowCount > 0) await shopNarrow.click({ timeout: 15_000 });
        else await shopWide.click({ timeout: 15_000 });
      }
    );

    await runPlaywrightStep(
      'V08_J6_CLICK',
      '在下拉里点击文案为「J6」的项',
      stepCtx,
      async () => {
        await listingPage.getByText('J6', { exact: true }).click({ timeout: 15_000 });
      }
    );

    await runPlaywrightStep(
      'V09_TBODY_WAIT',
      '在「产品信息」并列 div.body 内等待变体行（.tid_172 或 tbody 含 seller_sku:）',
      stepCtx,
      async () => {
        const bodyScope = listingProductInfoBody(listingPage);
        const byTid = bodyScope.locator('.tid_172 .body--wrapper tbody tr.vxe-body--row').first();
        const bySellerSku = bodyScope
          .locator('tbody tr.vxe-body--row')
          .filter({ hasText: /seller_sku:/i })
          .first();
        await byTid.or(bySellerSku).waitFor({ state: 'visible', timeout: 45_000 });
      }
    );

    await runPlaywrightStep(
      'V10_TABLE_STABLE',
      '表格出现后固定再等 3s',
      stepCtx,
      async () => {
        await listingPage.waitForTimeout(3000);
      }
    );

    const SkuNodeId = await runPlaywrightStep(
      'V11_SKU_NODEID',
      '列表页「产品信息」区块抓 SKU 对应的 Node ID',
      stepCtx,
      async () => {
        const productNodeId = listingPage.getByText('Node ID').first();
        const nextElement = productNodeId.locator('xpath=following-sibling::*[1]');
        const NodeIdInput = nextElement.locator('input').first();
        await NodeIdInput.waitFor({ state: 'visible', timeout: 10000 });
        const nodeId = await getValue(NodeIdInput);
        return nodeId;
      }
    );
    console.error(`[jinpu] 抓 SKU Node ID: ${SkuNodeId}`);

    const SkuTitle = await runPlaywrightStep(
      'V12_SKU_TITLE',
      '列表页「产品信息」区块抓 SKU 对应的标题',
      stepCtx,
      async () => {
        const targetElement = listingPage.getByText('产品标题').nth(1);
        const nextElement2 = targetElement.locator('xpath=following-sibling::*[1]');
        const textarea = nextElement2.locator('textarea').first();
        await textarea.waitFor({ state: 'visible', timeout: 10000 });
        const title = await getValue(textarea);
        return title;
      }
    );
    console.error(`[jinpu] 抓 SKU 标题: ${SkuTitle}`);

    await runPlaywrightStep(
      '13_SCRAPE',
      'evaluate：仅在「产品信息」并列 div.body 内抓取变体「产品状态」',
      stepCtx,
      async () => scrapeTid172VariantStatuses(listingPage)
    );

    const scraped = await runPlaywrightStep(
      'V13_SCRAPE',
      'evaluate：仅在「产品信息」并列 div.body 内抓取变体「产品状态」',
      stepCtx,
      async () => scrapeTid172VariantStatuses(listingPage)
    );
    const { ok, statuses, reason } = scraped;
    const statusIssue = !ok ? '变体销售状态异常' : classifyVariantStatuses(statuses);
    let out = statusIssue ?? baseLabel;

    let imagesCheck = 'skipped';
    if (!statusIssue) {
      const imagesOk = await runPlaywrightStep(
        'V13B_VARIANT_IMAGES',
        '逐行点「选择图片」，检查弹层 div.box 内 li.image_box img 是否有有效 src',
        stepCtx,
        async () => checkAllVariantsHaveListingImages(listingPage, { sku })
      );
      imagesCheck = imagesOk ? 'pass' : 'fail';
      if (!imagesOk) out = '变体没有图片';
    }

    console.error(
      `[jinpu] 变体 SKU=${sku} ok=${ok}${reason ? ` reason=${reason}` : ''} statuses=${JSON.stringify(statuses)} statusIssue=${statusIssue ?? '(无)'} imagesCheck=${imagesCheck} => ${out}`
    );
    debugLog('variant-flow 结束', { sku, out, scraped });

    await runPlaywrightStep(
      'V14_CLOSE_DIALOG',
      '关闭刊登向导：getByRole 取消 → getByRole 确定关闭',
      stepCtx,
      async () => {
        await listingPage.getByRole('button', { name: '取消' }).first().click({ timeout: 10_000 });
        await listingPage.getByRole('button', { name: '确定关闭' }).click({ timeout: 15_000 });
        await listingPage.waitForTimeout(500);
      }
    );
    console.error(`[jinpu] variant-flow 全部步骤完成，SKU=${sku} 最终结论: ${out}`);
    return { out, SkuNodeId, SkuTitle };
    // return out;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    const failedStep = err instanceof Error ? err.jinpuFailedStep : undefined;
    if (failedStep) {
      console.error(`[jinpu] variant-flow 结束于步骤 ${failedStep}，截图文件名含 FAIL-${failedStep}`);
    } else {
      console.error(`[jinpu] variant-flow 异常（未标记步骤，可能是 pause/内部 bug） SKU=${sku}: ${msg}`);
      await debugScreenshot(listingPage, `variant-fail-unmarked-SKU-${sku || 'unknown'}`);
    }
    try {
      await clickConfirmCloseIfPresent(listingPage);
      await listingPage.keyboard.press('Escape');
      await clickConfirmCloseIfPresent(listingPage);
      await listingPage.waitForTimeout(500);
    } catch {
      /* ignore */
    }
    return '变体销售状态异常';
  }
}

/** 把每行算好的第二列写回 matrix，再生成 sheet（首行非数字时第二列表头为「判断结果」） */
function writeSecondColumnToSheet(wb, sheetName, matrix, skuRows, labelsByRowIndex) {
  console.error('labelsByRowIndex', labelsByRowIndex);
  const next = matrix.map((row) => [...row]);
  for (const { rowIndex } of skuRows) {
    if (rowIndex >= next.length) continue;
    const label = labelsByRowIndex.get(rowIndex) ?? '';
    while (next[rowIndex].length < 2) next[rowIndex].push('');
    next[rowIndex][1] = label;
  }
  const headerA = String(next[0]?.[0] ?? '').trim();
  if (headerA && !/^\d+$/.test(headerA)) {
    while (next[0].length < 2) next[0].push('');
    if (!String(next[0][1] ?? '').trim()) next[0][1] = '判断结果';
  }
  wb.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(next);
}

/**
 * Windows 下若用 Excel/WPS 打开着目标 xlsx，写入时常报 EBUSY；稍作重试，仍失败则提示先关表格软件。
 */
async function writeXlsxWithRetry(wb, filePath, { attempts = 10, delayMs = 800 } = {}) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      XLSX.writeFile(wb, filePath);
      if (i > 1) console.error(`[jinpu] Excel 已写入成功（第 ${i} 次尝试）`);
      return;
    } catch (e) {
      lastErr = e;
      const code = e && e.code;
      const busy = code === 'EBUSY' || code === 'EPERM' || code === 'EACCES';
      if (busy && i < attempts) {
        console.error(
          `[jinpu] 写入被占用 (${code})，${delayMs}ms 后重试 ${i}/${attempts}：${filePath}`
        );
        await new Promise((r) => setTimeout(r, delayMs));
        continue;
      }
      if (busy) {
        console.error(
          `[jinpu] 仍无法写入「${filePath}」：请先关闭 Excel/WPS 中打开的本文件（或关闭预览/同步占用后再运行）。`
        );
      }
      throw e;
    }
  }
  throw lastErr;
}

/**
 * net::ERR_CONNECTION_RESET 等为网络层断开（防火墙/VPN/代理/线路不稳）。
 * 用 domcontentloaded + 超时 + 多次重试缓解偶发失败。
 */
async function gotoWithRetry(
  page,
  url,
  { attempts = 6, delayMs = 2500, timeout = 90_000 } = {}
) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout });
      if (i > 1) console.error(`[jinpu] 页面已打开（第 ${i} 次尝试）`);
      return;
    } catch (e) {
      lastErr = e;
      const msg = String(e?.message || e);
      const retryable = /ERR_CONNECTION|TIMED_OUT|timeout|Navigation|RESET|REFUSED|ABORTED/i.test(
        msg
      );
      if (retryable && i < attempts) {
        console.error(
          `[jinpu] 打开失败，${Math.round(delayMs / 1000)}s 后重试 ${i}/${attempts}：${msg.split('\n')[0]}`
        );
        await new Promise((r) => setTimeout(r, delayMs));
        continue;
      }
      console.error(
        `[jinpu] 无法打开「${url}」：多为网络中断、防火墙/VPN/公司代理重置连接。请用本机 Chrome 手动访问同一地址验证；仍失败则换网络或联系运维。`
      );
      throw e;
    }
  }
  throw lastErr;
}

/** 在列表页 DOM 中抓取 vxe 表格行，映射为与 LIST_TABLE_COLUMN_KEYS 对齐的对象 */
async function scrapeListingRows(listingPage) {
  return listingPage.evaluate((keys) => {
    const rows = Array.from(document.querySelectorAll('tr.vxe-body--row'));
    return rows.map((row) => {
      const tds = Array.from(row.querySelectorAll('td'));
      const cells = tds.map((td) => {
        const img = td.querySelector('img');
        if (img?.src) return img.src.trim();
        return (td.textContent || '').replace(/\s+/g, ' ').trim();
      });
      // 有的表格第一列是展开/勾选占位，列数会比表头多 1
      let start = 0;
      if (cells.length === keys.length + 1) start = 1;
      const obj = {};
      keys.forEach((key, i) => {
        obj[key] = cells[start + i] ?? '';
      });
      return obj;
    });
  }, LIST_TABLE_COLUMN_KEYS);
}

(async () => {
  console.error(
    `[jinpu] STEP_GAP_MS=${STEP_GAP_MS}（每步成功后停顿；弱网可调大、设 JINPU_STEP_GAP_MS=0 关闭、或 node ... --slow）`
  );
  const excelPath = resolveSkuExcelPath();
  const labelsByRowIndex = new Map();
  const { wb, sheetName, matrix, skuRows } = loadSkuRowsFromWorkbook(excelPath, labelsByRowIndex);

  const browser = await chromium.launch({ headless: false });
  try {
  const context = await browser.newContext();
  if (TRACE) {
    await context.tracing.start({ screenshots: true, snapshots: true, sources: true });
    debugLog('trace 已开始，结束时会写入 trace.zip');
  }
  const page = await context.newPage();

  // --- 登录主站（账号密码请按需替换，勿泄露） ---
  debugLog('打开登录页');
  await gotoWithRetry(page, 'https://saaserp-pos.yibainetwork.com/');
  await page.getByRole('textbox', { name: '请输入手机号或邮箱' }).fill('13011019088');
  await page.getByRole('textbox', { name: '请输入登录密码' }).fill('4545#￥w&AW');
  await page.getByRole('button', { name: '登录' }).click();
  // await page.waitForTimeout(10000);
  await page.waitForSelector('#app', { timeout: 30_000 });

  // 侧栏打开「刊登管理」新标签页
  await page.locator('#app > div > div.ui-content div:nth-child(1)  li:nth-child(6) > div > span').filter({ hasText: 'ERP' }).first().hover();
  const [listingPage] = await Promise.all([
    context.waitForEvent('page', { timeout: 60_000 }),
    page.getByText('刊登管理', { exact: true }).click(),
  ]);
  await listingPage.waitForLoadState('domcontentloaded');
  // await listingPage.waitForTimeout(10000);
  await listingPage.getByText('刊登管理', { exact: true }).waitFor({ state: 'visible', timeout: 30_000 });

  // 在新页进入「产品列表」
  await listingPage.locator('div').filter({ hasText: '刊登管理' }).nth(5).hover();
  await listingPage.getByText('产品列表', { exact: true }).click();
  await stepBreath(listingPage);

  const NETWORK_ERR_THRESHOLD = 5;
  const isNetworkError = (e) => /ERR_CONNECTION|TIMED_OUT|timeout|RESET|REFUSED|ABORTED|net::/i.test(
    e instanceof Error ? e.message : String(e)
  );

  const results = [];
  let networkErrCount = 0;
  try {
    for (const { sku, rowIndex } of skuRows) {
      debugLog('SKU 循环', { sku, rowIndex });
      if (labelsByRowIndex.has(rowIndex)) {
        debugLog(`SKU ${sku} 已有结果，跳过`, { sku, rowIndex, label: labelsByRowIndex.get(rowIndex) });
        continue;
      }

      try {
        await clickConfirmCloseIfPresent(listingPage);
        await stepBreath(listingPage);
        await listingPage.getByRole('textbox', { name: 'SKU' }).fill(sku);
        await stepBreath(listingPage);
        await listingPage.getByRole('button', { name: '查询' }).click();
        debugLog('已点查询，等待 10s');
        // await listingPage.waitForTimeout(10000);
        // await listingPage.locator('tr.vxe-body--row, .vxe-table-empty-text, .no-data, :text("暂无数据")').first().waitFor({ state: 'visible', timeout: 30_000 });
        const firstRowOrEmpty = listingPage.locator('tr.vxe-body--row, .vxe-table-empty-text, .no-data, :text("暂无数据")').first();
        await firstRowOrEmpty.waitFor({ state: 'visible', timeout: 30_000 });
        const noData = await listingPage.locator('.vxe-table-empty-text, .no-data, :text("暂无数据")').count() > 0;
        if (!noData) {
          const firstRowTds  = listingPage.locator('tr.vxe-body--row').first().locator('td');
          await firstRowTds.nth(3).waitFor({ state: 'visible', timeout: 10_000 });
          await firstRowTds.nth(8).waitFor({ state: 'attached', timeout: 10_000 });
          await listingPage.waitForTimeout(800);
        }
        const rows = await scrapeListingRows(listingPage);
        const first = rows[0] ?? {};
        let label = deriveSecondColumnLabel(first);
        debugLog('列表推导结论', { sku, label, listRowSample: first });
        // 调用 refineListingLabelWithVariantFlow 并返回 label, SkuNodeId, SkuTitle
        let SkuNodeId, SkuTitle;

        r = await refineListingLabelWithVariantFlow(listingPage, label, { sku });
        label = r.out;
        SkuNodeId = r.SkuNodeId;
        SkuTitle = r.SkuTitle;
        // ({label, SkuNodeId, SkuTitle} = await refineListingLabelWithVariantFlow(listingPage, label, { sku }));
        labelsByRowIndex.set(rowIndex, label, SkuNodeId, SkuTitle);
        console.error(`****************************************************`);
        console.error(`**---[jinpu] SKU=${sku} 处理完成，结论: ${label}, Node ID: ${SkuNodeId}, 标题: ${SkuTitle}---**`);
        console.error(`****************************************************`);
        results.push({ sku, rows, label, SkuNodeId, SkuTitle });
        networkErrCount = 0;
        if (confirmEnabled) await waitForConfirm(sku, label);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        console.error(`[jinpu] SKU=${sku} 处理失败，跳过: ${msg}`);
        labelsByRowIndex.set(rowIndex, '','','','处理异常');
        if (isNetworkError(err)) {
          networkErrCount += 1;
          if (networkErrCount >= NETWORK_ERR_THRESHOLD) {
            console.error(
              `[jinpu] 连续 ${networkErrCount} 个 SKU 因网络问题失败，停止后续处理。请检查网络/VPN后重新运行。`
            );
            break;
          }
        } else {
          networkErrCount = 0;
        }
        try {
          await clickConfirmCloseIfPresent(listingPage);
          await listingPage.keyboard.press('Escape');
          await clickConfirmCloseIfPresent(listingPage);
          await listingPage.waitForTimeout(500);
        } catch {
          /* ignore cleanup errors */
        }
      }
    }
  } finally {
    writeSecondColumnToSheet(wb, sheetName, matrix, skuRows, labelsByRowIndex);
    await writeXlsxWithRetry(wb, excelPath);
  }

  console.log(JSON.stringify(results, null, 2));
  console.error(`已写回第二列: ${excelPath}`);

  if (TRACE) {
    const tracePath = path.join(__dirname, '..', 'trace.zip');
    await context.tracing.stop({ path: tracePath });
    console.error(`[jinpu] trace 已保存: ${tracePath} （查看: npx playwright show-trace "${tracePath}"）`);
  }
  } finally {
    await browser.close();
  }
})();
