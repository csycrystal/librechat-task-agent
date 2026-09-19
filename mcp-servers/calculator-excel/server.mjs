/**
 * 任务智能Agent助手 —— MCP 服务器：计算器 + Excel 数据可视化
 *
 * 提供 4 个工具（stdio 传输，由 LibreChat 后端自动拉起）：
 *   1. calculate      安全数学表达式求值（词法分析 + 调度场算法，不使用 eval）
 *   2. read_excel     读取 .xlsx：工作表清单、数据预览（Markdown 表格）、数值列统计
 *   3. excel_chart    从 Excel 列生成 SVG 图表（柱状图/折线图/饼图）
 *   4. visualize_data 直接根据给定数据生成 SVG 图表
 *
 * 图表默认输出到 client/dist/assets/charts/，生产模式下可通过
 * /assets/charts/<文件名>.svg 直接在聊天中以内嵌 Markdown 图片渲染。
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
// xlsx 为 CJS 包且 readFile 是运行时挂载的，需取 default 互操作对象
import * as XLSX_NS from 'xlsx';
const XLSX = XLSX_NS.default ?? XLSX_NS;
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// 仓库根目录：mcp-servers/calculator-excel/ 的上上级
const REPO_ROOT = path.resolve(__dirname, '..', '..');
// 图表持久化目录：仓库根 data/charts（不随前端构建清空），后端经 /mcp-charts 静态托管
const DEFAULT_OUTPUT_DIR = path.join(REPO_ROOT, 'data', 'charts');
const PUBLIC_URL_PREFIX = '/mcp-charts';
const OUTPUT_DIR = process.env.MCP_CHART_OUTPUT_DIR || DEFAULT_OUTPUT_DIR;

const log = (...args) => process.stderr.write(`[calculator-excel] ${args.join(' ')}\n`);

/* ================================================================== */
/* 一、安全数学表达式求值（词法分析 + 调度场算法，不使用 eval）           */
/* ================================================================== */

const CONSTANTS = { PI: Math.PI, E: Math.E };
const FUNCTIONS = {
  sqrt: Math.sqrt, abs: Math.abs, round: Math.round, floor: Math.floor,
  ceil: Math.ceil, sin: Math.sin, cos: Math.cos, tan: Math.tan,
  asin: Math.asin, acos: Math.acos, atan: Math.atan,
  log: Math.log10, log2: Math.log2, ln: Math.log, exp: Math.exp,
  sign: Math.sign, cbrt: Math.cbrt,
  min: (...a) => Math.min(...a), max: (...a) => Math.max(...a),
  pow: (a, b) => Math.pow(a, b),
};
// 二元运算符优先级；'u-' 为一元负号，优先级最高、右结合
const PRECEDENCE = { '+': 1, '-': 1, '*': 2, '/': 2, '%': 2, '^': 3, 'u-': 4 };
const RIGHT_ASSOC = new Set(['^', 'u-']);

/** 常见全角符号与乘除号归一化 */
function normalizeExpression(input) {
  return String(input)
    .replace(/×/g, '*').replace(/÷/g, '/')
    .replace(/＋/g, '+').replace(/－/g, '-')
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/，/g, ',');
}

function tokenize(expr) {
  const tokens = [];
  const s = expr;
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < s.length && /[0-9.]/.test(s[j])) j++;
      // 科学计数法（如 1.5e-3）
      if (j < s.length && /[eE]/.test(s[j]) && j + 1 < s.length && /[0-9+\-]/.test(s[j + 1])) {
        j += 2;
        while (j < s.length && /[0-9]/.test(s[j])) j++;
      }
      const numStr = s.slice(i, j);
      const num = Number(numStr);
      if (Number.isNaN(num)) throw new Error(`无效数字: ${numStr}`);
      tokens.push({ type: 'num', value: num });
      i = j;
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < s.length && /[A-Za-z0-9_]/.test(s[j])) j++;
      tokens.push({ type: 'ident', value: s.slice(i, j) });
      i = j;
      continue;
    }
    if ('+-*/%^(),'.includes(ch)) {
      tokens.push({ type: 'op', value: ch });
      i++;
      continue;
    }
    throw new Error(`无法识别的字符 '${ch}'（位置 ${i}）`);
  }
  return tokens;
}

/** 调度场算法：中缀 token 序列 → 逆波兰序列 */
function toRPN(tokens) {
  const output = [];
  const stack = [];
  let prev = null; // 上一个输出/压栈的 token，用于一元负号判定

  for (const tok of tokens) {
    if (tok.type === 'num') {
      output.push(tok);
      prev = tok;
      continue;
    }
    if (tok.type === 'ident') {
      const upper = tok.value.toUpperCase();
      if (upper in CONSTANTS) {
        const numTok = { type: 'num', value: CONSTANTS[upper] };
        output.push(numTok);
        prev = numTok;
        continue;
      }
      const lower = tok.value.toLowerCase();
      if (lower in FUNCTIONS) {
        stack.push({ type: 'func', name: lower, arity: 1 });
        prev = { type: 'op', value: 'func' };
        continue;
      }
      throw new Error(`未知函数或常量: ${tok.value}`);
    }
    // op token
    const v = tok.value;
    if (v === '(') {
      stack.push({ type: '(' });
      prev = tok;
      continue;
    }
    if (v === ')') {
      let found = false;
      while (stack.length) {
        const top = stack[stack.length - 1];
        if (top.type === '(') { stack.pop(); found = true; break; }
        output.push(stack.pop());
      }
      if (!found) throw new Error('括号不匹配：缺少左括号');
      if (stack.length && stack[stack.length - 1].type === 'func') {
        output.push(stack.pop());
      }
      prev = tok;
      continue;
    }
    if (v === ',') {
      while (stack.length && stack[stack.length - 1].type !== '(') {
        output.push(stack.pop());
      }
      if (!stack.length) throw new Error('逗号出现在函数参数之外');
      // 该逗号属于最近一个未闭合函数的额外参数（函数在其 '(' 之下）
      for (let k = stack.length - 1; k >= 0; k--) {
        if (stack[k].type === '(') {
          if (k > 0 && stack[k - 1].type === 'func') stack[k - 1].arity++;
          break;
        }
      }
      prev = tok;
      continue;
    }
    // 普通二元运算符 / 一元负号
    let op = v;
    const unary = (v === '-' || v === '+') &&
      !(prev && (prev.type === 'num' || (prev.type === 'op' && (prev.value === ')' || prev.value === 'func'))));
    if (v === '+' && unary) { prev = tok; continue; } // 一元正号忽略
    if (unary) op = 'u-';
    while (stack.length) {
      const top = stack[stack.length - 1];
      if (top.type !== 'op') break;
      const tPrec = PRECEDENCE[top.value];
      const oPrec = PRECEDENCE[op];
      if (tPrec > oPrec || (tPrec === oPrec && !RIGHT_ASSOC.has(op))) {
        output.push(stack.pop());
      } else {
        break;
      }
    }
    stack.push({ type: 'op', value: op });
    prev = tok;
  }
  while (stack.length) {
    const top = stack.pop();
    if (top.type === '(') throw new Error('括号不匹配：缺少右括号');
    output.push(top);
  }
  return output;
}

function evalRPN(rpn) {
  const st = [];
  for (const tok of rpn) {
    if (tok.type === 'num') { st.push(tok.value); continue; }
    if (tok.type === 'func') {
      if (st.length < tok.arity) throw new Error(`函数 ${tok.name} 参数不足`);
      const args = st.splice(st.length - tok.arity);
      st.push(FUNCTIONS[tok.name](...args));
      continue;
    }
    if (tok.type === 'op') {
      if (tok.value === 'u-') {
        if (!st.length) throw new Error('表达式错误');
        st.push(-st.pop());
        continue;
      }
      if (st.length < 2) throw new Error('表达式错误：运算符缺少操作数');
      const b = st.pop();
      const a = st.pop();
      switch (tok.value) {
        case '+': st.push(a + b); break;
        case '-': st.push(a - b); break;
        case '*': st.push(a * b); break;
        case '/':
          if (b === 0) throw new Error('除数不能为零');
          st.push(a / b); break;
        case '%': st.push(a % b); break;
        case '^': st.push(Math.pow(a, b)); break;
        default: throw new Error(`未知运算符: ${tok.value}`);
      }
    }
  }
  if (st.length !== 1) throw new Error('表达式错误：操作数多余');
  return st[0];
}

function evaluateExpression(rawExpression) {
  const normalized = normalizeExpression(rawExpression);
  if (!normalized.trim()) throw new Error('表达式为空');
  return evalRPN(toRPN(tokenize(normalized)));
}

function formatNumber(n) {
  if (!Number.isFinite(n)) return String(n);
  if (Number.isInteger(n)) return String(n);
  return String(Math.round(n * 1e10) / 1e10);
}

/* ================================================================== */
/* 一·补、单位换算（长度/重量/面积/时间/容积 + 温度）                     */
/* ================================================================== */

// 各类别以基础单位为 1 的换算倍率
const UNIT_FACTORS = {
  length: {
    base: 'm',
    mm: 0.001, cm: 0.01, m: 1, km: 1000,
    inch: 0.0254, foot: 0.3048, yard: 0.9144, mile: 1609.344,
    li: 500, chi: 1 / 3, cun: 1 / 30,
  },
  weight: {
    base: 'g',
    mg: 0.001, g: 1, kg: 1000, t: 1e6,
    lb: 453.59237, oz: 28.349523125, jin: 500, liang: 50,
  },
  area: {
    base: 'm2',
    cm2: 1e-4, m2: 1, km2: 1e6, ha: 1e4, mu: 2000 / 3,
    ft2: 0.09290304, acre: 4046.8564224,
  },
  time: {
    base: 's',
    ms: 0.001, s: 1, min: 60, h: 3600, day: 86400, week: 604800,
  },
  volume: {
    base: 'l',
    ml: 0.001, l: 1, m3: 1000, gal: 3.785411784, qt: 0.946352946,
  },
};
// 中文/别名 → 标准键
const UNIT_ALIASES = {
  '毫米': 'mm', '厘米': 'cm', '公分': 'cm', '米': 'm', '千米': 'km', '公里': 'km',
  '英寸': 'inch', '英尺': 'foot', '码': 'yard', '英里': 'mile', '里': 'li', '尺': 'chi', '寸': 'cun',
  '毫克': 'mg', '克': 'g', '千克': 'kg', '公斤': 'kg', '吨': 't',
  '磅': 'lb', '盎司': 'oz', '斤': 'jin', '两': 'liang',
  '平方米': 'm2', '平米': 'm2', '平方公里': 'km2', '平方千米': 'km2',
  '公顷': 'ha', '亩': 'mu', '平方英尺': 'ft2', '英亩': 'acre', '平方厘米': 'cm2',
  '毫秒': 'ms', '秒': 's', '分钟': 'min', '分': 'min', '小时': 'h', '时': 'h',
  '天': 'day', '日': 'day', '周': 'week',
  '毫升': 'ml', '升': 'l', '立方米': 'm3', '加仑': 'gal', '夸脱': 'qt',
  '摄氏度': 'c', '摄氏': 'c', '度': 'c', '℃': 'c', 'c': 'c',
  '华氏度': 'f', '华氏': 'f', '℉': 'f', 'f': 'f',
  '开尔文': 'k', '开氏度': 'k', 'k': 'k',
};
const UNIT_LABELS = {
  mm: '毫米', cm: '厘米', m: '米', km: '千米/公里', inch: '英寸', foot: '英尺',
  yard: '码', mile: '英里', li: '市里', chi: '市尺', cun: '市寸',
  mg: '毫克', g: '克', kg: '千克/公斤', t: '吨', lb: '磅', oz: '盎司', jin: '市斤', liang: '两',
  cm2: '平方厘米', m2: '平方米', km2: '平方公里', ha: '公顷', mu: '亩',
  ft2: '平方英尺', acre: '英亩',
  ms: '毫秒', s: '秒', min: '分钟', h: '小时', day: '天', week: '周',
  ml: '毫升', l: '升', m3: '立方米', gal: '加仑', qt: '夸脱',
  c: '摄氏度(°C)', f: '华氏度(°F)', k: '开尔文(K)',
};

function normalizeUnit(raw) {
  const key = String(raw).trim().toLowerCase();
  return UNIT_ALIASES[key] ?? key;
}

function findUnitCategory(unit) {
  if (unit === 'c' || unit === 'f' || unit === 'k') return 'temperature';
  for (const [category, factors] of Object.entries(UNIT_FACTORS)) {
    if (unit in factors) return category;
  }
  return null;
}

function celsiusFrom(value, unit) {
  if (unit === 'c') return value;
  if (unit === 'f') return ((value - 32) * 5) / 9;
  return value - 273.15; // k
}
function celsiusTo(valueC, unit) {
  if (unit === 'c') return valueC;
  if (unit === 'f') return (valueC * 9) / 5 + 32;
  return valueC + 273.15;
}

function convertUnit(value, fromRaw, toRaw) {
  const from = normalizeUnit(fromRaw);
  const to = normalizeUnit(toRaw);
  const category = findUnitCategory(from);
  if (!category) throw new Error(`不支持的源单位: ${fromRaw}`);
  if (findUnitCategory(to) !== category) {
    throw new Error(`单位不匹配：${UNIT_LABELS[from] || from} 与 ${UNIT_LABELS[to] || to} 不属于同一类别`);
  }
  let result;
  if (category === 'temperature') {
    result = celsiusTo(celsiusFrom(value, from), to);
  } else {
    const factors = UNIT_FACTORS[category];
    result = (value * factors[from]) / factors[to];
  }
  return {
    result: Math.round(result * 1e8) / 1e8,
    category,
    fromLabel: UNIT_LABELS[from] || from,
    toLabel: UNIT_LABELS[to] || to,
  };
}

/* ================================================================== */
/* 二、SVG 图表生成（柱状图 / 折线图 / 饼图）                            */
/* ================================================================== */

const PALETTE = ['#4e79a7', '#f28e2b', '#e15759', '#76b7b2', '#59a14f', '#edc949', '#b07aa1', '#ff9da7'];
const FONT = "font-family=\"Microsoft YaHei, PingFang SC, sans-serif\"";

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function svgHeader(width, height) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" ${FONT}>` +
    `<rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff"/>`;
}

function renderTitle(title, width, y = 34) {
  if (!title) return '';
  return `<text x="${width / 2}" y="${y}" text-anchor="middle" font-size="20" fill="#1f2937" font-weight="bold">${escapeXml(title)}</text>`;
}

function renderLegend(seriesList, xRight, y) {
  return seriesList.map((s, idx) => {
    const name = escapeXml(s.name || `系列${idx + 1}`);
    const lx = xRight - 180;
    const ly = y + idx * 22;
    return `<rect x="${lx}" y="${ly - 11}" width="14" height="14" fill="${PALETTE[idx % PALETTE.length]}" rx="2"/>` +
      `<text x="${lx + 20}" y="${ly + 1}" font-size="13" fill="#374151">${name}</text>`;
  }).join('');
}

/** 柱状图：labels 为 x 轴类目，seriesList 为 {name, values[]} */
function renderBarChart(title, labels, seriesList) {
  const groupW = 90 * seriesList.length;
  const width = Math.max(720, 140 + labels.length * groupW);
  const height = 460;
  const m = { top: 70, right: 40, bottom: 80, left: 70 };
  const plotW = width - m.left - m.right;
  const plotH = height - m.top - m.bottom;

  const all = seriesList.flatMap((s) => s.values);
  const maxV = Math.max(0, ...all);
  const minV = Math.min(0, ...all);
  const range = maxV - minV || 1;
  const yOf = (v) => m.top + plotH * (1 - (v - minV) / range);

  let out = svgHeader(width, height) + renderTitle(title, width);
  // 水平网格线与 y 轴刻度
  for (let g = 0; g <= 5; g++) {
    const v = minV + (range * g) / 5;
    const y = yOf(v);
    out += `<line x1="${m.left}" y1="${y}" x2="${width - m.right}" y2="${y}" stroke="#e5e7eb" stroke-width="1"/>`;
    out += `<text x="${m.left - 8}" y="${y + 4}" text-anchor="end" font-size="12" fill="#6b7280">${formatNumber(Math.round(v * 100) / 100)}</text>`;
  }
  const bandW = plotW / Math.max(labels.length, 1);
  const barW = (bandW * 0.7) / seriesList.length;
  labels.forEach((label, li) => {
    seriesList.forEach((s, si) => {
      const v = Number(s.values[li]) || 0;
      const x = m.left + li * bandW + bandW * 0.15 + si * barW;
      const y0 = yOf(0);
      const y1 = yOf(v);
      const top = Math.min(y0, y1);
      const h = Math.max(Math.abs(y1 - y0), 1);
      out += `<rect x="${x}" y="${top}" width="${barW * 0.85}" height="${h}" fill="${PALETTE[si % PALETTE.length]}" rx="2"/>`;
      out += `<text x="${x + (barW * 0.85) / 2}" y="${top - 5}" text-anchor="middle" font-size="11" fill="#374151">${formatNumber(v)}</text>`;
    });
    const short = String(label).length > 8 ? String(label).slice(0, 7) + '…' : String(label);
    out += `<text x="${m.left + li * bandW + bandW / 2}" y="${height - m.bottom + 22}" text-anchor="middle" font-size="12" fill="#374151">${escapeXml(short)}</text>`;
  });
  out += renderLegend(seriesList, width - m.right, m.top - 46);
  out += '</svg>';
  return { svg: out, width, height };
}

/** 折线图 */
function renderLineChart(title, labels, seriesList) {
  const width = Math.max(720, 140 + labels.length * 70);
  const height = 460;
  const m = { top: 70, right: 40, bottom: 80, left: 70 };
  const plotW = width - m.left - m.right;
  const plotH = height - m.top - m.bottom;

  const all = seriesList.flatMap((s) => s.values);
  const maxV = Math.max(0, ...all);
  const minV = Math.min(0, ...all);
  const range = maxV - minV || 1;
  const yOf = (v) => m.top + plotH * (1 - (v - minV) / range);
  const xOf = (i) => labels.length === 1 ? m.left + plotW / 2 : m.left + (plotW * i) / (labels.length - 1);

  let out = svgHeader(width, height) + renderTitle(title, width);
  for (let g = 0; g <= 5; g++) {
    const v = minV + (range * g) / 5;
    const y = yOf(v);
    out += `<line x1="${m.left}" y1="${y}" x2="${width - m.right}" y2="${y}" stroke="#e5e7eb" stroke-width="1"/>`;
    out += `<text x="${m.left - 8}" y="${y + 4}" text-anchor="end" font-size="12" fill="#6b7280">${formatNumber(Math.round(v * 100) / 100)}</text>`;
  }
  seriesList.forEach((s, si) => {
    const pts = s.values.map((v, i) => `${xOf(i)},${yOf(Number(v) || 0)}`).join(' ');
    out += `<polyline points="${pts}" fill="none" stroke="${PALETTE[si % PALETTE.length]}" stroke-width="2.5" stroke-linejoin="round"/>`;
    s.values.forEach((v, i) => {
      out += `<circle cx="${xOf(i)}" cy="${yOf(Number(v) || 0)}" r="4" fill="${PALETTE[si % PALETTE.length]}"/>`;
      out += `<text x="${xOf(i)}" y="${yOf(Number(v) || 0) - 10}" text-anchor="middle" font-size="11" fill="#374151">${formatNumber(Number(v) || 0)}</text>`;
    });
  });
  labels.forEach((label, i) => {
    const short = String(label).length > 8 ? String(label).slice(0, 7) + '…' : String(label);
    out += `<text x="${xOf(i)}" y="${height - m.bottom + 22}" text-anchor="middle" font-size="12" fill="#374151">${escapeXml(short)}</text>`;
  });
  out += renderLegend(seriesList, width - m.right, m.top - 46);
  out += '</svg>';
  return { svg: out, width, height };
}

/** 饼图（取第一条系列） */
function renderPieChart(title, labels, values) {
  const width = 760;
  const height = 480;
  const cx = 250;
  const cy = 260;
  const r = 160;
  const total = values.reduce((a, b) => a + (Number(b) || 0), 0) || 1;

  let angle = -Math.PI / 2;
  let out = svgHeader(width, height) + renderTitle(title, width);
  values.forEach((v, i) => {
    const val = Number(v) || 0;
    const frac = val / total;
    const end = angle + frac * Math.PI * 2;
    const x0 = cx + r * Math.cos(angle);
    const y0 = cy + r * Math.sin(angle);
    const x1 = cx + r * Math.cos(end);
    const y1 = cy + r * Math.sin(end);
    const large = frac > 0.5 ? 1 : 0;
    if (frac >= 1) {
      out += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${PALETTE[0]}"/>`;
    } else if (frac > 0) {
      out += `<path d="M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z" ` +
        `fill="${PALETTE[i % PALETTE.length]}" stroke="#ffffff" stroke-width="2"/>`;
    }
    // 扇区外侧百分比标签
    if (frac >= 0.03) {
      const mid = (angle + end) / 2;
      const lx = cx + (r + 32) * Math.cos(mid);
      const ly = cy + (r + 32) * Math.sin(mid);
      out += `<text x="${lx}" y="${ly}" text-anchor="middle" font-size="12" fill="#374151">${(frac * 100).toFixed(1)}%</text>`;
    }
    angle = end;
  });
  // 图例
  labels.forEach((label, i) => {
    const val = Number(values[i]) || 0;
    const pct = ((val / total) * 100).toFixed(1);
    const lx = 470;
    const ly = 120 + i * 26;
    const short = String(label).length > 12 ? String(label).slice(0, 11) + '…' : String(label);
    out += `<rect x="${lx}" y="${ly - 12}" width="14" height="14" fill="${PALETTE[i % PALETTE.length]}" rx="2"/>` +
      `<text x="${lx + 20}" y="${ly}" font-size="13" fill="#374151">${escapeXml(short)}：${formatNumber(val)}（${pct}%）</text>`;
  });
  out += '</svg>';
  return { svg: out, width, height };
}

function generateChart(chartType, title, labels, seriesList) {
  if (chartType === 'pie') {
    return renderPieChart(title || '数据占比', labels, seriesList[0].values);
  }
  if (chartType === 'line') {
    return renderLineChart(title || '数据趋势', labels, seriesList);
  }
  if (chartType === 'area') {
    return renderAreaChart(title || '数据趋势', labels, seriesList);
  }
  if (chartType === 'scatter') {
    return renderScatterChart(title || '数据分布', labels, seriesList);
  }
  return renderBarChart(title || '数据对比', labels, seriesList);
}

/** 面积图：折线 + 渐变填充区域 */
function renderAreaChart(title, labels, seriesList) {
  const width = Math.max(720, 140 + labels.length * 70);
  const height = 460;
  const m = { top: 70, right: 40, bottom: 80, left: 70 };
  const plotW = width - m.left - m.right;
  const plotH = height - m.top - m.bottom;

  const all = seriesList.flatMap((s) => s.values);
  const maxV = Math.max(0, ...all);
  const minV = Math.min(0, ...all);
  const range = maxV - minV || 1;
  const yOf = (v) => m.top + plotH * (1 - (v - minV) / range);
  const xOf = (i) => labels.length === 1 ? m.left + plotW / 2 : m.left + (plotW * i) / (labels.length - 1);
  const baseline = yOf(0);

  let out = svgHeader(width, height) + renderTitle(title, width);
  for (let g = 0; g <= 5; g++) {
    const v = minV + (range * g) / 5;
    const y = yOf(v);
    out += `<line x1="${m.left}" y1="${y}" x2="${width - m.right}" y2="${y}" stroke="#e5e7eb" stroke-width="1"/>`;
    out += `<text x="${m.left - 8}" y="${y + 4}" text-anchor="end" font-size="12" fill="#6b7280">${formatNumber(Math.round(v * 100) / 100)}</text>`;
  }
  seriesList.forEach((s, si) => {
    const color = PALETTE[si % PALETTE.length];
    const gid = `areaGrad${si}`;
    out += `<defs><linearGradient id="${gid}" x1="0" y1="0" x2="0" y2="1">` +
      `<stop offset="0%" stop-color="${color}" stop-opacity="0.45"/>` +
      `<stop offset="100%" stop-color="${color}" stop-opacity="0.05"/>` +
      `</linearGradient></defs>`;
    const linePts = s.values.map((v, i) => `${xOf(i)},${yOf(Number(v) || 0)}`).join(' ');
    const firstX = xOf(0);
    const lastX = xOf(s.values.length - 1);
    out += `<polygon points="${firstX},${baseline} ${linePts} ${lastX},${baseline}" fill="url(#${gid})"/>`;
    out += `<polyline points="${linePts}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linejoin="round"/>`;
    s.values.forEach((v, i) => {
      out += `<circle cx="${xOf(i)}" cy="${yOf(Number(v) || 0)}" r="3.5" fill="${color}"/>`;
    });
  });
  labels.forEach((label, i) => {
    const short = String(label).length > 8 ? String(label).slice(0, 7) + '…' : String(label);
    out += `<text x="${xOf(i)}" y="${height - m.bottom + 22}" text-anchor="middle" font-size="12" fill="#374151">${escapeXml(short)}</text>`;
  });
  out += renderLegend(seriesList, width - m.right, m.top - 46);
  out += '</svg>';
  return { svg: out, width, height };
}

/** 散点图：labels 为 x 数值，seriesList 各系列为 y 数值 */
function renderScatterChart(title, labels, seriesList) {
  const width = 760;
  const height = 460;
  const m = { top: 70, right: 40, bottom: 70, left: 70 };
  const plotW = width - m.left - m.right;
  const plotH = height - m.top - m.bottom;

  const xs = labels.map(toNumber).filter((n) => !Number.isNaN(n));
  const ys = seriesList.flatMap((s) => s.values);
  if (!xs.length) throw new Error('散点图需要 x 轴为数值');
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(0, ...ys);
  const maxY = Math.max(0, ...ys);
  const rangeX = maxX - minX || 1;
  const rangeY = maxY - minY || 1;
  const xOf = (v) => m.left + plotW * ((v - minX) / rangeX);
  const yOf = (v) => m.top + plotH * (1 - (v - minY) / rangeY);

  let out = svgHeader(width, height) + renderTitle(title, width);
  for (let g = 0; g <= 5; g++) {
    const yv = minY + (rangeY * g) / 5;
    const y = yOf(yv);
    out += `<line x1="${m.left}" y1="${y}" x2="${width - m.right}" y2="${y}" stroke="#e5e7eb"/>`;
    out += `<text x="${m.left - 8}" y="${y + 4}" text-anchor="end" font-size="12" fill="#6b7280">${formatNumber(Math.round(yv * 100) / 100)}</text>`;
    const xv = minX + (rangeX * g) / 5;
    const x = xOf(xv);
    out += `<text x="${x}" y="${height - m.bottom + 22}" text-anchor="middle" font-size="12" fill="#6b7280">${formatNumber(Math.round(xv * 100) / 100)}</text>`;
  }
  out += `<line x1="${m.left}" y1="${yOf(0)}" x2="${width - m.right}" y2="${yOf(0)}" stroke="#9ca3af" stroke-width="1.2"/>`;
  out += `<line x1="${xOf(0)}" y1="${m.top}" x2="${xOf(0)}" y2="${height - m.bottom}" stroke="#9ca3af" stroke-width="1.2"/>`;
  seriesList.forEach((s, si) => {
    const color = PALETTE[si % PALETTE.length];
    s.values.forEach((v, i) => {
      const xv = toNumber(labels[i]);
      if (Number.isNaN(xv)) return;
      out += `<circle cx="${xOf(xv)}" cy="${yOf(Number(v) || 0)}" r="6" fill="${color}" fill-opacity="0.75" stroke="#ffffff" stroke-width="1.5"/>`;
    });
  });
  out += renderLegend(seriesList, width - m.right, m.top - 46);
  out += '</svg>';
  return { svg: out, width, height };
}

/** 保存 SVG；默认目录在仓库根 data/charts 下，可经 /mcp-charts/ 访问 */
function saveChartSvg(svg, prefix) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const filename = `${prefix}-${Date.now()}-${crypto.randomBytes(3).toString('hex')}.svg`;
  const absPath = path.join(OUTPUT_DIR, filename);
  fs.writeFileSync(absPath, svg, 'utf8');
  const webPath = OUTPUT_DIR === DEFAULT_OUTPUT_DIR ? `${PUBLIC_URL_PREFIX}/${filename}` : null;
  return { absPath, webPath };
}

/* ================================================================== */
/* 三、Excel 读取与图表数据提取                                          */
/* ================================================================== */

function toNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
  if (typeof v !== 'string') return NaN;
  const s = v.trim();
  if (!s) return NaN;
  const n = Number(s.replace(/,/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function isNumericLike(v) {
  return !Number.isNaN(toNumber(v));
}

/** 读取 xlsx/xls/csv 指定工作表，返回 {sheetName, headers, rows}；相对路径基于仓库根目录解析 */
function loadSheet(filePath, sheetName) {
  const resolved = path.isAbsolute(filePath) ? filePath : path.join(REPO_ROOT, filePath);
  if (!fs.existsSync(resolved)) {
    throw new Error(`文件不存在: ${resolved}`);
  }
  const ext = path.extname(resolved).toLowerCase();
  if (!['.xlsx', '.xls', '.csv'].includes(ext)) {
    throw new Error(`仅支持 .xlsx/.xls/.csv 文件，收到: ${ext || '无扩展名'}`);
  }
  // CSV：显式按 UTF-8 解析（无 BOM 时 SheetJS 会回退到系统代码页导致中文乱码）
  const workbook = ext === '.csv'
    ? XLSX.read(fs.readFileSync(resolved).toString('utf8'), { type: 'string', codepage: 65001 })
    : XLSX.readFile(resolved);
  if (!workbook.SheetNames.length) throw new Error('该 Excel 文件没有任何工作表');
  let sheet = sheetName;
  if (sheet) {
    const hit = workbook.SheetNames.find(
      (n) => n === sheet || n.toLowerCase() === String(sheet).toLowerCase(),
    );
    if (!hit) {
      throw new Error(`未找到工作表 "${sheet}"。可用工作表: ${workbook.SheetNames.join(', ')}`);
    }
    sheet = hit;
  } else {
    sheet = workbook.SheetNames[0];
  }
  const ws = workbook.Sheets[sheet];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  // 去掉末尾连续空行
  while (raw.length && raw[raw.length - 1].every((c) => c === '' || c == null)) raw.pop();
  if (!raw.length) throw new Error(`工作表 "${sheet}" 为空`);
  const headers = raw[0].map((h, i) => String(h).trim() || `列${i + 1}`);
  const rows = raw.slice(1).filter((r) => r.some((c) => c !== '' && c != null));
  return { sheetName: sheet, headers, rows };
}

/** 每个数值列的统计信息 */
function columnStats(headers, rows) {
  const stats = [];
  headers.forEach((h, ci) => {
    const nums = rows.map((r) => toNumber(r[ci])).filter((n) => !Number.isNaN(n));
    if (!nums.length) return;
    const sum = nums.reduce((a, b) => a + b, 0);
    stats.push({
      column: h,
      count: nums.length,
      sum: Math.round(sum * 100) / 100,
      mean: Math.round((sum / nums.length) * 100) / 100,
      min: Math.min(...nums),
      max: Math.max(...nums),
    });
  });
  return stats;
}

function markdownTable(headers, rows) {
  const head = `| ${headers.map(escapeXml).join(' | ')} |`;
  const sep = `| ${headers.map(() => '---').join(' | ')} |`;
  const body = rows
    .map((r) => `| ${headers.map((_, ci) => escapeXml(r[ci] ?? '')).join(' | ')} |`)
    .join('\n');
  return [head, sep, body].join('\n');
}

/** 数值占优（>=50% 且至少 1 个数值）的列索引 */
function numericColumns(headers, rows) {
  return headers
    .map((_, ci) => {
      const vals = rows.map((r) => r[ci]).filter((v) => v !== '' && v != null);
      if (!vals.length) return -1;
      const nums = vals.filter((v) => isNumericLike(v));
      return nums.length / vals.length >= 0.5 && nums.length > 0 ? ci : -1;
    })
    .filter((ci) => ci >= 0);
}

function findColumn(headers, col) {
  if (col == null) return -1;
  if (typeof col === 'number' && Number.isInteger(col)) {
    if (col < 0 || col >= headers.length) throw new Error(`列索引 ${col} 超出范围（共 ${headers.length} 列）`);
    return col;
  }
  const idx = headers.findIndex((h) => h === String(col).trim());
  if (idx === -1) throw new Error(`未找到列 "${col}"。可用列: ${headers.join(', ')}`);
  return idx;
}

/** 从表格提取图表数据：labels + seriesList。scatter 时 x 轴也使用数值列 */
function extractChartSeries(headers, rows, xCol, yCol, chartType = 'bar', maxRows = 50) {
  const dataRows = rows.slice(0, maxRows);
  const numeric = numericColumns(headers, rows);

  let xIdx = findColumn(headers, xCol);
  if (xIdx === -1) {
    if (chartType === 'scatter') {
      // 散点图 x 轴取第一个数值列
      xIdx = numeric[0] ?? -1;
    } else {
      const firstNonNumeric = headers.findIndex((_, ci) => !numeric.includes(ci) &&
        dataRows.some((r) => r[ci] !== '' && r[ci] != null));
      xIdx = firstNonNumeric;
    }
  }
  const labels = dataRows.map((r, i) => (xIdx >= 0 && r[xIdx] !== '' && r[xIdx] != null ? String(r[xIdx]) : `第${i + 1}行`));

  let seriesList;
  if (yCol != null) {
    const yIdx = findColumn(headers, yCol);
    seriesList = [{ name: headers[yIdx], values: dataRows.map((r) => toNumber(r[yIdx]) || 0) }];
  } else {
    const cols = numeric.filter((ci) => ci !== xIdx).slice(0, 3);
    if (!cols.length) throw new Error('没有找到可绘图的数值列，请通过 y_col 参数指定列');
    seriesList = cols.map((ci) => ({
      name: headers[ci],
      values: dataRows.map((r) => toNumber(r[ci]) || 0),
    }));
  }
  return { labels, seriesList, truncated: rows.length > maxRows };
}

/* ================================================================== */
/* 四、MCP 服务器与工具注册                                              */
/* ================================================================== */

const server = new McpServer({ name: 'calculator-excel', version: '1.0.0' });

server.registerTool(
  'calculate',
  {
    title: '计算器',
    description:
      '精确计算数学表达式。支持 + - * / % ^ 括号，函数 sqrt abs round floor ceil sin cos tan log ln exp min max pow，' +
      '常量 PI E。示例: (3+4)*5、sqrt(2)^10、max(12,36,48)、(12000*0.13)/1.13。' +
      '任何需要数值计算（尤其是多位数乘除、百分比、统计汇总）时都应使用本工具而不是心算。',
    inputSchema: {
      expression: z.string().describe('数学表达式，例如 (3+4)*5 或 1250*0.13'),
    },
  },
  async ({ expression }) => {
    try {
      const result = evaluateExpression(expression);
      return { content: [{ type: 'text', text: `计算结果: ${expression} = ${formatNumber(result)}` }] };
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: `计算失败: ${err.message}` }] };
    }
  },
);

server.registerTool(
  'convert_unit',
  {
    title: '单位换算',
    description:
      '精确的单位换算，支持长度（mm/cm/m/km/英寸/英尺/英里/里/尺）、重量（mg/g/kg/t/磅/盎司/斤/两）、' +
      '面积（平方米/平方公里/公顷/亩/平方英尺/英亩）、时间（毫秒/秒/分/时/天/周）、' +
      '容积（毫升/升/立方米/加仑）和温度（摄氏度C/华氏度F/开尔文K，可直接写中文"摄氏度""华氏度"）。' +
      '同类单位之间换算，如 100 英里→公里、32 华氏度→摄氏度、3 斤→克。',
    inputSchema: {
      value: z.number().describe('待换算的数值'),
      from: z.string().describe('源单位，如 mile、英里、km、摄氏度、f'),
      to: z.string().describe('目标单位，如 km、公里、m、摄氏度、c'),
    },
  },
  async ({ value, from, to }) => {
    try {
      const r = convertUnit(value, from, to);
      return {
        content: [{
          type: 'text',
          text: `换算结果（${r.category}）: ${formatNumber(value)} ${r.fromLabel} = ${formatNumber(r.result)} ${r.toLabel}`,
        }],
      };
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: `换算失败: ${err.message}` }] };
    }
  },
);

server.registerTool(
  'read_excel',
  {
    title: '读取 Excel',
    description:
      '读取 .xlsx/.xls/.csv 文件，返回工作表清单、前 N 行数据（Markdown 表格）与数值列统计（求和/均值/最大/最小）。' +
      '相对路径基于项目根目录解析。当用户提供 Excel/CSV 文件路径或需要分析表格数据时使用。',
    inputSchema: {
      file_path: z.string().describe('表格文件路径（.xlsx/.xls/.csv）'),
      sheet: z.string().optional().describe('工作表名称，默认第一个'),
      max_rows: z.number().int().min(1).max(100).default(20).describe('预览的数据行数，默认 20'),
    },
  },
  async ({ file_path, sheet, max_rows }) => {
    try {
      const { sheetName, headers, rows } = loadSheet(file_path, sheet);
      const preview = rows.slice(0, max_rows);
      const parts = [
        `文件: ${file_path}`,
        `工作表: ${sheetName}`,
        `数据规模: ${rows.length} 行 × ${headers.length} 列`,
        '',
        `## 数据预览（前 ${preview.length} 行）`,
        markdownTable(headers, preview),
      ];
      if (rows.length > preview.length) {
        parts.push('', `（其余 ${rows.length - preview.length} 行已省略）`);
      }
      const stats = columnStats(headers, rows);
      if (stats.length) {
        parts.push('', '## 数值列统计', markdownTable(['column', 'count', 'sum', 'mean', 'min', 'max'],
          stats.map((s) => [s.column, s.count, s.sum, s.mean, s.min, s.max])));
      }
      return { content: [{ type: 'text', text: parts.join('\n') }] };
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: `读取失败: ${err.message}` }] };
    }
  },
);

server.registerTool(
  'excel_chart',
  {
    title: 'Excel 图表',
    description:
      '从 Excel 表格数据生成可视化图表（bar 柱状图 / line 折线图 / pie 饼图），保存为 SVG 文件并返回访问链接。' +
      '未指定列时自动选择数值列。生成后应在回复中内嵌返回的 Markdown 图片链接。',
    inputSchema: {
      file_path: z.string().describe('Excel 文件路径（.xlsx）'),
      sheet: z.string().optional().describe('工作表名称，默认第一个'),
      x_col: z.union([z.string(), z.number()]).optional().describe('x 轴类目列（列名或索引），默认自动选择'),
      y_col: z.union([z.string(), z.number()]).optional().describe('数值列（列名或索引），默认自动选择数值列'),
      chart_type: z.enum(['bar', 'line', 'pie', 'area', 'scatter']).default('bar')
        .describe('图表类型：bar 柱状图/line 折线图/pie 饼图/area 面积图/scatter 散点图（散点图需两列数值）'),
      title: z.string().optional().describe('图表标题'),
    },
  },
  async ({ file_path, sheet, x_col, y_col, chart_type, title }) => {
    try {
      const { headers, rows } = loadSheet(file_path, sheet);
      const { labels, seriesList, truncated } = extractChartSeries(headers, rows, x_col, y_col, chart_type);
      const { svg } = generateChart(chart_type, title, labels, seriesList);
      const { absPath, webPath } = saveChartSvg(svg, `excel-${chart_type}`);
      const lines = [
        `图表已生成: ${absPath}`,
        webPath ? `网页访问地址: ${webPath}` : null,
        webPath ? `请在回复中原样包含以下 Markdown 图片以展示图表:\n![${title || '图表'}](${webPath})` : null,
        truncated ? '注意: 数据行数较多，图表仅使用前 50 行。' : null,
      ].filter(Boolean);
      return { content: [{ type: 'text', text: lines.join('\n') }] };
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: `生成图表失败: ${err.message}` }] };
    }
  },
);

server.registerTool(
  'visualize_data',
  {
    title: '数据可视化',
    description:
      '直接根据给定数据生成图表（bar 柱状图 / line 折线图 / pie 饼图），保存为 SVG 并返回访问链接。' +
      '适合对计算结果或汇总数据快速可视化。生成后应在回复中内嵌返回的 Markdown 图片链接。',
    inputSchema: {
      labels: z.array(z.union([z.string(), z.number()])).describe('x 轴标签数组，例如 ["一月","二月","三月"]'),
      values: z.array(z.number()).optional().describe('单系列数值数组（与 labels 等长）；多系列时用 series'),
      series: z.array(z.object({
        name: z.string(),
        values: z.array(z.number()),
      })).optional().describe('多系列数据，每个系列包含 name 与 values'),
      chart_type: z.enum(['bar', 'line', 'pie', 'area', 'scatter']).default('bar')
        .describe('图表类型：bar/line/pie/area/scatter；scatter 时 labels 与 values 均为数值'),
      title: z.string().optional().describe('图表标题'),
    },
  },
  async ({ labels, values, series, chart_type, title }) => {
    try {
      let seriesList = series;
      if (!seriesList || !seriesList.length) {
        if (!values || !values.length) {
          throw new Error('请提供 values（单系列）或 series（多系列）');
        }
        if (values.length !== labels.length) {
          throw new Error(`labels 长度(${labels.length}) 与 values 长度(${values.length}) 不一致`);
        }
        seriesList = [{ name: '数值', values }];
      }
      const { svg } = generateChart(chart_type, title, labels, seriesList);
      const { absPath, webPath } = saveChartSvg(svg, `data-${chart_type}`);
      const lines = [
        `图表已生成: ${absPath}`,
        webPath ? `网页访问地址: ${webPath}` : null,
        webPath ? `请在回复中原样包含以下 Markdown 图片以展示图表:\n![${title || '图表'}](${webPath})` : null,
      ].filter(Boolean);
      return { content: [{ type: 'text', text: lines.join('\n') }] };
    } catch (err) {
      return { isError: true, content: [{ type: 'text', text: `生成图表失败: ${err.message}` }] };
    }
  },
);

/* ---------------- 启动 ---------------- */

process.on('uncaughtException', (err) => log('uncaughtException:', err?.message));
process.on('unhandledRejection', (err) => log('unhandledRejection:', err?.message || err));

await server.connect(new StdioServerTransport());
log('MCP server "calculator-excel" started, chart output dir:', OUTPUT_DIR);
