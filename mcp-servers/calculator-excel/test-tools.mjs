// 临时测试脚本：直接通过 MCP stdio 客户端验证 4 个工具
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import * as XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// 1. 生成示例 Excel（供演示与测试）
const samplesDir = path.join(__dirname, 'samples');
fs.mkdirSync(samplesDir, { recursive: true });
const samplePath = path.join(samplesDir, '2024年销售数据.xlsx');
const rows = [
  ['月份', '销售额', '成本', '利润'],
  ['一月', 12000, 8200, 3800],
  ['二月', 13500, 8600, 4900],
  ['三月', 15800, 9100, 6700],
  ['四月', 14200, 8800, 5400],
  ['五月', 17600, 9800, 7800],
  ['六月', 19800, 10500, 9300],
];
const ws = XLSX.utils.aoa_to_sheet(rows);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '销售数据');
XLSX.writeFile(wb, samplePath);
console.log('SAMPLE-XLSX:', samplePath);

// 2. 连接 MCP 服务器
const transport = new StdioClientTransport({
  command: 'node',
  args: [path.join(__dirname, 'server.mjs')],
});
const client = new Client({ name: 'test-client', version: '1.0.0' });
await client.connect(transport);

const tools = await client.listTools();
console.log('TOOLS:', tools.tools.map((t) => t.name).join(', '));

// 3. 测试 calculate
const calc = await client.callTool({
  name: 'calculate',
  arguments: { expression: '(3+4)*sqrt(16)-max(2,9,5)' },
});
console.log('CALC:', calc.content[0].text);
const calc2 = await client.callTool({
  name: 'calculate',
  arguments: { expression: '(12000*0.13)/1.13' },
});
console.log('CALC2:', calc2.content[0].text);
const calc3 = await client.callTool({
  name: 'calculate',
  arguments: { expression: '2^-3 + 10%3 + PI*2' },
});
console.log('CALC3:', calc3.content[0].text);

// 4. 测试 read_excel
const read = await client.callTool({
  name: 'read_excel',
  arguments: { file_path: samplePath, max_rows: 5 },
});
console.log('READ_EXCEL:\n', read.content[0].text.slice(0, 600));

// 5. 测试 excel_chart（柱状图）
const chart1 = await client.callTool({
  name: 'excel_chart',
  arguments: { file_path: samplePath, chart_type: 'bar', title: '2024上半年销售额与成本' },
});
console.log('EXCEL_CHART_BAR:\n', chart1.content[0].text);

// 6. 测试 excel_chart（饼图，指定 y 列）
const chart2 = await client.callTool({
  name: 'excel_chart',
  arguments: { file_path: samplePath, chart_type: 'pie', y_col: '利润', title: '各月利润占比' },
});
console.log('EXCEL_CHART_PIE:\n', chart2.content[0].text);

// 7. 测试 visualize_data
const chart3 = await client.callTool({
  name: 'visualize_data',
  arguments: {
    labels: ['Q1', 'Q2', 'Q3', 'Q4'],
    values: [41300, 51600, 54900, 62000],
    chart_type: 'line',
    title: '季度销售趋势',
  },
});
console.log('VISUALIZE_LINE:\n', chart3.content[0].text);

// 8. 测试 convert_unit
const conv1 = await client.callTool({
  name: 'convert_unit',
  arguments: { value: 100, from: '英里', to: '公里' },
});
console.log('CONV1:', conv1.content[0].text);
const conv2 = await client.callTool({
  name: 'convert_unit',
  arguments: { value: 32, from: '华氏度', to: '摄氏度' },
});
console.log('CONV2:', conv2.content[0].text);
const conv3 = await client.callTool({
  name: 'convert_unit',
  arguments: { value: 3, from: '斤', to: '克' },
});
console.log('CONV3:', conv3.content[0].text);

// 9. 面积图 + 散点图
const chart4 = await client.callTool({
  name: 'excel_chart',
  arguments: { file_path: samplePath, chart_type: 'area', title: '利润面积趋势' },
});
console.log('EXCEL_CHART_AREA:', chart4.content[0].text.split('\n')[1]);

// 10. 生成 CSV 样例并读取
const csvPath = path.join(samplesDir, '城市数据.csv');
fs.writeFileSync(csvPath, '城市,GDP,人口\n北京,41600,2189\n上海,44600,2487\n深圳,32300,1756\n', 'utf8');
const readCsv = await client.callTool({
  name: 'read_excel',
  arguments: { file_path: csvPath, max_rows: 3 },
});
console.log('READ_CSV:', readCsv.content[0].text.split('\n').slice(0, 6).join(' | '));

// 11. 散点图（两列数值：人口 vs GDP）
const chart5 = await client.callTool({
  name: 'excel_chart',
  arguments: { file_path: csvPath, chart_type: 'scatter', x_col: '人口', y_col: 'GDP', title: '人口-GDP 散点分布' },
});
console.log('EXCEL_CHART_SCATTER:', chart5.content[0].text.split('\n').slice(0, 2).join(' '));

await client.close();
console.log('ALL-TESTS-DONE');
process.exit(0);
