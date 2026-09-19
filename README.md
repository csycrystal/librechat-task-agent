# 任务智能 Agent 助手 · LibreChat 二次开发

> 基于 [LibreChat](https://github.com/danny-avila/LibreChat) 二次开发的**私有化任务智能助手**：本地大模型（Ollama + Qwen2.5）驱动，通过 **MCP 协议**挂载自研工具链（智能计算 / Excel 分析 / 数据可视化 / 单位换算），并配备一套蓝绿色新拟物（Neumorphism）界面，支持亮色 / 暗色双模式。

![License](https://img.shields.io/badge/license-MIT-blue) ![Base](https://img.shields.io/badge/base-LibreChat%20v0.8.8-10b981) ![Model](https://img.shields.io/badge/model-Qwen2.5--3B%20%7C%20Ollama-orange) ![MCP](https://img.shields.io/badge/MCP-5%20tools-6366f1)

**在线演示**：<https://task-agent.vip.cpolar.cn>（运行在个人电脑上，电脑关机时离线；可自助注册体验）

---

## 原项目归属

- **本项目基于 LibreChat 二次开发**
- 原仓库地址：<https://github.com/danny-avila/LibreChat>
- 原项目开源协议：**MIT License**
- 原项目 README 保留于 [README.librechat-original.md](./README.librechat-original.md)，根目录 `LICENSE` 中原版权声明完整保留

> 本项目仅用于个人作品集学习展示，**不作任何商业用途**。

---

## 新增功能

### 1. MCP 自研工具链（5 个工具，stdio 接入）

位于 [mcp-servers/calculator-excel](./mcp-servers/calculator-excel/server.mjs)，基于官方 `@modelcontextprotocol/sdk` 实现：

| 工具 | 能力 |
|---|---|
| `calculate` | 安全表达式计算器：词法分析 + 调度场算法生成 AST 求值，**不使用 eval**，支持函数、括号、百分比 |
| `convert_unit` | 中英双语单位换算：长度 / 重量 / 温度（含℃℉开尔文）/ 面积 / 时间 / 容积 |
| `read_excel` | 读取 `.xlsx / .xls / .csv` 并输出结构化汇总；修复了 Windows 下 CSV 中文 GBK 乱码问题 |
| `excel_chart` | 基于 Excel 工作表数据直接生成图表 |
| `visualize_data` | 通用数据可视化，支持柱状图 / 折线图 / 饼图 / 面积图 / 散点图五种 SVG 图表 |

### 2. 蓝绿新拟物全局主题（亮 / 暗双模式）

在 [client/src/style.css](./client/src/style.css) 末尾**只追加、不覆盖**地新增 `neu-*` 设计系统：双向柔光阴影、凸起 / 内凹表面、渐变标题、圆形图标盘、交错入场动画，并通过 `prefers-reduced-motion` 适配无障碍需求。

### 3. 登录 / 注册页美化

[AuthLayout.tsx](./client/src/components/Auth/AuthLayout.tsx)、`LoginForm.tsx`、`Registration.tsx`：沉浸式背景 + 24px 圆角新拟物卡片 + 渐变标题 + 内凹输入框。

### 4. 欢迎屏快捷能力卡片

新增 [QuickStartCards.tsx](./client/src/components/Chat/QuickStartCards.tsx)：4 张卡片（智能计算 / Excel 分析 / 数据可视化 / 单位换算），点击即携带 `spec=task-agent` 整页导航并自动发起带工具调用的对话（修复了官方 `?submit=true` 仅在首屏消费参数导致的二次点击失效问题）。

### 5. 本地模型链路打通（Ollama → MCP 工具调用）

- 通过 OpenAI 兼容端点接入 Ollama，并在 `modelSpecs` 中定义默认模型规格 `task-agent`，由 agents 运行时自动展开注入 MCP 工具；
- 解决了 qwen / qwen2 架构模型不支持 function calling 的问题，改用 **Qwen2.5**（验证 `finish_reason=tool_calls` 真实往返成功）；
- 端到端验证：模型可自动调用计算器与图表工具完成多步任务。

### 6. 图表持久化与静态托管

MCP 生成的 SVG 落盘到 `data/charts/`，后端 [api/server/index.js](./api/server/index.js) 新增 `/mcp-charts` 静态托管路由，使 AI 生成的图表在对话中可长期访问。

### 7. 局域网开放 + 公网固定域名访问

- `HOST=0.0.0.0` 监听全部网卡 + Windows 防火墙入站规则，同一局域网内可直接访问；
- 通过 cpolar 内网穿透保留固定二级子域名，公网 HTTPS 地址永久不变，配合一键启动脚本（启动后自动弹窗并复制网址到剪贴板）。

---

## 技术栈

| 层 | 技术 |
|---|---|
| 前端 | React 18 · Vite · Tailwind CSS · TypeScript |
| 后端 | Node.js · Express · Mongoose · LibreChat monorepo（packages/api 等） |
| 数据库 | MongoDB 8 |
| 模型 | Ollama · Qwen2.5-3B-Instruct（本地运行，支持 tools） |
| 工具协议 | Model Context Protocol（MCP，stdio） |
| 图表 / 表格 | 自研 SVG 生成 · SheetJS(xlsx) |
| 网络 | cpolar 内网穿透（固定子域名 + HTTPS） |

---

## 本地运行

### 环境要求

- Node.js ≥ 18（推荐 20+）
- MongoDB
- [Ollama](https://ollama.com/) 并拉取工具调用模型：`ollama pull qwen2.5:3b`
- Windows / macOS / Linux 均可

### 安装与启动

```bash
# 1. 安装依赖（monorepo 根目录）
npm ci

# 2. 安装 MCP 工具服务依赖
cd mcp-servers/calculator-excel && npm install && cd ../../

# 3. 准备配置文件（仓库仅提供示例，真实配置不入库）
cp .env.example .env
cp librechat.example.yaml librechat.yaml

# 4. 构建前端
npm run build:client

# 5. 启动后端（同时托管前端产物）
npm run backend
```

打开 <http://localhost:3080> 注册首个账号即可使用。

### 关键配置

`.env`（核心项）：

```ini
HOST=0.0.0.0
OPENAI_API_KEY=ollama-local
# 只写 base URL，LibreChat 会自动追加 /chat/completions
OPENAI_REVERSE_PROXY=http://localhost:11434/v1
OPENAI_SUMMARIZE=false
```

`librechat.yaml`（MCP 工具 + 模型规格的最小示例）：

```yaml
mcpServers:
  calculator:
    type: stdio
    command: node
    args:
      - ./mcp-servers/calculator-excel/server.mjs

endpoints:
  custom: # 也可使用 openAI 端点，配合上面的 OPENAI_REVERSE_PROXY
    - name: Ollama
      apiKey: ollama-local
      baseURL: http://localhost:11434/v1
      models:
        default: [qwen2.5:3b]

modelSpecs:
  list:
    - name: task-agent
      label: 任务助手
      default: true
      preset:
        endpoint: openAI
        model: qwen2.5:3b
      mcpServers: [calculator]
```

> ⚠️ **模型必须是 qwen2.5 或更新架构**。qwen / qwen2（如 qwen:7b、qwen:1.8b）经 Ollama 实测不支持 tools，会直接返回 "model does not support tools"。

---

## 二改文件索引

| 类型 | 文件 |
|---|---|
| MCP 工具服务（新增） | `mcp-servers/calculator-excel/server.mjs` |
| 快捷卡片（新增） | `client/src/components/Chat/QuickStartCards.tsx` |
| 新拟物主题（追加） | `client/src/style.css` |
| 欢迎屏挂载 | `client/src/components/Chat/Landing.tsx` |
| 登录 / 注册页 | `client/src/components/Auth/AuthLayout.tsx`、`LoginForm.tsx`、`Registration.tsx` |
| 图表静态托管 | `api/server/index.js`（`/mcp-charts` 路由） |
| 忽略规则（追加） | `.gitignore`（排除 `.env`、`.tools/`、运行数据等） |

> 真实配置（`.env`、`librechat.yaml`）、本地数据库、模型文件与 cpolar 令牌均已通过 `.gitignore` 排除，不在仓库中。

---

## 免责声明

本仓库为个人学习与作品集展示项目，所使用的第三方开源项目（LibreChat 及其依赖）版权归原作者所有。原项目以 MIT License 授权，本项目在其基础上修改，同样遵循 MIT 协议，**不用于任何商业目的**。
