import {
  ArrowLeftRight,
  Calculator,
  ChartPie,
  Table2,
  type LucideIcon,
} from 'lucide-react';

/**
 * 欢迎屏快捷能力卡片（二次开发新增）。
 * 点击后通过整页导航进入官方 URL 自动提交流程（?prompt=...&submit=true，
 * 见 hooks/Input/useQueryParams.ts —— 该机制仅在页面加载时处理参数），
 * 直接发起对话并自动触发对应 MCP 工具。
 * 蓝绿新拟物风格，样式类定义在 style.css 的 neu-* 设计系统中。
 */
const QUICK_START_CARDS: {
  icon: LucideIcon;
  title: string;
  description: string;
  prompt: string;
}[] = [
  {
    icon: Calculator,
    title: '智能计算',
    description: '复杂四则运算 · 函数 · 百分比，自动调用计算器',
    prompt: '请使用计算器计算：(12800 × 0.85 + 3600) ÷ 1.13，并给出计算步骤',
  },
  {
    icon: Table2,
    title: 'Excel 分析',
    description: '读取表格 · 汇总统计 · 一键生成图表',
    prompt:
      '请读取 mcp-servers/calculator-excel/samples/2024年销售数据.xlsx，汇总上半年销售额、成本和利润，并用柱状图可视化',
  },
  {
    icon: ChartPie,
    title: '数据可视化',
    description: '柱状图 / 折线 / 饼图 / 面积图 / 散点图',
    prompt: '请用饼图展示各区域销售占比：华东 4200、华北 3100、华南 2800、西南 1900、东北 1400',
  },
  {
    icon: ArrowLeftRight,
    title: '单位换算',
    description: '长度 · 重量 · 温度 · 面积 · 汇率外常用单位',
    prompt: '请帮我换算：100 英里等于多少公里，32 华氏度等于多少摄氏度，3 斤等于多少克',
  },
];

export default function QuickStartCards() {
  const handlePick = (prompt: string) => {
    // useQueryParams 仅在页面加载时消费参数，因此用整页导航。
    // spec=task-agent 是 librechat.yaml 中绑定 calculator MCP 的默认模型规格，
    // 保证卡片发起的对话自动具备工具调用能力。
    const params = new URLSearchParams({ spec: 'task-agent', prompt, submit: 'true' });
    window.location.assign(`${window.location.pathname}?${params.toString()}`);
  };

  return (
    <div className="mt-9 flex w-full max-w-2xl flex-col items-center px-2">
      <p className="neu-rise mb-5 text-center text-xs font-medium tracking-wide text-text-secondary">
        已接入 MCP 工具 · 计算器 / Excel 分析 / 数据可视化 / 单位换算
      </p>
      <div className="grid w-full grid-cols-1 gap-4 sm:grid-cols-2">
        {QUICK_START_CARDS.map((card, i) => {
          const Icon = card.icon;
          return (
            <button
              key={card.title}
              type="button"
              onClick={() => handlePick(card.prompt)}
              className="neu-raised-sm neu-pressable neu-rise group flex items-start gap-3.5 p-4 text-left"
              style={{ animationDelay: `${140 + i * 90}ms` }}
              aria-label={`${card.title}示例：${card.description}`}
            >
              <span className="neu-disk flex size-11 shrink-0 items-center justify-center rounded-full">
                <Icon
                  className="size-5"
                  style={{ color: 'rgb(var(--neu-icon-text))' }}
                  aria-hidden="true"
                />
              </span>
              <span className="flex min-w-0 flex-col gap-1 pt-0.5">
                <span className="text-sm font-semibold text-text-primary">{card.title}</span>
                <span className="text-xs leading-5 text-text-secondary">{card.description}</span>
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
