# 技术工具栈

> 入场券也是解决复杂问题的抓手。

---

## 1. Excel进阶（建模核心）

### 必备技能
- INDEX-MATCH 替代 VLOOKUP（更灵活、更快）
- OFFSET + 命名范围：动态数据范围
- Power Query：数据清洗与合并（替代手工VBA）
- 数据表（DataTable）：敏感性分析利器
- 循环引用与迭代计算：三表联动的技术基础
- 自定义函数（Lambda）：复杂逻辑封装

### 建模规范
| 规范 | 说明 |
|------|------|
| 输入区 | 蓝色字体，所有假设集中一处 |
| 计算区 | 黑色字体，公式透明无硬编码 |
| 输出区 | 绿色字体，关键结果高亮 |
| 版本控制 | 每次修改记录变更日志 |
| 错误检查 | BALANCE CHECK行（资产=负债+权益） |

---

## 2. SQL数据提取（关键分水岭）

### 为什么SQL是分水岭
- 能直接连接ERP/CRM数据库，绕过导出Excel
- 处理百万行数据，Excel根本打不开
- 可重复、可审计的查询逻辑

### CFO常用SQL场景
```sql
-- 区域销售汇总
SELECT region, SUM(revenue), SUM(cogs), 
       SUM(revenue)-SUM(cogs) AS gross_profit,
       (SUM(revenue)-SUM(cogs))/SUM(revenue) AS gross_margin
FROM sales 
WHERE period BETWEEN '2025-01' AND '2025-12'
GROUP BY region
ORDER BY gross_margin DESC;

-- 客户集中度分析
SELECT customer, SUM(revenue) AS rev,
       SUM(revenue)/(SELECT SUM(revenue) FROM sales WHERE period='2025') AS pct
FROM sales WHERE period='2025'
GROUP BY customer
ORDER BY rev DESC LIMIT 20;
```

---

## 3. Python/Pandas（数据清洗与分析）

### 核心场景
- 百万行数据清洗：空值处理、异常值过滤、格式统一
- 多源数据合并：ERP + CRM + 外部数据
- 自动化报表：周报/月报脚本化
- 统计分析：趋势、分布、相关性

### 快速模板
```python
import pandas as pd

# 读取 + 清洗
df = pd.read_sql(query, engine)
df = df.dropna(subset=['revenue'])
df['margin'] = (df['revenue'] - df['cogs']) / df['revenue']

# 分组聚合
summary = df.groupby(['region','product']).agg(
    revenue=('revenue','sum'),
    margin=('margin','mean'),
    count=('revenue','count')
).round(4)

# 异常检测
anomalies = df[df['margin'] < df['margin'].quantile(0.05)]
```

---

## 4. Power BI / Tableau（可视化与看板）

### CFO看板设计原则
- **一页看全**：关键指标一屏展示
- **钻取设计**：总览→区域→产品→客户
- **异常高亮**：低于阈值自动标红
- **移动适配**：手机端可查看核心指标

### 必备看板
| 看板 | 核心指标 | 更新频率 |
|------|---------|---------|
| 经营驾驶舱 | 营收/利润/现金流/ROIC | 日/周 |
| 资本配置 | ROI by项目/预算执行率 | 月 |
| 风险监控 | 逾期率/拨备/LCR/汇率敞口 | 日 |
| 人力成本 | 人均产出/薪酬占比/编制执行 | 月 |
