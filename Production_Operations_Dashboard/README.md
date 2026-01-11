# Production Operations Dashboard

生产运营仪表板与日报自动化项目，围绕 Excel 生产数据建立统一的订单输入、需求汇总、生产规划与日报输出流程。

## 当前版本

- 最新 Excel：`data/FINAL_v70.xlsx`
- 中间汇总表：`14_Demand_Aggregation`（WIP_kg 汇总）
- 规划表：`14_Production_Planning`

## 项目结构

```
Production_Operations_Dashboard/
├── automation/           # Python 自动化脚本
│   ├── daily_report_automation.py
│   ├── enhance_dashboard.py
│   ├── create_production_planning_v2.py
│   └── requirements.txt
├── data/                 # Excel 文件与备份
│   └── FINAL_v70.xlsx
├── docs/                 # 项目文档与分析记录
│   ├── Excel_Structure_Analysis.md
│   ├── FINAL_v70_Analysis.md
│   └── CHANGELOG_v69_20260106.md
└── README.md             # 项目说明（本文件）
```

## 快速上手

1. **确认文件**：将最新 Excel 放入 `data/`（当前为 `FINAL_v70.xlsx`）。
2. **安装依赖**：
   ```bash
   pip install -r automation/requirements.txt
   ```
3. **运行日报自动化**：
   ```bash
   python automation/daily_report_automation.py
   ```

## 关键文档

- `docs/Excel_Structure_Analysis.md`: v70 工作表结构与范围
- `docs/FINAL_v70_Analysis.md`: 公式链检查与跨表引用关系
- `docs/CHANGELOG_v69_20260106.md`: v69 中间表变更记录

## 说明

如需继续扩展 Raw_kg → Cages 的计算链，建议先审阅 `FINAL_v70_Analysis.md` 中的缺失项，再决定是否补回公式或通过脚本生成。
