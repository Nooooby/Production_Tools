# Excel 文件结构分析（FINAL_v70.xlsx）

## 概述
- **文件名**: `FINAL_v70.xlsx`
- **总工作表数**: 14 个
- **核心新增**: `14_Demand_Aggregation` 中间表（替代 14_Production_Planning 的复杂 SUMPRODUCT）

---

## 工作表清单与范围（按工作簿元数据）

| 工作表 | 使用范围 | 公式数量 | 说明 |
|------|---------|--------|------|
| **00_SKU_Master** | A1:Q379 | 0 | SKU 主数据 |
| **00_Yield_Rates** | A1:U34 | 1 | 产率数据 |
| **01_Cages_Plan** | A2:O13 | 14 | 鸡笼库存/计划 |
| **02_TrayPack_Order** | B2:R265 | 1,841 | TrayPack 订单 |
| **03_BulkPack_Order** | B2:O201 | 432 | BulkPack 订单 |
| **04_Bagging_Order** | A1:R25 | 188 | Bagging 订单 |
| **05_Daily_Orders** | B1:N271 | 3,229 | 当日订单汇总 |
| **06_Resource_Plan** | A2:L55 | 153 | 原料计划 |
| **12_Executive_Dash** | A1:J26 | 21 | KPI 仪表板 |
| **13_Progress_Track** | A1:N108 | 0 | 进度追踪 |
| **14_Demand_Aggregation** | A1:E24 | 56 | 需求中间表 |
| **14_Production_Planning** | A1:I96 | 150 | 生产规划 |
| **15_5Day_Forecast** | A1:AS272 | 6,767 | 5 日预测 |
| **15_Weekly_Plan** | A1:I67 | 55 | 周计划 |

---

## 数据流简表（v70）

```
00_SKU_Master
   ↓
02_TrayPack_Order / 03_BulkPack_Order / 04_Bagging_Order
   ↓ (WIP_kg 汇总)
14_Demand_Aggregation
   ↓
14_Production_Planning
```

---

## 备注
- v70 中已存在 `14_Demand_Aggregation` 工作表，公式以 SUMIF 汇总 WIP_kg。
- 若需追踪完整公式链（Cases → WIP → Raw → Cages），请重点检查：
  - 订单表中的 WIP_kg 列
  - 14_Demand_Aggregation 的 SUMIF 汇总
  - 14_Production_Planning 的需求与库存对比区域
