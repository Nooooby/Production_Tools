# FINAL_v70.xlsx 结构与公式链分析

## 1) Cases → WIP → Raw → Cages 公式链检查

### 02_TrayPack_Order（TrayPack）
- Product_Group: `P` 列（XLOOKUP 取 00_SKU_Master）
- Avg_Case_Weight: `Q` 列（XLOOKUP 取 00_SKU_Master）
- **WIP_kg**: `R` 列（`R3 = M3 * Q3`）
- **缺失**: 未发现 Raw_kg / Cages 相关公式

### 03_BulkPack_Order（BulkPack）
- Product_Group: `K` 列（XLOOKUP 取 00_SKU_Master）
- Avg_Case_Weight: `L` 列（XLOOKUP 取 00_SKU_Master）
- Yield_Rate: `M` 列（XLOOKUP 取 00_Yield_Rates）
- **WIP_kg**: `N` 列（`N3 = I3 * L3`）
- **缺失**: 未发现 Raw_kg / Cages 相关公式

### 04_Bagging_Order（Bagging）
- Product_Group: `N` 列（XLOOKUP 取 00_SKU_Master）
- Avg_Case_Weight: `O` 列（XLOOKUP 取 00_SKU_Master）
- Yield_Rate: `P` 列（XLOOKUP 取 00_Yield_Rates）
- **WIP_kg**: `Q` 列（`Q3 = I3 * O3`）
- **缺失**: 未发现 Raw_kg / Cages 相关公式

### 05_Daily_Orders（Daily Orders）
- WIP_Code: `I` 列（XLOOKUP 取 00_SKU_Master）
- **WIP_kg**: `J` 列（`J3 = Avg_Case_Weight * N3`）
- **缺失**: 未发现 Raw_kg / Cages 相关公式

### 14_Demand_Aggregation（中间表）
- TrayPack WIP 汇总：`B5 = SUMIF('02_TrayPack_Order'!P:P, A5, '02_TrayPack_Order'!R:R)`
- BulkPack WIP 汇总：`C5 = SUMIF('03_BulkPack_Order'!K:K, A5, '03_BulkPack_Order'!N:N)`
- Bagging WIP 汇总：`D5 = SUMIF('04_Bagging_Order'!N:N, A5, '04_Bagging_Order'!Q:Q)`
- Total WIP：`E5 = B5 + C5 + D5`

### 14_Production_Planning（规划）
- Demand 区域（Row 30-38）使用 INDEX/MATCH 从 `14_Demand_Aggregation` 引用 WIP 需求
- 未检测到 `/680` 或 Raw_kg 公式，Column G 也未发现公式

**结论**: v70 中的公式链在 **WIP_kg 汇总**层面完整，但 **Raw_kg 与 Cages 的计算链条未在工作簿公式中出现**。

---

## 2) 跨表引用关系（按公式扫描）

| 工作表 | 直接引用的表 |
|------|--------------|
| 02_TrayPack_Order | 00_SKU_Master, 05_Daily_Orders |
| 03_BulkPack_Order | 00_SKU_Master, 00_Yield_Rates, 13_Progress_Track |
| 04_Bagging_Order | 00_SKU_Master, 00_Yield_Rates, 13_Progress_Track |
| 05_Daily_Orders | 00_SKU_Master, 02_TrayPack_Order, 13_Progress_Track |
| 06_Resource_Plan | 00_SKU_Master |
| 12_Executive_Dash | 01_Cages_Plan, 03_BulkPack_Order, 04_Bagging_Order, 05_Daily_Orders |
| 14_Demand_Aggregation | 02_TrayPack_Order, 03_BulkPack_Order, 04_Bagging_Order |
| 14_Production_Planning | 00_Yield_Rates, 01_Cages_Plan, 06_Resource_Plan, 14_Demand_Aggregation |
| 15_5Day_Forecast | 00_SKU_Master, 05_Daily_Orders |
| 15_Weekly_Plan | 15_5Day_Forecast |
