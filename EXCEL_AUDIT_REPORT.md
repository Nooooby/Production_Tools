# Excel 审计报告 - 生产物料规划逻辑

**日期**: 2026-01-01  
**文件**: v39_Dashboard_Enhanced.xlsx  
**审计主题**: 验证生产物料转换逻辑实现情况

---

## 执行摘要

**转换逻辑要求**:
```
Cages Needed = (Cases × Avg_Case_Weight) ÷ Yield% ÷ 680kg/cage
```

**审计结论**: ⚠️ **部分实现** - 发现多处缺陷和改进空间

---

## 关键发现

### 🔴 **严重问题 #1: 订单表缺少转换逻辑**

**问题**: 
- `02_TrayPack_Order` 表
- `03_BulkPack_Order` 表  
- `04_Bagging_Order` 表

这些表应该包含每个订单的完整转换计算（Cases → WIP → Cages），但目前**没有实现**。

**影响**: 
- 无法在订单级别追踪鸡笼需求
- 无法按产品分析 Yield 对笼子的影响
- 生产计划缺乏详细数据支撑

**现状**:
```
缺少列：
├─ Avg_Case_Weight (从 00_SKU_Master 拉取)
├─ Product_Group (产品分组)
├─ Yield_Rate (从 00_Yield_Rates 拉取)
├─ WIP_kg (Cases × Avg_Case_Weight)
├─ Raw_kg_Needed (WIP_kg ÷ Yield%)
└─ Cages_Needed (Raw_kg_Needed ÷ 680)
```

---

### 🔴 **严重问题 #2: 14_Production_Planning 使用不正确的聚合方式**

**现有实现** (`create_production_planning_v2.py`):

```excel
=AVERAGE('00_SKU_Master'!F:F)    ← 所有产品平均重量
=AVERAGE('00_Yield_Rates'!E:E)   ← 所有产品平均 Yield
```

**问题**: 
- 使用**全部产品的平均值**而不是**每个订单对应产品的实际值**
- BSB 产品: 31% Yield, ThighMeat: 10.5% Yield → 差异巨大！
- 平均 Yield 无法反映真实情况

**正确做法**:
```
对于每个订单：
1. 获取 SKU → 映射到 Product_Group
2. Product_Group → 查找对应的 Yield%
3. SKU → 查找对应的 Avg_Case_Weight
4. Cases × Weight ÷ Yield% ÷ 680 = Cages
5. 按产品/订单类型汇总
```

---

### 🟡 **中等问题 #3: 数据引用链断裂**

**链条应该是**:
```
05_Daily_Orders (M:M 订单数)
    ↓ (引用 SKU)
00_SKU_Master (E=Product_Group, F=Avg_Case_Weight)
    ↓
00_Yield_Rates (E=Yield%)
```

**当前状态**: 
- 03_BulkPack_Order 数据来自 `10_Cone_Line` 而非 `05_Daily_Orders`
- 没有清晰的产品分组参考
- 无法追踪从 SKU → Weight → Yield 的完整链路

---

### 🟡 **中等问题 #4: 缺少验证和约束**

**应该存在**:
- ✗ 数据验证规则 (Cases > 0?)
- ✗ 产品分组验证 (Product_Group 是否有效?)
- ✗ Yield 异常检测 (< 95%?)
- ✗ 计算错误提示 (#DIV/0! 处理?)

---

## 详细分析

### 表 1: 02_TrayPack_Order

**状态**: ⚠️ 结构不完整

**缺失功能**:
- [ ] Cases 输入列 (应来自 05_Daily_Orders M 列)
- [ ] Product_Group 列 (SKU → Product_Group 映射)
- [ ] Avg_Case_Weight 列 (VLOOKUP 从 00_SKU_Master!F)
- [ ] Yield_Rate 列 (VLOOKUP 从 00_Yield_Rates!E)
- [ ] WIP_kg 列 (=Cases * Avg_Case_Weight)
- [ ] Raw_kg_Needed 列 (=IF(Yield_Rate=0, 0, WIP_kg/Yield_Rate*100))
- [ ] Cages_Needed 列 (=IF(Raw_kg_Needed=0, 0, Raw_kg_Needed/680))

**应该添加的公式示例**:
```excel
Product_Group: =VLOOKUP(SKU, '00_SKU_Master'!B:E, 4, FALSE)
Avg_Case_Weight: =VLOOKUP(SKU, '00_SKU_Master'!B:F, 5, FALSE)
Yield_Rate: =VLOOKUP(Product_Group, '00_Yield_Rates'!D:E, 2, FALSE)
WIP_kg: =Cases*Avg_Case_Weight
Raw_kg_Needed: =IF(Yield_Rate=0, 0, WIP_kg/Yield_Rate*100)
Cages_Needed: =IF(Raw_kg_Needed=0, 0, ROUNDUP(Raw_kg_Needed/680, 0))
```

---

### 表 2: 03_BulkPack_Order

**状态**: ⚠️ 同 TrayPack_Order (结构不完整)

**额外问题**:
- 数据来自 `10_Cone_Line` M 列，需要验证这个引用是否正确
- 没有中间映射表或引用说明

---

### 表 3: 04_Bagging_Order

**状态**: ⚠️ 同上 (结构不完整)

**数据源**: I5:I22 (订单数)
**缺失**: 完整的转换计算链

---

### 表 4: 00_SKU_Master

**状态**: ✅ 良好 (数据源正确)

**提供数据**:
- B: SKU 编号
- E: Product_Group (产品分类)
- F: Avg_Case_Weight (平均每 case 重量)

**建议改进**:
- [ ] 在 B 列添加唯一性约束 (SKU 不重复)
- [ ] E 列验证 (Product_Group 只允许已定义的值)
- [ ] F 列验证 (Avg_Case_Weight > 0?)

---

### 表 5: 00_Yield_Rates

**状态**: ✅ 良好 (数据源正确)

**提供数据**:
- B: Product (产品名)
- E: Adjusted Yield% (调整后产率)

**建议改进**:
- [ ] 与 00_SKU_Master 的 Product_Group 建立明确映射
- [ ] 添加 Yield% 异常检测 (< 95% 标记警告)
- [ ] 添加历史 Yield 数据追踪

---

## 改进建议

### 优先级 1️⃣: 立即修复 (高风险)

#### 建议 1.1: 在订单表中添加转换逻辑列

**对象**: 02_TrayPack_Order, 03_BulkPack_Order, 04_Bagging_Order

**步骤**:
1. 在每个表中添加新列 (顺序如下):
   ```
   现有列 ... [新增以下列]
   ├─ Product_Group (从 SKU 查找)
   ├─ Avg_Case_Weight (从 SKU_Master 查找)
   ├─ Yield_Rate (从 Yield_Rates 查找)
   ├─ WIP_kg (Cases × Weight)
   ├─ Raw_kg_Needed (WIP_kg ÷ Yield%)
   └─ Cages_Needed (Raw_kg_Needed ÷ 680)
   ```

2. 编写查找公式:
   ```excel
   Product_Group: =VLOOKUP(A2,'00_SKU_Master'!B:E,4,0)
   Avg_Case_Weight: =VLOOKUP(A2,'00_SKU_Master'!B:F,5,0)
   Yield_Rate: =VLOOKUP(Product_Group,
                        '00_Yield_Rates'!D:E,2,0)
   WIP_kg: =Cases*Avg_Case_Weight
   Raw_kg_Needed: =IF(OR(Yield_Rate=0,Yield_Rate=""),
                       0,
                       WIP_kg/(Yield_Rate/100))
   Cages_Needed: =IF(Raw_kg_Needed=0,
                     0,
                     ROUNDUP(Raw_kg_Needed/680,0))
   ```

3. 向下复制公式到所有数据行

4. 验证没有 #REF! 或 #VALUE! 错误

---

#### 建议 1.2: 修正 14_Production_Planning 的聚合逻辑

**对象**: 14_Production_Planning 工作表

**修改前**:
```excel
TrayPack Cases: =SUMIF('05_Daily_Orders'!M:M,">0")
Avg_Case_Weight: =AVERAGE('00_SKU_Master'!F:F)  ← 错误！
Yield: =AVERAGE('00_Yield_Rates'!E:E)           ← 错误！
```

**修改后** (需要更复杂的公式):
```excel
TrayPack Cases: =SUMIF('05_Daily_Orders'!M:M,">0")
  ↓
对于 TrayPack WIP 计算，需要按产品分组求和：
= SUMPRODUCT(('05_Daily_Orders'!M:M > 0) * 
             VLOOKUP('05_Daily_Orders'!SKU_Col,
                    '00_SKU_Master'!B:F, 5, 0) *
             '05_Daily_Orders'!M:M)

类似地处理 Yield 加权平均
```

**或** (推荐):
- 在辅助区域创建 Pivot Table 或汇总表
- 按 Product_Group 汇总，再计算加权平均
- 引用汇总结果而非原始数据

---

### 优先级 2️⃣: 重要改进 (中风险)

#### 建议 2.1: 添加 Product_Group → Yield_Rate 映射表

创建简化的映射表:
```
Product_Group | Adjusted_Yield%
BSB           | 31%
ThighMeat     | 10.5%
...           | ...
```

这样可以使用简单的 VLOOKUP，而不是多层联接。

---

#### 建议 2.2: 在订单表中添加数据验证

对关键列添加规则:
- Cases: >= 0 的整数
- Product_Group: 下拉列表 (来自 SKU_Master)
- Yield_Rate: >= 0%, <= 100%
- Cages_Needed: >= 0 的整数

---

#### 建议 2.3: 添加 Yield 异常检测

在订单表或汇总中:
```excel
Yield_Status: =IF(Yield_Rate<0.95,
                  "⚠️ 低于 95% - 异常",
                  "✅ 正常")
```

---

### 优先级 3️⃣: 长期优化 (低风险)

#### 建议 3.1: 创建标准化计算模板

对所有订单类型创建统一的列结构和公式。

#### 建议 3.2: 添加历史数据追踪

保存每日订单的计算结果，用于趋势分析。

#### 建议 3.3: 性能优化

如果订单表行数增长，考虑:
- 使用 INDEX/MATCH 替代 VLOOKUP
- 添加辅助缓存表
- 考虑迁移到 Power Query

---

## 验证检查清单

使用以下清单验证改进:

- [ ] 02_TrayPack_Order 包含所有 7 列 (Cases...Cages_Needed)
- [ ] 03_BulkPack_Order 包含所有 7 列
- [ ] 04_Bagging_Order 包含所有 7 列
- [ ] 所有公式无 #REF! 错误
- [ ] 所有公式无 #VALUE! 错误
- [ ] 至少 5 行样本订单的计算结果正确
- [ ] 14_Production_Planning 的聚合值与订单表的求和一致
- [ ] Yield < 95% 的订单被正确标记
- [ ] 添加了数据验证规则
- [ ] 日志记录所有更改

---

## 数据整合性检查

**应验证的关键数据流**:

```
05_Daily_Orders (M 列)
  ├─ Cases 数据 → 应进入 02_TrayPack_Order
  ├─ SKU 映射 → Product_Group
  └─ Product_Group → Yield_Rates (E 列)

10_Cone_Line (M 列)
  └─ Cases 数据 → 应进入 03_BulkPack_Order

04_Bagging_Order (I5:I22)
  └─ Cases 数据 → 已在表中

所有汇总 → 14_Production_Planning
  ├─ TrayPack Cases 汇总
  ├─ BulkPack Cases 汇总
  ├─ Bagging Cases 汇总
  └─ 总 Cages 需求
```

---

## 摘要

| 问题 | 严重性 | 影响 | 修复时间 |
|------|------|------|---------|
| 订单表缺少转换逻辑 | 🔴 高 | 无法追踪订单级笼子需求 | 2-3 小时 |
| 聚合使用平均值 | 🔴 高 | Cages 计算不准确 | 1-2 小时 |
| 缺少数据验证 | 🟡 中 | 易发生数据错误 | 1 小时 |
| 缺少 Yield 异常检测 | 🟡 中 | Yield < 95% 无警告 | 1 小时 |
| 缺少历史追踪 | 🟢 低 | 无趋势分析 | 后续功能 |

**预计总修复时间**: 5-7 小时

---

*报告生成时间: 2026-01-01*

