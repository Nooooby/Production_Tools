# 📅 Excel 生产线轮换休息排班系统（v4.0）

> 一个基于 **Excel + Office Scripts** 的生产线排班与休息自动化系统  
> 适用于 **鸡肉加工厂 Tray Pack / Cut-Up 等多工种、多工位场景**

---

## 📌 项目背景（Why）

在鸡肉加工产线中，排班与休息计划通常存在以下问题：

- 人多、工位多，手排容易出错  
- 关键工位（Loader / Trimmer / X-Ray）容易在休息时断人  
- 每天排班逻辑相同，却要重复手工操作  
- 员工不知道 **自己在哪个工位、什么时候休息**

本项目的目标是：

> **让主管 5 分钟完成排班，让员工一眼看懂自己的当天安排**

---

## 🎯 核心目标（What）

- ✅ 自动分配 **4 个 Period 的工位**
- ✅ 智能安排 **分批轮换休息**
- ✅ 保证 **关键工位持续有人**
- ✅ 每人 **至少一次休息**
- ✅ 输出 **一张主表（每人一行）**

---

## 🧠 系统设计思路（How）

### 架构分层

```

配置层（Excel）
│
├─ 人员 / 工种 / 工位 / 休息窗口
│
计算层（Office Scripts）
│
├─ 排班算法
├─ 休息轮换算法
├─ 覆盖 & 容量检查
│
输出层（Excel）
│
└─ Master Schedule + 可视化提示

````

> ❌ 不使用 VBA  
> ❌ 不使用 Power Query  
> ✅ 只依赖 Office 365 自带功能

---

## 🗂 Excel 文件结构

```text
Department_Schedule_Template.xlsx
│
├── 00_Dept_Setup          # 部门与全局参数
├── 01_People              # 人员名单 & 工种
├── 02_Station_Config      # 工位配置
├── 03_Pool_Config         # 工种（Pool）规则
├── 04_Break_Windows       # 休息窗口配置
│
├── 05_Master_Schedule     # ⭐ 主排班表（每人一行）
│
├── 08_Capacity_Monitor    # 休息人数容量监控
├── 09_Coverage_Check      # 关键工位覆盖检查
└── 99_Instructions        # 使用说明
````

---

## 👥 支持的工种（Pool）

当前系统支持 **12 个工种**：

* 原有（5）
  `LOADER, PACKER, DETECT, TRIMMER, REWORK`

* 新增（7）
  `THIGHDEBONE, ONELEG, SCALE, HANGING, DRUMLINE, WINGLINE, BONECHECK`

每个工种可独立配置：

* MinStaff（最少在岗）
* MaxStaff（最多人数）
* BreakMinOnDuty（休息时至少留人）
* 对应工位列表
* 是否为关键工位（RequiresCover）

---

## 📊 输出示例（Master Schedule）

> **一张表解决所有人**

| Name | P1 Station | P1 Break    | P2 Station | P2 Break    | P3 Station | P3 Break    | P4 Station |
| ---- | ---------- | ----------- | ---------- | ----------- | ---------- | ----------- | ---------- |
| 张三   | LOADER 1   | 07:45–08:00 | LOADER 1   | 11:00–11:30 | LOADER 1   | 13:30–13:45 | LOADER 1   |
| 李四   | PACKER 5   | 08:00–08:15 | PACKER 5   | 11:30–12:00 | PACKER 5   | 13:45–14:00 | Unassigned |

**条件格式说明：**

* 🟢 绿色：正常工位
* 🟡 黄色：Unassigned
* ⚪ 灰色：Unavailable

---

## 🚀 使用方式（给现场主管）

1. 复制当天上班人员到 `01_People`
2. 根据当天情况调整：

   * 工位启用 / 停用
   * 休息窗口
3. 点击 **Run Script**
4. 打印 `05_Master_Schedule` 发给员工

> ⏱️ 100 人规模运行时间：**约 2–3 秒**

---

## 🔧 技术说明（给维护者）

* 技术栈：`Excel Office Scripts (TypeScript)`
* 核心逻辑：

  * Pool 校验与匹配
  * 关键工位覆盖检测
  * 休息窗口容量控制
  * 主排班表合并输出
* 所有脚本 **必须包含中文注释**
* 批量写入，避免单元格循环

---

## 📦 项目状态

* 当前版本：**v4.0（计划模式）**
* 已完成：

  * v3.3 稳定算法
  * v4.0 架构升级
* 阻塞点：

  * 新增 7 个工种的参数确认
  * 主排班表格式最终确认

---

## 🏁 最终愿景

> **主管：5 分钟排完班**
> **员工：一张表看明白**
> **产线：不断人、不停线**

---

## 📄 License

Internal use only
Designed for food processing production environments

```

---

如果你愿意，下一步我可以帮你：

- 🔹 再精简一版（给 GitHub 公开看）
- 🔹 写一个 `99_Instructions` 给现场主管
- 🔹 拆成 `README + TECH.md + USER_GUIDE.md`
- 🔹 中英双语 README（给总部 / IT）

你直接说：**「改成哪种」**即可。
```
