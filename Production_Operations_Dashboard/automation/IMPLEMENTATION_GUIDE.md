# Phase 3 实现指南 - 生产日报自动化系统

**状态**: ✅ Phase 3 Sub-task 1 - Daily Report Automation
**版本**: 1.0
**创建日期**: 2026-01-01
**最后更新**: 2026-01-01

---

## 📋 目录

1. [系统概述](#系统概述)
2. [架构设计](#架构设计)
3. [安装部署](#安装部署)
4. [配置指南](#配置指南)
5. [使用方法](#使用方法)
6. [故障排除](#故障排除)
7. [监控维护](#监控维护)

---

## 系统概述

### 功能特性

生产日报自动化系统实现以下核心功能：

| 功能 | 描述 | 状态 |
|------|------|------|
| **日报生成** | 自动从 Excel 提取当日数据，生成格式化日报 | ✅ |
| **Yield 监控** | 实时检测 Yield < 95% 项，自动分级警告 | ✅ |
| **邮件通知** | 生成邮件并发送给相关人员 | ✅ |
| **日志记录** | 完整的执行日志用于审计和故障排除 | ✅ |
| **定时执行** | Windows 任务计划程序每日自动执行 | ✅ |

### 文件位置

```
automation/
├── daily_report_automation.py      # 核心脚本（主程序）
├── requirements.txt                # Python 依赖列表
├── config_email.json              # 邮件配置模板
├── schedule_daily_report.bat       # 批处理脚本（任务调用）
├── setup_task_scheduler.ps1       # 任务计划程序设置
└── IMPLEMENTATION_GUIDE.md        # 本文件
```

---

## 架构设计

### 系统流程图

```
v39_Normalized.xlsx
    ↓
DailyReportGenerator (主类)
    ├── load_data()           → 加载 Excel 文件
    ├── extract_daily_data()  → 提取订单、Yield、进度数据
    ├── check_yield_alerts()  → 检测 Yield < 95% 警告
    ├── generate_report_file() → 生成 Excel 日报
    └── send_email_notification() → 发送邮件通知
    ↓
reports/Daily_Report_YYYY-MM-DD.xlsx
logs/daily_report_YYYYMMDD.log
邮件通知 (Manager, Quality, Supervisor)
```

### 数据流

```
数据来源:
├── 05_Daily_Orders      → 订单汇总 (总数、完成数、完成率)
├── 00_Yield_Rates       → Yield 分析 (平均、最低、警告检测)
└── 13_Progress_Track    → 生产进度 (总产量、部门产量)
    ↓
处理:
├── 计算关键指标
├── Yield < 95% 检测
├── 警告分级 (严重 <90%, 警告 90-95%)
└── 生成日报内容
    ↓
输出:
├── Excel 日报文件
├── 邮件通知
└── 日志文件
```

### 类设计

#### DailyReportGenerator

**主要方法**:

```python
__init__(excel_path)
    初始化生成器
    参数: Excel 文件路径

load_data() → bool
    加载 Excel 工作簿
    返回: 成功/失败

extract_daily_data() → bool
    从各工作表提取当日数据
    更新: self.daily_data 字典

check_yield_alerts() → bool
    检测 Yield < 95% 的警告项
    更新: self.alerts 列表

generate_report_file() → bool
    生成 Excel 日报文件
    创建: reports/Daily_Report_*.xlsx

send_email_notification() → bool
    发送邮件通知
    收件人: Config.RECIPIENT_LIST

run() → bool
    执行完整流程
    返回: 整体成功/失败
```

**属性**:

```python
excel_path: str           # Excel 文件路径
wb: Workbook             # 加载的 Excel 对象
report_date: str         # 报告日期 (YYYY-MM-DD)
report_datetime: str     # 报告时间戳 (YYYY-MM-DD HH:MM:SS)
daily_data: dict        # 当日数据
  ├── total_orders: int
  ├── completed_orders: int
  ├── completion_rate: float
  ├── total_cases: float
  ├── avg_yield: float
  └── min_yield: float
alerts: list            # Yield 警告列表
  └── {sku, description, yield_pct, level, gap, timestamp}
report_file: Path       # 生成的报告文件路径
```

---

## 安装部署

### 系统要求

- **操作系统**: Windows 10 / Windows Server 2016 或更高版本
- **Python**: 3.8 或更高版本
- **磁盘空间**: 至少 100 MB
- **网络**: 用于邮件发送（可选）

### 步骤 1: 安装 Python

1. 访问 https://www.python.org/downloads/
2. 下载 Python 3.10+ 版本
3. 安装时 **勾选 "Add Python to PATH"**
4. 验证安装:
   ```bash
   python --version
   ```

### 步骤 2: 安装 Python 依赖

在 PowerShell 中执行：

```powershell
cd C:\Projects\Production_management\Production_Operations_Dashboard\automation
pip install -r requirements.txt
```

或者单独安装：

```bash
pip install openpyxl==3.10.10
pip install pandas==2.0.3
pip install python-dateutil==2.8.2
```

### 步骤 3: 验证安装

```bash
python -c "import openpyxl; import pandas; print('✅ 依赖安装成功')"
```

### 步骤 4: 创建所需目录

```bash
cd C:\Projects\Production_management\Production_Operations_Dashboard
mkdir reports
mkdir logs
```

---

## 配置指南

### 配置 1: 邮件服务器

编辑 `daily_report_automation.py` 中的 `Config` 类：

```python
class Config:
    # SMTP 服务器配置
    SMTP_SERVER = "smtp.gmail.com"    # 邮件服务器地址
    SMTP_PORT = 587                  # SMTP 端口 (587 for TLS)
    SENDER_EMAIL = "production@company.com"  # 发件邮箱
    SENDER_PASSWORD = os.getenv("EMAIL_PASSWORD", "")  # 从环境变量读取

    # 收件人列表
    RECIPIENT_LIST = [
        "manager@company.com",
        "supervisor@company.com"
    ]

    YIELD_ALERT_RECIPIENTS = [
        "quality@company.com",
        "manager@company.com"
    ]
```

### 配置 2: 设置环境变量

为了安全起见，邮箱密码应通过环境变量设置，不要硬编码。

**Windows 设置步骤**:

1. 右键点击 "此电脑" → 属性
2. 左侧菜单 → 高级系统设置
3. 点击 "环境变量" 按钮
4. "新建" 用户变量:
   - 变量名: `EMAIL_PASSWORD`
   - 变量值: `你的邮箱密码或应用专用密码`
5. 点击确定，重启 PowerShell/CMD

**验证**:
```bash
echo %EMAIL_PASSWORD%
```

### 配置 3: Gmail 配置（如使用 Gmail）

如果使用 Gmail 账户：

1. 启用 2 步验证: https://myaccount.google.com/security
2. 生成应用专用密码: https://myaccount.google.com/apppasswords
3. 选择应用为 "Mail"，设备为 "Windows 电脑"
4. 复制生成的 16 位密码，设置为环境变量 `EMAIL_PASSWORD`

### 配置 4: Yield 警告阈值

根据业务需求调整阈值：

```python
YIELD_CRITICAL = 0.90      # < 90%: 严重告警（红色）
YIELD_WARNING = 0.95       # 90-95%: 警告（黄色）
```

---

## 使用方法

### 方法 1: 手动执行

在命令行中直接运行脚本：

```bash
cd C:\Projects\Production_management\Production_Operations_Dashboard\automation
python daily_report_automation.py
```

**预期输出**:
```
初始化日录生成器, 报告日期: 2026-01-01
加载 Excel 文件: C:\Projects\...
Excel 文件加载成功
开始提取当日数据...
...
✅ 日报已生成: ...\reports\Daily_Report_2026-01-01.xlsx
   警告项数: 2
   警告详情:
     - SKU123: 88.5% [严重]
     - SKU456: 92.3% [警告]
```

### 方法 2: 自动执行（推荐）

使用 Windows 任务计划程序每日自动运行。

#### 自动设置（简单方式）

以管理员身份运行 PowerShell：

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
& 'C:\Projects\Production_management\Production_Operations_Dashboard\automation\setup_task_scheduler.ps1'
```

#### 手动设置（详细方式）

1. 打开 "任务计划程序" (taskschd.msc)
2. 右侧 "创建任务..."
3. "常规" 标签:
   - 名称: `Production Daily Report`
   - 选中 "使用最高权限运行此任务"
4. "触发器" 标签:
   - 新建 → 每日
   - 时间: 17:00 (下午 5 点)
5. "操作" 标签:
   - 程序或脚本: `C:\Projects\Production_management\Production_Operations_Dashboard\automation\schedule_daily_report.bat`
   - 起始于: `C:\Projects\Production_management\Production_Operations_Dashboard\automation`
6. "条件" 标签:
   - 若有网络可用时运行 ✓
7. "设置" 标签:
   - 允许按需运行任务 ✓
   - 仅在用户登录时运行 ✓
8. 点击 "确定"

### 查看执行结果

#### 查看日报文件

```
C:\Projects\Production_management\Production_Operations_Dashboard\reports\
├── Daily_Report_2026-01-01.xlsx
├── Daily_Report_2026-01-02.xlsx
└── Daily_Report_2026-01-03.xlsx
```

#### 查看执行日志

```
C:\Projects\Production_management\Production_Operations_Dashboard\logs\
├── daily_report_20260101.log
├── daily_report_20260102.log
└── schedule.log
```

**日志内容示例**:
```
[2026-01-01 17:00:15] 初始化日报生成器, 报告日期: 2026-01-01
[2026-01-01 17:00:16] 加载 Excel 文件: C:\Projects\...
[2026-01-01 17:00:16] Excel 文件加载成功
[2026-01-01 17:00:17] 开始提取当日数据...
[2026-01-01 17:00:18] 订单统计: 总数=45, 完成=43, 完成率=95.6%
[2026-01-01 17:00:19] Yield 统计: 平均=94.2%, 最低=88.5%
[2026-01-01 17:00:20] 生产统计: 总产量=12500 cases
[2026-01-01 17:00:20] 当日数据提取完成
[2026-01-01 17:00:20] 开始检查 Yield 警告...
[2026-01-01 17:00:21] Yield 警告 [严重]: SKU123 (鸡胸肉) = 88.5%
[2026-01-01 17:00:21] Yield 警告 [警告]: SKU456 (鸡腿肉) = 92.3%
[2026-01-01 17:00:21] 检查完成, 发现 2 个警告
[2026-01-01 17:00:21] 开始生成日报文件...
[2026-01-01 17:00:22] 日报文件生成成功: C:\Projects\.../Daily_Report_2026-01-01.xlsx
[2026-01-01 17:00:22] 开始发送邮件通知...
[2026-01-01 17:00:23] 邮件发送成功, 收件人: quality@company.com, manager@company.com
[2026-01-01 17:00:23] ✅ 日报自动化流程完成成功
```

---

## 故障排除

### 问题 1: "ModuleNotFoundError: No module named 'openpyxl'"

**原因**: 依赖未安装

**解决**:
```bash
pip install -r requirements.txt
```

### 问题 2: "FileNotFoundError: [Errno 2] No such file or directory: 'v39_Normalized.xlsx'"

**原因**: Excel 文件路径不正确

**解决**:
1. 检查 Excel 文件是否存在于指定位置
2. 修改 `daily_report_automation.py` 中 `Config.EXCEL_PATH` 为正确路径
3. 确保路径中的文件夹存在

### 问题 3: 邮件发送失败

**原因**: 邮件配置不正确或网络问题

**解决步骤**:

1. 验证 EMAIL_PASSWORD 环境变量是否设置:
   ```bash
   echo %EMAIL_PASSWORD%
   ```

2. 如使用 Gmail，验证是否：
   - 启用了 2 步验证
   - 生成了应用专用密码
   - 使用的是 16 位专用密码而不是账户密码

3. 检查防火墙是否阻止 SMTP 连接:
   ```bash
   telnet smtp.gmail.com 587
   ```

4. 查看日志文件了解具体错误信息

### 问题 4: Excel 读取错误 "#N/A" 或 "ValueError"

**原因**: Excel 文件数据格式问题

**解决**:

1. 打开 v39_Normalized.xlsx 验证数据完整性
2. 确保以下工作表存在且格式正确：
   - 05_Daily_Orders
   - 00_Yield_Rates
   - 13_Progress_Track

3. 检查列名是否与代码中定义的一致

### 问题 5: 任务计划程序任务失败

**检查步骤**:

1. 打开任务计划程序 (taskschd.msc)
2. 查找任务 "Production Daily Report"
3. 右键 → 属性 → 历史记录
4. 查看最近的错误信息

**常见原因**:

- Python 未在 PATH 中 → 重新安装 Python 并勾选 "Add to PATH"
- 文件权限问题 → 确保任务以管理员身份运行
- 脚本路径错误 → 使用完整的绝对路径

---

## 监控维护

### 日常监控

**每周检查**:

1. 检查日报文件是否正常生成
   ```bash
   dir C:\Projects\Production_management\Production_Operations_Dashboard\reports
   ```

2. 查看日志文件了解执行状态
   ```bash
   tail -f C:\Projects\Production_management\Production_Operations_Dashboard\logs\daily_report_*.log
   ```

3. 验证邮件是否正常发送（查看收件箱）

**月度检查**:

1. 统计 Yield 警告次数和改进趋势
2. 收集用户反馈
3. 更新告警阈值（如需要）

### 日志分析

日志文件位置: `logs/daily_report_YYYYMMDD.log`

**日志级别**:

- `INFO`: 正常执行信息
- `WARNING`: 警告信息（如 Yield 低于阈值）
- `ERROR`: 错误信息（如文件读取失败）

**常见查询**:

```bash
# 查看某日执行状况
findstr "2026-01-01" logs\daily_report_*.log

# 查看所有错误
findstr /S "ERROR" logs\

# 查看 Yield 警告
findstr /S "警告" logs\
```

### 性能优化

如果执行时间过长（> 5 分钟）：

1. **优化 Excel 读取**: 确保 v39_Normalized.xlsx 文件已经过 Phase 2 优化
2. **分页处理**: 对大数据集使用 `chunksize` 参数
3. **缓存数据**: 在内存中缓存常用查询

### 备份和恢复

**备份日报**:

```bash
# 月度备份
robocopy C:\Projects\Production_management\Production_Operations_Dashboard\reports ^
         D:\Backups\Reports\2026-01 /S /Y
```

**恢复日报**:

如果日报生成失败，可以手动触发：

```bash
cd C:\Projects\Production_management\Production_Operations_Dashboard\automation
python daily_report_automation.py
```

---

## 后续优化计划

### Sub-task 2: Yield < 95% 自动警告

下一阶段将实现：

- ✅ **实时监控**: 每 30 分钟检查一次 Yield
- ✅ **即时告警**: Yield < 95% 时立即推送警告
- ✅ **多渠道通知**: 邮件、短信、系统通知
- ✅ **告警历史**: 记录所有告警事件

### Sub-task 3: 数据分析增强

后续将添加：

- 📈 **趋势分析**: 7 日移动平均、周对周对比
- 🔍 **异常检测**: 自动识别异常 Yield 和超期订单
- 🔮 **预测分析**: 明日产量和完成率预测
- 📊 **仪表板**: 实时 KPI 展示和警告面板

---

## 技术支持

### 常见问题

**Q: 如何修改执行时间?**
A: 在任务计划程序中编辑任务，修改触发器的时间。

**Q: 如何跳过邮件发送，仅生成日报?**
A: 在 `daily_report_automation.py` 中注释掉 `self.send_email_notification()` 行。

**Q: 如何添加更多收件人?**
A: 修改 `Config.RECIPIENT_LIST` 和 `Config.YIELD_ALERT_RECIPIENTS` 列表。

**Q: 日报文件的格式可以自定义吗?**
A: 可以，修改 `generate_report_file()` 方法中的 Excel 格式设置部分。

### 联系方式

如有问题，请查阅：

1. 日志文件: `logs/daily_report_*.log`
2. 本指南: `automation/IMPLEMENTATION_GUIDE.md`
3. 源代码注释: `automation/daily_report_automation.py`

---

**版本历史**

| 版本 | 日期 | 描述 |
|------|------|------|
| 1.0 | 2026-01-01 | 初始版本，实现日报生成和 Yield 检测 |

---

**最后更新**: 2026-01-01
**维护者**: Claude Code (AI Assistant)
**状态**: 生产就绪 ✅
