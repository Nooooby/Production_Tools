"""
生产日报自动化系统 - Phase 3 实现
Production Daily Report Automation System

功能:
1. 从 v39_Normalized.xlsx 提取当日生产数据
2. 计算关键指标 (订单数、产量、Yield等)
3. 检测 Yield < 95% 的警告项
4. 生成日报文件 (Excel 格式)
5. 发送邮件通知 (含警告详情)
6. 记录执行日志

作者: Claude Code
创建日期: 2026-01-01
"""

import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import logging
import os
from pathlib import Path
import json

# ============================================================================
# 配置部分
# ============================================================================

class Config:
    """系统配置"""

    # 文件路径
    EXCEL_PATH = r"C:\Projects\Production_management\Production_Operations_Dashboard\v39_Normalized.xlsx"
    REPORT_DIR = r"C:\Projects\Production_management\Production_Operations_Dashboard\reports"
    LOG_DIR = r"C:\Projects\Production_management\Production_Operations_Dashboard\logs"

    # 邮件配置
    SMTP_SERVER = "smtp.gmail.com"  # 需要配置实际的邮件服务器
    SMTP_PORT = 587
    SENDER_EMAIL = "production@company.com"  # 需要配置实际邮箱
    SENDER_PASSWORD = os.getenv("EMAIL_PASSWORD", "")  # 从环境变量读取
    RECIPIENT_LIST = [
        "manager@company.com",
        "supervisor@company.com"
    ]
    YIELD_ALERT_RECIPIENTS = [
        "quality@company.com",
        "manager@company.com"
    ]

    # Yield 警告阈值
    YIELD_CRITICAL = 0.90      # < 90%: 严重
    YIELD_WARNING = 0.95       # 90-95%: 警告

    # 日志配置
    LOG_LEVEL = logging.INFO
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'


# ============================================================================
# 日志设置
# ============================================================================

def setup_logging():
    """设置日志系统"""
    log_dir = Path(Config.LOG_DIR)
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"daily_report_{datetime.now().strftime('%Y%m%d')}.log"

    logging.basicConfig(
        level=Config.LOG_LEVEL,
        format=Config.LOG_FORMAT,
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

    return logging.getLogger(__name__)

logger = setup_logging()


# ============================================================================
# 日报生成器类
# ============================================================================

class DailyReportGenerator:
    """日报生成器 - 核心业务逻辑"""

    def __init__(self, excel_path):
        """初始化"""
        self.excel_path = excel_path
        self.wb = None
        self.report_date = datetime.now().strftime("%Y-%m-%d")
        self.report_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.daily_data = {}
        self.alerts = []
        self.report_file = None

        logger.info(f"初始化日报生成器, 报告日期: {self.report_date}")

    def load_data(self):
        """从 Excel 加载数据"""
        try:
            logger.info(f"加载 Excel 文件: {self.excel_path}")
            self.wb = openpyxl.load_workbook(self.excel_path, data_only=False)
            logger.info("Excel 文件加载成功")
            return True
        except Exception as e:
            logger.error(f"加载 Excel 失败: {str(e)}")
            return False

    def extract_daily_data(self):
        """从各工作表提取当日数据"""
        try:
            logger.info("开始提取当日数据...")

            # 1. 从 05_Daily_Orders 提取订单数据
            daily_orders_sheet = self.wb['05_Daily_Orders']
            df_orders = pd.read_excel(
                self.excel_path,
                sheet_name='05_Daily_Orders',
                header=0
            )

            self.daily_data['total_orders'] = len(df_orders)
            self.daily_data['completed_orders'] = len(df_orders[df_orders['Status'] == 'Completed']) \
                if 'Status' in df_orders.columns else 0
            self.daily_data['completion_rate'] = (
                self.daily_data['completed_orders'] / self.daily_data['total_orders'] * 100
            ) if self.daily_data['total_orders'] > 0 else 0

            logger.info(f"订单统计: 总数={self.daily_data['total_orders']}, "
                       f"完成={self.daily_data['completed_orders']}, "
                       f"完成率={self.daily_data['completion_rate']:.1f}%")

            # 2. 从 00_Yield_Rates 提取 Yield 数据
            df_yields = pd.read_excel(
                self.excel_path,
                sheet_name='00_Yield_Rates',
                header=0
            )

            if 'Yield' in df_yields.columns:
                yields = pd.to_numeric(df_yields['Yield'], errors='coerce')
                self.daily_data['avg_yield'] = yields.mean() * 100  # 转换为百分比
                self.daily_data['min_yield'] = yields.min() * 100
                logger.info(f"Yield 统计: 平均={self.daily_data['avg_yield']:.1f}%, "
                           f"最低={self.daily_data['min_yield']:.1f}%")

            # 3. 从 13_Progress_Track 提取进度数据
            df_progress = pd.read_excel(
                self.excel_path,
                sheet_name='13_Progress_Track',
                header=0
            )

            if 'Cases_Produced' in df_progress.columns:
                total_cases = pd.to_numeric(
                    df_progress['Cases_Produced'],
                    errors='coerce'
                ).sum()
                self.daily_data['total_cases'] = total_cases
                logger.info(f"生产统计: 总产量={total_cases:.0f} cases")

            logger.info("当日数据提取完成")
            return True

        except Exception as e:
            logger.error(f"提取当日数据失败: {str(e)}")
            return False

    def check_yield_alerts(self):
        """检查 Yield < 95% 的警告"""
        try:
            logger.info("开始检查 Yield 警告...")

            df_yields = pd.read_excel(
                self.excel_path,
                sheet_name='00_Yield_Rates',
                header=0
            )

            self.alerts = []

            # 检查每个 SKU 的 Yield
            for idx, row in df_yields.iterrows():
                sku = row.get('SKU', f"SKU_{idx}")
                sku_desc = row.get('SKU_Description', '')
                yield_value = row.get('Yield', 1.0)

                # 确保 yield_value 是数字
                try:
                    yield_pct = float(yield_value) * 100
                except (ValueError, TypeError):
                    continue

                # 检查警告条件
                if yield_pct < Config.YIELD_WARNING:
                    alert_level = "严重" if yield_pct < Config.YIELD_CRITICAL else "警告"

                    alert = {
                        'sku': sku,
                        'description': sku_desc,
                        'yield_pct': yield_pct,
                        'level': alert_level,
                        'gap': Config.YIELD_WARNING - yield_pct,
                        'timestamp': self.report_datetime
                    }

                    self.alerts.append(alert)
                    logger.warning(f"Yield 警告 [{alert_level}]: {sku} ({sku_desc}) = {yield_pct:.1f}%")

            logger.info(f"检查完成, 发现 {len(self.alerts)} 个警告")
            return True

        except Exception as e:
            logger.error(f"检查 Yield 警告失败: {str(e)}")
            return False

    def generate_report_file(self):
        """生成日报 Excel 文件"""
        try:
            logger.info("开始生成日报文件...")

            # 创建新的 Excel 工作簿
            report_wb = openpyxl.Workbook()
            report_ws = report_wb.active
            report_ws.title = "日报"

            # 设置列宽
            report_ws.column_dimensions['A'].width = 25
            report_ws.column_dimensions['B'].width = 20

            # 标题
            report_ws['A1'] = f"生产日报 - {self.report_date}"
            report_ws['A1'].font = openpyxl.styles.Font(size=14, bold=True)

            row = 3

            # 报告头信息
            report_ws[f'A{row}'] = "报告日期:"
            report_ws[f'B{row}'] = self.report_date
            row += 1

            report_ws[f'A{row}'] = "生成时间:"
            report_ws[f'B{row}'] = self.report_datetime
            row += 2

            # 核心数据部分
            report_ws[f'A{row}'] = "=== 一、生产概览 ==="
            row += 1

            report_ws[f'A{row}'] = "总订单数:"
            report_ws[f'B{row}'] = self.daily_data.get('total_orders', 0)
            row += 1

            report_ws[f'A{row}'] = "完成订单数:"
            report_ws[f'B{row}'] = self.daily_data.get('completed_orders', 0)
            row += 1

            report_ws[f'A{row}'] = "完成率:"
            report_ws[f'B{row}'] = f"{self.daily_data.get('completion_rate', 0):.1f}%"
            row += 1

            report_ws[f'A{row}'] = "总产量 (Cases):"
            report_ws[f'B{row}'] = f"{self.daily_data.get('total_cases', 0):.0f}"
            row += 2

            # Yield 分析
            report_ws[f'A{row}'] = "=== 二、Yield 分析 ==="
            row += 1

            report_ws[f'A{row}'] = "平均 Yield:"
            report_ws[f'B{row}'] = f"{self.daily_data.get('avg_yield', 0):.1f}%"
            row += 1

            report_ws[f'A{row}'] = "最低 Yield:"
            report_ws[f'B{row}'] = f"{self.daily_data.get('min_yield', 0):.1f}%"
            row += 2

            # 警告部分
            if self.alerts:
                report_ws[f'A{row}'] = "=== 三、Yield 警告 ⚠️ ==="
                row += 1

                report_ws[f'A{row}'] = "SKU"
                report_ws[f'B{row}'] = "Yield %"
                report_ws[f'C{row}'] = "级别"
                report_ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
                report_ws[f'B{row}'].font = openpyxl.styles.Font(bold=True)
                report_ws[f'C{row}'].font = openpyxl.styles.Font(bold=True)
                row += 1

                for alert in self.alerts:
                    report_ws[f'A{row}'] = alert['sku']
                    report_ws[f'B{row}'] = f"{alert['yield_pct']:.1f}%"
                    report_ws[f'C{row}'] = alert['level']

                    # 根据级别着色
                    if alert['level'] == "严重":
                        cell_fill = openpyxl.styles.PatternFill(
                            start_color="FF0000",
                            end_color="FF0000",
                            fill_type="solid"
                        )
                    else:
                        cell_fill = openpyxl.styles.PatternFill(
                            start_color="FFFF00",
                            end_color="FFFF00",
                            fill_type="solid"
                        )

                    report_ws[f'A{row}'].fill = cell_fill
                    report_ws[f'B{row}'].fill = cell_fill
                    report_ws[f'C{row}'].fill = cell_fill

                    row += 1

            # 建议部分
            row += 1
            report_ws[f'A{row}'] = "=== 四、建议 ==="
            row += 1

            if self.alerts:
                report_ws[f'A{row}'] = "1. 立即关注上述警告项目"
                row += 1
                report_ws[f'A{row}'] = "2. 检查生产工艺参数"
                row += 1
                report_ws[f'A{row}'] = "3. 增加质检频率"
                row += 1
            else:
                report_ws[f'A{row}'] = "所有指标正常，继续维持当前生产状态"
                row += 1

            # 保存文件
            report_dir = Path(Config.REPORT_DIR)
            report_dir.mkdir(parents=True, exist_ok=True)

            self.report_file = report_dir / f"Daily_Report_{self.report_date}.xlsx"
            report_wb.save(str(self.report_file))

            logger.info(f"日报文件生成成功: {self.report_file}")
            return True

        except Exception as e:
            logger.error(f"生成日报文件失败: {str(e)}")
            return False

    def send_email_notification(self):
        """发送邮件通知"""
        try:
            # 如果没有配置邮件参数，则跳过
            if not Config.SENDER_EMAIL or not Config.SENDER_PASSWORD:
                logger.warning("邮件配置不完整，跳过邮件发送")
                return True

            logger.info("开始发送邮件通知...")

            # 准备邮件内容
            if self.alerts:
                subject = f"【生产预警】{self.report_date} - Yield 低于目标"
                recipients = Config.YIELD_ALERT_RECIPIENTS
            else:
                subject = f"【日报】{self.report_date} - 生产日报"
                recipients = Config.RECIPIENT_LIST

            # 创建邮件
            msg = MIMEMultipart()
            msg['From'] = Config.SENDER_EMAIL
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = subject

            # 邮件正文
            body = self._generate_email_body()
            msg.attach(MIMEText(body, 'html', 'utf-8'))

            # 附加日报文件
            if self.report_file and self.report_file.exists():
                with open(self.report_file, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {self.report_file.name}'
                    )
                    msg.attach(part)

            # 发送邮件
            try:
                with smtplib.SMTP(Config.SMTP_SERVER, Config.SMTP_PORT) as server:
                    server.starttls()
                    server.login(Config.SENDER_EMAIL, Config.SENDER_PASSWORD)
                    server.send_message(msg)

                logger.info(f"邮件发送成功, 收件人: {', '.join(recipients)}")
                return True

            except smtplib.SMTPAuthenticationError:
                logger.error("邮件发送失败: 认证错误 (用户名或密码错误)")
                return False
            except smtplib.SMTPException as e:
                logger.error(f"邮件发送失败: {str(e)}")
                return False

        except Exception as e:
            logger.error(f"准备邮件失败: {str(e)}")
            return False

    def _generate_email_body(self):
        """生成邮件正文 (HTML 格式)"""
        alert_html = ""

        if self.alerts:
            alert_html = """
            <h2 style="color: #cc0000;">⚠️ Yield 警告项目</h2>
            <table style="border-collapse: collapse; width: 100%;">
                <tr style="background-color: #f0f0f0;">
                    <th style="border: 1px solid #ccc; padding: 8px;">SKU ID</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">Yield %</th>
                    <th style="border: 1px solid #ccc; padding: 8px;">级别</th>
                </tr>
            """

            for alert in self.alerts:
                bg_color = "#ffcccc" if alert['level'] == "严重" else "#ffff99"
                alert_html += f"""
                <tr style="background-color: {bg_color};">
                    <td style="border: 1px solid #ccc; padding: 8px;">{alert['sku']}</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">{alert['yield_pct']:.1f}%</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">{alert['level']}</td>
                </tr>
                """

            alert_html += "</table>"

        html_body = f"""
        <html>
            <head>
                <meta charset="utf-8">
            </head>
            <body style="font-family: Arial, sans-serif; line-height: 1.6;">
                <h1 style="color: #333;">生产日报 - {self.report_date}</h1>
                <p><strong>生成时间:</strong> {self.report_datetime}</p>

                <h2>📊 生产概览</h2>
                <ul>
                    <li>总订单数: {self.daily_data.get('total_orders', 0)}</li>
                    <li>完成订单数: {self.daily_data.get('completed_orders', 0)}</li>
                    <li>完成率: {self.daily_data.get('completion_rate', 0):.1f}%</li>
                    <li>总产量: {self.daily_data.get('total_cases', 0):.0f} Cases</li>
                </ul>

                <h2>📈 Yield 分析</h2>
                <ul>
                    <li>平均 Yield: {self.daily_data.get('avg_yield', 0):.1f}%</li>
                    <li>最低 Yield: {self.daily_data.get('min_yield', 0):.1f}%</li>
                </ul>

                {alert_html}

                <h2>建议行动</h2>
                <ol>
                    <li>检查上述警告项目的生产工艺</li>
                    <li>联系质量部门增加检查频率</li>
                    <li>查看详细报告文件了解更多信息</li>
                </ol>

                <hr>
                <p style="color: #666; font-size: 12px;">
                    本报告由自动化系统生成
                    <br>详细数据见附件: Daily_Report_{self.report_date}.xlsx
                </p>
            </body>
        </html>
        """

        return html_body

    def run(self):
        """执行完整流程"""
        logger.info("=" * 80)
        logger.info("开始执行日报自动化流程")
        logger.info("=" * 80)

        steps = [
            ("加载 Excel 数据", self.load_data),
            ("提取当日数据", self.extract_daily_data),
            ("检查 Yield 警告", self.check_yield_alerts),
            ("生成日报文件", self.generate_report_file),
            ("发送邮件通知", self.send_email_notification),
        ]

        success = True
        for step_name, step_func in steps:
            logger.info(f"执行: {step_name}...")
            if not step_func():
                logger.error(f"失败: {step_name}")
                success = False
                break
            logger.info(f"完成: {step_name}")

        logger.info("=" * 80)
        if success:
            logger.info("✅ 日报自动化流程完成成功")
        else:
            logger.info("❌ 日报自动化流程执行失败")
        logger.info("=" * 80)

        return success


# ============================================================================
# 主程序
# ============================================================================

def main():
    """主函数"""
    try:
        generator = DailyReportGenerator(Config.EXCEL_PATH)
        success = generator.run()

        if success and generator.report_file:
            print(f"\n✅ 日报已生成: {generator.report_file}")
            print(f"   警告项数: {len(generator.alerts)}")
            if generator.alerts:
                print("   警告详情:")
                for alert in generator.alerts:
                    print(f"     - {alert['sku']}: {alert['yield_pct']:.1f}% [{alert['level']}]")
        else:
            print("\n❌ 日报生成失败")
            return 1

        return 0

    except Exception as e:
        logger.error(f"程序执行异常: {str(e)}", exc_info=True)
        return 1


if __name__ == "__main__":
    exit_code = main()
    exit(exit_code)
