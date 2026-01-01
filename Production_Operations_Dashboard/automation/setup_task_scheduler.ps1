# ============================================================================
# Windows 任务计划程序设置脚本
#
# 用法: 以管理员身份运行此 PowerShell 脚本
# 命令: powershell -ExecutionPolicy Bypass -File setup_task_scheduler.ps1
#
# 功能: 自动在 Windows 任务计划程序中创建每日 17:00 运行的日报任务
# ============================================================================

param(
    [string]$TaskName = "Production_Daily_Report",
    [string]$ScriptPath = "C:\Projects\Production_management\Production_Operations_Dashboard\automation\schedule_daily_report.bat",
    [string]$TaskTime = "17:00",
    [string]$PythonPath = "python"
)

# ============================================================================
# 函数定义
# ============================================================================

function Check-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Create-DailyTask {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string]$Time
    )

    try {
        # 检查任务是否已存在
        $existingTask = Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue
        if ($existingTask) {
            Write-Host "⚠️  任务已存在: $Name"
            Write-Host "是否要删除并重新创建? (Y/N)" -ForegroundColor Yellow
            $response = Read-Host
            if ($response -eq 'Y' -or $response -eq 'y') {
                Unregister-ScheduledTask -TaskName $Name -Confirm:$false
                Write-Host "✅ 已删除旧任务"
            } else {
                Write-Host "❌ 取消操作"
                return $false
            }
        }

        # 创建任务触发器 (每天指定时间)
        $trigger = New-ScheduledTaskTrigger -Daily -At $Time

        # 创建任务操作
        $action = New-ScheduledTaskAction -Execute $ScriptPath

        # 创建任务设置
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -Compatibility Win8 `
            -StartWhenAvailable `
            -RunOnlyIfNetworkAvailable

        # 注册任务
        Register-ScheduledTask `
            -TaskName $Name `
            -Trigger $trigger `
            -Action $action `
            -Settings $settings `
            -Description "生产日报自动化 - 每日 $Time 运行" `
            -RunLevel Highest

        Write-Host "✅ 任务创建成功: $Name"
        Write-Host "   触发时间: 每天 $Time"
        Write-Host "   执行脚本: $ScriptPath"

        return $true
    }
    catch {
        Write-Host "❌ 创建任务失败: $_" -ForegroundColor Red
        return $false
    }
}

function Show-TaskStatus {
    param([string]$TaskName)

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($task) {
            Write-Host ""
            Write-Host "任务信息:" -ForegroundColor Cyan
            Write-Host "  名称: $($task.TaskName)"
            Write-Host "  状态: $($task.State)"
            Write-Host "  启用: $($task.Enabled)"

            # 显示最后运行信息
            $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
            if ($taskInfo.LastRunTime) {
                Write-Host "  最后运行: $($taskInfo.LastRunTime)"
                Write-Host "  最后结果: $($taskInfo.LastTaskResult)"
            }
        }
    }
    catch {
        Write-Host "⚠️  无法获取任务信息"
    }
}

function Test-PythonInstallation {
    try {
        $version = & python --version 2>&1
        Write-Host "✅ Python 已安装: $version"
        return $true
    }
    catch {
        Write-Host "❌ Python 未找到，请先安装 Python 3.8+"
        Write-Host "   下载地址: https://www.python.org/downloads/"
        return $false
    }
}

function Install-Dependencies {
    Write-Host ""
    Write-Host "安装 Python 依赖包..." -ForegroundColor Cyan

    try {
        & python -m pip install -r "$(Split-Path $ScriptPath)\requirements.txt" -q
        Write-Host "✅ 依赖包安装成功"
        return $true
    }
    catch {
        Write-Host "❌ 依赖包安装失败: $_" -ForegroundColor Red
        return $false
    }
}

function Show-Configuration {
    Write-Host ""
    Write-Host "当前配置:" -ForegroundColor Cyan
    Write-Host "  任务名称: $TaskName"
    Write-Host "  执行脚本: $ScriptPath"
    Write-Host "  触发时间: $TaskTime (每天)"
    Write-Host ""
}

# ============================================================================
# 主程序
# ============================================================================

Write-Host ""
Write-Host "======================================================================"
Write-Host "Windows 任务计划程序 - 生产日报自动化设置"
Write-Host "======================================================================"
Write-Host ""

# 检查管理员权限
if (-not (Check-Admin)) {
    Write-Host "❌ 错误: 此脚本需要以管理员身份运行"
    Write-Host "   请右键点击 PowerShell，选择'以管理员身份运行'"
    exit 1
}

# 检查 Python 安装
if (-not (Test-PythonInstallation)) {
    exit 1
}

# 显示当前配置
Show-Configuration

# 提示用户确认
Write-Host "是否继续? (Y/N)" -ForegroundColor Yellow
$confirm = Read-Host
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "❌ 取消操作"
    exit 0
}

Write-Host ""

# 安装依赖
if (-not (Install-Dependencies)) {
    exit 1
}

# 创建任务
if (-not (Create-DailyTask -Name $TaskName -ScriptPath $ScriptPath -Time $TaskTime)) {
    exit 1
}

# 显示任务状态
Show-TaskStatus -TaskName $TaskName

# 完成信息
Write-Host ""
Write-Host "======================================================================"
Write-Host "✅ 设置完成!"
Write-Host "======================================================================"
Write-Host ""
Write-Host "后续步骤:"
Write-Host "1. 配置邮件参数:"
Write-Host "   - 编辑 automation\daily_report_automation.py"
Write-Host "   - 修改 Config 类中的邮件设置"
Write-Host "2. 设置环境变量 EMAIL_PASSWORD:"
Write-Host "   - 右键'此电脑' > 属性 > 高级系统设置 > 环境变量"
Write-Host "   - 新建用户变量: EMAIL_PASSWORD = 你的邮箱密码"
Write-Host "3. 任务将在每天 $TaskTime 自动运行"
Write-Host ""
Write-Host "手动测试:"
Write-Host "   PowerShell: & '$ScriptPath'"
Write-Host "   或双击: $ScriptPath"
Write-Host ""
Write-Host "查看任务:"
Write-Host "   控制面板 > 管理工具 > 任务计划程序"
Write-Host "   或运行: taskschd.msc"
Write-Host ""

Write-Host "======================================================================"
