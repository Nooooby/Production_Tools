"""排班主入口模块。

按顺序执行：LoadInputs -> AssignStations -> InsertBreaks -> ValidateOverlaps -> ExportMaster。
每一步使用结构化参数并返回统一的 Result。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, Generic, Iterable, List, Optional, TypeVar


T = TypeVar("T")


@dataclass
class Result(Generic[T]):
    success: bool
    message: str
    data: Optional[T] = None
    errors: List[str] = field(default_factory=list)

    @staticmethod
    def ok(message: str, data: Optional[T] = None) -> "Result[T]":
        return Result(success=True, message=message, data=data)

    @staticmethod
    def fail(message: str, errors: Optional[Iterable[str]] = None) -> "Result[T]":
        return Result(success=False, message=message, errors=list(errors or []))


@dataclass
class ScheduleLogger:
    """可选日志写入器（支持 Log Sheet 或其他目标）。"""

    log_sheet: Optional[Any] = None

    def log(self, department: str, step: str, status: str, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if self.log_sheet is not None and hasattr(self.log_sheet, "append"):
            self.log_sheet.append([timestamp, department, step, status, message])
        else:
            # 退化为内存或控制台日志
            print(f"[{timestamp}] [{department}] {step} - {status}: {message}")


@dataclass
class LoadInputsParams:
    departments: Dict[str, List[Dict[str, Any]]]


@dataclass
class InputBundle:
    departments: Dict[str, List[Dict[str, Any]]]


@dataclass
class AssignStationsParams:
    inputs: InputBundle


@dataclass
class StationAssignment:
    departments: Dict[str, List[Dict[str, Any]]]


@dataclass
class InsertBreaksParams:
    assignments: StationAssignment
    break_rules: Dict[str, Any] = field(default_factory=dict)


@dataclass
class BreakPlan:
    schedule: Dict[str, List[Dict[str, Any]]]


@dataclass
class ValidateOverlapsParams:
    break_plan: BreakPlan


@dataclass
class ExportMasterParams:
    break_plan: BreakPlan
    output_options: Dict[str, Any] = field(default_factory=dict)


@dataclass
class RunScheduleParams:
    load_inputs: LoadInputsParams
    break_rules: Dict[str, Any] = field(default_factory=dict)
    output_options: Dict[str, Any] = field(default_factory=dict)
    log_sheet: Optional[Any] = None


# ============================================================================
# 核心流程函数（全部使用结构化参数，统一返回 Result）
# ============================================================================

def LoadInputs(params: LoadInputsParams, logger: Optional[ScheduleLogger] = None) -> Result[InputBundle]:
    if not params.departments:
        return Result.fail("未提供任何部门输入数据。")

    logger = logger or ScheduleLogger()
    for department, rows in params.departments.items():
        if not rows:
            logger.log(department, "LoadInputs", "WARN", "部门输入为空")
        else:
            logger.log(department, "LoadInputs", "OK", f"读取 {len(rows)} 行输入")

    return Result.ok("输入加载完成。", InputBundle(departments=params.departments))


def AssignStations(params: AssignStationsParams, logger: Optional[ScheduleLogger] = None) -> Result[StationAssignment]:
    logger = logger or ScheduleLogger()
    assignments: Dict[str, List[Dict[str, Any]]] = {}

    for department, rows in params.inputs.departments.items():
        assigned_rows = []
        for row in rows:
            assigned = dict(row)
            assigned.setdefault("station", "UNASSIGNED")
            assigned_rows.append(assigned)
        assignments[department] = assigned_rows
        logger.log(department, "AssignStations", "OK", f"已分配 {len(assigned_rows)} 条记录")

    return Result.ok("工位分配完成。", StationAssignment(assignments))


def InsertBreaks(params: InsertBreaksParams, logger: Optional[ScheduleLogger] = None) -> Result[BreakPlan]:
    logger = logger or ScheduleLogger()
    schedule: Dict[str, List[Dict[str, Any]]] = {}

    for department, rows in params.assignments.departments.items():
        updated_rows = []
        for row in rows:
            updated = dict(row)
            updated.setdefault("breaks", [])
            updated_rows.append(updated)
        schedule[department] = updated_rows
        logger.log(department, "InsertBreaks", "OK", f"已插入 {len(updated_rows)} 条记录的休息时间")

    return Result.ok("休息时间插入完成。", BreakPlan(schedule))


def ValidateOverlaps(params: ValidateOverlapsParams, logger: Optional[ScheduleLogger] = None) -> Result[BreakPlan]:
    logger = logger or ScheduleLogger()
    overlap_errors: List[str] = []

    for department, rows in params.break_plan.schedule.items():
        seen = set()
        for row in rows:
            key = (row.get("employee"), row.get("start"), row.get("end"))
            if key in seen:
                overlap_errors.append(f"{department} 存在重复排班: {key}")
            seen.add(key)

        if overlap_errors:
            logger.log(department, "ValidateOverlaps", "FAIL", "检测到排班冲突")
        else:
            logger.log(department, "ValidateOverlaps", "OK", "未检测到排班冲突")

    if overlap_errors:
        return Result.fail("排班冲突校验失败。", overlap_errors)

    return Result.ok("排班冲突校验通过。", params.break_plan)


def ExportMaster(params: ExportMasterParams, logger: Optional[ScheduleLogger] = None) -> Result[Dict[str, List[Dict[str, Any]]]]:
    logger = logger or ScheduleLogger()
    export_payload = params.break_plan.schedule

    for department in export_payload.keys():
        logger.log(department, "ExportMaster", "OK", "已生成导出数据")

    return Result.ok("主表导出数据准备完成。", export_payload)


# ============================================================================
# 主入口
# ============================================================================

def RunSchedule(params: RunScheduleParams) -> Result[Dict[str, List[Dict[str, Any]]]]:
    logger = ScheduleLogger(params.log_sheet)
    step_messages: List[str] = []

    load_result = LoadInputs(params.load_inputs, logger)
    step_messages.append(f"LoadInputs: {load_result.message}")
    if not load_result.success:
        return Result.fail("排班流程失败。", step_messages + load_result.errors)

    assign_result = AssignStations(AssignStationsParams(load_result.data), logger)
    step_messages.append(f"AssignStations: {assign_result.message}")
    if not assign_result.success:
        return Result.fail("排班流程失败。", step_messages + assign_result.errors)

    break_result = InsertBreaks(
        InsertBreaksParams(assign_result.data, break_rules=params.break_rules),
        logger,
    )
    step_messages.append(f"InsertBreaks: {break_result.message}")
    if not break_result.success:
        return Result.fail("排班流程失败。", step_messages + break_result.errors)

    validate_result = ValidateOverlaps(ValidateOverlapsParams(break_result.data), logger)
    step_messages.append(f"ValidateOverlaps: {validate_result.message}")
    if not validate_result.success:
        return Result.fail("排班流程失败。", step_messages + validate_result.errors)

    export_result = ExportMaster(
        ExportMasterParams(validate_result.data, output_options=params.output_options),
        logger,
    )
    step_messages.append(f"ExportMaster: {export_result.message}")
    if not export_result.success:
        return Result.fail("排班流程失败。", step_messages + export_result.errors)

    summary = " | ".join(step_messages)
    return Result.ok(f"排班流程完成。{summary}", export_result.data)
