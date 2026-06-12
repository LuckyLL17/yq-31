import os
import json
from datetime import datetime


def generate_report(batch_task, report_path=None):
    output_dir = batch_task.get("output_dir", "./output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    if report_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(output_dir, f"batch_report_{batch_task['batch_id']}_{timestamp}.txt")

    report_content = _build_text_report(batch_task)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_content)

    return report_path


def generate_json_report(batch_task, report_path=None):
    output_dir = batch_task.get("output_dir", "./output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    if report_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(output_dir, f"batch_report_{batch_task['batch_id']}_{timestamp}.json")

    report_data = _build_json_report(batch_task)

    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report_data, f, ensure_ascii=False, indent=2)

    return report_path


def _build_text_report(batch_task):
    lines = []
    sep = "=" * 70

    lines.append(sep)
    lines.append("  JSON 批量转 Excel - 处理报告")
    lines.append(sep)
    lines.append("")

    lines.append(f"任务ID: {batch_task['batch_id']}")
    lines.append(f"任务状态: {_format_status(batch_task['status'])}")
    lines.append(f"创建时间: {_format_datetime(batch_task.get('created_at'))}")
    lines.append(f"开始时间: {_format_datetime(batch_task.get('started_at'))}")
    lines.append(f"完成时间: {_format_datetime(batch_task.get('completed_at'))}")

    if batch_task.get("started_at") and batch_task.get("completed_at"):
        duration = _calc_duration(batch_task["started_at"], batch_task["completed_at"])
        lines.append(f"总耗时: {duration}")

    lines.append("")
    lines.append(f"来源类型: {'目录' if batch_task['source_type'] == 'directory' else '文件列表'}")
    lines.append(f"输出目录: {batch_task['output_dir']}")
    lines.append("")

    config = batch_task.get("config", {})
    if config:
        lines.append(sep)
        lines.append("  导出配置")
        lines.append(sep)
        lines.append("")
        lines.append(f"导出格式: {config.get('export_format', 'excel')}")
        lines.append(f"自动检测字段: {'是' if config.get('auto_detect_headers', True) else '否'}")

        split_cfg = config.get("split_config", {})
        if split_cfg.get("enabled"):
            rule_label = {
                "by_value": "按字段值",
                "by_range": "按数值范围",
                "by_custom": "自定义规则",
            }.get(split_cfg.get("split_rule", "by_value"), split_cfg.get("split_rule", "by_value"))
            lines.append("")
            lines.append("工作表拆分: 已启用")
            lines.append(f"  拆分规则: {rule_label}")
            lines.append(f"  拆分字段: {split_cfg.get('split_field', '')}")
            lines.append(f"  命名模板: {split_cfg.get('sheet_name_template', '{value}')}")
            lines.append(f"  汇总工作表: {'是' if split_cfg.get('include_all_sheet', True) else '否'}")
            if split_cfg.get("include_all_sheet", True):
                lines.append(f"  汇总表名称: {split_cfg.get('all_sheet_name', '全部数据')}")
            lines.append(f"  未分类标签: {split_cfg.get('empty_value_label', '未分类')}")
        else:
            lines.append("工作表拆分: 未启用")

        lines.append("")

    lines.append(sep)
    lines.append("  处理统计")
    lines.append(sep)
    lines.append("")

    total = batch_task["total_files"]
    completed = batch_task["completed_count"]
    failed = batch_task["failed_count"]
    skipped = batch_task["skipped_count"]
    success_rate = (completed / total * 100) if total > 0 else 0

    lines.append(f"总文件数: {total}")
    lines.append(f"成功: {completed} 个")
    lines.append(f"失败: {failed} 个")
    lines.append(f"跳过: {skipped} 个")
    lines.append(f"成功率: {success_rate:.1f}%")
    lines.append("")

    total_data = sum(f.get("data_count", 0) for f in batch_task["files"])
    lines.append(f"导出数据总条数: {total_data}")
    lines.append("")

    lines.append(sep)
    lines.append("  文件明细")
    lines.append(sep)
    lines.append("")

    for idx, f in enumerate(batch_task["files"], 1):
        status_icon = _get_status_icon(f["status"])
        filename = os.path.basename(f["json_path"])
        lines.append(f"{idx:3d}. {status_icon} {filename}")
        lines.append(f"     JSON: {f['json_path']}")
        lines.append(f"     Excel: {f['excel_path']}")
        lines.append(f"     状态: {_format_status(f['status'])}")

        if f["status"] == "completed":
            lines.append(f"     数据条数: {f['data_count']}")
            if f.get("duration_seconds"):
                lines.append(f"     耗时: {f['duration_seconds']}秒")
        elif f["status"] == "failed":
            lines.append(f"     错误信息: {f.get('error', '未知错误')}")

        lines.append("")

    failed_files = [f for f in batch_task["files"] if f["status"] == "failed"]
    if failed_files:
        lines.append(sep)
        lines.append("  失败文件列表")
        lines.append(sep)
        lines.append("")
        for idx, f in enumerate(failed_files, 1):
            filename = os.path.basename(f["json_path"])
            lines.append(f"{idx}. {filename}")
            lines.append(f"   错误: {f.get('error', '未知错误')}")
            lines.append("")

    lines.append(sep)
    lines.append(f"报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(sep)

    return "\n".join(lines)


def _build_json_report(batch_task):
    failed_files = [
        {
            "json_path": f["json_path"],
            "excel_path": f["excel_path"],
            "error": f.get("error", ""),
        }
        for f in batch_task["files"]
        if f["status"] == "failed"
    ]

    completed_files = [
        {
            "json_path": f["json_path"],
            "excel_path": f["excel_path"],
            "data_count": f["data_count"],
            "duration_seconds": f.get("duration_seconds", 0),
        }
        for f in batch_task["files"]
        if f["status"] == "completed"
    ]

    total_data = sum(f.get("data_count", 0) for f in batch_task["files"])
    total_duration = sum(f.get("duration_seconds", 0) for f in batch_task["files"])

    config = batch_task.get("config", {})
    split_info = None
    if config:
        split_cfg = config.get("split_config", {})
        if split_cfg.get("enabled"):
            split_info = {
                "enabled": True,
                "split_field": split_cfg.get("split_field", ""),
                "split_rule": split_cfg.get("split_rule", "by_value"),
                "sheet_name_template": split_cfg.get("sheet_name_template", "{value}"),
                "include_all_sheet": split_cfg.get("include_all_sheet", True),
                "all_sheet_name": split_cfg.get("all_sheet_name", "全部数据"),
                "empty_value_label": split_cfg.get("empty_value_label", "未分类"),
            }
        else:
            split_info = {"enabled": False}

    return {
        "batch_id": batch_task["batch_id"],
        "status": batch_task["status"],
        "created_at": batch_task.get("created_at"),
        "started_at": batch_task.get("started_at"),
        "completed_at": batch_task.get("completed_at"),
        "source_type": batch_task["source_type"],
        "sources": batch_task["sources"],
        "output_dir": batch_task["output_dir"],
        "export_config": {
            "format": config.get("export_format", "excel") if config else "excel",
            "auto_detect_headers": config.get("auto_detect_headers", True) if config else True,
            "split_config": split_info,
        },
        "statistics": {
            "total_files": batch_task["total_files"],
            "completed": batch_task["completed_count"],
            "failed": batch_task["failed_count"],
            "skipped": batch_task["skipped_count"],
            "success_rate": round((batch_task["completed_count"] / batch_task["total_files"] * 100) if batch_task["total_files"] > 0 else 0, 1),
            "total_data_count": total_data,
            "total_duration_seconds": round(total_duration, 2),
        },
        "completed_files": completed_files,
        "failed_files": failed_files,
        "all_files": batch_task["files"],
    }


def _format_status(status):
    status_map = {
        "pending": "待处理",
        "running": "处理中",
        "completed": "成功",
        "failed": "失败",
        "skipped": "已跳过",
        "paused": "已暂停",
    }
    return status_map.get(status, status)


def _get_status_icon(status):
    icon_map = {
        "pending": "⏳",
        "running": "🔄",
        "completed": "✅",
        "failed": "❌",
        "skipped": "⏭️",
    }
    return icon_map.get(status, "❓")


def _format_datetime(dt_str):
    if not dt_str:
        return "未开始"
    try:
        dt = datetime.fromisoformat(dt_str)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return dt_str


def _calc_duration(start_str, end_str):
    try:
        start = datetime.fromisoformat(start_str)
        end = datetime.fromisoformat(end_str)
        delta = end - start
        total_seconds = delta.total_seconds()

        if total_seconds < 60:
            return f"{total_seconds:.1f} 秒"
        elif total_seconds < 3600:
            minutes = int(total_seconds // 60)
            seconds = total_seconds % 60
            return f"{minutes} 分 {seconds:.1f} 秒"
        else:
            hours = int(total_seconds // 3600)
            minutes = int((total_seconds % 3600) // 60)
            seconds = total_seconds % 60
            return f"{hours} 时 {minutes} 分 {seconds:.1f} 秒"
    except (ValueError, TypeError):
        return "计算失败"


def print_summary(batch_task):
    total = batch_task["total_files"]
    completed = batch_task["completed_count"]
    failed = batch_task["failed_count"]
    skipped = batch_task["skipped_count"]
    success_rate = (completed / total * 100) if total > 0 else 0

    print("\n" + "=" * 60)
    print("  批量处理摘要")
    print("=" * 60)
    print(f"  任务ID: {batch_task['batch_id']}")
    print(f"  总文件数: {total}")
    print(f"  成功: {completed} 个")
    print(f"  失败: {failed} 个")
    print(f"  跳过: {skipped} 个")
    print(f"  成功率: {success_rate:.1f}%")
    print("=" * 60)
