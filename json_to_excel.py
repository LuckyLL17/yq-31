import json
import os
import sys
import copy
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from datetime import datetime

from config_manager import (
    load_config,
    save_config,
    get_default_config,
    validate_config,
)
from data_validator import (
    ValidationRule,
    validate_data,
    apply_validation_to_export,
    rules_from_config,
    ON_FAIL_MARK,
    ON_FAIL_ABORT,
    MARK_COLOR,
    MARK_BG_COLOR,
    ON_FAIL_LABELS,
)


def load_json(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"JSON文件不存在: {file_path}")

    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, dict):
        if "data" in data and isinstance(data["data"], list):
            data = data["data"]
        elif "items" in data and isinstance(data["items"], list):
            data = data["items"]
        elif "records" in data and isinstance(data["records"], list):
            data = data["records"]
        else:
            data = [data]
    elif not isinstance(data, list):
        data = [data]

    return data


def flatten_dict(d, parent_key="", sep="."):
    items = []
    if isinstance(d, dict):
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                items.append((new_key, json.dumps(v, ensure_ascii=False)))
            else:
                items.append((new_key, v))
    else:
        items.append((parent_key, d))
    return dict(items)


def auto_detect_headers(data):
    headers = []
    all_keys = set()
    for item in data:
        if isinstance(item, dict):
            flat = flatten_dict(item)
            all_keys.update(flat.keys())

    for key in sorted(all_keys):
        headers.append({
            "key": key,
            "label": key.replace(".", " / ").replace("_", " ").title(),
            "width": max(15, len(key) * 2 + 5),
        })
    return headers


def merge_headers(config_headers, auto_headers, config):
    if not config_headers:
        return auto_headers

    config_keys = {h["key"] for h in config_headers}
    extra_headers = [h for h in auto_headers if h["key"] not in config_keys]

    if config.get("auto_detect_headers", True) and extra_headers:
        return config_headers + extra_headers
    return config_headers


def _get_alignment(align_str):
    align_map = {
        "left": "left",
        "center": "center",
        "right": "right",
    }
    return align_map.get(align_str, "center")


def _create_border(style_config):
    border_style = style_config.get("border_style", "thin")
    border_color = style_config.get("border_color", "000000").replace("#", "")

    if border_style is None:
        return None

    side = Side(style=border_style, color=border_color)
    return Border(left=side, right=side, top=side, bottom=side)


def apply_header_style(ws, header_cells, config):
    header_style = config.get("header_style", {})

    font_name = header_style.get("font_name", "Microsoft YaHei")
    font_bold = header_style.get("font_bold", True)
    font_color = header_style.get("font_color", "FFFFFF").replace("#", "")
    font_size = header_style.get("font_size", 12)
    bg_color = header_style.get("bg_color", "4472C4").replace("#", "")
    alignment = _get_alignment(header_style.get("alignment", "center"))
    vertical_alignment = header_style.get("vertical_alignment", "center")

    font = Font(name=font_name, bold=font_bold, color=font_color, size=font_size)
    fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    alignment = Alignment(horizontal=alignment, vertical=vertical_alignment)
    border = _create_border(header_style)

    for cell in header_cells:
        cell.font = font
        cell.fill = fill
        cell.alignment = alignment
        if border:
            cell.border = border


def apply_data_style(ws, headers, data_rows_count, config):
    data_style = config.get("data_style", {})
    font_name = data_style.get("font_name", "Microsoft YaHei")
    font_bold = data_style.get("font_bold", False)
    font_color = data_style.get("font_color", "000000").replace("#", "")
    font_size = data_style.get("font_size", 11)
    wrap_text = data_style.get("wrap_text", True)
    alt_row_color = data_style.get("alt_row_color", "F2F2F2").replace("#", "")
    alignment = _get_alignment(data_style.get("alignment", "left"))
    vertical_alignment = data_style.get("vertical_alignment", "center")

    font = Font(name=font_name, bold=font_bold, color=font_color, size=font_size)
    alignment = Alignment(horizontal=alignment, vertical=vertical_alignment, wrap_text=wrap_text)
    border = _create_border(data_style)
    alt_fill = PatternFill(start_color=alt_row_color, end_color=alt_row_color, fill_type="solid")

    for row_idx in range(2, 2 + data_rows_count):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = font
            cell.alignment = alignment
            if border:
                cell.border = border
            if config.get("style_alt_rows", True) and row_idx % 2 == 0:
                cell.fill = alt_fill


def set_column_widths(ws, headers):
    for idx, header in enumerate(headers, start=1):
        width = header.get("width", 15)
        col_letter = chr(64 + idx) if idx <= 26 else _get_column_letter(idx)
        ws.column_dimensions[col_letter].width = width


def _get_column_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def extract_value(item, key):
    flat = flatten_dict(item)
    if key in flat:
        value = flat[key]
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False)
        return value

    if "." in key:
        parts = key.split(".")
        current = item
        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                return ""
        if isinstance(current, (dict, list)):
            return json.dumps(current, ensure_ascii=False)
        return current

    return ""


def export_to_excel(data, headers, config):
    output_path = config["excel_output_path"]
    sheet_name = config.get("sheet_name", "数据导出")

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    rules = rules_from_config(config)
    validation_result = None
    original_data = data
    original_indices = list(range(len(data)))

    if rules:
        print(f"\n正在执行数据校验（{len(rules)} 条规则）...")
        validation_result = validate_data(data, rules)
        print(validation_result.summary())

        if validation_result.aborted:
            print(f"\n❌ 校验中止，导出已取消: {validation_result.abort_reason}")
            _print_validation_errors(validation_result, data, max_display=10)
            return None

        if validation_result.skipped_rows:
            print(f"  将跳过 {len(validation_result.skipped_rows)} 行校验不通过的数据")

        valid_data, removed = apply_validation_to_export(data, validation_result, headers)
        if valid_data is None:
            print("\n❌ 校验中止，导出已取消")
            return None

        new_indices = []
        for idx in range(len(data)):
            if idx not in validation_result.skipped_rows:
                new_indices.append(idx)
        original_indices = new_indices
        data = valid_data
        print(f"  校验后有效数据: {len(data)} 条\n")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_labels = [h["label"] for h in headers]
    ws.append(header_labels)

    for item in data:
        if not isinstance(item, dict):
            continue
        row = [extract_value(item, h["key"]) for h in headers]
        ws.append(row)

    set_column_widths(ws, headers)

    if config.get("style_header", True):
        header_cells = ws[1]
        apply_header_style(ws, header_cells, config)

    apply_data_style(ws, headers, len(data), config)

    if validation_result and validation_result.marked_rows:
        _apply_validation_marks(ws, validation_result, headers, original_indices)

    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"Excel文件已成功导出: {os.path.abspath(output_path)}")
    if validation_result:
        skipped_count = len(validation_result.skipped_rows)
        marked_count = len(validation_result.marked_rows)
        print(f"共导出 {len(data)} 条数据，{len(headers)} 个字段", end="")
        if skipped_count > 0:
            print(f"，跳过 {skipped_count} 条", end="")
        if marked_count > 0:
            print(f"，标记 {marked_count} 条", end="")
        print()
    else:
        print(f"共导出 {len(data)} 条数据，{len(headers)} 个字段")
    return output_path


def _apply_validation_marks(ws, validation_result, headers, original_indices):
    mark_font = Font(color=MARK_COLOR, bold=True)
    mark_fill = PatternFill(start_color=MARK_BG_COLOR, end_color=MARK_BG_COLOR, fill_type="solid")

    header_key_to_col = {}
    for col_idx, header in enumerate(headers, start=1):
        header_key_to_col[header["key"]] = col_idx

    marked_in_export = set()
    for orig_idx in validation_result.marked_rows:
        if orig_idx in original_indices:
            export_row = original_indices.index(orig_idx) + 2
            marked_in_export.add(export_row)

            row_errors = validation_result.get_field_errors_for_row(orig_idx)
            error_msgs = []
            for field_key, errors in row_errors.items():
                for e in errors:
                    error_msgs.append(e.rule.message)

            comment_text = "; ".join(error_msgs)
            if len(comment_text) > 200:
                comment_text = comment_text[:200] + "..."

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=export_row, column=col_idx)
                cell.font = mark_font
                cell.fill = mark_fill

            first_cell = ws.cell(row=export_row, column=1)
            first_cell.comment = Comment(comment_text, "数据校验")

            for field_key, errors in row_errors.items():
                if field_key in header_key_to_col:
                    col_idx = header_key_to_col[field_key]
                    cell = ws.cell(row=export_row, column=col_idx)
                    cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")


def _print_validation_errors(validation_result, data, max_display=10):
    errors = validation_result.errors[:max_display]
    if not errors:
        return
    print("\n校验错误详情:")
    for e in errors:
        val_str = str(e.value)[:30] if e.value is not None else "(空)"
        on_fail_str = ON_FAIL_LABELS.get(e.rule.on_fail, "").split("（")[0]
        print(f"  行 {e.row_index + 1}: {e.rule.message} | 值: {val_str} | 处理: {on_fail_str}")
    if len(validation_result.errors) > max_display:
        print(f"  ... 还有 {len(validation_result.errors) - max_display} 个问题未显示")


def parse_args():
    parser = argparse.ArgumentParser(description="JSON 数据导出工具（支持 Excel/CSV/TSV/HTML/Markdown/JSON/PDF）")
    parser.add_argument(
        "-w", "--wizard",
        action="store_true",
        help="使用交互式配置向导",
    )
    parser.add_argument(
        "-p", "--preview",
        action="store_true",
        help="进入数据预览与筛选模式",
    )
    parser.add_argument(
        "-c", "--config",
        type=str,
        default=None,
        help="指定配置文件路径（默认: ./config.json）",
    )
    parser.add_argument(
        "-i", "--input",
        type=str,
        default=None,
        help="指定输入 JSON 文件路径",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="指定输出文件路径",
    )
    parser.add_argument(
        "-f", "--format",
        type=str,
        default=None,
        choices=["excel", "csv", "tsv", "html", "markdown", "json", "pdf"],
        help="指定导出格式 (excel/csv/tsv/html/markdown/json/pdf)",
    )
    parser.add_argument(
        "--list-formats",
        action="store_true",
        help="列出所有支持的导出格式",
    )
    parser.add_argument(
        "--save-config",
        type=str,
        default=None,
        help="将当前配置保存到指定文件",
    )
    parser.add_argument(
        "--list-templates",
        action="store_true",
        help="列出所有可用的样式模板",
    )
    parser.add_argument(
        "--apply-template",
        type=str,
        default=None,
        help="应用指定的样式模板",
    )
    parser.add_argument(
        "--template-manager",
        action="store_true",
        help="启动样式模板管理器",
    )
    parser.add_argument(
        "--save-as-template",
        type=str,
        default=None,
        help="将当前样式保存为模板（指定模板ID）",
    )
    parser.add_argument(
        "--template-name",
        type=str,
        default=None,
        help="保存模板时的名称（与 --save-as-template 配合使用）",
    )
    parser.add_argument(
        "--template-desc",
        type=str,
        default="",
        help="保存模板时的描述（与 --save-as-template 配合使用）",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="使用批量处理模式（交互式向导）",
    )
    parser.add_argument(
        "--batch-dir",
        type=str,
        default=None,
        help="批量处理：指定包含JSON文件的目录",
    )
    parser.add_argument(
        "--batch-files",
        type=str,
        default=None,
        help="批量处理：指定多个JSON文件，用逗号分隔",
    )
    parser.add_argument(
        "--batch-output",
        type=str,
        default=None,
        help="批量处理：指定输出目录",
    )
    parser.add_argument(
        "--batch-resume",
        type=str,
        default=None,
        help="批量处理：恢复指定ID的未完成任务",
    )
    parser.add_argument(
        "--batch-list",
        action="store_true",
        help="批量处理：列出所有历史任务",
    )
    parser.add_argument(
        "--batch-report",
        type=str,
        default=None,
        help="批量处理：生成指定任务ID的处理报告",
    )
    parser.add_argument(
        "--batch-report-format",
        type=str,
        default="txt",
        choices=["txt", "json", "both"],
        help="批量处理：报告格式 (txt/json/both)，默认txt",
    )
    parser.add_argument(
        "--batch-skip-existing",
        action="store_true",
        help="批量处理：跳过已存在的Excel文件",
    )
    parser.add_argument(
        "--validate",
        action="store_true",
        help="启动交互式数据校验规则配置",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="仅执行数据校验，不导出Excel",
    )
    return parser.parse_args()


def apply_cli_overrides(config, args):
    from multi_exporter import get_format_extension

    if args.input:
        config["json_file_path"] = args.input
    if args.format:
        config["export_format"] = args.format
    if args.output:
        fmt = config.get("export_format", "excel")
        if fmt == "excel":
            config["excel_output_path"] = args.output
        else:
            config[f"{fmt}_output_path"] = args.output
    return config


def run_with_config(config, data=None):
    from multi_exporter import export_data, EXPORT_FORMATS

    export_fmt = config.get("export_format", "excel")
    fmt_info = EXPORT_FORMATS.get(export_fmt, {})
    output_key = f"{export_fmt}_output_path" if export_fmt != "excel" else "excel_output_path"

    print("=" * 50)
    print(f"JSON 数据导出工具 - {fmt_info.get('label', export_fmt.upper())}")
    print("=" * 50)

    json_path = config["json_file_path"]
    output_path = config.get(output_key, "")

    print(f"JSON文件路径: {json_path}")
    print(f"输出格式: {fmt_info.get('label', export_fmt.upper())}")
    print(f"输出路径: {output_path}")
    print()

    try:
        if data is None:
            data = load_json(json_path)
            print(f"已加载 {len(data)} 条数据")
        else:
            print(f"使用筛选后的数据，共 {len(data)} 条")

        auto_headers = auto_detect_headers(data)
        print(f"自动检测到 {len(auto_headers)} 个字段")

        headers = merge_headers(config.get("default_headers", []), auto_headers, config)
        print(f"最终使用 {len(headers)} 个字段")
        print()

        export_data(
            data=data,
            headers=headers,
            config=config,
            fmt=export_fmt,
        )

        if export_fmt == "excel":
            prompt_save_as_template(config)

    except FileNotFoundError as e:
        print(f"错误: {e}")
        print("请检查JSON文件路径配置是否正确")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
        print("请检查JSON文件格式是否正确")
        sys.exit(1)
    except Exception as e:
        print(f"发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def prompt_save_as_template(config):
    from prompts import prompt_confirm, prompt_input
    from style_template_manager import create_template_from_config

    if not prompt_confirm("\n是否将当前样式保存为自定义模板？", default=False):
        return

    template_id = prompt_input("请输入模板ID (英文标识，如 'my_style')", required=True)
    name = prompt_input("请输入模板名称", required=True)
    description = prompt_input("请输入模板描述", required=False, default="")

    success, msg = create_template_from_config(
        template_id, name, description, config, overwrite=False
    )

    if not success and "已存在" in msg:
        if prompt_confirm(f"模板 '{template_id}' 已存在，是否覆盖？", default=False):
            success, msg = create_template_from_config(
                template_id, name, description, config, overwrite=True
            )

    print(f"\n{msg}")


def handle_template_operations(args, config):
    from style_template_manager import (
        list_templates,
        apply_template_to_config,
        create_template_from_config,
        get_border_style_name,
    )

    if args.list_templates:
        templates = list_templates()
        print("=" * 70)
        print("可用样式模板")
        print("=" * 70)
        for template_id, template in templates.items():
            is_builtin = template_id in ("default", "professional", "fresh", "warm", "elegant", "minimal")
            builtin_tag = " [内置]" if is_builtin else " [自定义]"
            header_bg = template.get("header_style", {}).get("bg_color", "#FFFFFF")
            alt_color = template.get("data_style", {}).get("alt_row_color", "#FFFFFF")
            border_style = template.get("header_style", {}).get("border_style", "thin")
            border_name = get_border_style_name(border_style) if border_style else "无"
            print(f"\n{template_id}{builtin_tag}")
            print(f"  名称: {template.get('name', '未命名')}")
            print(f"  描述: {template.get('description', '无描述')}")
            print(f"  表头背景: {header_bg} | 隔行色: {alt_color} | 边框: {border_name}")
        return True

    if args.template_manager:
        from style_template_wizard import run_style_template_manager
        run_style_template_manager(config)
        return True

    if args.apply_template:
        success, msg = apply_template_to_config(config, args.apply_template)
        print(msg)
        if not success:
            sys.exit(1)
        return False

    if args.save_as_template:
        template_id = args.save_as_template
        name = args.template_name or template_id
        description = args.template_desc
        success, msg = create_template_from_config(
            template_id, name, description, config, overwrite=False
        )
        print(msg)
        if not success:
            sys.exit(1)
        return True

    return False


def _handle_batch_operations(args):
    if args.batch:
        from batch_wizard import run_batch_wizard
        config = load_config(args.config) if args.config else get_default_config()
        run_batch_wizard(config)
        return True

    if args.batch_list:
        from batch_wizard import show_task_manager
        show_task_manager()
        return True

    if args.batch_resume:
        from batch_processor import resume_batch_process
        from report_generator import generate_report, generate_json_report, print_summary

        batch_id = args.batch_resume
        batch_task, result = resume_batch_process(batch_id)

        if batch_task is None:
            print(f"错误: {result}")
            sys.exit(1)

        fmt = args.batch_report_format
        if fmt in ("txt", "both"):
            report_path = generate_report(batch_task)
            print(f"\n📄 文本报告已生成: {os.path.abspath(report_path)}")
        if fmt in ("json", "both"):
            report_path = generate_json_report(batch_task)
            print(f"📄 JSON报告已生成: {os.path.abspath(report_path)}")

        print_summary(batch_task)
        return True

    if args.batch_report:
        from task_manager import load_batch_task
        from report_generator import generate_report, generate_json_report, print_summary

        batch_id = args.batch_report
        batch_task = load_batch_task(batch_id)
        if not batch_task:
            print(f"错误: 任务 {batch_id} 不存在")
            sys.exit(1)

        fmt = args.batch_report_format
        if fmt in ("txt", "both"):
            report_path = generate_report(batch_task)
            print(f"📄 文本报告已生成: {os.path.abspath(report_path)}")
        if fmt in ("json", "both"):
            report_path = generate_json_report(batch_task)
            print(f"📄 JSON报告已生成: {os.path.abspath(report_path)}")

        print_summary(batch_task)
        return True

    if args.batch_dir or args.batch_files:
        from batch_processor import start_batch_process
        from report_generator import generate_report, generate_json_report, print_summary

        config = load_config(args.config)
        config = apply_cli_overrides(config, args)

        output_dir = args.batch_output or "./output/batch"

        if args.batch_dir:
            source_type = "directory"
            sources = [args.batch_dir]
        else:
            source_type = "files"
            sources = [f.strip() for f in args.batch_files.split(",") if f.strip()]

        batch_task, result = start_batch_process(
            source_type=source_type,
            sources=sources,
            output_dir=output_dir,
            config=config,
            options={
                "generate_report": True,
                "report_format": args.batch_report_format,
                "skip_existing": args.batch_skip_existing,
            },
        )

        if batch_task is None:
            print(f"错误: {result}")
            sys.exit(1)

        fmt = args.batch_report_format
        if fmt in ("txt", "both"):
            report_path = generate_report(batch_task)
            print(f"\n📄 文本报告已生成: {os.path.abspath(report_path)}")
        if fmt in ("json", "both"):
            report_path = generate_json_report(batch_task)
            print(f"📄 JSON报告已生成: {os.path.abspath(report_path)}")

        print_summary(batch_task)
        return True

    return False


def main():
    args = parse_args()

    if args.list_formats:
        from multi_exporter import EXPORT_FORMATS
        print("=" * 70)
        print("支持的导出格式")
        print("=" * 70)
        for fmt_id, fmt_info in EXPORT_FORMATS.items():
            print(f"\n  {fmt_id:10s} - {fmt_info['label']}")
            print(f"              {fmt_info['description']}")
        print("\n使用 -f/--format 参数指定格式，例如: python json_to_excel.py -f csv")
        return

    if _handle_batch_operations(args):
        return

    if args.validate:
        from validation_wizard import run_validation_wizard
        config = load_config(args.config) if args.config else get_default_config()
        try:
            data = load_json(config["json_file_path"])
            auto_headers = auto_detect_headers(data)
            config = run_validation_wizard(config, available_fields=auto_headers)
        except Exception:
            config = run_validation_wizard(config)
        if args.save_config:
            save_config(config, args.save_config)
            print(f"\n配置已保存到: {os.path.abspath(args.save_config)}")
        return

    if args.validate_only:
        config = load_config(args.config) if args.config else get_default_config()
        config = apply_cli_overrides(config, args)
        rules = rules_from_config(config)
        if not rules:
            print("未配置任何校验规则，请先使用 --validate 配置规则")
            return
        try:
            data = load_json(config["json_file_path"])
            print(f"已加载 {len(data)} 条数据，正在执行校验...")
            result = validate_data(data, rules)
            print(f"\n{result.summary()}")
            _print_validation_errors(result, data)
        except Exception as e:
            print(f"校验执行失败: {e}")
        return

    if args.wizard:
        from wizard import run_wizard
        config = load_config(args.config) if args.config else get_default_config()
        config = run_wizard(config)

        errors = validate_config(config)
        if errors:
            print("\n❌ 配置验证失败:")
            for e in errors:
                print(f"  • {e}")
            sys.exit(1)

        if args.save_config:
            save_config(config, args.save_config)
            print(f"\n配置已保存到: {os.path.abspath(args.save_config)}")

        if prompt_confirm_execute():
            filtered_data = config.pop("_filtered_data", None)
            run_with_config(config, data=filtered_data)
        return

    config = load_config(args.config)
    config = apply_cli_overrides(config, args)

    template_exit = handle_template_operations(args, config)
    if template_exit and not args.apply_template:
        return

    if args.apply_template:
        if args.save_config:
            save_config(config, args.save_config)
            print(f"配置已保存到: {os.path.abspath(args.save_config)}")
        if prompt_confirm_execute():
            run_with_config(config)
        return

    if args.preview:
        from data_previewer import start_preview_mode
        from multi_exporter import export_data
        result = start_preview_mode(config)
        if result is not None and prompt_confirm_execute_after_preview(config):
            export_config = copy.deepcopy(config)
            auto_headers = auto_detect_headers(result)
            headers = merge_headers(config.get("default_headers", []), auto_headers, config)
            export_fmt = config.get("export_format", "excel")
            export_data(data=result, headers=headers, config=export_config, fmt=export_fmt)
            if export_fmt == "excel":
                prompt_save_as_template(export_config)
        return

    errors = validate_config(config)
    if errors:
        print("❌ 配置验证失败:")
        for e in errors:
            print(f"  • {e}")
        sys.exit(1)

    if args.save_config:
        save_config(config, args.save_config)
        print(f"配置已保存到: {os.path.abspath(args.save_config)}")

    if prompt_confirm_preview():
        from data_previewer import start_preview_mode
        from multi_exporter import export_data
        result = start_preview_mode(config)
        if result is not None and prompt_confirm_execute_after_preview(config):
            export_config = copy.deepcopy(config)
            auto_headers = auto_detect_headers(result)
            headers = merge_headers(config.get("default_headers", []), auto_headers, config)
            export_fmt = config.get("export_format", "excel")
            export_data(data=result, headers=headers, config=export_config, fmt=export_fmt)
            if export_fmt == "excel":
                prompt_save_as_template(export_config)
        return

    run_with_config(config)


def prompt_confirm_preview():
    while True:
        user_input = input("\n是否先预览并筛选数据？ (Y/n): ").strip().lower()
        if not user_input or user_input in ("y", "yes"):
            return True
        if user_input in ("n", "no"):
            return False
        print("  ⚠️  请输入 y 或 n")


def prompt_confirm_execute():
    while True:
        user_input = input("\n配置完成！是否立即执行导出？ (Y/n): ").strip().lower()
        if not user_input or user_input in ("y", "yes"):
            return True
        if user_input in ("n", "no"):
            print("配置已保存，您可以稍后运行 `python json_to_excel.py` 执行导出。")
            return False
        print("  ⚠️  请输入 y 或 n")


def prompt_confirm_execute_after_preview(config=None):
    from multi_exporter import EXPORT_FORMATS
    export_fmt = "excel"
    if config:
        export_fmt = config.get("export_format", "excel")
    fmt_label = EXPORT_FORMATS.get(export_fmt, {}).get("label", export_fmt.upper())

    while True:
        user_input = input(f"\n预览完成！是否导出筛选后的数据到 {fmt_label}？ (Y/n): ").strip().lower()
        if not user_input or user_input in ("y", "yes"):
            return True
        if user_input in ("n", "no"):
            return False
        print("  ⚠️  请输入 y 或 n")


if __name__ == "__main__":
    main()
