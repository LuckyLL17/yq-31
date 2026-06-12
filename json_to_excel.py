import json
import os
import sys
import copy
import argparse
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

from config_manager import (
    load_config,
    save_config,
    get_default_config,
    validate_config,
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


def apply_header_style(ws, header_cells, config):
    header_style = config.get("header_style", {})

    font_bold = header_style.get("font_bold", True)
    font_color = header_style.get("font_color", "FFFFFF").replace("#", "")
    font_size = header_style.get("font_size", 12)
    bg_color = header_style.get("bg_color", "4472C4").replace("#", "")
    alignment = _get_alignment(header_style.get("alignment", "center"))

    font = Font(bold=font_bold, color=font_color, size=font_size)
    fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    alignment = Alignment(horizontal=alignment, vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for cell in header_cells:
        cell.font = font
        cell.fill = fill
        cell.alignment = alignment
        cell.border = thin_border


def apply_data_style(ws, headers, data_rows_count, config):
    data_style = config.get("data_style", {})
    wrap_text = data_style.get("wrap_text", True)
    alt_row_color = data_style.get("alt_row_color", "F2F2F2").replace("#", "")

    alignment = Alignment(vertical="center", wrap_text=wrap_text)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    alt_fill = PatternFill(start_color=alt_row_color, end_color=alt_row_color, fill_type="solid")

    for row_idx in range(2, 2 + data_rows_count):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = alignment
            cell.border = thin_border
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

    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"Excel文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(data)} 条数据，{len(headers)} 个字段")


def parse_args():
    parser = argparse.ArgumentParser(description="JSON 转 Excel 工具")
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
        help="指定输出 Excel 文件路径",
    )
    parser.add_argument(
        "--save-config",
        type=str,
        default=None,
        help="将当前配置保存到指定文件",
    )
    return parser.parse_args()


def apply_cli_overrides(config, args):
    if args.input:
        config["json_file_path"] = args.input
    if args.output:
        config["excel_output_path"] = args.output
    return config


def run_with_config(config, data=None):
    print("=" * 50)
    print("JSON 转 Excel 工具")
    print("=" * 50)

    json_path = config["json_file_path"]
    excel_path = config["excel_output_path"]

    print(f"JSON文件路径: {json_path}")
    print(f"Excel输出路径: {excel_path}")
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

        export_to_excel(
            data=data,
            headers=headers,
            config=config,
        )

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


def main():
    args = parse_args()

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

    if args.preview:
        from data_previewer import start_preview_mode
        result = start_preview_mode(config)
        if result is not None and prompt_confirm_execute_after_preview():
            export_config = copy.deepcopy(config)
            auto_headers = auto_detect_headers(result)
            headers = merge_headers(config.get("default_headers", []), auto_headers, config)
            export_to_excel(data=result, headers=headers, config=export_config)
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
        result = start_preview_mode(config)
        if result is not None and prompt_confirm_execute_after_preview():
            export_config = copy.deepcopy(config)
            auto_headers = auto_detect_headers(result)
            headers = merge_headers(config.get("default_headers", []), auto_headers, config)
            export_to_excel(data=result, headers=headers, config=export_config)
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


def prompt_confirm_execute_after_preview():
    while True:
        user_input = input("\n预览完成！是否导出筛选后的数据到 Excel？ (Y/n): ").strip().lower()
        if not user_input or user_input in ("y", "yes"):
            return True
        if user_input in ("n", "no"):
            return False
        print("  ⚠️  请输入 y 或 n")


if __name__ == "__main__":
    main()
