import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime


CONFIG = {
    "json_file_path": "./data/sample_data.json",
    "excel_output_path": "./output/result.xlsx",
    "sheet_name": "数据导出",
    "default_headers": [
        {"key": "id", "label": "ID", "width": 10},
        {"key": "name", "label": "姓名", "width": 15},
        {"key": "age", "label": "年龄", "width": 10},
        {"key": "email", "label": "邮箱", "width": 30},
        {"key": "phone", "label": "电话", "width": 18},
        {"key": "address", "label": "地址", "width": 40},
        {"key": "department", "label": "部门", "width": 15},
        {"key": "position", "label": "职位", "width": 20},
        {"key": "salary", "label": "薪资", "width": 12},
        {"key": "join_date", "label": "入职日期", "width": 15},
    ],
    "auto_detect_headers": True,
    "style_header": True,
    "style_alt_rows": True,
}


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


def merge_headers(config_headers, auto_headers, data):
    if not config_headers:
        return auto_headers

    config_keys = {h["key"] for h in config_headers}
    extra_headers = [h for h in auto_headers if h["key"] not in config_keys]

    if CONFIG["auto_detect_headers"] and extra_headers:
        return config_headers + extra_headers
    return config_headers


def apply_header_style(ws, header_cells):
    font = Font(bold=True, color="FFFFFF", size=12)
    fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    alignment = Alignment(horizontal="center", vertical="center")
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


def apply_data_style(ws, headers, data_rows_count):
    alignment = Alignment(vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for row_idx in range(2, 2 + data_rows_count):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = alignment
            cell.border = thin_border
            if CONFIG["style_alt_rows"] and row_idx % 2 == 0:
                cell.fill = alt_fill


def set_column_widths(ws, headers):
    for idx, header in enumerate(headers, start=1):
        width = header.get("width", 15)
        ws.column_dimensions[chr(64 + idx)].width = width


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


def export_to_excel(data, headers, output_path, sheet_name):
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

    if CONFIG["style_header"]:
        header_cells = ws[1]
        apply_header_style(ws, header_cells)

    apply_data_style(ws, headers, len(data))

    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"Excel文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(data)} 条数据，{len(headers)} 个字段")


def main():
    print("=" * 50)
    print("JSON 转 Excel 工具")
    print("=" * 50)

    json_path = CONFIG["json_file_path"]
    excel_path = CONFIG["excel_output_path"]

    print(f"JSON文件路径: {json_path}")
    print(f"Excel输出路径: {excel_path}")
    print()

    try:
        data = load_json(json_path)
        print(f"已加载 {len(data)} 条数据")

        auto_headers = auto_detect_headers(data)
        print(f"自动检测到 {len(auto_headers)} 个字段")

        headers = merge_headers(CONFIG["default_headers"], auto_headers, data)
        print(f"最终使用 {len(headers)} 个字段")
        print()

        export_to_excel(
            data=data,
            headers=headers,
            output_path=excel_path,
            sheet_name=CONFIG["sheet_name"],
        )

    except FileNotFoundError as e:
        print(f"错误: {e}")
        print("请检查CONFIG中的json_file_path配置是否正确")
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
        print("请检查JSON文件格式是否正确")
    except Exception as e:
        print(f"发生错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
