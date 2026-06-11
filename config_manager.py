import json
import os
import copy


DEFAULT_CONFIG = {
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
    "header_style": {
        "font_bold": True,
        "font_color": "#FFFFFF",
        "font_size": 12,
        "bg_color": "#4472C4",
        "alignment": "center",
    },
    "data_style": {
        "wrap_text": True,
        "alt_row_color": "#F2F2F2",
    },
}

CONFIG_FILE_PATH = "./config.json"


def get_default_config():
    return copy.deepcopy(DEFAULT_CONFIG)


def load_config(config_path=None):
    path = config_path or CONFIG_FILE_PATH

    if not os.path.exists(path):
        return get_default_config()

    try:
        with open(path, "r", encoding="utf-8") as f:
            config = json.load(f)
        return merge_config(get_default_config(), config)
    except (json.JSONDecodeError, IOError) as e:
        print(f"警告: 配置文件加载失败，使用默认配置 - {e}")
        return get_default_config()


def save_config(config, config_path=None):
    path = config_path or CONFIG_FILE_PATH
    output_dir = os.path.dirname(path)

    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    return os.path.abspath(path)


def merge_config(base_config, override_config):
    merged = copy.deepcopy(base_config)
    for key, value in override_config.items():
        if isinstance(value, dict) and key in merged and isinstance(merged[key], dict):
            merged[key] = merge_config(merged[key], value)
        else:
            merged[key] = value
    return merged


def validate_config(config):
    errors = []

    if not config.get("json_file_path"):
        errors.append("JSON文件路径不能为空")

    if not config.get("excel_output_path"):
        errors.append("Excel输出路径不能为空")

    headers = config.get("default_headers", [])
    if not isinstance(headers, list):
        errors.append("字段配置格式错误，应为列表")
    else:
        for i, header in enumerate(headers):
            if not isinstance(header, dict):
                errors.append(f"第 {i + 1} 个字段配置格式错误")
                continue
            if "key" not in header:
                errors.append(f"第 {i + 1} 个字段缺少 key 属性")

    return errors


def get_header_by_key(config, key):
    for header in config.get("default_headers", []):
        if header.get("key") == key:
            return header
    return None


def update_header_width(config, key, width):
    header = get_header_by_key(config, key)
    if header:
        header["width"] = width
        return True
    return False


def update_header_label(config, key, label):
    header = get_header_by_key(config, key)
    if header:
        header["label"] = label
        return True
    return False


def add_header(config, key, label, width=15, position=None):
    new_header = {"key": key, "label": label, "width": width}
    headers = config.setdefault("default_headers", [])

    if position is None or position >= len(headers):
        headers.append(new_header)
    else:
        headers.insert(max(0, position), new_header)

    return new_header


def remove_header(config, key):
    headers = config.get("default_headers", [])
    for i, header in enumerate(headers):
        if header.get("key") == key:
            headers.pop(i)
            return True
    return False


def reorder_headers(config, new_order_keys):
    headers = config.get("default_headers", [])
    key_to_header = {h["key"]: h for h in headers}
    new_headers = []

    for key in new_order_keys:
        if key in key_to_header:
            new_headers.append(key_to_header[key])
            del key_to_header[key]

    for header in headers:
        if header["key"] in key_to_header:
            new_headers.append(header)

    config["default_headers"] = new_headers
    return new_headers


def validate_file_path(path, must_exist=True, file_type=None):
    if must_exist and not os.path.exists(path):
        return False, f"文件不存在: {path}"
    if must_exist and not os.path.isfile(path):
        return False, f"路径不是文件: {path}"
    if file_type and not path.lower().endswith(file_type.lower()):
        return False, f"文件类型不正确，应为 {file_type} 格式"
    return True, None
