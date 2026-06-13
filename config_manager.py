import json
import os
import copy


DEFAULT_CONFIG = {
    "json_file_path": "./data/sample_data.json",
    "export_format": "excel",
    "excel_output_path": "./output/result.xlsx",
    "csv_output_path": "./output/result.csv",
    "tsv_output_path": "./output/result.tsv",
    "html_output_path": "./output/result.html",
    "markdown_output_path": "./output/result.md",
    "json_output_path": "./output/result.json",
    "pdf_output_path": "./output/result.pdf",
    "sheet_name": "数据导出",
    "csv_config": {
        "encoding": "utf-8-sig",
        "delimiter": ",",
        "include_header": True,
        "quote_char": '"',
        "quoting": "minimal",
    },
    "tsv_config": {
        "encoding": "utf-8-sig",
        "include_header": True,
    },
    "html_config": {
        "title": "数据导出",
        "include_index": False,
        "pretty_print": True,
        "style": "default",
        "custom_css": "",
    },
    "markdown_config": {
        "title": "数据导出",
        "include_index": False,
        "max_col_width": 50,
    },
    "json_config": {
        "indent": 2,
        "ensure_ascii": False,
        "include_labels": False,
    },
    "pdf_config": {
        "title": "数据导出",
        "include_index": False,
        "page_size": "A4",
        "orientation": "portrait",
        "font_size": 10,
    },
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
        "font_name": "Microsoft YaHei",
        "font_bold": True,
        "font_color": "#FFFFFF",
        "font_size": 12,
        "bg_color": "#4472C4",
        "alignment": "center",
        "border_style": "thin",
        "border_color": "#000000",
    },
    "data_style": {
        "font_name": "Microsoft YaHei",
        "font_bold": False,
        "font_color": "#000000",
        "font_size": 11,
        "wrap_text": True,
        "alt_row_color": "#F2F2F2",
        "alignment": "left",
        "vertical_alignment": "center",
        "border_style": "thin",
        "border_color": "#000000",
    },
    "validation_rules": [],
    "validation_on_fail_default": "mark",
    "conditional_format_rules": [],
    "split_config": {
        "enabled": False,
        "split_field": "",
        "split_rule": "by_value",
        "sheet_name_template": "{value}",
        "include_all_sheet": True,
        "all_sheet_name": "全部数据",
        "empty_value_label": "未分类",
        "max_sheet_name_length": 31,
        "range_groups": [],
        "custom_rules": [],
    },
    "computed_columns": [],
    "pivot_config": {
        "enabled": False,
        "sheet_name": "数据透视表",
        "row_fields": [],
        "column_fields": [],
        "value_fields": [],
        "show_row_totals": True,
        "show_column_totals": True,
        "grand_total_label": "总计",
        "empty_value_label": "(空白)",
        "apply_style": True,
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

    export_format = config.get("export_format", "excel")
    valid_formats = {"excel", "csv", "tsv", "html", "markdown", "json", "pdf"}
    if export_format not in valid_formats:
        errors.append(f"导出格式无效，应为: {', '.join(sorted(valid_formats))}")

    output_path_key = f"{export_format}_output_path"
    if export_format == "excel":
        output_path_key = "excel_output_path"
    if not config.get(output_path_key):
        errors.append(f"{export_format.upper()}输出路径不能为空")

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

    validation_rules = config.get("validation_rules", [])
    if not isinstance(validation_rules, list):
        errors.append("校验规则格式错误，应为列表")
    else:
        valid_types = {"not_null", "format", "range", "regex"}
        valid_actions = {"mark", "skip", "abort"}
        for i, rule in enumerate(validation_rules):
            if not isinstance(rule, dict):
                errors.append(f"第 {i + 1} 条校验规则格式错误")
                continue
            if "field" not in rule or not rule["field"]:
                errors.append(f"第 {i + 1} 条校验规则缺少 field 属性")
            if "rule_type" not in rule or rule["rule_type"] not in valid_types:
                errors.append(f"第 {i + 1} 条校验规则类型无效，应为 {', '.join(valid_types)}")
            on_fail = rule.get("on_fail", "mark")
            if on_fail not in valid_actions:
                errors.append(f"第 {i + 1} 条校验规则处理方式无效，应为 {', '.join(valid_actions)}")

    split_config = config.get("split_config", {})
    if split_config.get("enabled"):
        if not split_config.get("split_field"):
            errors.append("启用拆分时必须指定 split_field（拆分字段）")

        valid_split_rules = {"by_value", "by_range", "by_custom"}
        split_rule = split_config.get("split_rule", "by_value")
        if split_rule not in valid_split_rules:
            errors.append(f"split_rule 无效，应为 {', '.join(sorted(valid_split_rules))}")

        if split_rule == "by_range":
            range_groups = split_config.get("range_groups", [])
            if not isinstance(range_groups, list) or len(range_groups) == 0:
                errors.append("使用 by_range 拆分时必须配置 range_groups")
            else:
                for gi, group in enumerate(range_groups):
                    if not isinstance(group, dict) or "name" not in group:
                        errors.append(f"第 {gi + 1} 个 range_group 缺少 name 属性")

        if split_rule == "by_custom":
            custom_rules = split_config.get("custom_rules", [])
            if not isinstance(custom_rules, list) or len(custom_rules) == 0:
                errors.append("使用 by_custom 拆分时必须配置 custom_rules")
            else:
                for ci, rule in enumerate(custom_rules):
                    if not isinstance(rule, dict):
                        errors.append(f"第 {ci + 1} 个 custom_rule 格式错误，应为字典")
                        continue
                    if "name" not in rule:
                        errors.append(f"第 {ci + 1} 个 custom_rule 缺少 name 属性")
                    if "values" not in rule and "condition" not in rule and "min" not in rule:
                        errors.append(f"第 {ci + 1} 个 custom_rule 缺少匹配条件（values/condition/min-max）")

    conditional_format_rules = config.get("conditional_format_rules", [])
    if not isinstance(conditional_format_rules, list):
        errors.append("条件格式规则格式错误，应为列表")
    else:
        from style_template_manager import (
            validate_conditional_rule,
            CF_RULE_TYPES,
            CF_STYLE_FIELDS,
        )
        for i, rule in enumerate(conditional_format_rules):
            if not isinstance(rule, dict):
                errors.append(f"第 {i + 1} 条条件格式规则格式错误")
                continue
            is_valid, rule_error = validate_conditional_rule(rule)
            if not is_valid:
                errors.append(f"第 {i + 1} 条条件格式规则错误: {rule_error}")

    computed_columns = config.get("computed_columns", [])
    if not isinstance(computed_columns, list):
        errors.append("计算列配置格式错误，应为列表")
    else:
        from computed_columns import validate_computed_column
        for i, cc in enumerate(computed_columns):
            if not isinstance(cc, dict):
                errors.append(f"第 {i + 1} 个计算列配置格式错误，应为字典")
                continue
            cc_errors = validate_computed_column(cc, i)
            errors.extend(cc_errors)

    pivot_config = config.get("pivot_config", {})
    if pivot_config.get("enabled"):
        valid_aggregates = {
            "sum", "count", "average", "max", "min",
            "product", "count_num", "stddev", "stddevp", "var", "varp"
        }

        row_fields = pivot_config.get("row_fields", [])
        if not isinstance(row_fields, list):
            errors.append("透视表 row_fields 格式错误，应为列表")
        elif len(row_fields) == 0:
            errors.append("启用透视表时至少需要配置一个行字段 (row_fields)")

        column_fields = pivot_config.get("column_fields", [])
        if not isinstance(column_fields, list):
            errors.append("透视表 column_fields 格式错误，应为列表")

        value_fields = pivot_config.get("value_fields", [])
        if not isinstance(value_fields, list):
            errors.append("透视表 value_fields 格式错误，应为列表")
        elif len(value_fields) == 0:
            errors.append("启用透视表时至少需要配置一个值字段 (value_fields)")
        else:
            for vi, vf in enumerate(value_fields):
                if not isinstance(vf, dict):
                    errors.append(f"第 {vi + 1} 个 value_field 格式错误，应为字典")
                    continue
                if "field" not in vf or not vf["field"]:
                    errors.append(f"第 {vi + 1} 个 value_field 缺少 field 属性")
                agg = vf.get("aggregate", "sum")
                if agg not in valid_aggregates:
                    errors.append(
                        f"第 {vi + 1} 个 value_field 的 aggregate 无效，"
                        f"应为: {', '.join(sorted(valid_aggregates))}"
                    )

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
