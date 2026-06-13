import json
import os
import copy


DEFAULT_TEMPLATES_DIR = "./style_templates"

BORDER_STYLES = [
    "thin",
    "medium",
    "thick",
    "dashed",
    "dotted",
    "double",
    "hair",
    "mediumDashed",
    "dashDot",
    "mediumDashDot",
    "dashDotDot",
    "mediumDashDotDot",
    "slantDashDot",
]

BORDER_STYLE_NAMES = {
    "thin": "细实线",
    "medium": "中实线",
    "thick": "粗实线",
    "dashed": "虚线",
    "dotted": "点线",
    "double": "双线",
    "hair": "发丝线",
    "mediumDashed": "中虚线",
    "dashDot": "点划线",
    "mediumDashDot": "中点划线",
    "dashDotDot": "双点划线",
    "mediumDashDotDot": "中双点划线",
    "slantDashDot": "斜点划线",
}

FONT_FAMILIES = [
    "Arial",
    "Times New Roman",
    "Calibri",
    "SimSun",
    "Microsoft YaHei",
    "SimHei",
    "KaiTi",
    "FangSong",
    "Tahoma",
    "Verdana",
    "Georgia",
    "Courier New",
]

DEFAULT_TEMPLATES = {
    "default": {
        "name": "默认样式",
        "description": "经典蓝色表头，灰色隔行（含薪资标红、年龄标绿、关键词高亮条件格式）",
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
        "style_header": True,
        "style_alt_rows": True,
        "conditional_format_rules": [
            {
                "field": "salary",
                "rule_type": "numeric",
                "operator": "gt",
                "value": 20000,
                "style": {"font_color": "#FF0000", "font_bold": True},
                "enabled": True,
                "priority": 10,
                "description": "薪资大于20000标红",
            },
            {
                "field": "age",
                "rule_type": "numeric",
                "operator": "lt",
                "value": 30,
                "style": {"font_color": "#00B050", "font_bold": True},
                "enabled": True,
                "priority": 5,
                "description": "年龄小于30标绿",
            },
            {
                "field": "*",
                "rule_type": "text",
                "operator": "contains",
                "value": "高级",
                "style": {"bg_color": "#FFFF00", "font_bold": True},
                "enabled": True,
                "priority": 1,
                "description": "包含'高级'关键词高亮",
            },
        ],
    },
    "professional": {
        "name": "专业商务",
        "description": "深灰色表头，白色隔行",
        "header_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": True,
            "font_color": "#FFFFFF",
            "font_size": 12,
            "bg_color": "#595959",
            "alignment": "center",
            "border_style": "medium",
            "border_color": "#333333",
        },
        "data_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": False,
            "font_color": "#333333",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#FFFFFF",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": "thin",
            "border_color": "#BFBFBF",
        },
        "style_header": True,
        "style_alt_rows": False,
    },
    "fresh": {
        "name": "清新绿色",
        "description": "绿色表头，浅绿隔行",
        "header_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": True,
            "font_color": "#FFFFFF",
            "font_size": 12,
            "bg_color": "#70AD47",
            "alignment": "center",
            "border_style": "thin",
            "border_color": "#507E32",
        },
        "data_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": False,
            "font_color": "#333333",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#E2EFDA",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": "thin",
            "border_color": "#C6E0B4",
        },
        "style_header": True,
        "style_alt_rows": True,
    },
    "warm": {
        "name": "温暖橙色",
        "description": "橙色表头，浅橙隔行",
        "header_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": True,
            "font_color": "#FFFFFF",
            "font_size": 12,
            "bg_color": "#ED7D31",
            "alignment": "center",
            "border_style": "thin",
            "border_color": "#C0504D",
        },
        "data_style": {
            "font_name": "Microsoft YaHei",
            "font_bold": False,
            "font_color": "#333333",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#FCE4D6",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": "thin",
            "border_color": "#F8CBAD",
        },
        "style_header": True,
        "style_alt_rows": True,
    },
    "elegant": {
        "name": "优雅紫色",
        "description": "紫色表头，浅紫隔行",
        "header_style": {
            "font_name": "Times New Roman",
            "font_bold": True,
            "font_color": "#FFFFFF",
            "font_size": 12,
            "bg_color": "#7030A0",
            "alignment": "center",
            "border_style": "double",
            "border_color": "#4B1F74",
        },
        "data_style": {
            "font_name": "Times New Roman",
            "font_bold": False,
            "font_color": "#333333",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#E7E6F7",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": "thin",
            "border_color": "#C9C4E8",
        },
        "style_header": True,
        "style_alt_rows": True,
    },
    "minimal": {
        "name": "极简风格",
        "description": "无边框，仅隔行变色",
        "header_style": {
            "font_name": "Calibri",
            "font_bold": True,
            "font_color": "#333333",
            "font_size": 12,
            "bg_color": "#FFFFFF",
            "alignment": "left",
            "border_style": None,
            "border_color": "#000000",
        },
        "data_style": {
            "font_name": "Calibri",
            "font_bold": False,
            "font_color": "#333333",
            "font_size": 11,
            "wrap_text": True,
            "alt_row_color": "#F9F9F9",
            "alignment": "left",
            "vertical_alignment": "center",
            "border_style": None,
            "border_color": "#000000",
        },
        "style_header": True,
        "style_alt_rows": True,
    },
}


def get_templates_dir():
    return DEFAULT_TEMPLATES_DIR


def ensure_templates_dir():
    os.makedirs(DEFAULT_TEMPLATES_DIR, exist_ok=True)


def get_template_path(template_id):
    return os.path.join(DEFAULT_TEMPLATES_DIR, f"{template_id}.json")


def list_templates():
    ensure_templates_dir()
    templates = {}

    for template_id, template in DEFAULT_TEMPLATES.items():
        templates[template_id] = copy.deepcopy(template)

    for filename in os.listdir(DEFAULT_TEMPLATES_DIR):
        if filename.endswith(".json"):
            template_id = filename[:-5]
            try:
                template = load_template(template_id)
                if template:
                    templates[template_id] = template
            except Exception:
                pass

    return templates


def load_template(template_id):
    if template_id in DEFAULT_TEMPLATES:
        return copy.deepcopy(DEFAULT_TEMPLATES[template_id])

    path = get_template_path(template_id)
    if not os.path.exists(path):
        return None

    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return None


def save_template(template_id, template, overwrite=False):
    ensure_templates_dir()

    if template_id in DEFAULT_TEMPLATES and not overwrite:
        return False, "不能覆盖内置模板"

    path = get_template_path(template_id)

    if os.path.exists(path) and not overwrite:
        return False, f"模板 '{template_id}' 已存在"

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        return True, f"模板已保存到: {os.path.abspath(path)}"
    except IOError as e:
        return False, f"保存失败: {e}"


def delete_template(template_id):
    if template_id in DEFAULT_TEMPLATES:
        return False, "不能删除内置模板"

    path = get_template_path(template_id)
    if not os.path.exists(path):
        return False, f"模板 '{template_id}' 不存在"

    try:
        os.remove(path)
        return True, f"模板 '{template_id}' 已删除"
    except IOError as e:
        return False, f"删除失败: {e}"


def create_template(template_id, name, description, header_style, data_style,
                    style_header=True, style_alt_rows=True,
                    conditional_format_rules=None, overwrite=False):
    template = {
        "name": name,
        "description": description,
        "header_style": header_style,
        "data_style": data_style,
        "style_header": style_header,
        "style_alt_rows": style_alt_rows,
        "conditional_format_rules": conditional_format_rules or [],
    }
    return save_template(template_id, template, overwrite)


def apply_template_to_config(config, template_id):
    template = load_template(template_id)
    if not template:
        return False, f"模板 '{template_id}' 不存在"

    config["style_header"] = template.get("style_header", True)
    config["style_alt_rows"] = template.get("style_alt_rows", True)
    config["header_style"] = copy.deepcopy(template.get("header_style", {}))
    config["data_style"] = copy.deepcopy(template.get("data_style", {}))
    config["conditional_format_rules"] = copy.deepcopy(template.get("conditional_format_rules", []))

    return True, f"已应用模板: {template.get('name', template_id)}"


def create_template_from_config(template_id, name, description, config, overwrite=False):
    header_style = copy.deepcopy(config.get("header_style", {}))
    data_style = copy.deepcopy(config.get("data_style", {}))
    conditional_format_rules = copy.deepcopy(config.get("conditional_format_rules", []))

    return create_template(
        template_id=template_id,
        name=name,
        description=description,
        header_style=header_style,
        data_style=data_style,
        style_header=config.get("style_header", True),
        style_alt_rows=config.get("style_alt_rows", True),
        conditional_format_rules=conditional_format_rules,
        overwrite=overwrite,
    )


def get_border_style_name(style):
    return BORDER_STYLE_NAMES.get(style, style)


def validate_border_style(style):
    if style is None:
        return True
    return style in BORDER_STYLES


def validate_font_name(font_name):
    if font_name is None:
        return True
    return font_name in FONT_FAMILIES


def get_default_header_style():
    return copy.deepcopy(DEFAULT_TEMPLATES["default"]["header_style"])


def get_default_data_style():
    return copy.deepcopy(DEFAULT_TEMPLATES["default"]["data_style"])


def merge_style(base_style, override_style):
    merged = copy.deepcopy(base_style)
    for key, value in override_style.items():
        merged[key] = value
    return merged


# ==================== 条件格式规则 ====================

CF_RULE_TYPE_NUMERIC = "numeric"
CF_RULE_TYPE_TEXT = "text"
CF_RULE_TYPE_NULL = "null"
CF_RULE_TYPE_CUSTOM = "custom"

CF_RULE_TYPES = [
    CF_RULE_TYPE_NUMERIC,
    CF_RULE_TYPE_TEXT,
    CF_RULE_TYPE_NULL,
    CF_RULE_TYPE_CUSTOM,
]

CF_RULE_TYPE_LABELS = {
    CF_RULE_TYPE_NUMERIC: "数值比较",
    CF_RULE_TYPE_TEXT: "文本匹配",
    CF_RULE_TYPE_NULL: "空值判断",
    CF_RULE_TYPE_CUSTOM: "自定义表达式",
}

CF_OP_GT = "gt"
CF_OP_LT = "lt"
CF_OP_EQ = "eq"
CF_OP_NE = "ne"
CF_OP_GTE = "gte"
CF_OP_LTE = "lte"
CF_OP_BETWEEN = "between"
CF_OP_NOT_BETWEEN = "not_between"
CF_OP_CONTAINS = "contains"
CF_OP_NOT_CONTAINS = "not_contains"
CF_OP_STARTS_WITH = "starts_with"
CF_OP_ENDS_WITH = "ends_with"
CF_OP_IS_NULL = "is_null"
CF_OP_IS_NOT_NULL = "is_not_null"

CF_NUMERIC_OPERATORS = [
    CF_OP_GT, CF_OP_LT, CF_OP_EQ, CF_OP_NE,
    CF_OP_GTE, CF_OP_LTE, CF_OP_BETWEEN, CF_OP_NOT_BETWEEN,
]

CF_TEXT_OPERATORS = [
    CF_OP_CONTAINS, CF_OP_NOT_CONTAINS,
    CF_OP_STARTS_WITH, CF_OP_ENDS_WITH,
    CF_OP_EQ, CF_OP_NE,
]

CF_NULL_OPERATORS = [
    CF_OP_IS_NULL, CF_OP_IS_NOT_NULL,
]

CF_OPERATOR_LABELS = {
    CF_OP_GT: "大于",
    CF_OP_LT: "小于",
    CF_OP_EQ: "等于",
    CF_OP_NE: "不等于",
    CF_OP_GTE: "大于等于",
    CF_OP_LTE: "小于等于",
    CF_OP_BETWEEN: "介于之间",
    CF_OP_NOT_BETWEEN: "不介于之间",
    CF_OP_CONTAINS: "包含",
    CF_OP_NOT_CONTAINS: "不包含",
    CF_OP_STARTS_WITH: "开头是",
    CF_OP_ENDS_WITH: "结尾是",
    CF_OP_IS_NULL: "为空",
    CF_OP_IS_NOT_NULL: "不为空",
}

CF_STYLE_FIELDS = [
    "font_name", "font_size", "font_bold", "font_italic",
    "font_color", "bg_color", "underline",
]

CF_STYLE_FIELD_LABELS = {
    "font_name": "字体",
    "font_size": "字号",
    "font_bold": "加粗",
    "font_italic": "斜体",
    "font_color": "字体颜色",
    "bg_color": "背景颜色",
    "underline": "下划线",
}


def _to_numeric(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _to_string(value):
    if value is None:
        return ""
    return str(value)


def evaluate_conditional_rule(value, rule):
    """
    评估单元格值是否匹配条件格式规则

    Args:
        value: 单元格值
        rule: 条件格式规则字典

    Returns:
        bool: 是否匹配规则
    """
    if not rule.get("enabled", True):
        return False

    rule_type = rule.get("rule_type")
    operator = rule.get("operator")
    rule_value = rule.get("value")
    rule_value2 = rule.get("value2")

    if rule_type == CF_RULE_TYPE_NUMERIC:
        num_val = _to_numeric(value)
        if num_val is None:
            return False

        if operator == CF_OP_GT:
            return num_val > float(rule_value)
        elif operator == CF_OP_LT:
            return num_val < float(rule_value)
        elif operator == CF_OP_EQ:
            return num_val == float(rule_value)
        elif operator == CF_OP_NE:
            return num_val != float(rule_value)
        elif operator == CF_OP_GTE:
            return num_val >= float(rule_value)
        elif operator == CF_OP_LTE:
            return num_val <= float(rule_value)
        elif operator == CF_OP_BETWEEN:
            return float(rule_value) <= num_val <= float(rule_value2)
        elif operator == CF_OP_NOT_BETWEEN:
            return num_val < float(rule_value) or num_val > float(rule_value2)

    elif rule_type == CF_RULE_TYPE_TEXT:
        str_val = _to_string(value)
        rule_str = _to_string(rule_value)

        if operator == CF_OP_CONTAINS:
            return rule_str in str_val
        elif operator == CF_OP_NOT_CONTAINS:
            return rule_str not in str_val
        elif operator == CF_OP_STARTS_WITH:
            return str_val.startswith(rule_str)
        elif operator == CF_OP_ENDS_WITH:
            return str_val.endswith(rule_str)
        elif operator == CF_OP_EQ:
            return str_val == rule_str
        elif operator == CF_OP_NE:
            return str_val != rule_str

    elif rule_type == CF_RULE_TYPE_NULL:
        is_null = value is None or value == ""
        if operator == CF_OP_IS_NULL:
            return is_null
        elif operator == CF_OP_IS_NOT_NULL:
            return not is_null

    elif rule_type == CF_RULE_TYPE_CUSTOM:
        condition = rule.get("condition", "")
        try:
            return eval(condition, {"__builtins__": {}}, {"value": value})
        except Exception:
            return False

    return False


def match_conditional_rules_for_field(field_key, value, rules):
    """
    为指定字段值匹配所有适用的条件格式规则

    Args:
        field_key: 字段名
        value: 单元格值
        rules: 条件格式规则列表

    Returns:
        list: 匹配的规则列表（按优先级排序）
    """
    matched = []
    for rule in rules:
        if not rule.get("enabled", True):
            continue

        rule_field = rule.get("field", "*")
        if rule_field != "*" and rule_field != field_key:
            continue

        if evaluate_conditional_rule(value, rule):
            matched.append(rule)

    matched.sort(key=lambda r: r.get("priority", 0), reverse=True)
    return matched


def merge_conditional_styles(matched_rules):
    """
    合并多个匹配规则的样式（后匹配的覆盖先匹配的）

    Args:
        matched_rules: 匹配的规则列表

    Returns:
        dict: 合并后的样式字典
    """
    merged = {}
    for rule in matched_rules:
        style = rule.get("style", {})
        for key, value in style.items():
            if key in CF_STYLE_FIELDS:
                merged[key] = value
    return merged


def create_conditional_rule(field, rule_type, operator, value, style,
                            value2=None, enabled=True, priority=0, description=""):
    """
    创建条件格式规则

    Args:
        field: 字段名，"*" 表示所有字段
        rule_type: 规则类型
        operator: 操作符
        value: 比较值
        style: 样式字典
        value2: 第二个比较值（用于 between 等）
        enabled: 是否启用
        priority: 优先级（数字越大越先应用）
        description: 规则描述

    Returns:
        dict: 条件格式规则字典
    """
    return {
        "field": field,
        "rule_type": rule_type,
        "operator": operator,
        "value": value,
        "value2": value2,
        "style": style,
        "enabled": enabled,
        "priority": priority,
        "description": description,
    }


def get_rule_description(rule):
    """
    获取规则的人类可读描述

    Args:
        rule: 条件格式规则字典

    Returns:
        str: 规则描述
    """
    field = rule.get("field", "*")
    rule_type = rule.get("rule_type")
    operator = rule.get("operator")
    value = rule.get("value")
    value2 = rule.get("value2")
    description = rule.get("description", "")

    if description:
        return description

    type_label = CF_RULE_TYPE_LABELS.get(rule_type, rule_type)
    op_label = CF_OPERATOR_LABELS.get(operator, operator)
    field_label = "所有字段" if field == "*" else f"字段 '{field}'"

    if operator in (CF_OP_BETWEEN, CF_OP_NOT_BETWEEN):
        return f"{field_label} {type_label}: {op_label} {value} 和 {value2}"
    elif operator in (CF_OP_IS_NULL, CF_OP_IS_NOT_NULL):
        return f"{field_label} {op_label}"
    else:
        return f"{field_label} {type_label}: {op_label} {value}"


def preset_conditional_rules():
    """
    提供常用的预设条件格式规则模板

    Returns:
        dict: 预设规则模板
    """
    return {
        "high_salary_red": {
            "name": "高薪标红",
            "description": "薪资大于20000标红",
            "rule": create_conditional_rule(
                field="salary",
                rule_type=CF_RULE_TYPE_NUMERIC,
                operator=CF_OP_GT,
                value=20000,
                style={"font_color": "#FF0000", "font_bold": True},
                description="薪资大于20000标红",
            )
        },
        "young_age_green": {
            "name": "年轻标绿",
            "description": "年龄小于30标绿",
            "rule": create_conditional_rule(
                field="age",
                rule_type=CF_RULE_TYPE_NUMERIC,
                operator=CF_OP_LT,
                value=30,
                style={"font_color": "#00B050", "font_bold": True},
                description="年龄小于30标绿",
            )
        },
        "keyword_highlight": {
            "name": "关键词高亮",
            "description": "包含'高级'关键词高亮",
            "rule": create_conditional_rule(
                field="*",
                rule_type=CF_RULE_TYPE_TEXT,
                operator=CF_OP_CONTAINS,
                value="高级",
                style={"bg_color": "#FFFF00", "font_bold": True},
                description="包含'高级'关键词高亮",
            )
        },
        "empty_warning": {
            "name": "空值警告",
            "description": "空值标黄警告",
            "rule": create_conditional_rule(
                field="*",
                rule_type=CF_RULE_TYPE_NULL,
                operator=CF_OP_IS_NULL,
                value=None,
                style={"bg_color": "#FFC7CE", "font_color": "#9C0006"},
                description="空值标黄警告",
            )
        },
        "top_performer": {
            "name": "高绩效突出",
            "description": "绩效大于等于90加粗并蓝底",
            "rule": create_conditional_rule(
                field="performance",
                rule_type=CF_RULE_TYPE_NUMERIC,
                operator=CF_OP_GTE,
                value=90,
                style={"bg_color": "#4472C4", "font_color": "#FFFFFF", "font_bold": True},
                description="绩效大于等于90加粗并蓝底",
            )
        },
    }


def validate_conditional_rule(rule):
    """
    验证条件格式规则是否有效

    Args:
        rule: 条件格式规则字典

    Returns:
        tuple: (is_valid, error_message)
    """
    if "field" not in rule:
        return False, "缺少字段名"

    rule_type = rule.get("rule_type")
    if rule_type not in CF_RULE_TYPES:
        return False, f"无效的规则类型: {rule_type}"

    operator = rule.get("operator")
    if rule_type == CF_RULE_TYPE_NUMERIC:
        if operator not in CF_NUMERIC_OPERATORS:
            return False, f"数值规则无效的操作符: {operator}"
    elif rule_type == CF_RULE_TYPE_TEXT:
        if operator not in CF_TEXT_OPERATORS:
            return False, f"文本规则无效的操作符: {operator}"
    elif rule_type == CF_RULE_TYPE_NULL:
        if operator not in CF_NULL_OPERATORS:
            return False, f"空值规则无效的操作符: {operator}"

    style = rule.get("style", {})
    if not style:
        return False, "规则缺少样式定义"

    for key in style:
        if key not in CF_STYLE_FIELDS:
            return False, f"无效的样式字段: {key}"

    return True, "规则有效"

