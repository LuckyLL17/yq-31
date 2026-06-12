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
        "description": "经典蓝色表头，灰色隔行",
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
                    style_header=True, style_alt_rows=True, overwrite=False):
    template = {
        "name": name,
        "description": description,
        "header_style": header_style,
        "data_style": data_style,
        "style_header": style_header,
        "style_alt_rows": style_alt_rows,
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

    return True, f"已应用模板: {template.get('name', template_id)}"


def create_template_from_config(template_id, name, description, config, overwrite=False):
    header_style = copy.deepcopy(config.get("header_style", {}))
    data_style = copy.deepcopy(config.get("data_style", {}))

    return create_template(
        template_id=template_id,
        name=name,
        description=description,
        header_style=header_style,
        data_style=data_style,
        style_header=config.get("style_header", True),
        style_alt_rows=config.get("style_alt_rows", True),
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
