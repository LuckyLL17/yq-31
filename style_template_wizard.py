import copy
import re

from prompts import (
    clear_screen,
    print_title,
    prompt_input,
    prompt_confirm,
    prompt_choice,
    prompt_number,
    prompt_color,
)
from style_template_manager import (
    list_templates,
    load_template,
    save_template,
    delete_template,
    create_template_from_config,
    apply_template_to_config,
    get_border_style_name,
    BORDER_STYLES,
    BORDER_STYLE_NAMES,
    FONT_FAMILIES,
    get_default_header_style,
    get_default_data_style,
)


def print_templates_list(templates):
    print("\n可用样式模板:")
    print("-" * 70)
    for i, (template_id, template) in enumerate(templates.items(), 1):
        is_builtin = template_id in ("default", "professional", "fresh", "warm", "elegant", "minimal")
        builtin_tag = " [内置]" if is_builtin else " [自定义]"
        header_bg = template.get("header_style", {}).get("bg_color", "#FFFFFF")
        alt_color = template.get("data_style", {}).get("alt_row_color", "#FFFFFF")
        border_style = template.get("header_style", {}).get("border_style", "thin")
        border_name = get_border_style_name(border_style) if border_style else "无"
        print(f"  {i:2d}. {template_id:15s} {builtin_tag:8s}")
        print(f"       名称: {template.get('name', '未命名')}")
        print(f"       描述: {template.get('description', '无描述')}")
        print(f"       表头背景: {header_bg} | 隔行色: {alt_color} | 边框: {border_name}")
        print()


def prompt_font_family(message, default="Microsoft YaHei"):
    print(f"\n{message}")
    for i, font in enumerate(FONT_FAMILIES, 1):
        print(f"  {i:2d}. {font}")

    while True:
        user_input = input(f"请选择 (1-{len(FONT_FAMILIES)}, 回车默认 {default}): ").strip()

        if not user_input:
            return default

        try:
            index = int(user_input) - 1
            if 0 <= index < len(FONT_FAMILIES):
                return FONT_FAMILIES[index]
        except ValueError:
            pass

        if user_input in FONT_FAMILIES:
            return user_input

        print(f"  ⚠️  请输入 1 到 {len(FONT_FAMILIES)} 之间的数字，或有效的字体名称")


def prompt_border_style(message, default="thin"):
    print(f"\n{message}")
    print("  0. 无边框")
    for i, style in enumerate(BORDER_STYLES, 1):
        name = BORDER_STYLE_NAMES.get(style, style)
        print(f"  {i:2d}. {style} ({name})")

    while True:
        default_index = BORDER_STYLES.index(default) + 1 if default in BORDER_STYLES else 0
        user_input = input(f"请选择 (0-{len(BORDER_STYLES)}, 回车默认 {default_index}): ").strip()

        if not user_input:
            return default

        try:
            index = int(user_input)
            if index == 0:
                return None
            if 1 <= index <= len(BORDER_STYLES):
                return BORDER_STYLES[index - 1]
        except ValueError:
            pass

        if user_input in BORDER_STYLES:
            return user_input
        if user_input.lower() in ("none", "null", "无"):
            return None

        print(f"  ⚠️  请输入 0 到 {len(BORDER_STYLES)} 之间的数字，或有效的边框样式")


def interactive_edit_style(style, style_name):
    style = copy.deepcopy(style)

    while True:
        clear_screen()
        print_title(f"编辑{style_name}样式")
        print(f"\n当前{style_name}样式:")
        print("-" * 50)
        for key, value in style.items():
            display_value = value
            if key == "border_style" and value is None:
                display_value = "无"
            elif key == "border_style":
                display_value = f"{value} ({get_border_style_name(value)})"
            print(f"  {key:20s}: {display_value}")

        print("\n操作选项:")
        print(f"  1. 修改字体 ({style.get('font_name', 'N/A')})")
        print(f"  2. 切换加粗 ({'是' if style.get('font_bold') else '否'})")
        print(f"  3. 修改字体颜色 ({style.get('font_color', 'N/A')})")
        print(f"  4. 修改字体大小 ({style.get('font_size', 'N/A')})")
        if "bg_color" in style:
            print(f"  5. 修改背景颜色 ({style.get('bg_color', 'N/A')})")
        print(f"  6. 修改水平对齐 ({style.get('alignment', 'N/A')})")
        print(f"  7. 修改垂直对齐 ({style.get('vertical_alignment', 'N/A')})")
        print(f"  8. 修改边框样式 ({style.get('border_style', 'N/A')})")
        print(f"  9. 修改边框颜色 ({style.get('border_color', 'N/A')})")
        if "wrap_text" in style:
            print(f"  10. 切换自动换行 ({'是' if style.get('wrap_text') else '否'})")
        if "alt_row_color" in style:
            print(f"  11. 修改隔行填充色 ({style.get('alt_row_color', 'N/A')})")
        print("  0. 完成编辑")

        choice = prompt_number("\n请选择操作", min_value=0, max_value=11, default=0)

        if choice == 0:
            break
        elif choice == 1:
            style["font_name"] = prompt_font_family("选择字体", default=style.get("font_name", "Microsoft YaHei"))
        elif choice == 2:
            style["font_bold"] = not style.get("font_bold", False)
        elif choice == 3:
            style["font_color"] = prompt_color("请输入字体颜色", default=style.get("font_color", "#000000"))
        elif choice == 4:
            style["font_size"] = prompt_number("请输入字体大小", min_value=8, max_value=36, default=style.get("font_size", 11))
        elif choice == 5 and "bg_color" in style:
            style["bg_color"] = prompt_color("请输入背景颜色", default=style.get("bg_color", "#FFFFFF"))
        elif choice == 6:
            style["alignment"] = prompt_choice(
                "选择水平对齐方式",
                [("left", "左对齐"), ("center", "居中对齐"), ("right", "右对齐")],
                default_index=0 if style.get("alignment") == "left" else (1 if style.get("alignment") == "center" else 2),
            )
        elif choice == 7:
            style["vertical_alignment"] = prompt_choice(
                "选择垂直对齐方式",
                [("top", "顶部对齐"), ("center", "居中对齐"), ("bottom", "底部对齐")],
                default_index=1,
            )
        elif choice == 8:
            style["border_style"] = prompt_border_style("选择边框样式", default=style.get("border_style", "thin"))
        elif choice == 9:
            style["border_color"] = prompt_color("请输入边框颜色", default=style.get("border_color", "#000000"))
        elif choice == 10 and "wrap_text" in style:
            style["wrap_text"] = not style.get("wrap_text", True)
        elif choice == 11 and "alt_row_color" in style:
            style["alt_row_color"] = prompt_color("请输入隔行填充色", default=style.get("alt_row_color", "#F2F2F2"))

    return style


def interactive_create_template():
    clear_screen()
    print_title("创建新样式模板")

    template_id = prompt_input("请输入模板ID (英文标识，如 'my_style')", required=True)
    name = prompt_input("请输入模板名称", required=True)
    description = prompt_input("请输入模板描述", required=False, default="")

    print("\n选择起始样式:")
    start_choice = prompt_choice(
        "请选择起始样式模板",
        [
            ("default", "默认样式"),
            ("professional", "专业商务"),
            ("fresh", "清新绿色"),
            ("warm", "温暖橙色"),
            ("elegant", "优雅紫色"),
            ("minimal", "极简风格"),
            ("blank", "空白样式"),
        ],
        default_index=0,
    )

    if start_choice == "blank":
        header_style = get_default_header_style()
        data_style = get_default_data_style()
    else:
        template = load_template(start_choice)
        header_style = template.get("header_style", get_default_header_style())
        data_style = template.get("data_style", get_default_data_style())

    if prompt_confirm("\n是否编辑表头样式？", default=True):
        header_style = interactive_edit_style(header_style, "表头")

    if prompt_confirm("\n是否编辑数据行样式？", default=True):
        data_style = interactive_edit_style(data_style, "数据行")

    style_header = prompt_confirm("是否启用表头样式？", default=True)
    style_alt_rows = prompt_confirm("是否启用隔行变色？", default=True)

    overwrite = False
    success, msg = save_template(
        template_id,
        {
            "name": name,
            "description": description,
            "header_style": header_style,
            "data_style": data_style,
            "style_header": style_header,
            "style_alt_rows": style_alt_rows,
        },
        overwrite=False,
    )

    if not success and "已存在" in msg:
        if prompt_confirm(f"模板 '{template_id}' 已存在，是否覆盖？", default=False):
            success, msg = save_template(
                template_id,
                {
                    "name": name,
                    "description": description,
                    "header_style": header_style,
                    "data_style": data_style,
                    "style_header": style_header,
                    "style_alt_rows": style_alt_rows,
                },
                overwrite=True,
            )

    print(f"\n{msg}")
    input("\n按回车继续...")
    return success


def interactive_edit_template():
    templates = list_templates()
    custom_templates = {k: v for k, v in templates.items()
                        if k not in ("default", "professional", "fresh", "warm", "elegant", "minimal")}

    if not custom_templates:
        print("\n⚠️  没有可编辑的自定义模板")
        input("\n按回车继续...")
        return False

    clear_screen()
    print_title("编辑样式模板")
    print_templates_list(custom_templates)

    template_ids = list(custom_templates.keys())
    idx = prompt_number(f"\n请选择要编辑的模板 (1-{len(template_ids)})", min_value=1, max_value=len(template_ids)) - 1
    template_id = template_ids[idx]
    template = load_template(template_id)

    name = prompt_input("请输入模板名称", default=template.get("name", ""))
    description = prompt_input("请输入模板描述", default=template.get("description", ""))

    header_style = template.get("header_style", get_default_header_style())
    data_style = template.get("data_style", get_default_data_style())

    if prompt_confirm("\n是否编辑表头样式？", default=True):
        header_style = interactive_edit_style(header_style, "表头")

    if prompt_confirm("\n是否编辑数据行样式？", default=True):
        data_style = interactive_edit_style(data_style, "数据行")

    style_header = prompt_confirm("是否启用表头样式？", default=template.get("style_header", True))
    style_alt_rows = prompt_confirm("是否启用隔行变色？", default=template.get("style_alt_rows", True))

    success, msg = save_template(
        template_id,
        {
            "name": name,
            "description": description,
            "header_style": header_style,
            "data_style": data_style,
            "style_header": style_header,
            "style_alt_rows": style_alt_rows,
        },
        overwrite=True,
    )

    print(f"\n{msg}")
    input("\n按回车继续...")
    return success


def interactive_delete_template():
    templates = list_templates()
    custom_templates = {k: v for k, v in templates.items()
                        if k not in ("default", "professional", "fresh", "warm", "elegant", "minimal")}

    if not custom_templates:
        print("\n⚠️  没有可删除的自定义模板")
        input("\n按回车继续...")
        return False

    clear_screen()
    print_title("删除样式模板")
    print_templates_list(custom_templates)

    template_ids = list(custom_templates.keys())
    idx = prompt_number(f"\n请选择要删除的模板 (1-{len(template_ids)})", min_value=1, max_value=len(template_ids)) - 1
    template_id = template_ids[idx]
    template = load_template(template_id)

    if prompt_confirm(f"确定要删除模板 '{template.get('name', template_id)}' 吗？此操作不可恢复。", default=False):
        success, msg = delete_template(template_id)
        print(f"\n{msg}")
        input("\n按回车继续...")
        return success

    print("\n已取消删除")
    input("\n按回车继续...")
    return False


def interactive_apply_template(config):
    templates = list_templates()

    clear_screen()
    print_title("应用样式模板")
    print_templates_list(templates)

    template_ids = list(templates.keys())
    idx = prompt_number(f"\n请选择要应用的模板 (1-{len(template_ids)})", min_value=1, max_value=len(template_ids)) - 1
    template_id = template_ids[idx]

    success, msg = apply_template_to_config(config, template_id)
    print(f"\n{msg}")
    input("\n按回车继续...")
    return success


def interactive_save_as_template(config):
    clear_screen()
    print_title("保存当前样式为模板")

    template_id = prompt_input("请输入模板ID (英文标识)", required=True)
    name = prompt_input("请输入模板名称", required=True)
    description = prompt_input("请输入模板描述", required=False, default="")

    success, msg = create_template_from_config(template_id, name, description, config, overwrite=False)

    if not success and "已存在" in msg:
        if prompt_confirm(f"模板 '{template_id}' 已存在，是否覆盖？", default=False):
            success, msg = create_template_from_config(template_id, name, description, config, overwrite=True)

    print(f"\n{msg}")
    input("\n按回车继续...")
    return success


def interactive_list_templates():
    templates = list_templates()
    clear_screen()
    print_title("样式模板列表")
    print_templates_list(templates)
    input("\n按回车继续...")


def run_style_template_manager(config=None):
    config = config or {}

    while True:
        clear_screen()
        print_title("Excel 样式模板管理")

        print("\n操作选项:")
        print("  1. 查看所有模板")
        print("  2. 创建新模板")
        print("  3. 编辑自定义模板")
        print("  4. 删除自定义模板")
        if config:
            print("  5. 应用模板到当前配置")
            print("  6. 保存当前样式为模板")
        print("  0. 退出")

        max_choice = 6 if config else 4
        choice = prompt_number("\n请选择操作", min_value=0, max_value=max_choice, default=0)

        if choice == 0:
            break
        elif choice == 1:
            interactive_list_templates()
        elif choice == 2:
            interactive_create_template()
        elif choice == 3:
            interactive_edit_template()
        elif choice == 4:
            interactive_delete_template()
        elif choice == 5 and config:
            interactive_apply_template(config)
        elif choice == 6 and config:
            interactive_save_as_template(config)

    return config


if __name__ == "__main__":
    run_style_template_manager()
