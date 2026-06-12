import json
import os

from prompts import (
    clear_screen,
    print_title,
    print_step,
    prompt_input,
    prompt_confirm,
    prompt_choice,
    prompt_multi_select,
    prompt_number,
    prompt_file_path,
    prompt_color,
    format_current_config,
)
from config_manager import (
    load_config,
    save_config,
    validate_config,
    get_default_config,
    update_header_width,
    update_header_label,
    add_header,
    remove_header,
    reorder_headers,
)
from style_template_manager import (
    apply_template_to_config,
    create_template_from_config,
    list_templates,
    get_border_style_name,
    BORDER_STYLES,
    BORDER_STYLE_NAMES,
    FONT_FAMILIES,
)

from json_to_excel import load_json, auto_detect_headers, flatten_dict
from multi_exporter import EXPORT_FORMATS, get_default_output_path


TOTAL_STEPS = 9


class ConfigWizard:
    def __init__(self, config=None):
        self.config = config or get_default_config()
        self.available_fields = []
        self.current_step = 0
        self.total_steps = 0

    def _next_step(self):
        self.current_step += 1

    def run(self):
        while True:
            clear_screen()
            print_title("JSON 数据导出 - 交互式配置向导")
            print("\n欢迎使用配置向导！我们将引导您完成所有设置。\n")

            self.current_step = 0
            should_load_config = prompt_confirm("是否加载已有配置文件？", default=True)
            self.total_steps = 9 + (1 if should_load_config else 0)

            if should_load_config:
                self._next_step()
                self.step_load_config()

            self._next_step()
            self.step_select_json_file()

            self._next_step()
            self.step_detect_fields()

            self._next_step()
            self.step_select_fields()

            self._next_step()
            self.step_configure_column_widths()

            self._next_step()
            self.step_configure_validation()

            self._next_step()
            self.step_configure_styles()

            self._next_step()
            self.step_configure_split()

            self._next_step()
            self.step_select_export_format()

            self._next_step()
            if self.step_review_and_confirm():
                break

        return self.config

    def step_load_config(self):
        print_step(self.current_step, self.total_steps, "加载配置文件")
        config_path = prompt_file_path(
            "请输入配置文件路径",
            default="./config.json",
            file_filter=".json",
            must_exist=True,
        )
        self.config = load_config(config_path)
        print(f"\n✅ 已加载配置文件: {os.path.abspath(config_path)}")
        input("\n按回车继续...")

    def step_select_json_file(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "选择 JSON 文件")

        default_path = self.config.get("json_file_path", "./data/sample_data.json")
        json_path = prompt_file_path(
            "请输入 JSON 数据文件路径",
            default=default_path,
            file_filter=".json",
            must_exist=True,
        )
        self.config["json_file_path"] = json_path
        print(f"\n✅ JSON 文件已设置: {os.path.abspath(json_path)}")
        input("\n按回车继续...")

    def step_detect_fields(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "检测数据字段")

        print("\n正在分析 JSON 数据结构...")
        try:
            data = load_json(self.config["json_file_path"])
            auto_headers = auto_detect_headers(data)
            self.available_fields = auto_headers
            self.sample_data = data[:3] if len(data) > 3 else data

            print(f"\n✅ 成功检测到 {len(auto_headers)} 个字段:")
            for h in auto_headers:
                sample_val = self._get_sample_value(h["key"])
                print(f"  • {h['key']:25s} - {h['label']:20s} 示例: {sample_val}")

            if prompt_confirm("\n是否需要手动添加未检测到的字段？", default=False):
                self._manual_add_fields()

            self._merge_config_headers_to_available()

        except Exception as e:
            print(f"\n❌ 检测字段时出错: {e}")
            if not prompt_confirm("是否继续使用现有配置？", default=True):
                raise

        input("\n按回车继续...")

    def step_select_fields(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "选择需要导出的字段")

        if not self.available_fields:
            print("\n⚠️  没有可用字段，使用默认配置")
            input("\n按回车继续...")
            return

        config_header_keys = {h["key"] for h in self.config.get("default_headers", [])}
        options = []
        default_selected_indices = []
        for idx, h in enumerate(self.available_fields):
            options.append((h["key"], f"{h['label']} ({h['key']})"))
            if h["key"] in config_header_keys:
                default_selected_indices.append(idx)

        selected_keys = prompt_multi_select(
            "请选择要导出的字段",
            options,
            min_selected=1,
            default_indices=default_selected_indices,
        )

        selected_headers = []
        for key in selected_keys:
            auto_h = next((h for h in self.available_fields if h["key"] == key), None)
            if auto_h:
                config_h = next(
                    (h for h in self.config.get("default_headers", []) if h["key"] == key),
                    None
                )
                selected_headers.append({
                    "key": key,
                    "label": config_h["label"] if config_h else auto_h["label"],
                    "width": config_h["width"] if config_h else auto_h["width"],
                })

        self.config["default_headers"] = selected_headers
        self.config["auto_detect_headers"] = prompt_confirm(
            "\n是否自动检测并添加 JSON 中的新字段？",
            default=self.config.get("auto_detect_headers", True),
        )

        print(f"\n✅ 已选择 {len(selected_headers)} 个字段")
        input("\n按回车继续...")

    def step_configure_column_widths(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "调整列宽和标签")

        headers = self.config.get("default_headers", [])
        if not headers:
            print("\n⚠️  没有可配置的字段")
            input("\n按回车继续...")
            return

        while True:
            print(f"\n当前字段配置（共 {len(headers)} 个）:")
            for i, h in enumerate(headers, 1):
                print(f"  {i:2d}. {h['label']:20s} | key: {h['key']:25s} | 宽度: {h['width']:3d}")

            print("\n操作选项:")
            print("  1. 修改字段标签")
            print("  2. 调整列宽")
            print("  3. 调整字段顺序")
            print("  4. 移除字段")
            print("  5. 全部使用自动宽度")
            print("  6. 完成配置，继续下一步")

            choice = prompt_choice(
                "\n请选择操作",
                [
                    ("label", "修改字段标签"),
                    ("width", "调整列宽"),
                    ("order", "调整字段顺序"),
                    ("remove", "移除字段"),
                    ("auto_width", "全部使用自动宽度"),
                    ("done", "完成配置，继续下一步"),
                ],
                default_index=5,
            )

            if choice == "done":
                break
            elif choice == "label":
                self._modify_label(headers)
            elif choice == "width":
                self._modify_width(headers)
            elif choice == "order":
                self._reorder_fields(headers)
            elif choice == "remove":
                self._remove_field(headers)
            elif choice == "auto_width":
                self._auto_width_all(headers)

            if choice != "done":
                input("\n按回车继续...")
                clear_screen()
                print_step(self.current_step, self.total_steps, "调整列宽和标签")

        print("\n✅ 字段配置完成")
        input("\n按回车继续...")

    def _prompt_font_family(self, message, default="Microsoft YaHei"):
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

    def _prompt_border_style(self, message, default="thin"):
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

    def _apply_template_interactive(self):
        templates = list_templates()
        print("\n可用样式模板:")
        template_list = []
        for i, (template_id, template) in enumerate(templates.items(), 1):
            is_builtin = template_id in ("default", "professional", "fresh", "warm", "elegant", "minimal")
            builtin_tag = " [内置]" if is_builtin else " [自定义]"
            print(f"  {i:2d}. {template_id:15s} {builtin_tag} - {template.get('name', '未命名')}")
            template_list.append(template_id)

        print(f"  0. 取消")

        choice = prompt_number(
            "请选择要应用的模板",
            min_value=0,
            max_value=len(template_list),
            default=0,
        )

        if choice == 0:
            return

        template_id = template_list[choice - 1]
        success, msg = apply_template_to_config(self.config, template_id)
        print(f"\n{msg}")

    def _save_as_template_interactive(self):
        template_id = prompt_input("请输入模板ID (英文标识)", required=True)
        name = prompt_input("请输入模板名称", required=True)
        description = prompt_input("请输入模板描述", required=False, default="")

        success, msg = create_template_from_config(
            template_id, name, description, self.config, overwrite=False
        )

        if not success and "已存在" in msg:
            if prompt_confirm(f"模板 '{template_id}' 已存在，是否覆盖？", default=False):
                success, msg = create_template_from_config(
                    template_id, name, description, self.config, overwrite=True
                )

        print(f"\n{msg}")

    def step_configure_validation(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置数据校验规则")

        existing_rules = self.config.get("validation_rules", [])
        if existing_rules:
            print(f"\n当前已配置 {len(existing_rules)} 条校验规则:")
            for i, rule in enumerate(existing_rules, 1):
                from data_validator import VALIDATION_TYPE_LABELS, ON_FAIL_LABELS
                type_label = VALIDATION_TYPE_LABELS.get(rule.get("rule_type", ""), rule.get("rule_type", ""))
                on_fail_label = ON_FAIL_LABELS.get(rule.get("on_fail", ""), "").split("（")[0]
                print(f"  {i}. [{type_label}] {rule.get('field', '')} → {on_fail_label}: {rule.get('message', '')}")
        else:
            print("\n当前未配置任何校验规则")

        if prompt_confirm("\n是否配置数据校验规则？", default=False):
            from validation_wizard import run_validation_wizard
            self.config = run_validation_wizard(self.config, available_fields=self.available_fields)

        print("\n✅ 校验规则配置完成")
        input("\n按回车继续...")

    def step_configure_styles(self):
        clear_screen()
        export_fmt = self.config.get("export_format", "excel")

        if export_fmt == "excel":
            print_step(self.current_step, self.total_steps, "配置 Excel 导出样式")
            self._config_excel_styles()
        else:
            print_step(self.current_step, self.total_steps, "配置通用选项")
            self._config_general_title()

        print("\n✅ 配置完成")
        input("\n按回车继续...")

    def step_configure_split(self):
        clear_screen()
        export_fmt = self.config.get("export_format", "excel")

        if export_fmt != "excel":
            print_step(self.current_step, self.total_steps, "配置 Excel 工作表拆分（跳过）")
            print("\n⚠️  工作表拆分功能仅适用于 Excel 导出格式，当前格式已跳过此步骤。")
            self.config.setdefault("split_config", {})["enabled"] = False
            input("\n按回车继续...")
            return

        print_step(self.current_step, self.total_steps, "配置 Excel 工作表拆分")

        split_cfg = self.config.setdefault("split_config", {})

        print("\n工作表拆分可以将数据按指定字段的值分组，分别导出到不同的工作表。")
        print("例如：按 '部门' 字段拆分，每个部门自动生成一个独立的工作表。")

        if not prompt_confirm("\n是否启用工作表拆分功能？", default=split_cfg.get("enabled", False)):
            split_cfg["enabled"] = False
            print("\n✅ 已禁用工作表拆分")
            input("\n按回车继续...")
            return

        split_cfg["enabled"] = True

        field_options = []
        if self.available_fields:
            for h in self.available_fields:
                field_options.append((h["key"], f"{h['label']} ({h['key']})"))
        else:
            for h in self.config.get("default_headers", []):
                field_options.append((h["key"], f"{h['label']} ({h['key']})"))

        if not field_options:
            print("\n⚠️  没有可用的字段，请先配置导出字段")
            split_cfg["enabled"] = False
            input("\n按回车继续...")
            return

        default_field_idx = 0
        current_field = split_cfg.get("split_field", "")
        for i, (k, _) in enumerate(field_options):
            if k == current_field:
                default_field_idx = i
                break

        split_cfg["split_field"] = prompt_choice(
            "\n请选择用于拆分的字段",
            field_options,
            default_index=default_field_idx,
        )

        split_rule = prompt_choice(
            "\n请选择拆分规则",
            [
                ("by_value", "按字段值拆分（每个不同值一个工作表，推荐）"),
                ("by_range", "按数值范围分组（适合薪资、年龄等数值字段）"),
                ("by_custom", "自定义规则分组（灵活组合多个条件）"),
            ],
            default_index=0 if split_cfg.get("split_rule", "by_value") == "by_value" else 0,
        )
        split_cfg["split_rule"] = split_rule

        split_cfg["include_all_sheet"] = prompt_confirm(
            "是否额外生成一个包含全部数据的汇总工作表？",
            default=split_cfg.get("include_all_sheet", True),
        )
        if split_cfg["include_all_sheet"]:
            split_cfg["all_sheet_name"] = prompt_input(
                "汇总工作表名称",
                default=split_cfg.get("all_sheet_name", "全部数据"),
            )

        split_cfg["empty_value_label"] = prompt_input(
            "空值或未匹配数据的分组名称",
            default=split_cfg.get("empty_value_label", "未分类"),
        )

        print("\n工作表命名模板可用占位符:")
        print("  {value}  - 分组名称/值")
        print("  {index}  - 分组序号（从1开始）")
        print("  {num}    - 同 {index}")
        print("  {count}  - 总组数")
        split_cfg["sheet_name_template"] = prompt_input(
            "工作表命名模板",
            default=split_cfg.get("sheet_name_template", "{value}"),
        )

        if split_rule == "by_range":
            self._config_split_range_groups(split_cfg)
        elif split_rule == "by_custom":
            self._config_split_custom_rules(split_cfg)

        print("\n✅ 拆分配置完成")
        input("\n按回车继续...")

    def _config_split_range_groups(self, split_cfg):
        print("\n配置数值范围分组:")
        print("  每个分组需要设置名称、最小值和最大值。")
        print("  例如: 名称='青年', min=0, max=30  表示年龄 0~30 岁")

        existing = split_cfg.get("range_groups", [])
        if not existing:
            existing = [
                {"name": "低", "min": None, "max": 1000, "include_min": True, "include_max": False},
                {"name": "中", "min": 1000, "max": 5000, "include_min": True, "include_max": False},
                {"name": "高", "min": 5000, "max": None, "include_min": True, "include_max": False},
            ]

        while True:
            clear_screen()
            print_step(self.current_step, self.total_steps, "配置数值范围分组")

            if existing:
                print(f"\n当前已配置 {len(existing)} 个范围分组:")
                for i, g in enumerate(existing, 1):
                    min_str = f"{g.get('min', '-∞')}"
                    max_str = f"{g.get('max', '+∞')}"
                    left_b = "[" if g.get("include_min", True) else "("
                    right_b = "]" if g.get("include_max", False) else ")"
                    print(f"  {i}. {g['name']:20s}  范围: {left_b}{min_str}, {max_str}{right_b}")
            else:
                print("\n当前未配置任何范围分组")

            print("\n操作选项:")
            print("  1. 添加分组")
            print("  2. 修改分组")
            print("  3. 删除分组")
            print("  4. 使用示例分组（低/中/高）")
            print("  0. 完成配置")

            choice = prompt_number(
                "\n请选择操作",
                min_value=0,
                max_value=4,
                default=0,
            )

            if choice == 0:
                if not existing:
                    print("  ⚠️  至少需要配置一个分组")
                    input("\n按回车继续...")
                    continue
                break
            elif choice == 1:
                group = self._prompt_range_group()
                if group:
                    existing.append(group)
                    print(f"  ✅ 已添加分组: {group['name']}")
            elif choice == 2:
                if not existing:
                    print("  ⚠️  没有可修改的分组")
                else:
                    idx = prompt_number("请选择要修改的分组序号", min_value=1, max_value=len(existing)) - 1
                    group = self._prompt_range_group(existing[idx])
                    if group:
                        existing[idx] = group
                        print(f"  ✅ 已更新分组: {group['name']}")
            elif choice == 3:
                if not existing:
                    print("  ⚠️  没有可删除的分组")
                else:
                    idx = prompt_number("请选择要删除的分组序号", min_value=1, max_value=len(existing)) - 1
                    name = existing[idx]["name"]
                    if prompt_confirm(f"确定删除分组 '{name}' 吗？", default=False):
                        existing.pop(idx)
                        print(f"  ✅ 已删除分组: {name}")
            elif choice == 4:
                existing = [
                    {"name": "低", "min": None, "max": 1000, "include_min": True, "include_max": False},
                    {"name": "中", "min": 1000, "max": 5000, "include_min": True, "include_max": False},
                    {"name": "高", "min": 5000, "max": None, "include_min": True, "include_max": False},
                ]
                print("  ✅ 已使用示例分组")

            if choice != 0:
                input("\n按回车继续...")

        split_cfg["range_groups"] = existing

    def _prompt_range_group(self, existing=None):
        existing = existing or {}
        print("\n配置范围分组:")

        name = prompt_input(
            "分组名称",
            default=existing.get("name", ""),
        )

        min_str = prompt_input(
            "最小值（留空表示无下限）",
            default=str(existing.get("min", "")) if existing.get("min") is not None else "",
            required=False,
        )
        min_val = float(min_str) if min_str else None

        max_str = prompt_input(
            "最大值（留空表示无上限）",
            default=str(existing.get("max", "")) if existing.get("max") is not None else "",
            required=False,
        )
        max_val = float(max_str) if max_str else None

        include_min = prompt_confirm(
            "是否包含最小值？（即 >= min）",
            default=existing.get("include_min", True),
        )
        include_max = prompt_confirm(
            "是否包含最大值？（即 <= max）",
            default=existing.get("include_max", False),
        )

        if min_val is None and max_val is None:
            print("  ⚠️  最小值和最大值不能同时为空")
            return None

        return {
            "name": name,
            "min": min_val,
            "max": max_val,
            "include_min": include_min,
            "include_max": include_max,
        }

    def _config_split_custom_rules(self, split_cfg):
        print("\n配置自定义分组规则:")
        print("  支持三种匹配方式:")
        print("    1. 指定值列表 - 例如 values: ['技术部', '产品部']")
        print("    2. 数值区间   - 例如 min: 20, max: 30")
        print("    3. 条件表达式 - 例如 condition: 'value and len(str(value)) > 5'")

        existing = split_cfg.get("custom_rules", [])
        if not existing:
            existing = []

        while True:
            clear_screen()
            print_step(self.current_step, self.total_steps, "配置自定义分组规则")

            if existing:
                print(f"\n当前已配置 {len(existing)} 条自定义规则:")
                for i, rule in enumerate(existing, 1):
                    desc = self._describe_custom_rule(rule)
                    print(f"  {i}. {rule['name']:20s}  条件: {desc}")
            else:
                print("\n当前未配置任何自定义规则")

            print("\n操作选项:")
            print("  1. 添加规则（值列表方式）")
            print("  2. 添加规则（数值区间方式）")
            print("  3. 添加规则（表达式方式）")
            print("  4. 删除规则")
            print("  0. 完成配置")

            choice = prompt_number(
                "\n请选择操作",
                min_value=0,
                max_value=4,
                default=0,
            )

            if choice == 0:
                if not existing:
                    print("  ⚠️  至少需要配置一条规则")
                    input("\n按回车继续...")
                    continue
                break
            elif choice == 1:
                rule = self._prompt_custom_values_rule()
                if rule:
                    existing.append(rule)
                    print(f"  ✅ 已添加规则: {rule['name']}")
            elif choice == 2:
                rule = self._prompt_custom_range_rule()
                if rule:
                    existing.append(rule)
                    print(f"  ✅ 已添加规则: {rule['name']}")
            elif choice == 3:
                rule = self._prompt_custom_expr_rule()
                if rule:
                    existing.append(rule)
                    print(f"  ✅ 已添加规则: {rule['name']}")
            elif choice == 4:
                if not existing:
                    print("  ⚠️  没有可删除的规则")
                else:
                    idx = prompt_number("请选择要删除的规则序号", min_value=1, max_value=len(existing)) - 1
                    name = existing[idx]["name"]
                    if prompt_confirm(f"确定删除规则 '{name}' 吗？", default=False):
                        existing.pop(idx)
                        print(f"  ✅ 已删除规则: {name}")

            if choice != 0:
                input("\n按回车继续...")

        split_cfg["custom_rules"] = existing

        split_cfg["fallback_group_name"] = prompt_input(
            "\n未匹配任何规则的数据归入的分组名",
            default=split_cfg.get("fallback_group_name", split_cfg.get("empty_value_label", "其他")),
        )

    def _describe_custom_rule(self, rule):
        if "values" in rule:
            vals = rule["values"]
            if isinstance(vals, list):
                return f"值在 [{', '.join(str(v) for v in vals)}]"
            return f"值等于 {vals}"
        if "min" in rule or "max" in rule:
            parts = []
            if rule.get("min") is not None:
                op = ">=" if rule.get("include_min", True) else ">"
                parts.append(f"{op} {rule['min']}")
            if rule.get("max") is not None:
                op = "<=" if rule.get("include_max", False) else "<"
                parts.append(f"{op} {rule['max']}")
            return " 且 ".join(parts) if parts else "任意"
        if "condition" in rule:
            return f"表达式: {rule['condition']}"
        return "未知条件"

    def _prompt_custom_values_rule(self):
        print("\n添加值列表匹配规则:")
        name = prompt_input("分组名称", required=True)
        values_str = prompt_input(
            "输入匹配的值（多个值用逗号分隔，例如: 技术部,产品部,设计部）",
            required=True,
        )
        values = [v.strip() for v in values_str.split(",") if v.strip()]
        if not values:
            print("  ⚠️  至少需要输入一个值")
            return None
        return {"name": name, "values": values}

    def _prompt_custom_range_rule(self):
        print("\n添加数值区间匹配规则:")
        name = prompt_input("分组名称", required=True)

        min_str = prompt_input("最小值（留空表示无下限）", required=False)
        min_val = float(min_str) if min_str else None

        max_str = prompt_input("最大值（留空表示无上限）", required=False)
        max_val = float(max_str) if max_str else None

        if min_val is None and max_val is None:
            print("  ⚠️  最小值和最大值不能同时为空")
            return None

        include_min = True
        include_max = False
        if min_val is not None:
            include_min = prompt_confirm("是否包含最小值？(>= min)", default=True)
        if max_val is not None:
            include_max = prompt_confirm("是否包含最大值？(<= max)", default=False)

        return {
            "name": name,
            "min": min_val,
            "max": max_val,
            "include_min": include_min,
            "include_max": include_max,
        }

    def _prompt_custom_expr_rule(self):
        print("\n添加表达式匹配规则:")
        print("  使用 'value' 代表当前字段值，例如:")
        print("    value.startswith('A')    - 值以 A 开头")
        print("    len(str(value)) > 10     - 值的长度大于10")
        print("    value in ['X','Y']       - 值是 X 或 Y")

        name = prompt_input("分组名称", required=True)
        condition = prompt_input(
            "Python 表达式（使用 value 变量）",
            required=True,
        )
        return {"name": name, "condition": condition}

    def _config_general_title(self):
        for cfg_key in ["html_config", "markdown_config", "pdf_config"]:
            cfg = self.config.setdefault(cfg_key, {})
            cfg["title"] = prompt_input(
                "请输入导出文档标题",
                default=cfg.get("title", "数据导出"),
            )

    def _config_excel_styles(self):
        self.config["sheet_name"] = prompt_input(
            "请输入工作表名称",
            default=self.config.get("sheet_name", "数据导出"),
        )

        while True:
            clear_screen()
            print_step(self.current_step, self.total_steps, "配置 Excel 导出样式")

            print("\n样式快捷操作:")
            print("  1. 应用样式模板")
            print("  2. 保存当前样式为模板")
            print("  3. 配置表头样式")
            print("  4. 配置数据行样式")
            print("  0. 完成样式配置")

            choice = prompt_number(
                "\n请选择操作",
                min_value=0,
                max_value=4,
                default=0,
            )

            if choice == 0:
                break
            elif choice == 1:
                self._apply_template_interactive()
                input("\n按回车继续...")
            elif choice == 2:
                self._save_as_template_interactive()
                input("\n按回车继续...")
            elif choice == 3:
                self.config["style_header"] = prompt_confirm(
                    "是否启用表头样式？",
                    default=self.config.get("style_header", True),
                )
                if self.config["style_header"]:
                    self._config_header_style()
                input("\n按回车继续...")
            elif choice == 4:
                self.config["style_alt_rows"] = prompt_confirm(
                    "是否启用隔行变色？",
                    default=self.config.get("style_alt_rows", True),
                )
                self._config_data_style()
                input("\n按回车继续...")

    def step_select_export_format(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "选择导出格式")

        print("\n支持的导出格式:")
        format_list = []
        for idx, (fmt_id, fmt_info) in enumerate(EXPORT_FORMATS.items(), 1):
            print(f"  {idx:2d}. {fmt_info['label']:20s} - {fmt_info['description']}")
            format_list.append(fmt_id)

        default_fmt = self.config.get("export_format", "excel")
        default_idx = format_list.index(default_fmt) if default_fmt in format_list else 0

        fmt_idx = prompt_number(
            "\n请选择导出格式",
            min_value=1,
            max_value=len(format_list),
            default=default_idx + 1,
        ) - 1
        selected_fmt = format_list[fmt_idx]
        self.config["export_format"] = selected_fmt
        fmt_info = EXPORT_FORMATS[selected_fmt]

        output_key = f"{selected_fmt}_output_path" if selected_fmt != "excel" else "excel_output_path"
        default_path = self.config.get(output_key, get_default_output_path(selected_fmt))
        self.config[output_key] = prompt_file_path(
            f"请输入 {fmt_info['label'].split('(')[0].strip()} 输出文件路径",
            default=default_path,
            file_filter=fmt_info["extension"],
            must_exist=False,
        )

        if prompt_confirm(f"\n是否配置 {fmt_info['label'].split('(')[0].strip()} 的高级选项？", default=False):
            self._config_format_options(selected_fmt)

        print(f"\n✅ 已选择导出格式: {fmt_info['label']}")
        input("\n按回车继续...")

    def _config_format_options(self, fmt):
        if fmt == "csv":
            self._config_csv_options()
        elif fmt == "tsv":
            self._config_tsv_options()
        elif fmt == "html":
            self._config_html_options()
        elif fmt == "markdown":
            self._config_markdown_options()
        elif fmt == "json":
            self._config_json_options()
        elif fmt == "pdf":
            self._config_pdf_options()
        elif fmt == "excel":
            print("\nExcel 的样式选项已在之前步骤中配置")

    def _config_csv_options(self):
        cfg = self.config.setdefault("csv_config", {})
        print("\nCSV 配置选项:")
        cfg["encoding"] = prompt_choice(
            "文件编码",
            [
                ("utf-8-sig", "UTF-8 with BOM (Excel 兼容，推荐)"),
                ("utf-8", "UTF-8"),
                ("gbk", "GBK (简体中文 Windows)"),
                ("gb18030", "GB18030"),
            ],
            default_index=0 if cfg.get("encoding", "utf-8-sig") == "utf-8-sig" else 0,
        )
        cfg["delimiter"] = prompt_choice(
            "字段分隔符",
            [
                (",", "逗号 (,)"),
                (";", "分号 (;)"),
                ("\t", "制表符 (Tab)"),
                ("|", "竖线 (|)"),
            ],
            default_index=0,
        )
        cfg["include_header"] = prompt_confirm(
            "包含表头行？",
            default=cfg.get("include_header", True),
        )
        cfg["quoting"] = prompt_choice(
            "引号策略",
            [
                ("minimal", "仅在需要时使用引号 (推荐)"),
                ("all", "所有字段都加引号"),
                ("nonnumeric", "非数字字段加引号"),
                ("none", "不使用引号"),
            ],
            default_index=0,
        )

    def _config_tsv_options(self):
        cfg = self.config.setdefault("tsv_config", {})
        print("\nTSV 配置选项:")
        cfg["encoding"] = prompt_choice(
            "文件编码",
            [
                ("utf-8-sig", "UTF-8 with BOM (推荐)"),
                ("utf-8", "UTF-8"),
                ("gbk", "GBK"),
            ],
            default_index=0,
        )
        cfg["include_header"] = prompt_confirm(
            "包含表头行？",
            default=cfg.get("include_header", True),
        )

    def _config_html_options(self):
        cfg = self.config.setdefault("html_config", {})
        print("\nHTML 配置选项:")
        cfg["title"] = prompt_input(
            "页面标题",
            default=cfg.get("title", "数据导出"),
        )
        cfg["include_index"] = prompt_confirm(
            "显示行号列？",
            default=cfg.get("include_index", False),
        )
        cfg["pretty_print"] = prompt_confirm(
            "格式化 HTML 源码（缩进换行）？",
            default=cfg.get("pretty_print", True),
        )
        cfg["style"] = prompt_choice(
            "页面样式",
            [
                ("default", "默认美观样式 (推荐)"),
                ("compact", "紧凑样式"),
                ("none", "无样式"),
            ],
            default_index=0,
        )
        if prompt_confirm("是否使用自定义 CSS？", default=False):
            cfg["custom_css"] = prompt_input(
                "输入自定义 CSS 内容（留空清除）",
                default=cfg.get("custom_css", ""),
                required=False,
            )
        else:
            cfg["custom_css"] = ""

    def _config_markdown_options(self):
        cfg = self.config.setdefault("markdown_config", {})
        print("\nMarkdown 配置选项:")
        cfg["title"] = prompt_input(
            "文档标题",
            default=cfg.get("title", "数据导出"),
        )
        cfg["include_index"] = prompt_confirm(
            "显示行号列？",
            default=cfg.get("include_index", False),
        )
        cfg["max_col_width"] = prompt_number(
            "单元格最大字符宽度（超出截断）",
            default=cfg.get("max_col_width", 50),
            min_value=10,
            max_value=500,
        )

    def _config_json_options(self):
        cfg = self.config.setdefault("json_config", {})
        print("\nJSON 配置选项:")
        indent_opt = prompt_choice(
            "格式化缩进",
            [
                (2, "2 空格 (推荐)"),
                (4, "4 空格"),
                (0, "不缩进（压缩）"),
            ],
            default_index=0,
        )
        cfg["indent"] = indent_opt
        cfg["ensure_ascii"] = prompt_confirm(
            "转义非 ASCII 字符（中文转为 \\uXXXX）？",
            default=cfg.get("ensure_ascii", False),
        )
        cfg["include_labels"] = prompt_confirm(
            "使用字段标签（label）作为键名？否则使用原始 key",
            default=cfg.get("include_labels", False),
        )

    def _config_pdf_options(self):
        cfg = self.config.setdefault("pdf_config", {})
        print("\nPDF 配置选项:")
        cfg["title"] = prompt_input(
            "文档标题",
            default=cfg.get("title", "数据导出"),
        )
        cfg["include_index"] = prompt_confirm(
            "显示行号列？",
            default=cfg.get("include_index", False),
        )
        cfg["page_size"] = prompt_choice(
            "纸张大小",
            [
                ("A4", "A4 (推荐)"),
                ("letter", "Letter"),
            ],
            default_index=0,
        )
        cfg["orientation"] = prompt_choice(
            "页面方向",
            [
                ("portrait", "纵向 (Portrait)"),
                ("landscape", "横向 (Landscape)"),
            ],
            default_index=0,
        )
        cfg["font_size"] = prompt_number(
            "正文字号",
            default=cfg.get("font_size", 10),
            min_value=6,
            max_value=24,
        )

    def step_review_and_confirm(self):
        clear_screen()
        print_title("配置确认")

        export_fmt = self.config.get("export_format", "excel")
        fmt_info = EXPORT_FORMATS.get(export_fmt, {})
        output_key = f"{export_fmt}_output_path" if export_fmt != "excel" else "excel_output_path"

        print("\n当前配置摘要:")
        print(f"\n  JSON 文件: {self.config['json_file_path']}")
        print(f"  导出格式: {fmt_info.get('label', export_fmt)}")
        print(f"  输出文件: {self.config.get(output_key, '未设置')}")
        print(f"  导出字段: {len(self.config['default_headers'])} 个")
        print(f"  自动检测新字段: {'是' if self.config['auto_detect_headers'] else '否'}")
        if export_fmt == "excel":
            print(f"  工作表名: {self.config['sheet_name']}")
            print(f"  表头样式: {'启用' if self.config['style_header'] else '禁用'}")
            print(f"  隔行变色: {'启用' if self.config['style_alt_rows'] else '禁用'}")

            split_cfg = self.config.get("split_config", {})
            if split_cfg.get("enabled"):
                rule_label = {
                    "by_value": "按字段值",
                    "by_range": "按数值范围",
                    "by_custom": "自定义规则",
                }.get(split_cfg.get("split_rule", "by_value"), split_cfg.get("split_rule", "by_value"))
                print(f"  工作表拆分: 已启用（{rule_label}）")
                print(f"    拆分字段: {split_cfg.get('split_field', '')}")
                print(f"    Sheet命名: {split_cfg.get('sheet_name_template', '{value}')}")
                print(f"    汇总Sheet: {'是' if split_cfg.get('include_all_sheet', True) else '否'}")
            else:
                print("  工作表拆分: 未启用")

        validation_rules = self.config.get("validation_rules", [])
        if validation_rules:
            print(f"  数据校验: {len(validation_rules)} 条规则")
        else:
            print("  数据校验: 未配置")

        print("\n详细字段列表:")
        for h in self.config["default_headers"]:
            print(f"  • {h['label']} ({h['key']}) - 宽度: {h['width']}")

        errors = validate_config(self.config)
        if errors:
            print("\n⚠️  配置存在以下问题:")
            for e in errors:
                print(f"  • {e}")
            if not prompt_confirm("是否继续？", default=False):
                return False

        print("\n操作选项:")
        print("  1. 预览并筛选数据")
        print("  2. 直接保存配置并继续")
        choice = prompt_choice(
            "请选择操作",
            [
                ("preview", "预览并筛选数据"),
                ("continue", "直接保存配置并继续"),
            ],
            default_index=1,
        )

        if choice == "preview":
            try:
                from data_previewer import run_data_preview
                from json_to_excel import load_json
                data = load_json(self.config["json_file_path"])
                print(f"\n正在启动预览，共 {len(data)} 条数据...")
                input("\n按回车继续...")
                filtered = run_data_preview(data, self.config)
                if filtered is not None and len(filtered) < len(data):
                    if prompt_confirm(f"当前筛选后有 {len(filtered)} 条数据，是否使用筛选后的数据进行导出？", default=True):
                        self.config["_filtered_data"] = filtered
                        print(f"\n✅ 已设置使用筛选后的 {len(filtered)} 条数据")
            except Exception as e:
                print(f"\n⚠️  预览失败: {e}")
                input("\n按回车继续...")

        save = prompt_confirm("\n是否保存配置到 config.json？", default=True)
        if save:
            save_path = save_config(self.config)
            print(f"\n✅ 配置已保存到: {save_path}")

        return True

    def _get_sample_value(self, key):
        if not self.sample_data:
            return "N/A"
        for item in self.sample_data:
            flat = flatten_dict(item)
            if key in flat:
                val = str(flat[key])
                return val[:30] + "..." if len(val) > 30 else val
        return "(空)"

    def _manual_add_fields(self):
        while True:
            key = prompt_input("\n请输入字段 key（留空结束）", required=False)
            if not key:
                break
            label = prompt_input("请输入字段显示名称", default=key.replace("_", " ").title())
            width = prompt_number("请输入列宽", default=15, min_value=5, max_value=200)
            add_header(self.config, key, label, width)
            existing = next((h for h in self.available_fields if h["key"] == key), None)
            if existing:
                existing["label"] = label
                existing["width"] = width
            else:
                self.available_fields.append({
                    "key": key,
                    "label": label,
                    "width": width,
                })
            print(f"✅ 已添加字段: {label} ({key})")

    def _merge_config_headers_to_available(self):
        available_keys = {h["key"] for h in self.available_fields}
        for config_header in self.config.get("default_headers", []):
            if config_header["key"] not in available_keys:
                self.available_fields.append({
                    "key": config_header["key"],
                    "label": config_header.get("label", config_header["key"]),
                    "width": config_header.get("width", 15),
                })

    def _modify_label(self, headers):
        idx = prompt_number("请选择要修改的字段序号", min_value=1, max_value=len(headers)) - 1
        header = headers[idx]
        new_label = prompt_input(
            f"当前标签: {header['label']}，请输入新标签",
            default=header["label"],
        )
        update_header_label(self.config, header["key"], new_label)
        print(f"✅ 标签已更新为: {new_label}")

    def _modify_width(self, headers):
        idx = prompt_number("请选择要修改的字段序号", min_value=1, max_value=len(headers)) - 1
        header = headers[idx]
        new_width = prompt_number(
            f"当前宽度: {header['width']}，请输入新宽度 (5-200)",
            default=header["width"],
            min_value=5,
            max_value=200,
        )
        update_header_width(self.config, header["key"], new_width)
        print(f"✅ 宽度已更新为: {new_width}")

    def _reorder_fields(self, headers):
        print("\n当前顺序:")
        for i, h in enumerate(headers, 1):
            print(f"  {i}. {h['label']}")

        print("\n请输入新的顺序（用逗号分隔序号，例如: 3,1,2,5,4）")
        while True:
            order_input = prompt_input("新顺序")
            try:
                indices = [int(x.strip()) - 1 for x in order_input.split(",")]
                if len(indices) != len(headers) or len(set(indices)) != len(headers):
                    print(f"⚠️  请输入 1 到 {len(headers)} 的所有序号，且不重复")
                    continue
                if any(i < 0 or i >= len(headers) for i in indices):
                    print(f"⚠️  序号必须在 1 到 {len(headers)} 之间")
                    continue
                new_order_keys = [headers[i]["key"] for i in indices]
                reorder_headers(self.config, new_order_keys)
                print("✅ 顺序已更新")
                break
            except ValueError:
                print("⚠️  请输入有效的数字，用逗号分隔")

    def _remove_field(self, headers):
        if len(headers) <= 1:
            print("⚠️  至少需要保留一个字段")
            return

        idx = prompt_number("请选择要移除的字段序号", min_value=1, max_value=len(headers)) - 1
        header = headers[idx]
        if prompt_confirm(f"确定要移除字段 '{header['label']}' 吗？", default=False):
            remove_header(self.config, header["key"])
            print(f"✅ 已移除字段: {header['label']}")

    def _auto_width_all(self, headers):
        if not self.available_fields:
            print("⚠️  没有可用的参考数据")
            return

        auto_h_map = {h["key"]: h for h in self.available_fields}
        for header in headers:
            if header["key"] in auto_h_map:
                auto_w = auto_h_map[header["key"]]["width"]
                update_header_width(self.config, header["key"], auto_w)
        print("✅ 已全部使用自动宽度")

    def _config_data_style(self):
        data_style = self.config.setdefault("data_style", {})

        print("\n数据行字体配置:")
        data_style["font_name"] = self._prompt_font_family(
            "选择数据行字体",
            default=data_style.get("font_name", "Microsoft YaHei"),
        )

        data_style["font_bold"] = prompt_confirm(
            "数据行字体加粗？",
            default=data_style.get("font_bold", False),
        )

        data_style["font_size"] = prompt_number(
            "数据行字体大小",
            default=data_style.get("font_size", 11),
            min_value=8,
            max_value=24,
        )

        data_style["font_color"] = prompt_color(
            "数据行字体颜色",
            default=data_style.get("font_color", "#000000"),
        )

        data_style["wrap_text"] = prompt_confirm(
            "数据行自动换行？",
            default=data_style.get("wrap_text", True),
        )

        if self.config.get("style_alt_rows", True):
            data_style["alt_row_color"] = prompt_color(
                "隔行背景色（十六进制）",
                default=data_style.get("alt_row_color", "#F2F2F2"),
            )

        data_style["alignment"] = prompt_choice(
            "数据行水平对齐方式",
            [
                ("left", "左对齐 (推荐)"),
                ("center", "居中对齐"),
                ("right", "右对齐"),
            ],
            default_index=0,
        )

        data_style["vertical_alignment"] = prompt_choice(
            "数据行垂直对齐方式",
            [
                ("top", "顶部对齐"),
                ("center", "居中对齐 (推荐)"),
                ("bottom", "底部对齐"),
            ],
            default_index=1,
        )

        print("\n数据行边框配置:")
        data_style["border_style"] = self._prompt_border_style(
            "选择数据行边框样式",
            default=data_style.get("border_style", "thin"),
        )

        if data_style["border_style"] is not None:
            data_style["border_color"] = prompt_color(
                "数据行边框颜色",
                default=data_style.get("border_color", "#000000"),
            )

    def _config_header_style(self):
        header_style = self.config.setdefault("header_style", {})

        print("\n表头字体配置:")
        header_style["font_name"] = self._prompt_font_family(
            "选择表头字体",
            default=header_style.get("font_name", "Microsoft YaHei"),
        )

        header_style["font_bold"] = prompt_confirm(
            "表头字体加粗？",
            default=header_style.get("font_bold", True),
        )

        header_style["font_size"] = prompt_number(
            "表头字体大小",
            default=header_style.get("font_size", 12),
            min_value=8,
            max_value=24,
        )

        header_style["font_color"] = prompt_color(
            "表头字体颜色",
            default=header_style.get("font_color", "#FFFFFF"),
        )

        header_style["bg_color"] = prompt_color(
            "表头背景颜色",
            default=header_style.get("bg_color", "#4472C4"),
        )

        header_style["alignment"] = prompt_choice(
            "表头文字对齐方式",
            [
                ("left", "左对齐"),
                ("center", "居中对齐 (推荐)"),
                ("right", "右对齐"),
            ],
            default_index=1,
        )

        header_style["vertical_alignment"] = prompt_choice(
            "表头垂直对齐方式",
            [
                ("top", "顶部对齐"),
                ("center", "居中对齐 (推荐)"),
                ("bottom", "底部对齐"),
            ],
            default_index=1,
        )

        print("\n表头边框配置:")
        header_style["border_style"] = self._prompt_border_style(
            "选择表头边框样式",
            default=header_style.get("border_style", "thin"),
        )

        if header_style["border_style"] is not None:
            header_style["border_color"] = prompt_color(
                "表头边框颜色",
                default=header_style.get("border_color", "#000000"),
            )


def run_wizard(config=None):
    wizard = ConfigWizard(config)
    return wizard.run()


if __name__ == "__main__":
    run_wizard()
