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

from json_to_excel import load_json, auto_detect_headers, flatten_dict


TOTAL_STEPS = 6


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
            print_title("JSON 转 Excel - 交互式配置向导")
            print("\n欢迎使用配置向导！我们将引导您完成所有设置。\n")

            self.current_step = 0
            should_load_config = prompt_confirm("是否加载已有配置文件？", default=True)
            self.total_steps = 6 + (1 if should_load_config else 0)

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
            self.step_configure_styles()

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

    def step_configure_styles(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置导出样式")

        self.config["sheet_name"] = prompt_input(
            "请输入工作表名称",
            default=self.config.get("sheet_name", "数据导出"),
        )

        self.config["excel_output_path"] = prompt_file_path(
            "请输入 Excel 输出文件路径",
            default=self.config.get("excel_output_path", "./output/result.xlsx"),
            file_filter=".xlsx",
            must_exist=False,
        )

        self.config["style_header"] = prompt_confirm(
            "是否启用表头样式？",
            default=self.config.get("style_header", True),
        )

        if self.config["style_header"]:
            self._config_header_style()

        self.config["style_alt_rows"] = prompt_confirm(
            "是否启用隔行变色？",
            default=self.config.get("style_alt_rows", True),
        )

        if self.config["style_alt_rows"]:
            data_style = self.config.setdefault("data_style", {})
            data_style["alt_row_color"] = prompt_color(
                "请输入隔行背景色（十六进制）",
                default=data_style.get("alt_row_color", "#F2F2F2"),
            )

        print("\n✅ 样式配置完成")
        input("\n按回车继续...")

    def step_review_and_confirm(self):
        clear_screen()
        print_title("配置确认")

        print("\n当前配置摘要:")
        print(f"\n  JSON 文件: {self.config['json_file_path']}")
        print(f"  输出文件: {self.config['excel_output_path']}")
        print(f"  工作表名: {self.config['sheet_name']}")
        print(f"  导出字段: {len(self.config['default_headers'])} 个")
        print(f"  自动检测新字段: {'是' if self.config['auto_detect_headers'] else '否'}")
        print(f"  表头样式: {'启用' if self.config['style_header'] else '禁用'}")
        print(f"  隔行变色: {'启用' if self.config['style_alt_rows'] else '禁用'}")

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

    def _config_header_style(self):
        header_style = self.config.setdefault("header_style", {})

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


def run_wizard(config=None):
    wizard = ConfigWizard(config)
    return wizard.run()


if __name__ == "__main__":
    run_wizard()
