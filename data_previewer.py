import re
import json
import os
import copy
from datetime import datetime

from json_to_excel import load_json, auto_detect_headers, flatten_dict, extract_value
from prompts import clear_screen, print_title, prompt_input, prompt_choice, prompt_number, prompt_confirm


class FilterCondition:
    def __init__(self, field, operator, value, value2=None):
        self.field = field
        self.operator = operator
        self.value = value
        self.value2 = value2

    def __repr__(self):
        if self.operator == "between":
            return f"{self.field} 在 [{self.value}, {self.value2}] 之间"
        op_map = {
            "eq": "=",
            "ne": "≠",
            "gt": ">",
            "gte": "≥",
            "lt": "<",
            "lte": "≤",
            "contains": "包含",
            "icontains": "包含(不区分大小写)",
            "startswith": "以...开头",
            "endswith": "以...结尾",
            "regex": "匹配正则",
            "in": "属于列表",
            "isnull": "为空",
            "notnull": "不为空",
        }
        op_str = op_map.get(self.operator, self.operator)
        return f"{self.field} {op_str} {self.value}"


class DataPreviewer:
    def __init__(self, data, headers=None):
        self.original_data = list(data)
        self.filtered_data = list(data)
        self.headers = headers or auto_detect_headers(data)
        self.conditions = []
        self.preview_count = 10
        self.page = 1
        self.page_size = 10
        self.visible_fields = [h["key"] for h in self.headers]

    def apply_filters(self):
        result = list(self.original_data)
        for cond in self.conditions:
            result = [item for item in result if self._match_condition(item, cond)]
        self.filtered_data = result
        self.page = 1
        return result

    def _match_condition(self, item, condition):
        flat = flatten_dict(item)
        value = flat.get(condition.field, "")
        return self._compare(value, condition.operator, condition.value, condition.value2)

    def _compare(self, value, operator, target, target2=None):
        if operator in ("isnull", "notnull"):
            is_null = value is None or value == ""
            return is_null if operator == "isnull" else not is_null

        if operator in ("in",):
            if isinstance(target, list):
                return str(value) in [str(t) for t in target]
            return False

        if operator == "between":
            try:
                num_val = float(value)
                num_min = float(target)
                num_max = float(target2)
                return num_min <= num_val <= num_max
            except (ValueError, TypeError):
                return str(target) <= str(value) <= str(target2)

        if operator in ("contains", "icontains", "startswith", "endswith", "regex"):
            str_val = str(value) if value is not None else ""
            str_target = str(target) if target is not None else ""
            if operator == "contains":
                return str_target in str_val
            if operator == "icontains":
                return str_target.lower() in str_val.lower()
            if operator == "startswith":
                return str_val.startswith(str_target)
            if operator == "endswith":
                return str_val.endswith(str_target)
            if operator == "regex":
                try:
                    return re.search(str_target, str_val) is not None
                except re.error:
                    return False

        num_operators = ("gt", "gte", "lt", "lte")
        if operator in num_operators or operator in ("eq", "ne"):
            try:
                num_val = float(value)
                num_target = float(target)
                if operator == "gt":
                    return num_val > num_target
                if operator == "gte":
                    return num_val >= num_target
                if operator == "lt":
                    return num_val < num_target
                if operator == "lte":
                    return num_val <= num_target
                if operator == "eq":
                    return num_val == num_target
                if operator == "ne":
                    return num_val != num_target
            except (ValueError, TypeError):
                if operator == "eq":
                    return str(value) == str(target)
                if operator == "ne":
                    return str(value) != str(target)
                return False

        return False

    def add_condition(self, condition):
        self.conditions.append(condition)
        self.apply_filters()

    def remove_condition(self, index):
        if 0 <= index < len(self.conditions):
            self.conditions.pop(index)
            self.apply_filters()
            return True
        return False

    def clear_conditions(self):
        self.conditions.clear()
        self.filtered_data = list(self.original_data)
        self.page = 1

    def get_paged_data(self):
        start = (self.page - 1) * self.page_size
        end = start + self.page_size
        return self.filtered_data[start:end]

    def total_pages(self):
        return max(1, (len(self.filtered_data) + self.page_size - 1) // self.page_size)


def _truncate_str(s, max_len=20):
    if s is None:
        return ""
    s = str(s)
    if len(s) <= max_len:
        return s
    return s[: max_len - 3] + "..."


def _calc_col_widths(previewer, data_rows):
    col_widths = {}
    for h in previewer.headers:
        if h["key"] not in previewer.visible_fields:
            continue
        label_width = len(h["label"])
        max_val_width = 0
        for item in data_rows:
            val = extract_value(item, h["key"])
            val_str = str(val) if val is not None else ""
            max_val_width = max(max_val_width, min(len(val_str), 30))
        col_widths[h["key"]] = max(label_width, max_val_width, 8)
    return col_widths


def _print_table(previewer, data_rows):
    visible_headers = [h for h in previewer.headers if h["key"] in previewer.visible_fields]
    if not visible_headers:
        print("  (没有可显示的字段)")
        return

    col_widths = _calc_col_widths(previewer, data_rows)

    separator = "+"
    header_line = "|"
    for h in visible_headers:
        w = col_widths[h["key"]]
        separator += "-" * (w + 2) + "+"
        header_line += " " + h["label"].ljust(w) + " |"

    print("  " + separator)
    print("  " + header_line)
    print("  " + separator)

    for item in data_rows:
        row_line = "|"
        for h in visible_headers:
            w = col_widths[h["key"]]
            val = extract_value(item, h["key"])
            val_str = _truncate_str(val, w)
            row_line += " " + val_str.ljust(w) + " |"
        print("  " + row_line)

    print("  " + separator)


def _print_status_bar(previewer):
    total = len(previewer.filtered_data)
    orig_total = len(previewer.original_data)
    total_p = previewer.total_pages()
    pct = (total / orig_total * 100) if orig_total > 0 else 0
    print(f"\n  📊 数据统计: 显示 {total}/{orig_total} 条 ({pct:.1f}%) | "
          f"筛选条件: {len(previewer.conditions)} 个 | "
          f"第 {previewer.page}/{total_p} 页")


def _print_conditions(previewer):
    if not previewer.conditions:
        print("\n  (当前无筛选条件)")
        return
    print("\n  当前筛选条件:")
    for i, cond in enumerate(previewer.conditions, 1):
        print(f"    {i}. {cond}")


def _print_menu():
    print("\n  ┌─────────────────────────────────────────────────────┐")
    print("  │  数据预览命令菜单                                     │")
    print("  ├─────────────────────────────────────────────────────┤")
    print("  │  p / page    设置每页显示条数                         │")
    print("  │  n / next    下一页                                   │")
    print("  │  b / back    上一页                                   │")
    print("  │  g / goto    跳转到指定页                              │")
    print("  │  ─────────────────────────────────────────────────   │")
    print("  │  f / filter  添加筛选条件                              │")
    print("  │  r / remove  移除筛选条件                              │")
    print("  │  c / clear   清除所有筛选条件                          │")
    print("  │  ─────────────────────────────────────────────────   │")
    print("  │  s / show    选择显示字段                              │")
    print("  │  a / all     查看原始完整数据（前N条）                 │")
    print("  │  ─────────────────────────────────────────────────   │")
    print("  │  e / export  导出筛选后的数据为 Excel                  │")
    print("  │  q / quit    退出预览                                  │")
    print("  └─────────────────────────────────────────────────────┘")


def _select_field(previewer, message="请选择字段"):
    options = [(h["key"], f"{h['label']} ({h['key']})") for h in previewer.headers]
    options.insert(0, ("__cancel__", "取消"))
    choice = prompt_choice(message, options, default_index=0)
    if choice == "__cancel__":
        return None
    return choice


def _select_operator(field_type="any"):
    all_ops = [
        ("eq", "等于 (=)"),
        ("ne", "不等于 (≠)"),
        ("gt", "大于 (>)"),
        ("gte", "大于等于 (≥)"),
        ("lt", "小于 (<)"),
        ("lte", "小于等于 (≤)"),
        ("between", "在...范围之间"),
        ("contains", "包含文本"),
        ("icontains", "包含(不区分大小写)"),
        ("startswith", "以...开头"),
        ("endswith", "以...结尾"),
        ("regex", "正则匹配"),
        ("in", "在列表中"),
        ("isnull", "为空"),
        ("notnull", "不为空"),
    ]

    if field_type == "numeric":
        ops = [
            ("eq", "等于 (=)"),
            ("ne", "不等于 (≠)"),
            ("gt", "大于 (>)"),
            ("gte", "大于等于 (≥)"),
            ("lt", "小于 (<)"),
            ("lte", "小于等于 (≤)"),
            ("between", "在...范围之间"),
            ("isnull", "为空"),
            ("notnull", "不为空"),
        ]
    elif field_type == "string":
        ops = [
            ("eq", "等于"),
            ("ne", "不等于"),
            ("contains", "包含文本"),
            ("icontains", "包含(不区分大小写)"),
            ("startswith", "以...开头"),
            ("endswith", "以...结尾"),
            ("regex", "正则匹配"),
            ("in", "在列表中(逗号分隔)"),
            ("isnull", "为空"),
            ("notnull", "不为空"),
        ]
    else:
        ops = all_ops

    ops.insert(0, ("__cancel__", "取消"))
    return prompt_choice("请选择比较运算符", ops, default_index=0)


def _detect_field_type(previewer, field_key):
    for item in previewer.original_data[:50]:
        flat = flatten_dict(item)
        val = flat.get(field_key, "")
        if val is None or val == "":
            continue
        if isinstance(val, (int, float)):
            return "numeric"
        if isinstance(val, bool):
            return "boolean"
    return "string"


def _add_filter_condition(previewer):
    clear_screen()
    print_title("添加筛选条件")
    _print_conditions(previewer)
    print()

    field = _select_field(previewer, "请选择要筛选的字段")
    if not field:
        return

    field_type = _detect_field_type(previewer, field)
    operator = _select_operator(field_type)
    if operator == "__cancel__":
        return

    value = None
    value2 = None

    if operator in ("isnull", "notnull"):
        pass
    elif operator == "between":
        value = prompt_input("请输入最小值", required=True)
        value2 = prompt_input("请输入最大值", required=True)
    elif operator == "in":
        raw = prompt_input("请输入多个值（用逗号分隔）", required=True)
        value = [v.strip() for v in raw.split(",") if v.strip()]
    else:
        value = prompt_input("请输入比较值", required=True)

    cond = FilterCondition(field, operator, value, value2)
    previewer.add_condition(cond)
    print(f"\n✅ 已添加筛选条件: {cond}")
    input("\n按回车继续...")


def _remove_filter_condition(previewer):
    if not previewer.conditions:
        print("\n  ⚠️  当前没有筛选条件")
        input("\n按回车继续...")
        return

    clear_screen()
    print_title("移除筛选条件")
    _print_conditions(previewer)
    print()

    idx = prompt_number(
        f"请选择要移除的条件编号 (1-{len(previewer.conditions)})",
        min_value=1,
        max_value=len(previewer.conditions),
    )
    if previewer.remove_condition(idx - 1):
        print(f"\n✅ 已移除条件 #{idx}")
    input("\n按回车继续...")


def _select_visible_fields(previewer):
    clear_screen()
    print_title("选择显示字段")

    options = []
    default_indices = []
    for i, h in enumerate(previewer.headers):
        options.append((h["key"], f"{h['label']} ({h['key']})"))
        if h["key"] in previewer.visible_fields:
            default_indices.append(i)

    from prompts import prompt_multi_select
    selected = prompt_multi_select(
        "请选择要显示的字段",
        options,
        min_selected=1,
        default_indices=default_indices,
    )
    previewer.visible_fields = selected
    print(f"\n✅ 已设置显示 {len(selected)} 个字段")
    input("\n按回车继续...")


def _preview_top_n(previewer):
    clear_screen()
    print_title("查看前 N 条原始数据")
    n = prompt_number("请输入要查看的条数", default=20, min_value=1, max_value=1000)
    rows = previewer.original_data[:n]
    print(f"\n  显示前 {n} 条原始数据 (共 {len(previewer.original_data)} 条):\n")
    _print_table(previewer, rows)
    input("\n按回车继续...")


def _export_filtered(previewer, config):
    from json_to_excel import export_to_excel, merge_headers

    clear_screen()
    print_title("导出筛选后的数据")

    default_path = config.get("excel_output_path", "./output/filtered_result.xlsx")
    if default_path.endswith(".xlsx"):
        default_path = default_path.replace(".xlsx", "_filtered.xlsx")

    output_path = prompt_input("请输入输出文件路径", default=default_path, required=True)

    export_config = copy.deepcopy(config)
    export_config["excel_output_path"] = output_path

    visible_headers = [h for h in previewer.headers if h["key"] in previewer.visible_fields]
    auto_headers = auto_detect_headers(previewer.filtered_data)
    headers = merge_headers(visible_headers, auto_headers, export_config)

    try:
        export_to_excel(
            data=previewer.filtered_data,
            headers=headers,
            config=export_config,
        )
        print(f"\n✅ 已导出 {len(previewer.filtered_data)} 条数据到: {os.path.abspath(output_path)}")
    except Exception as e:
        print(f"\n❌ 导出失败: {e}")

    input("\n按回车继续...")


def run_data_preview(data, config=None):
    config = config or {}
    previewer = DataPreviewer(data)

    while True:
        clear_screen()
        print_title("JSON 数据预览与筛选")

        _print_conditions(previewer)
        _print_status_bar(previewer)

        paged_data = previewer.get_paged_data()
        if paged_data:
            print(f"\n  当前页数据 ({len(paged_data)} 条):")
            _print_table(previewer, paged_data)
        else:
            print("\n  (当前页无数据)")

        _print_menu()

        cmd = input("\n  请输入命令: ").strip().lower()

        if cmd in ("q", "quit", "exit"):
            print("\n  退出预览模式")
            return previewer.filtered_data

        elif cmd in ("p", "page"):
            new_size = prompt_number("请输入每页显示条数", default=previewer.page_size, min_value=1, max_value=200)
            previewer.page_size = new_size
            previewer.page = 1

        elif cmd in ("n", "next"):
            if previewer.page < previewer.total_pages():
                previewer.page += 1
            else:
                print("\n  ⚠️  已经是最后一页")

        elif cmd in ("b", "back", "prev"):
            if previewer.page > 1:
                previewer.page -= 1
            else:
                print("\n  ⚠️  已经是第一页")

        elif cmd in ("g", "goto"):
            target = prompt_number(
                f"请输入页码 (1-{previewer.total_pages()})",
                min_value=1,
                max_value=previewer.total_pages(),
            )
            previewer.page = target

        elif cmd in ("f", "filter"):
            _add_filter_condition(previewer)

        elif cmd in ("r", "remove"):
            _remove_filter_condition(previewer)

        elif cmd in ("c", "clear"):
            if previewer.conditions:
                if prompt_confirm("确定要清除所有筛选条件吗？", default=True):
                    previewer.clear_conditions()
                    print("\n  ✅ 已清除所有筛选条件")
                    input("\n按回车继续...")
            else:
                print("\n  ⚠️  当前没有筛选条件")
                input("\n按回车继续...")

        elif cmd in ("s", "show"):
            _select_visible_fields(previewer)

        elif cmd in ("a", "all"):
            _preview_top_n(previewer)

        elif cmd in ("e", "export"):
            _export_filtered(previewer, config)

        elif cmd == "":
            pass

        else:
            print(f"\n  ⚠️  未知命令: {cmd}")
            input("\n按回车继续...")


def start_preview_mode(config):
    json_path = config.get("json_file_path", "./data/sample_data.json")
    if not os.path.exists(json_path):
        print(f"❌ JSON 文件不存在: {json_path}")
        return None

    print(f"\n正在加载数据: {json_path}")
    try:
        data = load_json(json_path)
        print(f"✅ 已加载 {len(data)} 条数据")
    except Exception as e:
        print(f"❌ 加载数据失败: {e}")
        return None

    return run_data_preview(data, config)
