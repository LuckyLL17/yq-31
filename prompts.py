import os
import re


def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def print_title(title):
    line = "=" * (len(title) + 4)
    print(f"\n{line}")
    print(f"  {title}")
    print(line)


def print_step(step_num, total_steps, description):
    print(f"\n[步骤 {step_num}/{total_steps}] {description}")
    print("-" * 40)


def prompt_input(message, default=None, validator=None, required=True):
    while True:
        display_default = f" [{default}]" if default is not None else ""
        user_input = input(f"{message}{display_default}: ").strip()

        if not user_input:
            if default is not None:
                return default
            if not required:
                return None
            print("  ⚠️  此项为必填项，请输入内容")
            continue

        if validator is not None:
            is_valid, error_msg = validator(user_input)
            if not is_valid:
                print(f"  ⚠️  {error_msg}")
                continue

        return user_input


def prompt_confirm(message, default=True):
    yes_str = "Y/n" if default else "y/N"
    while True:
        user_input = input(f"{message} ({yes_str}): ").strip().lower()

        if not user_input:
            return default

        if user_input in ("y", "yes"):
            return True
        if user_input in ("n", "no"):
            return False

        print("  ⚠️  请输入 y 或 n")


def prompt_choice(message, choices, default_index=0):
    while True:
        print(f"\n{message}")
        for i, choice in enumerate(choices):
            marker = "  " if i != default_index else "• "
            display = choice[1] if isinstance(choice, tuple) else str(choice)
            print(f"  {i + 1}. {marker}{display}")

        user_input = input(f"请选择 (1-{len(choices)}, 回车默认 {default_index + 1}): ").strip()

        if not user_input:
            selected = choices[default_index]
            return selected[0] if isinstance(selected, tuple) else selected

        try:
            index = int(user_input) - 1
            if 0 <= index < len(choices):
                selected = choices[index]
                return selected[0] if isinstance(selected, tuple) else selected
        except ValueError:
            pass

        print(f"  ⚠️  请输入 1 到 {len(choices)} 之间的数字")


def prompt_multi_select(message, options, min_selected=1, default_indices=None):
    selected_indices = set(default_indices or [])

    while True:
        print(f"\n{message}")
        print(f"  (可多选，输入序号切换选中状态，输入 'a' 全选，输入 'n' 取消全选，回车确认)")
        print(f"  最少选择 {min_selected} 项")
        print()

        for i, option in enumerate(options):
            status = "✓" if i in selected_indices else " "
            display = option[1] if isinstance(option, tuple) else str(option)
            print(f"  [{status}] {i + 1}. {display}")

        user_input = input("\n输入选择: ").strip().lower()

        if not user_input:
            if len(selected_indices) >= min_selected:
                break
            print(f"  ⚠️  请至少选择 {min_selected} 项")
            continue

        if user_input == "a":
            selected_indices = set(range(len(options)))
            continue

        if user_input == "n":
            selected_indices.clear()
            continue

        try:
            index = int(user_input) - 1
            if 0 <= index < len(options):
                if index in selected_indices:
                    selected_indices.remove(index)
                else:
                    selected_indices.add(index)
            else:
                print(f"  ⚠️  请输入 1 到 {len(options)} 之间的数字")
        except ValueError:
            print("  ⚠️  请输入有效的序号")

    selected = []
    for i in sorted(selected_indices):
        opt = options[i]
        if isinstance(opt, tuple):
            selected.append(opt[0])
        else:
            selected.append(opt)
    return selected


def prompt_number(message, default=None, min_value=None, max_value=None, required=True):
    def validator(value):
        try:
            num = int(value) if "." not in value else float(value)
            if min_value is not None and num < min_value:
                return False, f"数值不能小于 {min_value}"
            if max_value is not None and num > max_value:
                return False, f"数值不能大于 {max_value}"
            return True, None
        except ValueError:
            return False, "请输入有效的数字"

    result = prompt_input(message, default=str(default) if default is not None else None,
                          validator=validator, required=required)
    if result is None:
        return None
    return int(result) if "." not in result else float(result)


def prompt_file_path(message, default=None, file_filter=None, must_exist=True, required=True):
    def validator(path):
        if must_exist and not os.path.exists(path):
            return False, f"文件不存在: {path}"
        if must_exist and not os.path.isfile(path):
            return False, f"路径不是文件: {path}"
        if file_filter and isinstance(file_filter, str):
            if not path.lower().endswith(file_filter.lower()):
                return False, f"请选择 {file_filter} 格式的文件"
        return True, None

    return prompt_input(message, default=default, validator=validator, required=required)


def prompt_color(message, default="#4472C4"):
    hex_pattern = re.compile(r'^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$')

    def validator(color):
        if hex_pattern.match(color):
            return True, None
        return False, "请输入有效的十六进制颜色代码（如 #4472C4 或 #FFF）"

    return prompt_input(message, default=default, validator=validator)


def format_current_config(config):
    lines = []
    for key, value in config.items():
        if isinstance(value, list):
            lines.append(f"  {key}: {len(value)} 项")
        elif isinstance(value, bool):
            lines.append(f"  {key}: {'是' if value else '否'}")
        else:
            lines.append(f"  {key}: {value}")
    return "\n".join(lines)
