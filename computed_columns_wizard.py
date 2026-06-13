from prompts import clear_screen, print_title, print_step, prompt_input, prompt_choice, prompt_number, prompt_confirm
from computed_columns import FORMULA_TYPES, DATE_DIFF_UNITS, describe_computed_column


def run_computed_columns_wizard(config, available_fields=None):
    clear_screen()
    print_title("计算列配置向导")

    print("\n计算列可以根据已有字段的值自动计算新列。")
    print("例如：工龄 = 当前日期 - 入职日期、年薪 = 薪资 × 12")

    field_options = []
    if available_fields:
        field_options = [(h["key"], f"{h['label']} ({h['key']})") for h in available_fields]
    else:
        field_options = [(h["key"], f"{h['label']} ({h['key']})") for h in config.get("default_headers", [])]

    current_cc = config.get("computed_columns", [])

    while True:
        clear_screen()
        print_title("计算列配置向导")

        if current_cc:
            print(f"\n当前已配置 {len(current_cc)} 个计算列:")
            for i, cc in enumerate(current_cc, 1):
                enabled_str = "" if cc.get("enabled", True) else " [已禁用]"
                print(f"  {i}. {describe_computed_column(cc)}{enabled_str}")
        else:
            print("\n当前未配置任何计算列")

        print("\n操作选项:")
        print("  1. 添加算术计算列（如：年薪 = 薪资 * 12）")
        print("  2. 添加日期差值计算列（如：工龄 = 当前日期 - 入职日期）")
        print("  3. 添加日期加减计算列（如：到期日 = 开始日期 + 30天）")
        print("  4. 添加字符串拼接计算列")
        print("  5. 添加条件表达式计算列（如：薪资等级 = IF 薪资>10000 THEN 高 ELSE 低）")
        print("  6. 添加取整计算列")
        print("  7. 删除计算列")
        print("  8. 启用/禁用计算列")
        print("  0. 保存并退出")

        choice = prompt_number(
            "\n请选择操作",
            min_value=0,
            max_value=8,
            default=0,
        )

        if choice == 0:
            break
        elif choice == 1:
            cc = _prompt_arithmetic_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 2:
            cc = _prompt_date_diff_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 3:
            cc = _prompt_date_add_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 4:
            cc = _prompt_concat_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 5:
            cc = _prompt_conditional_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 6:
            cc = _prompt_round_cc(field_options)
            if cc:
                current_cc.append(cc)
                print(f"  ✅ 已添加计算列: {cc['label']}")
        elif choice == 7:
            _delete_computed_column(current_cc)
        elif choice == 8:
            _toggle_computed_column(current_cc)

        config["computed_columns"] = current_cc
        if choice != 0:
            input("\n按回车继续...")

    config["computed_columns"] = current_cc
    return config


def _prompt_cc_base():
    key = prompt_input("计算列字段 key（英文标识，如 annual_salary）", required=True)
    label = prompt_input("计算列显示名称（如 年薪）", required=True)
    width = prompt_number("列宽", default=15, min_value=5, max_value=200)
    return key, label, width


def _prompt_arithmetic_cc(field_options):
    print("\n添加算术计算列:")
    print("  支持基本算术运算: +, -, *, /, %, 以及函数: abs, round, min, max, int, float, ceil, floor, sqrt")
    print("  示例: salary * 12, salary / 12, abs(salary - bonus)")

    key, label, width = _prompt_cc_base()

    print("\n可用字段:")
    for i, (k, desc) in enumerate(field_options, 1):
        print(f"  {i}. {desc}")

    print("\n请在公式中使用字段的 key 值，例如 salary * 12")
    formula = prompt_input("请输入算术公式", required=True)

    ref_fields = []
    for fk, _ in field_options:
        if fk in formula:
            ref_fields.append(fk)

    if not ref_fields:
        print("  ⚠️  公式中未引用任何字段，请重新配置")
        return None

    print(f"  自动检测到引用字段: {', '.join(ref_fields)}")

    return {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "arithmetic",
        "formula": formula,
        "referenced_fields": ref_fields,
    }


def _prompt_date_diff_cc(field_options):
    print("\n添加日期差值计算列:")
    print("  计算两个日期之间的差值，如工龄 = 当前日期 - 入职日期")

    key, label, width = _prompt_cc_base()

    print("\n可用字段:")
    for i, (k, desc) in enumerate(field_options, 1):
        print(f"  {i}. {desc}")

    start_field = prompt_input("起始日期字段 key（如 join_date）", required=True)
    use_current = prompt_confirm("结束日期是否使用当前日期？", default=True)
    end_field = ""
    if not use_current:
        end_field = prompt_input("结束日期字段 key", required=True)

    unit_options = [(k, f"{v} ({k})") for k, v in DATE_DIFF_UNITS.items()]
    unit = prompt_choice("请选择差值单位", unit_options, default_index=2)

    cc = {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "date_diff",
        "start_field": start_field,
        "use_current_date": use_current,
        "unit": unit,
    }
    if not use_current:
        cc["end_field"] = end_field
    if use_current or end_field:
        cc["referenced_fields"] = [start_field] + ([end_field] if end_field else [])
    return cc


def _prompt_date_add_cc(field_options):
    print("\n添加日期加减计算列:")
    print("  在日期上加上或减去一定时间，如到期日 = 开始日期 + 30天")

    key, label, width = _prompt_cc_base()

    print("\n可用字段:")
    for i, (k, desc) in enumerate(field_options, 1):
        print(f"  {i}. {desc}")

    base_field = prompt_input("基准日期字段 key", required=True)
    value_str = prompt_input("加减的数值（负数表示减去）", default="0")
    try:
        value = float(value_str)
    except ValueError:
        print("  ⚠️  请输入有效数字")
        return None

    unit_options = [(k, f"{v} ({k})") for k, v in DATE_DIFF_UNITS.items()]
    unit = prompt_choice("请选择加减单位", unit_options, default_index=0)

    return {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "date_add",
        "base_field": base_field,
        "value": value,
        "unit": unit,
        "referenced_fields": [base_field],
    }


def _prompt_concat_cc(field_options):
    print("\n添加字符串拼接计算列:")
    print("  将多个字段或文本拼接成一个新值")

    key, label, width = _prompt_cc_base()

    separator = prompt_input("分隔符（留空表示无分隔）", required=False, default="")

    print("\n可用字段:")
    for i, (k, desc) in enumerate(field_options, 1):
        print(f"  {i}. {desc}")

    parts = []
    ref_fields = []
    while True:
        print(f"\n当前拼接部分: {len(parts)} 个")
        for i, p in enumerate(parts, 1):
            ptype = "字段" if p["type"] == "field" else "文本"
            print(f"  {i}. [{ptype}] {p['value']}")
        print("  0. 完成添加")

        add_choice = prompt_number("请选择添加类型 (1=字段, 2=固定文本, 0=完成)", min_value=0, max_value=2, default=0)
        if add_choice == 0:
            break
        elif add_choice == 1:
            field_key = prompt_input("请输入字段 key", required=True)
            parts.append({"type": "field", "value": field_key})
            ref_fields.append(field_key)
        elif add_choice == 2:
            text = prompt_input("请输入固定文本", required=True)
            parts.append({"type": "text", "value": text})

    if not parts:
        print("  ⚠️  至少需要添加一个拼接部分")
        return None

    return {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "concat",
        "parts": parts,
        "separator": separator,
        "referenced_fields": ref_fields,
    }


def _prompt_conditional_cc(field_options):
    print("\n添加条件表达式计算列:")
    print("  根据条件返回不同的值，如: IF 薪资 > 10000 THEN '高' ELSE '低'")

    key, label, width = _prompt_cc_base()

    print("\n可用字段:")
    for i, (k, desc) in enumerate(field_options, 1):
        print(f"  {i}. {desc}")

    cond_type = prompt_choice(
        "\n请选择条件类型",
        [
            ("compare", "比较运算（=, !=, >, >=, <, <=, contains）"),
            ("is_null", "为空判断"),
            ("not_null", "不为空判断"),
            ("range", "数值范围判断"),
            ("in_values", "值列表判断"),
        ],
        default_index=0,
    )

    condition = {"type": cond_type}

    if cond_type in ("compare", "range", "in_values"):
        field = prompt_input("判断的字段 key", required=True)
        condition["field"] = field

        if cond_type == "compare":
            operator = prompt_choice(
                "比较运算符",
                [
                    ("==", "等于 (==)"),
                    ("!=", "不等于 (!=)"),
                    (">", "大于 (>)"),
                    (">=", "大于等于 (>=)"),
                    ("<", "小于 (<)"),
                    ("<=", "小于等于 (<=)"),
                    ("contains", "包含"),
                ],
                default_index=0,
            )
            condition["operator"] = operator
            compare_value = prompt_input("比较值", required=True)
            try:
                compare_value = float(compare_value)
                if compare_value == int(compare_value):
                    compare_value = int(compare_value)
            except ValueError:
                pass
            condition["value"] = compare_value

        elif cond_type == "range":
            min_str = prompt_input("最小值（留空无下限）", required=False)
            max_str = prompt_input("最大值（留空无上限）", required=False)
            if min_str:
                condition["min"] = float(min_str)
            if max_str:
                condition["max"] = float(max_str)

        elif cond_type == "in_values":
            values_str = prompt_input("值列表（逗号分隔，如: 技术部,产品部）", required=True)
            condition["values"] = [v.strip() for v in values_str.split(",") if v.strip()]

    elif cond_type in ("is_null", "not_null"):
        field = prompt_input("判断的字段 key", required=True)
        condition["field"] = field

    true_value = prompt_input("条件为真时的返回值", required=True)
    false_value = prompt_input("条件为假时的返回值", default="")

    try:
        true_value = float(true_value)
        if true_value == int(true_value):
            true_value = int(true_value)
    except ValueError:
        pass
    try:
        false_value = float(false_value)
        if false_value == int(false_value):
            false_value = int(false_value)
    except ValueError:
        pass

    ref_fields = []
    if "field" in condition:
        ref_fields.append(condition["field"])

    return {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "conditional",
        "condition": condition,
        "true_value": true_value,
        "false_value": false_value,
        "referenced_fields": ref_fields,
    }


def _prompt_round_cc(field_options):
    print("\n添加取整计算列:")
    print("  对数值字段进行四舍五入或取整操作")

    key, label, width = _prompt_cc_base()

    field = prompt_input("要取整的字段 key", required=True)
    method = prompt_choice(
        "取整方式",
        [
            ("round", "四舍五入 (round)"),
            ("ceil", "向上取整 (ceil)"),
            ("floor", "向下取整 (floor)"),
        ],
        default_index=0,
    )
    precision = 0
    if method == "round":
        precision = prompt_number("保留小数位数", default=0, min_value=0, max_value=10)

    return {
        "key": key,
        "label": label,
        "width": width,
        "enabled": True,
        "formula_type": "round",
        "field": field,
        "method": method,
        "precision": precision,
        "referenced_fields": [field],
    }


def _delete_computed_column(computed_columns):
    if not computed_columns:
        print("  ⚠️  没有可删除的计算列")
        return
    for i, cc in enumerate(computed_columns, 1):
        print(f"  {i}. {describe_computed_column(cc)}")
    idx = prompt_number("请选择要删除的计算列序号", min_value=1, max_value=len(computed_columns)) - 1
    name = computed_columns[idx].get("label", computed_columns[idx].get("key", ""))
    if prompt_confirm(f"确定删除计算列 '{name}' 吗？", default=False):
        computed_columns.pop(idx)
        print(f"  ✅ 已删除计算列: {name}")


def _toggle_computed_column(computed_columns):
    if not computed_columns:
        print("  ⚠️  没有可操作的计算列")
        return
    for i, cc in enumerate(computed_columns, 1):
        enabled_str = "✓ 启用" if cc.get("enabled", True) else "✗ 禁用"
        print(f"  {i}. [{enabled_str}] {describe_computed_column(cc)}")
    idx = prompt_number("请选择要切换状态的计算列序号", min_value=1, max_value=len(computed_columns)) - 1
    computed_columns[idx]["enabled"] = not computed_columns[idx].get("enabled", True)
    status = "启用" if computed_columns[idx]["enabled"] else "禁用"
    print(f"  ✅ 已{status}计算列: {computed_columns[idx].get('label', '')}")
