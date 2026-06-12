from prompts import (
    clear_screen,
    print_title,
    print_step,
    prompt_input,
    prompt_confirm,
    prompt_choice,
    prompt_multi_select,
    prompt_number,
)
from data_validator import (
    ValidationRule,
    VALIDATION_TYPES,
    VALIDATION_TYPE_LABELS,
    FORMAT_TYPES,
    FORMAT_LABELS,
    ON_FAIL_ACTIONS,
    ON_FAIL_LABELS,
    ON_FAIL_MARK,
    ON_FAIL_SKIP,
    ON_FAIL_ABORT,
    VALIDATION_TYPE_NOT_NULL,
    VALIDATION_TYPE_FORMAT,
    VALIDATION_TYPE_RANGE,
    VALIDATION_TYPE_REGEX,
    rules_to_config,
    rules_from_config,
)


class ValidationWizard:
    def __init__(self, config=None, available_fields=None):
        self.config = config or {}
        self.available_fields = available_fields or []
        self.rules = rules_from_config(self.config)

    def run(self):
        while True:
            clear_screen()
            print_title("数据校验规则配置")

            if self.rules:
                print(f"\n当前已配置 {len(self.rules)} 条校验规则:\n")
                for i, rule in enumerate(self.rules, 1):
                    on_fail_short = ON_FAIL_LABELS.get(rule.on_fail, rule.on_fail).split("（")[0]
                    print(f"  {i}. {rule}")
                print()
            else:
                print("\n  当前未配置任何校验规则\n")

            print("操作选项:")
            print("  1. 添加校验规则")
            print("  2. 修改校验规则")
            print("  3. 删除校验规则")
            print("  4. 清空所有规则")
            print("  5. 预览校验效果（需先加载数据）")
            print("  0. 完成配置")

            choice = prompt_number(
                "\n请选择操作",
                min_value=0,
                max_value=5,
                default=0,
            )

            if choice == 0:
                self._save_to_config()
                break
            elif choice == 1:
                self._add_rule()
            elif choice == 2:
                if self.rules:
                    self._edit_rule()
                else:
                    print("\n  ⚠️  没有可修改的规则")
                    input("\n按回车继续...")
            elif choice == 3:
                if self.rules:
                    self._delete_rule()
                else:
                    print("\n  ⚠️  没有可删除的规则")
                    input("\n按回车继续...")
            elif choice == 4:
                if self.rules:
                    if prompt_confirm("确定要清空所有校验规则吗？", default=False):
                        self.rules.clear()
                        print("\n  ✅ 已清空所有校验规则")
                        input("\n按回车继续...")
                else:
                    print("\n  ⚠️  没有可清空的规则")
                    input("\n按回车继续...")
            elif choice == 5:
                self._preview_validation()

        return self.config

    def _save_to_config(self):
        self.config["validation_rules"] = rules_to_config(self.rules)
        print(f"\n✅ 已保存 {len(self.rules)} 条校验规则")

    def _select_field(self):
        if not self.available_fields:
            field = prompt_input("请输入要校验的字段 key", required=True)
            return field

        options = [(h["key"], f"{h.get('label', h['key'])} ({h['key']})") for h in self.available_fields]
        options.append(("__custom__", "自定义字段..."))
        choice = prompt_choice("请选择要校验的字段", options, default_index=0)

        if choice == "__custom__":
            return prompt_input("请输入字段 key", required=True)
        return choice

    def _select_rule_type(self):
        options = [(rt, VALIDATION_TYPE_LABELS[rt]) for rt in VALIDATION_TYPES]
        return prompt_choice("请选择校验类型", options, default_index=0)

    def _select_on_fail(self):
        options = [(a, ON_FAIL_LABELS[a]) for a in ON_FAIL_ACTIONS]
        return prompt_choice("校验不通过时的处理方式", options, default_index=0)

    def _add_rule(self):
        clear_screen()
        print_title("添加校验规则")

        field = self._select_field()
        if not field:
            return

        rule_type = self._select_rule_type()
        params = self._configure_params(rule_type)

        on_fail = self._select_on_fail()

        default_msg = ValidationRule(field, rule_type, params, on_fail).message
        message = prompt_input("请输入校验失败提示信息", default=default_msg, required=False)

        rule = ValidationRule(field, rule_type, params, on_fail, message)
        self.rules.append(rule)

        print(f"\n✅ 已添加校验规则: {rule}")
        input("\n按回车继续...")

    def _configure_params(self, rule_type):
        params = {}

        if rule_type == VALIDATION_TYPE_FORMAT:
            options = [(ft, FORMAT_LABELS[ft]) for ft in FORMAT_TYPES]
            fmt = prompt_choice("请选择格式类型", options, default_index=0)
            params["format"] = fmt

        elif rule_type == VALIDATION_TYPE_RANGE:
            has_min = prompt_confirm("是否设置最小值？", default=True)
            if has_min:
                params["min"] = prompt_number("请输入最小值", default=0, required=True)
            has_max = prompt_confirm("是否设置最大值？", default=True)
            if has_max:
                params["max"] = prompt_number("请输入最大值", default=100, required=True)
            if "min" not in params and "max" not in params:
                params["min"] = 0
                print("  ⚠️  至少需要设置最小值或最大值之一，已默认设置最小值为 0")

        elif rule_type == VALIDATION_TYPE_REGEX:
            params["pattern"] = prompt_input("请输入正则表达式", required=True)
            case_sensitive = prompt_confirm("是否区分大小写？", default=True)
            params["case_sensitive"] = case_sensitive
            if not case_sensitive:
                params["pattern"] = "(?i)" + params["pattern"]

        return params

    def _edit_rule(self):
        clear_screen()
        print_title("修改校验规则")

        print("\n当前校验规则:")
        for i, rule in enumerate(self.rules, 1):
            print(f"  {i}. {rule}")

        idx = prompt_number(
            f"\n请选择要修改的规则编号 (1-{len(self.rules)})",
            min_value=1,
            max_value=len(self.rules),
        ) - 1

        rule = self.rules[idx]
        print(f"\n当前规则: {rule}")
        print("\n可修改的项目:")
        print("  1. 校验类型和参数")
        print("  2. 处理方式")
        print("  3. 提示信息")
        print("  0. 取消")

        choice = prompt_number("请选择要修改的项目", min_value=0, max_value=3, default=0)

        if choice == 1:
            rule_type = self._select_rule_type()
            params = self._configure_params(rule_type)
            rule.rule_type = rule_type
            rule.params = params
            rule.message = ValidationRule(rule.field, rule_type, params, rule.on_fail).message
        elif choice == 2:
            rule.on_fail = self._select_on_fail()
        elif choice == 3:
            rule.message = prompt_input("请输入新的提示信息", default=rule.message, required=False)

        print(f"\n✅ 规则已更新: {rule}")
        input("\n按回车继续...")

    def _delete_rule(self):
        clear_screen()
        print_title("删除校验规则")

        print("\n当前校验规则:")
        for i, rule in enumerate(self.rules, 1):
            print(f"  {i}. {rule}")

        idx = prompt_number(
            f"\n请选择要删除的规则编号 (1-{len(self.rules)})",
            min_value=1,
            max_value=len(self.rules),
        ) - 1

        rule = self.rules[idx]
        if prompt_confirm(f"确定要删除规则: {rule}？", default=False):
            self.rules.pop(idx)
            print("\n✅ 已删除该规则")

        input("\n按回车继续...")

    def _preview_validation(self):
        json_path = self.config.get("json_file_path", "")
        if not json_path:
            print("\n  ⚠️  未设置 JSON 文件路径，无法预览校验效果")
            input("\n按回车继续...")
            return

        try:
            from json_to_excel import load_json
            from data_validator import validate_data

            data = load_json(json_path)
            print(f"\n已加载 {len(data)} 条数据，正在执行校验...")

            result = validate_data(data, self.rules)
            print(f"\n{result.summary()}")

            if result.has_errors:
                print("\n校验详情 (最多显示 20 条):")
                for i, error in enumerate(result.errors[:20]):
                    val_str = str(error.value)[:30] if error.value is not None else "(空)"
                    on_fail_str = ON_FAIL_LABELS.get(error.rule.on_fail, "").split("（")[0]
                    print(f"  行 {error.row_index + 1}: {error.rule.message} | 值: {val_str} | 处理: {on_fail_str}")

                if len(result.errors) > 20:
                    print(f"  ... 还有 {len(result.errors) - 20} 个问题未显示")

        except Exception as e:
            print(f"\n  ❌ 预览校验失败: {e}")

        input("\n按回车继续...")


def run_validation_wizard(config, available_fields=None):
    wizard = ValidationWizard(config, available_fields)
    return wizard.run()
