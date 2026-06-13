import re
import math
from datetime import datetime, date


FORMULA_TYPES = {
    "arithmetic": "算术运算（加减乘除、取模等）",
    "date_diff": "日期差值（计算两个日期之间的差）",
    "date_add": "日期加减（日期加上或减去天数/月数/年数）",
    "concat": "字符串拼接",
    "conditional": "条件表达式（根据条件返回不同值）",
    "round": "四舍五入 / 取整",
}

DATE_DIFF_UNITS = {
    "days": "天",
    "months": "月",
    "years": "年",
}

RESULT_TYPES = {
    "number": "数字",
    "string": "文本",
    "date": "日期",
}

SAFE_FUNCTIONS = {
    "abs": abs,
    "round": round,
    "min": min,
    "max": max,
    "int": int,
    "float": float,
    "str": str,
    "len": len,
    "math": math,
    "math_ceil": math.ceil,
    "math_floor": math.floor,
    "math_sqrt": math.sqrt,
    "math_log": math.log,
    "math_log10": math.log10,
    "math_pow": math.pow,
}


def _parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, (datetime, date)):
        return value if isinstance(value, date) else value.date()
    if isinstance(value, (int, float)):
        try:
            return date.fromordinal(int(value) + 693594)
        except (ValueError, OverflowError):
            return None
    try:
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
            try:
                return datetime.strptime(str(value), fmt).date()
            except ValueError:
                continue
    except Exception:
        pass
    return None


def _date_diff(start_date, end_date, unit="days"):
    if start_date is None or end_date is None:
        return None

    d1 = _parse_date(start_date)
    d2 = _parse_date(end_date)
    if d1 is None or d2 is None:
        return None

    delta = d2 - d1
    total_days = delta.days

    if unit == "days":
        return total_days
    elif unit == "months":
        months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
        day_adjust = d2.day - d1.day
        if day_adjust < 0:
            months -= 1
        return months
    elif unit == "years":
        years = d2.year - d1.year
        if (d2.month, d2.day) < (d1.month, d1.day):
            years -= 1
        return years
    return total_days


def _date_add(base_date, value, unit="days"):
    if base_date is None:
        return None

    d = _parse_date(base_date)
    if d is None:
        return None

    if unit == "days":
        from datetime import timedelta
        return d + timedelta(days=int(value))
    elif unit == "months":
        month = d.month + int(value)
        year = d.year + (month - 1) // 12
        month = (month - 1) % 12 + 1
        import calendar
        max_day = calendar.monthrange(year, month)[1]
        day = min(d.day, max_day)
        return date(year, month, day)
    elif unit == "years":
        year = d.year + int(value)
        import calendar
        max_day = calendar.monthrange(year, d.month)[1]
        day = min(d.day, max_day)
        return date(year, d.month, day)
    return d


def _to_number(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def evaluate_arithmetic(formula, field_values):
    expr = formula
    field_keys = sorted(field_values.keys(), key=len, reverse=True)
    for key in field_keys:
        val = field_values[key]
        num = _to_number(val)
        if num is not None:
            expr = expr.replace(key, str(num))
        else:
            expr = expr.replace(key, "0")

    safe_dict = {
        "__builtins__": {},
        "abs": abs,
        "round": round,
        "min": min,
        "max": max,
        "int": int,
        "float": float,
        "ceil": math.ceil,
        "floor": math.floor,
        "sqrt": math.sqrt,
        "log": math.log,
        "log10": math.log10,
        "pow": math.pow,
    }

    try:
        result = eval(expr, safe_dict)
        if isinstance(result, float) and result == int(result):
            return int(result)
        return result
    except Exception:
        return None


def evaluate_date_diff(formula_config, field_values, now=None):
    start_field = formula_config.get("start_field", "")
    end_field = formula_config.get("end_field", "")
    unit = formula_config.get("unit", "days")
    use_current_date = formula_config.get("use_current_date", False)

    start_val = field_values.get(start_field)
    end_val = field_values.get(end_field)

    if use_current_date:
        end_val = now or date.today()

    return _date_diff(start_val, end_val, unit)


def evaluate_date_add(formula_config, field_values):
    base_field = formula_config.get("base_field", "")
    value = formula_config.get("value", 0)
    unit = formula_config.get("unit", "days")

    base_val = field_values.get(base_field)
    if isinstance(value, str) and value in field_values:
        value = _to_number(field_values[value])
        if value is None:
            value = 0

    return _date_add(base_val, value, unit)


def evaluate_concat(formula_config, field_values):
    parts = formula_config.get("parts", [])
    separator = formula_config.get("separator", "")

    result_parts = []
    for part in parts:
        if part.get("type") == "field":
            val = field_values.get(part.get("value", ""), "")
            result_parts.append(str(val) if val is not None else "")
        elif part.get("type") == "text":
            result_parts.append(str(part.get("value", "")))

    return separator.join(result_parts)


def evaluate_conditional(formula_config, field_values):
    condition = formula_config.get("condition", {})
    true_value = formula_config.get("true_value", "")
    false_value = formula_config.get("false_value", "")

    cond_type = condition.get("type", "compare")
    field = condition.get("field", "")
    field_val = field_values.get(field)

    if cond_type == "compare":
        operator = condition.get("operator", "==")
        compare_value = condition.get("value")

        num_val = _to_number(field_val)
        num_cmp = _to_number(compare_value)

        if num_val is not None and num_cmp is not None:
            if operator == "==":
                matched = num_val == num_cmp
            elif operator == "!=":
                matched = num_val != num_cmp
            elif operator == ">":
                matched = num_val > num_cmp
            elif operator == ">=":
                matched = num_val >= num_cmp
            elif operator == "<":
                matched = num_val < num_cmp
            elif operator == "<=":
                matched = num_val <= num_cmp
            else:
                matched = False
        else:
            str_val = str(field_val) if field_val is not None else ""
            str_cmp = str(compare_value) if compare_value is not None else ""
            if operator == "==":
                matched = str_val == str_cmp
            elif operator == "!=":
                matched = str_val != str_cmp
            elif operator == "contains":
                matched = str_cmp in str_val
            else:
                matched = False

    elif cond_type == "is_null":
        matched = field_val is None or field_val == ""
    elif cond_type == "not_null":
        matched = field_val is not None and field_val != ""
    elif cond_type == "range":
        num_val = _to_number(field_val)
        min_val = condition.get("min")
        max_val = condition.get("max")
        if num_val is not None:
            matched = True
            if min_val is not None:
                matched = matched and num_val >= min_val
            if max_val is not None:
                matched = matched and num_val < max_val
        else:
            matched = False
    elif cond_type == "in_values":
        values = condition.get("values", [])
        matched = field_val in values
    else:
        matched = False

    result_val = true_value if matched else false_value

    if isinstance(result_val, str) and result_val in field_values:
        return field_values[result_val]

    return result_val


def evaluate_round(formula_config, field_values):
    field = formula_config.get("field", "")
    precision = formula_config.get("precision", 0)
    method = formula_config.get("method", "round")

    val = _to_number(field_values.get(field))
    if val is None:
        return None

    if method == "round":
        return round(val, precision)
    elif method == "ceil":
        return math.ceil(val)
    elif method == "floor":
        return math.floor(val)
    return round(val, precision)


def evaluate_computed_column(col_config, item, extract_value_func, now=None):
    formula_type = col_config.get("formula_type", "arithmetic")
    formula = col_config.get("formula", "")

    field_values = {}
    referenced_fields = col_config.get("referenced_fields", [])
    for field_key in referenced_fields:
        field_values[field_key] = extract_value_func(item, field_key)

    if formula_type == "arithmetic":
        return evaluate_arithmetic(formula, field_values)
    elif formula_type == "date_diff":
        return evaluate_date_diff(col_config, field_values, now)
    elif formula_type == "date_add":
        return evaluate_date_add(col_config, field_values)
    elif formula_type == "concat":
        return evaluate_concat(col_config, field_values)
    elif formula_type == "conditional":
        return evaluate_conditional(col_config, field_values)
    elif formula_type == "round":
        return evaluate_round(col_config, field_values)
    return None


def apply_computed_columns(data, headers, config, extract_value_func):
    computed_columns = config.get("computed_columns", [])
    if not computed_columns:
        return data, headers

    now = date.today()
    new_headers = list(headers)

    for cc in computed_columns:
        if not cc.get("enabled", True):
            continue
        new_headers.append({
            "key": cc["key"],
            "label": cc.get("label", cc["key"]),
            "width": cc.get("width", 15),
        })

    computed_cache = {}
    for item in data:
        if not isinstance(item, dict):
            continue
        row_values = {}
        for cc in computed_columns:
            if not cc.get("enabled", True):
                continue
            result = evaluate_computed_column(cc, item, extract_value_func, now)
            row_values[cc["key"]] = result
        computed_cache[id(item)] = row_values

    return computed_cache, new_headers


def validate_computed_column(col_config, index):
    errors = []

    if not col_config.get("key"):
        errors.append(f"第 {index + 1} 个计算列缺少 key 属性")

    if not col_config.get("label"):
        errors.append(f"第 {index + 1} 个计算列缺少 label 属性")

    formula_type = col_config.get("formula_type", "arithmetic")
    if formula_type not in FORMULA_TYPES:
        errors.append(f"第 {index + 1} 个计算列的 formula_type 无效，应为: {', '.join(FORMULA_TYPES.keys())}")

    if formula_type == "arithmetic":
        formula = col_config.get("formula", "")
        if not formula:
            errors.append(f"第 {index + 1} 个算术计算列缺少 formula 属性")
        else:
            ref_fields = col_config.get("referenced_fields", [])
            if not ref_fields:
                errors.append(f"第 {index + 1} 个算术计算列缺少 referenced_fields 属性")

    elif formula_type == "date_diff":
        start_field = col_config.get("start_field", "")
        end_field = col_config.get("end_field", "")
        use_current = col_config.get("use_current_date", False)
        if not start_field:
            errors.append(f"第 {index + 1} 个日期差值计算列缺少 start_field 属性")
        if not end_field and not use_current:
            errors.append(f"第 {index + 1} 个日期差值计算列需指定 end_field 或 use_current_date")
        unit = col_config.get("unit", "days")
        if unit not in DATE_DIFF_UNITS:
            errors.append(f"第 {index + 1} 个日期差值计算列的 unit 无效，应为: {', '.join(DATE_DIFF_UNITS.keys())}")

    elif formula_type == "date_add":
        base_field = col_config.get("base_field", "")
        if not base_field:
            errors.append(f"第 {index + 1} 个日期加减计算列缺少 base_field 属性")
        unit = col_config.get("unit", "days")
        if unit not in DATE_DIFF_UNITS:
            errors.append(f"第 {index + 1} 个日期加减计算列的 unit 无效")

    elif formula_type == "concat":
        parts = col_config.get("parts", [])
        if not parts:
            errors.append(f"第 {index + 1} 个字符串拼接计算列缺少 parts 属性")

    elif formula_type == "conditional":
        condition = col_config.get("condition", {})
        if not condition:
            errors.append(f"第 {index + 1} 个条件表达式计算列缺少 condition 属性")

    elif formula_type == "round":
        field = col_config.get("field", "")
        if not field:
            errors.append(f"第 {index + 1} 个取整计算列缺少 field 属性")

    return errors


def describe_computed_column(col_config):
    formula_type = col_config.get("formula_type", "arithmetic")
    label = col_config.get("label", col_config.get("key", ""))
    type_label = FORMULA_TYPES.get(formula_type, formula_type).split("（")[0]

    if formula_type == "arithmetic":
        formula = col_config.get("formula", "")
        return f"{label} = {formula} [{type_label}]"
    elif formula_type == "date_diff":
        start = col_config.get("start_field", "?")
        end = col_config.get("end_field", "")
        unit = col_config.get("unit", "days")
        use_current = col_config.get("use_current_date", False)
        end_display = "当前日期" if use_current else end
        return f"{label} = {end_display} - {start} ({DATE_DIFF_UNITS.get(unit, unit)}) [{type_label}]"
    elif formula_type == "date_add":
        base = col_config.get("base_field", "?")
        value = col_config.get("value", 0)
        unit = col_config.get("unit", "days")
        return f"{label} = {base} + {value}{DATE_DIFF_UNITS.get(unit, unit)} [{type_label}]"
    elif formula_type == "concat":
        parts = col_config.get("parts", [])
        sep = col_config.get("separator", "")
        parts_desc = sep.join(
            "{" + p.get("value", "") + "}" if p.get("type") == "field" else p.get("value", "")
            for p in parts
        )
        return f"{label} = \"{parts_desc}\" [{type_label}]"
    elif formula_type == "conditional":
        condition = col_config.get("condition", {})
        true_v = col_config.get("true_value", "")
        false_v = col_config.get("false_value", "")
        field = condition.get("field", "?")
        op = condition.get("operator", "==")
        val = condition.get("value", "")
        return f"{label} = IF {field} {op} {val} THEN {true_v} ELSE {false_v} [{type_label}]"
    elif formula_type == "round":
        field = col_config.get("field", "?")
        precision = col_config.get("precision", 0)
        method = col_config.get("method", "round")
        return f"{label} = {method}({field}, {precision}) [{type_label}]"
    return f"{label} [{type_label}]"
