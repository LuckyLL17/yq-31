import re
import copy
import json


def _flatten_dict(d, parent_key="", sep="."):
    items = []
    if isinstance(d, dict):
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(_flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                items.append((new_key, json.dumps(v, ensure_ascii=False)))
            else:
                items.append((new_key, v))
    else:
        items.append((parent_key, d))
    return dict(items)


VALIDATION_TYPE_NOT_NULL = "not_null"
VALIDATION_TYPE_FORMAT = "format"
VALIDATION_TYPE_RANGE = "range"
VALIDATION_TYPE_REGEX = "regex"

VALIDATION_TYPES = [
    VALIDATION_TYPE_NOT_NULL,
    VALIDATION_TYPE_FORMAT,
    VALIDATION_TYPE_RANGE,
    VALIDATION_TYPE_REGEX,
]

VALIDATION_TYPE_LABELS = {
    VALIDATION_TYPE_NOT_NULL: "非空校验",
    VALIDATION_TYPE_FORMAT: "格式校验",
    VALIDATION_TYPE_RANGE: "范围校验",
    VALIDATION_TYPE_REGEX: "正则表达式校验",
}

FORMAT_EMAIL = "email"
FORMAT_PHONE = "phone"
FORMAT_URL = "url"
FORMAT_DATE = "date"
FORMAT_NUMBER = "number"
FORMAT_ID_CARD = "id_card"

FORMAT_TYPES = [
    FORMAT_EMAIL,
    FORMAT_PHONE,
    FORMAT_URL,
    FORMAT_DATE,
    FORMAT_NUMBER,
    FORMAT_ID_CARD,
]

FORMAT_LABELS = {
    FORMAT_EMAIL: "邮箱地址",
    FORMAT_PHONE: "手机号码",
    FORMAT_URL: "URL网址",
    FORMAT_DATE: "日期 (YYYY-MM-DD)",
    FORMAT_NUMBER: "数字",
    FORMAT_ID_CARD: "身份证号",
}

FORMAT_PATTERNS = {
    FORMAT_EMAIL: r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
    FORMAT_PHONE: r'^1[3-9]\d{9}$',
    FORMAT_URL: r'^https?://[^\s<>]+$',
    FORMAT_DATE: r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}$',
    FORMAT_NUMBER: r'^-?\d+(\.\d+)?$',
    FORMAT_ID_CARD: r'^\d{17}[\dXx]$',
}

ON_FAIL_MARK = "mark"
ON_FAIL_SKIP = "skip"
ON_FAIL_ABORT = "abort"

ON_FAIL_ACTIONS = [ON_FAIL_MARK, ON_FAIL_SKIP, ON_FAIL_ABORT]

ON_FAIL_LABELS = {
    ON_FAIL_MARK: "标记（继续导出，高亮问题数据）",
    ON_FAIL_SKIP: "跳过（不导出该行数据）",
    ON_FAIL_ABORT: "中止（停止整个导出）",
}

MARK_COLOR = "FF0000"
MARK_BG_COLOR = "FFFF00"


class ValidationRule:
    def __init__(self, field, rule_type, params=None, on_fail=ON_FAIL_MARK, message=""):
        self.field = field
        self.rule_type = rule_type
        self.params = params or {}
        self.on_fail = on_fail
        self.message = message or self._default_message()

    def _default_message(self):
        if self.rule_type == VALIDATION_TYPE_NOT_NULL:
            return f"字段 '{self.field}' 不能为空"
        elif self.rule_type == VALIDATION_TYPE_FORMAT:
            fmt = self.params.get("format", "")
            fmt_label = FORMAT_LABELS.get(fmt, fmt)
            return f"字段 '{self.field}' 格式不正确，应为 {fmt_label}"
        elif self.rule_type == VALIDATION_TYPE_RANGE:
            min_val = self.params.get("min")
            max_val = self.params.get("max")
            if min_val is not None and max_val is not None:
                return f"字段 '{self.field}' 应在 [{min_val}, {max_val}] 范围内"
            elif min_val is not None:
                return f"字段 '{self.field}' 不能小于 {min_val}"
            elif max_val is not None:
                return f"字段 '{self.field}' 不能大于 {max_val}"
            return f"字段 '{self.field}' 范围校验失败"
        elif self.rule_type == VALIDATION_TYPE_REGEX:
            return f"字段 '{self.field}' 不匹配正则表达式"
        return f"字段 '{self.field}' 校验失败"

    def to_dict(self):
        return {
            "field": self.field,
            "rule_type": self.rule_type,
            "params": self.params,
            "on_fail": self.on_fail,
            "message": self.message,
        }

    @classmethod
    def from_dict(cls, d):
        return cls(
            field=d["field"],
            rule_type=d["rule_type"],
            params=d.get("params", {}),
            on_fail=d.get("on_fail", ON_FAIL_MARK),
            message=d.get("message", ""),
        )

    def __repr__(self):
        type_label = VALIDATION_TYPE_LABELS.get(self.rule_type, self.rule_type)
        on_fail_label = ON_FAIL_LABELS.get(self.on_fail, self.on_fail).split("（")[0]
        return f"[{type_label}] {self.field} → {on_fail_label}: {self.message}"


class ValidationError:
    def __init__(self, row_index, rule, value, field_key):
        self.row_index = row_index
        self.rule = rule
        self.value = value
        self.field_key = field_key

    def __repr__(self):
        val_str = str(self.value)[:30] if self.value is not None else "(空)"
        return f"第 {self.row_index + 1} 行, {self.rule}"


class ValidationResult:
    def __init__(self):
        self.errors = []
        self.marked_rows = set()
        self.skipped_rows = set()
        self.aborted = False
        self.abort_reason = None

    def add_error(self, error):
        self.errors.append(error)
        if error.rule.on_fail == ON_FAIL_MARK:
            self.marked_rows.add(error.row_index)
        elif error.rule.on_fail == ON_FAIL_SKIP:
            self.skipped_rows.add(error.row_index)
        elif error.rule.on_fail == ON_FAIL_ABORT:
            self.aborted = True
            self.abort_reason = error.rule.message

    @property
    def has_errors(self):
        return len(self.errors) > 0

    @property
    def marked_indices(self):
        return sorted(self.marked_rows)

    @property
    def skipped_indices(self):
        return sorted(self.skipped_rows)

    def get_errors_for_row(self, row_index):
        return [e for e in self.errors if e.row_index == row_index]

    def get_field_errors_for_row(self, row_index):
        result = {}
        for e in self.errors:
            if e.row_index == row_index:
                if e.field_key not in result:
                    result[e.field_key] = []
                result[e.field_key].append(e)
        return result

    def summary(self):
        lines = []
        total_errors = len(self.errors)
        marked = len(self.marked_rows)
        skipped = len(self.skipped_rows)

        lines.append(f"校验完成: 共发现 {total_errors} 个问题")

        if marked > 0:
            lines.append(f"  标记行: {marked} 行")
        if skipped > 0:
            lines.append(f"  跳过行: {skipped} 行")
        if self.aborted:
            lines.append(f"  已中止: {self.abort_reason}")

        by_rule_type = {}
        for e in self.errors:
            rt = e.rule.rule_type
            by_rule_type[rt] = by_rule_type.get(rt, 0) + 1

        if by_rule_type:
            lines.append("  按类型统计:")
            for rt, count in by_rule_type.items():
                label = VALIDATION_TYPE_LABELS.get(rt, rt)
                lines.append(f"    {label}: {count} 个")

        return "\n".join(lines)


def validate_value(value, rule):
    if rule.rule_type == VALIDATION_TYPE_NOT_NULL:
        return _validate_not_null(value)
    elif rule.rule_type == VALIDATION_TYPE_FORMAT:
        return _validate_format(value, rule.params)
    elif rule.rule_type == VALIDATION_TYPE_RANGE:
        return _validate_range(value, rule.params)
    elif rule.rule_type == VALIDATION_TYPE_REGEX:
        return _validate_regex(value, rule.params)
    return True


def _validate_not_null(value):
    if value is None:
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def _validate_format(value, params):
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return True

    fmt = params.get("format", "")
    pattern = FORMAT_PATTERNS.get(fmt)
    if pattern is None:
        return True

    str_val = str(value)
    try:
        return re.search(pattern, str_val) is not None
    except re.error:
        return True


def _validate_range(value, params):
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return True

    min_val = params.get("min")
    max_val = params.get("max")

    try:
        num_val = float(value)
    except (ValueError, TypeError):
        return False

    if min_val is not None and num_val < float(min_val):
        return False
    if max_val is not None and num_val > float(max_val):
        return False
    return True


def _validate_regex(value, params):
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return True

    pattern = params.get("pattern", "")
    if not pattern:
        return True

    str_val = str(value)
    try:
        return re.search(pattern, str_val) is not None
    except re.error:
        return True


def validate_data(data, rules, headers=None):
    result = ValidationResult()

    if not rules:
        return result

    for row_idx, item in enumerate(data):
        if not isinstance(item, dict):
            continue

        flat = _flatten_dict(item)

        for rule in rules:
            value = flat.get(rule.field, "")
            if not validate_value(value, rule):
                error = ValidationError(
                    row_index=row_idx,
                    rule=rule,
                    value=value,
                    field_key=rule.field,
                )
                result.add_error(error)

                if result.aborted:
                    return result

    return result


def apply_validation_to_export(data, result, headers):
    if not result.has_errors:
        return data, []

    if result.aborted:
        return None, result.skipped_indices

    valid_data = []
    removed_indices = set()

    for idx, item in enumerate(data):
        if idx in result.skipped_rows:
            removed_indices.add(idx)
            continue
        valid_data.append(item)

    return valid_data, sorted(removed_indices)


def rules_from_config(config):
    rules = []
    for rule_dict in config.get("validation_rules", []):
        rules.append(ValidationRule.from_dict(rule_dict))
    return rules


def rules_to_config(rules):
    return [r.to_dict() for r in rules]
