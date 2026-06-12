import os
import copy

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
)
from config_manager import (
    get_default_config,
    load_config,
    save_config,
    validate_config,
)
from style_template_manager import (
    apply_template_to_config,
    list_templates,
    get_border_style_name,
)
from task_manager import (
    list_batch_tasks,
    load_batch_task,
    delete_batch_task,
    create_batch_task,
    BATCH_STATUS_PENDING,
    BATCH_STATUS_RUNNING,
    BATCH_STATUS_PAUSED,
    BATCH_STATUS_COMPLETED,
    BATCH_STATUS_FAILED,
)
from batch_processor import BatchProcessor, resume_batch_process
from report_generator import generate_report, generate_json_report, print_summary


TOTAL_STEPS = 7


class BatchConfigWizard:
    def __init__(self, config=None):
        self.config = config or get_default_config()
        self.source_type = "files"
        self.sources = []
        self.output_dir = "./output/batch"
        self.options = {
            "skip_existing": False,
            "generate_report": True,
            "report_format": "txt",
        }

    def run(self):
        while True:
            clear_screen()
            print_title("JSON 批量转 Excel - 配置向导")
            print("\n欢迎使用批量转换向导！请按照步骤进行配置。\n")

            self.current_step = 0
            self.total_steps = TOTAL_STEPS

            self._next_step()
            self.step_select_source_type()

            self._next_step()
            self.step_select_sources()

            self._next_step()
            self.step_config_output()

            self._next_step()
            self.step_config_style()

            self._next_step()
            self.step_config_split()

            self._next_step()
            self.step_config_options()

            self._next_step()
            if self.step_review_and_confirm():
                break

        return self._get_batch_config()

    def _next_step(self):
        self.current_step += 1

    def step_select_source_type(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "选择来源类型")

        choice = prompt_choice(
            "请选择JSON来源方式",
            [
                ("directory", "选择目录（自动扫描目录下所有JSON文件）"),
                ("files", "选择多个文件（手动指定多个JSON文件）"),
            ],
            default_index=0,
        )

        self.source_type = choice
        print(f"\n✅ 已选择: {'目录模式' if choice == 'directory' else '多文件模式'}")
        input("\n按回车继续...")

    def step_select_sources(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "选择来源")

        self.sources = []

        if self.source_type == "directory":
            while True:
                dir_path = prompt_file_path(
                    "请输入JSON文件目录路径",
                    default="./data",
                    must_exist=True,
                )
                if os.path.isdir(dir_path):
                    self.sources.append(os.path.abspath(dir_path))
                    count = self._count_json_in_dir(dir_path)
                    print(f"\n✅ 已添加目录: {os.path.abspath(dir_path)}")
                    print(f"   该目录下找到 {count} 个JSON文件")

                    if not prompt_confirm("是否添加更多目录？", default=False):
                        break
                else:
                    print("  ⚠️  路径不是有效的目录")
        else:
            print("\n请输入多个JSON文件路径（每行一个，输入空行结束）")
            while True:
                file_path = prompt_input(
                    f"文件 {len(self.sources) + 1}",
                    required=False,
                )
                if not file_path:
                    if len(self.sources) > 0:
                        break
                    print("  ⚠️  至少需要添加一个文件")
                    continue

                if os.path.isfile(file_path) and file_path.lower().endswith(".json"):
                    self.sources.append(os.path.abspath(file_path))
                    print(f"  ✅ 已添加: {os.path.basename(file_path)}")
                else:
                    print("  ⚠️  文件不存在或不是JSON格式")

        print(f"\n✅ 共选择 {len(self.sources)} 个来源")
        total_files = self._count_total_files()
        print(f"   预计处理 {total_files} 个JSON文件")
        input("\n按回车继续...")

    def step_config_output(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置输出")

        self.output_dir = prompt_file_path(
            "请输入Excel输出目录",
            default=self.output_dir,
            must_exist=False,
        )

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        print(f"\n✅ 输出目录: {os.path.abspath(self.output_dir)}")
        input("\n按回车继续...")

    def step_config_style(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置导出样式")

        while True:
            clear_screen()
            print_step(self.current_step, self.total_steps, "配置导出样式")

            print("\n样式配置:")
            print("  1. 应用样式模板")
            print("  2. 配置表头样式")
            print("  3. 配置数据行样式")
            print("  0. 完成样式配置")

            choice = prompt_number(
                "\n请选择操作",
                min_value=0,
                max_value=3,
                default=0,
            )

            if choice == 0:
                break
            elif choice == 1:
                self._apply_template_interactive()
                input("\n按回车继续...")
            elif choice == 2:
                self._config_header_style()
                input("\n按回车继续...")
            elif choice == 3:
                self._config_data_style()
                input("\n按回车继续...")

        print("\n✅ 样式配置完成")
        input("\n按回车继续...")

    def step_config_split(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置 Excel 工作表拆分")

        print("\n工作表拆分可以将数据按指定字段的值分组，分别导出到不同的工作表。")
        print("例如：按 '部门' 字段拆分，每个部门自动生成一个独立的工作表。")

        split_cfg = self.config.setdefault("split_config", {})

        if not prompt_confirm("\n是否启用工作表拆分功能？", default=split_cfg.get("enabled", False)):
            split_cfg["enabled"] = False
            print("\n✅ 已禁用工作表拆分")
            input("\n按回车继续...")
            return

        split_cfg["enabled"] = True

        from wizard import ConfigWizard
        dummy = ConfigWizard(self.config)
        dummy.current_step = self.current_step
        dummy.total_steps = self.total_steps
        dummy.available_fields = []

        split_cfg["split_field"] = prompt_input(
            "\n请输入用于拆分的字段 key（例如: department, age, salary）",
            default=split_cfg.get("split_field", ""),
            required=True,
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
            dummy._config_split_range_groups(split_cfg)
        elif split_rule == "by_custom":
            dummy._config_split_custom_rules(split_cfg)

        print("\n✅ 拆分配置完成")
        input("\n按回车继续...")

    def step_config_options(self):
        clear_screen()
        print_step(self.current_step, self.total_steps, "配置批量选项")

        self.options["skip_existing"] = prompt_confirm(
            "是否跳过已存在的Excel文件？",
            default=self.options.get("skip_existing", False),
        )

        self.options["generate_report"] = prompt_confirm(
            "是否生成处理报告？",
            default=self.options.get("generate_report", True),
        )

        if self.options["generate_report"]:
            self.options["report_format"] = prompt_choice(
                "报告格式",
                [
                    ("txt", "文本格式 (.txt)"),
                    ("json", "JSON格式 (.json)"),
                    ("both", "同时生成两种格式"),
                ],
                default_index=0,
            )

        self.config["auto_detect_headers"] = prompt_confirm(
            "是否自动检测并添加新字段？",
            default=self.config.get("auto_detect_headers", True),
        )

        print("\n✅ 选项配置完成")
        input("\n按回车继续...")

    def step_review_and_confirm(self):
        clear_screen()
        print_title("批量配置确认")

        total_files = self._count_total_files()

        print("\n当前配置摘要:")
        print(f"\n  来源类型: {'目录' if self.source_type == 'directory' else '多文件'}")
        print(f"  来源数量: {len(self.sources)} 个")
        print(f"  预计文件数: {total_files} 个")
        print(f"  输出目录: {os.path.abspath(self.output_dir)}")
        print(f"  自动检测字段: {'是' if self.config['auto_detect_headers'] else '否'}")
        print(f"  跳过已存在: {'是' if self.options['skip_existing'] else '否'}")
        print(f"  生成报告: {'是' if self.options['generate_report'] else '否'}")

        if self.options["generate_report"]:
            fmt_map = {"txt": "文本格式", "json": "JSON格式", "both": "两种格式"}
            print(f"  报告格式: {fmt_map.get(self.options['report_format'], '文本格式')}")

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

        print("\n来源列表:")
        for i, src in enumerate(self.sources, 1):
            print(f"  {i}. {src}")

        print("\n操作选项:")
        print("  1. 开始批量转换")
        print("  2. 重新配置")
        print("  3. 取消")

        choice = prompt_choice(
            "请选择操作",
            [
                ("start", "开始批量转换"),
                ("retry", "重新配置"),
                ("cancel", "取消"),
            ],
            default_index=0,
        )

        if choice == "start":
            return True
        elif choice == "retry":
            return False
        else:
            raise SystemExit(0)

    def _get_batch_config(self):
        return {
            "source_type": self.source_type,
            "sources": self.sources,
            "output_dir": self.output_dir,
            "config": self.config,
            "options": self.options,
        }

    def _count_json_in_dir(self, dir_path):
        count = 0
        for root, _, files in os.walk(dir_path):
            for f in files:
                if f.lower().endswith(".json"):
                    count += 1
        return count

    def _count_total_files(self):
        if self.source_type == "directory":
            total = 0
            for src in self.sources:
                total += self._count_json_in_dir(src)
            return total
        else:
            return len(self.sources)

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

    def _config_header_style(self):
        from wizard import ConfigWizard
        dummy_wizard = ConfigWizard(self.config)
        dummy_wizard._config_header_style()

    def _config_data_style(self):
        from wizard import ConfigWizard
        dummy_wizard = ConfigWizard(self.config)
        self.config["style_alt_rows"] = prompt_confirm(
            "是否启用隔行变色？",
            default=self.config.get("style_alt_rows", True),
        )
        dummy_wizard._config_data_style()


def run_batch_wizard(config=None):
    wizard = BatchConfigWizard(config)
    batch_config = wizard.run()

    batch_task = create_batch_task(
        source_type=batch_config["source_type"],
        sources=batch_config["sources"],
        output_dir=batch_config["output_dir"],
        config=batch_config["config"],
        options=batch_config["options"],
    )

    print(f"\n批量任务已创建，任务ID: {batch_task['batch_id']}")
    print(f"共发现 {batch_task['total_files']} 个JSON文件")

    if prompt_confirm("\n是否立即开始批量转换？", default=True):
        processor = BatchProcessor(batch_task=batch_task)
        success, result = processor.run(config=batch_config["config"])

        if batch_config["options"].get("generate_report"):
            fmt = batch_config["options"].get("report_format", "txt")
            if fmt in ("txt", "both"):
                report_path = generate_report(batch_task)
                print(f"\n📄 文本报告已生成: {os.path.abspath(report_path)}")
            if fmt in ("json", "both"):
                report_path = generate_json_report(batch_task)
                print(f"📄 JSON报告已生成: {os.path.abspath(report_path)}")

        print_summary(batch_task)

    return batch_task


def show_task_manager():
    while True:
        clear_screen()
        print_title("批量任务管理")

        tasks = list_batch_tasks()

        if not tasks:
            print("\n暂无批量任务记录")
            input("\n按回车返回...")
            return

        print(f"\n共 {len(tasks)} 个任务记录：\n")

        for i, task in enumerate(tasks, 1):
            status_icon = _get_batch_status_icon(task["status"])
            status_text = _format_batch_status(task["status"])
            created = task["created_at"][:19] if task["created_at"] else "未知"
            print(f"  {i:2d}. {status_icon} 任务ID: {task['batch_id']}")
            print(f"      状态: {status_text} | 文件数: {task['total_files']} "
                  f"| 成功: {task['completed_count']} | 失败: {task['failed_count']}")
            print(f"      创建时间: {created} | 输出目录: {task['output_dir']}")
            print()

        print("操作选项:")
        print("  1. 恢复任务")
        print("  2. 查看详情")
        print("  3. 删除任务")
        print("  0. 返回")

        choice = prompt_number("\n请选择操作", min_value=0, max_value=3, default=0)

        if choice == 0:
            return
        elif choice == 1:
            task_idx = prompt_number("选择要恢复的任务序号", min_value=1, max_value=len(tasks)) - 1
            batch_id = tasks[task_idx]["batch_id"]
            _handle_resume_task(batch_id)
        elif choice == 2:
            task_idx = prompt_number("选择要查看的任务序号", min_value=1, max_value=len(tasks)) - 1
            batch_id = tasks[task_idx]["batch_id"]
            _show_task_detail(batch_id)
        elif choice == 3:
            task_idx = prompt_number("选择要删除的任务序号", min_value=1, max_value=len(tasks)) - 1
            batch_id = tasks[task_idx]["batch_id"]
            if prompt_confirm(f"确定要删除任务 {batch_id} 吗？", default=False):
                delete_batch_task(batch_id)
                print(f"\n✅ 任务 {batch_id} 已删除")
                input("\n按回车继续...")


def _handle_resume_task(batch_id):
    clear_screen()
    print_title(f"恢复任务 - {batch_id}")

    batch_task = load_batch_task(batch_id)
    if not batch_task:
        print("\n❌ 任务不存在")
        input("\n按回车返回...")
        return

    if batch_task["status"] == BATCH_STATUS_COMPLETED:
        print("\nℹ️  任务已完成")
        print_summary(batch_task)
        input("\n按回车返回...")
        return

    pending_count = batch_task["total_files"] - batch_task["completed_count"] - batch_task["failed_count"] - batch_task["skipped_count"]
    print(f"\n当前进度: {batch_task['completed_count']}/{batch_task['total_files']} 已完成")
    print(f"待处理文件: {pending_count} 个")
    print(f"失败文件: {batch_task['failed_count']} 个")

    if prompt_confirm("\n是否继续处理？", default=True):
        config = batch_task.get("config", {})
        processor = BatchProcessor(batch_task=batch_task)
        success, result = processor.run(config=config)

        options = batch_task.get("options", {})
        if options.get("generate_report"):
            fmt = options.get("report_format", "txt")
            if fmt in ("txt", "both"):
                report_path = generate_report(batch_task)
                print(f"\n📄 文本报告已生成: {os.path.abspath(report_path)}")
            if fmt in ("json", "both"):
                report_path = generate_json_report(batch_task)
                print(f"📄 JSON报告已生成: {os.path.abspath(report_path)}")

        print_summary(batch_task)
        input("\n按回车返回...")


def _show_task_detail(batch_id):
    clear_screen()
    print_title(f"任务详情 - {batch_id}")

    batch_task = load_batch_task(batch_id)
    if not batch_task:
        print("\n❌ 任务不存在")
        input("\n按回车返回...")
        return

    print(f"\n状态: {_format_batch_status(batch_task['status'])}")
    print(f"总文件数: {batch_task['total_files']}")
    print(f"成功: {batch_task['completed_count']}")
    print(f"失败: {batch_task['failed_count']}")
    print(f"跳过: {batch_task['skipped_count']}")

    print("\n文件列表:")
    for i, f in enumerate(batch_task["files"], 1):
        icon = _get_file_status_icon(f["status"])
        filename = os.path.basename(f["json_path"])
        print(f"  {i:3d}. {icon} {filename}")
        if f["status"] == "completed":
            print(f"       数据: {f['data_count']} 条 | 耗时: {f.get('duration_seconds', 0)}秒")
        elif f["status"] == "failed":
            print(f"       错误: {f.get('error', '未知错误')}")

    input("\n按回车返回...")


def _get_batch_status_icon(status):
    icon_map = {
        BATCH_STATUS_PENDING: "⏳",
        BATCH_STATUS_RUNNING: "🔄",
        BATCH_STATUS_PAUSED: "⏸️",
        BATCH_STATUS_COMPLETED: "✅",
        BATCH_STATUS_FAILED: "❌",
    }
    return icon_map.get(status, "❓")


def _format_batch_status(status):
    status_map = {
        BATCH_STATUS_PENDING: "待处理",
        BATCH_STATUS_RUNNING: "处理中",
        BATCH_STATUS_PAUSED: "已暂停",
        BATCH_STATUS_COMPLETED: "已完成",
        BATCH_STATUS_FAILED: "部分失败",
    }
    return status_map.get(status, status)


def _get_file_status_icon(status):
    icon_map = {
        "pending": "⏳",
        "running": "🔄",
        "completed": "✅",
        "failed": "❌",
        "skipped": "⏭️",
    }
    return icon_map.get(status, "❓")


if __name__ == "__main__":
    run_batch_wizard()
