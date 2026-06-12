import os
import copy
import json
import sys
import signal
from datetime import datetime

from json_to_excel import load_json, auto_detect_headers, merge_headers, export_to_excel
from task_manager import (
    create_batch_task,
    load_batch_task,
    update_file_status,
    update_batch_status,
    get_pending_files,
    get_progress,
    resume_batch_task,
    TASK_STATUS_COMPLETED,
    TASK_STATUS_FAILED,
    TASK_STATUS_SKIPPED,
    TASK_STATUS_RUNNING,
    BATCH_STATUS_RUNNING,
    BATCH_STATUS_COMPLETED,
    BATCH_STATUS_FAILED,
    BATCH_STATUS_PAUSED,
)


class BatchProcessor:
    def __init__(self, batch_task=None, on_progress=None, on_file_complete=None, on_file_failed=None):
        self.batch_task = batch_task
        self.on_progress = on_progress
        self.on_file_complete = on_file_complete
        self.on_file_failed = on_file_failed
        self._stop_requested = False
        self._current_json_path = None

    def _signal_handler(self, signum, frame):
        print("\n\n收到中断信号，正在安全停止...")
        self._stop_requested = True

    def process_file(self, file_item, config, options=None):
        options = options or {}
        json_path = file_item["json_path"]
        excel_path = file_item["excel_path"]
        self._current_json_path = json_path

        skip_existing = options.get("skip_existing", False)
        if skip_existing and os.path.exists(excel_path):
            update_file_status(self.batch_task, json_path, TASK_STATUS_SKIPPED)
            if self.on_file_complete:
                self.on_file_complete(file_item, 0)
            return True, "已跳过（文件已存在）"

        update_file_status(self.batch_task, json_path, TASK_STATUS_RUNNING)

        try:
            if not os.path.exists(json_path):
                raise FileNotFoundError(f"JSON文件不存在: {json_path}")

            data = load_json(json_path)

            auto_headers = auto_detect_headers(data)
            headers = merge_headers(config.get("default_headers", []), auto_headers, config)

            file_config = copy.deepcopy(config)
            file_config["excel_output_path"] = excel_path

            output_dir = os.path.dirname(excel_path)
            if output_dir and not os.path.exists(output_dir):
                batch_output_dir = os.path.abspath(self.batch_task.get("output_dir", "./output"))
                try:
                    common = os.path.commonpath([batch_output_dir, os.path.abspath(output_dir)])
                    if common == batch_output_dir:
                        os.makedirs(output_dir, exist_ok=True)
                    else:
                        output_dir = batch_output_dir
                        safe_name = os.path.splitext(os.path.basename(excel_path))[0]
                        excel_path = os.path.join(batch_output_dir, safe_name + ".xlsx")
                        file_config["excel_output_path"] = excel_path
                        file_item["excel_path"] = excel_path
                        if not os.path.exists(output_dir):
                            os.makedirs(output_dir, exist_ok=True)
                except ValueError:
                    output_dir = batch_output_dir
                    if not os.path.exists(output_dir):
                        os.makedirs(output_dir, exist_ok=True)

            export_to_excel(data=data, headers=headers, config=file_config)

            update_file_status(self.batch_task, json_path, TASK_STATUS_COMPLETED, data_count=len(data))

            if self.on_file_complete:
                self.on_file_complete(file_item, len(data))

            return True, len(data)

        except Exception as e:
            error_msg = str(e)
            update_file_status(self.batch_task, json_path, TASK_STATUS_FAILED, error=error_msg)

            if self.on_file_failed:
                self.on_file_failed(file_item, error_msg)

            return False, error_msg

        finally:
            self._current_json_path = None

    def run(self, config=None, resume=False):
        if self.batch_task is None:
            raise ValueError("没有指定批量任务")

        if config is None:
            config = self.batch_task.get("config", {})

        options = self.batch_task.get("options", {})

        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)

        update_batch_status(self.batch_task, BATCH_STATUS_RUNNING)

        pending_files = get_pending_files(self.batch_task)

        if not pending_files:
            update_batch_status(self.batch_task, BATCH_STATUS_COMPLETED)
            return True, "没有待处理的文件"

        print(f"\n开始批量处理，共 {len(pending_files)} 个文件待处理")
        print("=" * 60)

        success_count = 0
        fail_count = 0

        for idx, file_item in enumerate(pending_files, 1):
            if self._stop_requested:
                update_batch_status(self.batch_task, BATCH_STATUS_PAUSED)
                print(f"\n批量处理已暂停，已处理 {idx - 1} 个文件")
                print(f"任务ID: {self.batch_task['batch_id']}，可通过 --batch-resume 恢复")
                return False, "用户中断"

            json_path = file_item["json_path"]
            rel_path = os.path.basename(json_path)
            print(f"\n[{idx}/{len(pending_files)}] 处理: {rel_path}")

            success, result = self.process_file(file_item, config, options)

            if success:
                success_count += 1
                if file_item["status"] == TASK_STATUS_SKIPPED:
                    print(f"  ⏭️  {result}")
                else:
                    print(f"  ✅ 成功 - {result} 条数据")
            else:
                fail_count += 1
                print(f"  ❌ 失败 - {result}")

            if self.on_progress:
                progress = get_progress(self.batch_task)
                self.on_progress(progress)

        all_completed = fail_count == 0
        status = BATCH_STATUS_COMPLETED if all_completed else BATCH_STATUS_FAILED
        update_batch_status(self.batch_task, status)

        print("\n" + "=" * 60)
        print(f"批量处理完成！")
        print(f"  总计: {self.batch_task['total_files']} 个文件")
        print(f"  成功: {self.batch_task['completed_count']} 个")
        print(f"  失败: {self.batch_task['failed_count']} 个")
        print(f"  跳过: {self.batch_task['skipped_count']} 个")

        return all_completed, {
            "success": success_count,
            "failed": fail_count,
            "total": self.batch_task["total_files"],
        }


def start_batch_process(source_type, sources, output_dir, config, options=None):
    batch_task = create_batch_task(source_type, sources, output_dir, config, options)

    if batch_task["total_files"] == 0:
        return None, "未找到任何JSON文件"

    processor = BatchProcessor(batch_task=batch_task)
    success, result = processor.run(config=config)

    return batch_task, result


def resume_batch_process(batch_id):
    batch_task, error = resume_batch_task(batch_id)
    if error:
        return None, error

    config = batch_task.get("config", {})
    processor = BatchProcessor(batch_task=batch_task)
    success, result = processor.run(config=config)

    return batch_task, result


def get_batch_info(batch_id):
    batch_task = load_batch_task(batch_id)
    if not batch_task:
        return None

    progress = get_progress(batch_task)
    return {
        "batch_task": batch_task,
        "progress": progress,
    }
