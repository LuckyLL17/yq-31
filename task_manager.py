import json
import os
import uuid
import time
from datetime import datetime


TASK_STATUS_PENDING = "pending"
TASK_STATUS_RUNNING = "running"
TASK_STATUS_COMPLETED = "completed"
TASK_STATUS_FAILED = "failed"
TASK_STATUS_SKIPPED = "skipped"

BATCH_STATUS_PENDING = "pending"
BATCH_STATUS_RUNNING = "running"
BATCH_STATUS_PAUSED = "paused"
BATCH_STATUS_COMPLETED = "completed"
BATCH_STATUS_FAILED = "failed"

TASKS_DIR = "./.batch_tasks"


def _ensure_tasks_dir():
    if not os.path.exists(TASKS_DIR):
        os.makedirs(TASKS_DIR, exist_ok=True)


def _get_task_file_path(batch_id):
    return os.path.join(TASKS_DIR, f"batch_{batch_id}.json")


def create_batch_task(source_type, sources, output_dir, config, options=None):
    _ensure_tasks_dir()

    batch_id = str(uuid.uuid4())[:8]
    now = datetime.now().isoformat()

    file_items = []
    if source_type == "directory":
        for src in sources:
            if os.path.isdir(src):
                for root, _, files in os.walk(src):
                    for f in files:
                        if f.lower().endswith(".json"):
                            full_path = os.path.join(root, f)
                            file_items.append(_create_file_item(full_path, output_dir, src))
            elif os.path.isfile(src) and src.lower().endswith(".json"):
                file_items.append(_create_file_item(src, output_dir))
    else:
        for src in sources:
            if os.path.isfile(src) and src.lower().endswith(".json"):
                file_items.append(_create_file_item(src, output_dir))

    batch_task = {
        "batch_id": batch_id,
        "source_type": source_type,
        "sources": sources,
        "output_dir": output_dir,
        "total_files": len(file_items),
        "completed_count": 0,
        "failed_count": 0,
        "skipped_count": 0,
        "status": BATCH_STATUS_PENDING,
        "created_at": now,
        "started_at": None,
        "completed_at": None,
        "config": config,
        "options": options or {},
        "files": file_items,
    }

    _save_batch_task(batch_task)
    return batch_task


def _create_file_item(json_path, output_dir, base_dir=None):
    if base_dir:
        rel_path = os.path.relpath(json_path, base_dir)
        base_name = os.path.splitext(rel_path)[0]
        excel_path = os.path.join(output_dir, base_name + ".xlsx")
    else:
        base_name = os.path.splitext(os.path.basename(json_path))[0]
        excel_path = os.path.join(output_dir, base_name + ".xlsx")

    return {
        "json_path": os.path.abspath(json_path),
        "excel_path": os.path.abspath(excel_path),
        "status": TASK_STATUS_PENDING,
        "data_count": 0,
        "error": None,
        "started_at": None,
        "completed_at": None,
        "duration_seconds": 0,
    }


def _save_batch_task(batch_task):
    file_path = _get_task_file_path(batch_task["batch_id"])
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(batch_task, f, ensure_ascii=False, indent=2)


def load_batch_task(batch_id):
    file_path = _get_task_file_path(batch_id)
    if not os.path.exists(file_path):
        return None
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def update_file_status(batch_task, json_path, status, data_count=0, error=None):
    for file_item in batch_task["files"]:
        if file_item["json_path"] == json_path:
            file_item["status"] = status
            file_item["data_count"] = data_count
            file_item["error"] = error

            if status == TASK_STATUS_RUNNING:
                file_item["started_at"] = datetime.now().isoformat()
            elif status in (TASK_STATUS_COMPLETED, TASK_STATUS_FAILED, TASK_STATUS_SKIPPED):
                file_item["completed_at"] = datetime.now().isoformat()
                if file_item["started_at"]:
                    start = datetime.fromisoformat(file_item["started_at"])
                    end = datetime.fromisoformat(file_item["completed_at"])
                    file_item["duration_seconds"] = round((end - start).total_seconds(), 2)

            _update_batch_counts(batch_task)
            _save_batch_task(batch_task)
            break


def _update_batch_counts(batch_task):
    completed = 0
    failed = 0
    skipped = 0

    for f in batch_task["files"]:
        if f["status"] == TASK_STATUS_COMPLETED:
            completed += 1
        elif f["status"] == TASK_STATUS_FAILED:
            failed += 1
        elif f["status"] == TASK_STATUS_SKIPPED:
            skipped += 1

    batch_task["completed_count"] = completed
    batch_task["failed_count"] = failed
    batch_task["skipped_count"] = skipped


def update_batch_status(batch_task, status):
    batch_task["status"] = status
    now = datetime.now().isoformat()

    if status == BATCH_STATUS_RUNNING and not batch_task["started_at"]:
        batch_task["started_at"] = now
    elif status in (BATCH_STATUS_COMPLETED, BATCH_STATUS_FAILED):
        batch_task["completed_at"] = now

    _save_batch_task(batch_task)


def get_pending_files(batch_task):
    return [f for f in batch_task["files"] if f["status"] in (TASK_STATUS_PENDING, TASK_STATUS_FAILED)]


def get_progress(batch_task):
    total = batch_task["total_files"]
    completed = batch_task["completed_count"]
    failed = batch_task["failed_count"]
    skipped = batch_task["skipped_count"]
    processed = completed + failed + skipped
    percent = (processed / total * 100) if total > 0 else 0

    return {
        "total": total,
        "completed": completed,
        "failed": failed,
        "skipped": skipped,
        "processed": processed,
        "percent": round(percent, 1),
    }


def list_batch_tasks():
    _ensure_tasks_dir()
    tasks = []
    for filename in os.listdir(TASKS_DIR):
        if filename.startswith("batch_") and filename.endswith(".json"):
            batch_id = filename[6:-5]
            task = load_batch_task(batch_id)
            if task:
                tasks.append({
                    "batch_id": task["batch_id"],
                    "status": task["status"],
                    "total_files": task["total_files"],
                    "completed_count": task["completed_count"],
                    "failed_count": task["failed_count"],
                    "created_at": task["created_at"],
                    "output_dir": task["output_dir"],
                })
    tasks.sort(key=lambda x: x["created_at"], reverse=True)
    return tasks


def delete_batch_task(batch_id):
    file_path = _get_task_file_path(batch_id)
    if os.path.exists(file_path):
        os.remove(file_path)
        return True
    return False


def resume_batch_task(batch_id):
    batch_task = load_batch_task(batch_id)
    if not batch_task:
        return None, "任务不存在"

    if batch_task["status"] == BATCH_STATUS_COMPLETED:
        return batch_task, "任务已完成，无需恢复"

    return batch_task, None
