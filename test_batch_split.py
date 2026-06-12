import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config_manager import get_default_config
from batch_processor import start_batch_process


def main():
    config = get_default_config()
    config["export_format"] = "excel"

    config["split_config"] = {
        "enabled": True,
        "split_field": "department",
        "split_rule": "by_value",
        "sheet_name_template": "{value}",
        "include_all_sheet": True,
        "all_sheet_name": "全部数据",
        "empty_value_label": "未分类",
        "max_sheet_name_length": 31,
    }

    output_dir = tempfile.mkdtemp(prefix="batch_split_test_")
    print(f"测试输出目录: {output_dir}")

    print(f"\n配置中的 split_config: {config['split_config']}")
    print(f"split_config.enabled: {config['split_config']['enabled']}")

    batch_task, result = start_batch_process(
        source_type="directory",
        sources=["./data/batch_test"],
        output_dir=output_dir,
        config=config,
        options={
            "generate_report": False,
            "skip_existing": False,
        },
    )

    if batch_task is None:
        print(f"\n❌ 创建批量任务失败: {result}")
        return 1

    print(f"\n批量任务 ID: {batch_task['batch_id']}")
    print(f"任务中保存的 split_config.enabled: {batch_task['config'].get('split_config', {}).get('enabled')}")
    print(f"任务中保存的 split_field: {batch_task['config'].get('split_config', {}).get('split_field')}")

    print(f"\n生成的文件:")
    for f in batch_task["files"]:
        print(f"  - {os.path.basename(f['output_path'])} (状态: {f['status']})")
        if f["status"] == "completed":
            print(f"    数据条数: {f.get('data_count', 0)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
