import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from json_to_excel import load_json, auto_detect_headers, merge_headers
from multi_exporter import export_data, EXPORT_FORMATS
from config_manager import get_default_config


def test_all_formats():
    json_path = "./data/sample_data.json"
    output_dir = "./output/test_formats"
    os.makedirs(output_dir, exist_ok=True)

    print("加载数据...")
    data = load_json(json_path)
    print(f"已加载 {len(data)} 条数据")

    auto_headers = auto_detect_headers(data)
    print(f"自动检测到 {len(auto_headers)} 个字段")

    base_config = get_default_config()
    headers = merge_headers(base_config.get("default_headers", []), auto_headers, base_config)
    print(f"使用 {len(headers)} 个字段\n")

    test_formats = ["csv", "tsv", "html", "markdown", "json", "excel"]

    for fmt in test_formats:
        print("=" * 60)
        print(f"测试导出格式: {EXPORT_FORMATS[fmt]['label']}")
        print("=" * 60)

        config = get_default_config()
        config["export_format"] = fmt
        ext = EXPORT_FORMATS[fmt]["extension"]
        output_key = f"{fmt}_output_path" if fmt != "excel" else "excel_output_path"
        config[output_key] = os.path.join(output_dir, f"test_result{ext}")

        try:
            result = export_data(data, headers, config, fmt=fmt)
            if result and os.path.exists(config[output_key]):
                size = os.path.getsize(config[output_key])
                print(f"✅ 成功！文件大小: {size} 字节\n")
            else:
                print(f"⚠️  导出完成但未确认文件存在\n")
        except Exception as e:
            print(f"❌ 失败: {e}\n")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 60)
    print("测试 PDF 格式（需要 reportlab 库）...")
    print("=" * 60)
    try:
        config = get_default_config()
        config["export_format"] = "pdf"
        config["pdf_output_path"] = os.path.join(output_dir, "test_result.pdf")
        result = export_data(data, headers, config, fmt="pdf")
        if result and os.path.exists(config["pdf_output_path"]):
            size = os.path.getsize(config["pdf_output_path"])
            print(f"✅ PDF 导出成功！文件大小: {size} 字节")
        else:
            print(f"⚠️  PDF 导出完成但未确认文件存在")
    except ImportError as e:
        print(f"ℹ️  reportlab 未安装，PDF 测试跳过: {e}")
    except Exception as e:
        print(f"❌ PDF 导出失败: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "=" * 60)
    print(f"测试完成！输出目录: {os.path.abspath(output_dir)}")
    print("=" * 60)


if __name__ == "__main__":
    test_all_formats()
