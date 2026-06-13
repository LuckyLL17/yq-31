import csv
import json
import os
import sys
import html
from datetime import datetime

from json_to_excel import extract_value, flatten_dict, ProgressTracker
from computed_columns import apply_computed_columns


EXPORT_FORMATS = {
    "excel": {
        "label": "Excel (.xlsx)",
        "extension": ".xlsx",
        "description": "Microsoft Excel 电子表格（功能最完整）",
    },
    "csv": {
        "label": "CSV (.csv)",
        "extension": ".csv",
        "description": "逗号分隔值，通用性最强",
    },
    "tsv": {
        "label": "TSV (.tsv)",
        "extension": ".tsv",
        "description": "制表符分隔值，适合复制到 Excel",
    },
    "html": {
        "label": "HTML (.html)",
        "extension": ".html",
        "description": "网页格式，可在浏览器中查看",
    },
    "markdown": {
        "label": "Markdown (.md)",
        "extension": ".md",
        "description": "Markdown 表格，适合文档",
    },
    "json": {
        "label": "JSON (.json)",
        "extension": ".json",
        "description": "格式化 JSON，便于程序处理",
    },
    "pdf": {
        "label": "PDF (.pdf)",
        "extension": ".pdf",
        "description": "PDF 文档，适合打印和分享",
    },
}


def get_format_extension(fmt):
    return EXPORT_FORMATS.get(fmt, {}).get("extension", ".xlsx")


def get_default_output_path(fmt, base_dir="./output"):
    ext = get_format_extension(fmt)
    return os.path.join(base_dir, f"result{ext}")


def _prepare_rows(data, headers, computed_cache=None, progress=None, progress_step=1):
    rows = []
    for item in data:
        if not isinstance(item, dict):
            if progress:
                progress.update(progress_step)
            continue
        row = []
        for h_idx, h in enumerate(headers):
            if computed_cache and h["key"] in computed_cache.get(id(item), {}):
                val = computed_cache[id(item)][h["key"]]
            else:
                val = extract_value(item, h["key"])
            if val is None:
                val = ""
            row.append(val)
            if progress and h_idx == 0:
                preview_val = str(val) if val is not None else ""
                progress.set_field(h["key"])
                progress.set_row_preview(preview_val)
        rows.append(row)
        if progress:
            progress.update(progress_step)
    return rows


def export_to_csv(data, headers, config, computed_cache=None):
    output_path = config.get("csv_output_path") or config.get("output_path", "./output/result.csv")
    csv_config = config.get("csv_config", {})
    encoding = csv_config.get("encoding", "utf-8-sig")
    delimiter = csv_config.get("delimiter", ",")
    include_header = csv_config.get("include_header", True)
    quote_char = csv_config.get("quote_char", '"')
    quoting = csv_config.get("quoting", "minimal")

    quoting_map = {
        "all": csv.QUOTE_ALL,
        "minimal": csv.QUOTE_MINIMAL,
        "nonnumeric": csv.QUOTE_NONNUMERIC,
        "none": csv.QUOTE_NONE,
    }

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出CSV", unit="行")
    rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
    progress.finish()

    with open(output_path, "w", encoding=encoding, newline="") as f:
        writer = csv.writer(
            f,
            delimiter=delimiter,
            quotechar=quote_char,
            quoting=quoting_map.get(quoting, csv.QUOTE_MINIMAL),
        )
        if include_header:
            writer.writerow([h["label"] for h in headers])
        writer.writerows(rows)

    print(f"CSV 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(rows)} 条数据，{len(headers)} 个字段")
    return output_path


def export_to_tsv(data, headers, config, computed_cache=None):
    output_path = config.get("tsv_output_path") or config.get("output_path", "./output/result.tsv")
    tsv_config = config.get("tsv_config", {})
    encoding = tsv_config.get("encoding", "utf-8-sig")
    include_header = tsv_config.get("include_header", True)

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出TSV", unit="行")
    rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
    progress.finish()

    with open(output_path, "w", encoding=encoding, newline="") as f:
        writer = csv.writer(f, delimiter="\t", quotechar='"', quoting=csv.QUOTE_MINIMAL)
        if include_header:
            writer.writerow([h["label"] for h in headers])
        writer.writerows(rows)

    print(f"TSV 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(rows)} 条数据，{len(headers)} 个字段")
    return output_path


def export_to_json(data, headers, config, computed_cache=None):
    output_path = config.get("json_output_path") or config.get("output_path", "./output/result.json")
    json_config = config.get("json_config", {})
    indent = json_config.get("indent", 2)
    ensure_ascii = json_config.get("ensure_ascii", False)
    include_labels = json_config.get("include_labels", False)

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出JSON", unit="行")
    export_data_list = []
    for item in data:
        if not isinstance(item, dict):
            progress.update()
            continue
        if include_labels:
            row_obj = {}
            for h_idx, h in enumerate(headers):
                if computed_cache and h["key"] in computed_cache.get(id(item), {}):
                    val = computed_cache[id(item)][h["key"]]
                else:
                    val = extract_value(item, h["key"])
                row_obj[h["label"]] = val
                if h_idx == 0:
                    preview_val = str(val) if val is not None else ""
                    progress.set_field(h["key"])
                    progress.set_row_preview(preview_val)
            export_data_list.append(row_obj)
        else:
            row_obj = {}
            for h_idx, h in enumerate(headers):
                if computed_cache and h["key"] in computed_cache.get(id(item), {}):
                    val = computed_cache[id(item)][h["key"]]
                else:
                    val = extract_value(item, h["key"])
                row_obj[h["key"]] = val
                if h_idx == 0:
                    preview_val = str(val) if val is not None else ""
                    progress.set_field(h["key"])
                    progress.set_row_preview(preview_val)
            export_data_list.append(row_obj)
        progress.update()
    progress.finish()

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(export_data_list, f, ensure_ascii=ensure_ascii, indent=indent if indent > 0 else None)

    print(f"JSON 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(export_data_list)} 条数据，{len(headers)} 个字段")
    return output_path


def export_to_markdown(data, headers, config, computed_cache=None):
    output_path = config.get("markdown_output_path") or config.get("output_path", "./output/result.md")
    md_config = config.get("markdown_config", {})
    title = md_config.get("title", "数据导出")
    include_index = md_config.get("include_index", False)
    max_col_width = md_config.get("max_col_width", 50)

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出Markdown", unit="行")
    rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
    progress.finish()

    def _truncate(s, width):
        s = str(s)
        if len(s) > width:
            return s[: width - 3] + "..."
        return s

    lines = []
    if title:
        lines.append(f"# {title}")
        lines.append("")
        lines.append(f"> 导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"> 数据条数: {len(rows)}")
        lines.append("")

    display_headers = []
    if include_index:
        display_headers.append("#")
    for h in headers:
        display_headers.append(_truncate(h["label"], max_col_width))

    lines.append("| " + " | ".join(display_headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(display_headers)) + " |")

    for idx, row in enumerate(rows, 1):
        display_row = []
        if include_index:
            display_row.append(str(idx))
        for val in row:
            cell = str(val).replace("\n", " ").replace("|", "\\|")
            display_row.append(_truncate(cell, max_col_width))
        lines.append("| " + " | ".join(display_row) + " |")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"Markdown 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(rows)} 条数据，{len(headers)} 个字段")
    return output_path


def export_to_html(data, headers, config, computed_cache=None):
    output_path = config.get("html_output_path") or config.get("output_path", "./output/result.html")
    html_config = config.get("html_config", {})
    title = html_config.get("title", "数据导出")
    include_index = html_config.get("include_index", False)
    pretty_print = html_config.get("pretty_print", True)
    style = html_config.get("style", "default")
    custom_css = html_config.get("custom_css", "")

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出HTML", unit="行")
    rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
    progress.finish()

    default_css = """
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif; margin: 20px; }
        h1 { color: #333; border-bottom: 2px solid #4472C4; padding-bottom: 10px; }
        .meta { color: #666; margin-bottom: 20px; font-size: 14px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        th { background-color: #4472C4; color: white; padding: 12px 15px; text-align: left; font-weight: bold; }
        td { padding: 10px 15px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f1f1f1; }
        .summary { margin-top: 20px; color: #666; font-size: 14px; }
    </style>
    """

    compact_css = """
    <style>
        body { font-family: monospace; margin: 10px; }
        table { border-collapse: collapse; font-size: 12px; }
        th, td { border: 1px solid #999; padding: 4px 8px; }
        th { background: #ccc; }
    </style>
    """

    no_style_css = "<style></style>"

    style_map = {
        "default": default_css,
        "compact": compact_css,
        "none": no_style_css,
    }

    css = style_map.get(style, default_css)
    if custom_css:
        css = f"<style>{custom_css}</style>"

    header_cells = []
    if include_index:
        header_cells.append("<th>#</th>")
    for h in headers:
        header_cells.append(f"<th>{html.escape(str(h['label']))}</th>")

    body_rows = []
    for idx, row in enumerate(rows, 1):
        cells = []
        if include_index:
            cells.append(f"<td>{idx}</td>")
        for val in row:
            cells.append(f"<td>{html.escape(str(val))}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    indent = "  " if pretty_print else ""
    newline = "\n" if pretty_print else ""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html.escape(title)}</title>
{css}
</head>
<body>
<h1>{html.escape(title)}</h1>
<div class="meta">
  <div>导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
  <div>数据条数: {len(rows)}</div>
  <div>字段数量: {len(headers)}</div>
</div>
<table>
  <thead>
    <tr>{newline.join(header_cells)}</tr>
  </thead>
  <tbody>
    {newline.join(body_rows)}
  </tbody>
</table>
<div class="summary">共 {len(rows)} 条记录</div>
</body>
</html>
"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"HTML 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(rows)} 条数据，{len(headers)} 个字段")
    return output_path


def export_to_pdf(data, headers, config, computed_cache=None):
    output_path = config.get("pdf_output_path") or config.get("output_path", "./output/result.pdf")
    pdf_config = config.get("pdf_config", {})
    title = pdf_config.get("title", "数据导出")
    include_index = pdf_config.get("include_index", False)
    page_size = pdf_config.get("page_size", "A4")
    orientation = pdf_config.get("orientation", "portrait")
    font_size = pdf_config.get("font_size", 10)

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    progress = ProgressTracker(total=len(data), description="📝 导出PDF", unit="行")
    rows = _prepare_rows(data, headers, computed_cache=computed_cache, progress=progress)
    progress.finish()

    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        print("⚠️  PDF 导出需要安装 reportlab 库")
        print("请运行: pip install reportlab")
        print("正在尝试使用 HTML 中转方式...")
        return _export_pdf_via_html(data, headers, config, output_path, title, include_index, computed_cache=computed_cache)

    try:
        font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]
        font_registered = False
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    pdfmetrics.registerFont(TTFont("ChineseFont", fp))
                    font_registered = True
                    break
                except Exception:
                    continue
    except Exception:
        font_registered = False

    page_size_map = {"A4": A4, "letter": letter}
    ps = page_size_map.get(page_size, A4)
    if orientation == "landscape":
        ps = (ps[1], ps[0])

    doc = SimpleDocTemplate(
        output_path,
        pagesize=ps,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )

    styles = getSampleStyleSheet()
    elements = []

    if font_registered:
        title_font = "ChineseFont"
        body_font = "ChineseFont"
        title_style = ParagraphStyle(
            "CustomTitle",
            parent=styles["Title"],
            fontName=title_font,
            fontSize=18,
            spaceAfter=12,
        )
        normal_style = ParagraphStyle(
            "CustomNormal",
            parent=styles["Normal"],
            fontName=body_font,
            fontSize=font_size,
        )
    else:
        title_style = styles["Title"]
        normal_style = styles["Normal"]

    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 6 * mm))
    meta_text = (
        f"导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  |  "
        f"数据: {len(rows)} 条  |  字段: {len(headers)} 个"
    )
    elements.append(Paragraph(meta_text, normal_style))
    elements.append(Spacer(1, 8 * mm))

    table_data = []
    header_row = []
    if include_index:
        header_row.append("#")
    for h in headers:
        header_row.append(str(h["label"]))
    table_data.append(header_row)

    for idx, row in enumerate(rows, 1):
        data_row = []
        if include_index:
            data_row.append(str(idx))
        for val in row:
            cell = str(val)
            if len(cell) > 100:
                cell = cell[:97] + "..."
            data_row.append(cell)
        table_data.append(data_row)

    col_count = len(headers) + (1 if include_index else 0)
    available_width = ps[0] - 40 * mm
    col_widths = [available_width / col_count] * col_count

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, 0), font_size + 1),
        ("FONTSIZE", (0, 1), (-1, -1), font_size),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("TOPPADDING", (0, 1), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F9F9F9")]),
    ]
    if font_registered:
        table_style.append(("FONTNAME", (0, 0), (-1, -1), body_font))
    table.setStyle(TableStyle(table_style))
    elements.append(table)

    doc.build(elements)
    print(f"PDF 文件已成功导出: {os.path.abspath(output_path)}")
    print(f"共导出 {len(rows)} 条数据，{len(headers)} 个字段")
    return output_path


def _export_pdf_via_html(data, headers, config, output_path, title, include_index, computed_cache=None):
    try:
        import weasyprint
    except ImportError:
        print("❌ 无法导出 PDF: 需要安装 reportlab 或 weasyprint")
        print("请运行: pip install reportlab  或  pip install weasyprint")
        return None

    html_path = output_path + ".tmp.html"
    temp_config = dict(config)
    temp_config["html_output_path"] = html_path
    temp_config["html_config"] = config.get("pdf_config", {})
    export_to_html(data, headers, temp_config, computed_cache=computed_cache)

    weasyprint.HTML(filename=html_path).write_pdf(output_path)

    try:
        os.remove(html_path)
    except Exception:
        pass

    print(f"PDF 文件已成功导出 (via HTML): {os.path.abspath(output_path)}")
    return output_path


def export_data(data, headers, config, fmt=None):
    if fmt is None:
        fmt = config.get("export_format", "excel")

    computed_cache = None
    computed_columns = config.get("computed_columns", [])
    if computed_columns:
        enabled_cc = [cc for cc in computed_columns if cc.get("enabled", True)]
        if enabled_cc:
            print(f"正在计算 {len(enabled_cc)} 个计算列...")
            computed_cache, headers = apply_computed_columns(data, headers, config, extract_value)
            print(f"  已添加 {len(enabled_cc)} 个计算列: {', '.join(cc['label'] for cc in enabled_cc)}\n")

    if fmt == "excel":
        from json_to_excel import export_to_excel
        return export_to_excel(data, headers, config)
    elif fmt == "csv":
        return export_to_csv(data, headers, config, computed_cache=computed_cache)
    elif fmt == "tsv":
        return export_to_tsv(data, headers, config, computed_cache=computed_cache)
    elif fmt == "html":
        return export_to_html(data, headers, config, computed_cache=computed_cache)
    elif fmt == "markdown":
        return export_to_markdown(data, headers, config, computed_cache=computed_cache)
    elif fmt == "json":
        return export_to_json(data, headers, config, computed_cache=computed_cache)
    elif fmt == "pdf":
        return export_to_pdf(data, headers, config, computed_cache=computed_cache)
    else:
        raise ValueError(f"不支持的导出格式: {fmt}")
