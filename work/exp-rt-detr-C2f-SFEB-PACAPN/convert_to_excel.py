import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "work")


def parse_txt(filepath, model_name=None):
    """解析单个模型txt文件，返回模型信息字典"""
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    if model_name is None:
        model_name = os.path.splitext(os.path.basename(filepath))[0]

    result = {"model": model_name}

    # 解析 Model Info 表
    info_match = re.search(
        r"\|\s*([\d.]+)\s*\|\s*([\d,]+)\s*\|\s*([\d.]+)s\s*\|\s*([\d.]+)s\s*\|\s*([\d.]+)s\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+MB)\s*\|",
        content
    )
    if info_match:
        result["gflops"] = float(info_match.group(1))
        result["parameters"] = info_match.group(2).strip()
        result["preprocess_time"] = float(info_match.group(3))
        result["inference_time"] = float(info_match.group(4))
        result["postprocess_time"] = float(info_match.group(5))
        result["fps_total"] = float(info_match.group(6))
        result["fps_inference"] = float(info_match.group(7))
        result["model_size"] = info_match.group(8).strip()

    # 解析 Model Metrics 表
    lines = content.split("\n")
    classes_data = {}
    for line in lines:
        metric_match = re.match(
            r"\|\s*(\S+(?:\(\S+\))?)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|",
            line
        )
        if metric_match:
            cls_name = metric_match.group(1).strip()
            precision = float(metric_match.group(2))
            recall = float(metric_match.group(3))
            f1 = float(metric_match.group(4))
            map50 = float(metric_match.group(5))
            map75 = float(metric_match.group(6))
            map50_95 = float(metric_match.group(7))
            classes_data[cls_name] = {
                "precision": precision,
                "recall": recall,
                "f1": f1,
                "map50": map50,
                "map75": map75,
                "map50_95": map50_95,
            }

    result["classes"] = classes_data
    return result


def load_all_models(data_dir=None):
    """加载所有模型数据。
    支持两种模式：
    1. 新模式：递归扫描 data_dir 下所有 paper_data.txt，自动定位
    2. 旧模式：直接扫描 data_dir 下的 .txt 文件（兜底）
    """
    if data_dir is None:
        data_dir = INPUT_DIR

    models = []

    # 新模式：递归扫描 paper_data.txt（不限嵌套深度）
    for root, dirs, files in os.walk(data_dir):
        # 跳过隐藏目录和 __pycache__
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__pycache__']
        for fname in files:
            if fname == "paper_data.txt":
                paper_path = os.path.join(root, fname)
                # 模型名 = paper_data.txt 相对于 data_dir 的父目录路径（/ 替换为 _）
                if os.path.normpath(root) == os.path.normpath(data_dir):
                    model_name = os.path.basename(data_dir.rstrip(os.sep))
                else:
                    rel_path = os.path.relpath(root, data_dir)
                    model_name = rel_path.replace(os.sep, '_')
                data = parse_txt(paper_path, model_name=model_name)
                models.append(data)

    # 旧模式：如果没有找到任何 paper_data.txt，则扫描 .txt 文件
    if not models:
        for fname in sorted(os.listdir(data_dir)):
            if fname.endswith(".txt"):
                filepath = os.path.join(data_dir, fname)
                data = parse_txt(filepath)
                models.append(data)

    return models


# 样式定义
HEADER_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SUB_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
SUBTITLE_FONT = Font(name="微软雅黑", size=12, bold=True, color="2F5496")
DATA_FONT = Font(name="微软雅黑", size=11)
BEST_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # 浅黄 - 最优
WORST_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # 浅红 - 最差
MODEL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ALL_ROW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # 浅黄 - 平均行
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")


def apply_cell_style(cell, font=DATA_FONT, fill=None, alignment=CENTER_ALIGN, border=THIN_BORDER, number_format=None):
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = alignment
    cell.border = border
    if number_format:
        cell.number_format = number_format


def set_header_row(ws, row, col, headers, fill=HEADER_FILL):
    for i, h in enumerate(headers):
        cell = ws.cell(row=row, column=col + i, value=h)
        apply_cell_style(cell, font=HEADER_FONT, fill=fill)


def auto_width(ws, min_width=10, max_width=22):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.value:
                # 中文字符算2个宽度
                val = str(cell.value)
                length = 0
                for ch in val:
                    length += 2 if ord(ch) > 127 else 1
                max_len = max(max_len, length)
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 4, max_width))


def highlight_best_worst(ws, row_start, row_end, col, higher_is_better=True):
    """在指定列范围内高亮最优值和最差值"""
    values = []
    for r in range(row_start, row_end + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, (int, float)):
            values.append((r, v))
    if not values:
        return
    values.sort(key=lambda x: x[1], reverse=higher_is_better)
    best_val = values[0][1]
    worst_val = values[-1][1]
    for r, v in values:
        if v == best_val:
            ws.cell(row=r, column=col).fill = BEST_FILL
        elif v == worst_val:
            ws.cell(row=r, column=col).fill = WORST_FILL


def build_sheet1_overview(wb, models):
    """Sheet 1: 综合性能对比"""
    ws = wb.active
    ws.title = "综合性能对比"

    # 标题
    ws.merge_cells("A1:O1")
    title_cell = ws.cell(row=1, column=1, value="模型综合性能对比")
    apply_cell_style(title_cell, font=TITLE_FONT, alignment=Alignment(horizontal="center", vertical="center"))
    title_cell.border = Border()

    # 表头
    headers = [
        "模型名称",
        "GFLOPs",
        "参数量",
        "FPS(总)",
        "FPS(推理)",
        "模型大小",
        "Precision(avg)",
        "Recall(avg)",
        "F1-Score(avg)",
        "mAP50(avg)",
        "mAP75(avg)",
        "mAP50-95(avg)",
        "mAP50(Platelets)",
        "mAP50(RBC)",
        "mAP50(WBC)",
    ]
    header_row = 3
    set_header_row(ws, header_row, 1, headers)

    data_start = header_row + 1
    for idx, m in enumerate(models):
        row = data_start + idx
        all_data = m["classes"].get("all(平均数据)", {})

        model_name = m["model"]
        if "yolov8" in model_name.lower():
            model_name = "YOLOv8-DETR"
        elif "yolov5" in model_name.lower():
            model_name = "YOLOv5-DETR"

        data = [
            model_name,
            m.get("gflops"),
            m.get("parameters"),
            m.get("fps_total"),
            m.get("fps_inference"),
            m.get("model_size"),
            all_data.get("precision"),
            all_data.get("recall"),
            all_data.get("f1"),
            all_data.get("map50"),
            all_data.get("map75"),
            all_data.get("map50_95"),
            m["classes"].get("Platelets", {}).get("map50"),
            m["classes"].get("RBC", {}).get("map50"),
            m["classes"].get("WBC", {}).get("map50"),
        ]
        for j, val in enumerate(data):
            cell = ws.cell(row=row, column=j + 1, value=val)
            fill = MODEL_FILL if j == 0 else None
            apply_cell_style(cell, fill=fill)
            if j == 0:
                cell.alignment = LEFT_ALIGN

    data_end = data_start + len(models) - 1

    # 高亮最优/最差
    # GFLOPs/参数量/模型大小(越小越好), FPS/精度(越大越好)
    higher_better = {
        2: False,   # GFLOPs
        3: False,   # 参数量 (string, skip)
        4: True,    # FPS(总)
        5: True,    # FPS(推理)
        6: False,   # 模型大小 (string, skip)
        7: True,    # Precision
        8: True,    # Recall
        9: True,    # F1
        10: True,   # mAP50
        11: True,   # mAP75
        12: True,   # mAP50-95
        13: True,   # mAP50(Platelets)
        14: True,   # mAP50(RBC)
        15: True,   # mAP50(WBC)
    }
    for col, higher in higher_better.items():
        highlight_best_worst(ws, data_start, data_end, col, higher_is_better=higher)

    auto_width(ws, min_width=12, max_width=22)
    ws.freeze_panes = "B4"


def build_sheet2_per_class(wb, models):
    """Sheet 2: 各类别指标对比"""
    ws = wb.create_sheet("各类别指标对比")

    metrics = [
        ("Precision", "precision"),
        ("Recall", "recall"),
        ("F1-Score", "f1"),
        ("mAP50", "map50"),
        ("mAP75", "map75"),
        ("mAP50-95", "map50_95"),
    ]
    classes = ["Platelets", "RBC", "WBC", "all(平均数据)"]

    current_row = 1
    for metric_name, metric_key in metrics:
        # 子标题
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=1 + len(models))
        subtitle_cell = ws.cell(row=current_row, column=1, value=f"【{metric_name}】")
        apply_cell_style(subtitle_cell, font=SUBTITLE_FONT, alignment=LEFT_ALIGN)
        subtitle_cell.border = Border()
        current_row += 1

        # 表头: 类别 | 模型1 | 模型2 | ...
        sub_headers = ["类别"] + [m["model"] for m in models]
        set_header_row(ws, current_row, 1, sub_headers, fill=SUB_HEADER_FILL)
        current_row += 1

        data_start_row = current_row
        for cls_name in classes:
            row_data = [cls_name]
            for m in models:
                cls_data = m["classes"].get(cls_name, {})
                row_data.append(cls_data.get(metric_key))
            for j, val in enumerate(row_data):
                cell = ws.cell(row=current_row, column=j + 1, value=val)
                fill = None
                if j == 0:
                    if cls_name == "all(平均数据)":
                        fill = ALL_ROW_FILL
                    else:
                        fill = MODEL_FILL
                apply_cell_style(cell, fill=fill)
                if j == 0:
                    cell.alignment = LEFT_ALIGN
            current_row += 1

        data_end_row = current_row - 1

        # 高亮最优/最差 (非平均行)
        for col_idx in range(2, 2 + len(models)):
            # 只比较前三行(Platelets, RBC, WBC)，不包含平均行
            values = []
            for r in range(data_start_row, data_start_row + 3):
                v = ws.cell(row=r, column=col_idx).value
                if isinstance(v, (int, float)):
                    values.append((r, v))
            if not values:
                continue
            values.sort(key=lambda x: x[1], reverse=True)
            best_val = values[0][1]
            worst_val = values[-1][1]
            for r, v in values:
                if v == best_val:
                    ws.cell(row=r, column=col_idx).fill = BEST_FILL
                elif v == worst_val:
                    ws.cell(row=r, column=col_idx).fill = WORST_FILL

        # 空行分隔
        current_row += 1

    auto_width(ws, min_width=12, max_width=22)
    ws.freeze_panes = "B4"


def build_sheet3_model_vs_class(wb, models):
    """Sheet 3: 模型×类别矩阵 - 以模型为行、类别为列，按指标分块"""
    ws = wb.create_sheet("模型×类别矩阵")

    all_classes = ["Platelets", "RBC", "WBC"]
    metric_groups = [
        ("Precision", "precision"),
        ("Recall", "recall"),
        ("F1-Score", "f1"),
        ("mAP50", "map50"),
        ("mAP75", "map75"),
        ("mAP50-95", "map50_95"),
    ]

    current_row = 1
    for metric_name, metric_key in metric_groups:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=1 + len(all_classes))
        subtitle_cell = ws.cell(row=current_row, column=1, value=f"【{metric_name}】")
        apply_cell_style(subtitle_cell, font=SUBTITLE_FONT, alignment=LEFT_ALIGN)
        subtitle_cell.border = Border()
        current_row += 1

        # 表头
        sub_headers = ["模型"] + all_classes
        set_header_row(ws, current_row, 1, sub_headers, fill=SUB_HEADER_FILL)
        current_row += 1

        data_start_row = current_row
        for m in models:
            model_name = m["model"]
            if "yolov8" in model_name.lower():
                model_name = "YOLOv8-DETR"
            elif "yolov5" in model_name.lower():
                model_name = "YOLOv5-DETR"

            row_data = [model_name]
            for cls_name in all_classes:
                cls_data = m["classes"].get(cls_name, {})
                row_data.append(cls_data.get(metric_key))
            for j, val in enumerate(row_data):
                cell = ws.cell(row=current_row, column=j + 1, value=val)
                fill = MODEL_FILL if j == 0 else None
                apply_cell_style(cell, fill=fill)
                if j == 0:
                    cell.alignment = LEFT_ALIGN
            current_row += 1

        data_end_row = current_row - 1

        # 高亮最优/最差
        for col_idx in range(2, 2 + len(all_classes)):
            highlight_best_worst(ws, data_start_row, data_end_row, col_idx, higher_is_better=True)

        current_row += 1  # 空行分隔

    auto_width(ws, min_width=12, max_width=22)
    ws.freeze_panes = "A2"


def main():
    # 支持命令行参数指定目标文件夹（自动拼接含空格的路径）
    if len(sys.argv) > 1:
        data_dir = os.path.abspath(' '.join(sys.argv[1:]))
    else:
        data_dir = INPUT_DIR

    if not os.path.isdir(data_dir):
        print(f"错误：目录不存在 - {data_dir}")
        return

    output_file = os.path.join(data_dir, "模型性能对比.xlsx")

    models = load_all_models(data_dir)
    if not models:
        print("未找到任何模型数据！")
        print("支持的模式：")
        print("  1. 子目录模式：扫描子目录中的 paper_data.txt")
        print("  2. 旧模式：扫描目录中的 .txt 文件")
        return

    # 按 mAP50(avg) 降序排列
    models.sort(key=lambda m: m["classes"].get("all(平均数据)", {}).get("map50", 0), reverse=True)

    print(f"数据目录: {data_dir}")
    print(f"找到 {len(models)} 个模型 (按mAP50降序排列):")
    for m in models:
        all_data = m["classes"].get("all(平均数据)", {})
        print(f"  - {m['model']}  (mAP50: {all_data.get('map50', 'N/A')})")

    wb = Workbook()

    build_sheet1_overview(wb, models)
    build_sheet2_per_class(wb, models)
    build_sheet3_model_vs_class(wb, models)

    wb.save(output_file)
    print(f"\nExcel 已保存至: {output_file}")


if __name__ == "__main__":
    main()