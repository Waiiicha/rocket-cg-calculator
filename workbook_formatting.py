from __future__ import annotations

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation


HEADER_FILL = "1F4E78"
INPUT_FILL = "FFF2CC"
NOTE_FILL = "F2F2F2"
INFO_FILL = "D9EAF7"
OUTPUT_FILL = "E2F0D9"
WHITE = "FFFFFF"
BLUE = "0000FF"
BLACK = "000000"
GRID = "BFBFBF"
INPUT_RESERVE_ROWS = 500


SHEET_NOTES = {
    "总体参数": "填写全局计算参数。时间单位为 s，长度单位为 m，质量单位为 kg。",
    "级事件": "填写各子级或助推器的点火、关机和分离时间。分离规则为 t >= ts 时剔除该级/助推器。",
    "不变质量": "填写结构、设备、有效载荷等不含推进剂和增压气体的部组件质量特性。",
    "可抛质量": "填写整流罩、热级环等按时间抛离的部件。无可抛部件时可留空。",
    "贮箱基本参数": "填写每个贮箱的基本参数。等效圆柱只需填写等效半径R和等效高度H；QJ第一类需补充几何表。",
    "贮箱几何_第一类": "按 QJ 1080A 图3填写第一类外轮廓半径模型参数。等效圆柱贮箱不需要填写本表。",
    "贮箱几何_第二类": "按 QJ 1080A 图4填写内部扣除体积半径模型参数。没有扣除体积时选择 否。",
    "外行导管": "填写外行导管内推进剂对 Jx 的贡献。没有外行导管时可留空或填0。",
    "运输状态": "运输状态输入预留表，当前计算流程暂未接入输出。",
}


HEADER_COMMENTS = {
    "类型": "选择 子级 或 助推器。",
    "编号": "子级或助推器编号，应与其他表中的所属编号一致。",
    "所属编号": "对应子级或助推器编号。",
    "贮箱编号": "同一子级/助推器内的贮箱序号。贮箱ID规则：类型-所属编号-贮箱编号。",
    "几何类型": "选择 等效圆柱 或 QJ第一类。等效圆柱不需要填写QJ几何表。",
    "等效半径R": "仅几何类型为 等效圆柱 时使用，单位 m。",
    "等效高度H": "仅几何类型为 等效圆柱 时使用，单位 m。",
    "M": "质量，单位 kg。",
    "X": "质心 X 坐标，OXYZ 坐标系，单位 m。",
    "Y": "质心 Y 坐标，OXYZ 坐标系，单位 m。",
    "Z": "质心 Z 坐标，OXYZ 坐标系，单位 m。",
    "Jx": "绕部组件自身质心主轴 X 的中心主惯量，单位 kg·m^2。",
    "Jy": "绕部组件自身质心主轴 Y 的中心主惯量，单位 kg·m^2。",
    "Jz": "绕部组件自身质心主轴 Z 的中心主惯量，单位 kg·m^2。",
    "点火时间tf": "发动机开始工作时间，单位 s。",
    "关机时间tb": "发动机关机时间，单位 s。",
    "分离时间ts": "该级/助推器分离时间。程序采用 t >= ts 时剔除。",
    "抛离时间": "可抛部件抛离时间，单位 s。",
    "推进剂密度rho_phi": "推进剂密度，单位 kg/m^3，必须大于0。",
    "气体密度rho_g": "增压气体密度，单位 kg/m^3。",
    "秒流量mdot": "推进剂消耗质量流量，单位 kg/s。",
    "加注量Mf": "发动机点火前贮箱内推进剂/可变质量近似加注量，单位 kg。",
    "关机剩余量Mr": "发动机关机后剩余可变质量，单位 kg。",
    "启动后剩余量Mef": "启动活门打开后的剩余推进剂质量，预留用于QJ底部液位公式。",
    "坐标原点X或Loki": "贮箱局部坐标原点到全局 O 点的 X 向距离；程序使用 X_global = Loki - X_local。",
    "总容积Vo": "贮箱总容积，单位 m^3，建议与几何积分满箱体积一致。",
    "满箱Xo": "满箱推进剂质心在贮箱局部坐标系下的 X 坐标，单位 m。",
    "满箱Jyo": "满箱推进剂绕局部 Y 轴中心主惯量，单位 kg·m^2。",
    "是否启用扣除": "选择 是 或 否。否 表示第二类扣除半径 R2(x)=0。",
    "semi_a11": "图3中的 ā11，椭圆横向半轴，不是 tan(beta)。",
    "semi_a22": "图4中的 ā22，椭圆横向半轴，不是 tan(beta)。",
    "beta11_deg": "角度，单位为度，程序内部转换为弧度。",
    "beta12_deg": "角度，单位为度，程序内部转换为弧度。",
    "beta21_deg": "角度，单位为度，程序内部转换为弧度。",
    "beta22_deg": "角度，单位为度，程序内部转换为弧度。",
}


def add_instructions_sheet(wb) -> None:
    if "填写说明" in wb.sheetnames:
        del wb["填写说明"]
    ws = wb.create_sheet("填写说明", 0)
    rows = [
        ["主题", "说明"],
        ["模板定位", "Excel 只负责输入和结果查看，复杂计算由 Python 完成。"],
        ["坐标系", "O 为火箭理论顶点；X 轴沿火箭中心线指向底部为正；惯性积按 0 处理。"],
        ["单位", "质量 kg；长度/坐标 m；时间 s；密度 kg/m^3；流量 kg/s；转动惯量 kg·m^2。"],
        ["填写顺序", "总体参数 -> 级事件 -> 不变质量 -> 贮箱基本参数 -> 贮箱几何/外行导管/可抛质量。"],
        ["贮箱ID", "规则为 类型-所属编号-贮箱编号，例如 子级-1-1、助推器-1-2。"],
        ["几何类型", "等效圆柱：填写等效半径R和等效高度H；QJ第一类：填写贮箱几何_第一类，必要时填写第二类扣除。"],
        ["级间分离", "程序采用 t < ts 保留，t >= ts 剔除该级/助推器。"],
        ["颜色", "浅黄色/蓝字为输入区；灰色为备注说明；深蓝色为表头。"],
        ["Starship样例", "input_starship_example.xlsx 是公开数据近似验证样例，不代表 SpaceX 官方内部质量特性模型。"],
    ]
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 100
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    apply_grid(ws)
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        row[0].fill = PatternFill("solid", fgColor=INFO_FILL)
        row[0].font = Font(name="Arial", bold=True, color=BLACK)
        row[1].alignment = Alignment(vertical="center", wrap_text=True)


WIDTHS = {
    "参数": 24,
    "值": 40,
    "单位": 12,
    "说明": 56,
    "备注": 46,
    "名称": 26,
    "组件名称": 28,
    "部段名称": 24,
    "贮箱ID": 18,
    "类型": 12,
    "几何类型": 16,
    "坐标原点X或Loki": 20,
    "推进剂密度rho_phi": 20,
    "气体密度rho_g": 18,
    "秒流量mdot": 16,
    "启动后剩余量Mef": 18,
    "等效半径R": 16,
    "等效高度H": 16,
    "总容积Vo": 16,
    "满箱Xo": 16,
    "满箱Jyo": 18,
    "事件类型": 18,
    "对象": 18,
    "事件时间": 16,
    "匹配时间": 16,
}


def apply_grid(ws, max_row: int | None = None, max_col: int | None = None) -> None:
    thin = Side(style="thin", color=GRID)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    max_row = max_row or ws.max_row
    max_col = max_col or ws.max_column
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=False)


def add_dropdown(ws, cell_range: str, options: list[str]) -> None:
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.error = "请从下拉列表中选择。"
    dv.errorTitle = "无效选项"
    dv.prompt = "请选择：" + "、".join(options)
    dv.promptTitle = "可选值"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def add_decimal_validation(ws, cell_range: str, min_value: float = 0.0, allow_blank: bool = True) -> None:
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(min_value), allow_blank=allow_blank)
    dv.error = f"请输入大于等于 {min_value} 的数值。"
    dv.errorTitle = "数值范围错误"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def add_integer_validation(ws, cell_range: str, min_value: int = 0, allow_blank: bool = True) -> None:
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=str(min_value), allow_blank=allow_blank)
    dv.error = f"请输入大于等于 {min_value} 的整数。"
    dv.errorTitle = "整数范围错误"
    ws.add_data_validation(dv)
    dv.add(cell_range)


def header_map(ws) -> dict[str, int]:
    return {str(cell.value): cell.column for cell in ws[1] if cell.value is not None}


def range_for_col(ws, col_name: str, start: int = 2, end: int = 500) -> str | None:
    h = header_map(ws)
    if col_name not in h:
        return None
    letter = ws.cell(1, h[col_name]).column_letter
    return f"{letter}{start}:{letter}{end}"


def apply_validations(ws) -> None:
    for col in ["类型"]:
        rng = range_for_col(ws, col)
        if rng:
            add_dropdown(ws, rng, ["子级", "助推器"])
    rng = range_for_col(ws, "几何类型")
    if rng:
        add_dropdown(ws, rng, ["等效圆柱", "QJ第一类"])
    rng = range_for_col(ws, "是否启用扣除")
    if rng:
        add_dropdown(ws, rng, ["否", "是"])

    nonnegative_keywords = ["M", "质量", "密度", "秒流量", "加注量", "剩余量", "时间", "半径", "高度", "容积", "Jx", "Jy", "Jz", "R", "H", "h", "b", "semi_a"]
    integer_cols = {"编号", "所属编号", "贮箱编号", "导管编号"}
    for name, col in header_map(ws).items():
        rng = f"{ws.cell(1, col).column_letter}2:{ws.cell(1, col).column_letter}500"
        if name in integer_cols:
            add_integer_validation(ws, rng, 0)
        elif any(k in name for k in nonnegative_keywords):
            if name not in {"X", "Y", "Z", "Xcq", "Ycq", "Zcq", "坐标原点X或Loki", "坐标原点Y", "坐标原点Z"}:
                add_decimal_validation(ws, rng, 0)


def format_worksheet(ws) -> None:
    reserve_rows = max(INPUT_RESERVE_ROWS, ws.max_row)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A1:{ws.cell(1, ws.max_column).column_letter}{reserve_rows}"
    apply_grid(ws, reserve_rows, ws.max_column)
    if ws.max_row >= 1:
        ws.row_dimensions[1].height = 34
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if cell.value in HEADER_COMMENTS:
            cell.comment = Comment(HEADER_COMMENTS[cell.value], "OpenCode")
    for row in ws.iter_rows(min_row=2, max_row=reserve_rows, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = Font(name="Arial", size=10, color=BLUE)
            cell.fill = PatternFill("solid", fgColor=INPUT_FILL)
            cell.alignment = Alignment(vertical="center", wrap_text=False)
    if ws.max_column >= 1:
        for cell in ws.iter_cols(min_col=ws.max_column, max_col=ws.max_column, min_row=2, max_row=reserve_rows):
            for c in cell:
                c.fill = PatternFill("solid", fgColor=NOTE_FILL)
                c.font = Font(name="Arial", size=10, color=BLACK)
    for col in ws.columns:
        header = str(col[0].value) if col[0].value is not None else ""
        width = WIDTHS.get(header, max(14, min(max(len(str(c.value)) if c.value is not None else 0 for c in col[: min(len(col), 50)]) + 4, 42)))
        ws.column_dimensions[col[0].column_letter].width = width
    apply_validations(ws)


def add_sheet_notes(wb) -> None:
    for ws in wb.worksheets:
        if ws.title == "填写说明":
            continue
        note = SHEET_NOTES.get(ws.title)
        if note:
            ws["A1"].comment = Comment((ws["A1"].comment.text + "\n\n" if ws["A1"].comment else "") + note, "OpenCode")


def format_input_workbook(path: str) -> None:
    wb = load_workbook(path)
    add_instructions_sheet(wb)
    for ws in wb.worksheets:
        if ws.title != "填写说明":
            format_worksheet(ws)
    add_sheet_notes(wb)
    wb.save(path)


def format_output_workbook(path: str) -> None:
    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        max_row = max(ws.max_row + 20, 100)
        max_col = ws.max_column
        ws.auto_filter.ref = f"A1:{ws.cell(1, max_col).column_letter}{ws.max_row}"
        apply_grid(ws, max_row, max_col)
        ws.row_dimensions[1].height = 32
        for cell in ws[1]:
            cell.font = Font(name="Arial", bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.font = Font(name="Arial", size=10, color=BLACK)
                cell.fill = PatternFill("solid", fgColor=OUTPUT_FILL)
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                header = str(ws.cell(1, cell.column).value)
                if header in {"M", "M_v", "M_phi", "M_g", "Jx", "Jy", "Jz", "Jx_v", "Jy_v", "Jz_v", "Jx_phi", "Jy_phi", "Jz_phi"}:
                    cell.number_format = '#,##0.000'
                elif header in {"X", "Y", "Z", "h", "X_v", "Y_v", "Z_v", "X_phi", "Y_phi", "Z_phi", "X_g"}:
                    cell.number_format = '0.000'
                elif header in {"t", "事件时间", "匹配时间", "前一时间", "后一时间"}:
                    cell.number_format = '0.0'
        for col in ws.columns:
            header = str(col[0].value) if col[0].value is not None else ""
            width = WIDTHS.get(header, max(14, min(max(len(str(c.value)) if c.value is not None else 0 for c in col[: min(len(col), 80)]) + 4, 42)))
            ws.column_dimensions[col[0].column_letter].width = width
    wb.save(path)
