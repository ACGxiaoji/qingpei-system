# -*- coding: utf-8 -*-
"""
把“铁塔征租地补偿协议书.docx”里的内容填写到补偿表 xlsx。

使用方法：
1. 双击运行本脚本。
2. 选择青赔目录里的人名文件夹。
3. 选择该文件夹里的协议书。
4. 模板可选“自动判断 / 含培肥费 / 无培肥费”。
5. 点击生成，补偿表会另存到该人名文件夹里。

说明：
- 脚本通过复制 Excel 模板再填写单元格，尽量保持模板原格式不变。
- 协议书没有写明的信息不会乱填，例如地上附着物名称、规格等可能需要生成后手动补充。
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from shutil import copy2


def relaunch_with_bundled_python() -> None:
    runtime = (
        Path(os.environ.get("USERPROFILE", ""))
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "python"
        / "python.exe"
    )
    current = Path(sys.executable).resolve()
    if runtime.exists() and current != runtime.resolve():
        completed = subprocess.run([str(runtime), str(Path(__file__).resolve()), *sys.argv[1:]])
        raise SystemExit(completed.returncode)


relaunch_with_bundled_python()


try:
    from docx import Document
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
    from openpyxl.styles import Alignment
except ImportError as exc:
    print("缺少必要组件，无法运行。")
    print("请确认已安装：python-docx openpyxl")
    print(f"具体缺少：{exc}")
    input("按回车键退出...")
    raise SystemExit(1)


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "00_模板"


def template_file(*names: str) -> Path:
    """按顺序查找模板，兼容已使用过的旧文件名。"""
    for name in names:
        path = TEMPLATE_DIR / name
        if path.exists():
            return path
    return TEMPLATE_DIR / names[0]


TEMPLATE_WITH_FERTILIZER = TEMPLATE_DIR / "模板.xlsx"
TEMPLATE_NO_FERTILIZER = TEMPLATE_DIR / "模板（无培肥费版）.xlsx"
TEMPLATE_WITH_FERTILIZER_V2 = template_file(
    "征租地补偿表(有培肥).xlsx",
    "某风电项目征租地补偿表(有培肥).xlsx",
)
TEMPLATE_NO_FERTILIZER_V2 = template_file(
    "征租地补偿表(无培肥).xlsx",
    "某风电项目征租地补偿表(无培肥).xlsx",
)


@dataclass
class AgreementData:
    name: str = ""
    id_number: str = ""
    town: str = ""
    village: str = ""
    team: str = ""
    tower_ids: str = ""
    tower_count: int = 1
    bank_name: str = ""
    bank_card: str = ""
    permanent_area: float | None = None
    permanent_unit: float | None = None
    permanent_amount: float | None = None
    rent_area: float | None = None
    rent_years: float | None = None
    construction_area: float | None = None
    construction_unit: float | None = None
    construction_amount: float | None = None
    road_area: float | None = None
    road_unit: float | None = None
    road_amount: float | None = None
    ground_count: float | None = None
    ground_unit: float | None = None
    ground_amount: float | None = None
    attachment_unit: float | None = None
    attachment_area: float | None = None
    attachment_amount: float | None = None
    total_amount: float | None = None
    fertilizer_area: float | None = None
    fertilizer_unit: float | None = None
    fertilizer_amount: float | None = None
    has_fertilizer: bool = False


def normalize_text(text: str) -> str:
    return text.replace("\t", " ").replace("\u3000", " ")


def compact_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", normalize_text(text)).strip()


def to_number(value: str | None) -> float | None:
    if value is None:
        return None
    value = value.strip().replace(",", "")
    if not value or value in {"/", "\\", "／"}:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", value)
    if not match:
        return None
    return float(match.group(0))


def clean_number(value: float | None):
    if value is None:
        return None
    if abs(value - int(value)) < 1e-9:
        return int(value)
    return value


def amount_upper(value: float | None) -> str:
    if value is None:
        return ""
    value = round(float(value) + 1e-8, 2)
    if abs(value) < 1e-9:
        return ""
    digits = "零壹贰叁肆伍陆柒捌玖"
    units = ["", "拾", "佰", "仟"]
    sections = ["", "万", "亿"]
    integer = int(value)
    cents = int(round((value - integer) * 100))

    def section_text(number: int) -> str:
        text = ""
        zero = False
        for index in range(4):
            digit = number % 10
            if digit == 0:
                if text:
                    zero = True
            else:
                if zero:
                    text = digits[0] + text
                    zero = False
                text = digits[digit] + units[index] + text
            number //= 10
        return text

    if integer == 0:
        result = "零元"
    else:
        result_parts: list[str] = []
        section_index = 0
        need_zero = False
        while integer:
            section = integer % 10000
            if section == 0:
                need_zero = bool(result_parts)
            else:
                part = section_text(section) + sections[section_index]
                if need_zero:
                    result_parts.insert(0, digits[0])
                    need_zero = False
                result_parts.insert(0, part)
            integer //= 10000
            section_index += 1
        result = "".join(result_parts) + "元"

    jiao = cents // 10
    fen = cents % 10
    if cents == 0:
        return "（大写）" + result + "整"
    tail = ""
    if jiao:
        tail += digits[jiao] + "角"
    elif fen:
        tail += "零"
    if fen:
        tail += digits[fen] + "分"
    return "（大写）" + result + tail


def first_match(pattern: str, text: str) -> str:
    match = re.search(pattern, text, re.S)
    return match.group(1).strip() if match else ""


def parse_party(line: str, data: AgreementData) -> None:
    parse_location(line, data)
    match = re.search(
        r"某县\s*(?P<town>[\u4e00-\u9fff]{1,12})\s*镇\s*(?P<village>[\u4e00-\u9fff]{1,12})\s*村\s*(?P<team>[\u4e00-\u9fff]{1,12})\s*队\s*(?P<name>[\u4e00-\u9fff\s]+)",
        line,
    )
    if not match:
        return
    data.town = re.sub(r"\s+", "", match.group("town"))
    data.village = re.sub(r"\s+", "", match.group("village"))
    data.team = re.sub(r"\s+", "", match.group("team"))
    data.name = re.sub(r"\s+", "", match.group("name"))


def location_parts(line: str) -> tuple[str, str, str] | None:
    match = re.search(
        r"(?:某县\s*)?(?P<town>[\u4e00-\u9fff]{1,12})\s*(?:镇|乡|街道)\s*(?P<village>[\u4e00-\u9fff]{1,12}?)\s*(?:村|社区)(?:\s*(?P<team>[\u4e00-\u9fff0-9]{1,12}?)\s*(?:队|组))?",
        line,
    )
    if not match:
        return None
    return (
        re.sub(r"\s+", "", match.group("town")),
        re.sub(r"\s+", "", match.group("village")),
        re.sub(r"\s+", "", match.group("team") or ""),
    )


def parse_location(line: str, data: AgreementData) -> None:
    parts = location_parts(line)
    if parts is None:
        return
    town, village, team = parts
    data.town = data.town or town
    data.village = data.village or village
    data.team = data.team or team


def parse_tower_ids(value: str) -> str:
    value = re.sub(r"\s*铁塔.*$", "", value)
    value = re.sub(r"[，,、]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    if value in {"\\", "/", "／"}:
        return ""
    return value


def tower_count(tower_ids: str) -> int:
    parts = [p for p in re.split(r"\s+", tower_ids.strip()) if p]
    return max(1, len(parts))


def read_docx_text(path: Path) -> list[str]:
    document = Document(str(path))
    lines: list[str] = []
    seen: set[str] = set()

    def add_line(value: str) -> None:
        text = compact_spaces(value)
        if text and text not in seen:
            seen.add(text)
            lines.append(text)

    for paragraph in document.paragraphs:
        add_line(paragraph.text)
    # 部分旧协议把地址、塔号和合计金额写在表格单元格中。
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                add_line(cell.text)
    return lines


def parse_agreement(path: Path) -> AgreementData:
    lines = read_docx_text(path)
    text = "\n".join(lines)
    data = AgreementData()

    for line in lines:
        # 协议版本很多，地址可能出现在甲方信息或正文中，队别也可能留空。
        parse_location(line, data)
        if "土地坐落在" in line:
            # 只识别“土地坐落在”之后的铁塔位置，避免把前面的协议正文识别成镇名。
            tower_parts = location_parts(line.split("土地坐落在", 1)[1])
            if tower_parts is not None:
                data.town, data.village, data.team = tower_parts
        if "被征租方" in line or "被租方" in line or "被租用方" in line:
            parse_party(line, data)
        elif line.startswith("身份证号码"):
            value = first_match(r"身份证号码[:：]\s*([0-9Xx]{15,18})", line)
            data.id_number = value.upper()
        elif line.startswith("征租地位置") or line.startswith("租地位置"):
            data.tower_ids = parse_tower_ids(first_match(r"(?:征租地位置|租地位置)[:：]\s*(.*?)\s*铁塔", line))
            data.tower_count = tower_count(data.tower_ids)
        elif line.startswith("征地面积") or line.startswith("长期租地面积"):
            vals = re.findall(r"([\\/／\d.]+)\s*亩", line)
            if vals:
                data.permanent_area = to_number(vals[0])
        elif line.startswith("租地面积") or line.startswith("临时租地面积"):
            match = re.search(r"(?:租地面积|临时租地面积)[:：]\s*([\\/／\d.]+)\s*亩.*?租期\s*([\\/／\d.]+)\s*年", line)
            if match:
                data.rent_area = to_number(match.group(1))
                data.rent_years = to_number(match.group(2))
        elif line.startswith("其中施工用地面积"):
            match = re.search(r"施工用地面积\s*([\\/／\d.]+)\s*亩.*?道路用地面积\s*([\\/／\d.]+)\s*亩", line)
            if match:
                data.construction_area = to_number(match.group(1))
                data.road_area = to_number(match.group(2))
        elif line.startswith("林地："):
            match = re.search(r"林地[:：]\s*([\\/／\d.]+)\s*元/亩×\s*([\\/／\d.]+)\s*亩＝\s*([\\/／\d.]+)", line)
            if match:
                data.permanent_unit = to_number(match.group(1))
                data.permanent_area = data.permanent_area or to_number(match.group(2))
                data.permanent_amount = to_number(match.group(3))
        elif line.startswith("施工用地："):
            match = re.search(r"施工用地[:：]\s*([\\/／\d.]+)\s*元/亩×\s*([\\/／\d.]+)\s*年×\s*([\\/／\d.]+)\s*亩＝\s*([\\/／\d.]+)", line)
            if match:
                data.construction_unit = to_number(match.group(1))
                data.rent_years = data.rent_years or to_number(match.group(2))
                data.construction_area = data.construction_area or to_number(match.group(3))
                data.construction_amount = to_number(match.group(4))
        elif line.startswith("道路用地："):
            match = re.search(r"道路用地[:：]\s*([\\/／\d.]+)\s*元/亩×\s*([\\/／\d.]+)\s*年×\s*([\\/／\d.]+)\s*亩＝\s*([\\/／\d.]+)", line)
            if match:
                data.road_unit = to_number(match.group(1))
                data.rent_years = data.rent_years or to_number(match.group(2))
                data.road_area = data.road_area or to_number(match.group(3))
                data.road_amount = to_number(match.group(4))
        elif line.startswith("接地："):
            match = re.search(r"接地[:：]\s*([\\/／\d.]+)\s*元/基×\s*([\\/／\d.]+)\s*基＝\s*([\\/／\d.]+)", line)
            if match:
                data.ground_unit = to_number(match.group(1))
                data.ground_count = to_number(match.group(2))
                data.ground_amount = to_number(match.group(3))
        elif line.startswith("地上附作物：") or re.match(r"^[\\/／\d.]+\s*元/亩×", line):
            match = re.search(r"(?:地上附作物[:：]\s*)?([\\/／\d.]+)\s*元/亩×\s*([\\/／\d.]+)\s*亩＝\s*([\\/／\d.]+)", line)
            if match:
                area = to_number(match.group(2))
                amount = to_number(match.group(3))
                unit = to_number(match.group(1))
                if area is None and amount is None and unit is None:
                    continue
                data.attachment_unit = data.attachment_unit or unit
                data.attachment_area = (data.attachment_area or 0) + (area or 0)
                data.attachment_amount = (data.attachment_amount or 0) + (amount or 0)
        elif line.startswith("培肥补偿费"):
            match = re.search(r"培肥补偿费[:：]\s*([\\/／\d.]+)\s*元/亩×\s*([\\/／\d.]+)\s*亩＝\s*([\\/／\d.]+)", line)
            if match:
                data.fertilizer_unit = to_number(match.group(1))
                data.fertilizer_area = to_number(match.group(2))
                data.fertilizer_amount = to_number(match.group(3))
                data.has_fertilizer = True
        elif line.startswith("以上补偿费用合计为"):
            data.total_amount = to_number(line)
        elif line.startswith("姓") and "名" in line:
            value = first_match(r"姓\s*名[:：]\s*([\u4e00-\u9fff\s]+)", line)
            if value:
                data.name = re.sub(r"\s+", "", value)
        elif line.startswith("身份证号"):
            value = first_match(r"身份证号[:：]\s*([0-9Xx]{15,18})", line)
            if value:
                data.id_number = value.upper()
        elif line.startswith("卡"):
            value = first_match(r"卡\s*号[:：]\s*([0-9 ]{10,25})", line)
            if value:
                data.bank_card = re.sub(r"\D", "", value)
        elif line.startswith("银"):
            value = first_match(r"银\s*行[:：]\s*(.+)", line)
            if value:
                data.bank_name = value.strip()

    known = sum(
        x or 0
        for x in [
            data.permanent_amount,
            data.construction_amount,
            data.road_amount,
            data.ground_amount,
            data.attachment_amount,
        ]
    )
    if data.total_amount is not None:
        diff = round(data.total_amount - known, 6)
        area = (data.construction_area or 0) + (data.road_area or 0)
        expected = area * 1500 if area else 0
        if diff > 0.05 and (not expected or abs(diff - expected) < 0.2):
            data.has_fertilizer = True
            data.fertilizer_area = area or None
            data.fertilizer_unit = 1500
            data.fertilizer_amount = diff

    return data


def choose_template(data: AgreementData, template_choice: str) -> Path:
    if template_choice == "with":
        return TEMPLATE_WITH_FERTILIZER
    if template_choice == "without":
        return TEMPLATE_NO_FERTILIZER
    if template_choice == "with_v2":
        return TEMPLATE_WITH_FERTILIZER_V2
    if template_choice == "without_v2":
        return TEMPLATE_NO_FERTILIZER_V2
    return TEMPLATE_WITH_FERTILIZER if data.has_fertilizer else TEMPLATE_NO_FERTILIZER


def is_with_fertilizer_template(template_path: Path) -> bool:
    return template_path in {TEMPLATE_WITH_FERTILIZER, TEMPLATE_WITH_FERTILIZER_V2}


def template_label(template_path: Path) -> str:
    labels = {
        TEMPLATE_WITH_FERTILIZER: "含培肥费",
        TEMPLATE_NO_FERTILIZER: "无培肥费",
        TEMPLATE_WITH_FERTILIZER_V2: "含培肥费2.0",
        TEMPLATE_NO_FERTILIZER_V2: "无培肥费2.0",
    }
    return labels.get(template_path, template_path.stem)


def set_if(ws, cell: str, value) -> None:
    if isinstance(value, str) and value.strip() in {"\\", "/", "／"}:
        return
    if value is not None and value != "":
        ws[cell] = value


def clear_cells(ws, cells: list[str]) -> None:
    for cell in cells:
        if isinstance(ws[cell], MergedCell):
            continue
        ws[cell] = None


def make_rich_text(parts: list[tuple[str, bool]], size: float = 14.0) -> CellRichText:
    rich = CellRichText()
    for text, underline in parts:
        font = InlineFont(rFont="宋体", sz=size, u="single" if underline else None)
        rich.append(TextBlock(font, text))
    return rich


def route_line_rich(data: AgreementData) -> CellRichText:
    return make_rich_text(
        [
            ("线路路经：某县", False),
            (f"   {data.town}  ", True),
            ("镇", False),
            (f"  {data.village} ", True),
            ("村", False),
            (f"  {data.team}  ", True),
            ("队，共", False),
            (f"  {data.tower_count}  ", True),
            ("基铁塔。", False),
        ]
    )


def reason_line_rich(data: AgreementData) -> CellRichText:
    return make_rich_text(
        [
            ("补偿原因：\n    因某新能源风电项目EPC工程35kV集电线路工程经过某县", False),
            (f"   {data.town} ", True),
            ("镇", False),
            (f" {data.village} ", True),
            ("村", False),
            (f"   {data.team}  ", True),
            ("队，需对铁塔基础租地青苗和地上附着物补偿。", False),
        ]
    )


def clear_template_sample_values(wb, with_fertilizer: bool) -> None:
    ws = wb["兑现表"]
    ws["M2"] = "年           月        日"
    clear_cells(
        ws,
        [
            "B5",
            "C5",
            "D5",
            "E5",
            "F5",
            "G5",
            "H5",
            "I5",
            "J5",
            "F7",
            "G7",
            "H7",
            "I7",
            "J7",
            "K7",
            "L7",
            "M7",
            "N7",
            "O7",
            "G8",
            "H8",
            "I8",
            "J8",
            "F10",
            "G10",
            "H10",
            "I10",
            "J10",
            "K10",
            "L10",
            "M10",
            "N10",
            "O10",
            "F11",
            "G11",
            "H11",
            "I11",
            "J11",
            "P5",
            "P16",
        ],
    )

    land = wb["土地结算表"]
    clear_cells(
        land,
        [
            "B5",
            "C5",
            "D6",
            "D7",
            "D8",
            "D9",
            "E5",
            "E6",
            "E7",
            "E8",
            "E9",
            "F5",
            "F6",
            "F7",
            "F8",
            "F9",
            "G5",
            "H5",
            "I5",
            "J5",
            "C10",
            "C11",
            "H11",
            "L5",
            "L6",
            "L7",
            "L8",
            "L9",
        ],
    )

    attach = wb["附作物结算表"]
    clear_cells(attach, ["B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5"])


def nearly_equal(a: float | None, b: float | None, tolerance: float = 0.0001) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) <= tolerance


def fit_tower_cell(cell, tower_ids: str) -> None:
    text = str(tower_ids or "")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, shrink_to_fit=True)
    if not text:
        return
    length = len(text)
    size = 11
    if length > 34:
        size = 8
    elif length > 24:
        size = 9
    elif length > 16:
        size = 10
    cell.font = cell.font.copy(sz=size)


def sum_values(values: list[float | None]) -> float | None:
    known = [value for value in values if value is not None]
    if not known:
        return None
    return sum(known)


def write_land_total(ws, total_amount: float | None) -> None:
    if total_amount is None:
        return

    total_value = clean_number(total_amount)
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or "").strip() == "总金额:":
                amount_cell = ws.cell(cell.row, 3)
                upper_cell = ws.cell(cell.row, 8)
                amount_cell.value = total_value
                upper_cell.value = amount_upper(total_amount)
                return


def apply_number_formats(wb) -> None:
    number_format = "General"

    ws = wb["兑现表"]
    for addr in ["K7", "L7", "K10", "L10"]:
        ws[addr].number_format = "General"
    for row in range(5, ws.max_row + 1):
        for col in ["G", "H", "I", "J", "M", "N", "O", "P"]:
            ws[f"{col}{row}"].number_format = number_format

    land = wb["土地结算表"]
    for row in range(5, land.max_row + 1):
        for col in ["C", "D", "E", "F", "G", "H", "I", "J"]:
            land[f"{col}{row}"].number_format = number_format

    attach = wb["附作物结算表"]
    for row in range(5, attach.max_row + 1):
        for col in ["C", "E", "F", "G", "H", "I", "J"]:
            attach[f"{col}{row}"].number_format = number_format


def fill_common_workbook(wb, data: AgreementData, with_fertilizer: bool) -> None:
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    clear_template_sample_values(wb, with_fertilizer)

    ws = wb["兑现表"]
    land_place = "".join(part for part in [data.town + "镇" if data.town else "", data.village + "村" if data.village else ""])
    tower_place = land_place + (data.team + "队" if data.team else "")

    set_if(ws, "B5", data.tower_ids)
    fit_tower_cell(ws["B5"], data.tower_ids)
    set_if(ws, "C5", land_place)
    set_if(ws, "D5", data.name)
    set_if(ws, "E5", data.id_number)

    if data.permanent_area is not None:
        set_if(ws, "F5", "林地")
        set_if(ws, "G5", clean_number(data.permanent_area))
        set_if(ws, "H5", "20年")
        set_if(ws, "I5", clean_number(data.permanent_unit))
        if data.permanent_unit is not None:
            ws["J5"] = "=G5*I5"

    if data.construction_area is not None:
        set_if(ws, "F7", "林地")
        set_if(ws, "G7", clean_number(data.construction_area))
        set_if(ws, "H7", clean_number(data.rent_years))
        set_if(ws, "I7", clean_number(data.construction_unit))
        if data.rent_years is not None and data.construction_unit is not None:
            ws["J7"] = "=G7*H7*I7"

    if data.ground_count is not None:
        set_if(ws, "G8", f"{clean_number(data.ground_count)}基" if data.ground_count is not None else None)
        set_if(ws, "H8", "/")
        set_if(ws, "I8", clean_number(data.ground_unit))
        set_if(ws, "J8", clean_number(data.ground_amount))

    if data.road_area is not None:
        set_if(ws, "F10", "林地")
        set_if(ws, "G10", clean_number(data.road_area))
        set_if(ws, "H10", clean_number(data.rent_years))
        set_if(ws, "I10", clean_number(data.road_unit))
        if data.rent_years is not None and data.road_unit is not None:
            ws["J10"] = "=G10*H10*I10"

    if data.attachment_area is not None and data.attachment_unit is not None:
        if nearly_equal(data.attachment_area, (data.construction_area or 0) + (data.road_area or 0)) and data.construction_area and data.road_area:
            set_if(ws, "M7", clean_number(data.construction_area))
            set_if(ws, "N7", clean_number(data.attachment_unit))
            ws["O7"] = "=M7*N7"
            set_if(ws, "M10", clean_number(data.road_area))
            set_if(ws, "N10", clean_number(data.attachment_unit))
            ws["O10"] = "=M10*N10"
        else:
            set_if(ws, "M7", clean_number(data.attachment_area))
            set_if(ws, "N7", clean_number(data.attachment_unit))
            ws["O7"] = "=M7*N7"

    has_cash_amount = any(
        value is not None
        for value in [
            data.permanent_amount,
            data.construction_amount,
            data.road_amount,
            data.ground_amount,
            data.attachment_amount,
            data.fertilizer_amount,
        ]
    )

    if with_fertilizer:
        if data.fertilizer_area is not None:
            set_if(ws, "F11", "林地")
            if data.construction_area is not None or data.road_area is not None:
                ws["G11"] = "=G10+G7"
            else:
                set_if(ws, "G11", clean_number(data.fertilizer_area))
            set_if(ws, "H11", "/")
            set_if(ws, "I11", 1500)
            ws["J11"] = "=G11*I11"
        if has_cash_amount and not ws["P5"].value:
            ws["P5"] = "=SUM(J5,J7,J8,J10,J11,O5,O7,O8,O10)"
        if has_cash_amount and not ws["P16"].value:
            ws["P16"] = "=P5"
    else:
        if has_cash_amount:
            ws["P5"] = "=SUM(J5,J7,J8,J10,O5,O7,O10)"
            ws["P16"] = "=P5"

    land = wb["土地结算表"]
    land["A3"] = route_line_rich(data)
    land["C4"] = "长期租地面积（亩）"
    land["D4"] = "临时租地2年面积（亩）"
    has_land_rows = any(
        value is not None
        for value in [
            data.permanent_area,
            data.construction_area,
            data.road_area,
            data.ground_count,
            data.fertilizer_area,
            data.permanent_amount,
            data.construction_amount,
            data.road_amount,
            data.ground_amount,
            data.fertilizer_amount,
        ]
    )
    if has_land_rows:
        set_if(land, "B5", data.name)
        set_if(land, "G5", data.tower_ids)
        fit_tower_cell(land["G5"], data.tower_ids)
        set_if(land, "H5", tower_place)
        set_if(land, "I5", data.bank_name)
        set_if(land, "J5", data.bank_card)

    if data.permanent_area is not None:
        land["C5"] = clean_number(data.permanent_area)
        land["E5"] = clean_number(data.permanent_unit)
        land["F5"] = clean_number(data.permanent_amount)
        land["L5"] = "长期租地补偿"

    if data.construction_area is not None:
        land["D6"] = clean_number(data.construction_area)
        land["E6"] = clean_number(data.construction_unit)
        land["F6"] = clean_number(data.construction_amount)
        land["L6"] = "临时用地补偿"

    if data.ground_count is not None:
        land["E7"] = clean_number(data.ground_unit)
        land["F7"] = clean_number(data.ground_amount)
        land["L7"] = "接地线补偿"

    if with_fertilizer:
        if data.road_area is not None:
            land["D8"] = clean_number(data.road_area)
            land["E8"] = clean_number(data.road_unit)
            land["F8"] = clean_number(data.road_amount)
            land["L8"] = "施工道路补偿"
        if data.fertilizer_area is not None:
            land["D9"] = clean_number(data.fertilizer_area)
            land["E9"] = 1500
            land["F9"] = clean_number(data.fertilizer_amount)
            land["L9"] = "培肥补偿"
        land_total = sum_values(
            [
                data.permanent_amount,
                data.construction_amount,
                data.ground_amount,
                data.road_amount,
                data.fertilizer_amount,
            ]
        )
    else:
        if data.road_area is not None:
            land["D8"] = clean_number(data.road_area)
            land["E8"] = clean_number(data.road_unit)
            land["F8"] = clean_number(data.road_amount)
            land["L8"] = "施工道路补偿"
        land_total = sum_values(
            [
                data.permanent_amount,
                data.construction_amount,
                data.ground_amount,
                data.road_amount,
            ]
        )
    land["C10"] = clean_number(land_total)
    write_land_total(land, land_total)

    attach = wb["附作物结算表"]
    attach["A3"] = reason_line_rich(data)
    set_if(attach, "B5", data.name)
    set_if(attach, "E5", clean_number(data.attachment_area))
    set_if(attach, "F5", clean_number(data.attachment_unit))
    if data.attachment_area is not None and data.attachment_unit is not None:
        attach["G5"] = "=E5*F5"
    set_if(attach, "H5", data.tower_ids)
    fit_tower_cell(attach["H5"], data.tower_ids)
    set_if(attach, "I5", data.bank_name)
    set_if(attach, "J5", data.bank_card)
    attachment_total = data.attachment_amount
    if attachment_total is None and data.attachment_area is not None and data.attachment_unit is not None:
        attachment_total = data.attachment_area * data.attachment_unit
    attach["C9"] = clean_number(attachment_total)
    attach["C10"] = clean_number(attachment_total)
    attach["H10"] = amount_upper(attachment_total)
    apply_number_formats(wb)


def generate_table(agreement_path: Path, output_path: Path, template_choice: str = "auto") -> tuple[Path, AgreementData, str]:
    data = parse_agreement(agreement_path)
    template_path = choose_template(data, template_choice)
    if not template_path.exists():
        raise FileNotFoundError(f"没有找到模板：{template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    copy2(template_path, output_path)
    wb = load_workbook(output_path, rich_text=True)
    with_fertilizer = is_with_fertilizer_template(template_path)
    fill_common_workbook(wb, data, with_fertilizer)
    wb.save(output_path)
    return output_path, data, template_label(template_path)


def default_output_path(person_dir: Path, data: AgreementData) -> Path:
    name = data.name or person_dir.name
    return person_dir / f"某风电项目-{name}补偿表.xlsx"


def person_dirs() -> list[Path]:
    return [
        p
        for p in sorted(BASE_DIR.iterdir(), key=lambda x: x.name)
        if p.is_dir() and not p.name.startswith("00_") and not p.name.startswith("__") and p.name != "证件提取"
    ]


def agreement_files(person_dir: Path) -> list[Path]:
    return [
        p
        for p in sorted(person_dir.glob("*.docx"), key=lambda x: x.name)
        if not p.name.startswith("~$") and ("协议" in p.name or "征租地" in p.name)
    ]


def run_gui() -> int:
    import tkinter as tk
    from tkinter import messagebox

    dirs = person_dirs()
    if not dirs:
        messagebox.showerror("没有文件夹", "青赔目录里没有找到人名文件夹。")
        return 1

    root = tk.Tk()
    root.title("协议书生成补偿表")
    root.geometry("680x460")
    root.minsize(620, 420)

    tk.Label(root, text="选择人名文件夹：", anchor="w").pack(fill="x", padx=14, pady=(12, 4))
    folder_box = tk.Listbox(root, height=8, exportselection=False)
    for folder in dirs:
        folder_box.insert(tk.END, folder.name)
    folder_box.selection_set(0)
    folder_box.pack(fill="both", expand=True, padx=14, pady=4)

    tk.Label(root, text="选择协议书：", anchor="w").pack(fill="x", padx=14, pady=(8, 4))
    doc_var = tk.StringVar()
    doc_menu = tk.OptionMenu(root, doc_var, "")
    doc_menu.pack(fill="x", padx=14)

    template_var = tk.StringVar(value="自动判断")
    tk.Label(root, text="模板选择：", anchor="w").pack(fill="x", padx=14, pady=(8, 4))
    tk.OptionMenu(root, template_var, "自动判断", "含培肥费", "无培肥费", "含培肥费2.0", "无培肥费2.0").pack(fill="x", padx=14)

    status_var = tk.StringVar(value="")
    tk.Label(root, textvariable=status_var, anchor="w", fg="#333333").pack(fill="x", padx=14, pady=(8, 4))

    current_docs: list[Path] = []

    def refresh_docs(*_):
        nonlocal current_docs
        selection = folder_box.curselection()
        if not selection:
            return
        folder = dirs[selection[0]]
        current_docs = agreement_files(folder)
        menu = doc_menu["menu"]
        menu.delete(0, "end")
        if not current_docs:
            doc_var.set("")
            status_var.set("这个文件夹里没有找到协议书 docx。")
            return
        for doc in current_docs:
            menu.add_command(label=doc.name, command=lambda value=doc.name: doc_var.set(value))
        doc_var.set(current_docs[0].name)
        status_var.set("")

    folder_box.bind("<<ListboxSelect>>", refresh_docs)
    refresh_docs()

    def generate():
        selection = folder_box.curselection()
        if not selection:
            messagebox.showwarning("请选择", "请先选择人名文件夹。")
            return
        folder = dirs[selection[0]]
        doc_name = doc_var.get()
        agreement_path = next((p for p in current_docs if p.name == doc_name), None)
        if agreement_path is None:
            messagebox.showwarning("请选择", "请先选择协议书。")
            return

        try:
            data = parse_agreement(agreement_path)
            output_path = default_output_path(folder, data)
            if output_path.exists():
                stem = output_path.stem
                output_path = output_path.with_name(stem + "_新.xlsx")

            choice = {
                "自动判断": "auto",
                "含培肥费": "with",
                "无培肥费": "without",
                "含培肥费2.0": "with_v2",
                "无培肥费2.0": "without_v2",
            }[template_var.get()]
            output, _, template_label = generate_table(agreement_path, output_path, choice)
        except PermissionError:
            messagebox.showerror("生成失败", "目标 Excel 可能正在打开，请关闭后重试。")
            return
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))
            return

        messagebox.showinfo("完成", f"已生成：\n{output}\n\n使用模板：{template_label}")

    buttons = tk.Frame(root)
    buttons.pack(pady=12)
    tk.Button(buttons, text="生成补偿表", width=16, command=generate).pack(side="left", padx=8)
    tk.Button(buttons, text="退出", width=12, command=root.destroy).pack(side="left", padx=8)

    root.mainloop()
    return 0


def run_cli(args: argparse.Namespace) -> int:
    agreement_path = Path(args.agreement).resolve()
    if not agreement_path.exists():
        print(f"协议书不存在：{agreement_path}")
        return 1
    data = parse_agreement(agreement_path)
    output = Path(args.output).resolve() if args.output else default_output_path(agreement_path.parent, data)
    out, parsed, template_label = generate_table(agreement_path, output, args.template)
    print(f"已生成：{out}")
    print(f"模板：{template_label}")
    print(f"姓名：{parsed.name} 身份证：{parsed.id_number} 铁塔：{parsed.tower_ids}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="把协议书内容填写到补偿表")
    parser.add_argument("--no-gui", action="store_true", help="不用窗口")
    parser.add_argument("--agreement", help="协议书 docx 路径")
    parser.add_argument("--output", help="输出 xlsx 路径")
    parser.add_argument("--template", choices=["auto", "with", "without", "with_v2", "without_v2"], default="auto", help="模板选择")
    args = parser.parse_args()
    if args.no_gui:
        return run_cli(args)
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
