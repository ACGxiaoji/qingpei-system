# -*- coding: utf-8 -*-
"""
一键生成青赔资料

作用：
1. 从“证件提取/人名文件夹”里的证件图片生成身份证、银行卡 Word。
2. 识别证件资料，补充缺少的协议信息，生成铁塔租地补偿协议书。
3. 根据协议书生成补偿表 Excel。
4. 三个新文件统一放到“青赔/人名/”文件夹里。

旧脚本不会被修改，本脚本只调用它们已有的识别、排版、转表能力。
"""

from __future__ import annotations

import os
import re
import importlib
import math
import random
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
from copy import copy, deepcopy
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from itertools import combinations
from pathlib import Path
from typing import Any, Callable


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
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.shared import Cm, Pt
    from openpyxl import load_workbook
except ImportError as exc:
    print("缺少 Word/窗口处理组件，无法运行。")
    print("请先安装：pip install python-docx")
    print(f"具体缺少：{exc}")
    input("按回车键退出...")
    raise SystemExit(1)


BASE_DIR = Path(__file__).resolve().parent
CERT_DIR = BASE_DIR / "证件提取"
TEMPLATE_DIR = BASE_DIR / "00_模板"
MASTER_EXTRACT_DIR_NAME = "_总文档整理"
LOCAL_COMPONENT_DIR = BASE_DIR / "_运行组件"
LOCAL_PY_PACKAGES = LOCAL_COMPONENT_DIR / "python_packages"
OCR_PIP_PACKAGES = ["rapidocr-onnxruntime", "onnxruntime", "opencv-python"]
PIP_INDEX_URL = "https://pypi.tuna.tsinghua.edu.cn/simple"

COMMON_ATTACHMENT_OPTIONS = [
    ("桉树", "棵"),
    ("速生桉", "棵"),
    ("马尾松", "棵"),
    ("松树", "棵"),
    ("杉树", "棵"),
    ("杂木", "棵"),
    ("相思树", "棵"),
    ("竹子", "棵"),
    ("荔枝树", "棵"),
    ("龙眼树", "棵"),
    ("芒果树", "棵"),
    ("八角树", "棵"),
    ("油茶树", "棵"),
    ("红椎", "棵"),
    ("香蕉", "亩"),
    ("甘蔗", "亩"),
    ("花生", "亩"),
    ("玉米", "亩"),
    ("木薯", "亩"),
    ("其他", "亩"),
]

AGREEMENT_TEMPLATE_WITH = TEMPLATE_DIR / "某风电项目风电场铁塔征租地补偿协议书.docx"
AGREEMENT_TEMPLATE_WITHOUT = TEMPLATE_DIR / "模板（无培肥费版本）.docx"
AGREEMENT_TEMPLATE_WITH_V2 = TEMPLATE_DIR / "模板（有培肥2.0）.docx"
AGREEMENT_TEMPLATE_WITHOUT_V2 = TEMPLATE_DIR / "模板（无培肥2.0）.docx"

TEMPLATE_OPTIONS = {
    "without": {
        "label": "无培肥费版本",
        "agreement": AGREEMENT_TEMPLATE_WITHOUT,
        "table": "without",
        "with_fertilizer": False,
        "version": "old",
    },
    "with": {
        "label": "有培肥费版本",
        "agreement": AGREEMENT_TEMPLATE_WITH,
        "table": "with",
        "with_fertilizer": True,
        "version": "old",
    },
    "without_v2": {
        "label": "无培肥2.0",
        "agreement": AGREEMENT_TEMPLATE_WITHOUT_V2,
        "table": "without_v2",
        "with_fertilizer": False,
        "version": "2.0",
    },
    "with_v2": {
        "label": "有培肥2.0",
        "agreement": AGREEMENT_TEMPLATE_WITH_V2,
        "table": "with_v2",
        "with_fertilizer": True,
        "version": "2.0",
    },
}

sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(CERT_DIR))
if LOCAL_PY_PACKAGES.exists():
    sys.path.insert(0, str(LOCAL_PY_PACKAGES))

try:
    import agreement_to_table
    import fill_template
    import generate_id_doc
    import organize_master_doc as master_doc_tools
    import qingpei_db
except Exception as exc:
    print("旧脚本加载失败，无法运行。")
    print(f"具体原因：{exc}")
    input("按回车键退出...")
    raise SystemExit(1)

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


@dataclass
class GeneratedFiles:
    cert_doc: Path | None
    agreement_doc: Path
    table_xlsx: Path | None
    output_dir: Path
    warnings: list[str]


@dataclass
class TablePersonEntry:
    label: str
    source_path: Path
    row_numbers: list[int]
    data: Any
    tower_ids: str = ""
    attachment_items: list[dict[str, str]] = field(default_factory=list)
    seedling_total: str = ""
    payment_status: str = ""


def safe_filename(value: str) -> str:
    value = (value or "").strip() or "未命名"
    return re.sub(r'[\\/:*?"<>|]+', "_", value)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(1, 1000):
        candidate = path.with_name(f"{stem}_新{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"同名文件太多，无法自动命名：{path}")


def clean_number_text(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = text.replace("，", ".").replace(",", "")
    try:
        number = float(text)
    except ValueError:
        return text
    if abs(number - round(number)) < 0.0000001:
        return str(int(round(number)))
    return f"{number:.6f}".rstrip("0").rstrip(".")


def to_float(value: Any) -> float | None:
    text = clean_number_text(value)
    if not text or text in {"/", "／"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def multiply(*values: Any) -> float | None:
    result = 1.0
    for value in values:
        number = to_float(value)
        if number is None:
            return None
        result *= number
    return result


def amount_text(value: float | None) -> str:
    if value is None:
        return ""
    return clean_number_text(round(value, 6))


def rent_area(data: dict[str, str]) -> str:
    temporary = clean_number_text(data.get("temporary_area", ""))
    if temporary:
        return temporary
    construction = to_float(data.get("construction_area"))
    road = to_float(data.get("road_area"))
    if construction is not None or road is not None:
        return clean_number_text((construction or 0) + (road or 0))
    return ""


def location_to_team_address(address: str, town: str, village: str, team: str) -> str:
    """将甲方地址整理为镇、村、队层级，供铁塔地名默认使用。"""
    town, village, team = qingpei_db.normalize_location_parts(town, village, team)
    address_town, address_village, address_team = qingpei_db.split_address(address)
    town = town or address_town
    village = village or address_village
    team = team or address_team
    if town or village or team:
        parts = []
        if town:
            parts.append(f"{town}镇")
        if village:
            parts.append(f"{village}村")
        if team:
            parts.append(f"{team}队")
        return "".join(parts)
    return re.sub(r"\s+", "", address)


def normalize_form_data(data: dict[str, str]) -> dict[str, str]:
    normalized = dict(data)
    town, village, team = qingpei_db.normalize_location_parts(
        normalized.get("town", ""), normalized.get("village", ""), normalized.get("team", "")
    )
    normalized["town"] = town
    normalized["village"] = village
    normalized["team"] = team
    tower_location = re.sub(r"\s+", "", normalized.get("tower_location", ""))
    if not tower_location:
        tower_location = location_to_team_address(
            normalized.get("address", ""), town, village, team
        )
    tower_town, tower_village, tower_team = qingpei_db.split_address(tower_location)
    normalized["tower_location"] = tower_location
    normalized["tower_town"] = tower_town
    normalized["tower_village"] = tower_village
    normalized["tower_team"] = tower_team
    temporary = clean_number_text(normalized.get("temporary_area", ""))
    construction = clean_number_text(normalized.get("construction_area", ""))
    road = clean_number_text(normalized.get("road_area", ""))
    if temporary and not construction and not road:
        normalized["construction_area"] = temporary
    elif not temporary:
        calculated = rent_area(normalized)
        if calculated:
            normalized["temporary_area"] = calculated
    return normalized


def parse_attachment_items(data: dict[str, str]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    index = 1
    while True:
        suffix = "" if index == 1 else f"_{index}"
        keys = [
            f"attachment_type{suffix}",
            f"attachment_name{suffix}",
            f"attachment_spec{suffix}",
            f"attachment_area{suffix}",
            f"attachment_unit{suffix}",
            f"attachment_land_use{suffix}",
        ]
        if index > 1 and not any(key in data for key in keys):
            break

        item = {
            "type": str(data.get(keys[0], "")).strip(),
            "name": str(data.get(keys[1], "")).strip(),
            "spec": str(data.get(keys[2], "")).strip(),
            "area": clean_number_text(data.get(keys[3], "")),
            "unit": clean_number_text(data.get(keys[4], "")),
            "land_use": str(data.get(keys[5], "施工用地")).strip() or "施工用地",
        }
        if any(item.get(key) for key in ["type", "name", "spec", "area", "unit"]):
            items.append(item)
        index += 1
    return items


def attachment_item_amount(item: dict[str, str]) -> float | None:
    return multiply(item.get("unit"), item.get("area"))


def attachment_total_amount(items: list[dict[str, str]]) -> float | None:
    total = 0.0
    has_value = False
    for item in items:
        amount = attachment_item_amount(item)
        if amount is None:
            continue
        total += amount
        has_value = True
    return total if has_value else None


SEEDLING_PLANTS = (
    ("龙眼树", "棵", Decimal("800"), 1),
    ("荔枝树", "棵", Decimal("800"), 1),
    ("桉树", "14cm", Decimal("3200"), 1000),
    ("桉树", "10cm", Decimal("3000"), 1000),
)


def seedling_price_units(plant: tuple[str, str, Decimal, int]) -> int:
    return int(plant[2] * 1000 / plant[3])


def parse_seedling_budget(value: str) -> int:
    text = str(value or "").strip().replace(",", "")
    if not text:
        raise ValueError("请先填写青苗总额。")
    try:
        budget = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("青苗总额请输入数字，例如 50000 或 50000.2。") from exc
    if budget <= 0:
        raise ValueError("青苗总额必须大于 0。")
    units = budget * 1000
    if units != units.to_integral_value():
        raise ValueError("青苗总额最多可输入小数点后三位。")
    return int(units)


def sample_two_seedling_quantities(first: int, second: int, total: int, random_source: random.Random) -> tuple[int, int] | None:
    if total < 0:
        return None
    divisor = math.gcd(first, second)
    if total % divisor:
        return None
    first_reduced = first // divisor
    second_reduced = second // divisor
    total_reduced = total // divisor
    first_base = 0 if second_reduced == 1 else (total_reduced * pow(first_reduced, -1, second_reduced)) % second_reduced
    second_base = (total_reduced - first_reduced * first_base) // second_reduced
    if second_base < 0:
        return None
    max_step = min(
        (total_reduced // first_reduced - first_base) // second_reduced,
        second_base // first_reduced,
    )
    if max_step < 0:
        return None
    step = random_source.randint(0, max_step)
    return first_base + second_reduced * step, second_base - first_reduced * step


def sample_three_seedling_quantities(first: int, second: int, third: int, total: int, random_source: random.Random) -> tuple[int, int, int] | None:
    remaining = total - first - second - third
    if remaining < 0:
        return None
    pair_divisor = math.gcd(first, second)
    all_divisor = math.gcd(pair_divisor, third)
    if remaining % all_divisor:
        return None
    modulus = pair_divisor // all_divisor
    third_base = 0 if modulus == 1 else ((remaining // all_divisor) * pow((third // all_divisor) % modulus, -1, modulus)) % modulus
    max_third = remaining // third
    if third_base > max_third:
        return None
    count = (max_third - third_base) // modulus + 1
    for _ in range(min(count, 500)):
        third_extra = third_base + random_source.randrange(count) * modulus
        sampled = sample_two_seedling_quantities(first, second, remaining - third * third_extra, random_source)
        if sampled is not None:
            return sampled[0] + 1, sampled[1] + 1, third_extra + 1
    return None


def decimal_text(value: Decimal, places: int = 3) -> str:
    text = f"{value:.{places}f}"
    return text.rstrip("0").rstrip(".") if "." in text else text


def simulate_seedling_items(
    total_units: int,
    requested_count: int | None = None,
    random_source: random.Random | None = None,
) -> list[dict[str, str]] | None:
    if requested_count is not None and requested_count not in (1, 2, 3):
        raise ValueError("预算模拟的种类数只能选择 1、2 或 3 种。")
    source = random_source or random.SystemRandom()
    candidates: list[tuple[tuple[int, int], ...]] = []
    group_sizes = (requested_count,) if requested_count is not None else range(1, 4)
    for group_size in group_sizes:
        for indexes in combinations(range(len(SEEDLING_PLANTS)), group_size):
            if 2 in indexes and 3 in indexes:
                continue
            prices = [seedling_price_units(SEEDLING_PLANTS[index]) for index in indexes]
            if group_size == 1:
                quantity = total_units // prices[0]
                quantities = (quantity,) if quantity > 0 and quantity * prices[0] == total_units else None
            elif group_size == 2:
                sampled = sample_two_seedling_quantities(prices[0], prices[1], total_units - prices[0] - prices[1], source)
                quantities = (sampled[0] + 1, sampled[1] + 1) if sampled is not None else None
            else:
                quantities = sample_three_seedling_quantities(prices[0], prices[1], prices[2], total_units, source)
            if quantities is not None:
                candidates.append(tuple(zip(indexes, quantities)))
    if not candidates:
        return None

    items: list[dict[str, str]] = []
    for index, quantity in source.choice(candidates):
        name, spec, price, scale = SEEDLING_PLANTS[index]
        amount = Decimal(quantity) / scale
        items.append(
            {
                "land_use": "施工用地",
                "name": name,
                "spec": spec,
                "area": decimal_text(amount),
                "unit": decimal_text(price),
            }
        )
    return items


def excel_number_value(value: Any):
    number = to_float(value)
    if number is None:
        return None
    if abs(number - round(number)) < 0.0000001:
        return int(round(number))
    return number


def field_value(data: dict[str, str], key: str, blank: str = "\\") -> str:
    text = clean_number_text(data.get(key, "")) if key.endswith(("_area", "_unit", "_years", "_count")) else str(data.get(key, "")).strip()
    return text if text else blank


def underline_value(value: str, width: int, left: int = 1, right: int = 1) -> str:
    text = str(value or "").strip()
    is_blank = not text
    if is_blank:
        return (" " * min(left, 1)) + "\\"
    if not is_blank and re.fullmatch(r"\d+(?:\.\d+)?", text):
        return (" " * min(left, 1)) + text
    if not is_blank and len(text) >= width:
        return (" " * min(left, 1)) + text
    wanted = max(width, len(text) + left + right)
    rest = wanted - len(text)
    left_spaces = max(left, rest // 2)
    right_spaces = max(right, rest - left_spaces)
    return (" " * left_spaces) + text + (" " * right_spaces)


def has_input_value(data: dict[str, str], key: str) -> bool:
    text = clean_number_text(data.get(key, ""))
    return bool(text and text not in {"\\", "/", "／", "0", "0.0", "0.00"})


def should_write_attachment_row(data: dict[str, str], area_key: str) -> bool:
    if has_input_value(data, area_key):
        return True
    if not has_input_value(data, "attachment_area"):
        return False
    construction = to_float(data.get("construction_area"))
    road = to_float(data.get("road_area"))
    attachment = to_float(data.get("attachment_area"))
    if attachment is None:
        return False
    if area_key == "construction_area" and construction is not None and abs(attachment - construction) < 0.0001:
        return True
    if area_key == "road_area" and road is not None and abs(attachment - road) < 0.0001:
        return True
    if construction is not None and road is None and area_key == "construction_area":
        return True
    if road is not None and construction is None and area_key == "road_area":
        return True
    return False


def rewrite_paragraph(paragraph, parts: list[tuple[str, bool]], rpr=None) -> None:
    if rpr is None and paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        rpr = deepcopy(paragraph.runs[0]._r.rPr)

    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)

    for text, underline in parts:
        run = paragraph.add_run(text)
        if rpr is not None:
            run._r.insert(0, deepcopy(rpr))
        run.underline = bool(underline)
        stripped = text.strip()
        if underline and len(stripped) > 30:
            run.font.size = Pt(8 if len(stripped) > 45 else 9)


def insert_paragraph_after(paragraph, parts: list[tuple[str, bool]]) -> Paragraph:
    source_ppr = deepcopy(paragraph._p.pPr) if paragraph._p.pPr is not None else None
    source_rpr = deepcopy(paragraph.runs[0]._r.rPr) if paragraph.runs and paragraph.runs[0]._r.rPr is not None else None
    new_p = OxmlElement("w:p")
    if source_ppr is not None:
        new_p.append(source_ppr)
    paragraph._p.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if paragraph.style is not None:
        new_paragraph.style = paragraph.style
    if paragraph.paragraph_format is not None:
        new_paragraph.paragraph_format.alignment = paragraph.paragraph_format.alignment
        new_paragraph.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
        new_paragraph.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
        new_paragraph.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
        new_paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_before
        new_paragraph.paragraph_format.space_after = paragraph.paragraph_format.space_after
        new_paragraph.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
    rewrite_paragraph(new_paragraph, parts, rpr=source_rpr)
    return new_paragraph


def attachment_line_parts(item: dict[str, str], include_label: bool = True) -> list[tuple[str, bool]]:
    unit = item.get("unit", "")
    area = item.get("area", "")
    amount = amount_text(attachment_item_amount(item))
    measure = "棵" if item.get("spec", "").strip() == "棵" else "亩"
    prefix = "地上附作物：  " if include_label else "              "
    return [
        (prefix, False),
        (underline_value(unit, 6, left=1, right=1), True),
        (f"元/{measure}×", False),
        (underline_value(area, 6, left=1, right=1), True),
        (f"{measure}＝", False),
        (underline_value(amount, 7, left=1, right=1), True),
        ("  元；", False),
    ]


def find_paragraph(document: Document, startswith: str):
    for paragraph in document.paragraphs:
        if paragraph.text.strip().startswith(startswith):
            return paragraph
    return None


def set_line(document: Document, startswith: str, parts: list[tuple[str, bool]]) -> None:
    paragraph = find_paragraph(document, startswith)
    if paragraph is not None:
        rewrite_paragraph(paragraph, parts)


def set_land_policy_price(document: Document, unit_price: str = "") -> None:
    unit_price = clean_number_text(unit_price)
    if not unit_price:
        return
    pattern = re.compile(r"(征地片区地价林地)\s*(?:[\\/／\d.]*)\s*(元/亩)")
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if "征地片区地价林地" not in text or "元/亩" not in text:
            continue
        match = pattern.search(text)
        if not match:
            continue
        rewrite_paragraph(
            paragraph,
            [
                (text[: match.start(1)] + match.group(1), False),
                (underline_value(unit_price, 6, left=1, right=1), True),
                (match.group(2) + text[match.end(2) :], False),
            ],
        )
        return


def compact_agreement_pages(document: Document, template_path: Path, with_fertilizer: bool) -> None:
    """Keep agreements near six pages while preserving signature space."""
    for section in document.sections:
        if section.top_margin is not None and section.top_margin.cm > 2.0:
            section.top_margin = Cm(2.0)
        if section.bottom_margin is not None and section.bottom_margin.cm > 2.0:
            section.bottom_margin = Cm(2.0)

    is_v2_template = "2.0" in template_path.name
    start_other_terms_on_new_page = is_v2_template != with_fertilizer
    for paragraph in document.paragraphs:
        if paragraph.text.strip().startswith("六、违约责任"):
            paragraph.paragraph_format.page_break_before = True
            continue
        if paragraph.text.strip().startswith("八、其他条款"):
            paragraph.paragraph_format.page_break_before = start_other_terms_on_new_page
            break


def display_name(name: str, gap: str = " ") -> str:
    name = (name or "").strip()
    if len(name) <= 1:
        return name
    return gap.join(name)


def template_info(choice: str) -> dict[str, Any]:
    return TEMPLATE_OPTIONS.get(choice, TEMPLATE_OPTIONS["without"])


def is_fertilizer_choice(choice: str) -> bool:
    return bool(template_info(choice)["with_fertilizer"])


def build_agreement(template_path: Path, output_path: Path, data: dict[str, str], with_fertilizer: bool) -> None:
    data = normalize_form_data(data)
    if not template_path.exists():
        raise FileNotFoundError(f"没有找到协议模板：{template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    document = Document(str(output_path))

    name = data.get("name", "").strip()
    town = data.get("town", "").strip()
    village = data.get("village", "").strip()
    team = data.get("team", "").strip()
    tower_town = data.get("tower_town", "").strip() or town
    tower_village = data.get("tower_village", "").strip() or village
    tower_team = data.get("tower_team", "").strip() or team
    tower = data.get("tower_ids", "").strip()
    id_number = data.get("id_number", "").strip()
    bank_card = data.get("bank_card", "").strip()
    bank_name = data.get("bank_name", "").strip()
    is_v2_template = "2.0" in template_path.name
    use_word = "租用" if is_v2_template else "征租"

    permanent_area = field_value(data, "permanent_area")
    construction_area = field_value(data, "construction_area")
    road_area_value = field_value(data, "road_area")
    temporary_area = field_value(data, "temporary_area")
    rent_area_value = rent_area(data) or "\\"
    rent_years = field_value(data, "rent_years")
    permanent_unit_source = clean_number_text(data.get("permanent_unit", ""))
    permanent_unit = field_value(data, "permanent_unit")
    construction_unit = field_value(data, "construction_unit")
    road_unit = field_value(data, "road_unit")
    ground_unit = field_value(data, "ground_unit")
    ground_count = field_value(data, "ground_count")
    attachment_unit = field_value(data, "attachment_unit")
    attachment_area = field_value(data, "attachment_area")
    fertilizer_unit = field_value(data, "fertilizer_unit")
    fertilizer_area = field_value(data, "fertilizer_area")
    attachment_items = parse_attachment_items(data)
    if not attachment_items:
        attachment_items = [
            {
                "type": "",
                "name": str(data.get("attachment_name", "")).strip(),
                "spec": str(data.get("attachment_spec", "")).strip(),
                "area": clean_number_text(data.get("attachment_area", "")),
                "unit": clean_number_text(data.get("attachment_unit", "")),
            }
        ] if any(str(data.get(key, "")).strip() for key in ["attachment_name", "attachment_spec", "attachment_area", "attachment_unit"]) else []

    permanent_amount = multiply(permanent_unit_source, data.get("permanent_area"))
    construction_amount = multiply(data.get("construction_unit"), data.get("rent_years"), data.get("construction_area"))
    road_amount = multiply(data.get("road_unit"), data.get("rent_years"), data.get("road_area"))
    ground_amount = multiply(data.get("ground_unit"), data.get("ground_count"))
    attachment_amount = attachment_total_amount(attachment_items)
    fertilizer_amount = multiply(data.get("fertilizer_unit"), data.get("fertilizer_area")) if with_fertilizer else None
    total_amount = sum(value or 0 for value in [permanent_amount, construction_amount, road_amount, ground_amount, attachment_amount, fertilizer_amount])
    total_text = amount_text(total_amount) if total_amount else ""

    set_line(
        document,
        "被征租方",
        [
            ("被征租方(甲方)：某县", False),
            (underline_value(town, 4), True),
            ("镇", False),
            (underline_value(village, 4), True),
            ("村", False),
            (underline_value(team, 5), True),
            ("队", False),
            (underline_value(name, 5), True),
            ("  ", True),
        ],
    )
    set_line(
        document,
        "被租方",
        [
            ("被租方(甲方)：某县", False),
            (underline_value(town, 4), True),
            ("镇", False),
            (underline_value(village, 4), True),
            ("村", False),
            (underline_value(team, 5), True),
            ("队", False),
            (underline_value(name, 5), True),
            ("  ", True),
        ],
    )
    set_line(
        document,
        "被租用方",
        [
            ("被租用方(甲方)：某县", False),
            (underline_value(town, 4), True),
            ("镇", False),
            (underline_value(village, 4), True),
            ("村", False),
            (underline_value(team, 5), True),
            ("队", False),
            (underline_value(name, 5), True),
            ("  ", True),
        ],
    )
    set_line(document, "身份证号码", [("身份证号码：  ", False), (underline_value(id_number, 22, left=2, right=2), True), ("  ", True)])
    set_line(document, "见证方", [("见证方：某县  ", False), (underline_value(town, 6), True), ("  镇人民政府", False)])
    set_land_policy_price(document, permanent_unit_source)
    set_line(
        document,
        "经见证方证明",
        [
            ("经见证方证明，甲方有权代表甲方家庭所有成员签订和领取土地坐落在", False),
            (underline_value(tower_town, 4), True),
            ("镇", True),
            (underline_value(tower_village, 4), True),
            ("村", True),
            (underline_value(tower_team, 5), True),
            ("队", True),
            (underline_value(tower, 0, left=1, right=0), True),
            (" 铁塔", True),
            (f"的全部补偿款。{use_word}范围以实地测量结果为准。", False),
        ],
    )
    set_line(document, "征租地位置", [("征租地位置：  ", False), (underline_value(tower, 0, left=1, right=0), True), ("铁塔。", False)])
    set_line(document, "租地位置", [("租地位置：  ", False), (underline_value(tower, 0, left=1, right=0), True), ("铁塔。", False)])
    set_line(
        document,
        "征地面积",
        [("征地面积：    ", False), (underline_value(permanent_area, 8, left=2, right=2), True), ("  亩（林地  ", False), (underline_value(permanent_area, 8, left=2, right=2), True), ("  亩）。", False)],
    )
    set_line(
        document,
        "长期租地面积",
        [("长期租地面积：    ", False), (underline_value(permanent_area, 8, left=2, right=2), True), ("  亩（林地  ", False), (underline_value(permanent_area, 8, left=2, right=2), True), ("，租期  20 年，自本协议生效之日起计算）。", False)],
    )
    set_line(
        document,
        "租地面积",
        [("租地面积：   ", False), (underline_value(rent_area_value, 8, left=2, right=2), True), ("  亩（租期  ", False), (underline_value(rent_years, 4, left=1, right=1), True), (" 年）", False)],
    )
    set_line(
        document,
        "临时租地面积",
        [("临时租地面积：   ", False), (underline_value(temporary_area if temporary_area != "\\" else rent_area_value, 8, left=2, right=2), True), ("  亩（租期  ", False), (underline_value(rent_years, 4, left=1, right=1), True), ("  年，自本协议生效之日起计算）", False)],
    )
    set_line(
        document,
        "其中施工用地面积",
        [("其中施工用地面积  ", False), (underline_value(construction_area, 8, left=2, right=2), True), ("  亩，道路用地面积  ", False), (underline_value(road_area_value, 8, left=2, right=2), True), ("  亩。", False)],
    )
    set_line(
        document,
        "林地：",
        [
            ("林地：    ", False),
            (underline_value(permanent_unit, 6, left=1, right=1), True),
            ("元/亩×", False),
            (underline_value(permanent_area, 6, left=1, right=1), True),
            ("亩＝", False),
            (underline_value(amount_text(permanent_amount), 7, left=1, right=1), True),
            ("  元；", False),
        ],
    )
    set_line(
        document,
        "施工用地",
        [
            ("施工用地：   ", False),
            (underline_value(construction_unit, 5, left=1, right=1), True),
            ("元/亩×", False),
            (underline_value(rent_years, 3, left=1, right=1), True),
            ("年×", False),
            (underline_value(construction_area, 6, left=1, right=1), True),
            ("亩＝", False),
            (underline_value(amount_text(construction_amount), 7, left=1, right=1), True),
            ("  元；", False),
        ],
    )
    set_line(
        document,
        "道路用地",
        [
            ("道路用地：   ", False),
            (underline_value(road_unit, 5, left=1, right=1), True),
            ("元/亩×", False),
            (underline_value(rent_years, 3, left=1, right=1), True),
            ("年×", False),
            (underline_value(road_area_value, 6, left=1, right=1), True),
            ("亩＝", False),
            (underline_value(amount_text(road_amount), 7, left=1, right=1), True),
            ("  元；", False),
        ],
    )
    set_line(
        document,
        "接地：",
        [
            ("接地：  ", False),
            (underline_value(ground_unit, 5, left=1, right=1), True),
            ("元/基×", False),
            (underline_value(ground_count, 4, left=1, right=1), True),
            ("基＝", False),
            (underline_value(amount_text(ground_amount), 7, left=1, right=1), True),
            ("  元；", False),
        ],
    )
    attachment_anchor = find_paragraph(document, "地上附作物")
    if attachment_anchor is not None and attachment_items:
        current = attachment_anchor
        for index, item in enumerate(attachment_items):
            parts = attachment_line_parts(item, include_label=index == 0)
            if index == 0:
                rewrite_paragraph(current, parts)
            else:
                current = insert_paragraph_after(current, parts)
    elif attachment_anchor is not None:
        rewrite_paragraph(
            attachment_anchor,
            [
                ("地上附作物：  ", False),
                (underline_value("\\", 6, left=1, right=1), True),
                ("元/亩×", False),
                (underline_value("\\", 6, left=1, right=1), True),
                ("亩＝", False),
                (underline_value("\\", 7, left=1, right=1), True),
                ("  元；", False),
            ],
        )
    if with_fertilizer:
        set_line(
            document,
            "培肥补偿费",
            [
                ("培肥补偿费：", False),
                (underline_value(fertilizer_unit, 5, left=0, right=1), True),
                ("元/亩×", False),
                (underline_value(fertilizer_area, 5, left=1, right=1), True),
                ("亩＝", False),
                (underline_value(amount_text(fertilizer_amount), 6, left=1, right=1), True),
                ("  元；", False),
            ],
        )
    set_line(document, "以上补偿费用合计为", [("以上补偿费用合计为：  ", False), (underline_value(total_text, 9, left=1, right=2), True), ("  元。", False)])

    set_line(document, "姓    名", [("姓    名：", False), (underline_value(name, 16, left=7, right=6), True)])
    set_line(document, "身份证号：", [("身份证号： ", False), (underline_value(id_number, 21, left=1, right=2), True), (" ", True)])
    set_line(document, "卡", [("卡    号： ", False), (underline_value(bank_card, 23, left=1, right=2), True), (" ", True)])
    set_line(document, "银", [("银    行： ", False), (underline_value(bank_name, 16, left=1, right=2), True), ("    ", True)])

    compact_agreement_pages(document, template_path, with_fertilizer)
    document.save(str(output_path))


def person_folders() -> list[Path]:
    if not CERT_DIR.exists():
        return []
    folders: list[Path] = []
    for path in sorted(CERT_DIR.iterdir(), key=lambda p: p.name):
        if not path.is_dir():
            continue
        if path.name.startswith(".") or path.name.startswith("_"):
            continue
        if path.name == MASTER_EXTRACT_DIR_NAME:
            continue
        folders.append(path)

    master_root = CERT_DIR / MASTER_EXTRACT_DIR_NAME
    if master_root.exists():
        for path in sorted(master_root.iterdir(), key=lambda p: p.name):
            if not path.is_dir():
                continue
            if path.name.startswith(".") or path.name.startswith("_") or path.name == master_doc_tools.CACHE_DIR_NAME:
                continue
            folders.append(path)
    return folders


def person_folder_label(path: Path) -> str:
    if path.parent.name == MASTER_EXTRACT_DIR_NAME:
        return f"{path.name}（总文档）"
    return path.name


def table_files() -> list[Path]:
    if not CERT_DIR.exists():
        return []
    files: list[Path] = []
    for path in sorted(CERT_DIR.glob("*.xlsx"), key=lambda p: p.name):
        if path.name.startswith("~$") or path.name.startswith("_"):
            continue
        files.append(path)
    return files


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def table_cell_text(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def table_merged_cell_text(sheet, row_index: int, column_index: int) -> str:
    value = sheet.cell(row_index, column_index).value
    if value not in (None, ""):
        return table_cell_text(value)
    for merged in sheet.merged_cells.ranges:
        if merged.min_row <= row_index <= merged.max_row and merged.min_col <= column_index <= merged.max_col:
            return table_cell_text(sheet.cell(merged.min_row, merged.min_col).value)
    return ""


def find_table_header(ws) -> tuple[int, dict[str, int]]:
    wanted = {"被租方", "位置"}
    for row_index in range(1, min(ws.max_row, 30) + 1):
        headers: dict[str, int] = {}
        for col_index in range(1, ws.max_column + 1):
            text = normalize_header(ws.cell(row_index, col_index).value)
            if text:
                headers[text] = col_index
        if wanted.issubset(headers):
            return row_index, headers
    raise ValueError("表格里没有找到包含“被租方、位置”的表头行。")


def get_header_col(headers: dict[str, int], *names: str) -> int | None:
    normalized = {normalize_header(name) for name in names}
    for header, col_index in headers.items():
        if header in normalized:
            return col_index
    return None


def get_seedling_fee_col(headers: dict[str, int]) -> int | None:
    for header, col_index in headers.items():
        if "青苗" in header and ("补偿费" in header or "补偿金额" in header):
            return col_index
    return None


def normalize_bank_main_name(value: str) -> str:
    text = re.sub(r"\s+", "", str(value or ""))
    for name in [
        "中国农业银行",
        "中国工商银行",
        "中国建设银行",
        "中国银行",
        "中国邮政储蓄银行",
        "交通银行",
        "招商银行",
        "广西北部湾银行",
        "广西农村信用社",
        "农村商业银行",
        "农村信用社",
        "桂林银行",
    ]:
        if name in text:
            return name
    return text


def normalize_table_location(location: str) -> str:
    location = re.sub(r"\s+", "", location or "")
    if not location:
        return ""
    if "县" not in location and re.search(r"(镇|乡|街道)", location):
        return "某县" + location
    return location


def table_tower_text(pole_no: str, pile_no: str) -> str:
    pole_no = table_cell_text(pole_no)
    pile_no = table_cell_text(pile_no)
    if pole_no and pile_no and pole_no != pile_no:
        return f"{pole_no}/{pile_no}"
    return pile_no or pole_no


def parse_attachment_items_from_workbook(workbook) -> dict[str, list[dict[str, str]]]:
    if "附作物结算表" not in workbook.sheetnames:
        return {}
    sheet = workbook["附作物结算表"]
    person_name = table_cell_text(sheet["B5"].value)
    if not person_name:
        return {}
    items: list[dict[str, str]] = []
    for row_index in range(5, sheet.max_row + 1):
        name = table_cell_text(sheet[f"C{row_index}"].value)
        if name in {"", "合计", "总金额:"}:
            continue
        spec = table_cell_text(sheet[f"D{row_index}"].value)
        area = clean_number_text(sheet[f"E{row_index}"].value)
        unit = clean_number_text(sheet[f"F{row_index}"].value)
        if not any([spec, area, unit]):
            continue
        items.append(
            {
                "land_use": "施工用地",
                "name": name,
                "spec": spec,
                "area": area,
                "unit": unit,
            }
        )
    return {person_name: items} if items else {}


def parse_table_people(path: Path) -> list[TablePersonEntry]:
    wb = load_workbook(path, data_only=True)
    entries: list[TablePersonEntry] = []
    attachment_items_by_name = parse_attachment_items_from_workbook(wb)
    for ws in wb.worksheets:
        try:
            header_row, headers = find_table_header(ws)
        except ValueError:
            continue
        name_col = get_header_col(headers, "被租方", "姓名", "户主姓名")
        location_col = get_header_col(headers, "位置", "地名", "地址")
        id_col = get_header_col(headers, "被租方身份证", "身份证号", "身份证号码")
        bank_col = get_header_col(headers, "银行卡/存折", "银行卡号", "户主账号")
        bank_name_col = get_header_col(headers, "支付银行", "开户银行", "银行名称")
        pole_col = get_header_col(headers, "杆塔号", "塔号")
        pile_col = get_header_col(headers, "铁塔桩号", "铁塔编号", "桩号")
        fee_col = get_header_col(headers, "补偿费用", "补偿金额")
        seedling_fee_col = get_seedling_fee_col(headers)
        payment_col = get_header_col(headers, "支付情况", "支付状态")
        if not name_col:
            continue

        for row_index in range(header_row + 1, ws.max_row + 1):
            name = table_cell_text(ws.cell(row_index, name_col).value)
            if not name:
                continue
            id_number = table_cell_text(ws.cell(row_index, id_col).value).upper() if id_col else ""
            bank_card = fill_template.normalize_digits(table_cell_text(ws.cell(row_index, bank_col).value)) if bank_col else ""
            bank_name = normalize_bank_main_name(ws.cell(row_index, bank_name_col).value) if bank_name_col else ""
            location = normalize_table_location(table_cell_text(ws.cell(row_index, location_col).value) if location_col else "")
            town, village, team = qingpei_db.split_address(location)
            tower = table_tower_text(
                ws.cell(row_index, pole_col).value if pole_col else "",
                ws.cell(row_index, pile_col).value if pile_col else "",
            )
            fee = clean_number_text(ws.cell(row_index, fee_col).value) if fee_col else ""
            seedling_total = clean_number_text(ws.cell(row_index, seedling_fee_col).value) if seedling_fee_col else ""
            payment_status = table_merged_cell_text(ws, row_index, payment_col) if payment_col else ""
            data = fill_template.PersonData(
                name=name,
                address=location,
                town=town,
                village=village,
                team=team,
                id_number=id_number,
                bank_card=bank_card,
                bank_name=bank_name,
            )
            entry_label = f"第{row_index}行 - {name}"
            if tower:
                entry_label += f" - {tower}"
            if fee:
                entry_label += f" - {fee}元"
            entries.append(
                TablePersonEntry(
                    entry_label,
                    path,
                    [row_index],
                    data,
                    tower,
                    attachment_items=deepcopy(attachment_items_by_name.get(name, [])),
                    seedling_total=seedling_total,
                    payment_status=payment_status,
                )
            )
    return entries


def extract_data(folder: Path):
    return fill_template.extract_person_data(folder, use_ocr=True)


def add_local_component_path() -> None:
    if not LOCAL_PY_PACKAGES.exists():
        return
    package_path = str(LOCAL_PY_PACKAGES)
    if package_path not in sys.path:
        sys.path.insert(0, package_path)


def reset_ocr_cache() -> None:
    for module in (fill_template, generate_id_doc):
        if hasattr(module, "_OCR_ENGINE"):
            module._OCR_ENGINE = None


def local_ocr_components_ready() -> bool:
    return all(
        (LOCAL_PY_PACKAGES / name).exists()
        for name in ("rapidocr_onnxruntime", "onnxruntime", "cv2")
    )


def missing_ocr_imports() -> list[str]:
    add_local_component_path()
    missing: list[str] = []
    for module_name in ("rapidocr_onnxruntime", "onnxruntime", "cv2"):
        try:
            importlib.import_module(module_name)
        except Exception:
            missing.append(module_name)
    return missing


def install_ocr_components(progress: Callable[[int, str], None] | None = None) -> str:
    if progress is not None:
        progress(10, "正在准备 OCR 固定组件目录")
    LOCAL_PY_PACKAGES.mkdir(parents=True, exist_ok=True)
    add_local_component_path()

    command = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--upgrade",
        "--target",
        str(LOCAL_PY_PACKAGES),
        *OCR_PIP_PACKAGES,
        "-i",
        PIP_INDEX_URL,
    ]
    if progress is not None:
        progress(14, "正在下载/修复 OCR 组件，首次可能需要稍等")
    completed = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "OCR 组件自动修复失败。\n\n"
            "可以检查网络后重新打开软件再试。\n\n"
            f"修复日志：\n{completed.stdout[-3000:]}"
        )
    importlib.invalidate_caches()
    add_local_component_path()
    reset_ocr_cache()
    if progress is not None:
        progress(32, "OCR 组件修复完成，正在验证")
    return completed.stdout


def ensure_ocr_ready(progress: Callable[[int, str], None] | None = None) -> bool:
    if progress is not None:
        progress(6, "正在检查 OCR 组件")
    add_local_component_path()
    reset_ocr_cache()

    if local_ocr_components_ready() and not missing_ocr_imports() and ocr_available():
        if progress is not None:
            progress(36, "OCR 组件正常")
        return True

    try:
        install_ocr_components(progress=progress)
    except Exception:
        reset_ocr_cache()
        if ocr_available():
            if progress is not None:
                progress(36, "OCR 当前可用，固定组件稍后可再修复")
            return True
        raise

    reset_ocr_cache()
    ready = not missing_ocr_imports() and ocr_available()
    if progress is not None:
        progress(36, "OCR 组件正常" if ready else "OCR 组件仍未正常加载")
    return ready


def ocr_available() -> bool:
    add_local_component_path()
    if getattr(fill_template, "_OCR_ENGINE", None) is False:
        reset_ocr_cache()
    return bool(fill_template.get_ocr_engine())


def run_self_check(progress: Callable[[int, str], None] | None = None) -> list[str]:
    def notify(value: int, message: str) -> None:
        if progress is not None:
            progress(value, message)

    lines: list[str] = []
    if not ensure_ocr_ready(progress=notify):
        raise RuntimeError("OCR 组件没有正常加载，自动识别不可用。")
    notify(8, "正在检查 OCR 组件")
    if not ocr_available():
        raise RuntimeError("OCR 组件没有正常加载，自动识别不可用。")
    lines.append("OCR 组件正常")

    notify(20, "正在检查人名文件夹列表")
    folders = person_folders()
    labels = [person_folder_label(path) for path in folders]
    if any(label.startswith("_") or "缓存" in label for label in labels):
        raise RuntimeError("下拉列表里混入了缓存或临时目录。")
    lines.append(f"可选择人名文件夹：{len(labels)} 个")

    master_root = CERT_DIR / MASTER_EXTRACT_DIR_NAME
    if master_root.exists():
        people = [path for path in master_root.iterdir() if path.is_dir() and not path.name.startswith("_")]
        image_count = 0
        for person in people:
            image_count += len([path for path in person.iterdir() if path.is_file() and path.suffix.lower() in generate_id_doc.IMAGE_EXTENSIONS])
        lines.append(f"总文档整理结果：{len(people)} 人，{image_count} 张图片")

    notify(42, "正在检查地址识别")
    address_cases = [
        (CERT_DIR / "童鹏", ("松旺", "周北", "老村")),
        (master_root / "刘丽珍", ("双旺", "周旺", "桐斜")),
        (master_root / "卜晓华", ("沙陂", "那新", "大坡")),
        (master_root / "张鑫", ("沙陂", "八壁", "射欠上高")),
        (master_root / "黄孙锦", ("沙陂", "良村", "亚妈坡")),
    ]
    checked = 0
    build_sample: tuple[Path, Any] | None = None
    for folder, expected in address_cases:
        if not folder.exists():
            continue
        data = extract_data(folder)
        got = (data.town, data.village, data.team)
        if got != expected:
            raise RuntimeError(f"{folder.name} 镇村队识别异常：识别为 {got}，应为 {expected}。")
        if not data.name or not data.id_number:
            raise RuntimeError(f"{folder.name} 自动识别缺少姓名或身份证号。")
        checked += 1
        if build_sample is None:
            build_sample = (folder, data)
    lines.append(f"典型地址样本：{checked} 个通过")

    notify(72, "正在检查临时生成")
    if build_sample is not None:
        folder, data = build_sample
        form = {
            "name": data.name or folder.name,
            "address": data.address,
            "town": data.town,
            "village": data.village,
            "team": data.team,
            "id_number": data.id_number,
            "bank_card": data.bank_card,
            "bank_name": data.bank_name,
            "tower_ids": "T001",
            "permanent_area": "1",
            "temporary_area": "1",
            "construction_area": "1",
            "road_area": "",
            "rent_years": "20",
            "permanent_unit": "48600",
            "construction_unit": "400",
            "road_unit": "400",
            "ground_count": "",
            "ground_unit": "400",
            "attachment_area": "",
            "attachment_unit": "",
            "template_choice": "without",
        }
        global BASE_DIR
        old_base = BASE_DIR
        with tempfile.TemporaryDirectory(prefix="qingpei_self_check_") as temp_root:
            BASE_DIR = Path(temp_root)
            try:
                result = build_all(folder, form)
            finally:
                BASE_DIR = old_base
            if not result.cert_doc or not result.cert_doc.exists():
                raise RuntimeError("临时生成证件 Word 失败。")
            if not result.agreement_doc.exists():
                raise RuntimeError("临时生成协议书失败。")
            if not result.table_xlsx or not result.table_xlsx.exists():
                raise RuntimeError("临时生成补偿表失败。")
        lines.append("临时生成证件 Word、协议书、补偿表正常")
    else:
        lines.append("未找到典型样本，已跳过临时生成检查")

    notify(100, "自检完成")
    return lines


def build_no_cert_doc(person_name: str, output_path: Path) -> None:
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(11)

    paragraph = document.add_paragraph()
    paragraph.alignment = 1
    run = paragraph.add_run("无身份证和银行卡文档")
    run.bold = True
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(16)

    name_paragraph = document.add_paragraph()
    name_paragraph.alignment = 1
    name_run = name_paragraph.add_run(f"姓名：{person_name}")
    name_run.font.name = "宋体"
    name_run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    name_run.font.size = Pt(12)
    document.save(str(output_path))


def generate_cert_doc(person_folder: Path, person_name: str, output_dir: Path, no_id_doc: bool = False) -> Path | None:
    output_path = unique_path(output_dir / f"{safe_filename(person_name)}.docx")
    if no_id_doc:
        build_no_cert_doc(person_name, output_path)
        return output_path

    images = generate_id_doc.collect_images(person_folder)
    if not images:
        return None
    arranged = generate_id_doc.arrange_images(images)
    generate_id_doc.build_document(person_name, arranged, output_path)
    return output_path


def copy_cell_style(source, target) -> None:
    target._style = copy(source._style)
    target.number_format = source.number_format
    target.alignment = copy(source.alignment)
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.protection = copy(source.protection)


def copy_row_style(ws, source_row: int, target_row: int, max_col: int = 12) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    ws.row_dimensions[target_row].hidden = ws.row_dimensions[source_row].hidden
    for col in range(1, max_col + 1):
        copy_cell_style(ws.cell(source_row, col), ws.cell(target_row, col))


def clear_worksheet_merges(ws) -> None:
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))


def clear_attachment_sheet_merges(ws) -> None:
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 5:
            ws.unmerge_cells(str(merged))


def extend_cash_detail_merges(ws, last_detail_row: int) -> None:
    for col_index, col_letter in [(3, "C"), (4, "D"), (5, "E"), (16, "P")]:
        for merged in list(ws.merged_cells.ranges):
            if merged.min_col == col_index and merged.max_col == col_index and merged.min_row == 5:
                ws.unmerge_cells(str(merged))
        ws.merge_cells(f"{col_letter}5:{col_letter}{last_detail_row}")


def ensure_cash_attachment_rows(ws, items_count: int, has_fertilizer_row: bool) -> tuple[list[int], int]:
    base_rows = [7, 9, 10, 12, 13, 14, 15] if has_fertilizer_row else [7, 9, 10, 11, 12, 13, 14, 15]
    extra = max(0, items_count - len(base_rows))
    if extra:
        ws.insert_rows(16, amount=extra)
        for offset in range(extra):
            target_row = 16 + offset
            copy_row_style(ws, 15, target_row, max_col=17)
        last_detail_row = 15 + extra
        extend_cash_detail_merges(ws, last_detail_row)
    else:
        last_detail_row = 15
    return base_rows + list(range(16, 16 + extra)), last_detail_row + 1


def attachment_cash_group(item: dict[str, str]) -> str:
    land_use = str(item.get("land_use", "")).strip()
    return "road" if "路" in land_use else "construction"


def clear_cash_dynamic_merges(ws) -> None:
    summary_ranges = {(1, 4), (5, 11), (12, 15)}
    for merged in list(ws.merged_cells.ranges):
        is_detail_merge = (
            merged.min_row == 5
            and merged.min_col == merged.max_col
            and merged.min_col in {3, 4, 5, 16}
        )
        is_summary_merge = (
            merged.min_row == merged.max_row
            and merged.min_row >= 16
            and (merged.min_col, merged.max_col) in summary_ranges
        )
        if is_detail_merge or is_summary_merge:
            ws.unmerge_cells(str(merged))


def apply_cash_dynamic_merges(ws, last_detail_row: int, total_row: int) -> None:
    for col_letter in ["C", "D", "E", "P"]:
        ws.merge_cells(f"{col_letter}5:{col_letter}{last_detail_row}")
    ws.merge_cells(f"A{total_row}:D{total_row}")
    ws.merge_cells(f"E{total_row}:K{total_row}")
    ws.merge_cells(f"L{total_row}:O{total_row}")


def copy_blank_cash_row(ws, source_row: int, target_row: int, label: str) -> None:
    copy_row_style(ws, source_row, target_row, max_col=17)
    for col in range(1, 18):
        ws.cell(target_row, col).value = None
    ws[f"B{target_row}"] = label


def prepare_cash_attachment_slots(ws, construction_count: int, road_count: int) -> tuple[list[int], list[int], int]:
    total_row_height = ws.row_dimensions[16].height
    signature_row_height = ws.row_dimensions[17].height
    clear_cash_dynamic_merges(ws)
    construction_extra = max(0, construction_count - 2)
    if construction_extra:
        ws.insert_rows(10, amount=construction_extra)
        for offset in range(construction_extra):
            copy_blank_cash_row(ws, 9, 10 + offset, "施工用地")

    road_row = 10 + construction_extra
    road_extra = max(0, road_count - 1)
    if road_extra:
        ws.insert_rows(road_row + 1, amount=road_extra)
        for offset in range(road_extra):
            copy_blank_cash_row(ws, road_row, road_row + 1 + offset, "施工道路")

    total_row = 16 + construction_extra + road_extra
    last_detail_row = total_row - 1
    apply_cash_dynamic_merges(ws, last_detail_row, total_row)
    ws.row_dimensions[total_row].height = total_row_height
    if total_row + 1 <= ws.max_row and signature_row_height is not None:
        ws.row_dimensions[total_row + 1].height = signature_row_height
    construction_rows = [7, 9] + list(range(10, 10 + construction_extra))
    road_rows = [road_row] + list(range(road_row + 1, road_row + 1 + road_extra))
    return construction_rows, road_rows, total_row


def update_cash_total_text_formula(ws, total_row: int) -> None:
    ws[f"P{total_row}"] = "=P5"
    total_text_cell = ws[f"E{total_row}"]
    if isinstance(total_text_cell.value, str):
        total_text_cell.value = re.sub(r"P\d+", f"P{total_row}", total_text_cell.value)


def filled_cash_land_cells(ws, last_detail_row: int) -> list[str]:
    cells: list[str] = []
    for row in range(5, last_detail_row + 1):
        cell = f"J{row}"
        if ws[cell].value not in (None, "", "\\", "/", "／"):
            cells.append(cell)
    return cells


def has_cell_value(ws, address: str) -> bool:
    return ws[address].value not in (None, "", "\\", "/", "／")


def repair_cash_land_formulas(ws, total_row: int, road_row: int) -> None:
    if has_cell_value(ws, "G5") and has_cell_value(ws, "I5"):
        ws["J5"] = "=G5*I5"
    if has_cell_value(ws, "G7") and has_cell_value(ws, "H7") and has_cell_value(ws, "I7"):
        ws["J7"] = "=G7*H7*I7"
    if has_cell_value(ws, f"G{road_row}") and has_cell_value(ws, f"H{road_row}") and has_cell_value(ws, f"I{road_row}"):
        ws[f"J{road_row}"] = f"=G{road_row}*H{road_row}*I{road_row}"

    for row in range(5, total_row):
        if ws[f"B{row}"].value == "培肥费":
            if has_cell_value(ws, "G7") or has_cell_value(ws, f"G{road_row}"):
                ws[f"G{row}"] = f"=G{road_row}+G7"
            if not has_cell_value(ws, f"I{row}"):
                ws[f"I{row}"] = 1500
            ws[f"J{row}"] = f"=G{row}*I{row}"
            break


def renumber_cash_rows(ws, construction_rows: list[int], road_rows: list[int], total_row: int) -> None:
    ws["A5"] = 1
    serial = 2
    ordered_rows = [7, 8, 9] + construction_rows[2:] + road_rows
    for row in ordered_rows:
        if row >= total_row:
            continue
        ws[f"A{row}"] = serial
        serial += 1
    for row in range(5, total_row):
        if ws[f"B{row}"].value == "培肥费":
            ws[f"A{row}"] = serial
            break


def rebuild_attachment_sheet(ws, data: dict[str, str]) -> None:
    items = parse_attachment_items(data)
    if not items:
        return

    clear_attachment_sheet_merges(ws)
    extra = max(0, len(items) - 2)
    if extra:
        ws.insert_rows(9, amount=extra * 2)
        for index in range(extra):
            target_top = 9 + index * 2
            copy_row_style(ws, 7, target_top)
            copy_row_style(ws, 8, target_top + 1)

    block_slots = max(2, len(items))
    last_item_bottom = 4 + block_slots * 2
    summary_top = 9 + extra * 2
    summary_bottom = summary_top + 1
    note_row = summary_bottom + 1

    ws.merge_cells(f"B5:B{last_item_bottom}")
    ws.merge_cells(f"I5:I{summary_top}")
    ws.merge_cells(f"J5:K{summary_top}")

    for index in range(block_slots):
        top = 5 + index * 2
        for col in ["A", "C", "D", "E", "F", "G", "H"]:
            ws.merge_cells(f"{col}{top}:{col}{top + 1}")

    ws.merge_cells(f"C{summary_top}:G{summary_top}")
    ws.merge_cells(f"A{summary_bottom}:B{summary_bottom}")
    ws.merge_cells(f"C{summary_bottom}:G{summary_bottom}")
    ws.merge_cells(f"H{summary_bottom}:L{summary_bottom}")
    ws.merge_cells(f"A{note_row}:L{note_row}")
    ws.merge_cells(f"L5:L{summary_top}")

    total = 0.0
    for index, item in enumerate(items, start=1):
        top = 5 + (index - 1) * 2
        amount = attachment_item_amount(item)
        if amount is not None:
            total += amount
        ws[f"A{top}"] = index
        ws[f"C{top}"] = item.get("name") or item.get("type") or ""
        ws[f"D{top}"] = item.get("spec", "")
        ws[f"E{top}"] = excel_number_value(item.get("area"))
        ws[f"F{top}"] = excel_number_value(item.get("unit"))
        ws[f"G{top}"] = excel_number_value(amount)
        ws[f"H{top}"] = data.get("tower_ids", "").strip()
        if index == 1:
            ws["B5"] = data.get("name", "").strip()
            ws["I5"] = data.get("bank_name", "").strip()
            ws["J5"] = data.get("bank_card", "").strip()
    if items and ws["H7"].value is not None:
        ws["H5"].font = copy(ws["H7"].font)
        ws["H5"].alignment = copy(ws["H7"].alignment)

    ws[f"B{summary_top}"] = "合计"
    ws[f"C{summary_top}"] = excel_number_value(total)
    ws[f"A{summary_bottom}"] = "总金额:"
    ws[f"C{summary_bottom}"] = excel_number_value(total)
    ws[f"H{summary_bottom}"] = agreement_to_table.amount_upper(total)


def write_cash_attachment_rows(ws, data: dict[str, str]) -> None:
    items = parse_attachment_items(data)

    construction_items = [item for item in items if attachment_cash_group(item) == "construction"]
    road_items = [item for item in items if attachment_cash_group(item) == "road"]
    construction_rows, road_rows, total_row = prepare_cash_attachment_slots(ws, len(construction_items), len(road_items))
    last_detail_row = total_row - 1
    repair_cash_land_formulas(ws, total_row, road_rows[0])

    for row in range(7, last_detail_row + 1):
        for col in ["K", "L", "M", "N", "O"]:
            ws[f"{col}{row}"] = None

    amount_cells: list[str] = []
    if not items:
        land_cells = filled_cash_land_cells(ws, last_detail_row)
        if land_cells:
            ws["P5"] = f"=SUM({','.join(land_cells)})"
            update_cash_total_text_formula(ws, total_row)
        return

    for item, row in [*zip(construction_items, construction_rows), *zip(road_items, road_rows)]:
        amount = attachment_item_amount(item)
        ws[f"K{row}"] = item.get("name", "") or item.get("type", "")
        ws[f"L{row}"] = item.get("spec", "")
        ws[f"M{row}"] = excel_number_value(item.get("area"))
        ws[f"N{row}"] = excel_number_value(item.get("unit"))
        ws[f"O{row}"] = excel_number_value(amount)
        amount_cells.append(f"O{row}")

    land_cells = filled_cash_land_cells(ws, last_detail_row)
    sum_cells = land_cells + amount_cells
    if sum_cells:
        ws["P5"] = f"=SUM({','.join(sum_cells)})"
        update_cash_total_text_formula(ws, total_row)
    renumber_cash_rows(ws, construction_rows, road_rows, total_row)


def apply_manual_table_fields(table_path: Path, data: dict[str, str]) -> None:
    attachment_items = parse_attachment_items(data)
    primary_item = attachment_items[0] if attachment_items else {}
    attachment_name = data.get("attachment_name", "").strip() or primary_item.get("name", "").strip() or primary_item.get("type", "").strip()
    attachment_spec = data.get("attachment_spec", "").strip() or primary_item.get("spec", "").strip()
    if attachment_name in {"\\", "/", "／"}:
        attachment_name = ""
    if attachment_spec in {"\\", "/", "／"}:
        attachment_spec = ""
    wb = load_workbook(table_path, rich_text=True)
    cash = wb.worksheets[0]
    attach = wb.worksheets[2]
    write_construction_attachment = cash["M7"].value not in (None, "", "\\", "/", "／")
    write_road_attachment = cash["M10"].value not in (None, "", "\\", "/", "／")
    if not write_construction_attachment and not write_road_attachment:
        write_construction_attachment = should_write_attachment_row(data, "construction_area")
        write_road_attachment = should_write_attachment_row(data, "road_area")

    if attachment_name:
        if write_construction_attachment:
            cash["K7"] = attachment_name
        if write_road_attachment:
            cash["K10"] = attachment_name
        attach["C5"] = attachment_name
    if attachment_spec:
        if write_construction_attachment:
            cash["L7"] = attachment_spec
        if write_road_attachment:
            cash["L10"] = attachment_spec
    write_cash_attachment_rows(cash, data)
    rebuild_attachment_sheet(attach, data)

    agreement_to_table.apply_number_formats(wb)
    wb.save(table_path)


ProgressCallback = Callable[[int, str, str], None]


def build_all(
    person_folder: Path,
    form_data: dict[str, str],
    progress: ProgressCallback | None = None,
    output_root: Path | None = None,
) -> GeneratedFiles:
    def notify(value: int, message: str, detail: str = "") -> None:
        if progress is not None:
            progress(value, message, detail)

    warnings: list[str] = []
    form_data = normalize_form_data(form_data)
    no_id_doc = form_data.get("no_id_doc", "").strip() in {"1", "true", "True", "是", "yes", "on"}
    name = form_data.get("name", "").strip()
    if not name and not no_id_doc:
        name = person_folder.name
    if not name:
        raise ValueError("请先填写姓名，生成的文件夹和文件名会按这个姓名命名。")
    notify(8, "正在准备输出文件夹", "正在创建或打开青赔目录下的人名文件夹。")
    output_dir = (output_root or BASE_DIR) / safe_filename(name)
    output_dir.mkdir(parents=True, exist_ok=True)

    template_choice = form_data.get("template_choice", "without_v2")
    template_info = TEMPLATE_OPTIONS.get(template_choice, TEMPLATE_OPTIONS["without"])
    with_fertilizer = bool(template_info["with_fertilizer"])
    template_path = template_info["agreement"]

    notify(18, "正在整理证件 Word", "正在识别证件图片并排入身份证、银行卡文档。")
    cert_doc = generate_cert_doc(person_folder, name, output_dir, no_id_doc=no_id_doc)
    if cert_doc is None:
        warnings.append("没有找到可用证件图片，已跳过证件 Word。")
        notify(32, "证件 Word 已跳过", "没有找到可用证件图片，继续生成协议书。")
    else:
        notify(32, "证件 Word 已完成", f"已保存：{cert_doc.name}")

    notify(42, "正在填写协议书", "正在套用所选模板，并保持正文、签字盖章格式。")
    agreement_path = unique_path(output_dir / f"{safe_filename(name)}-铁塔租地补偿协议书.docx")
    build_agreement(template_path, agreement_path, form_data, with_fertilizer)
    notify(62, "协议书已完成", f"已保存：{agreement_path.name}")

    notify(72, "正在生成补偿表", "正在写入兑现表、土地结算表和附作物结算表。")
    table_path = unique_path(output_dir / f"某风电项目-{safe_filename(name)}补偿表.xlsx")
    try:
        agreement_to_table.generate_table(agreement_path, table_path, str(template_info["table"]))
        apply_manual_table_fields(table_path, form_data)
        notify(90, "补偿表已完成", f"已保存：{table_path.name}")
    except Exception as exc:
        table_path = None
        warnings.append(f"补偿表生成失败：{exc}")
        notify(90, "补偿表生成失败", "协议书已生成，补偿表失败信息会在完成提示里列出。")

    notify(98, "正在完成归档", "正在检查生成结果并准备完成提示。")
    return GeneratedFiles(cert_doc=cert_doc, agreement_doc=agreement_path, table_xlsx=table_path, output_dir=output_dir, warnings=warnings)


class ScrollFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg="#f4f6f8")
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="App.TFrame")
        self.inner.bind("<Configure>", lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vbar.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self._resize_inner)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _resize_inner(self, event) -> None:
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event) -> None:
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass


class App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("青赔系统")
        self.root.geometry("1020x740")
        self.root.minsize(860, 620)
        self.configure_styles()
        self.brand_logo = None
        self.load_brand_logo()

        self.selected_folder: Path | None = None
        self.extracted_data: Any = None
        self.manual_no_cert_doc = False
        self.table_input_mode = False
        self.vars: dict[str, tk.StringVar] = {}
        self.required_labels: dict[str, str] = {}
        self.missing_var: tk.StringVar | None = None
        self._syncing_seedling_total = False
        self.loading_active = False
        self.loading_phase = 0
        self.loading_progress = 0
        self.generation_wait_active = False
        self.database_records: dict[str, Any] = {}

        qingpei_db.initialize()

        self.main = ttk.Frame(self.root, padding=18, style="App.TFrame")
        self.main.pack(fill="both", expand=True)
        self.show_start()

    def configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.bg_color = "#f4f6f8"
        self.card_color = "#ffffff"
        self.panel_color = "#ffffff"
        self.panel_line = "#d6b459"
        self.accent_color = "#bd8719"
        self.accent_dark = "#9b6710"
        self.text_color = "#1d2935"
        self.muted_color = "#667583"
        self.root.configure(bg=self.bg_color)
        default_font = ("Microsoft YaHei UI", 10)
        style.configure("App.TFrame", background=self.bg_color)
        style.configure("Card.TFrame", background=self.card_color)
        style.configure("Toolbar.TFrame", background=self.bg_color)
        style.configure("Section.TLabelframe", background=self.card_color, bordercolor="#d7dce2", relief="solid")
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"), foreground=self.accent_color, background=self.bg_color)
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 18, "bold"), foreground=self.text_color, background=self.bg_color)
        style.configure("SubHeader.TLabel", font=default_font, foreground=self.muted_color, background=self.bg_color)
        style.configure("Hint.TLabel", font=default_font, foreground=self.muted_color, background=self.card_color)
        style.configure("Status.TLabel", font=default_font, foreground="#6f5520", background="#fff8e8")
        style.configure("TLabel", font=default_font, background=self.card_color, foreground=self.text_color)
        style.configure("TEntry", padding=5, fieldbackground="#ffffff", foreground=self.text_color, insertcolor=self.accent_color)
        style.configure("TCombobox", padding=5, fieldbackground="#ffffff", foreground=self.text_color, background="#ffffff", arrowcolor=self.accent_color)
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")], foreground=[("readonly", self.text_color)], selectbackground=[("readonly", "#f5e1aa")], selectforeground=[("readonly", self.text_color)])
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(14, 7), background="#ffffff", foreground=self.text_color, bordercolor="#d7dce2")
        style.map("TButton", background=[("active", "#fff5dc")], foreground=[("active", "#5f4105")])
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(18, 8), background=self.accent_color, foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#d8a638"), ("disabled", "#d9bf84")], foreground=[("disabled", "#765a20")])
        style.configure("TRadiobutton", font=default_font, background=self.card_color, foreground=self.text_color)
        style.configure("TCheckbutton", font=default_font, background=self.card_color, foreground=self.text_color)
        style.configure("Tech.Horizontal.TProgressbar", troughcolor="#ebeef1", background=self.accent_color, bordercolor="#d6b459", lightcolor="#e9c863", darkcolor=self.accent_dark)
        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground=self.text_color, rowheight=30, bordercolor="#d7dce2")
        style.configure("Treeview.Heading", background="#f1f3f5", foreground="#715215", font=("Microsoft YaHei UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#f4deb0")], foreground=[("selected", "#3f2b06")])

    def load_brand_logo(self) -> None:
        if Image is None or ImageTk is None:
            return
        logo_path = BASE_DIR / "assets" / "entry-logo.jpg"
        if not logo_path.exists():
            return
        try:
            image = Image.open(logo_path).convert("RGB")
            image.thumbnail((72, 72), Image.Resampling.LANCZOS)
            self.brand_logo = ImageTk.PhotoImage(image)
        except Exception:
            self.brand_logo = None

    def clear(self) -> None:
        self.missing_var = None
        self.loading_canvas = None
        for child in self.main.winfo_children():
            child.destroy()

    def add_header(self, title: str, subtitle: str = "") -> None:
        banner = tk.Frame(self.main, bg=self.panel_color, padx=18, pady=12, highlightthickness=1, highlightbackground="#6d5620")
        banner.pack(fill="x", pady=(0, 14))
        title_row = tk.Frame(banner, bg=self.panel_color)
        title_row.pack(fill="x")
        if self.brand_logo is not None:
            tk.Label(title_row, image=self.brand_logo, bg=self.panel_color).pack(side="left", padx=(0, 12))
        tk.Label(
            title_row,
            text=title,
            font=("Microsoft YaHei UI", 20, "bold"),
            fg=self.text_color,
            bg=self.panel_color,
            anchor="w",
        ).pack(side="left", anchor="w")
        line = tk.Frame(banner, bg=self.panel_line, height=2)
        line.pack(fill="x", pady=(8, 0))
        tk.Frame(line, bg=self.accent_color, width=180, height=2).pack(side="left")
        if subtitle:
            tk.Label(
                banner,
                text=subtitle,
                font=("Microsoft YaHei UI", 10),
                fg=self.muted_color,
                bg=self.panel_color,
                anchor="w",
                justify="left",
                wraplength=900,
            ).pack(anchor="w", pady=(6, 0))
    def add_author_footer(self) -> None:
        footer = tk.Frame(self.main, bg=self.bg_color)
        footer.pack(fill="x", pady=(12, 0))
        tk.Frame(footer, bg="#d6b459", height=1).pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Label(
            footer,
            text="小呆鸡",
            font=("Microsoft YaHei UI", 9),
            fg="#8a7750",
            bg=self.bg_color,
        ).pack(side="right")

    def add_step_chips(self, parent, steps: list[str], active: int = 0) -> None:
        row = tk.Frame(parent, bg=self.bg_color)
        row.pack(fill="x", pady=(0, 14))
        for index, text in enumerate(steps):
            bg = self.accent_color if index == active else "#ffffff"
            fg = "#ffffff" if index == active else "#687684"
            chip = tk.Label(row, text=text, bg=bg, fg=fg, font=("Microsoft YaHei UI", 9, "bold"), padx=12, pady=5)
            chip.pack(side="left", padx=(0, 8))

    def start_import_existing_data(self) -> None:
        if not messagebox.askyesno(
            "扫描旧资料",
            "系统会扫描青赔目录下现有的人名文件夹，把 Word、Excel、图片和 PDF 的位置登记到资料库。\n\n不会移动、删除或修改原文件，是否开始？",
        ):
            return

        def worker() -> None:
            try:
                imported, files_count = qingpei_db.scan_existing_folders()
                self.root.after(
                    0,
                    lambda: (
                        messagebox.showinfo("扫描完成", f"已新增登记 {imported} 户，已关联 {files_count} 个资料文件。"),
                        self.show_database(),
                    ),
                )
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: messagebox.showerror("扫描失败", error))

        threading.Thread(target=worker, daemon=True).start()

    def show_database(self) -> None:
        self.clear()
        self.add_header("资料库")

        search_bar = ttk.Frame(self.main, style="Toolbar.TFrame")
        search_bar.pack(fill="x", pady=(0, 10))
        ttk.Label(search_bar, text="查找").pack(side="left", padx=(0, 8))
        self.database_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_bar, textvariable=self.database_search_var)
        search_entry.pack(side="left", fill="x", expand=True)
        search_entry.bind("<Return>", lambda _event: self.refresh_database_records())
        ttk.Button(search_bar, text="查询", command=self.refresh_database_records).pack(side="left", padx=(8, 0))
        ttk.Button(search_bar, text="扫描本地资料", command=self.start_import_existing_data, style="Primary.TButton").pack(side="left", padx=(8, 0))

        filter_bar = ttk.Frame(self.main, style="Toolbar.TFrame")
        filter_bar.pack(fill="x", pady=(0, 10))
        view_value = qingpei_db.get_preference("database_view", "按人名")
        payment_value = qingpei_db.get_preference("database_payment_filter", "全部")
        archive_value = qingpei_db.get_preference("database_archive_filter", "全部")
        self.database_view_var = tk.StringVar(value=view_value if view_value in {"按人名", "按桩号", "按地名"} else "按人名")
        self.database_payment_var = tk.StringVar(value=payment_value if payment_value in {"全部", "已付款", "未付款"} else "全部")
        self.database_archive_var = tk.StringVar(value=archive_value if archive_value in {"全部", "已归档", "未归档"} else "全部")
        for label, variable, values in [
            ("查看", self.database_view_var, ["按人名", "按桩号", "按地名"]),
            ("付款", self.database_payment_var, ["全部", "已付款", "未付款"]),
            ("归档", self.database_archive_var, ["全部", "已归档", "未归档"]),
        ]:
            ttk.Label(filter_bar, text=label).pack(side="left", padx=(0, 5))
            selector = ttk.Combobox(filter_bar, textvariable=variable, values=values, state="readonly", width=10)
            selector.pack(side="left", padx=(0, 12))
            selector.bind("<<ComboboxSelected>>", lambda _event: self.apply_database_filters())

        table_frame = ttk.Frame(self.main, style="Card.TFrame", padding=8)
        table_frame.pack(fill="both", expand=True)
        columns = ("name", "tower", "location", "total", "payment", "archive", "files", "updated")
        self.database_tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "name": ("姓名", 95), "tower": ("桩号", 125), "location": ("地名", 150),
            "total": ("补偿金额", 105), "payment": ("付款", 75), "archive": ("归档", 75),
            "files": ("资料", 55), "updated": ("更新时间", 145),
        }
        for key, (title, width) in headings.items():
            self.database_tree.heading(key, text=title)
            self.database_tree.column(key, width=width, anchor="center", stretch=key in {"tower", "location", "updated"})
        ybar = ttk.Scrollbar(table_frame, orient="vertical", command=self.database_tree.yview)
        xbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.database_tree.xview)
        self.database_tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.database_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.database_tree.bind("<Double-1>", lambda _event: self.open_selected_database_folder())

        bottom_bar = ttk.Frame(self.main, style="Toolbar.TFrame")
        bottom_bar.pack(fill="x", pady=(10, 0))
        ttk.Button(bottom_bar, text="返回录入", command=self.show_start).pack(side="left")
        ttk.Button(bottom_bar, text="打开资料文件夹", command=self.open_selected_database_folder).pack(side="right")
        ttk.Button(bottom_bar, text="付款/归档", command=self.edit_selected_database_status).pack(side="right", padx=(0, 8))
        ttk.Button(bottom_bar, text="导出汇总", command=self.export_database_records, style="Primary.TButton").pack(side="right", padx=(0, 8))
        self.refresh_database_records()

    @staticmethod
    def payment_label(value: str) -> str:
        return "已付款" if str(value or "").strip() in {"已付款", "已支付", "是"} else "未付款"

    @staticmethod
    def archive_label(value: str) -> str:
        return "已归档" if str(value or "").strip() in {"已归档", "是"} else "未归档"

    def apply_database_filters(self) -> None:
        qingpei_db.set_preference("database_view", self.database_view_var.get())
        qingpei_db.set_preference("database_payment_filter", self.database_payment_var.get())
        qingpei_db.set_preference("database_archive_filter", self.database_archive_var.get())
        self.refresh_database_records()

    def refresh_database_records(self) -> None:
        if not hasattr(self, "database_tree"):
            return
        keyword = self.database_search_var.get() if hasattr(self, "database_search_var") else ""
        rows = qingpei_db.list_records(keyword)
        payment_filter = self.database_payment_var.get() if hasattr(self, "database_payment_var") else "全部"
        archive_filter = self.database_archive_var.get() if hasattr(self, "database_archive_var") else "全部"
        if payment_filter != "全部":
            rows = [row for row in rows if self.payment_label(row["payment_status"]) == payment_filter]
        if archive_filter != "全部":
            rows = [row for row in rows if self.archive_label(row["archive_status"]) == archive_filter]

        view = self.database_view_var.get() if hasattr(self, "database_view_var") else "按人名"
        if view == "按桩号":
            rows = sorted(rows, key=lambda row: (str(row["tower_ids"] or "未识别"), str(row["name"] or "")))
        elif view == "按地名":
            rows = sorted(rows, key=lambda row: (" ".join(str(row[key] or "") for key in ("town", "village")) or "未识别", str(row["name"] or "")))
        else:
            rows = sorted(rows, key=lambda row: str(row["name"] or ""))
        self.database_records = {str(row["id"]): row for row in rows}
        for item in self.database_tree.get_children():
            self.database_tree.delete(item)
        for row in rows:
            location = " ".join(part for part in [row["town"], row["village"]] if part)
            values = (
                row["name"], row["tower_ids"], location, f"{float(row['total_amount'] or 0):.2f}",
                self.payment_label(row["payment_status"]), self.archive_label(row["archive_status"]), row["file_count"], row["updated_at"],
            )
            self.database_tree.insert("", "end", iid=str(row["id"]), values=values)

    def selected_database_row(self):
        if not hasattr(self, "database_tree"):
            return None
        selection = self.database_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先在资料库中选择一位人员。")
            return None
        return self.database_records.get(selection[0])

    def open_selected_database_folder(self) -> None:
        row = self.selected_database_row()
        if row is None:
            return
        folder = Path(row["folder_path"])
        if folder.is_file():
            folder = folder.parent
        if not folder.exists():
            messagebox.showerror("文件夹不存在", f"资料文件夹不存在或已被移动：\n{folder}")
            return
        subprocess.run(["explorer", str(folder)], check=False)

    def edit_selected_database_status(self) -> None:
        row = self.selected_database_row()
        if row is None:
            return
        window = tk.Toplevel(self.root)
        window.title(f"资料状态 - {row['name']}")
        window.resizable(False, False)
        window.configure(bg=self.bg_color)
        panel = ttk.Frame(window, padding=18, style="App.TFrame")
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text=f"{row['name']} 的资料状态", style="Header.TLabel").pack(anchor="w", pady=(0, 14))
        payment_var = tk.StringVar(value=self.payment_label(row["payment_status"]))
        archive_var = tk.StringVar(value=self.archive_label(row["archive_status"]))
        note_var = tk.StringVar(value=row["missing_notes"])
        tower_var = tk.StringVar(value=row["tower_ids"])
        amount_var = tk.StringVar(value=f"{float(row['total_amount'] or 0):.2f}" if float(row["total_amount"] or 0) else "")
        for label, variable in [("铁塔编号", tower_var), ("补偿总额", amount_var)]:
            line = ttk.Frame(panel)
            line.pack(fill="x", pady=5)
            ttk.Label(line, text=label, width=10, anchor="e").pack(side="left", padx=(0, 8))
            ttk.Entry(line, textvariable=variable, width=32).pack(side="left", fill="x", expand=True)
        for label, variable, values in [
            ("付款", payment_var, ["未付款", "已付款"]),
            ("归档", archive_var, ["未归档", "已归档"]),
        ]:
            line = ttk.Frame(panel)
            line.pack(fill="x", pady=5)
            ttk.Label(line, text=label, width=10, anchor="e").pack(side="left", padx=(0, 8))
            ttk.Combobox(line, textvariable=variable, values=values, state="readonly", width=18).pack(side="left", fill="x", expand=True)
        line = ttk.Frame(panel)
        line.pack(fill="x", pady=5)
        ttk.Label(line, text="缺少资料备注", width=10, anchor="e").pack(side="left", padx=(0, 8))
        ttk.Entry(line, textvariable=note_var, width=32).pack(side="left", fill="x", expand=True)

        def save() -> None:
            amount_text = amount_var.get().strip().replace("，", ".").replace(",", "")
            try:
                total_amount = float(amount_text) if amount_text else 0.0
            except ValueError:
                messagebox.showerror("金额格式不对", "补偿总额请填写数字，例如：12500.50")
                return
            qingpei_db.update_compensation(int(row["id"]), tower_var.get(), total_amount)
            qingpei_db.update_status(int(row["id"]), row["sign_status"], payment_var.get(), archive_var.get(), note_var.get().strip())
            window.destroy()
            self.refresh_database_records()

        buttons = ttk.Frame(panel)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="取消", command=window.destroy).pack(side="right")
        ttk.Button(buttons, text="保存", command=save, style="Primary.TButton").pack(side="right", padx=(0, 8))

    def export_database_records(self) -> None:
        try:
            output_path = qingpei_db.export_records()
        except Exception:
            messagebox.showerror("导出失败", traceback.format_exc())
            return
        messagebox.showinfo("导出完成", f"汇总表已导出到：\n{output_path}")

    def show_start(self) -> None:
        self.clear()
        self.loading_active = False
        self.table_input_mode = False
        self.add_header("青赔系统")
        self.add_step_chips(self.main, ["选择资料", "核对信息", "生成文件"], active=0)

        bottom_bar = ttk.Frame(self.main, style="Toolbar.TFrame")
        bottom_bar.pack(side="bottom", fill="x", pady=(10, 0))

        self.status_var = tk.StringVar(value="选择资料和处理方式后，点击“下一步”开始。")
        self.status_box = tk.Frame(bottom_bar, bg="#fff8e8", padx=12, pady=9, highlightthickness=1, highlightbackground="#e2c779")
        status_box = self.status_box
        status_box.pack(fill="x", pady=(0, 8))
        self.status_label = tk.Label(
            status_box,
            textvariable=self.status_var,
            font=("Microsoft YaHei UI", 10),
            fg="#725923",
            bg="#fff8e8",
            anchor="w",
            justify="left",
            wraplength=880,
        )
        self.status_label.pack(anchor="w")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(status_box, mode="determinate", maximum=100, variable=self.progress_var, style="Tech.Horizontal.TProgressbar")

        buttons = ttk.Frame(bottom_bar, style="Toolbar.TFrame")
        buttons.pack(fill="x")
        tk.Label(
            buttons,
            text="小呆鸡",
            font=("Microsoft YaHei UI", 8),
            fg="#8a7750",
            bg=self.bg_color,
        ).pack(side="left")
        self.next_button = ttk.Button(buttons, text="下一步", command=self.start_extract, style="Primary.TButton")
        self.next_button.pack(side="right")
        ttk.Button(buttons, text="退出", command=self.root.destroy).pack(side="right", padx=(0, 8))

        scroll = ScrollFrame(self.main)
        scroll.pack(fill="both", expand=True)
        body = scroll.inner

        folder_frame = self.add_section(body, "选择人员资料")
        row = ttk.Frame(folder_frame)
        row.pack(fill="x")
        ttk.Label(row, text="当前资料", width=12, anchor="e").pack(side="left", padx=(0, 8))
        folders = person_folders()
        self.folder_map = {person_folder_label(path): path for path in folders}
        self.folder_var = tk.StringVar(value=next(iter(self.folder_map), ""))
        self.folder_combo = ttk.Combobox(row, textvariable=self.folder_var, values=list(self.folder_map), state="readonly")
        self.folder_combo.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览", command=self.browse_folder).pack(side="left", padx=(8, 0))

        mode_frame = self.add_section(body, "选择处理方式")
        self.manual_doc_var = tk.StringVar(value="auto")
        mode_grid = ttk.Frame(mode_frame)
        mode_grid.pack(fill="x")
        for column, (title, hint, value) in enumerate([
            ("自动识别", "读取身份证、银行卡或存折图片", "auto"),
            ("手动填写", "适合没有证件图片的资料", "manual"),
            ("表格识别", "从已有 Excel 中选择人员", "table"),
        ]):
            mode_grid.columnconfigure(column, weight=1)
            option = tk.Frame(mode_grid, bg="#ffffff", padx=10, pady=8, highlightthickness=1, highlightbackground="#d7dce2")
            option.grid(row=0, column=column, padx=(0, 8) if column < 2 else 0, sticky="nsew")
            ttk.Radiobutton(
                option,
                text=title,
                variable=self.manual_doc_var,
                value=value,
                command=self.update_start_mode,
            ).pack(anchor="w")
            tk.Label(option, text=hint, font=("Microsoft YaHei UI", 9), fg="#667583", bg="#ffffff", anchor="w").pack(anchor="w", pady=(3, 0))

        self.table_frame = self.add_section(body, "Excel 表格识别")
        ttk.Label(self.table_frame, text="选择青赔 Excel 表和人员；缺少的证件信息会留空，下一步可手动补充。", style="Hint.TLabel", wraplength=880).pack(anchor="w", pady=(0, 8))
        table_row = ttk.Frame(self.table_frame)
        table_row.pack(fill="x", pady=3)
        ttk.Label(table_row, text="表格文件", width=12, anchor="e").pack(side="left", padx=(0, 8))
        self.table_file_map = {path.name: path for path in table_files()}
        self.table_file_var = tk.StringVar(value=next(iter(self.table_file_map), ""))
        self.table_file_combo = ttk.Combobox(table_row, textvariable=self.table_file_var, values=list(self.table_file_map), state="readonly")
        self.table_file_combo.pack(side="left", fill="x", expand=True)
        self.table_file_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_table_people())
        ttk.Button(table_row, text="刷新表格", command=self.refresh_table_files).pack(side="left", padx=(8, 0))

        person_row = ttk.Frame(self.table_frame)
        person_row.pack(fill="x", pady=3)
        ttk.Label(person_row, text="表格人员", width=12, anchor="e").pack(side="left", padx=(0, 8))
        self.table_person_var = tk.StringVar(value="")
        self.table_person_combo = ttk.Combobox(person_row, textvariable=self.table_person_var, values=[], state="readonly")
        self.table_person_combo.pack(side="left", fill="x", expand=True)
        self.table_entry_map: dict[str, TablePersonEntry] = {}
        self.refresh_table_people()

        self.tools_frame = self.add_section(body, "资料中心")
        tool_row = ttk.Frame(self.tools_frame)
        tool_row.pack(fill="x")
        ttk.Button(tool_row, text="资料库", command=self.show_database, style="Primary.TButton").pack(side="left")
        self.self_check_button = ttk.Button(tool_row, text="自检", command=self.start_self_check)
        self.self_check_button.pack(side="left", padx=(8, 0))
        self.update_start_mode()

    def draw_loading_bar(self) -> None:
        if not getattr(self, "loading_canvas", None):
            return
        canvas = self.loading_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 1)
        height = max(canvas.winfo_height(), 10)
        canvas.create_rectangle(0, height // 2 - 1, width, height // 2 + 1, fill="#54441b", outline="")
        if self.loading_active:
            filled = int(width * min(max(self.loading_progress, 0), 100) / 100)
            canvas.create_rectangle(0, 1, filled, height - 1, fill=self.accent_color, outline="")
            if filled:
                canvas.create_rectangle(max(0, filled - 18), 1, filled, height - 1, fill="#f7d66f", outline="")
        else:
            canvas.create_rectangle(0, 1, 80, height - 1, fill="#8f7226", outline="")

    def start_loading_animation(self, message: str) -> None:
        self.loading_active = True
        self.loading_phase = 0
        self.loading_progress = 8
        self.progress_var.set(self.loading_progress)
        self.status_box.configure(bg="#fff8e8", highlightbackground=self.accent_color)
        self.status_label.configure(bg="#fff8e8", fg="#725923")
        self.progress.pack(fill="x", pady=(8, 0))
        self.animate_loading(message)

    def animate_loading(self, message: str) -> None:
        if not self.loading_active:
            return
        dots = "." * ((self.loading_phase % 3) + 1)
        if self.loading_progress < 88:
            self.loading_progress += 2
            self.progress_var.set(self.loading_progress)
        self.status_var.set(f"{message}{dots}  {int(self.loading_progress)}%")
        self.draw_loading_bar()
        self.loading_phase += 1
        self.root.after(300, lambda: self.animate_loading(message))

    def set_loading_progress(self, value: int, message: str | None = None) -> None:
        if not self.loading_active:
            return
        self.loading_progress = max(0, min(100, value))
        if hasattr(self, "progress_var"):
            self.progress_var.set(self.loading_progress)
        if message:
            self.status_var.set(f"{message}  {int(self.loading_progress)}%")
        self.draw_loading_bar()

    def stop_loading_animation(self, message: str = "") -> None:
        if self.loading_active:
            self.loading_progress = 100
            if hasattr(self, "progress_var"):
                self.progress_var.set(100)
            self.draw_loading_bar()
        self.loading_active = False
        if hasattr(self, "progress"):
            self.progress.pack_forget()
        if message:
            self.status_var.set(message)
        self.draw_loading_bar()

    def start_generation_wait(self, person_name: str) -> None:
        self.generation_wait_active = True
        self.generation_wait_phase = 0
        self.generation_wait_progress = 0
        self.generation_wait_message_base = "正在初始化生成任务"

        win = tk.Toplevel(self.root)
        self.generation_wait_window = win
        win.title("正在生成资料")
        win.geometry("620x380")
        win.resizable(False, False)
        win.configure(bg=self.panel_color)
        win.transient(self.root)
        win.protocol("WM_DELETE_WINDOW", lambda: None)

        try:
            x = self.root.winfo_rootx() + max((self.root.winfo_width() - 620) // 2, 0)
            y = self.root.winfo_rooty() + max((self.root.winfo_height() - 380) // 2, 0)
            win.geometry(f"620x380+{x}+{y}")
        except tk.TclError:
            pass

        outer = tk.Frame(win, bg=self.card_color, padx=22, pady=20, highlightthickness=1, highlightbackground=self.panel_line)
        outer.pack(fill="both", expand=True, padx=12, pady=12)

        top = tk.Frame(outer, bg=self.card_color)
        top.pack(fill="x")
        tk.Label(
            top,
            text="青赔系统",
            font=("Consolas", 9, "bold"),
            fg="#171207",
            bg=self.accent_color,
            padx=10,
            pady=3,
        ).pack(side="left")
        tk.Label(
            top,
            text="by 小呆鸡",
            font=("Microsoft YaHei UI", 8),
            fg=self.muted_color,
            bg=self.card_color,
        ).pack(side="right")

        tk.Label(
            outer,
            text="资料生成中",
            font=("Microsoft YaHei UI", 20, "bold"),
            fg=self.text_color,
            bg=self.card_color,
        ).pack(anchor="w", pady=(16, 4))
        tk.Label(
            outer,
            text=f"当前人员：{person_name or '未命名'}",
            font=("Microsoft YaHei UI", 10),
            fg="#7b6841",
            bg=self.card_color,
        ).pack(anchor="w")

        self.generation_wait_message_var = tk.StringVar(value="正在初始化生成任务")
        self.generation_wait_detail_var = tk.StringVar(value="证件 Word、协议书、补偿表会依次生成，请不要重复点击。")
        self.generation_wait_percent_var = tk.StringVar(value="0%")
        status_row = tk.Frame(outer, bg=self.card_color)
        status_row.pack(fill="x", pady=(22, 2))
        tk.Label(
            status_row,
            textvariable=self.generation_wait_message_var,
            font=("Microsoft YaHei UI", 12, "bold"),
            fg=self.accent_color,
            bg=self.card_color,
        ).pack(side="left", anchor="w")
        tk.Label(
            status_row,
            textvariable=self.generation_wait_percent_var,
            font=("Consolas", 18, "bold"),
            fg=self.text_color,
            bg=self.card_color,
        ).pack(side="right")
        tk.Label(
            outer,
            textvariable=self.generation_wait_detail_var,
            font=("Microsoft YaHei UI", 9),
            fg=self.muted_color,
            bg=self.card_color,
            wraplength=500,
            justify="left",
        ).pack(anchor="w")

        self.generation_wait_progress_var = tk.DoubleVar(value=self.generation_wait_progress)
        ttk.Progressbar(
            outer,
            mode="determinate",
            maximum=100,
            variable=self.generation_wait_progress_var,
            style="Tech.Horizontal.TProgressbar",
        ).pack(fill="x", pady=(20, 14))

        line = tk.Frame(outer, bg="#e8ddc3", height=1)
        line.pack(fill="x", pady=(0, 14))
        tk.Frame(line, bg=self.accent_color, width=190, height=1).pack(side="left")

        steps = tk.Frame(outer, bg=self.card_color)
        steps.pack(fill="x")
        self.generation_step_labels = []
        for text in ["01 证件 Word", "02 协议书", "03 补偿表", "04 保存完成"]:
            label = tk.Label(
                steps,
                text=text,
                font=("Microsoft YaHei UI", 9, "bold"),
                fg="#687684",
                bg="#f1f3f5",
                padx=13,
                pady=6,
            )
            label.pack(side="left", padx=(0, 8))
            self.generation_step_labels.append(label)

        tk.Label(
            outer,
            text="生成完成后会自动弹出保存结果。",
            font=("Microsoft YaHei UI", 9),
            fg="#8a7750",
            bg=self.card_color,
        ).pack(anchor="w", pady=(18, 0))

        self.set_generation_wait_progress(3, "正在初始化生成任务", "正在准备后台生成流程。", 0)
        self.animate_generation_wait()

    def animate_generation_wait(self) -> None:
        if not self.generation_wait_active:
            return
        self.generation_wait_phase += 1
        message = getattr(self, "generation_wait_message_base", "正在生成资料")
        dots = "." * ((self.generation_wait_phase % 3) + 1)
        if hasattr(self, "generation_wait_message_var"):
            self.generation_wait_message_var.set(f"{message}{dots}")

        self.root.after(360, self.animate_generation_wait)

    def set_generation_wait_progress(self, value: int | float, message: str = "", detail: str = "", step_index: int | None = None) -> None:
        value = max(0, min(100, float(value)))
        current = float(getattr(self, "generation_wait_progress", 0))
        if value < current and current < 100:
            value = current
        self.generation_wait_progress = value
        if message:
            self.generation_wait_message_base = message
            if hasattr(self, "generation_wait_message_var"):
                self.generation_wait_message_var.set(message)
        if detail and hasattr(self, "generation_wait_detail_var"):
            self.generation_wait_detail_var.set(detail)
        if hasattr(self, "generation_wait_progress_var"):
            self.generation_wait_progress_var.set(value)
        if hasattr(self, "generation_wait_percent_var"):
            self.generation_wait_percent_var.set(f"{int(round(value))}%")

        if step_index is None:
            if value < 38:
                step_index = 0
            elif value < 68:
                step_index = 1
            elif value < 94:
                step_index = 2
            else:
                step_index = 3

        for index, label in enumerate(getattr(self, "generation_step_labels", [])):
            if index < step_index:
                label.configure(bg="#f1e1b7", fg="#6b4c10")
            elif index == step_index:
                label.configure(bg=self.accent_color, fg="#ffffff")
            else:
                label.configure(bg="#f1f3f5", fg="#687684")

    def stop_generation_wait(self, final_message: str = "") -> None:
        if final_message:
            if "失败" in final_message:
                self.set_generation_wait_progress(
                    getattr(self, "generation_wait_progress", 0),
                    final_message,
                    "生成过程遇到问题，正在打开错误提示。",
                )
            else:
                self.set_generation_wait_progress(100, final_message, "所有文件已保存完成。", 3)
        self.generation_wait_active = False
        win = getattr(self, "generation_wait_window", None)
        if win is not None:
            try:
                win.update_idletasks()
                win.destroy()
            except tk.TclError:
                pass
        self.generation_wait_window = None

    def refresh_folder_choices(self, preferred: Path | None = None) -> None:
        folders = person_folders()
        self.folder_map = {person_folder_label(path): path for path in folders}
        labels = list(self.folder_map)
        if hasattr(self, "folder_combo"):
            self.folder_combo.configure(values=labels)
        preferred_label = person_folder_label(preferred) if preferred else ""
        current = self.folder_var.get() if hasattr(self, "folder_var") else ""
        if preferred_label and preferred_label in self.folder_map:
            self.folder_var.set(preferred_label)
        elif current in self.folder_map:
            self.folder_var.set(current)
        else:
            self.folder_var.set(labels[0] if labels else "")

    def refresh_table_files(self) -> None:
        self.table_file_map = {path.name: path for path in table_files()}
        labels = list(self.table_file_map)
        if hasattr(self, "table_file_combo"):
            self.table_file_combo.configure(values=labels)
        current = self.table_file_var.get() if hasattr(self, "table_file_var") else ""
        self.table_file_var.set(current if current in self.table_file_map else (labels[0] if labels else ""))
        self.refresh_table_people()

    def refresh_table_people(self) -> None:
        self.table_entry_map = {}
        table_path = self.table_file_map.get(self.table_file_var.get()) if hasattr(self, "table_file_map") else None
        if table_path and table_path.exists():
            try:
                for entry in parse_table_people(table_path):
                    label = entry.label
                    if label in self.table_entry_map:
                        label = f"{label}（{len(self.table_entry_map) + 1}）"
                    self.table_entry_map[label] = entry
            except Exception:
                self.table_entry_map = {}
        labels = list(self.table_entry_map)
        if hasattr(self, "table_person_combo"):
            self.table_person_combo.configure(values=labels)
        current = self.table_person_var.get() if hasattr(self, "table_person_var") else ""
        if hasattr(self, "table_person_var"):
            self.table_person_var.set(current if current in self.table_entry_map else (labels[0] if labels else ""))

    def update_start_mode(self) -> None:
        if not hasattr(self, "table_frame") or not hasattr(self, "manual_doc_var"):
            return
        if self.manual_doc_var.get() == "table":
            self.table_frame.pack(fill="x", pady=(0, 12), before=self.tools_frame)
        else:
            self.table_frame.pack_forget()

    def start_organize_master_doc(self) -> None:
        if not CERT_DIR.exists():
            messagebox.showerror("缺少文件夹", f"没有找到证件提取文件夹：\n{CERT_DIR}")
            return
        source = CERT_DIR / master_doc_tools.MASTER_DOC_NAME
        source_docx = CERT_DIR / master_doc_tools.MASTER_DOCX_NAME
        if not source.exists() and not source_docx.exists():
            messagebox.showerror("缺少总文档", f"没有找到总证件文档：\n{source}")
            return
        self.next_button.configure(state="disabled")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="disabled")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="disabled")
        self.start_loading_animation("正在整理总证件文档")

        def worker() -> None:
            try:
                def progress(value: int, message: str) -> None:
                    self.root.after(0, lambda value=value, message=message: self.set_loading_progress(value, message))

                result = master_doc_tools.organize_master_document(CERT_DIR, progress=progress)
                self.root.after(0, lambda: self.organize_master_done(result))
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: self.organize_master_failed(error))

        threading.Thread(target=worker, daemon=True).start()

    def start_self_check(self) -> None:
        self.next_button.configure(state="disabled")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="disabled")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="disabled")
        self.start_loading_animation("正在自检软件功能")

        def worker() -> None:
            try:
                def progress(value: int, message: str) -> None:
                    self.root.after(0, lambda value=value, message=message: self.set_loading_progress(value, message))

                result = run_self_check(progress=progress)
                self.root.after(0, lambda: self.self_check_done(result))
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: self.self_check_failed(error))

        threading.Thread(target=worker, daemon=True).start()

    def self_check_done(self, result: list[str]) -> None:
        self.stop_loading_animation("自检完成，功能正常。")
        self.next_button.configure(state="normal")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="normal")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="normal")
        self.refresh_folder_choices()
        messagebox.showinfo("自检完成", "检查通过：\n" + "\n".join(f"- {line}" for line in result))

    def self_check_failed(self, error: str) -> None:
        self.stop_loading_animation("自检失败，可以查看提示后处理。")
        self.next_button.configure(state="normal")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="normal")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="normal")
        messagebox.showerror("自检失败", error)

    def organize_master_done(self, result) -> None:
        self.stop_loading_animation(f"总文档整理完成：{result.people_count} 人，{result.image_count} 张图片。")
        self.next_button.configure(state="normal")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="normal")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="normal")
        self.refresh_folder_choices()
        messagebox.showinfo(
            "整理完成",
            f"已整理 {result.people_count} 人，{result.image_count} 张图片。\n现在可以在下拉框选择带“（总文档）”的人名。",
        )

    def organize_master_failed(self, error: str) -> None:
        self.stop_loading_animation("总文档整理失败，可以查看提示后重试。")
        self.next_button.configure(state="normal")
        if hasattr(self, "organize_button"):
            self.organize_button.configure(state="normal")
        if hasattr(self, "self_check_button"):
            self.self_check_button.configure(state="normal")
        messagebox.showerror("整理失败", error)

    def browse_folder(self) -> None:
        selected = filedialog.askdirectory(initialdir=str(CERT_DIR if CERT_DIR.exists() else BASE_DIR), title="选择人名文件夹")
        if not selected:
            return
        path = Path(selected)
        label = person_folder_label(path)
        self.folder_map[label] = path
        self.folder_combo.configure(values=list(self.folder_map))
        self.folder_var.set(label)

    def start_extract(self) -> None:
        mode = self.manual_doc_var.get()
        self.table_input_mode = mode == "table"
        if self.table_input_mode:
            entry = self.table_entry_map.get(self.table_person_var.get()) if hasattr(self, "table_entry_map") else None
            if entry is None:
                messagebox.showerror("提示", "请先选择表格文件和表格人员。")
                return
            data = deepcopy(entry.data)
            self.selected_folder = CERT_DIR
            self.manual_no_cert_doc = True
            self.show_form(data)
            if "tower_ids" in self.vars and entry.tower_ids:
                self.vars["tower_ids"].set(entry.tower_ids)
            if entry.attachment_items:
                self.set_attachment_items(entry.attachment_items)
                self.update_seedling_total()
                self.seedling_status_var.set("已从来源表的附作物结算表带入青苗信息和青苗合计。")
            elif entry.seedling_total:
                self.vars["seedling_total"].set(entry.seedling_total)
                self.seedling_status_var.set("已从来源表的补偿费用 青苗 列带入青苗总额，可直接预算模拟或核对修改。")
            return

        folder = self.folder_map.get(self.folder_var.get())
        if folder is None or not folder.is_dir():
            messagebox.showerror("提示", "请先选择一个有效的人名文件夹。")
            return
        self.selected_folder = folder
        self.manual_no_cert_doc = mode == "manual"
        if self.manual_no_cert_doc:
            self.show_form(fill_template.PersonData())
            return
        self.next_button.configure(state="disabled")
        self.start_loading_animation("正在准备 OCR 识别组件")

        def worker() -> None:
            try:
                def progress(value: int, message: str) -> None:
                    self.root.after(0, lambda value=value, message=message: self.set_loading_progress(value, message))

                if not ensure_ocr_ready(progress=progress):
                    raise RuntimeError("OCR 组件没有正常加载，自动识别不可用。")
                self.root.after(0, lambda: self.set_loading_progress(42, "正在读取证件图片"))
                data = extract_data(folder)
                self.root.after(0, lambda: self.set_loading_progress(92, "识别完成，正在整理资料"))
                self.root.after(0, lambda: self.show_form(data))
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: self.extract_failed(error))

        threading.Thread(target=worker, daemon=True).start()
        return
        if not ocr_available():
            messagebox.showerror(
                "缺少 OCR 识别组件",
                "当前选择的是自动识别模式，但 OCR 组件没有正常加载。\n\n"
                "请先安装或修复 rapidocr-onnxruntime、onnxruntime、opencv-python，"
                "然后重新打开软件再试。",
            )
            return
        self.next_button.configure(state="disabled")
        self.start_loading_animation("正在扫描证件图片并识别资料")

        def worker() -> None:
            try:
                self.root.after(0, lambda: self.set_loading_progress(20, "正在读取证件图片"))
                data = extract_data(folder)
                self.root.after(0, lambda: self.set_loading_progress(92, "识别完成，正在整理资料"))
                self.root.after(0, lambda: self.show_form(data))
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: self.extract_failed(error))

        threading.Thread(target=worker, daemon=True).start()

    def extract_failed(self, error: str) -> None:
        self.next_button.configure(state="normal")
        self.stop_loading_animation("识别失败，可以查看提示后重试。")
        messagebox.showerror("识别失败", error)

    def make_var(self, key: str, value: str = "") -> tk.StringVar:
        var = tk.StringVar(value=value or "")
        var.trace_add("write", lambda *_: self.form_value_changed(key))
        self.vars[key] = var
        return var

    def form_value_changed(self, key: str) -> None:
        self.update_missing()
        if key.startswith(("attachment_name", "attachment_spec", "attachment_area", "attachment_unit", "attachment_land_use")):
            self.update_seedling_total()

    def update_seedling_total(self) -> None:
        if self._syncing_seedling_total or "seedling_total" not in self.vars:
            return
        items = parse_attachment_items(self.collect_form_data())
        total = attachment_total_amount(items)
        if total is None:
            return
        self._syncing_seedling_total = True
        try:
            self.vars["seedling_total"].set(amount_text(total))
        finally:
            self._syncing_seedling_total = False

    def set_attachment_items(self, items: list[dict[str, str]]) -> None:
        self._syncing_seedling_total = True
        try:
            for key, var in self.vars.items():
                if re.match(r"attachment_(?:land_use|name|spec|area|unit)_\d+$", key):
                    var.set("")
            for frame in getattr(self, "attachment_extra_rows", []):
                frame.destroy()
            self.attachment_extra_rows = []
            self.attachment_row_seq = 2

            first = items[0]
            for field in ["land_use", "name", "spec", "area", "unit"]:
                self.vars[f"attachment_{field}"].set(first.get(field, ""))
            for item in items[1:]:
                self.add_attachment_row(item)
        finally:
            self._syncing_seedling_total = False

    def simulate_seedlings(self) -> None:
        try:
            total_units = parse_seedling_budget(self.vars["seedling_total"].get())
        except ValueError as exc:
            messagebox.showwarning("无法模拟", str(exc), parent=self.root)
            return
        count_text = self.vars["seedling_count"].get()
        requested_count = None
        if count_text and not count_text.startswith("自动"):
            match = re.match(r"(\d+)", count_text)
            if match:
                requested_count = int(match.group(1))
        try:
            items = simulate_seedling_items(total_units, requested_count)
        except ValueError as exc:
            messagebox.showwarning("无法模拟", str(exc), parent=self.root)
            return
        if items is None:
            messagebox.showwarning(
                "无法模拟",
                f"该青苗总额不能按 {requested_count or '所选'} 种类精确组合。请调整总额或种类数后重试。",
                parent=self.root,
            )
            return
        existing_items = parse_attachment_items(self.collect_form_data())
        if existing_items and not messagebox.askyesno(
            "覆盖现有附作物",
            "模拟会替换当前地上附作物的内容，是否继续？",
            parent=self.root,
        ):
            return
        self.set_attachment_items(items)
        self._syncing_seedling_total = True
        try:
            self.vars["seedling_total"].set(decimal_text(Decimal(total_units) / 1000))
        finally:
            self._syncing_seedling_total = False
        self.seedling_status_var.set(f"已随机填入 {len(items)} 种青苗，合计金额已严格匹配。")
        self.update_missing()

    def add_entry(self, parent, key: str, label: str, value: str = "", required: bool = True) -> ttk.Frame:
        if required:
            self.required_labels[key] = label
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=5)
        mark = " *" if required else ""
        ttk.Label(row, text=label + mark, width=16, anchor="e").pack(side="left", padx=(0, 8))
        ttk.Entry(row, textvariable=self.make_var(key, value)).pack(side="left", fill="x", expand=True)
        return row

    def add_check(self, parent, key: str, label: str, value: bool = False) -> ttk.Frame:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        variable = self.make_var(key, "1" if value else "")
        ttk.Checkbutton(row, text=label, variable=variable, onvalue="1", offvalue="", command=self.update_missing).pack(anchor="w", padx=(112, 0))
        return row

    def add_attachment_row(self, values: dict[str, str] | None = None) -> None:
        if not hasattr(self, "attachment_rows_frame"):
            return
        values = values or {}
        index = getattr(self, "attachment_row_seq", 2)
        self.attachment_row_seq = index + 1
        suffix = f"_{index}"

        row = ttk.Frame(self.attachment_rows_frame)
        row.pack(fill="x", pady=3)
        row.columnconfigure(1, weight=1)
        row.columnconfigure(2, weight=1)
        row.columnconfigure(3, weight=1)
        row.columnconfigure(4, weight=1)
        row.columnconfigure(5, weight=1)

        ttk.Label(row, text=f"{index}.", width=4, anchor="e").grid(row=0, column=0, padx=(0, 8))
        land_use_var = self.make_var(f"attachment_land_use{suffix}", values.get("land_use", "施工用地") or "施工用地")
        ttk.Combobox(
            row,
            textvariable=land_use_var,
            values=["施工用地", "道路用地"],
            state="readonly",
            width=10,
        ).grid(row=0, column=1, padx=(0, 8), sticky="ew")
        name_var = self.make_var(f"attachment_name{suffix}", values.get("name", ""))
        spec_var = self.make_var(f"attachment_spec{suffix}", values.get("spec", ""))
        name_box = ttk.Combobox(
            row,
            textvariable=name_var,
            values=[name for name, _ in COMMON_ATTACHMENT_OPTIONS],
            state="normal",
            width=14,
        )
        name_box.grid(row=0, column=2, padx=(0, 8), sticky="ew")
        name_box.bind("<<ComboboxSelected>>", lambda _event: self.set_attachment_default_spec(name_var, spec_var))
        ttk.Entry(row, textvariable=spec_var, width=14).grid(row=0, column=3, padx=(0, 8), sticky="ew")
        ttk.Entry(row, textvariable=self.make_var(f"attachment_area{suffix}", values.get("area", "")), width=10).grid(
            row=0, column=4, padx=(0, 8), sticky="ew"
        )
        ttk.Entry(row, textvariable=self.make_var(f"attachment_unit{suffix}", values.get("unit", "")), width=10).grid(
            row=0, column=5, padx=(0, 8), sticky="ew"
        )
        fields = [("name", "attachment_name", 14), ("spec", "attachment_spec", 14), ("area", "attachment_area", 10), ("unit", "attachment_unit", 10)]
        ttk.Button(row, text="×", width=3, command=lambda frame=row, vars=[
            self.vars[f"attachment_land_use{suffix}"],
            self.vars[f"attachment_name{suffix}"],
            self.vars[f"attachment_spec{suffix}"],
            self.vars[f"attachment_area{suffix}"],
            self.vars[f"attachment_unit{suffix}"],
        ]: self.remove_attachment_row(frame, vars)).grid(row=0, column=len(fields) + 2)

        if not hasattr(self, "attachment_extra_rows"):
            self.attachment_extra_rows = []
        self.attachment_extra_rows.append(row)

    def remove_attachment_row(self, frame, variables: list[tk.StringVar]) -> None:
        for var in variables:
            var.set("")
        if hasattr(self, "attachment_extra_rows") and frame in self.attachment_extra_rows:
            self.attachment_extra_rows.remove(frame)
        frame.destroy()

    def set_attachment_default_spec(self, name_var: tk.StringVar, spec_var: tk.StringVar) -> None:
        selected_name = name_var.get().strip()
        for name, default_spec in COMMON_ATTACHMENT_OPTIONS:
            if selected_name == name:
                spec_var.set(default_spec)
                return

    def add_section(self, parent, title: str) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text=title, padding=12, style="Section.TLabelframe")
        frame.pack(fill="x", pady=(8, 8), padx=2)
        return frame

    def show_form(self, data) -> None:
        self.stop_loading_animation()
        self.extracted_data = data
        self.clear()
        self.vars = {}
        self.required_labels = {}
        self.make_var("no_id_doc", "1" if self.manual_no_cert_doc else "")

        top = ttk.Frame(self.main)
        top.pack(fill="x", pady=(0, 10))
        if self.table_input_mode:
            current_text = "当前模式：按表格识别录入，缺少信息可手动补充"
        elif self.manual_no_cert_doc:
            current_text = "当前模式：无身份证和银行卡文档，手动填入"
        else:
            current_text = f"当前人员：{self.selected_folder.name if self.selected_folder else ''}"
        ttk.Label(top, text=current_text, style="Header.TLabel").pack(side="left")
        ttk.Button(top, text="重新选择", command=self.show_start).pack(side="right")
        self.add_step_chips(self.main, ["选择资料", "核对信息", "生成文件"], active=1)

        if self.table_input_mode:
            intro_text = "已从表格带入已有信息，请补充面积、单价、证件资料等缺少内容。带 * 的项目建议填写。"
        elif self.manual_no_cert_doc:
            intro_text = "请填写姓名、身份证号、银行卡号等资料。带 * 的项目建议填写，缺少时会在下方提示。"
        else:
            intro_text = "已完成识别，请核对识别结果；识别不准的地方可以直接修改。带 * 的项目建议填写。"
        ttk.Label(self.main, text=intro_text, style="SubHeader.TLabel", wraplength=900).pack(anchor="w", pady=(0, 10))

        scroll = ScrollFrame(self.main)
        scroll.pack(fill="both", expand=True)
        body = scroll.inner

        template_frame = self.add_section(body, "1. 模板版本")
        ttk.Label(template_frame, text="根据实际资料选择协议和表格版本。选择无培肥时，培肥输入项会自动隐藏。", style="Hint.TLabel", wraplength=820).pack(anchor="w", pady=(0, 8))
        self.make_var("template_choice", "without_v2")
        for value in ["without", "with", "without_v2", "with_v2"]:
            ttk.Radiobutton(
                template_frame,
                text=str(TEMPLATE_OPTIONS[value]["label"]),
                variable=self.vars["template_choice"],
                value=value,
                command=self.apply_template_defaults,
            ).pack(side="left", padx=(0, 18))

        identity = self.add_section(body, "2. 人员、甲方地址和账户")
        default_name = data.name or ("" if self.manual_no_cert_doc else (self.selected_folder.name if self.selected_folder else ""))
        self.add_entry(identity, "name", "姓名", default_name)
        self.add_entry(identity, "address", "甲方地址", data.address)
        self.add_entry(identity, "town", "甲方镇", data.town)
        self.add_entry(identity, "village", "甲方村", data.village)
        self.add_entry(identity, "team", "甲方队", data.team)
        self.add_entry(identity, "id_number", "身份证号", data.id_number)
        self.add_entry(identity, "bank_card", "银行卡号", data.bank_card)
        self.add_entry(identity, "bank_name", "银行名称", data.bank_name)

        land = self.add_section(body, "3. 铁塔位置、编号和面积")
        self.add_entry(
            land,
            "tower_location",
            "铁塔位置（镇村队）",
            location_to_team_address(data.address, data.town, data.village, data.team),
        )
        self.add_entry(land, "tower_ids", "铁塔编号", "")
        self.add_entry(land, "permanent_area", "长期租地面积(亩)", "")
        self.add_entry(land, "temporary_area", "临时租地面积(亩)", "")
        self.add_entry(land, "construction_area", "其中施工用地(亩)", "", required=False)
        self.add_entry(land, "road_area", "其中道路用地(亩)", "", required=False)
        self.add_entry(land, "rent_years", "租期(年)", "" if self.table_input_mode else "2")

        price = self.add_section(body, "4. 单价、接地和附作物")
        self.add_entry(price, "permanent_unit", "林地单价", "" if self.table_input_mode else "48600")
        self.add_entry(price, "construction_unit", "施工用地单价", "" if self.table_input_mode else "400")
        self.add_entry(price, "road_unit", "道路用地单价", "" if self.table_input_mode else "400")
        self.add_entry(price, "ground_count", "接地数量(基)", "")
        self.add_entry(price, "ground_unit", "接地单价", "" if self.table_input_mode else "400")

        self.required_labels["attachment_area"] = "附作物面积(亩)"
        self.required_labels["attachment_unit"] = "附作物单价"
        attachment_panel = tk.Frame(price, bg="#fffaf0", padx=12, pady=10, highlightthickness=1, highlightbackground=self.panel_line)
        attachment_panel.pack(fill="x", pady=(10, 6))
        attachment_top = tk.Frame(attachment_panel, bg="#fffaf0")
        attachment_top.pack(fill="x", pady=(0, 8))
        tk.Label(
            attachment_top,
            text="地上附作物",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg=self.accent_color,
            bg="#fffaf0",
        ).pack(side="left")
        ttk.Button(attachment_top, text="+ 添加附作物", command=self.add_attachment_row).pack(side="right")

        header = tk.Frame(attachment_panel, bg="#f5e6bf", padx=6, pady=4)
        header.pack(fill="x", pady=(0, 4))
        header.columnconfigure(1, weight=1)
        header.columnconfigure(2, weight=1)
        header.columnconfigure(3, weight=1)
        header.columnconfigure(4, weight=1)
        header.columnconfigure(5, weight=1)
        for col, (text, width) in enumerate([("项", 4), ("用地", 10), ("名称", 14), ("规格", 14), ("面积(亩) *", 10), ("单价 *", 10), ("", 4)]):
            tk.Label(header, text=text, width=width, anchor="center", bg="#f5e6bf", fg="#705014", font=("Microsoft YaHei UI", 9, "bold")).grid(
                row=0, column=col, padx=(0, 8), sticky="ew"
            )

        first_attachment_row = ttk.Frame(attachment_panel)
        first_attachment_row.pack(fill="x", pady=3)
        for col in range(1, 6):
            first_attachment_row.columnconfigure(col, weight=1)
        ttk.Label(first_attachment_row, text="1.", width=4, anchor="e").grid(row=0, column=0, padx=(0, 8))
        ttk.Combobox(
            first_attachment_row,
            textvariable=self.make_var("attachment_land_use", "" if self.table_input_mode else "施工用地"),
            values=["", "施工用地", "道路用地"] if self.table_input_mode else ["施工用地", "道路用地"],
            state="readonly",
            width=10,
        ).grid(row=0, column=1, padx=(0, 8), sticky="ew")
        first_name_var = self.make_var("attachment_name", "")
        first_spec_var = self.make_var("attachment_spec", "")
        first_name_box = ttk.Combobox(
            first_attachment_row,
            textvariable=first_name_var,
            values=[name for name, _ in COMMON_ATTACHMENT_OPTIONS],
            state="normal",
            width=14,
        )
        first_name_box.grid(row=0, column=2, padx=(0, 8), sticky="ew")
        first_name_box.bind(
            "<<ComboboxSelected>>",
            lambda _event: self.set_attachment_default_spec(first_name_var, first_spec_var),
        )
        ttk.Entry(first_attachment_row, textvariable=first_spec_var, width=14).grid(row=0, column=3, padx=(0, 8), sticky="ew")
        ttk.Entry(first_attachment_row, textvariable=self.make_var("attachment_area", ""), width=10).grid(row=0, column=4, padx=(0, 8), sticky="ew")
        ttk.Entry(first_attachment_row, textvariable=self.make_var("attachment_unit", ""), width=10).grid(row=0, column=5, padx=(0, 8), sticky="ew")
        ttk.Label(first_attachment_row, text="", width=4).grid(row=0, column=6)

        self.attachment_rows_frame = ttk.Frame(attachment_panel)
        self.attachment_rows_frame.pack(fill="x", pady=(1, 0))
        self.attachment_extra_rows = []
        self.attachment_row_seq = 2

        seedling_panel = tk.Frame(price, bg="#eef8ef", padx=12, pady=10, highlightthickness=1, highlightbackground="#a7cfa9")
        seedling_panel.pack(fill="x", pady=(4, 8))
        seedling_row = tk.Frame(seedling_panel, bg="#eef8ef")
        seedling_row.pack(fill="x")
        tk.Label(
            seedling_row,
            text="青苗总额 元",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#2f6b38",
            bg="#eef8ef",
        ).pack(side="left")
        ttk.Entry(seedling_row, textvariable=self.make_var("seedling_total", ""), width=18).pack(side="left", padx=(12, 8))
        tk.Label(
            seedling_row,
            text="模拟种类数",
            font=("Microsoft YaHei UI", 10),
            fg="#2f6b38",
            bg="#eef8ef",
        ).pack(side="left", padx=(12, 6))
        ttk.Combobox(
            seedling_row,
            textvariable=self.make_var("seedling_count", "自动（1-3种）"),
            values=["自动（1-3种）", "1种", "2种", "3种"],
            state="readonly",
            width=15,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(seedling_row, text="预算模拟", command=self.simulate_seedlings).pack(side="left")
        self.seedling_status_var = tk.StringVar(value="手填青苗的数量或面积、单价后，青苗总额会自动合计。")
        tk.Label(
            seedling_panel,
            textvariable=self.seedling_status_var,
            font=("Microsoft YaHei UI", 9),
            fg="#4d7251",
            bg="#eef8ef",
            anchor="w",
        ).pack(fill="x", pady=(7, 0))
        self.fertilizer_rows = [
            self.add_entry(price, "fertilizer_area", "培肥面积(亩)", "", required=False),
            self.add_entry(price, "fertilizer_unit", "培肥单价", "" if self.table_input_mode else "1500", required=False),
        ]

        missing_frame = self.add_section(body, "5. 待补充信息")
        self.missing_var = tk.StringVar()
        ttk.Label(missing_frame, textvariable=self.missing_var, style="Status.TLabel", wraplength=850, justify="left").pack(anchor="w", fill="x")

        button_bar = ttk.Frame(self.main, style="Toolbar.TFrame")
        button_bar.pack(fill="x", pady=(10, 0))
        self.generate_button = ttk.Button(button_bar, text="生成全部文件", command=self.start_generate, style="Primary.TButton")
        self.generate_button.pack(side="right")
        ttk.Button(button_bar, text="退出", command=self.root.destroy).pack(side="right", padx=(0, 8))
        self.add_author_footer()

        self.apply_template_defaults()
        self.update_missing()

    def apply_template_defaults(self) -> None:
        choice = self.vars.get("template_choice").get() if "template_choice" in self.vars else "without_v2"
        with_fertilizer = is_fertilizer_choice(choice)
        self.update_fertilizer_fields(with_fertilizer)
        if not self.table_input_mode and with_fertilizer and self.vars.get("fertilizer_area") and not self.vars["fertilizer_area"].get().strip():
            self.vars["fertilizer_area"].set(self.vars.get("construction_area", tk.StringVar()).get().strip())
        self.update_missing()

    def update_fertilizer_fields(self, show: bool) -> None:
        for row in getattr(self, "fertilizer_rows", []):
            if show:
                row.pack(fill="x", pady=4)
            else:
                row.pack_forget()
        if not show and self.vars.get("fertilizer_area"):
            self.vars["fertilizer_area"].set("")

    def collect_form_data(self) -> dict[str, str]:
        return {key: var.get().strip() for key, var in self.vars.items()}

    def update_missing(self) -> None:
        if self.missing_var is None:
            return
        data = normalize_form_data(self.collect_form_data())
        missing = [label for key, label in self.required_labels.items() if not data.get(key, "").strip()]
        if is_fertilizer_choice(data.get("template_choice", "without")):
            for key, label in [("fertilizer_area", "培肥面积(亩)"), ("fertilizer_unit", "培肥单价")]:
                if not data.get(key, "").strip():
                    missing.append(label)
        if missing:
            self.missing_var.set("待补充：" + "、".join(missing))
        else:
            self.missing_var.set("信息已基本完整，可以生成文件。")

    def start_generate(self) -> None:
        if self.selected_folder is None:
            messagebox.showerror("提示", "请先选择人名文件夹。")
            return
        data = self.collect_form_data()
        data = normalize_form_data(data)
        if self.manual_no_cert_doc and not data.get("name", "").strip():
            messagebox.showerror("提示", "请先填写姓名，生成的文件夹和文件名会按这个姓名命名。")
            return
        missing = [label for key, label in self.required_labels.items() if not data.get(key, "").strip()]
        if is_fertilizer_choice(data.get("template_choice", "without")):
            for key, label in [("fertilizer_area", "培肥面积(亩)"), ("fertilizer_unit", "培肥单价")]:
                if not data.get(key, "").strip():
                    missing.append(label)
        if missing:
            ok = messagebox.askyesno("还有信息未填", "以下信息未填：\n" + "、".join(missing) + "\n\n继续生成时这些位置会留空，是否继续？")
            if not ok:
                return

        self.generate_button.configure(state="disabled")
        self.start_generation_wait(data.get("name", "") or (self.selected_folder.name if self.selected_folder else ""))

        def worker() -> None:
            try:
                def progress(value: int, message: str, detail: str = "") -> None:
                    self.root.after(
                        0,
                        lambda value=value, message=message, detail=detail: self.set_generation_wait_progress(value, message, detail),
                    )

                result = build_all(self.selected_folder, data, progress=progress)
                try:
                    qingpei_db.save_generated_record(
                        data,
                        result.output_dir,
                        [result.cert_doc, result.agreement_doc, result.table_xlsx],
                    )
                except Exception as exc:
                    result.warnings.append(f"资料库登记失败：{exc}")
                self.root.after(0, lambda: self.generate_done(result))
            except Exception:
                error = traceback.format_exc()
                self.root.after(0, lambda: self.generate_failed(error))

        threading.Thread(target=worker, daemon=True).start()

    def generate_done(self, result: GeneratedFiles) -> None:
        self.stop_generation_wait("生成完成")
        self.generate_button.configure(state="normal")
        lines = [f"已生成到：{result.output_dir}", ""]
        if result.cert_doc:
            lines.append(f"证件 Word：{result.cert_doc.name}")
        lines.append(f"协议书：{result.agreement_doc.name}")
        if result.table_xlsx:
            lines.append(f"补偿表：{result.table_xlsx.name}")
        if result.warnings:
            lines.append("")
            lines.extend(result.warnings)
        messagebox.showinfo("完成", "\n".join(lines))

    def generate_failed(self, error: str) -> None:
        self.stop_generation_wait("生成失败")
        self.generate_button.configure(state="normal")
        messagebox.showerror("生成失败", error)

    def run(self) -> int:
        if not CERT_DIR.exists():
            messagebox.showerror("缺少文件夹", f"没有找到证件提取文件夹：\n{CERT_DIR}")
            return 1
        self.root.mainloop()
        return 0


def main() -> int:
    return App().run()


if __name__ == "__main__":
    raise SystemExit(main())
