# -*- coding: utf-8 -*-
"""青赔资料本地资料库。

数据仅保存在青赔目录下的 _数据库 文件夹，不需要安装数据库软件。
"""

from __future__ import annotations

import csv
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
DATABASE_DIR = BASE_DIR / "_数据库"
DATABASE_PATH = DATABASE_DIR / "青赔资料库.sqlite"
EXPORT_DIR = DATABASE_DIR / "导出报表"

EXCLUDED_FOLDER_NAMES = {
    "00_模板",
    "00_重复文件待确认",
    "00_项目汇总",
    "_test_output",
    "__pycache__",
    "_自动测试输出",
    "_运行组件",
    "证件提取",
    "_数据库",
}

LOCATION_ALIASES: dict[str, tuple[str, str]] = {
    "松旺周北": ("松旺", "周北"),
}


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: Any) -> str:
    return str(value or "").strip()


def normalize_location_parts(town: str, village: str, team: str = "") -> tuple[str, str, str]:
    """统一镇、村、队字段，兼容地址被写进同一个字段的情况。"""
    town = re.sub(r"\s+", "", clean_text(town))
    village = re.sub(r"\s+", "", clean_text(village))
    team = re.sub(r"\s+", "", clean_text(team))

    combined_match = re.fullmatch(r"(.+?)(?:镇|乡|街道)(.+)", town)
    if combined_match:
        town = combined_match.group(1)
        if not village:
            village = combined_match.group(2)

    town = re.sub(r"(?:镇|乡|街道)$", "", town)
    village = re.sub(r"(?:村|社区)$", "", village)
    team = re.sub(r"(?:队|组)$", "", team)
    if not village and town in LOCATION_ALIASES:
        town, village = LOCATION_ALIASES[town]
    return town, village, team


def number(value: Any) -> float:
    text = clean_text(value).replace(",", "").replace("，", ".")
    if not text or text in {"/", "／"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def calculate_total(data: dict[str, str]) -> float:
    permanent = number(data.get("permanent_area")) * number(data.get("permanent_unit"))
    construction = number(data.get("construction_area") or data.get("temporary_area")) * number(data.get("construction_unit"))
    road = number(data.get("road_area")) * number(data.get("road_unit"))
    ground = number(data.get("ground_count")) * number(data.get("ground_unit"))
    fertilizer = number(data.get("fertilizer_area")) * number(data.get("fertilizer_unit"))
    attachments = 0.0
    index = 1
    while True:
        suffix = "" if index == 1 else f"_{index}"
        keys = [f"attachment_area{suffix}", f"attachment_unit{suffix}"]
        if index > 1 and not any(key in data for key in keys):
            break
        attachments += number(data.get(keys[0])) * number(data.get(keys[1]))
        index += 1
    return round(permanent + construction + road + ground + fertilizer + attachments, 2)


def connect() -> sqlite3.Connection:
    DATABASE_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def initialize() -> None:
    with connect() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS persons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                address TEXT DEFAULT '',
                town TEXT DEFAULT '',
                village TEXT DEFAULT '',
                team TEXT DEFAULT '',
                id_number TEXT DEFAULT '',
                bank_name TEXT DEFAULT '',
                bank_card TEXT DEFAULT '',
                folder_path TEXT DEFAULT '',
                sign_status TEXT DEFAULT '待签字',
                payment_status TEXT DEFAULT '未支付',
                archive_status TEXT DEFAULT '待归档',
                missing_notes TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(name, folder_path)
            );

            CREATE TABLE IF NOT EXISTS compensations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person_id INTEGER NOT NULL,
                tower_ids TEXT DEFAULT '',
                permanent_area TEXT DEFAULT '',
                temporary_area TEXT DEFAULT '',
                construction_area TEXT DEFAULT '',
                road_area TEXT DEFAULT '',
                rent_years TEXT DEFAULT '',
                permanent_unit TEXT DEFAULT '',
                construction_unit TEXT DEFAULT '',
                road_unit TEXT DEFAULT '',
                ground_count TEXT DEFAULT '',
                ground_unit TEXT DEFAULT '',
                fertilizer_area TEXT DEFAULT '',
                fertilizer_unit TEXT DEFAULT '',
                total_amount REAL DEFAULT 0,
                total_source TEXT DEFAULT '',
                template_choice TEXT DEFAULT '',
                updated_at TEXT NOT NULL,
                FOREIGN KEY(person_id) REFERENCES persons(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS attachments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person_id INTEGER NOT NULL,
                file_type TEXT NOT NULL,
                file_path TEXT NOT NULL,
                exists_flag INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                UNIQUE(person_id, file_path),
                FOREIGN KEY(person_id) REFERENCES persons(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS preferences (
                preference_key TEXT PRIMARY KEY,
                preference_value TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL
            );
            """
        )
        columns = {row["name"] for row in connection.execute("PRAGMA table_info(compensations)")}
        if "total_source" not in columns:
            connection.execute("ALTER TABLE compensations ADD COLUMN total_source TEXT DEFAULT ''")
        for row in connection.execute("SELECT id, town, village, team FROM persons"):
            town, village, team = normalize_location_parts(row["town"], row["village"], row["team"])
            if (town, village, team) != (row["town"], row["village"], row["team"]):
                connection.execute(
                    "UPDATE persons SET town=?, village=?, team=?, updated_at=? WHERE id=?",
                    (town, village, team, now_text(), int(row["id"])),
                )


def attachment_type(path: Path) -> str:
    name = path.name
    if "协议" in name and path.suffix.lower() == ".docx":
        return "补偿协议"
    if "补偿表" in name and path.suffix.lower() == ".xlsx":
        return "补偿表"
    if path.suffix.lower() in {".jpg", ".jpeg", ".png", ".pdf"}:
        return "证件或凭证"
    if path.suffix.lower() == ".docx":
        return "Word资料"
    if path.suffix.lower() == ".xlsx":
        return "Excel资料"
    return "其他资料"


def record_files(connection: sqlite3.Connection, person_id: int, files: list[Path]) -> None:
    created_at = now_text()
    for path in files:
        connection.execute(
            """
            INSERT INTO attachments (person_id, file_type, file_path, exists_flag, created_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(person_id, file_path) DO UPDATE SET
                file_type = excluded.file_type,
                exists_flag = excluded.exists_flag
            """,
            (person_id, attachment_type(path), str(path), int(path.exists()), created_at),
        )


def sheet_number(sheet, cell_address: str) -> float | None:
    value = sheet[cell_address].value
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def formula_number(sheet, cell_address: str, memo: dict[str, float | None] | None = None) -> float | None:
    """计算补偿表中的简单加减乘除和 SUM 公式，避免使用 Excel 的旧缓存值。"""
    memo = memo if memo is not None else {}
    address = cell_address.upper()
    if address in memo:
        return memo[address]
    value = sheet[address].value
    if isinstance(value, bool):
        memo[address] = None
        return None
    if isinstance(value, (int, float)):
        memo[address] = float(value)
        return memo[address]
    if not isinstance(value, str) or not value.startswith("="):
        memo[address] = None
        return None

    expression = value[1:].replace(" ", "").upper()

    def resolve(token: str) -> float | None:
        if re.fullmatch(r"[A-Z]+\d+", token):
            return formula_number(sheet, token, memo)
        try:
            return float(token)
        except ValueError:
            return None

    result: float | None = None
    sum_match = re.fullmatch(r"SUM\((.+)\)", expression)
    if sum_match:
        values = [resolve(token) for token in sum_match.group(1).split(",")]
        result = sum(item for item in values if item is not None) if any(item is not None for item in values) else None
    elif re.fullmatch(r"[A-Z0-9.+*/-]+", expression):
        pieces = re.split(r"([+*/-])", expression)
        first = resolve(pieces[0])
        if first is not None:
            result = first
            for operator, token in zip(pieces[1::2], pieces[2::2]):
                operand = resolve(token)
                if operand is None:
                    result = None
                    break
                if operator == "+":
                    result += operand
                elif operator == "-":
                    result -= operand
                elif operator == "*":
                    result *= operand
                elif operator == "/":
                    result = result / operand if operand else None
                    if result is None:
                        break
    memo[address] = result
    return result


def extract_from_compensation_table(path: Path) -> tuple[str, float | None]:
    """读取补偿表中的塔号和兑现表总计，公式实时计算而不是读取旧缓存。"""
    try:
        workbook = load_workbook(path, data_only=False, read_only=True)
    except Exception:
        return "", None

    try:
        cash = workbook["兑现表"] if "兑现表" in workbook.sheetnames else None
        land = workbook["土地结算表"] if "土地结算表" in workbook.sheetnames else None
        attachment = workbook["附作物结算表"] if "附作物结算表" in workbook.sheetnames else None
        tower_ids = ""
        for sheet, address in [(cash, "B5"), (land, "G5"), (attachment, "H5")]:
            if sheet is not None:
                tower_ids = clean_text(sheet[address].value)
                if tower_ids:
                    break

        direct_total = None
        if cash is not None:
            for address in ("P16", "P5"):
                direct_total = formula_number(cash, address)
                if direct_total is not None:
                    break
        if direct_total is not None:
            return tower_ids, round(direct_total, 2)

        totals = []
        if land is not None:
            land_total = formula_number(land, "C10")
            if land_total is not None:
                totals.append(land_total)
        if attachment is not None:
            attachment_total = formula_number(attachment, "C10")
            if attachment_total is not None:
                totals.append(attachment_total)
        return tower_ids, round(sum(totals), 2) if totals else None
    finally:
        workbook.close()


def extract_from_agreement(path: Path) -> tuple[str, float | None]:
    """使用现有协议解析器识别铁塔编号和协议合计金额。"""
    try:
        import agreement_to_table

        data = agreement_to_table.parse_agreement(path)
        total = round(float(data.total_amount), 2) if data.total_amount is not None else None
        return clean_text(data.tower_ids), total
    except Exception:
        return "", None


def merge_tower_ids(values: list[str]) -> str:
    """合并多份协议中的塔号，并保持第一次出现的顺序。"""
    seen: set[str] = set()
    tower_ids: list[str] = []
    for value in values:
        for tower_id in re.split(r"[\s,，、/;；]+", clean_text(value)):
            if tower_id and tower_id not in seen:
                seen.add(tower_id)
                tower_ids.append(tower_id)
    return "、".join(tower_ids)


def build_address(town: str, village: str, team: str) -> str:
    town, village, team = normalize_location_parts(town, village, team)
    parts = ["某县"]
    if clean_text(town):
        parts.append(f"{clean_text(town)}镇")
    if clean_text(village):
        parts.append(f"{clean_text(village)}村")
    if clean_text(team):
        parts.append(f"{clean_text(team)}队")
    return "".join(parts) if len(parts) > 1 else ""


def extract_location_from_agreement(files: list[Path]) -> tuple[str, str, str, str]:
    for path in files:
        if path.suffix.lower() != ".docx":
            continue
        try:
            import agreement_to_table

            data = agreement_to_table.parse_agreement(path)
            town, village, team = normalize_location_parts(data.town, data.village, data.team)
            if not (town or village or team):
                town, village, team = split_address(clean_text(getattr(data, "address", "")))
            if town or village or team:
                return town, village, team, build_address(town, village, team)
        except Exception:
            continue
    return "", "", "", ""


def extract_identity_and_bank(files: list[Path]) -> tuple[str, str, str]:
    """从协议或证件 Word 中补充身份证和收款信息。"""
    id_number = ""
    bank_name = ""
    bank_card = ""
    for path in files:
        if path.suffix.lower() != ".docx":
            continue
        try:
            import agreement_to_table

            data = agreement_to_table.parse_agreement(path)
        except Exception:
            continue
        id_number = id_number or clean_text(data.id_number).upper()
        bank_name = bank_name or clean_text(data.bank_name)
        bank_card = bank_card or re.sub(r"\s+", "", clean_text(data.bank_card))
        if id_number and bank_name and bank_card:
            break
    return id_number, bank_name, bank_card


def detect_tower_and_total(files: list[Path]) -> tuple[str, float | None, str]:
    tables = sorted(
        (path for path in files if path.suffix.lower() == ".xlsx"),
        key=lambda path: ("补偿表" not in path.name, path.name),
    )
    agreements = sorted(
        (path for path in files if path.suffix.lower() == ".docx" and "协议" in path.name),
        key=lambda path: path.name,
    )
    agreement_values = [extract_from_agreement(path) for path in agreements]
    agreement_towers = merge_tower_ids([tower_ids for tower_ids, _ in agreement_values])
    agreement_totals = [total for _, total in agreement_values if total is not None]
    if agreement_towers or agreement_totals:
        source = "多份协议合计" if len(agreement_values) > 1 else "协议书"
        return agreement_towers, round(sum(agreement_totals), 2) if agreement_totals else None, source

    tower_ids = ""
    total_amount = None
    for path in tables:
        table_tower, table_total = extract_from_compensation_table(path)
        tower_ids = tower_ids or table_tower
        if table_total is not None:
            return tower_ids, table_total, "补偿表"
    return tower_ids, total_amount, ""


def source_priority(source: str) -> int:
    return {"补偿表": 1, "项目汇总表": 2, "自动计算": 3, "协议书": 4, "多份协议合计": 5, "人工确认": 9}.get(source, 0)


def save_scanned_compensation(connection: sqlite3.Connection, person_id: int, tower_ids: str, total_amount: float | None, source: str = "") -> None:
    row = connection.execute(
        "SELECT id, tower_ids, total_amount, total_source FROM compensations WHERE person_id = ? ORDER BY id DESC LIMIT 1",
        (person_id,),
    ).fetchone()
    timestamp = now_text()
    if row is None:
        connection.execute(
            "INSERT INTO compensations (person_id, tower_ids, total_amount, total_source, updated_at) VALUES (?, ?, ?, ?, ?)",
            (person_id, tower_ids, total_amount or 0, source, timestamp),
        )
        return

    current_source = clean_text(row["total_source"])
    can_replace_tower = bool(tower_ids) and (
        not clean_text(row["tower_ids"]) or source_priority(source) >= source_priority(current_source)
    )
    detected_tower = tower_ids if can_replace_tower else clean_text(row["tower_ids"])
    can_replace_total = total_amount is not None and source_priority(source) >= source_priority(current_source)
    detected_total = total_amount if can_replace_total else float(row["total_amount"] or 0)
    detected_source = source if can_replace_total else current_source
    connection.execute(
        "UPDATE compensations SET tower_ids=?, total_amount=?, total_source=?, updated_at=? WHERE id=?",
        (detected_tower, detected_total, detected_source, timestamp, int(row["id"])),
    )


def save_generated_record(data: dict[str, str], output_dir: Path, generated_files: list[Path | None]) -> int:
    initialize()
    timestamp = now_text()
    name = clean_text(data.get("name")) or output_dir.name
    folder_path = str(output_dir)
    fields = [
        "address", "town", "village", "team", "id_number", "bank_name", "bank_card",
    ]
    with connect() as connection:
        existing = connection.execute(
            "SELECT id FROM persons WHERE name = ? AND folder_path = ?", (name, folder_path)
        ).fetchone()
        values = [clean_text(data.get(field)) for field in fields]
        if existing:
            person_id = int(existing["id"])
            connection.execute(
                """
                UPDATE persons SET address=?, town=?, village=?, team=?, id_number=?, bank_name=?, bank_card=?, updated_at=?
                WHERE id=?
                """,
                (*values, timestamp, person_id),
            )
        else:
            cursor = connection.execute(
                """
                INSERT INTO persons (name, address, town, village, team, id_number, bank_name, bank_card, folder_path, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (name, *values, folder_path, timestamp, timestamp),
            )
            person_id = int(cursor.lastrowid)

        compensation_fields = [
            "tower_ids", "permanent_area", "temporary_area", "construction_area", "road_area", "rent_years",
            "permanent_unit", "construction_unit", "road_unit", "ground_count", "ground_unit",
            "fertilizer_area", "fertilizer_unit", "template_choice",
        ]
        compensation_values = [clean_text(data.get(field)) for field in compensation_fields]
        connection.execute("DELETE FROM compensations WHERE person_id = ?", (person_id,))
        connection.execute(
            """
            INSERT INTO compensations (
                person_id, tower_ids, permanent_area, temporary_area, construction_area, road_area, rent_years,
                permanent_unit, construction_unit, road_unit, ground_count, ground_unit, fertilizer_area,
                fertilizer_unit, template_choice, total_amount, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (person_id, *compensation_values[:-1], compensation_values[-1], calculate_total(data), timestamp),
        )
        record_files(connection, person_id, [path for path in generated_files if path])
    return person_id


def table_cell_text(value: Any) -> str:
    if value is None:
        return ""
    return clean_text(value).replace("\n", "").replace("\r", "")


def find_column(headers: dict[str, int], *names: str) -> int | None:
    for name in names:
        normalized = re.sub(r"\s+", "", name)
        for header, column in headers.items():
            if normalized == header or normalized in header or header in normalized:
                return column
    return None


def split_address(address: str) -> tuple[str, str, str]:
    text = re.sub(r"\s+", "", address)
    tail = text.split("县")[-1] if "县" in text else text
    town_match = re.search(r"(?P<town>[\u4e00-\u9fff]{1,12}?)(?:镇|乡|街道)", tail)
    town = town_match.group("town") if town_match else ""
    village_tail = tail[town_match.end():] if town_match else tail
    village_match = re.search(r"(?P<village>[\u4e00-\u9fff]{1,12}?)(?:村|社区)", village_tail)
    village = village_match.group("village") if village_match else ""
    team_tail = village_tail[village_match.end():] if village_match else village_tail
    team_match = re.search(r"(?P<team>[\u4e00-\u9fff0-9]{1,12}?)(?:队|组)", team_tail)
    team = team_match.group("team") if team_match else ""
    town, village, team = normalize_location_parts(town, village, team)
    if not town and not village and tail in LOCATION_ALIASES:
        town, village = LOCATION_ALIASES[tail]
    return town, village, team


def summary_rows(path: Path) -> list[dict[str, str | float | None]]:
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return []
    rows: list[dict[str, str | float | None]] = []
    try:
        for sheet in workbook.worksheets:
            headers: dict[str, int] | None = None
            header_row = 0
            for row_index, row in enumerate(sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30), values_only=True), start=1):
                candidate = {
                    re.sub(r"\s+", "", table_cell_text(value)): index
                    for index, value in enumerate(row, start=1)
                    if table_cell_text(value)
                }
                if find_column(candidate, "被租方", "姓名", "户主姓名") and find_column(candidate, "补偿费用", "补偿金额", "补偿合计"):
                    headers = candidate
                    header_row = row_index
                    break
            if not headers:
                continue
            name_col = find_column(headers, "被租方", "姓名", "户主姓名")
            address_col = find_column(headers, "位置", "地名", "地址")
            pole_col = find_column(headers, "杆塔号", "塔号")
            pile_col = find_column(headers, "铁塔桩号", "铁塔编号", "桩号")
            amount_col = find_column(headers, "补偿费用", "补偿金额", "补偿合计")
            for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                name = table_cell_text(row[name_col - 1]) if name_col and len(row) >= name_col else ""
                if not name:
                    continue
                address = table_cell_text(row[address_col - 1]) if address_col and len(row) >= address_col else ""
                pole = table_cell_text(row[pole_col - 1]) if pole_col and len(row) >= pole_col else ""
                pile = table_cell_text(row[pile_col - 1]) if pile_col and len(row) >= pile_col else ""
                tower = "/".join(part for part in [pole, pile] if part and part != pole[:])
                if pole and pile and pole != pile:
                    tower = f"{pole}/{pile}"
                elif pole or pile:
                    tower = pole or pile
                raw_amount = row[amount_col - 1] if amount_col and len(row) >= amount_col else None
                amount = number(raw_amount) if raw_amount not in (None, "") else None
                rows.append({"name": name, "address": address, "tower_ids": tower, "total_amount": amount})
    finally:
        workbook.close()
    return rows


def apply_project_summary_data() -> int:
    """用项目汇总表补充完整地址，并在没有协议总额时校正补偿总额。"""
    changed = 0
    workbook_paths = [
        path
        for folder in BASE_DIR.iterdir()
        if folder.is_dir() and folder.name.startswith("00_")
        for path in folder.glob("*.xlsx")
    ]
    for workbook_path in workbook_paths:
        for item in summary_rows(workbook_path):
            name = clean_text(item["name"])
            if not name:
                continue
            address = clean_text(item["address"])
            town, village, team = split_address(address)
            with connect() as connection:
                people = connection.execute("SELECT id FROM persons WHERE name = ?", (name,)).fetchall()
                for person in people:
                    person_id = int(person["id"])
                    connection.execute(
                        """
                        UPDATE persons SET
                            address=CASE WHEN ? <> '' THEN ? ELSE address END,
                            town=CASE WHEN ? <> '' THEN ? ELSE town END,
                            village=CASE WHEN ? <> '' THEN ? ELSE village END,
                            team=CASE WHEN ? <> '' THEN ? ELSE team END,
                            updated_at=?
                        WHERE id=?
                        """,
                        (address, address, town, town, village, village, team, team, now_text(), person_id),
                    )
                    save_scanned_compensation(
                        connection,
                        person_id,
                        clean_text(item["tower_ids"]),
                        item["total_amount"] if isinstance(item["total_amount"], float) else None,
                        "项目汇总表",
                    )
                    changed += 1
    return changed


def normalized_tower_key(value: str) -> str:
    return "".join(re.findall(r"[A-Z0-9+\-]+", clean_text(value).upper()))


def agreement_records(files: list[Path], fallback_name: str) -> list[dict[str, Any]]:
    """从多份协议中提取独立资料记录，并合并内容完全重复的协议文件。"""
    records: list[dict[str, Any]] = []
    seen: dict[tuple[str, str, float | None], dict[str, Any]] = {}
    for path in sorted((item for item in files if item.suffix.lower() == ".docx"), key=lambda item: item.name):
        try:
            import agreement_to_table

            data = agreement_to_table.parse_agreement(path)
        except Exception:
            continue
        name = clean_text(data.name) or fallback_name
        tower_ids = clean_text(data.tower_ids)
        total_amount = round(float(data.total_amount), 2) if data.total_amount is not None else None
        if not tower_ids and total_amount is None:
            continue
        key = (name, normalized_tower_key(tower_ids), total_amount)
        town, village, team = normalize_location_parts(
            clean_text(getattr(data, "town", "")),
            clean_text(getattr(data, "village", "")),
            clean_text(getattr(data, "team", "")),
        )
        address = clean_text(getattr(data, "address", "")) or build_address(town, village, team)
        record = seen.get(key)
        if record is None:
            record = {
                "name": name,
                "tower_ids": tower_ids,
                "total_amount": total_amount,
                "town": town,
                "village": village,
                "team": team,
                "address": address,
                "agreement": path,
                "files": [path],
            }
            seen[key] = record
            records.append(record)
        else:
            record["files"].append(path)
            if "协议" in path.name and "协议" not in record["agreement"].name:
                record["agreement"] = path
    return records


def related_files_for_agreements(files: list[Path], records: list[dict[str, Any]]) -> None:
    """将补偿表、证件 Word 和图片按协议内容或文件顺序归入对应资料记录。"""
    if not records:
        return
    tables = [path for path in files if path.suffix.lower() == ".xlsx"]
    for table in tables:
        table_towers, table_total = extract_from_compensation_table(table)
        matches = [
            record for record in records
            if (table_total is not None and record["total_amount"] == table_total)
            or (normalized_tower_key(table_towers) and normalized_tower_key(table_towers) == normalized_tower_key(record["tower_ids"]))
        ]
        if len(matches) == 1:
            matches[0]["files"].append(table)

    linked_docs = {path for record in records for path in record["files"]}
    certificates = [
        path for path in files
        if path not in linked_docs and path.suffix.lower() in {".docx", ".jpg", ".jpeg", ".png", ".pdf"}
    ]
    for index, path in enumerate(sorted(certificates, key=lambda item: item.name)):
        records[min(index, len(records) - 1)]["files"].append(path)


def save_scanned_person(
    connection: sqlite3.Connection,
    name: str,
    record_path: Path,
    address: str,
    town: str,
    village: str,
    team: str,
    id_number: str,
    bank_name: str,
    bank_card: str,
    files: list[Path],
    tower_ids: str,
    total_amount: float | None,
    total_source: str,
) -> tuple[int, bool]:
    timestamp = now_text()
    row = connection.execute(
        "SELECT id FROM persons WHERE name = ? AND folder_path = ?", (name, str(record_path))
    ).fetchone()
    if row:
        person_id = int(row["id"])
        connection.execute(
            """
            UPDATE persons SET
                address=CASE WHEN ? <> '' THEN ? ELSE address END,
                town=CASE WHEN ? <> '' THEN ? ELSE town END,
                village=CASE WHEN ? <> '' THEN ? ELSE village END,
                team=CASE WHEN ? <> '' THEN ? ELSE team END,
                id_number=CASE WHEN ? <> '' THEN ? ELSE id_number END,
                bank_name=CASE WHEN ? <> '' THEN ? ELSE bank_name END,
                bank_card=CASE WHEN ? <> '' THEN ? ELSE bank_card END,
                updated_at=? WHERE id=?
            """,
            (address, address, town, town, village, village, team, team,
             id_number, id_number, bank_name, bank_name, bank_card, bank_card, timestamp, person_id),
        )
        imported = False
    else:
        cursor = connection.execute(
            """
            INSERT INTO persons (name, address, town, village, team, id_number, bank_name, bank_card, folder_path, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (name, address, town, village, team, id_number, bank_name, bank_card, str(record_path), timestamp, timestamp),
        )
        person_id = int(cursor.lastrowid)
        imported = True
    connection.execute("DELETE FROM attachments WHERE person_id = ? AND file_path LIKE ?", (person_id, "%~$%"))
    record_files(connection, person_id, files)
    save_scanned_compensation(connection, person_id, tower_ids, total_amount, total_source)
    return person_id, imported


def scan_existing_folders() -> tuple[int, int]:
    initialize()
    imported = 0
    files_count = 0
    for folder in BASE_DIR.iterdir():
        if not folder.is_dir() or folder.name in EXCLUDED_FOLDER_NAMES:
            continue
        files = [
            path for path in folder.rglob("*")
            if path.is_file()
            and not path.name.startswith("~$")
            and path.suffix.lower() in {".docx", ".xlsx", ".jpg", ".jpeg", ".png", ".pdf"}
        ]
        if not files:
            continue
        tower_ids, total_amount, total_source = detect_tower_and_total(files)
        town, village, team, address = extract_location_from_agreement(files)
        id_number, bank_name, bank_card = extract_identity_and_bank(files)
        with connect() as connection:
            records = agreement_records(files, folder.name)
            if len(records) > 1:
                related_files_for_agreements(files, records)
                legacy = connection.execute(
                    "SELECT id FROM persons WHERE name = ? AND folder_path = ?", (folder.name, str(folder))
                ).fetchone()
                if legacy:
                    legacy_id = int(legacy["id"])
                    connection.execute("DELETE FROM attachments WHERE person_id = ?", (legacy_id,))
                    connection.execute("DELETE FROM compensations WHERE person_id = ?", (legacy_id,))
                    connection.execute("DELETE FROM persons WHERE id = ?", (legacy_id,))
                for record in records:
                    person_id, was_imported = save_scanned_person(
                        connection,
                        record["name"],
                        record["agreement"],
                        record["address"] or address,
                        record["town"] or town,
                        record["village"] or village,
                        record["team"] or team,
                        id_number,
                        bank_name,
                        bank_card,
                        record["files"],
                        record["tower_ids"],
                        record["total_amount"],
                        "协议书",
                    )
                    imported += int(was_imported)
            else:
                _, was_imported = save_scanned_person(
                    connection, folder.name, folder, address, town, village, team,
                    id_number, bank_name, bank_card, files, tower_ids, total_amount, total_source,
                )
                imported += int(was_imported)
        files_count += len(files)
    apply_project_summary_data()
    return imported, files_count


def list_records(keyword: str = "") -> list[sqlite3.Row]:
    initialize()
    term = f"%{clean_text(keyword)}%"
    with connect() as connection:
        return connection.execute(
            """
            SELECT persons.*, COALESCE(compensations.tower_ids, '') AS tower_ids,
                   COALESCE(compensations.total_amount, 0) AS total_amount,
                   COALESCE(compensations.total_source, '') AS total_source,
                   (SELECT COUNT(*) FROM attachments WHERE attachments.person_id = persons.id) AS file_count
            FROM persons
            LEFT JOIN compensations ON compensations.person_id = persons.id
            WHERE persons.name LIKE ? OR persons.town LIKE ? OR persons.village LIKE ?
                  OR persons.folder_path LIKE ? OR COALESCE(compensations.tower_ids, '') LIKE ?
            ORDER BY persons.updated_at DESC, persons.name COLLATE NOCASE
            """,
            (term, term, term, term, term),
        ).fetchall()


def get_preference(key: str, default: str = "") -> str:
    initialize()
    with connect() as connection:
        row = connection.execute(
            "SELECT preference_value FROM preferences WHERE preference_key = ?", (clean_text(key),)
        ).fetchone()
    return clean_text(row["preference_value"]) if row else default


def set_preference(key: str, value: str) -> None:
    initialize()
    with connect() as connection:
        connection.execute(
            """
            INSERT INTO preferences (preference_key, preference_value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(preference_key) DO UPDATE SET
                preference_value = excluded.preference_value,
                updated_at = excluded.updated_at
            """,
            (clean_text(key), clean_text(value), now_text()),
        )


def list_attachments(person_id: int) -> list[sqlite3.Row]:
    initialize()
    with connect() as connection:
        return connection.execute(
            """
            SELECT id, file_type, file_path, exists_flag, created_at
            FROM attachments WHERE person_id = ?
            ORDER BY
                CASE file_type
                    WHEN '证件或凭证' THEN 1
                    WHEN '补偿协议' THEN 2
                    WHEN '补偿表' THEN 3
                    ELSE 4
                END,
                file_path
            """,
            (person_id,),
        ).fetchall()


def list_agreements(person_id: int) -> list[dict[str, Any]]:
    """返回一人名下的每一份协议，供详情页逐份展示。"""
    agreements: list[dict[str, Any]] = []
    for attachment in list_attachments(person_id):
        if attachment["file_type"] != "补偿协议":
            continue
        path = Path(attachment["file_path"])
        tower_ids, total_amount = extract_from_agreement(path) if path.is_file() else ("", None)
        agreements.append({
            "id": int(attachment["id"]),
            "file_path": str(path),
            "exists_flag": bool(attachment["exists_flag"]) and path.is_file(),
            "tower_ids": tower_ids,
            "total_amount": total_amount,
        })
    return agreements


def get_attachment(attachment_id: int) -> sqlite3.Row | None:
    initialize()
    with connect() as connection:
        return connection.execute(
            "SELECT id, person_id, file_type, file_path, exists_flag FROM attachments WHERE id = ?",
            (attachment_id,),
        ).fetchone()


def update_status(person_id: int, sign_status: str, payment_status: str, archive_status: str, missing_notes: str) -> None:
    initialize()
    with connect() as connection:
        connection.execute(
            """
            UPDATE persons SET sign_status=?, payment_status=?, archive_status=?, missing_notes=?, updated_at=?
            WHERE id=?
            """,
            (sign_status, payment_status, archive_status, missing_notes, now_text(), person_id),
        )


def update_compensation(person_id: int, tower_ids: str, total_amount: float) -> None:
    initialize()
    with connect() as connection:
        row = connection.execute(
            "SELECT id FROM compensations WHERE person_id = ? ORDER BY id DESC LIMIT 1", (person_id,)
        ).fetchone()
        if row is None:
            connection.execute(
            "INSERT INTO compensations (person_id, tower_ids, total_amount, total_source, updated_at) VALUES (?, ?, ?, ?, ?)",
            (person_id, clean_text(tower_ids), total_amount, "人工确认", now_text()),
            )
            return
        connection.execute(
            "UPDATE compensations SET tower_ids=?, total_amount=?, total_source=?, updated_at=? WHERE id=?",
            (clean_text(tower_ids), total_amount, "人工确认", now_text(), int(row["id"])),
        )


def update_person_location(person_id: int, address: str, town: str, village: str, team: str) -> None:
    initialize()
    with connect() as connection:
        connection.execute(
            """
            UPDATE persons SET address=?, town=?, village=?, team=?, updated_at=? WHERE id=?
            """,
            (clean_text(address), clean_text(town), clean_text(village), clean_text(team), now_text(), person_id),
        )


def update_person_identity(person_id: int, id_number: str, bank_name: str, bank_card: str) -> None:
    initialize()
    with connect() as connection:
        connection.execute(
            """
            UPDATE persons SET id_number=?, bank_name=?, bank_card=?, updated_at=? WHERE id=?
            """,
            (
                re.sub(r"\s+", "", clean_text(id_number)).upper(),
                clean_text(bank_name),
                re.sub(r"\s+", "", clean_text(bank_card)),
                now_text(),
                person_id,
            ),
        )


def create_web_record(
    name: str,
    tower_ids: str,
    total_amount: float,
    town: str = "",
    village: str = "",
    team: str = "",
    missing_notes: str = "",
) -> int:
    """创建网页录入的简要档案，后续生成资料时会自动合并到同名文件夹。"""
    initialize()
    name = clean_text(name)
    if not name:
        raise ValueError("请填写姓名。")
    folder_path = str(BASE_DIR / name)
    timestamp = now_text()
    with connect() as connection:
        row = connection.execute(
            "SELECT id FROM persons WHERE name = ? AND folder_path = ?", (name, folder_path)
        ).fetchone()
        if row is None:
            cursor = connection.execute(
                """
                INSERT INTO persons (name, town, village, team, folder_path, missing_notes, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (name, clean_text(town), clean_text(village), clean_text(team), folder_path, clean_text(missing_notes), timestamp, timestamp),
            )
            person_id = int(cursor.lastrowid)
        else:
            person_id = int(row["id"])
            connection.execute(
                """
                UPDATE persons SET town=?, village=?, team=?, missing_notes=?, updated_at=? WHERE id=?
                """,
                (clean_text(town), clean_text(village), clean_text(team), clean_text(missing_notes), timestamp, person_id),
            )
        update_compensation(person_id, tower_ids, total_amount)
    return person_id


def export_records() -> Path:
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = EXPORT_DIR / f"青赔资料汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    rows = list_records()
    headers = ["姓名", "铁塔编号", "镇", "村", "队", "补偿金额", "签字状态", "支付状态", "归档状态", "资料数量", "资料文件夹", "更新时间"]
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([
                row["name"], row["tower_ids"], row["town"], row["village"], row["team"],
                f"{float(row['total_amount'] or 0):.2f}", row["sign_status"], row["payment_status"],
                row["archive_status"], row["file_count"], row["folder_path"], row["updated_at"],
            ])
    return output_path
