# -*- coding: utf-8 -*-
"""
从“证件提取”里选择人名文件夹，把资料填写到“00_模板\\模板（无培肥费版本）.docx”，
并另存到“青赔”文件夹下，文件名为“人名.docx”。

使用方法：
1. 双击运行本脚本。
2. 在窗口里选择要处理的人名文件夹。
3. 检查/补充姓名、地址、身份证号、银行卡号、银行名称。
4. 点击生成。

说明：
- 脚本会尽量从所选人名文件夹里的 Word 文档读取文字资料。
- 如果证件资料只是图片，电脑没有 OCR 时无法自动读出身份证号和银行卡号，
  这时在弹窗里手动填写即可。
- 哪个信息不填，就不改模板里对应位置，模板其它内容保持不变。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path


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
except ImportError as exc:
    print("缺少 Word 处理组件，无法运行。")
    print("请先安装：pip install python-docx")
    print(f"具体缺少：{exc}")
    input("按回车键退出...")
    raise SystemExit(1)


SCRIPT_DIR = Path(__file__).resolve().parent
QP_DIR = SCRIPT_DIR.parent
TEMPLATE_DIR = QP_DIR / "00_模板"
TEMPLATE_CANDIDATES = [
    TEMPLATE_DIR / "模板（无培肥费版本）.docx",
    TEMPLATE_DIR / "模板（无培肥费版）.docx",
    QP_DIR / "00_模板的模板（无培肥费版本）" / "模板（无培肥费版本）.docx",
]

TEMP_WORD_PREFIX = "~$"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


@dataclass
class PersonData:
    name: str = ""
    address: str = ""
    town: str = ""
    village: str = ""
    team: str = ""
    id_number: str = ""
    bank_card: str = ""
    bank_name: str = ""
    output_name: str = ""


def find_template() -> Path:
    for path in TEMPLATE_CANDIDATES:
        if path.exists() and not path.name.startswith(TEMP_WORD_PREFIX):
            return path
    raise FileNotFoundError("没有找到模板：00_模板\\模板（无培肥费版本）.docx")


def safe_filename(name: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name).strip(" .")
    return safe or "青赔协议"


def person_dirs() -> list[Path]:
    ignored = {"__pycache__"}
    dirs: list[Path] = []
    for path in sorted(SCRIPT_DIR.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_dir():
            continue
        if path.name.startswith(".") or path.name.startswith("_") or path.name in ignored:
            continue
        dirs.append(path)
    return dirs


def docx_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(folder.glob("*.docx"), key=lambda p: p.name.lower()):
        if path.name.startswith(TEMP_WORD_PREFIX):
            continue
        files.append(path)
    return files


def image_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
            files.append(path)
    return files


def read_docx_text(path: Path) -> str:
    try:
        document = Document(str(path))
    except Exception:
        return ""

    chunks: list[str] = []
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            chunks.append(paragraph.text.strip())

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    chunks.append(text)

    return "\n".join(chunks)


_OCR_ENGINE = None


def get_ocr_engine():
    global _OCR_ENGINE
    if _OCR_ENGINE is not None:
        return _OCR_ENGINE

    try:
        from rapidocr_onnxruntime import RapidOCR
    except Exception:
        _OCR_ENGINE = False
        return None

    try:
        _OCR_ENGINE = RapidOCR()
    except Exception:
        _OCR_ENGINE = False
        return None
    return _OCR_ENGINE


def read_image_ocr_text(path: Path) -> str:
    ocr = get_ocr_engine()
    if not ocr:
        return ""

    try:
        result, _ = ocr(str(path))
    except Exception:
        return ""

    if not result:
        return ""

    lines: list[str] = []
    for item in result:
        if len(item) >= 2 and str(item[1]).strip():
            lines.append(str(item[1]).strip())
    return "\n".join(lines)


def read_bank_card_enhanced_text(path: Path) -> str:
    ocr = get_ocr_engine()
    if not ocr:
        return ""

    try:
        from PIL import Image, ImageEnhance, ImageOps
    except Exception:
        return ""

    try:
        image = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    except Exception:
        return ""

    width, height = image.size
    crop_boxes = [
        ("bank_lower_middle", (0, int(height * 0.42), width, int(height * 0.68))),
        ("bank_middle_band", (0, int(height * 0.36), width, int(height * 0.72))),
    ]

    lines: list[str] = []
    with tempfile.TemporaryDirectory(prefix="bank_ocr_") as temp_root:
        temp_dir = Path(temp_root)
        for crop_name, box in crop_boxes:
            crop = image.crop(box)
            gray = ImageOps.grayscale(crop)
            scale = 3 if max(gray.size) < 2600 else 2
            big = gray.resize((gray.width * scale, gray.height * scale), Image.Resampling.LANCZOS)
            big = ImageEnhance.Contrast(big).enhance(2.5)
            big = ImageEnhance.Sharpness(big).enhance(2.0)

            variants = [
                ("raw", crop),
                ("gray", big.convert("RGB")),
            ]

            for variant_name, variant in variants:
                image_path = temp_dir / f"{crop_name}_{variant_name}.png"
                try:
                    variant.save(image_path)
                    result, _ = ocr(str(image_path))
                except Exception:
                    continue
                if not result:
                    continue
                for item in result:
                    if len(item) >= 2 and str(item[1]).strip():
                        lines.append(str(item[1]).strip())

    return "\n".join(lines)


def is_likely_bank_image(path: Path, ocr_text: str) -> bool:
    name = path.stem.lower()
    bank_words = [
        "银行卡",
        "银行",
        "bank",
        "card",
        "银联",
        "unionpay",
        "农村信用社",
        "农商银行",
        "邮政",
        "信用社",
        "debit",
        "atm",
        "闪付",
    ]
    compact = re.sub(r"\s+", "", ocr_text.lower())
    return any(word.lower() in name or word.lower() in compact for word in bank_words)


def read_folder_ocr_text(folder: Path) -> str:
    chunks: list[str] = []
    for path in image_files(folder):
        text = read_image_ocr_text(path)
        bank_text = read_bank_card_enhanced_text(path) if is_likely_bank_image(path, text) else ""
        combined = "\n".join(part for part in [text, bank_text] if part)
        if combined:
            chunks.append(f"【{path.name}】\n{combined}")
    return "\n".join(chunks)


def first_match(pattern: str, text: str, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    if not match:
        return ""
    return match.group(1).strip()


def normalize_digits(value: str) -> str:
    return re.sub(r"\D", "", value)


def extract_id_number(text: str) -> str:
    match = re.search(r"(?<!\d)(\d{17}[\dXx])(?![0-9A-Za-z])", text)
    return match.group(1).upper() if match else ""


def extract_bank_card(text: str, id_number: str = "") -> str:
    candidates: list[str] = []
    scores: dict[str, int] = {}

    def add_candidate(digits: str, score: int) -> None:
        if not digits:
            return
        candidates.append(digits)
        scores[digits] = scores.get(digits, 0) + score

    def looks_like_bank_card(raw: str) -> str:
        if re.search(r"\d{4}[./-]\d{2}[./-]\d{2}\s*[-至到]\s*\d{4}[./-]\d{2}[./-]\d{2}", raw):
            return ""
        if re.search(r"\d{2}\s*/\s*\d{2}", raw):
            return ""

        digits = normalize_digits(raw)
        if not (16 <= len(digits) <= 23):
            return ""
        if digits == id_number:
            return ""
        if id_number and digits in id_number:
            return ""
        # 国内常见银行卡多以 6、9 开头；身份证有效期等日期拼接也可能刚好是 16 位，要过滤掉。
        if not digits.startswith(("6", "9")):
            return ""
        if re.fullmatch(r"20\d{14,21}", digits):
            return ""
        return digits

    for match in re.finditer(r"(?<!\d)((?:\d[\s-]?){16,23})(?!\d)", text):
        digits = looks_like_bank_card(match.group(1))
        if digits:
            add_candidate(digits, 8 if " " not in match.group(1) and "-" not in match.group(1) else 6)
    for line in text.splitlines():
        clean = line.strip()
        if not clean:
            continue
        if any(ch.isalpha() for ch in clean):
            continue
        digits = looks_like_bank_card(clean)
        if digits:
            add_candidate(digits, 10 if clean == digits else 7)

    fragments: list[str] = []
    for line in text.splitlines():
        clean = line.strip()
        if re.search(r"\d{2}\s*/\s*\d{2}", clean):
            continue
        if re.search(r"\d{4}[./-]\d{2}[./-]\d{2}\s*[-至到]\s*\d{4}[./-]\d{2}[./-]\d{2}", clean):
            continue
        for token in re.findall(r"\d{4,15}", clean):
            if token == id_number:
                continue
            fragments.append(token)

    for idx, first in enumerate(fragments):
        if not (len(first) == 6 and first.startswith(("62", "60", "61", "63"))):
            continue
        for second in fragments[idx + 1 : idx + 6]:
            if not (10 <= len(second) <= 14):
                continue
            digits = looks_like_bank_card(first + second)
            if digits and 18 <= len(digits) <= 20:
                # 拼接识别容易把装饰数字误拼进卡号，权重低于完整连续识别。
                add_candidate(digits, 3)

    if not candidates:
        return ""
    candidates = list(dict.fromkeys(candidates))
    candidates.sort(
        key=lambda item: (
            scores.get(item, 0),
            item.startswith("62"),
            -abs(len(item) - 19),
            -len(item),
        ),
        reverse=True,
    )
    return candidates[0]


def extract_bank_name(text: str) -> str:
    explicit = first_match(r"银\s*行\s*[:：]\s*([^\n\r]+)", text)
    if explicit:
        return explicit

    banks = [
        "广西农村信用社",
        "农村信用社",
        "广西农商联合银行",
        "农商联合银行",
        "广西农信",
        "中国邮政储蓄银行",
        "中国邮政银行",
        "邮政储蓄银行",
        "中国农业银行",
        "中国工商银行",
        "中国建设银行",
        "中国银行",
        "交通银行",
        "桂林银行",
        "北部湾银行",
    ]
    compact = re.sub(r"\s+", "", text)
    for bank in banks:
        if bank in compact:
            return bank
    return ""


def extract_address(text: str, person_name: str = "") -> str:
    address_label_pattern = r"(?:住\s*址|进\s*址|迸\s*址)"
    address = first_match(address_label_pattern + r"\s*[:：]?\s*([^\n\r]+)", text)
    if address:
        lines = [line.strip() for line in text.splitlines()]
        for idx, line in enumerate(lines):
            if re.search(address_label_pattern, line) and idx + 1 < len(lines):
                compact_address = re.sub(address_label_pattern, "", address).strip()
                extra_parts: list[str] = []
                for next_line in lines[idx + 1 : idx + 8]:
                    next_line = next_line.strip()
                    if not next_line:
                        continue
                    if re.search(r"(公民身份号码|GUNGHMINZ|SINHFWN)", next_line):
                        break
                    if re.search(r"(性别|民族|出生|姓名|有效期限|签发机关|卡号|账号|帐号|账户|银行|银行卡|存折)", next_line):
                        continue
                    if re.fullmatch(r"[A-Za-z\\s]+", next_line):
                        continue
                    chinese = re.sub(r"[^\u4e00-\u9fff0-9号队组屯村镇县乡]", "", next_line)
                    date_tail = re.search(r"\d{4}年\d{1,2}月\d{1,2}日(.+)", chinese)
                    if date_tail:
                        chinese = date_tail.group(1)
                    elif re.search(r"\d{4}年\d{1,2}月\d{1,2}日", chinese):
                        continue
                    if not chinese:
                        continue
                    if re.fullmatch(r"[生民男女汉]+", chinese):
                        continue
                    if person_name and chinese == person_name:
                        break
                    if chinese not in compact_address and chinese not in extra_parts:
                        extra_parts.append(chinese)
                    if "号" in chinese or "队" in chinese:
                        break
                return normalize_address_text(compact_address + "".join(extra_parts))
        return normalize_address_text(address)

    party = first_match(r"被征租方\s*\(甲方\)\s*[:：]\s*([^\n\r]+)", text)
    if party:
        party = party.strip()
        if person_name and party.endswith(person_name):
            party = party[: -len(person_name)].strip()
        return party
    return ""


def strip_address_noise_digits(address: str) -> str:
    address = address or ""
    address = re.sub(r"\d{4}年\d{1,2}月\d{1,2}日", "", address)
    address = re.sub(r"(?<=[\u4e00-\u9fff])(?:\d[\s-]*){6,}(?=[\u4e00-\u9fff])", "", address)
    address = re.sub(r"(?<=[\u4e00-\u9fff])(?:\d[\s-]*){6,}(?=号)", "", address)
    return address


def normalize_address_text(address: str) -> str:
    address = re.sub(r"\s+", "", address or "")
    address = strip_address_noise_digits(address)
    replacements = {
        "博自县": "某县",
        "博日县": "某县",
        "博向县": "某县",
        "松旺镇周比村": "松旺镇周北村",
    }
    for wrong, right in replacements.items():
        address = address.replace(wrong, right)
    if not re.search(r"(镇|乡|街道|村|社区|队|组|屯)", address) and re.search(r"(卡|银行|账号|帐号|账户|存折)", address):
        return ""
    return address


def split_address_parts(address: str) -> tuple[str, str, str]:
    compact = re.sub(r"\s+", "", address or "")
    compact = strip_address_noise_digits(compact)
    compact = compact.replace("广西壮族自治区", "").replace("广西", "")
    compact = compact.replace("玉林市", "").replace("某县", "")
    compact = re.sub(r"^.*?县", "", compact)

    town = ""
    village = ""
    team = ""

    town_match = re.search(r"([\u4e00-\u9fff]{1,12})(?:镇|乡|街道)", compact)
    if town_match:
        town = town_match.group(1)
        after_town = compact[town_match.end() :]
    else:
        after_town = compact

    village_candidates: list[tuple[int, str, str]] = []
    for village_match in re.finditer(r"(村|社区|居委会|居委)", after_town):
        candidate_village = after_town[: village_match.start()]
        remaining = after_town[village_match.end() :]
        if not candidate_village:
            continue
        team_candidate_match = re.search(r"([\u4e00-\u9fff0-9]{1,20})(?:队|组|屯)", remaining)
        score = 0
        if team_candidate_match and team_candidate_match.group(1):
            score += 10
            if team_candidate_match.group(1).startswith("村"):
                score -= 8
        if remaining and not remaining.startswith(("队", "组", "屯", "号")):
            score += 2
        village_candidates.append((score, candidate_village, remaining))

    if village_candidates:
        _score, village, after_village = max(village_candidates, key=lambda item: item[0])
    else:
        after_village = after_town

    team_match = re.search(r"([\u4e00-\u9fff0-9]{1,20})(?:队|组|屯)", after_village)
    if team_match:
        team = team_match.group(1)
        team = re.sub(r".*?\d{4}年\d{1,2}月\d{1,2}日", "", team)
        team = re.sub(r"^(出生|性别|民族)+", "", team)
        # 有些 OCR 会把“某某村队”识别成队名结尾多一个“村”，
        # 但“老村队”本身就是正确队名，短队名不能一刀切去掉。
        if team.endswith("村") and len(team) > 2:
            team = team[:-1]
        # OCR 有时会把门牌“卡号/户号/账号”接在村名后，不能当成队名。
        if re.search(r"(卡号|账号|帐号|户号|门牌|身份证|银行卡)", team):
            team = ""
    return town, village, team


INVALID_NAME_WORDS = {
    "姓名",
    "性别",
    "民族",
    "出生",
    "住址",
    "公民",
    "身份号码",
    "居民身份证",
    "身份证",
    "身份证正面",
    "身份证反面",
    "银行卡",
    "银行",
    "国徽面",
    "人像面",
    "正面",
    "反面",
}


def is_valid_person_name(candidate: str) -> bool:
    clean = re.sub(r"[^\u4e00-\u9fff·]", "", candidate or "")
    if not (2 <= len(clean) <= 5):
        return False
    if clean in INVALID_NAME_WORDS:
        return False
    if re.search(r"(姓名|性别|民族|出生|住址|公民|身份|号码|银行|正面|反面|国徽|人像)", clean):
        return False
    return True


def extract_name(text: str, folder_name: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        if re.search(r"姓\s*名", line):
            inline = re.sub(r".*?姓\s*名\s*[:：]?", "", line).strip()
            candidates = [inline]
            if idx > 0:
                candidates.insert(0, lines[idx - 1].strip())
            if idx + 1 < len(lines):
                candidates.append(lines[idx + 1].strip())
            for candidate in candidates:
                clean = re.sub(r"[^\u4e00-\u9fff·]", "", candidate)
                if is_valid_person_name(clean):
                    return clean

    if is_valid_person_name(folder_name):
        return folder_name

    for candidate in re.findall(r"[\u4e00-\u9fff·]{2,5}", text):
        if is_valid_person_name(candidate):
            return candidate
    return folder_name


def extract_person_data(folder: Path, use_ocr: bool = True) -> PersonData:
    text_parts: list[str] = []
    for path in docx_files(folder):
        text = read_docx_text(path)
        if text:
            text_parts.append(text)

    if use_ocr:
        ocr_text = read_folder_ocr_text(folder)
        if ocr_text:
            text_parts.append(ocr_text)

    text_all = "\n".join(text_parts)
    data = PersonData()
    data.name = extract_name(text_all, folder.name)
    data.id_number = extract_id_number(text_all)
    data.bank_card = extract_bank_card(text_all, data.id_number)
    data.bank_name = extract_bank_name(text_all)
    data.address = extract_address(text_all, data.name)
    data.town, data.village, data.team = split_address_parts(data.address)
    return data


def set_paragraph_text(paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def set_run(paragraph, index: int, text: str) -> None:
    if len(paragraph.runs) > index:
        paragraph.runs[index].text = text


def spaced_name(name: str, gap: str = " ") -> str:
    clean = re.sub(r"\s+", "", name or "")
    if not clean:
        return ""
    return gap.join(clean)


def iter_all_paragraphs(document: Document):
    for paragraph in document.paragraphs:
        yield paragraph

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph

    for section in document.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                yield paragraph
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            yield paragraph


def fill_document(template_path: Path, output_path: Path, data: PersonData) -> None:
    work_path = output_path
    if template_path.resolve() != output_path.resolve():
        shutil.copy2(template_path, output_path)

    document = Document(str(work_path))

    town = data.town.strip()
    village = data.village.strip()
    team = data.team.strip()
    party_name = spaced_name(data.name, " ")
    account_name = spaced_name(data.name, "  ")

    if len(document.paragraphs) > 3 and any([town, village, team, party_name]):
        p = document.paragraphs[3]
        set_run(p, 3, f" {town}" if town else "   ")
        set_run(p, 7, f" {village}" if village else "    ")
        set_run(p, 10, team if team else "    ")
        set_run(p, 11, "队" if team else "    ")
        if len(p.runs) > 11 and team:
            p.runs[11].underline = None
        set_run(p, 13, "")
        set_run(p, 14, "")
        set_run(p, 15, f"{party_name}    " if party_name else "    ")

    if len(document.paragraphs) > 4 and data.id_number:
        p = document.paragraphs[4]
        set_run(p, 2, f" {data.id_number} ")
        set_run(p, 3, " ")

    if len(document.paragraphs) > 7 and town:
        p = document.paragraphs[7]
        set_run(p, 3, f"  {town}  ")

    if len(document.paragraphs) > 10 and any([town, village, team]):
        p = document.paragraphs[10]
        set_run(p, 9, f" {town}  " if town else "   ")
        set_run(p, 11, f"  {village}  " if village else "    ")
        set_run(p, 14, f" {team} " if team else "  ")

    if len(document.paragraphs) > 36 and account_name:
        p = document.paragraphs[36]
        set_run(p, 3, f"{account_name}            ")

    if len(document.paragraphs) > 37 and data.id_number:
        p = document.paragraphs[37]
        set_run(p, 2, f" {data.id_number}  ")
        set_run(p, 3, " ")
        set_run(p, 4, " ")
        set_run(p, 5, "")

    if len(document.paragraphs) > 38 and data.bank_card:
        p = document.paragraphs[38]
        set_run(p, 4, f"  {data.bank_card} ")
        set_run(p, 5, "")

    if len(document.paragraphs) > 39 and data.bank_name:
        p = document.paragraphs[39]
        set_run(p, 2, f" {data.bank_name}       ")
        set_run(p, 3, "")
        set_run(p, 4, "")
        set_run(p, 5, "")

    document.save(str(output_path))


def generate_for_folder(folder: Path, data: PersonData) -> Path:
    template = find_template()
    if data.output_name:
        out_name = safe_filename(data.output_name) + ".docx"
    else:
        out_name = safe_filename(data.name or folder.name) + "-铁塔征租地补偿协议书.docx"
    output_path = QP_DIR / out_name
    fill_document(template, output_path, data)
    return output_path


def run_cli(args: argparse.Namespace) -> int:
    folder = SCRIPT_DIR / args.person if args.person else None
    if folder is None or not folder.is_dir():
        print("请用 --person 指定证件提取里的人名文件夹。")
        return 1

    data = extract_person_data(folder, use_ocr=not args.no_ocr)
    if args.name:
        data.name = args.name
    if args.address:
        data.address = args.address
        data.town, data.village, data.team = split_address_parts(data.address)
    if args.town:
        data.town = args.town
    if args.village:
        data.village = args.village
    if args.team:
        data.team = args.team
    if args.id_number:
        data.id_number = args.id_number
    if args.bank_card:
        data.bank_card = args.bank_card
    if args.bank_name:
        data.bank_name = args.bank_name
    if args.output_name:
        data.output_name = args.output_name

    if args.print_data:
        print(f"人名：{data.name}")
        print(f"地址：{data.address}")
        print(f"镇：{data.town}")
        print(f"村：{data.village}")
        print(f"队：{data.team}")
        print(f"身份证号：{data.id_number}")
        print(f"银行卡号：{data.bank_card}")
        print(f"银行名称：{data.bank_name}")
        return 0

    output = generate_for_folder(folder, data)
    print(f"已生成：{output}")
    return 0


def run_gui() -> int:
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception as exc:
        print(f"无法打开窗口：{exc}")
        return 1

    folders = person_dirs()
    if not folders:
        messagebox.showerror("没有人名文件夹", "证件提取目录里没有找到人名文件夹。")
        return 1

    root = tk.Tk()
    root.title("填写青赔协议模板")
    root.geometry("560x420")
    root.resizable(False, False)

    tk.Label(root, text="选择证件提取里的人名文件夹：", anchor="w").pack(fill="x", padx=14, pady=(12, 4))

    listbox = tk.Listbox(root, height=10, exportselection=False)
    for folder in folders:
        listbox.insert(tk.END, folder.name)
    listbox.selection_set(0)
    listbox.pack(fill="both", expand=True, padx=14, pady=4)

    tip = (
        "说明：脚本会先从所选文件夹里的 Word 尝试读取资料。"
        "如果没有读到，可以在下一步手动填写；空着的信息不会写入模板。"
    )
    tk.Label(root, text=tip, anchor="w", justify="left", wraplength=520).pack(fill="x", padx=14, pady=(4, 8))

    def selected_folder() -> Path | None:
        selection = listbox.curselection()
        if not selection:
            return None
        return folders[selection[0]]

    def show_form(folder: Path, data: PersonData) -> None:
        form = tk.Toplevel(root)
        form.title(f"确认资料 - {folder.name}")
        form.geometry("700x520")
        form.minsize(650, 500)
        form.resizable(True, True)
        form.transient(root)
        form.grab_set()

        fields = [
            ("name", "人名", data.name),
            ("address", "地址", data.address),
            ("town", "镇", data.town),
            ("village", "村", data.village),
            ("team", "队", data.team),
            ("id_number", "身份证号", data.id_number),
            ("bank_card", "银行卡号", data.bank_card),
            ("bank_name", "银行名称", data.bank_name),
        ]
        entries: dict[str, tk.Entry] = {}

        tk.Label(form, text="请确认资料；没有的信息可以留空。", anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(12, 8)
        )

        for row, (key, label, value) in enumerate(fields, start=1):
            tk.Label(form, text=label + "：", width=10, anchor="e").grid(row=row, column=0, padx=(14, 6), pady=6)
            entry = tk.Entry(form, width=62)
            entry.insert(0, value)
            entry.grid(row=row, column=1, padx=(0, 14), pady=6, sticky="ew")
            entries[key] = entry

        def do_generate() -> None:
            new_data = PersonData(
                name=entries["name"].get().strip() or folder.name,
                address=entries["address"].get().strip(),
                town=entries["town"].get().strip(),
                village=entries["village"].get().strip(),
                team=entries["team"].get().strip(),
                id_number=entries["id_number"].get().strip().upper(),
                bank_card=normalize_digits(entries["bank_card"].get().strip()),
                bank_name=entries["bank_name"].get().strip(),
            )
            try:
                output = generate_for_folder(folder, new_data)
            except PermissionError:
                messagebox.showerror("生成失败", "目标文件可能正在被 Word 打开，请关闭后重试。")
                return
            except Exception as exc:
                messagebox.showerror("生成失败", str(exc))
                return

            messagebox.showinfo("完成", f"已生成：\n{output}")
            form.destroy()

        button_frame = tk.Frame(form)
        button_frame.grid(row=len(fields) + 1, column=0, columnspan=2, pady=16)
        tk.Button(button_frame, text="生成", width=14, command=do_generate).pack(side="left", padx=8)
        tk.Button(button_frame, text="取消", width=14, command=form.destroy).pack(side="left", padx=8)

    def open_form() -> None:
        folder = selected_folder()
        if folder is None:
            messagebox.showwarning("请选择", "请先选择一个人名文件夹。")
            return

        loading = tk.Toplevel(root)
        loading.title("正在识别")
        loading.geometry("360x120")
        loading.resizable(False, False)
        loading.transient(root)
        loading.grab_set()
        tk.Label(loading, text=f"正在识别 {folder.name} 的证件资料，请稍等...", pady=24).pack(fill="both", expand=True)

        next_button.config(state="disabled")

        def worker() -> None:
            try:
                data = extract_person_data(folder)
                error = None
            except Exception as exc:
                data = None
                error = exc

            def finish() -> None:
                next_button.config(state="normal")
                loading.destroy()
                if error is not None:
                    messagebox.showerror("识别失败", str(error))
                    return
                show_form(folder, data)

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    buttons = tk.Frame(root)
    buttons.pack(pady=(4, 14))
    next_button = tk.Button(buttons, text="下一步", width=14, command=open_form)
    next_button.pack(side="left", padx=8)
    tk.Button(buttons, text="退出", width=14, command=root.destroy).pack(side="left", padx=8)

    root.mainloop()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="选择人名文件夹并填写青赔协议模板")
    parser.add_argument("--no-gui", action="store_true", help="不用窗口，改用命令行参数")
    parser.add_argument("--person", help="证件提取里的人名文件夹名称")
    parser.add_argument("--name", help="人名")
    parser.add_argument("--address", help="地址")
    parser.add_argument("--town", help="镇，例如：松旺")
    parser.add_argument("--village", help="村，例如：潭连")
    parser.add_argument("--team", help="队，例如：山岱角")
    parser.add_argument("--id-number", help="身份证号")
    parser.add_argument("--bank-card", help="银行卡号")
    parser.add_argument("--bank-name", help="银行名称")
    parser.add_argument("--output-name", help="输出文件名，不需要写 .docx")
    parser.add_argument("--no-ocr", action="store_true", help="不从图片 OCR 识别资料")
    parser.add_argument("--print-data", action="store_true", help="只打印识别资料，不生成文件")
    args = parser.parse_args()

    if args.no_gui:
        return run_cli(args)
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
