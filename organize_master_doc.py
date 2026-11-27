# -*- coding: utf-8 -*-
"""
把一个汇总 Word 里的多人身份证、银行卡图片整理成人名文件夹。

输出目录默认为：
证件提取/_总文档整理/人名/

整理后的图片命名会复用现有软件规则：
- 身份证正面：国徽面
- 身份证反面：人像面
- 银行卡
"""

from __future__ import annotations

import posixpath
import re
import shutil
import subprocess
import xml.etree.ElementTree as ET
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from PIL import Image, ImageOps

import generate_id_doc


MASTER_DOC_NAME = "青赔身份证、银行卡.doc"
MASTER_DOCX_NAME = "青赔身份证、银行卡.docx"
OUTPUT_DIR_NAME = "_总文档整理"
CACHE_DIR_NAME = "_缓存"
STAGING_DIR_NAME = "_临时整理中"
BACKUP_DIR_NAME = "_整理前备份"

IMAGE_LABELS_3 = ["身份证正面", "银行卡", "身份证反面"]
IMAGE_LABELS_2 = ["身份证正面", "身份证反面"]
IMAGE_LABELS_1 = ["身份证正面"]
DOC_IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}

CONTENT_ORDER = {"身份证正面": 0, "身份证反面": 1, "银行卡": 2}


@dataclass
class MasterDocResult:
    output_root: Path
    people_count: int
    image_count: int
    skipped_count: int
    source_docx: Path


def safe_name(text: str) -> str:
    text = re.sub(r"\s+", "", text or "")
    text = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", text).strip(" ._")
    return text or "未命名"


def looks_like_person_name(text: str) -> bool:
    clean = re.sub(r"\s+", "", text or "")
    if not (2 <= len(clean) <= 8):
        return False
    if not re.fullmatch(r"[\u4e00-\u9fff·]+", clean):
        return False
    blocked = {
        "姓名",
        "身份证",
        "银行卡",
        "身份证正面",
        "身份证反面",
        "青赔身份证",
        "开户银行",
        "户主账号",
    }
    return clean not in blocked


def unique_folder(root: Path, name: str, used: set[str]) -> Path:
    base = safe_name(name)
    candidate = base
    index = 2
    while candidate in used:
        candidate = f"{base}_{index}"
        index += 1
    used.add(candidate)
    return root / candidate


def quote_powershell(value: Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def convert_doc_to_docx(source: Path, cache_dir: Path) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    if source.suffix.lower() == ".docx":
        return source

    output = cache_dir / MASTER_DOCX_NAME
    if output.exists() and output.stat().st_mtime >= source.stat().st_mtime:
        return output

    script = f"""
$src={quote_powershell(source)}
$dst={quote_powershell(output)}
$word=$null
$doc=$null
try {{
  $word=New-Object -ComObject Word.Application
  $word.Visible=$false
  $word.DisplayAlerts=0
  $doc=$word.Documents.Open($src, $false, $true)
  $doc.SaveAs2($dst, 16)
  $doc.Close($false)
  $word.Quit()
}} finally {{
  if ($doc -ne $null) {{ try {{ $doc.Close($false) }} catch {{}} }}
  if ($word -ne $null) {{ try {{ $word.Quit() }} catch {{}} }}
}}
"""
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )
    if not output.exists():
        detail = (completed.stderr or completed.stdout or "").strip()
        raise RuntimeError("无法把总证件文档转换成 docx。请确认电脑安装了 Microsoft Word。" + (f"\n{detail}" if detail else ""))
    return output


def relationship_targets(zip_file: ZipFile) -> dict[str, str]:
    rel_root = ET.fromstring(zip_file.read("word/_rels/document.xml.rels"))
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    targets: dict[str, str] = {}
    for rel in rel_root.findall(f"{rel_ns}Relationship"):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            targets[rid] = target
    return targets


def image_path_from_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("word", target))


def parse_groups(docx_path: Path) -> list[tuple[str, list[str]]]:
    with ZipFile(docx_path) as zip_file:
        root = ET.fromstring(zip_file.read("word/document.xml"))

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "v": "urn:schemas-microsoft-com:vml",
    }
    groups: list[tuple[str, list[str]]] = []
    current_name = ""
    current_images: list[str] = []

    for paragraph in root.findall(".//w:body/w:p", ns):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", ns)).strip()
        image_ids = [
            node.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            for node in paragraph.findall(".//v:imagedata", ns)
        ]
        image_ids = [rid for rid in image_ids if rid]

        if looks_like_person_name(text):
            if current_name and current_images:
                groups.append((safe_name(current_name), current_images))
            current_name = safe_name(text)
            current_images = []

        if image_ids and current_name:
            current_images.extend(image_ids)

    if current_name and current_images:
        groups.append((safe_name(current_name), current_images))
    return groups


def labels_for_count(count: int) -> list[str]:
    if count == 1:
        return IMAGE_LABELS_1
    if count == 2:
        return IMAGE_LABELS_2
    labels = IMAGE_LABELS_3[:]
    for index in range(4, count + 1):
        labels.append(f"其他证件{index - 3}")
    return labels


def classify_ocr_text(text: str) -> tuple[str | None, int]:
    compact = re.sub(r"\s+", "", text or "").lower()
    if not compact:
        return None, 0

    bank_words = (
        "银行卡",
        "银行",
        "存折",
        "账号",
        "帐号",
        "账户",
        "户名",
        "开户",
        "unionpay",
        "bank",
        "debit",
        "银联",
        "储蓄卡",
        "借记卡",
        "信用社",
        "农村信用",
        "农商",
        "金穗",
        "个人结算",
        "印鉴",
        "利率",
    )
    front_words = ("中华人民共和国", "居民身份证", "签发机关", "有效期限")
    back_words = ("姓名", "性别", "民族", "出生", "住址", "公民身份号码")

    bank_score = sum(12 for word in bank_words if word.lower() in compact)
    front_score = sum(14 for word in front_words if word in compact)
    back_score = sum(10 for word in back_words if word in compact)
    if bank_score and any(word in compact for word in ("存折", "账号", "帐号", "账户", "户名", "开户", "信用社", "农村信用")):
        bank_score += 35
    if re.search(r"\d{17}[\dx]", compact):
        if bank_score:
            bank_score += 5
        else:
            back_score += 25

    scores = [("银行卡", bank_score), ("身份证正面", front_score), ("身份证反面", back_score)]
    label, score = max(scores, key=lambda item: item[1])
    if score <= 0:
        return None, min(len(compact), 8)
    return label, score + min(len(compact), 20)


def image_format_for_suffix(suffix: str) -> tuple[str, str]:
    if suffix.lower() == ".png":
        return "PNG", ".png"
    return "JPEG", ".jpg"


def save_candidate(image: Image.Image, path: Path, image_format: str) -> None:
    if image_format == "PNG":
        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA")
        image.save(path, "PNG")
    else:
        if image.mode != "RGB":
            image = image.convert("RGB")
        image.save(path, "JPEG", quality=94, optimize=True)


def normalize_and_classify_image(data: bytes, suffix: str, temp_dir: Path, stem: str) -> tuple[Path, str | None]:
    image = Image.open(BytesIO(data))
    image = ImageOps.exif_transpose(image)
    image_format, normalized_suffix = image_format_for_suffix(suffix)

    if image.width >= image.height:
        rotation_candidates = [0, 180, 90, 270]
    else:
        rotation_candidates = [90, 270, 0, 180]

    best_path: Path | None = None
    best_label: str | None = None
    best_score = -1
    for angle in rotation_candidates:
        candidate = image.rotate(angle, expand=True) if angle else image.copy()
        candidate.thumbnail((generate_id_doc.MAX_IMAGE_PIXEL_EDGE, generate_id_doc.MAX_IMAGE_PIXEL_EDGE), Image.Resampling.LANCZOS)
        candidate_path = temp_dir / f"{stem}_r{angle}{normalized_suffix}"
        save_candidate(candidate, candidate_path, image_format)
        ocr_text = generate_id_doc.read_ocr_text(candidate_path)
        label, score = classify_ocr_text(ocr_text)
        if label:
            score += 100
        if candidate_path.stat().st_size > 0 and candidate_path.exists() and score > best_score:
            best_path = candidate_path
            best_label = label
            best_score = score
        if label:
            return candidate_path, label

    if best_path is None:
        fallback = image.rotate(90, expand=True) if image.height > image.width else image.copy()
        fallback_path = temp_dir / f"{stem}_normalized{normalized_suffix}"
        save_candidate(fallback, fallback_path, image_format)
        return fallback_path, None
    return best_path, best_label


def image_info_from_path(path: Path, explicit_type: str | None) -> generate_id_doc.ImageInfo | None:
    try:
        with Image.open(path) as image:
            image = ImageOps.exif_transpose(image)
            width, height = image.size
    except Exception:
        return None
    if width <= 0 or height <= 0:
        return None
    return generate_id_doc.ImageInfo(
        path=path,
        width_px=width,
        height_px=height,
        ratio=width / height,
        area=width * height,
        explicit_type=explicit_type,
    )


def arrange_normalized_images(images: list[generate_id_doc.ImageInfo]) -> list[tuple[str, generate_id_doc.ImageInfo]]:
    arranged = generate_id_doc.arrange_images(images)
    return sorted(arranged, key=lambda item: (CONTENT_ORDER.get(item[0], 99), item[1].path.name.lower()))


def organize_master_document(cert_dir: Path, progress=None) -> MasterDocResult:
    source = cert_dir / MASTER_DOC_NAME
    if not source.exists():
        source = cert_dir / MASTER_DOCX_NAME
    if not source.exists():
        raise FileNotFoundError(f"没有找到总证件文档：{cert_dir / MASTER_DOC_NAME}")

    output_root = cert_dir / OUTPUT_DIR_NAME
    cache_dir = output_root / CACHE_DIR_NAME
    if progress:
        progress(5, "正在准备总证件文档")
    source_docx = convert_doc_to_docx(source, cache_dir)

    if progress:
        progress(18, "正在读取总文档里的人名和图片")
    groups = parse_groups(source_docx)
    if not groups:
        raise RuntimeError("总证件文档里没有识别到可整理的人名和图片。")

    output_root.mkdir(parents=True, exist_ok=True)
    staging_root = output_root / STAGING_DIR_NAME
    backup_root = output_root / BACKUP_DIR_NAME
    for temp_root in (staging_root, backup_root):
        if temp_root.exists():
            shutil.rmtree(temp_root, ignore_errors=True)
    staging_root.mkdir(parents=True, exist_ok=True)

    people_count = 0
    image_count = 0
    skipped_count = 0
    used_names: set[str] = set()
    normalized_root = cache_dir / "_旋正识别"
    if normalized_root.exists():
        shutil.rmtree(normalized_root, ignore_errors=True)
    normalized_root.mkdir(parents=True, exist_ok=True)

    with ZipFile(source_docx) as zip_file:
        targets = relationship_targets(zip_file)
        total = len(groups)
        for group_index, (name, image_ids) in enumerate(groups, 1):
            person_dir = unique_folder(staging_root, name, used_names)
            person_dir.mkdir(parents=True, exist_ok=True)
            temp_person_dir = normalized_root / f"{group_index:03d}_{safe_name(name)}"
            temp_person_dir.mkdir(parents=True, exist_ok=True)
            normalized_images: list[generate_id_doc.ImageInfo] = []
            wrote = 0
            for image_index, rid in enumerate(image_ids):
                target = targets.get(rid)
                if not target:
                    skipped_count += 1
                    continue
                archive_path = image_path_from_target(target)
                try:
                    data = zip_file.read(archive_path)
                except KeyError:
                    skipped_count += 1
                    continue
                suffix = Path(archive_path).suffix.lower() or ".jpg"
                if suffix not in DOC_IMAGE_SUFFIXES:
                    skipped_count += 1
                    continue
                try:
                    normalized_path, detected_label = normalize_and_classify_image(
                        data,
                        suffix,
                        temp_person_dir,
                        f"{image_index + 1:02d}",
                    )
                    image_info = image_info_from_path(normalized_path, detected_label)
                except Exception:
                    skipped_count += 1
                    continue
                if image_info is None:
                    skipped_count += 1
                    continue
                normalized_images.append(image_info)

            arranged_images = arrange_normalized_images(normalized_images)
            for arranged_index, (label, image_info) in enumerate(arranged_images, 1):
                suffix = image_info.path.suffix.lower() or ".jpg"
                image_name = f"{arranged_index:02d}_{label}{suffix}"
                shutil.copyfile(image_info.path, person_dir / image_name)
                wrote += 1
                image_count += 1
            if wrote:
                people_count += 1
            else:
                shutil.rmtree(person_dir, ignore_errors=True)
            if progress:
                value = 20 + int(group_index / max(total, 1) * 75)
                progress(value, f"正在整理：{name}")

    if people_count == 0 or image_count == 0:
        shutil.rmtree(staging_root, ignore_errors=True)
        raise RuntimeError("没有成功整理出任何证件图片，旧整理结果已保留。")

    if progress:
        progress(96, "正在替换旧整理结果")
    backup_root.mkdir(parents=True, exist_ok=True)
    moved_new: list[Path] = []
    try:
        for child in output_root.iterdir():
            if child.name in {CACHE_DIR_NAME, STAGING_DIR_NAME, BACKUP_DIR_NAME}:
                continue
            shutil.move(str(child), str(backup_root / child.name))
        for child in staging_root.iterdir():
            target = output_root / child.name
            shutil.move(str(child), str(target))
            moved_new.append(target)
        shutil.rmtree(staging_root, ignore_errors=True)
        shutil.rmtree(backup_root, ignore_errors=True)
    except Exception:
        for child in moved_new:
            if child.exists():
                if child.is_dir():
                    shutil.rmtree(child, ignore_errors=True)
                else:
                    child.unlink(missing_ok=True)
        if backup_root.exists():
            for child in backup_root.iterdir():
                target = output_root / child.name
                if not target.exists():
                    shutil.move(str(child), str(target))
        raise

    return MasterDocResult(
        output_root=output_root,
        people_count=people_count,
        image_count=image_count,
        skipped_count=skipped_count,
        source_docx=source_docx,
    )


if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    result = organize_master_document(base)
    print(f"已整理 {result.people_count} 人，{result.image_count} 张图片")
    print(result.output_root)
