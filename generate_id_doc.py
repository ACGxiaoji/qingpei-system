# -*- coding: utf-8 -*-
"""
把“证件提取”目录下各个人名文件夹里的证件图片，批量生成一份 Word。

使用方法：
1. 把每个人的身份证、银行卡图片放到对应的人名文件夹里。
2. 双击运行本脚本，或在命令行执行：python 生成青赔证件Word.py
3. 每个人名文件夹里会生成：青赔身份证、银行卡.docx

说明：
- 文件名里包含“身份证正面/国徽面/背面/反面/front”等，会优先排到身份证正面。
- 文件名里包含“身份证反面/人像面/头像面/back”等，会优先排到身份证反面。
- 文件名里包含“银行卡/银行/bank/储蓄卡”等，会优先排到银行卡。
- 如果文件名无法识别，脚本会按图片比例和顺序自动排版。
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import tempfile
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
    from PIL import Image, ImageOps
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ImportError as exc:
    print("缺少必要组件，无法生成 Word。")
    print("请先安装组件：pip install python-docx pillow")
    print(f"具体缺少：{exc}")
    input("按回车键退出...")
    raise SystemExit(1)


# ===== 可按需要调整的设置 =====
# 输出文件名使用人名文件夹命名，例如：林泉.docx
OUTPUT_FILENAME_TEMPLATE = "{name}.docx"
OLD_OUTPUT_FILENAME = "青赔身份证、银行卡.docx"
TITLE_TEXT = "青赔身份证、银行卡"

# 页面设置
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGIN_CM = 1.6

# 图片显示大小：三张证件图通常可以放在一页内。
MAX_IMAGE_WIDTH_CM = 13.2
MAX_IMAGE_HEIGHT_CM = 6.6

# 为了避免生成的 Word 过大，图片会在嵌入前等比压缩到这个像素以内。
MAX_IMAGE_PIXEL_EDGE = 2200

# 是否覆盖已生成的 Word。
OVERWRITE_EXISTING = True

# 支持的图片格式。
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
_OCR_ENGINE = None


FRONT_KEYWORDS = (
    "身份证正面",
    "国徽面",
    "正面",
    "front",
    "id_front",
    "sfz_zm",
)
BACK_KEYWORDS = (
    "身份证反面",
    "人像面",
    "头像面",
    "背面",
    "反面",
    "back",
    "id_back",
    "sfz_fm",
)
BANK_KEYWORDS = (
    "银行卡",
    "银行",
    "开户",
    "储蓄卡",
    "借记卡",
    "bank",
    "bankcard",
)


@dataclass
class ImageInfo:
    path: Path
    width_px: int
    height_px: int
    ratio: float
    area: int
    explicit_type: str | None = None


def normalize_name(text: str) -> str:
    return re.sub(r"[\s_\-（）()【】\[\].]+", "", text.lower())


def classify_by_filename(path: Path) -> str | None:
    name = normalize_name(path.stem)
    if any(normalize_name(k) in name for k in BANK_KEYWORDS):
        return "银行卡"
    if any(normalize_name(k) in name for k in BACK_KEYWORDS):
        return "身份证反面"
    if any(normalize_name(k) in name for k in FRONT_KEYWORDS):
        return "身份证正面"
    return None


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


def read_ocr_text(path: Path) -> str:
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
        if len(item) >= 2:
            lines.append(str(item[1]))
    return "\n".join(lines)


def classify_by_content(path: Path) -> str | None:
    text = read_ocr_text(path)
    if not text:
        return None
    compact = re.sub(r"\s+", "", text).lower()

    bank_words = (
        "银行卡",
        "银行",
        "unionpay",
        "bank",
        "debit",
        "银联",
        "储蓄卡",
        "借记卡",
        "信用社",
        "农商",
        "金穗",
    )
    if any(word.lower() in compact for word in bank_words):
        return "银行卡"

    front_words = ("中华人民共和国", "居民身份证", "签发机关", "有效期限")
    if any(word in compact for word in front_words):
        return "身份证正面"

    back_words = ("姓名", "性别", "民族", "出生", "住址", "公民身份号码")
    if any(word in compact for word in back_words) or re.search(r"\d{17}[\dx]", compact):
        return "身份证反面"

    return None


def read_image_info(path: Path) -> ImageInfo | None:
    try:
        with Image.open(path) as image:
            image = ImageOps.exif_transpose(image)
            width, height = image.size
    except Exception as exc:
        print(f"  跳过无法读取的图片：{path.name}（{exc}）")
        return None

    if width <= 0 or height <= 0:
        print(f"  跳过尺寸异常的图片：{path.name}")
        return None

    return ImageInfo(
        path=path,
        width_px=width,
        height_px=height,
        ratio=width / height,
        area=width * height,
        explicit_type=classify_by_content(path) or classify_by_filename(path),
    )


def collect_images(person_dir: Path) -> list[ImageInfo]:
    images: list[ImageInfo] = []
    for path in sorted(person_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        info = read_image_info(path)
        if info is not None:
            images.append(info)
    return images


def arrange_images(images: list[ImageInfo]) -> list[tuple[str, ImageInfo]]:
    buckets: dict[str, list[ImageInfo]] = {
        "身份证正面": [],
        "身份证反面": [],
        "银行卡": [],
    }
    unknown: list[ImageInfo] = []

    for image in images:
        if image.explicit_type in buckets:
            buckets[image.explicit_type].append(image)
        else:
            unknown.append(image)

    arranged: list[tuple[str, ImageInfo]] = []

    for label in ("身份证正面", "身份证反面", "银行卡"):
        if buckets[label]:
            arranged.append((label, buckets[label].pop(0)))

    for label in ("身份证正面", "身份证反面", "银行卡"):
        if label not in {item[0] for item in arranged} and unknown:
            arranged.append((label, unknown.pop(0)))

    extra_index = 1
    for image in unknown:
        arranged.append((f"其他证件{extra_index}", image))
        extra_index += 1

    for label in ("身份证正面", "身份证反面", "银行卡"):
        for image in buckets[label]:
            arranged.append((label, image))

    return arranged


def set_run_font(run, size_pt: float = 11, bold: bool = False) -> None:
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size_pt)
    run.bold = bold


def add_center_text(document: Document, text: str, size_pt: float, bold: bool = False, after_pt: float = 4) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(after_pt)
    run = paragraph.add_run(text)
    set_run_font(run, size_pt=size_pt, bold=bold)


def prepare_image_for_word(info: ImageInfo, temp_dir: Path) -> tuple[Path, int, int]:
    with Image.open(info.path) as image:
        image = ImageOps.exif_transpose(image)
        if image.height > image.width:
            image = image.rotate(90, expand=True)
        image.thumbnail((MAX_IMAGE_PIXEL_EDGE, MAX_IMAGE_PIXEL_EDGE), Image.Resampling.LANCZOS)

        width, height = image.size
        safe_stem = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", info.path.stem).strip("_") or "image"

        if image.mode in ("RGBA", "LA") or info.path.suffix.lower() in {".png", ".gif", ".webp"}:
            if image.mode not in ("RGB", "RGBA"):
                image = image.convert("RGBA")
            out_path = temp_dir / f"{safe_stem}.png"
            image.save(out_path, "PNG")
        else:
            if image.mode != "RGB":
                image = image.convert("RGB")
            out_path = temp_dir / f"{safe_stem}.jpg"
            image.save(out_path, "JPEG", quality=92, optimize=True)

    return out_path, width, height


def display_size_cm(width_px: int, height_px: int) -> tuple[float, float]:
    width_cm = MAX_IMAGE_WIDTH_CM
    height_cm = width_cm * height_px / width_px

    if height_cm > MAX_IMAGE_HEIGHT_CM:
        height_cm = MAX_IMAGE_HEIGHT_CM
        width_cm = height_cm * width_px / height_px

    return width_cm, height_cm


def build_document(person_name: str, arranged: list[tuple[str, ImageInfo]], output_path: Path) -> None:
    document = Document()

    section = document.sections[0]
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.top_margin = Cm(MARGIN_CM)
    section.bottom_margin = Cm(MARGIN_CM)
    section.left_margin = Cm(MARGIN_CM)
    section.right_margin = Cm(MARGIN_CM)

    normal = document.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(11)

    add_center_text(document, TITLE_TEXT, size_pt=18, bold=True, after_pt=2)
    add_center_text(document, f"姓名：{person_name}", size_pt=12, bold=False, after_pt=8)

    with tempfile.TemporaryDirectory(prefix="qp_doc_images_") as temp_root:
        temp_dir = Path(temp_root)
        for index, (label, image_info) in enumerate(arranged):
            if index:
                spacer = document.add_paragraph()
                spacer.paragraph_format.space_after = Pt(2)

            add_center_text(document, label, size_pt=11, bold=True, after_pt=2)
            prepared_path, width_px, height_px = prepare_image_for_word(image_info, temp_dir)
            width_cm, height_cm = display_size_cm(width_px, height_px)

            paragraph = document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_after = Pt(4)
            run = paragraph.add_run()
            run.add_picture(str(prepared_path), width=Cm(width_cm), height=Cm(height_cm))

    document.save(output_path)


def person_directories(base_dir: Path) -> list[Path]:
    ignored = {"__pycache__"}
    dirs: list[Path] = []
    for path in sorted(base_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_dir():
            continue
        if path.name in ignored or path.name.startswith("_") or path.name.startswith("."):
            continue
        dirs.append(path)
    return dirs


def safe_output_filename(person_name: str) -> str:
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", person_name).strip(" .")
    if not safe_name:
        safe_name = "证件"
    return OUTPUT_FILENAME_TEMPLATE.format(name=safe_name)


def run(base_dir: Path) -> int:
    print(f"处理目录：{base_dir}")
    dirs = person_directories(base_dir)
    if not dirs:
        print("没有找到人名文件夹。")
        return 1

    success = 0
    skipped = 0

    for person_dir in dirs:
        print(f"\n开始处理：{person_dir.name}")
        output_path = person_dir / safe_output_filename(person_dir.name)

        if output_path.exists() and not OVERWRITE_EXISTING:
            print(f"  已存在，跳过：{output_path.name}")
            skipped += 1
            continue

        images = collect_images(person_dir)
        if not images:
            print("  没有找到图片，跳过。")
            skipped += 1
            continue

        arranged = arrange_images(images)
        try:
            build_document(person_dir.name, arranged, output_path)
        except PermissionError:
            print(f"  生成失败：{output_path.name} 可能正被 Word 打开，请关闭后重试。")
            skipped += 1
            continue
        except Exception as exc:
            print(f"  生成失败：{exc}")
            skipped += 1
            continue

        print(f"  已生成：{output_path.name}")
        old_output_path = person_dir / OLD_OUTPUT_FILENAME
        if old_output_path.exists() and old_output_path != output_path:
            try:
                old_output_path.unlink()
                print(f"  已清理旧命名文件：{old_output_path.name}")
            except PermissionError:
                print(f"  旧命名文件未清理：{old_output_path.name} 可能正被 Word 打开。")
        for label, image in arranged:
            print(f"    {label}：{image.path.name}")
        success += 1

    print(f"\n完成：成功 {success} 个，跳过/失败 {skipped} 个。")
    return 0 if success else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="批量生成青赔身份证、银行卡 Word")
    parser.add_argument("--base-dir", type=Path, default=Path(__file__).resolve().parent, help="证件提取目录")
    parser.add_argument("--no-pause", action="store_true", help="运行结束后不等待回车")
    args = parser.parse_args()

    code = run(args.base_dir)
    if not args.no_pause:
        input("\n按回车键退出...")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
