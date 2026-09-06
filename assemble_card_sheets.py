#!/usr/bin/env python3
"""
Assemble card-front images onto landscape US Letter PDF sheets, extract a
card-back image from a source PDF, and generate matching duplex back pages.

Default layout:
  - Paper: 11 x 8.5 in (landscape US Letter)
  - Final card image: 63 x 88 mm, upright
  - Black bleed: 3 mm embedded inside each side of the 63 x 88 mm image
  - Card spacing: approximately 0.127 mm
  - Result: 4 x 2 cards per sheet, centered

The visible outer black boundary remains 63 x 88 mm rather than becoming
69 x 94 mm.

Each batch of eight cards is written as a separate output. Card backs are
included by default; use --single-sided (or --no-card-backs) to generate only
front pages. For example, 22 selected cards produce OutputMerge.pdf,
OutputMerge_2.pdf, and OutputMerge_3.pdf.

Card selection is controlled by cards.txt. Example:
  4 Beza, the Bounding Spring (BLB) 287
  2x Elves of Deep Shadow (RAV) 161
  x1 Sol Ring (CMM) 410
  Ancient Den (EOC) 148

The leading quantity is optional and may use 4, 4x, or x4; omitted quantities
default to 1. Only the quantity and card name are used; the set code and
collector number are ignored.
Image filenames may either match the card name exactly or start with
it, followed by metadata such as "(Normal) [BLB] {287}".

Duplex alignment defaults to automatic long-edge binding. On a landscape
sheet, long-edge duplex flips the paper vertically, so back-page card slots are
mirrored vertically. Use --duplex-binding short-edge for horizontal mirroring,
or --duplex-binding none if you manually reinsert paper without an automatic
flip.

By default, the script looks for cards.txt, images, and
horizontal_cardBack_3x3.pdf in the same folder as this script. It writes
merge1.pdf, merge1_backs.pdf, OutputMerge.pdf, and numbered batch files there
as needed. In single-sided mode, no card-back PDF is required and no
merge1_backs.pdf files are created.
"""

from __future__ import annotations

import argparse
import math
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable

try:
    import numpy as np
    from PIL import Image, ImageOps
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
except ImportError as exc:  # pragma: no cover - depends on user's environment
    missing = getattr(exc, "name", None) or str(exc)
    raise SystemExit(
        f"Missing required Python package: {missing}\n"
        "Install the dependencies with:\n"
        "  python -m pip install pillow numpy pypdf reportlab"
    ) from exc


BACK_PDF_NAME = "horizontal_cardBack_3x3.pdf"
CARDS_TXT_NAME = "cards.txt"
FRONTS_PDF_NAME = "merge1.pdf"
OUTPUT_PDF_NAME = "OutputMerge.pdf"

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
    ".bmp",
    ".tif",
    ".tiff",
}

# The leading quantity is optional and defaults to 1. Accept common deck-list
# forms such as "4 Card", "4x Card", and "x4 Card". The optional trailing group
# consumes and ignores " (SET) collector-number".
CARD_LIST_LINE_RE = re.compile(
    r"^\s*(?:(?P<count>\d+)x?\s+|x(?P<count_prefix>\d+)\s+)?"
    r"(?P<name>.*?)(?:\s+\([^()]*\)\s+\S+)?\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Layout:
    paper_width_mm: float
    paper_height_mm: float
    card_width_mm: float
    card_height_mm: float
    bleed_mm: float
    margin_mm: float
    columns: int
    rows: int
    page_margin_x_mm: float
    page_margin_y_mm: float

    @property
    def cards_per_page(self) -> int:
        return self.columns * self.rows

    @property
    def slot_width_mm(self) -> float:
        # Bleed is drawn into the spacing and does not consume grid capacity.
        return self.card_width_mm + self.margin_mm

    @property
    def slot_height_mm(self) -> float:
        # Bleed is drawn into the spacing and does not consume grid capacity.
        return self.card_height_mm + self.margin_mm


@dataclass(frozen=True)
class CardRequest:
    name: str
    quantity: int
    line_number: int


def natural_key(value: Path | str) -> list[object]:
    """Sort like Windows Explorer: card2 before card10."""
    text = str(value).casefold()
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", text)]


def calculate_layout(
    paper_width_mm: float,
    paper_height_mm: float,
    card_width_mm: float,
    card_height_mm: float,
    bleed_mm: float,
    margin_mm: float,
) -> Layout:
    """
    Calculate the card grid from the cut-card size plus spacing.

    The black bleed is drawn outward from each card into the spacing instead of
    increasing the grid pitch. With 63 x 88 mm cards and 3 mm spacing on
    landscape letter, this preserves the expected 4 x 2 arrangement.
    """
    slot_width = card_width_mm + margin_mm
    slot_height = card_height_mm + margin_mm

    columns = math.floor(paper_width_mm / slot_width)
    rows = math.floor(paper_height_mm / slot_height)

    if columns < 1 or rows < 1:
        raise ValueError(
            "No complete card fits on the requested paper size. "
            f"Slot size is {slot_width:.3f} x {slot_height:.3f} mm; "
            f"paper is {paper_width_mm:.3f} x {paper_height_mm:.3f} mm."
        )

    page_margin_x = (paper_width_mm - (columns * slot_width) + margin_mm) / 2
    page_margin_y = (paper_height_mm - (rows * slot_height) + margin_mm) / 2

    if (2 * bleed_mm) >= card_width_mm or (2 * bleed_mm) >= card_height_mm:
        raise ValueError(
            "Bleed must be smaller than half the card dimensions so artwork "
            "remains inside the black border."
        )

    return Layout(
        paper_width_mm=paper_width_mm,
        paper_height_mm=paper_height_mm,
        card_width_mm=card_width_mm,
        card_height_mm=card_height_mm,
        bleed_mm=bleed_mm,
        margin_mm=margin_mm,
        columns=columns,
        rows=rows,
        page_margin_x_mm=page_margin_x,
        page_margin_y_mm=page_margin_y,
    )


def collect_images(input_folder: Path, recursive: bool) -> list[Path]:
    candidates: Iterable[Path] = input_folder.rglob("*") if recursive else input_folder.iterdir()
    images = [
        path
        for path in candidates
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return sorted(
        images,
        key=lambda path: natural_key(path.relative_to(input_folder)),
    )


def normalize_card_name(value: str) -> str:
    """Normalize names for case-insensitive, whitespace-tolerant matching."""
    normalized = unicodedata.normalize("NFKC", value).casefold()
    normalized = normalized.replace("_", " ")
    return re.sub(r"\s+", " ", normalized).strip()


def filename_matches_card_name(image_path: Path, card_name: str) -> bool:
    """
    Match an exact card-name stem or a stem that starts with the card name and
    then continues with metadata, such as:
      Beza, the Bounding Spring (Normal) [BLB] {287}.jpg
    """
    stem = normalize_card_name(image_path.stem)
    name = normalize_card_name(card_name)

    if stem == name:
        return True
    if not stem.startswith(name):
        return False

    remainder = stem[len(name) :]
    return remainder.startswith((" ", "(", "[", "{", "-"))


def parse_cards_txt(cards_txt: Path) -> tuple[list[CardRequest], list[str]]:
    requests: list[CardRequest] = []
    errors: list[str] = []

    with cards_txt.open("r", encoding="utf-8-sig") as cards_file:
        for line_number, line in enumerate(cards_file, start=1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue

            match = CARD_LIST_LINE_RE.match(line)
            if not match:
                errors.append(
                    f"ERROR: Could not parse cards.txt line {line_number}: {stripped}"
                )
                continue

            quantity_text = match.group("count") or match.group("count_prefix")
            quantity = int(quantity_text) if quantity_text else 1
            name = match.group("name").strip()
            if quantity < 1:
                errors.append(
                    f"ERROR: Quantity must be at least 1 on cards.txt line "
                    f"{line_number}: {stripped}"
                )
                continue
            if not name:
                errors.append(
                    f"ERROR: Missing card name on cards.txt line {line_number}: {stripped}"
                )
                continue

            requests.append(
                CardRequest(name=name, quantity=quantity, line_number=line_number)
            )

    return requests, errors


def select_images_from_cards_txt(
    requests: list[CardRequest], image_paths: list[Path]
) -> tuple[list[Path], list[CardRequest], int]:
    """
    Expand cards.txt quantities into the ordered image list sent to the PDF.
    Returns (selected copies, missing requests, ignored image count).
    """
    selected: list[Path] = []
    missing: list[CardRequest] = []
    used_images: set[Path] = set()

    for request in requests:
        matches = [
            image_path
            for image_path in image_paths
            if filename_matches_card_name(image_path, request.name)
        ]

        if not matches:
            missing.append(request)
            continue

        requested_name = normalize_card_name(request.name)
        exact_matches = [
            image_path
            for image_path in matches
            if normalize_card_name(image_path.stem) == requested_name
        ]
        selected_image = exact_matches[0] if exact_matches else matches[0]

        if len(matches) > 1:
            print(
                f"WARNING: {len(matches)} image files match "
                f'"{request.name}"; using "{selected_image.name}".',
                file=sys.stderr,
            )

        used_images.add(selected_image.resolve())
        selected.extend([selected_image] * request.quantity)

    ignored_count = sum(
        1 for image_path in image_paths if image_path.resolve() not in used_images
    )
    return selected, missing, ignored_count


def auto_crop_black_canvas(
    image: Image.Image,
    threshold: int = 20,
    edge_fraction: float = 0.005,
    minimum_trim_fraction: float = 0.08,
    padding_px: int = 2,
) -> tuple[Image.Image, bool]:
    """
    Remove a substantial black scanner/canvas border while avoiding normal
    black-bordered card images.

    A crop is applied only when both dimensions shrink by at least
    minimum_trim_fraction. The sample scan trims about 13% horizontally and
    11% vertically; an ordinary card's printed black border trims much less.
    """
    rgb = np.asarray(image.convert("RGB"))
    nonblack = rgb.max(axis=2) > threshold

    row_has_content = nonblack.mean(axis=1) > edge_fraction
    column_has_content = nonblack.mean(axis=0) > edge_fraction

    content_rows = np.flatnonzero(row_has_content)
    content_columns = np.flatnonzero(column_has_content)
    if content_rows.size == 0 or content_columns.size == 0:
        return image, False

    height, width = nonblack.shape
    left = max(0, int(content_columns[0]) - padding_px)
    top = max(0, int(content_rows[0]) - padding_px)
    right = min(width, int(content_columns[-1]) + 1 + padding_px)
    bottom = min(height, int(content_rows[-1]) + 1 + padding_px)

    horizontal_trim = (left + (width - right)) / width
    vertical_trim = (top + (height - bottom)) / height

    if horizontal_trim >= minimum_trim_fraction and vertical_trim >= minimum_trim_fraction:
        return image.crop((left, top, right, bottom)), True

    return image, False


def prepare_card_image(
    path: Path,
    card_width_mm: float,
    card_height_mm: float,
    bleed_mm: float,
    dpi: int,
    jpeg_quality: int,
    auto_crop: bool,
) -> tuple[ImageReader, BytesIO, bool]:
    """Crop, orient, resize, add black bleed, and JPEG-encode one card."""
    with Image.open(path) as opened:
        image = ImageOps.exif_transpose(opened)
        image = image.convert("RGB")
        was_cropped = False

        if auto_crop:
            image, was_cropped = auto_crop_black_canvas(image)

        final_width_px = max(1, round(card_width_mm / 25.4 * dpi))
        final_height_px = max(1, round(card_height_mm / 25.4 * dpi))
        bleed_px = max(0, round(bleed_mm / 25.4 * dpi))

        # Match the original PDF's convention: the final image is 63 x 88 mm,
        # and the 3 mm bleed is embedded inside that image as a black border.
        artwork_width_px = max(1, final_width_px - (2 * bleed_px))
        artwork_height_px = max(1, final_height_px - (2 * bleed_px))
        if image.size != (artwork_width_px, artwork_height_px):
            image = image.resize(
                (artwork_width_px, artwork_height_px),
                Image.Resampling.LANCZOS,
            )

        if bleed_px > 0:
            image_with_bleed = Image.new(
                "RGB",
                (final_width_px, final_height_px),
                (0, 0, 0),
            )
            image_with_bleed.paste(image, (bleed_px, bleed_px))
            image = image_with_bleed

        image_bytes = BytesIO()
        image.save(
            image_bytes,
            format="JPEG",
            quality=jpeg_quality,
            subsampling=0,
            dpi=(dpi, dpi),
        )
        image_bytes.seek(0)
        return ImageReader(image_bytes), image_bytes, was_cropped


def extract_card_back_image(
    back_pdf: Path,
    card_width_mm: float,
    card_height_mm: float,
) -> Image.Image:
    """Extract the image whose aspect ratio best matches one upright card back."""
    expected_ratio = card_width_mm / card_height_mm
    best_image: Image.Image | None = None
    best_score: float | None = None

    reader = PdfReader(str(back_pdf))
    for page in reader.pages:
        try:
            page_images = list(page.images)
        except Exception:
            continue

        for extracted_image in page_images:
            try:
                image = extracted_image.image
            except Exception:
                try:
                    image = Image.open(BytesIO(extracted_image.data))
                except Exception:
                    continue

            image = ImageOps.exif_transpose(image).convert("RGB").copy()
            if image.width == 0 or image.height == 0:
                continue

            ratio = image.width / image.height
            score = abs(ratio - expected_ratio)
            if best_score is None or score < best_score:
                best_image = image
                best_score = score

    if best_image is None or best_score is None or best_score > 0.08:
        raise ValueError(
            f"Could not find an upright card-sized image inside {back_pdf}. "
            "The back PDF must contain individual card-back images."
        )

    return best_image


def prepare_back_image(
    image: Image.Image,
    card_width_mm: float,
    card_height_mm: float,
    dpi: int,
    jpeg_quality: int,
) -> tuple[ImageReader, BytesIO]:
    """Resize and encode a card back without adding another bleed border."""
    image = image.convert("RGB")
    target_size = (
        max(1, round(card_width_mm / 25.4 * dpi)),
        max(1, round(card_height_mm / 25.4 * dpi)),
    )
    if image.size != target_size:
        image = image.resize(target_size, Image.Resampling.LANCZOS)

    image_bytes = BytesIO()
    image.save(
        image_bytes,
        format="JPEG",
        quality=jpeg_quality,
        subsampling=0,
        dpi=(dpi, dpi),
    )
    image_bytes.seek(0)
    return ImageReader(image_bytes), image_bytes


def draw_cutting_guides(pdf: canvas.Canvas, layout: Layout) -> None:
    """Draw thin full-page card-boundary guides behind the card images."""
    page_width_pt = layout.paper_width_mm * mm
    page_height_pt = layout.paper_height_mm * mm
    card_width_pt = layout.card_width_mm * mm
    card_height_pt = layout.card_height_mm * mm

    pdf.saveState()
    pdf.setStrokeColorRGB(0, 0, 0)
    pdf.setLineWidth(0.2)

    for column in range(layout.columns):
        left_pt = (layout.page_margin_x_mm + (column * layout.slot_width_mm)) * mm
        right_pt = left_pt + card_width_pt
        pdf.line(left_pt, 0, left_pt, page_height_pt)
        pdf.line(right_pt, 0, right_pt, page_height_pt)

    for row in range(layout.rows):
        top_pt = page_height_pt - (
            (layout.page_margin_y_mm + (row * layout.slot_height_mm)) * mm
        )
        bottom_pt = top_pt - card_height_pt
        pdf.line(0, top_pt, page_width_pt, top_pt)
        pdf.line(0, bottom_pt, page_width_pt, bottom_pt)

    pdf.restoreState()


def build_fronts_pdf(
    image_paths: list[Path],
    output_pdf: Path,
    layout: Layout,
    dpi: int,
    jpeg_quality: int,
    auto_crop: bool,
    report_auto_crop: bool = True,
    cut_guides: bool = True,
) -> int:
    """Create the landscape front-card PDF. Returns its page count."""
    page_width_pt = layout.paper_width_mm * mm
    page_height_pt = layout.paper_height_mm * mm
    image_width_pt = layout.card_width_mm * mm
    image_height_pt = layout.card_height_mm * mm

    pdf = canvas.Canvas(
        str(output_pdf),
        pagesize=(page_width_pt, page_height_pt),
        pageCompression=1,
    )
    pdf.setTitle("Card Front Sheets")
    pdf.setSubject(
        f"{layout.columns} x {layout.rows} upright cards on landscape US Letter"
    )
    if cut_guides:
        draw_cutting_guides(pdf, layout)

    # Keep buffers alive until the PDF is saved; ReportLab stores image data by reference.
    image_buffers: list[BytesIO] = []
    prepared_cache: dict[Path, tuple[ImageReader, BytesIO, bool]] = {}
    cropped_count = 0

    for index, image_path in enumerate(image_paths):
        slot_index = index % layout.cards_per_page
        if index > 0 and slot_index == 0:
            pdf.showPage()
            if cut_guides:
                draw_cutting_guides(pdf, layout)

        column = slot_index % layout.columns
        row = slot_index // layout.columns

        # The prepared image already contains its black bleed border and is
        # placed at the final 63 x 88 mm size, matching the original PDF.
        image_left_mm = layout.page_margin_x_mm + (column * layout.slot_width_mm)
        image_top_mm = layout.page_margin_y_mm + (row * layout.slot_height_mm)
        image_left_pt = image_left_mm * mm
        image_bottom_pt = page_height_pt - (image_top_mm * mm) - image_height_pt

        cache_key = image_path.resolve()
        if cache_key not in prepared_cache:
            prepared_cache[cache_key] = prepare_card_image(
                image_path,
                card_width_mm=layout.card_width_mm,
                card_height_mm=layout.card_height_mm,
                bleed_mm=layout.bleed_mm,
                dpi=dpi,
                jpeg_quality=jpeg_quality,
                auto_crop=auto_crop,
            )
            image_reader, image_buffer, was_cropped = prepared_cache[cache_key]
            image_buffers.append(image_buffer)
            cropped_count += int(was_cropped)
        else:
            image_reader, _, _ = prepared_cache[cache_key]

        pdf.drawImage(
            image_reader,
            image_left_pt,
            image_bottom_pt,
            width=image_width_pt,
            height=image_height_pt,
            preserveAspectRatio=False,
            mask="auto",
        )

    pdf.save()

    if auto_crop and report_auto_crop:
        print(
            "Auto-cropped black scanner canvas from "
            f"{cropped_count} unique image(s)."
        )

    return math.ceil(len(image_paths) / layout.cards_per_page)


def get_back_slot(
    column: int,
    row: int,
    layout: Layout,
    duplex_binding: str,
) -> tuple[int, int]:
    """Map a front slot to the physical slot reached by a duplex paper flip."""
    if duplex_binding == "long-edge":
        # Landscape paper flipped around its long horizontal edge maps top to
        # bottom while preserving the left/right column.
        return column, (layout.rows - 1 - row)
    if duplex_binding == "short-edge":
        # Landscape paper flipped around its short vertical edge maps left to
        # right while preserving the top/bottom row.
        return (layout.columns - 1 - column), row
    return column, row


def build_backs_pdf(
    back_image: Image.Image,
    card_count: int,
    output_pdf: Path,
    layout: Layout,
    dpi: int,
    jpeg_quality: int,
    cut_guides: bool = True,
    duplex_binding: str = "long-edge",
) -> int:
    """Create back pages whose slots compensate for automatic duplex flipping."""
    page_width_pt = layout.paper_width_mm * mm
    page_height_pt = layout.paper_height_mm * mm
    image_width_pt = layout.card_width_mm * mm
    image_height_pt = layout.card_height_mm * mm

    pdf = canvas.Canvas(
        str(output_pdf),
        pagesize=(page_width_pt, page_height_pt),
        pageCompression=1,
    )
    pdf.setTitle("Card Back Sheets")
    pdf.setSubject(
        f"{layout.columns} x {layout.rows} upright card backs on landscape US Letter"
    )
    if cut_guides:
        draw_cutting_guides(pdf, layout)

    back_reader, back_buffer = prepare_back_image(
        back_image,
        card_width_mm=layout.card_width_mm,
        card_height_mm=layout.card_height_mm,
        dpi=dpi,
        jpeg_quality=jpeg_quality,
    )

    for index in range(card_count):
        slot_index = index % layout.cards_per_page
        if index > 0 and slot_index == 0:
            pdf.showPage()
            if cut_guides:
                draw_cutting_guides(pdf, layout)

        front_column = slot_index % layout.columns
        front_row = slot_index // layout.columns
        column, row = get_back_slot(
            front_column,
            front_row,
            layout,
            duplex_binding,
        )
        image_left_pt = (
            layout.page_margin_x_mm + (column * layout.slot_width_mm)
        ) * mm
        image_top_mm = layout.page_margin_y_mm + (row * layout.slot_height_mm)
        image_bottom_pt = page_height_pt - (image_top_mm * mm) - image_height_pt

        pdf.drawImage(
            back_reader,
            image_left_pt,
            image_bottom_pt,
            width=image_width_pt,
            height=image_height_pt,
            preserveAspectRatio=False,
            mask="auto",
        )

    # Keep the source buffer alive until the document is complete.
    pdf.save()
    back_buffer.close()
    return math.ceil(card_count / layout.cards_per_page)


def read_pdf_pages(path: Path) -> list[object]:
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise ValueError(f"PDF is encrypted and could not be opened: {path}") from exc
    return list(reader.pages)


def merge_pdfs(fronts_pdf: Path, back_pdf: Path, output_pdf: Path) -> tuple[int, int, int]:
    """Append the back PDF after the fronts PDF, matching Merge-PDF behavior."""
    fronts_pages = read_pdf_pages(fronts_pdf)
    back_pages = read_pdf_pages(back_pdf)

    writer = PdfWriter()
    for page in fronts_pages:
        writer.add_page(page)
    for page in back_pages:
        writer.add_page(page)

    with output_pdf.open("wb") as output_stream:
        writer.write(output_stream)

    return len(fronts_pages), len(back_pages), len(fronts_pages) + len(back_pages)


def chunk_items(items: list[Path], chunk_size: int) -> list[list[Path]]:
    """Split selected card copies into one-grid batches."""
    return [items[index : index + chunk_size] for index in range(0, len(items), chunk_size)]


def numbered_batch_path(base_path: Path, batch_number: int) -> Path:
    """Keep the first batch's existing name, then append _2, _3, and so on."""
    if batch_number == 1:
        return base_path
    return base_path.with_name(f"{base_path.stem}_{batch_number}{base_path.suffix}")


def resolve_input_path(path_text: str | None, fallback: Path) -> Path:
    if path_text:
        return Path(path_text).expanduser().resolve()
    return fallback.expanduser().resolve()


def resolve_default_file(
    path_text: str | None,
    filename: str,
    input_folder: Path,
    script_folder: Path,
    option_name: str,
) -> Path:
    if path_text:
        candidate = Path(path_text).expanduser().resolve()
        if not candidate.is_file():
            raise FileNotFoundError(f"File was not found: {candidate}")
        return candidate

    candidates = [
        input_folder / filename,
        script_folder / filename,
        Path.cwd() / filename,
    ]
    unique_candidates: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        resolved = candidate.expanduser().resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique_candidates.append(resolved)

    for candidate in unique_candidates:
        if candidate.is_file():
            return candidate

    searched = "\n  ".join(str(path) for path in unique_candidates)
    raise FileNotFoundError(
        f"Could not find {filename}. Searched:\n  {searched}\n"
        f"Use {option_name} to specify its location."
    )


def parse_arguments() -> argparse.Namespace:
    script_folder = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(
        description=(
            "Lay out upright 63 x 88 mm card fronts with 3 mm black bleed on "
            f"landscape US Letter, using quantities from {CARDS_TXT_NAME}, then "
            f"generate matching back pages from {BACK_PDF_NAME}."
        )
    )
    parser.add_argument(
        "--input-folder",
        help=(
            "Folder containing card-front images. Default: the folder containing "
            "this script."
        ),
    )
    parser.add_argument(
        "--cards-txt",
        help=(
            f"Path to the card-quantity list. Default: {CARDS_TXT_NAME} in the "
            "input folder, script folder, or current folder."
        ),
    )
    parser.add_argument(
        "--back-pdf",
        help=(
            "Path to a PDF containing the card-back image. Its page layout is "
            f"ignored and backs are re-centered to the front grid. Default: "
            f"{BACK_PDF_NAME} in the input folder, script folder, or current folder. "
            "Not required with --single-sided."
        ),
    )
    parser.add_argument(
        "--single-sided",
        "--no-card-backs",
        dest="single_sided",
        action="store_true",
        help=(
            "Generate only card-front pages. The card-back PDF is not required, "
            "back pages are not generated, and each output contains up to eight "
            "single-sided cards."
        ),
    )
    parser.add_argument(
        "--fronts-pdf",
        help=(
            "Base path for intermediate per-batch fronts PDFs. Later batches "
            f"receive _2, _3, etc. Default: script folder/{FRONTS_PDF_NAME}."
        ),
    )
    parser.add_argument(
        "--output",
        help=(
            "Base path for final merged PDFs. Later batches receive _2, _3, "
            f"etc. Default: script folder/{OUTPUT_PDF_NAME}."
        ),
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Also search subfolders of --input-folder.",
    )
    parser.add_argument(
        "--no-auto-crop",
        action="store_true",
        help=(
            "Do not remove substantial black scanner/canvas borders. Use this if "
            "your images are already tightly cropped."
        ),
    )
    parser.add_argument(
        "--no-cut-guides",
        action="store_true",
        help="Do not draw thin card-boundary cutting guides.",
    )
    parser.add_argument(
        "--duplex-binding",
        choices=("long-edge", "short-edge", "none"),
        default="long-edge",
        help=(
            "Automatic duplex flip used by the printer. long-edge mirrors back "
            "slots vertically for landscape paper; short-edge mirrors them "
            "horizontally; none preserves the front coordinates. Default: long-edge."
        ),
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=600,
        help="Image resolution embedded in the fronts PDF. Default: 600.",
    )
    parser.add_argument(
        "--jpeg-quality",
        type=int,
        default=95,
        choices=range(1, 101),
        metavar="1-100",
        help="JPEG quality for embedded card images. Default: 95.",
    )

    layout_group = parser.add_argument_group("layout overrides")
    layout_group.add_argument("--paper-width-mm", type=float, default=279.4)
    layout_group.add_argument("--paper-height-mm", type=float, default=215.9)
    layout_group.add_argument("--card-width-mm", type=float, default=63.0)
    layout_group.add_argument("--card-height-mm", type=float, default=88.0)
    layout_group.add_argument("--bleed-mm", type=float, default=3.0)
    layout_group.add_argument(
        "--margin-mm",
        type=float,
        default=0.127,
        help="Gap between card images. Default: 0.127 mm.",
    )

    args = parser.parse_args()

    args.script_folder = script_folder
    args.input_folder_path = resolve_input_path(args.input_folder, script_folder)
    args.fronts_pdf_path = resolve_input_path(
        args.fronts_pdf, script_folder / FRONTS_PDF_NAME
    )
    args.output_path = resolve_input_path(args.output, script_folder / OUTPUT_PDF_NAME)
    return args


def main() -> int:
    args = parse_arguments()

    if args.dpi < 72:
        raise SystemExit("--dpi must be at least 72.")
    if args.bleed_mm < 0 or args.margin_mm < 0:
        raise SystemExit("Bleed and margin cannot be negative.")

    input_folder: Path = args.input_folder_path
    fronts_pdf: Path = args.fronts_pdf_path
    output_pdf: Path = args.output_path

    if not input_folder.is_dir():
        raise SystemExit(f"Input folder was not found: {input_folder}")

    try:
        cards_txt = resolve_default_file(
            args.cards_txt,
            CARDS_TXT_NAME,
            input_folder,
            args.script_folder,
            "--cards-txt",
        )
        back_pdf = None
        if not args.single_sided:
            back_pdf = resolve_default_file(
                args.back_pdf,
                BACK_PDF_NAME,
                input_folder,
                args.script_folder,
                "--back-pdf",
            )
    except (FileNotFoundError, ValueError) as exc:
        raise SystemExit(str(exc)) from exc

    distinct_paths = {
        "fronts PDF": fronts_pdf,
        "output PDF": output_pdf,
    }
    if back_pdf is not None:
        distinct_paths["back PDF"] = back_pdf
    if len({path for path in distinct_paths.values()}) != len(distinct_paths):
        details = "\n".join(f"  {name}: {path}" for name, path in distinct_paths.items())
        raise SystemExit(f"Fronts, back, and output PDF paths must be different:\n{details}")

    try:
        card_requests, parse_errors = parse_cards_txt(cards_txt)
    except OSError as exc:
        raise SystemExit(f"Could not read {cards_txt}: {exc}") from exc

    for parse_error in parse_errors:
        print(parse_error, file=sys.stderr)

    if not card_requests:
        raise SystemExit(f"No valid card entries were found in {cards_txt}")

    image_paths = collect_images(input_folder, args.recursive)
    if not image_paths:
        suffixes = ", ".join(sorted(IMAGE_EXTENSIONS))
        raise SystemExit(
            f"No card-front images found in {input_folder}\n"
            f"Supported extensions: {suffixes}"
        )

    selected_image_paths, missing_requests, ignored_count = select_images_from_cards_txt(
        card_requests, image_paths
    )

    for missing_request in missing_requests:
        print(
            f'ERROR: Card from cards.txt was not found: "{missing_request.name}" '
            f"(line {missing_request.line_number}, quantity "
            f"{missing_request.quantity})",
            file=sys.stderr,
        )

    if not selected_image_paths:
        raise SystemExit("No cards from cards.txt could be matched to image files.")

    try:
        layout = calculate_layout(
            paper_width_mm=args.paper_width_mm,
            paper_height_mm=args.paper_height_mm,
            card_width_mm=args.card_width_mm,
            card_height_mm=args.card_height_mm,
            bleed_mm=args.bleed_mm,
            margin_mm=args.margin_mm,
        )
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc

    card_batches = chunk_items(selected_image_paths, layout.cards_per_page)

    back_card_image = None
    if not args.single_sided:
        if back_pdf is None:
            raise SystemExit("A card-back PDF is required unless --single-sided is used.")
        try:
            back_card_image = extract_card_back_image(
                back_pdf,
                card_width_mm=args.card_width_mm,
                card_height_mm=args.card_height_mm,
            )
        except Exception as exc:
            raise SystemExit(f"Could not prepare the card-back image: {exc}") from exc

    fronts_pdf.parent.mkdir(parents=True, exist_ok=True)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    requested_copy_count = sum(request.quantity for request in card_requests)
    selected_copy_count = len(selected_image_paths)

    print(f"Input folder: {input_folder}")
    print(f"Cards list: {cards_txt}")
    print(
        f"Found {len(image_paths)} image file(s); "
        f"{ignored_count} not selected by cards.txt and ignored."
    )
    print(
        f"Cards.txt requests: {len(card_requests)} line(s), "
        f"{requested_copy_count} requested copies; "
        f"selected {selected_copy_count} copies."
    )
    if missing_requests:
        print(
            f"WARNING: {len(missing_requests)} cards.txt line(s) had no matching "
            "image and were skipped.",
            file=sys.stderr,
        )
    siding = (
        "single-sided; no card backs"
        if args.single_sided
        else f"duplex binding {args.duplex_binding}"
    )
    print(
        "Layout: "
        f"{layout.columns} x {layout.rows} = {layout.cards_per_page} cards/page; "
        f"card {layout.card_width_mm:g} x {layout.card_height_mm:g} mm; "
        f"black bleed {layout.bleed_mm:g} mm per side; "
        f"margin {layout.margin_mm:g} mm; "
        f"{siding}."
    )

    batch_results: list[dict[str, object]] = []
    try:
        for batch_number, batch_images in enumerate(card_batches, start=1):
            batch_fronts_pdf = numbered_batch_path(fronts_pdf, batch_number)
            batch_output_pdf = numbered_batch_path(output_pdf, batch_number)

            front_page_count = build_fronts_pdf(
                image_paths=batch_images,
                output_pdf=batch_fronts_pdf,
                layout=layout,
                dpi=args.dpi,
                jpeg_quality=args.jpeg_quality,
                auto_crop=not args.no_auto_crop,
                report_auto_crop=(batch_number == 1),
                cut_guides=not args.no_cut_guides,
            )

            batch_backs_pdf = None
            back_page_count = 0
            if args.single_sided:
                shutil.copyfile(batch_fronts_pdf, batch_output_pdf)
                merged_front_pages = front_page_count
                total_pages = front_page_count
            else:
                batch_backs_pdf = numbered_batch_path(
                    fronts_pdf.with_name(f"{fronts_pdf.stem}_backs{fronts_pdf.suffix}"),
                    batch_number,
                )
                if back_card_image is None:
                    raise RuntimeError("Card-back image was not prepared for duplex output.")

                back_page_count = build_backs_pdf(
                    back_image=back_card_image,
                    card_count=len(batch_images),
                    output_pdf=batch_backs_pdf,
                    layout=layout,
                    dpi=args.dpi,
                    jpeg_quality=args.jpeg_quality,
                    cut_guides=not args.no_cut_guides,
                    duplex_binding=args.duplex_binding,
                )
                merged_front_pages, merged_back_pages, total_pages = merge_pdfs(
                    fronts_pdf=batch_fronts_pdf,
                    back_pdf=batch_backs_pdf,
                    output_pdf=batch_output_pdf,
                )

                if merged_back_pages != back_page_count:
                    raise RuntimeError(
                        "Back PDF page count changed unexpectedly while merging: "
                        f"expected {back_page_count}, got {merged_back_pages}."
                    )

                if merged_front_pages != front_page_count:
                    raise RuntimeError(
                        "Front PDF page count changed unexpectedly while merging: "
                        f"expected {front_page_count}, got {merged_front_pages}."
                    )

            batch_results.append(
                {
                    "batch_number": batch_number,
                    "card_count": len(batch_images),
                    "fronts_pdf": batch_fronts_pdf,
                    "front_pages": merged_front_pages,
                    "backs_pdf": batch_backs_pdf,
                    "back_pages": back_page_count,
                    "output_pdf": batch_output_pdf,
                    "total_pages": total_pages,
                }
            )
    except Exception as exc:
        raise SystemExit(f"Failed to build the PDF: {exc}") from exc

    output_kind = "single-sided" if args.single_sided else "merged front/back"
    print(
        f"Created {len(batch_results)} {output_kind} PDF(s) from "
        f"{len(card_batches)} batch(es) of up to {layout.cards_per_page} cards:"
    )
    for result in batch_results:
        message = (
            f"  Batch {result['batch_number']}: "
            f"{result['card_count']} card(s); "
            f"fronts {result['fronts_pdf']} ({result['front_pages']} page(s)); "
        )
        if result["backs_pdf"] is not None:
            message += f"backs {result['backs_pdf']} ({result['back_pages']} page(s)); "
        message += f"output {result['output_pdf']} ({result['total_pages']} page(s))"
        print(message)
    return 0


if __name__ == "__main__":
    sys.exit(main())