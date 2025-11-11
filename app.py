#!/usr/bin/env python3
"""GUI application to generate script-style slides from PPTX notes."""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Callable, Iterable, List, Optional

import sys

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QHBoxLayout,
)
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Cm, Pt

MAX_CHARS_PER_SLIDE = 200
OUTPUT_FILENAME = "スクリプトスライド_自動生成.pptx"
FONT_NAME = "メイリオ"
FONT_SIZE_PT = 40
PAGE_INDICATOR_COLOR = RGBColor(0x00, 0xB0, 0xF0)
DEFAULT_FONT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
TEXTBOX_POSITION = {
    "left": Cm(0.79),
    "top": Cm(0.8),
    "width": Cm(32.31),
    "height": Cm(15.6),
}
THUMBNAIL_WIDTH_CM = 8.0
THUMBNAIL_MARGIN_CM = 0.5
DEFAULT_THUMBNAIL_DPI = 150
EMU_PER_INCH = 914400
SPEAKER_PATTERN = re.compile(r"^\s*(話者\d+)[:：]\s*(.*)$")

SPEAKER_COLORS = {
    "話者1": RGBColor(0xFF, 0xFF, 0x00),
    "話者2": RGBColor(0x00, 0xFF, 0xFF),
    "話者3": RGBColor(0x00, 0xF9, 0x00),
}

FONT_PATH_CANDIDATES = [
    "C:/Windows/Fonts/meiryo.ttc",
    "C:/Windows/Fonts/meiryo.ttf",
    "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/Library/Fonts/Osaka.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]

_FONT_CACHE: dict[int, ImageFont.ImageFont] = {}


@dataclass
class Segment:
    """Represents a single sentence or fragment with optional speaker metadata."""

    text: str
    speaker: Optional[str]


def segment_line(text: str, max_chars: int) -> List[str]:
    """Split a line into punctuation-aware segments within the max length."""
    pattern = re.compile(r"[^。、，,.！？!?]+[。、，,.！？!?]?")
    segments: List[str] = []
    for token in pattern.findall(text):
        trimmed = token.strip()
        if not trimmed:
            continue
        if len(trimmed) <= max_chars:
            segments.append(trimmed)
        else:
            for start in range(0, len(trimmed), max_chars):
                segments.append(trimmed[start : start + max_chars])
    if not segments and text:
        # Fallback when regex fails (e.g., single punctuation)
        for start in range(0, len(text), max_chars):
            segments.append(text[start : start + max_chars])
    return segments


def build_segments(notes_text: str, max_chars: int) -> List[Segment]:
    """Convert raw notes into ordered segments with speaker metadata."""
    segments: List[Segment] = []
    normalized = notes_text.replace("\r\n", "\n").replace("\r", "\n")
    for raw_line in normalized.split("\n"):
        stripped = raw_line.strip()
        if not stripped:
            segments.append(Segment("", None))
            continue
        match = SPEAKER_PATTERN.match(stripped)
        if match:
            speaker_label = match.group(1)
            content = match.group(2).strip()
        else:
            speaker_label = None
            content = stripped
        line_segments = segment_line(content, max_chars)
        if not line_segments:
            segments.append(Segment(content, speaker_label))
            continue
        for idx, part in enumerate(line_segments):
            display_text = part if idx > 0 or not speaker_label else f"{speaker_label}：{part}"
            segments.append(Segment(display_text, speaker_label))
    return segments


def chunk_segments(segments: Iterable[Segment], max_chars: int) -> List[List[Segment]]:
    """Group segments into chunks that respect the character limit."""
    chunks: List[List[Segment]] = []
    current: List[Segment] = []
    current_len = 0

    for seg in segments:
        seg_len = len(seg.text)
        additional = max(seg_len, 1)  # blank lines count as 1
        if current:
            additional += 1  # newline between segments
        if current and current_len + additional > max_chars:
            chunks.append(current)
            current = [seg]
            current_len = seg_len
        else:
            current.append(seg)
            if current_len == 0:
                current_len = seg_len
            else:
                current_len += additional
    if current:
        chunks.append(current)
    return chunks


def speaker_color(speaker: Optional[str]) -> RGBColor:
    return SPEAKER_COLORS.get(speaker or "", DEFAULT_FONT_COLOR)


def log(message: str, reporter: Optional[Callable[[str], None]]) -> None:
    print(message)
    if reporter:
        reporter(message)


def emu_to_px(value: int, dpi: int = DEFAULT_THUMBNAIL_DPI) -> int:
    return max(1, int(round(value * dpi / EMU_PER_INCH)))


def get_font(size_pt: float) -> ImageFont.ImageFont:
    rounded = int(round(size_pt)) if size_pt else 20
    if rounded in _FONT_CACHE:
        return _FONT_CACHE[rounded]
    for candidate in FONT_PATH_CANDIDATES:
        if Path(candidate).exists():
            try:
                font = ImageFont.truetype(candidate, rounded)
                _FONT_CACHE[rounded] = font
                return font
            except OSError:
                continue
    font = ImageFont.load_default()
    _FONT_CACHE[rounded] = font
    return font


def rgb_color_tuple(color: Optional[RGBColor], default=(0, 0, 0)) -> tuple[int, int, int]:
    if color is None:
        return default
    try:
        return (color[0], color[1], color[2])
    except (TypeError, IndexError):
        return default


def draw_text_block(
    image: Image.Image,
    shape,
    dpi: int = DEFAULT_THUMBNAIL_DPI,
) -> bool:
    if not shape.has_text_frame:
        return False
    text_frame = shape.text_frame
    raw_text = text_frame.text
    if not raw_text.strip():
        return False
    draw = ImageDraw.Draw(image)
    left = emu_to_px(int(shape.left), dpi)
    top = emu_to_px(int(shape.top), dpi)
    width = emu_to_px(int(shape.width), dpi)

    first_run = None
    for paragraph in text_frame.paragraphs:
        if paragraph.runs:
            first_run = paragraph.runs[0]
            break
    font_size = (
        first_run.font.size.pt
        if first_run and first_run.font.size
        else 24
    )
    color = rgb_color_tuple(first_run.font.color.rgb if first_run and first_run.font.color and first_run.font.color.rgb else None)
    font = get_font(font_size)

    lines: List[str] = []
    for paragraph in raw_text.splitlines():
        if not paragraph:
            lines.append("")
            continue
        current = ""
        for char in paragraph:
            test = current + char
            if draw.textlength(test, font=font) <= width:
                current = test
            else:
                if current:
                    lines.append(current)
                current = char
        if current:
            lines.append(current)
    line_height = font.getbbox("あ")[3] if hasattr(font, "getbbox") else font.size
    y = top
    drew_text = False
    for line in lines:
        draw.text((left, y), line, font=font, fill=color)
        y += line_height
        if line:
            drew_text = True

    return drew_text


def draw_shape_fill(image: Image.Image, shape, dpi: int = DEFAULT_THUMBNAIL_DPI) -> bool:
    fill = shape.fill
    if not fill or fill.type != MSO_FILL.SOLID:
        return False
    color = rgb_color_tuple(fill.fore_color.rgb if fill.fore_color.type is not None else None, default=(255, 255, 255))
    left = emu_to_px(int(shape.left), dpi)
    top = emu_to_px(int(shape.top), dpi)
    width = emu_to_px(int(shape.width), dpi)
    height = emu_to_px(int(shape.height), dpi)
    ImageDraw.Draw(image).rectangle(
        [left, top, left + width, top + height],
        fill=color,
    )
    return True


def draw_picture(image: Image.Image, shape, dpi: int = DEFAULT_THUMBNAIL_DPI) -> bool:
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    try:
        blob = shape.image.blob
    except Exception:
        return False
    with Image.open(BytesIO(blob)) as pic:
        pic = pic.convert("RGBA")
        width = emu_to_px(int(shape.width), dpi)
        height = emu_to_px(int(shape.height), dpi)
        if width > 0 and height > 0:
            pic = pic.resize((width, height), Image.LANCZOS)
        left = emu_to_px(int(shape.left), dpi)
        top = emu_to_px(int(shape.top), dpi)
        image.paste(pic, (left, top), pic if pic.mode == "RGBA" else None)
    return True


def slide_background_color(slide) -> tuple[int, int, int]:
    fill = slide.background.fill
    if fill and fill.type == MSO_FILL.SOLID:
        return rgb_color_tuple(fill.fore_color.rgb if fill.fore_color and fill.fore_color.rgb else None, default=(255, 255, 255))
    return (30, 30, 30)


def _draw_placeholder_notice(image: Image.Image) -> None:
    draw = ImageDraw.Draw(image)
    overlay_color = (40, 40, 40)
    draw.rectangle([(0, 0), (image.width, image.height)], fill=overlay_color)
    title_font = get_font(48)
    body_font = get_font(28)
    lines = [
        "サムネイルを内部描画しましたが",
        "表示できる要素が見つかりませんでした",
        "LibreOffice をインストールすると正確なプレビューが作成できます",
    ]
    current_y = image.height // 2 - (len(lines) * 50) // 2
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=body_font)
        text_width = bbox[2] - bbox[0]
        draw.text(
            ((image.width - text_width) / 2, current_y),
            line,
            font=body_font,
            fill=(220, 220, 220),
        )
        current_y += 50


def render_slide_to_image(
    slide,
    slide_width: int,
    slide_height: int,
    output_path: Path,
    dpi: int = DEFAULT_THUMBNAIL_DPI,
) -> bool:
    width_px = emu_to_px(slide_width, dpi)
    height_px = emu_to_px(slide_height, dpi)
    background = slide_background_color(slide)
    image = Image.new("RGB", (width_px, height_px), color=background)

    drawn_any = False
    for shape in slide.shapes:
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if draw_picture(image, shape, dpi):
                    drawn_any = True
            else:
                if draw_shape_fill(image, shape, dpi):
                    drawn_any = True
        except Exception:
            continue

    for shape in slide.shapes:
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                continue
            if draw_text_block(image, shape, dpi):
                drawn_any = True
        except Exception:
            continue

    if not drawn_any:
        _draw_placeholder_notice(image)

    image.save(output_path)
    return drawn_any


def ensure_blank_presentation() -> Presentation:
    prs = Presentation()
    if prs.slides:
        slide_id_list = prs.slides._sldIdLst  # type: ignore[attr-defined]
        slide_id = slide_id_list[0]
        slide_id_list.remove(slide_id)
    return prs


def apply_background(slide) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x00, 0x00, 0x00)


def add_textbox(slide, chunk: List[Segment]) -> None:
    textbox = slide.shapes.add_textbox(
        TEXTBOX_POSITION["left"],
        TEXTBOX_POSITION["top"],
        TEXTBOX_POSITION["width"],
        TEXTBOX_POSITION["height"],
    )
    text_frame = textbox.text_frame
    text_frame.text = ""
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.NONE
    text_frame.vertical_anchor = MSO_ANCHOR.TOP
    text_frame.margin_bottom = Pt(0)
    text_frame.margin_top = Pt(0)
    text_frame.margin_left = Pt(0)
    text_frame.margin_right = Pt(0)

    for idx, segment in enumerate(chunk):
        paragraph = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
        paragraph.text = segment.text
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_after = Pt(0)
        paragraph.space_before = Pt(0)
        font = paragraph.font
        font.name = FONT_NAME
        font.size = Pt(FONT_SIZE_PT)
        font.bold = False
        font.color.rgb = speaker_color(segment.speaker)


def add_page_indicator(slide, index: int, total: int, slide_width, slide_height) -> None:
    if total <= 1:
        return
    margin = Cm(THUMBNAIL_MARGIN_CM)
    indicator_width = Cm(4)
    indicator_height = Cm(1.5)
    textbox = slide.shapes.add_textbox(
        slide_width - indicator_width - margin,
        slide_height - indicator_height - margin,
        indicator_width,
        indicator_height,
    )
    text_frame = textbox.text_frame
    text_frame.text = ""
    text_frame.word_wrap = False
    text_frame.auto_size = MSO_AUTO_SIZE.NONE
    paragraph = text_frame.paragraphs[0]
    paragraph.text = f"{index}/{total}"
    paragraph.alignment = PP_ALIGN.RIGHT
    font = paragraph.font
    font.name = FONT_NAME
    font.size = Pt(FONT_SIZE_PT)
    font.bold = False
    font.color.rgb = PAGE_INDICATOR_COLOR


def add_thumbnail(slide, image_path: Path, slide_width, slide_height) -> None:
    if not image_path.exists():
        return
    margin = Cm(THUMBNAIL_MARGIN_CM)
    with Image.open(image_path) as img:
        width_cm = THUMBNAIL_WIDTH_CM
        height_cm = width_cm * img.height / img.width
    width = Cm(width_cm)
    height = Cm(height_cm)
    left = slide_width - width - margin
    top = slide_height - height - margin
    slide.shapes.add_picture(str(image_path), left, top, width=width, height=height)


def create_placeholder_thumbnail(index: int, output_dir: Path) -> Path:
    width, height = 1600, 900
    image = Image.new("RGB", (width, height), color=(30, 30, 30))
    draw = ImageDraw.Draw(image)
    title = f"Slide {index}"
    body_lines = [
        "LibreOffice/soffice が見つからないため",
        "サムネイルを簡易表示に切り替えました",
        "フルプレビューを得るには LibreOffice をインストールしてください",
    ]
    title_font = get_font(56)
    body_font = get_font(28)

    title_bbox = draw.textbbox((0, 0), title, font=title_font)
    title_width = title_bbox[2] - title_bbox[0]
    draw.text(
        ((width - title_width) / 2, height * 0.3),
        title,
        font=title_font,
        fill=(240, 240, 240),
    )

    start_y = height * 0.5
    line_spacing = 45
    for idx, line in enumerate(body_lines):
        bbox = draw.textbbox((0, 0), line, font=body_font)
        text_width = bbox[2] - bbox[0]
        draw.text(
            ((width - text_width) / 2, start_y + idx * line_spacing),
            line,
            font=body_font,
            fill=(200, 200, 200),
        )
    placeholder_path = output_dir / f"placeholder_slide_{index}.png"
    image.save(placeholder_path)
    return placeholder_path


def generate_thumbnails(
    prs: Presentation,
    pptx_path: Path,
    reporter: Optional[Callable[[str], None]],
) -> List[Optional[Path]]:
    slide_count = len(prs.slides)
    thumbnails: List[Optional[Path]] = [None] * slide_count
    persistent_dir = Path(tempfile.mkdtemp(prefix="pptx_thumbs_"))

    soffice = next((cmd for cmd in ("soffice", "libreoffice") if shutil.which(cmd)), None)
    if soffice:
        with tempfile.TemporaryDirectory(prefix="pptx_lo_") as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            try:
                subprocess.run(
                    [
                        soffice,
                        "--headless",
                        "--convert-to",
                        "png",
                        "--outdir",
                        str(tmp_dir),
                        str(pptx_path),
                    ],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                png_files = sorted(tmp_dir.glob("*.png"))
                if len(png_files) == slide_count:
                    for idx, path in enumerate(png_files):
                        dest = persistent_dir / f"slide_{idx + 1:03d}.png"
                        shutil.copy2(path, dest)
                        thumbnails[idx] = dest
                else:
                    log(
                        "LibreOffice で生成されたサムネイル数がスライド数と一致しません。内部レンダリングを使用します。",
                        reporter,
                    )
            except (subprocess.CalledProcessError, FileNotFoundError) as exc:
                log(f"サムネイル生成に失敗しました: {exc}. 内部レンダリングを使用します。", reporter)
    else:
        log(
            "LibreOffice/soffice が見つかりません。内部レンダリングでサムネイルを生成します。",
            reporter,
        )

    for idx, slide in enumerate(prs.slides, start=1):
        if thumbnails[idx - 1] is not None:
            continue
        dest = persistent_dir / f"slide_{idx:03d}.png"
        try:
            rendered = render_slide_to_image(slide, prs.slide_width, prs.slide_height, dest)
            if rendered:
                thumbnails[idx - 1] = dest
            else:
                log(
                    f"スライド {idx}: 内部レンダリング結果に表示要素がありません。プレースホルダーに切り替えます。",
                    reporter,
                )
                thumbnails[idx - 1] = create_placeholder_thumbnail(idx, persistent_dir)
        except Exception as exc:
            log(
                f"スライド {idx} のレンダリングに失敗しました: {exc}. プレースホルダーに切り替えます。",
                reporter,
            )
            thumbnails[idx - 1] = create_placeholder_thumbnail(idx, persistent_dir)

    return thumbnails


def generate_script_slides(input_file: Path, output_dir: Path, reporter: Optional[Callable[[str], None]]) -> Path:
    prs = Presentation(str(input_file))
    output_prs = ensure_blank_presentation()
    output_prs.slide_width = prs.slide_width
    output_prs.slide_height = prs.slide_height
    blank_layout = output_prs.slide_layouts[6]

    thumbnails = generate_thumbnails(prs, input_file, reporter)

    created = 0
    for slide_index, slide in enumerate(prs.slides, start=1):
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide and slide.notes_slide.notes_text_frame else ""
        if not notes.strip():
            log(f"スライド {slide_index}: ノートが空のためスキップします。", reporter)
            continue
        segments = build_segments(notes, MAX_CHARS_PER_SLIDE)
        chunks = chunk_segments(segments, MAX_CHARS_PER_SLIDE)
        log(
            f"スライド {slide_index}: {len(segments)} セグメント -> {len(chunks)} 枚に分割", reporter
        )
        for part_idx, chunk in enumerate(chunks, start=1):
            new_slide = output_prs.slides.add_slide(blank_layout)
            apply_background(new_slide)
            add_textbox(new_slide, chunk)
            add_page_indicator(new_slide, part_idx, len(chunks), output_prs.slide_width, output_prs.slide_height)
            thumbnail_path = thumbnails[slide_index - 1] if slide_index - 1 < len(thumbnails) else None
            if thumbnail_path:
                add_thumbnail(new_slide, thumbnail_path, output_prs.slide_width, output_prs.slide_height)
            created += 1

    if created == 0:
        raise ValueError("ノートから生成できるスライドがありませんでした。")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_FILENAME
    output_prs.save(str(output_path))
    log(f"生成完了: {created} 枚 -> {output_path}", reporter)
    return output_path


class ConversionThread(QThread):
    progress = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, input_file: Path, output_dir: Path):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir

    def run(self) -> None:  # noqa: D401
        try:
            generate_script_slides(self.input_file, self.output_dir, self.progress.emit)
            self.finished.emit(True, "処理が完了しました。")
        except Exception as exc:  # pylint: disable=broad-except
            self.finished.emit(False, f"エラーが発生しました: {exc}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PPTX スクリプトスライド生成ツール")
        self.resize(900, 600)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("変換するPowerPointファイル(.pptx)を選択")

        input_button = QPushButton("参照…")
        input_button.clicked.connect(self.choose_input_file)  # type: ignore[arg-type]

        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("入力 PPTX"))
        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(input_button)

        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("保存先フォルダ (未指定の場合は入力ファイルと同じフォルダ)")

        output_button = QPushButton("参照…")
        output_button.clicked.connect(self.choose_output_dir)  # type: ignore[arg-type]

        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("出力フォルダ"))
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(output_button)

        self.convert_button = QPushButton("変換")
        self.convert_button.setDefault(True)
        self.convert_button.clicked.connect(self.start_conversion)  # type: ignore[arg-type]

        exit_button = QPushButton("終了")
        exit_button.clicked.connect(self.close)  # type: ignore[arg-type]

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.convert_button)
        button_layout.addWidget(exit_button)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("ここに処理ログが表示されます")

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        header_label = QLabel("PowerPointファイルを選んで変換してください。")
        header_font = QFont()
        header_font.setPointSize(12)
        header_label.setFont(header_font)
        header_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        layout.addWidget(header_label)
        layout.addLayout(input_layout)
        layout.addLayout(output_layout)
        layout.addLayout(button_layout)
        layout.addWidget(self.log_view)

        self.setCentralWidget(central_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.menu_bar = QMenuBar()
        self.setMenuBar(self.menu_bar)

        file_menu = self.menu_bar.addMenu("ファイル")
        open_action = QAction("入力ファイルを開く", self)
        open_action.triggered.connect(self.choose_input_file)  # type: ignore[arg-type]
        file_menu.addAction(open_action)

        output_action = QAction("出力フォルダを選択", self)
        output_action.triggered.connect(self.choose_output_dir)  # type: ignore[arg-type]
        file_menu.addAction(output_action)

        file_menu.addSeparator()

        exit_action = QAction("終了", self)
        exit_action.triggered.connect(self.close)  # type: ignore[arg-type]
        file_menu.addAction(exit_action)

        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        toolbar.addAction(open_action)
        toolbar.addAction(output_action)
        toolbar.addSeparator()
        toolbar.addAction(exit_action)
        self.addToolBar(toolbar)

        self.worker: Optional[ConversionThread] = None

    def append_log(self, message: str) -> None:
        self.log_view.append(message)
        self.log_view.ensureCursorVisible()

    def choose_input_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "入力 PPTX を選択",
            "",
            "PowerPoint ファイル (*.pptx)",
        )
        if file_path:
            self.input_edit.setText(file_path)
            if not self.output_edit.text():
                self.output_edit.setText(str(Path(file_path).parent))

    def choose_output_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self,
            "出力フォルダを選択",
            "",
        )
        if directory:
            self.output_edit.setText(directory)

    def start_conversion(self) -> None:
        input_path = self.input_edit.text().strip()
        if not input_path:
            QMessageBox.warning(self, "入力ファイル", "入力ファイルを選択してください。")
            return
        input_file = Path(input_path)
        if not input_file.exists():
            QMessageBox.critical(self, "入力ファイル", "入力ファイルが見つかりません。")
            return

        output_path_value = self.output_edit.text().strip()
        output_dir = Path(output_path_value) if output_path_value else input_file.parent

        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                QMessageBox.critical(self, "出力フォルダ", f"出力フォルダを作成できません: {exc}")
                return

        self.log_view.clear()
        self.append_log("変換を開始します...")
        self.status_bar.showMessage("変換中...")
        self.convert_button.setEnabled(False)

        self.worker = ConversionThread(input_file, output_dir)
        self.worker.progress.connect(self.append_log)
        self.worker.finished.connect(self.finish_conversion)
        self.worker.start()

    def finish_conversion(self, success: bool, message: str) -> None:
        self.append_log(message)
        self.status_bar.showMessage(message, 5000)
        self.convert_button.setEnabled(True)
        QMessageBox.information(self, "結果", message if success else f"失敗: {message}")
        self.worker = None


def run_app() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run_app()
