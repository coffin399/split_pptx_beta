#!/usr/bin/env python3
"""GUI application to generate script-style slides from PPTX notes."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, List, Optional

import PySimpleGUI as sg
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
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
SPEAKER_PATTERN = re.compile(r"^\s*(話者\d+)[:：]\s*(.*)$")

SPEAKER_COLORS = {
    "話者1": RGBColor(0xFF, 0xFF, 0x00),
    "話者2": RGBColor(0x00, 0xFF, 0xFF),
    "話者3": RGBColor(0x00, 0xF9, 0x00),
}


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


def create_placeholder_thumbnail(index: int, temp_dir: Path) -> Path:
    width, height = 1600, 900
    image = Image.new("RGB", (width, height), color=(30, 30, 30))
    draw = ImageDraw.Draw(image)
    text = f"Slide {index} preview"
    text_color = (200, 200, 200)
    text_size = draw.textlength(text)  # type: ignore[attr-defined]
    if text_size:
        draw.text(((width - text_size) / 2, height / 2 - 20), text, fill=text_color)
    placeholder_path = temp_dir / f"placeholder_slide_{index}.png"
    image.save(placeholder_path)
    return placeholder_path


def generate_thumbnails(pptx_path: Path, slide_count: int, reporter: Optional[Callable[[str], None]]) -> List[Optional[Path]]:
    thumbnails: List[Optional[Path]] = [None] * slide_count
    with tempfile.TemporaryDirectory(prefix="pptx_thumbs_") as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)
        soffice = next((cmd for cmd in ("soffice", "libreoffice") if shutil.which(cmd)), None)
        if soffice:
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
                        thumbnails[idx] = path
                else:
                    log(
                        "LibreOffice で生成されたサムネイル数がスライド数と一致しません。プレースホルダーを使用します。",
                        reporter,
                    )
            except (subprocess.CalledProcessError, FileNotFoundError) as exc:
                log(f"サムネイル生成に失敗しました: {exc}. プレースホルダーを使用します。", reporter)
        else:
            log(
                "LibreOffice/soffice が見つかりません。サムネイルはプレースホルダーで代替します。",
                reporter,
            )
        for idx in range(slide_count):
            if thumbnails[idx] is None:
                thumbnails[idx] = create_placeholder_thumbnail(idx + 1, tmp_dir)
        # Copy placeholder files to persistent temp so caller can use after context exit
        persistent_dir = Path(tempfile.mkdtemp(prefix="pptx_thumbs_persist_"))
        copied: List[Optional[Path]] = []
        for path in thumbnails:
            if path is None:
                copied.append(None)
                continue
            new_path = persistent_dir / path.name
            shutil.copy2(path, new_path)
            copied.append(new_path)
        return copied


def generate_script_slides(input_file: Path, output_dir: Path, reporter: Optional[Callable[[str], None]]) -> Path:
    prs = Presentation(str(input_file))
    output_prs = ensure_blank_presentation()
    output_prs.slide_width = prs.slide_width
    output_prs.slide_height = prs.slide_height
    blank_layout = output_prs.slide_layouts[6]

    thumbnails = generate_thumbnails(input_file, len(prs.slides), reporter)

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


def create_window() -> sg.Window:
    sg.theme("DarkBlue3")
    layout = [
        [sg.Text("PowerPointファイルを選んで変換してください。")],
        [
            sg.Text("入力 PPTX", size=(12, 1)),
            sg.Input(key="-INPUT-", enable_events=True),
            sg.FileBrowse(file_types=(("PowerPoint", "*.pptx"),)),
        ],
        [
            sg.Text("出力フォルダ", size=(12, 1)),
            sg.Input(key="-OUTPUT-"),
            sg.FolderBrowse(),
        ],
        [sg.Button("変換", key="-CONVERT-"), sg.Button("終了", key="-EXIT-")],
        [sg.Multiline("", size=(80, 20), key="-LOG-", autoscroll=True, write_only=True)],
    ]
    return sg.Window("PPTX スクリプトスライド生成ツール", layout, finalize=True)


def run_app() -> None:
    window = create_window()

    def reporter(message: str) -> None:
        window["-LOG-"].print(message)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "-EXIT-"):
            break
        if event == "-CONVERT-":
            input_path = values.get("-INPUT-")
            if not input_path:
                reporter("入力ファイルを選択してください。")
                continue
            input_file = Path(input_path)
            if not input_file.exists():
                reporter("入力ファイルが見つかりません。")
                continue
            output_path_value = values.get("-OUTPUT-")
            output_dir = Path(output_path_value) if output_path_value else input_file.parent
            window["-CONVERT-"].update(disabled=True)
            window["-LOG-"].update("")
            try:
                generate_script_slides(input_file, output_dir, reporter)
                reporter("処理が完了しました。")
            except Exception as exc:  # pylint: disable=broad-except
                reporter(f"エラーが発生しました: {exc}")
            finally:
                window["-CONVERT-"].update(disabled=False)
    window.close()


if __name__ == "__main__":
    run_app()
