from flask import (
    Flask,
    after_this_request,
    jsonify,
    request,
    send_file,
    send_from_directory,
)
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional
from collections import deque
from werkzeug.utils import secure_filename
import os
import io
import re
import tempfile
import threading
import time
import uuid
import shutil
import gc

try:
    import psutil  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    psutil = None

app = Flask(__name__)

# -----------------------------
# 設定（あなたの要望すべて反映）
# -----------------------------

# 1スライドあたりの最大文字数
MAX_CHARS_PER_SLIDE = 150

# テキスト領域設定（右下枠と被らない範囲）
TEXT_LEFT_CM = 0.79
TEXT_TOP_CM = 0.80
TEXT_WIDTH_CM = 25.2   # 枠にかからない右端まで
TEXT_HEIGHT_CM = 15.6

# 右下の枠設定（位置・サイズ）
FRAME_LEFT_CM = 25.87
FRAME_TOP_CM = 14.55
FRAME_WIDTH_CM = 8.0
FRAME_HEIGHT_CM = 4.5

# ページ番号（分割時のみ表示）
PAGE_LEFT_CM = 21.94
PAGE_TOP_CM = 16.93
PAGE_COLOR = RGBColor(0x00, 0x9D, 0xFF)  # #009DFF
PAGE_FONT_SIZE_PT = 32
PAGE_FONT_BOLD = True

# Web API settings
UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "split_pptx_uploads"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
MAX_MEMORY_MB = 400
CLEANUP_INTERVAL = 300


@dataclass
class ConversionJob:
    task_id: str
    input_path: Path
    output_dir: Path
    output_filename: str
    temp_dir: Path
    created_at: float
    log_callback: Callable[[str], None]


# In-memory structures
task_status: Dict[str, Dict[str, object]] = {}
conversion_queue: deque[ConversionJob] = deque()
queue_lock = threading.Lock()
queue_event = threading.Event()
worker_thread: Optional[threading.Thread] = None
last_cleanup = time.time()

# 明示的な話者ごとの色（固定）
NAME_FIXED_COLORS = {
    "仲條": RGBColor(0x00, 0xFD, 0xFF),  # #00FDFF（水色）
    "三村": RGBColor(0xFF, 0xFF, 0xFF),  # #FFFFFF（白）
    "星野": RGBColor(0xFF, 0xFF, 0x00),  # #FFFF00（黄色）
}

# 明示指定以外の話者に自動割り当て（以降固定）
AUTO_COLOR_POOL = [
    RGBColor(0xFF, 0x40, 0xFF),  # ピンク
    RGBColor(0xFF, 0xA5, 0x00),  # オレンジ
    RGBColor(0xFF, 0xFB, 0x00),  # 黄（予備）
]
name_color_map = {}
_auto_color_idx = 0


# =============================
# Backend helpers
# =============================


def _timestamp() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def _add_log(task_id: str, message: str) -> None:
    entry = task_status.get(task_id)
    if not entry:
        return
    logs: List[str] = entry.setdefault("logs", [])  # type: ignore[assignment]
    log_line = f"{_timestamp()} {message}"
    if not logs or logs[-1] != log_line:
        logs.append(log_line)


def _update_queue_positions_locked() -> None:
    for index, job in enumerate(conversion_queue):
        status = task_status.get(job.task_id)
        if status is not None:
            status["queue_position"] = index


def _enqueue_job(job: ConversionJob) -> int:
    with queue_lock:
        conversion_queue.append(job)
        _update_queue_positions_locked()
        queue_event.set()
        return conversion_queue.index(job)


def _dequeue_job() -> Optional[ConversionJob]:
    with queue_lock:
        if conversion_queue:
            job = conversion_queue.popleft()
            _update_queue_positions_locked()
            if not conversion_queue:
                queue_event.clear()
            return job
        queue_event.clear()
        return None


def _get_memory_usage_mb() -> float:
    if psutil is None:
        return 0.0
    try:
        process = psutil.Process()
        rss = process.memory_info().rss
        for child in process.children(recursive=True):
            try:
                rss += child.memory_info().rss
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return rss / 1024 / 1024
    except Exception:
        return 0.0


def _force_gc() -> None:
    try:
        gc.collect()
    except Exception:
        pass


def _cleanup_task_files(task_id: str) -> None:
    status = task_status.pop(task_id, None)
    if not status:
        return
    file_path_value = status.get("file_path")
    if file_path_value:
        try:
            path = Path(file_path_value)
            if path.exists():
                path.unlink()
            shutil.rmtree(path.parent, ignore_errors=True)
        except Exception:
            pass
    _force_gc()


def _auto_cleanup_if_needed() -> None:
    global last_cleanup
    current_time = time.time()
    if current_time - last_cleanup < CLEANUP_INTERVAL:
        return

    memory_mb = _get_memory_usage_mb()
    if memory_mb > MAX_MEMORY_MB:
        old_tasks: List[str] = []
        for task_id, status in task_status.items():
            created_at = status.get("created_at", 0)
            if status.get("status") == "completed" and created_at < current_time - 3600:
                old_tasks.append(task_id)
        for task_id in old_tasks:
            _cleanup_task_files(task_id)
        _force_gc()

    last_cleanup = current_time

# =============================
# 補助関数群
# =============================

def get_color_for_name(name: str):
    """話者ごとに色を固定して返す"""
    global _auto_color_idx
    if not name:
        return RGBColor(0xFF, 0xFF, 0xFF)  # デフォルト白

    # 明示指定があればそれを使う
    if name in NAME_FIXED_COLORS:
        return NAME_FIXED_COLORS[name]

    # 既存登録があればそれを再利用
    if name in name_color_map:
        return name_color_map[name]

    # 新規話者 → 自動カラー割り当て
    color = AUTO_COLOR_POOL[_auto_color_idx % len(AUTO_COLOR_POOL)]
    name_color_map[name] = color
    _auto_color_idx += 1
    return color


def clean_text(text: str) -> str:
    """ノート欄から制御文字を除去"""
    return text.replace("\x0b", "").strip()


def parse_notes_into_segments(note_text: str):
    """
    ノートを《名前》単位で抽出 [(名前, テキスト)]。
    連続する同一話者は結合して扱う。
    """
    segments = []
    current_name = None
    buffer = []

    for raw in note_text.splitlines():
        line = raw.strip()
        if not line:
            continue

        match = re.match(r"《(.+?)》", line)
        if match:
            # 直前バッファを保存
            if buffer:
                joined = "".join(buffer).strip()
                if joined:
                    segments.append((current_name, joined))
                buffer = []
            current_name = match.group(1)
            rest = line[match.end():]
            if rest:
                buffer.append(rest)
        else:
            buffer.append(line)

    if buffer:
        joined = "".join(buffer).strip()
        if joined:
            segments.append((current_name, joined))

    # 連続する同名を結合
    merged = []
    for name, text in segments:
        if merged and merged[-1][0] == name:
            merged[-1] = (name, merged[-1][1] + text)
        else:
            merged.append((name, text))
    return merged


def pack_segments_into_chunks(segments, max_len=MAX_CHARS_PER_SLIDE):
    """
    同じスライド内で話者を区切らず150文字単位で分割。
    1チャンク内に収まる文字数が<=max_len。
    """
    chunks = []
    cur = []
    cur_len = 0

    def flush():
        nonlocal cur, cur_len
        if cur:
            chunks.append(cur)
            cur = []
            cur_len = 0

    for name, text in segments:
        i = 0
        while i < len(text):
            remain = max_len - cur_len
            if remain <= 0:
                flush()
                remain = max_len
            take = min(remain, len(text) - i)
            part = text[i:i + take]
            cur.append((name, part))
            cur_len += len(part)
            i += take
    flush()
    return chunks


def add_photo_frame(slide):
    """右下に固定枠を追加"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(FRAME_LEFT_CM),
        Cm(FRAME_TOP_CM),
        Cm(FRAME_WIDTH_CM),
        Cm(FRAME_HEIGHT_CM)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)
    shape.line.color.rgb = RGBColor(0x64, 0x64, 0x64)
    shape.line.width = Pt(2)
    return shape


def add_page_indicator(slide, index, total):
    """分割時のみページ番号を表示"""
    if total <= 1:
        return
    tb = slide.shapes.add_textbox(Cm(PAGE_LEFT_CM), Cm(PAGE_TOP_CM), Cm(4), Cm(1.5))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{index}/{total}"
    p.alignment = PP_ALIGN.LEFT
    font = p.font
    font.name = "メイリオ"
    font.size = Pt(PAGE_FONT_SIZE_PT)
    font.bold = PAGE_FONT_BOLD
    font.color.rgb = PAGE_COLOR


# =============================
# スライド生成
# =============================

def create_script_slides(notes):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)

    for note in notes:
        segments = parse_notes_into_segments(note)
        chunks = pack_segments_into_chunks(segments, MAX_CHARS_PER_SLIDE)
        total_parts = len(chunks)

        for idx, chunk in enumerate(chunks, start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 背景を黒に
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # テキストボックス
            txBox = slide.shapes.add_textbox(
                Cm(TEXT_LEFT_CM),
                Cm(TEXT_TOP_CM),
                Cm(TEXT_WIDTH_CM),
                Cm(TEXT_HEIGHT_CM)
            )
            tf = txBox.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.space_before = Pt(0)
            p.space_after = Pt(0)

            # 各話者テキストをrun単位で追加
            for name, part in chunk:
                run = p.add_run()
                prefix = f"《{name}》" if name else ""
                run.text = prefix + part
                f = run.font
                f.name = "メイリオ"
                f.size = Pt(40)
                f.bold = True
                f.color.rgb = get_color_for_name(name)

            # 右下の枠
            add_photo_frame(slide)

            # ページ番号
            add_page_indicator(slide, idx, total_parts)

    return prs


def generate_script_slides(
    input_path: Path,
    temp_dir: Path,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Path:
    def log(message: str) -> None:
        if log_callback:
            log_callback(message)

    log("PPTXファイルを解析中...")
    prs = Presentation(str(input_path))

    notes: List[str] = []
    for slide in prs.slides:
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            text = slide.notes_slide.notes_text_frame.text
            cleaned = clean_text(text)
            if cleaned:
                notes.append(cleaned)

    if not notes:
        log("ノートが見つかりませんでしたが空のプレゼンを生成します。")

    log("スクリプトスライドを生成しています...")
    new_prs = create_script_slides(notes)
    output_path = temp_dir / "スクリプトスライド_自動生成.pptx"
    new_prs.save(output_path)
    log("スクリプトスライドの生成が完了しました。")
    return output_path


def _process_conversion_job(job: ConversionJob) -> None:
    task_id = job.task_id
    status = task_status.get(task_id)
    if status is None:
        shutil.rmtree(job.temp_dir, ignore_errors=True)
        return

    status.update(
        {
            "status": "processing",
            "message": "Processing slides...",
            "queue_position": 0,
        }
    )
    _add_log(task_id, "変換を開始しました。")

    try:
        result_path = generate_script_slides(job.input_path, job.temp_dir, job.log_callback)

        job.output_dir.mkdir(parents=True, exist_ok=True)
        final_path = job.output_dir / job.output_filename
        if result_path != final_path:
            shutil.copy2(result_path, final_path)

        status.update(
            {
                "status": "completed",
                "message": "Conversion completed successfully!",
                "download_url": f"/download/{task_id}",
                "file_path": str(final_path),
                "queue_position": None,
            }
        )
        _add_log(task_id, "変換が正常に完了しました。")
    except Exception as exc:  # noqa: BLE001
        status.update(
            {
                "status": "failed",
                "message": f"Conversion failed: {exc}",
                "download_url": None,
                "queue_position": None,
            }
        )
        _add_log(task_id, f"変換に失敗しました: {exc}")
    finally:
        try:
            shutil.rmtree(job.temp_dir, ignore_errors=True)
        finally:
            _force_gc()


def _worker_loop() -> None:
    while True:
        queue_event.wait()
        while True:
            job = _dequeue_job()
            if job is None:
                break
            _process_conversion_job(job)


def _ensure_worker_started() -> None:
    global worker_thread
    if worker_thread is None or not worker_thread.is_alive():
        worker_thread = threading.Thread(target=_worker_loop, daemon=True)
        worker_thread.start()


def _serialize_status(task_id: str, status: Dict[str, object]) -> Dict[str, object]:
    response = {
        "task_id": task_id,
        "status": status.get("status", "unknown"),
        "message": status.get("message", ""),
        "download_url": status.get("download_url"),
        "logs": status.get("logs", []),
        "queue_position": status.get("queue_position"),
    }
    return response


# =============================
# Flask API ルーティング
# =============================


@app.route("/", methods=["GET", "HEAD"])
def serve_index():
    return send_from_directory("static", "index.html")


@app.route("/health", methods=["GET"])
def health() -> "jsonify":
    _auto_cleanup_if_needed()
    memory_usage = round(_get_memory_usage_mb(), 2)
    return jsonify(
        {
            "status": "healthy",
            "service": "PPTX Script Slides API",
            "memory_usage_mb": memory_usage,
            "active_tasks": len(task_status),
            "queue_length": len(conversion_queue),
        }
    )


@app.route("/convert", methods=["POST"])
def convert_pptx():
    _auto_cleanup_if_needed()

    memory_usage = _get_memory_usage_mb()
    if memory_usage > MAX_MEMORY_MB:
        return (
            jsonify(
                {
                    "detail": f"Server memory usage too high ({memory_usage:.1f}MB). Please try again later.",
                }
            ),
            503,
        )

    if "file" not in request.files:
        return jsonify({"detail": "No file uploaded"}), 400

    file_storage = request.files["file"]
    if not file_storage.filename:
        return jsonify({"detail": "No file selected"}), 400

    filename = secure_filename(file_storage.filename)
    if not filename.lower().endswith(".pptx"):
        return jsonify({"detail": "Only .pptx files are supported"}), 400

    task_id = str(uuid.uuid4())
    temp_dir = Path(tempfile.mkdtemp(prefix=f"pptx_{task_id}_"))
    input_path = temp_dir / filename
    output_dir = UPLOAD_FOLDER / task_id
    output_filename = "スクリプトスライド_自動生成.pptx"

    try:
        file_storage.stream.seek(0)
        with open(input_path, "wb") as buffer:
            while True:
                chunk = file_storage.stream.read(4 * 1024 * 1024)
                if not chunk:
                    break
                buffer.write(chunk)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"detail": f"Failed to save uploaded file: {exc}"}), 500

    def log_callback(message: str) -> None:
        _add_log(task_id, message)

    job = ConversionJob(
        task_id=task_id,
        input_path=input_path,
        output_dir=output_dir,
        output_filename=output_filename,
        temp_dir=temp_dir,
        created_at=time.time(),
        log_callback=log_callback,
    )

    status_entry = {
        "status": "queued",
        "message": "Waiting for available converter...",
        "download_url": None,
        "created_at": job.created_at,
        "logs": [],
        "file_path": None,
        "queue_position": len(conversion_queue),
    }
    task_status[task_id] = status_entry
    _add_log(task_id, "Conversion request accepted. Added to queue.")

    _ensure_worker_started()
    position = _enqueue_job(job)
    status_entry["queue_position"] = position

    response_status = "queued" if position > 0 else "processing"
    response_message = "Queued for processing" if position > 0 else "Started processing"

    return jsonify(
        {
            "task_id": task_id,
            "status": response_status,
            "message": response_message,
            "download_url": None,
            "logs": status_entry.get("logs", []),
            "queue_position": position if position > 0 else None,
        }
    )


@app.route("/status/<task_id>", methods=["GET"])
def get_status(task_id: str):
    status = task_status.get(task_id)
    if status is None:
        return jsonify({"detail": "Task not found"}), 404
    return jsonify(_serialize_status(task_id, status))


@app.route("/download/<task_id>", methods=["GET"])
def download_file(task_id: str):
    status = task_status.get(task_id)
    if status is None:
        return jsonify({"detail": "Task not found"}), 404

    if status.get("status") != "completed":
        return jsonify({"detail": "Conversion not completed"}), 400

    file_path_value = status.get("file_path")
    if not file_path_value:
        return jsonify({"detail": "File not available"}), 404

    file_path = Path(file_path_value)
    if not file_path.exists():
        return jsonify({"detail": "File not found on disk"}), 404

    @after_this_request
    def cleanup(response):  # type: ignore[override]
        _cleanup_task_files(task_id)
        return response

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_path.name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/cleanup/<task_id>", methods=["DELETE"])
def cleanup_task(task_id: str):
    if task_id not in task_status:
        return jsonify({"detail": "Task not found"}), 404
    _cleanup_task_files(task_id)
    return jsonify({"message": "Task cleaned up successfully"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=True)