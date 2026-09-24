from __future__ import annotations

import copy
from datetime import datetime
import json
import math
import os
import re
import sys
import tempfile
import time
import traceback
import unicodedata
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Iterable

try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover - 起動時に利用者へ案内する
    try:
        import fitz  # type: ignore  # PyMuPDFの旧インポート名
    except ImportError:
        fitz = None  # type: ignore[assignment]


APP_NAME = "PDF リンク配置マネージャー"
APP_VERSION = "3.1.2"
REQUIRED_PYMUPDF_VERSION = "1.27.2.3"
MM_TO_PT = 72.0 / 25.4
HANDLE_RADIUS = 6
LOG_DIRECTORY_NAME = "PDFリンク配置マネージャー_logs"
BORDER_WIDTH_PRESETS = {
    "太い": "2.0",
    "標準": "1.0",
    "細い": "0.5",
}
HIGHLIGHT_MODE_LABELS = {
    "none": "なし",
    "invert": "反転",
    "outline": "アウトライン",
    "push": "プッシュ",
}
HIGHLIGHT_MODE_PDF = {
    "invert": "/I",
    "none": "/N",
    "outline": "/O",
    "push": "/P",
}
HIGHLIGHT_LABEL_TO_MODE = {
    label: mode for mode, label in HIGHLIGHT_MODE_LABELS.items()
}
BORDER_DISPLAY_LABELS = ("ボックスを表示", "ボックスを非表示")
BORDER_STYLE_LABELS = {
    "solid": "実線",
    "dashed": "破線",
    "underline": "下線",
}
BORDER_STYLE_PDF = {
    "solid": "/S",
    "dashed": "/D",
    "underline": "/U",
}
BORDER_STYLE_LABEL_TO_MODE = {
    label: mode for mode, label in BORDER_STYLE_LABELS.items()
}
DESTINATION_VIEW_LABELS = {
    "fit_page": "全体表示",
    "actual_size": "100%表示",
    "fit_width": "幅に合わせる",
    "fit_visible_width": "描画領域の幅に合わせる",
}
DESTINATION_VIEW_LABEL_TO_MODE = {
    label: mode for mode, label in DESTINATION_VIEW_LABELS.items()
}
NAMED_DESTINATION_OUTPUT_LABELS = {
    "preserve": "元の名前を保持",
    "convert": "ページ番号へ変換",
}
NAMED_DESTINATION_OUTPUT_LABEL_TO_MODE = {
    label: mode for mode, label in NAMED_DESTINATION_OUTPUT_LABELS.items()
}


class ValidationError(ValueError):
    pass


def require_pymupdf() -> None:
    """文字位置の回帰を避けるため、安全確認済みPyMuPDF版を検査する。"""
    if fitz is None:
        raise ValidationError(
            "PyMuPDF がインストールされていません。\n"
            "`python -m pip install --force-reinstall -r requirements.txt` "
            "を実行してください。"
        )
    installed_version = str(
        getattr(fitz, "VersionBind", getattr(fitz, "__version__", "unknown"))
    )
    if installed_version != REQUIRED_PYMUPDF_VERSION:
        raise ValidationError(
            "PDF内の文字・図形などの位置ずれを避けるため、"
            f"PyMuPDF {REQUIRED_PYMUPDF_VERSION} に固定しています。\n"
            f"現在のバージョン: {installed_version}\n\n"
            "次を実行してから再起動してください。\n"
            "python -m pip install --force-reinstall -r requirements.txt"
        )


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(value, high))


def zoomed_page_origin(
    page_origin: tuple[float, float],
    content_anchor: tuple[float, float],
    old_scale: float,
    new_scale: float,
) -> tuple[float, float]:
    """拡大縮小の前後で、カーソル下のページ座標を同じ位置に保つ。"""
    if old_scale <= 0 or new_scale <= 0:
        raise ValueError("拡大率は0より大きい必要があります。")
    page_x = (content_anchor[0] - page_origin[0]) / old_scale
    page_y = (content_anchor[1] - page_origin[1]) / old_scale
    return (
        content_anchor[0] - page_x * new_scale,
        content_anchor[1] - page_y * new_scale,
    )


def rects_overlap(a: fitz.Rect, b: fitz.Rect, tolerance: float = 0.01) -> bool:
    intersection = a & b
    return (
        not intersection.is_empty
        and intersection.width > tolerance
        and intersection.height > tolerance
    )


def parse_page_spec(spec: str, page_count: int) -> list[int]:
    """利用者向け1始まりの指定を、重複のない0始まりページ番号へ変換する。"""
    text = spec.strip().lower().replace(" ", "")
    if not text:
        raise ValidationError("対象ページを入力してください。")
    if text in {"all", "すべて", "全ページ", "*"}:
        return list(range(page_count))
    if text in {"odd", "奇数", "奇数ページ"}:
        return list(range(0, page_count, 2))
    if text in {"even", "偶数", "偶数ページ"}:
        return list(range(1, page_count, 2))

    result: list[int] = []
    seen: set[int] = set()
    for part in text.split(","):
        if not part:
            raise ValidationError(f"ページ指定「{spec}」の形式が正しくありません。")
        if "-" in part:
            pair = part.split("-")
            if len(pair) != 2 or not pair[0].isdigit() or not pair[1].isdigit():
                raise ValidationError(f"ページ範囲「{part}」の形式が正しくありません。")
            start, end = int(pair[0]), int(pair[1])
            if start > end:
                raise ValidationError(f"ページ範囲「{part}」は開始が終了を超えています。")
            numbers: Iterable[int] = range(start, end + 1)
        else:
            if not part.isdigit():
                raise ValidationError(f"ページ「{part}」は数値ではありません。")
            numbers = (int(part),)
        for number in numbers:
            if number < 1 or number > page_count:
                raise ValidationError(
                    f"ページ {number} はPDFの範囲（1～{page_count}）外です。"
                )
            zero_based = number - 1
            if zero_based not in seen:
                seen.add(zero_based)
                result.append(zero_based)
    return result


def hex_to_pdf_color(value: str) -> tuple[float, float, float]:
    text = value.strip().lstrip("#")
    if len(text) != 6:
        raise ValidationError("枠色は #RRGGBB 形式で指定してください。")
    try:
        components = tuple(int(text[i : i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError as exc:
        raise ValidationError("枠色は #RRGGBB 形式で指定してください。") from exc
    return components  # type: ignore[return-value]


@dataclass
class LinkRule:
    name: str = ""
    page_spec: str = "全ページ"
    x: float = 0.0
    y: float = 0.0
    width: float = MM_TO_PT
    height: float = MM_TO_PT
    link_type: str = "page"
    target_page: str | int = "1"
    url: str = "https://"
    destination_view: str = "fit_page"
    border_visible: bool = False
    border_color: str = "#FF0000"
    border_width: float = 1.0
    border_style: str = "solid"
    highlight_mode: str = "none"
    enabled: bool = True
    source_page: int | None = None
    source_xref: int | None = None
    source_named_destination: str | None = None
    source_named_destination_pdf: str | None = None
    deleted: bool = False
    rule_id: str = field(default_factory=lambda: uuid.uuid4().hex)

    def rect(self) -> fitz.Rect:
        return fitz.Rect(self.x, self.y, self.x + self.width, self.y + self.height)

    @classmethod
    def from_dict(cls, data: dict) -> "LinkRule":
        allowed = {item.name for item in cls.__dataclass_fields__.values()}
        clean = {key: value for key, value in data.items() if key in allowed}
        if "target_page" in clean:
            clean["target_page"] = str(clean["target_page"])
        return cls(**clean)


@dataclass
class ValidationReport:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    insertion_count: int = 0

    @property
    def ok(self) -> bool:
        return not self.errors


def rectangle_context(
    rect: fitz.Rect | None,
    *,
    page_number: int | None = None,
) -> str:
    parts: list[str] = []
    if page_number is not None:
        parts.append(f"ページ={page_number}")
    if rect is not None:
        parts.append(
            "X={:.2f}pt ({:.2f}mm), Y={:.2f}pt ({:.2f}mm), "
            "幅={:.2f}pt ({:.2f}mm), 高さ={:.2f}pt ({:.2f}mm)".format(
                rect.x0,
                rect.x0 / MM_TO_PT,
                rect.y0,
                rect.y0 / MM_TO_PT,
                rect.width,
                rect.width / MM_TO_PT,
                rect.height,
                rect.height / MM_TO_PT,
            )
        )
    return " | ".join(parts)


def rule_context(rule: LinkRule, page_number: int | None = None) -> str:
    page_text = (
        f"ページ={page_number}"
        if page_number is not None
        else f"対象ページ={rule.page_spec}"
    )
    return (
        f"設定名={rule.name or '(空欄)'} | {page_text} | "
        + rectangle_context(rule.rect())
    )


def write_diagnostic_log(
    operation: str,
    messages: Iterable[str],
    *,
    source_path: Path | None = None,
    output_path: Path | None = None,
    exception: BaseException | None = None,
) -> Path:
    """詳細を小さなメッセージボックスではなくUTF-8ログへ保存する。"""
    message_list = [str(message) for message in messages]
    timestamp = datetime.now().astimezone()
    slug = re.sub(r"[^0-9A-Za-z_-]+", "_", operation).strip("_") or "log"
    filename = f"{timestamp:%Y%m%d_%H%M%S_%f}_{slug}.log"
    candidate_bases: list[Path] = []
    if output_path is not None:
        candidate_bases.append(output_path.parent)
    if source_path is not None:
        candidate_bases.append(source_path.parent)
    candidate_bases.append(Path(tempfile.gettempdir()))

    lines = [
        f"{APP_NAME} {APP_VERSION} エラーログ",
        f"日時: {timestamp.isoformat(timespec='seconds')}",
        f"操作: {operation}",
        f"入力PDF: {source_path if source_path is not None else '(なし)'}",
        f"出力PDF: {output_path if output_path is not None else '(なし)'}",
        f"記録件数: {len(message_list)}",
        "",
    ]
    lines.extend(
        f"[{index:04d}] {message}"
        for index, message in enumerate(message_list, start=1)
    )
    if exception is not None:
        lines.extend(
            [
                "",
                f"例外: {type(exception).__name__}: {exception}",
                "",
                "トレースバック:",
                traceback.format_exc(),
            ]
        )
    content = "\n".join(lines).rstrip() + "\n"

    last_error: OSError | None = None
    tried: set[Path] = set()
    for base in candidate_bases:
        resolved_base = base.resolve()
        if resolved_base in tried:
            continue
        tried.add(resolved_base)
        log_directory = resolved_base / LOG_DIRECTORY_NAME
        try:
            log_directory.mkdir(parents=True, exist_ok=True)
            log_path = log_directory / filename
            log_path.write_text(content, encoding="utf-8-sig")
            return log_path
        except OSError as exc:
            last_error = exc
    raise OSError("エラーログの保存先を作成できませんでした。") from last_error

def display_rect_to_insertion(page: fitz.Page, rect: fitz.Rect) -> fitz.Rect:
    """表示座標をinsert_linkが期待する非回転座標へ変換する。"""
    if page.rotation:
        return fitz.Rect(rect) * page.derotation_matrix
    return fitz.Rect(rect)


def display_point_to_insertion(page: fitz.Page, point: fitz.Point) -> fitz.Point:
    if page.rotation:
        return fitz.Point(point) * page.derotation_matrix
    return fitz.Point(point)


def resolve_target_page(document: fitz.Document, target: str | int) -> int:
    """数値ページまたは名前付き移動先を0始まりページへ解決する。"""
    text = str(target).strip()
    if not text:
        raise ValidationError("PDF内ページまたは移動先名を入力してください。")
    if text.isdigit():
        page_index = int(text) - 1
    else:
        destination = document.resolve_names().get(text)
        if not destination:
            raise ValidationError(f"移動先名「{text}」がPDF内に見つかりません。")
        page_index = int(destination.get("page", -1))
    if not 0 <= page_index < document.page_count:
        raise ValidationError(
            f"リンク先「{text}」はPDFのページへ解決できません。"
        )
    return page_index


def destination_view_from_source(
    document: fitz.Document, xref: int, link: dict
) -> str:
    """既存リンクのPDF移動先配列から表示方法を推定する。"""
    try:
        source = document.xref_object(xref, compressed=True) if xref > 0 else ""
    except Exception:
        source = ""
    _, destination_value = destination_value_from_source(document, xref)
    searchable = " ".join(
        (
            source,
            destination_value,
            str(link.get("dest", "")),
            str(link.get("view", "")),
        )
    )
    if re.search(r"/XYZ\s*null\s*null\s*1(?:\.0+)?(?=[]/\s]|$)", searchable):
        return "actual_size"
    if "/FitBH" in searchable or re.search(r"\bFitBH\b", searchable):
        return "fit_visible_width"
    if "/FitH" in searchable or re.search(r"\bFitH\b", searchable):
        return "fit_width"
    if re.search(r"(?:/|\b)Fit(?![A-Z])", searchable):
        return "fit_page"
    try:
        if math.isclose(float(link.get("zoom", 0) or 0), 1.0, abs_tol=0.001):
            return "actual_size"
    except (TypeError, ValueError):
        pass
    return "fit_page"


def destination_value_from_source(
    document: fitz.Document, xref: int
) -> tuple[str, str]:
    """注釈の /Dest または直接・間接 /A の /D を取得する。"""
    if xref <= 0:
        return "null", "null"
    for key in ("Dest", "A/D"):
        value_type, value = _safe_xref_get_key(document, xref, key)
        if value_type not in {"null", "none"} and value not in {"", "null"}:
            if value_type == "xref":
                try:
                    referred_xref = int(value.split()[0])
                    return value_type, document.xref_object(
                        referred_xref, compressed=True
                    )
                except (ValueError, RuntimeError):
                    pass
            return value_type, value
    action_type, action_value = _safe_xref_get_key(document, xref, "A")
    if action_type == "xref":
        try:
            action_xref = int(action_value.split()[0])
            value_type, value = _safe_xref_get_key(document, action_xref, "D")
            if value_type not in {"null", "none"} and value not in {"", "null"}:
                if value_type == "xref":
                    destination_xref = int(value.split()[0])
                    return value_type, document.xref_object(
                        destination_xref, compressed=True
                    )
                return value_type, value
        except (ValueError, RuntimeError):
            pass
    return "null", "null"


def direct_destination_page_from_source(
    document: fitz.Document, xref: int
) -> int | None:
    """PDFの直接移動先配列に含まれるページ参照を0始まりで返す。"""
    if xref <= 0:
        return None
    _, source = destination_value_from_source(document, xref)
    match = re.search(
        r"\[?\s*(\d+)\s+0\s+R\s*/(?:XYZ|Fit(?:B?[HV])?|FitR)\b",
        source,
    )
    if match is None:
        return None
    page_xref = int(match.group(1))
    for page_index in range(document.page_count):
        if document.page_xref(page_index) == page_xref:
            return page_index
    return None


def named_destination_name_from_source(
    document: fitz.Document, xref: int
) -> str | None:
    value_type, value = destination_value_from_source(document, xref)
    if value_type == "name" and value not in {"", "null"}:
        return value[1:] if value.startswith("/") else value
    if value_type == "string" and value not in {"", "null"}:
        return value
    return None


def named_destination_pdf_from_source(
    document: fitz.Document, xref: int
) -> str | None:
    """名前付き移動先を再設定できるPDF値として取得する。"""
    if xref <= 0:
        return None
    value_type, value = destination_value_from_source(document, xref)
    if value_type == "name" and value not in {"", "null"}:
        # xref_get_key() は #20 などを復号して返すため、そのまま /Name として
        # 再利用すると空白や記号を含む名前が不正になる。PDF文字列へ安全に直す。
        return fitz.get_pdf_str(value[1:] if value.startswith("/") else value)
    if value_type == "string" and value not in {"", "null"}:
        return fitz.get_pdf_str(value)
    return None


def uri_from_source(document: fitz.Document, xref: int) -> str | None:
    if xref <= 0:
        return None
    value_type, value = _safe_xref_get_key(document, xref, "A/URI")
    if value_type in {"string", "name"} and value not in {"", "null"}:
        return value[1:] if value_type == "name" and value.startswith("/") else value
    return None


def action_kind_from_source(document: fitz.Document, xref: int) -> str:
    if xref <= 0:
        return "不明"
    value_type, value = _safe_xref_get_key(document, xref, "A/S")
    if value_type == "name" and value not in {"", "null"}:
        return value[1:] if value.startswith("/") else value
    if _safe_xref_get_key(document, xref, "Dest")[0] not in {"null", "none"}:
        return "Dest"
    return "不明"


def _safe_xref_get_key(
    document: fitz.Document, xref: int, key: str
) -> tuple[str, str]:
    """空・破損したActionの子キーを読めない場合も、当該注釈だけ未対応にする。"""
    try:
        return document.xref_get_key(xref, key)
    except Exception:
        return "null", "null"


def named_destination_pdf_for_rule(rule: LinkRule) -> str:
    """既存値を優先し、変更された名前は安全なPDF文字列へ変換する。"""
    name = str(rule.target_page).strip()
    if (
        rule.source_named_destination == name
        and rule.source_named_destination_pdf
    ):
        return rule.source_named_destination_pdf
    return fitz.get_pdf_str(name)


def destination_array(
    document: fitz.Document, target_page_index: int, view_mode: str
) -> str:
    page_xref = document.page_xref(target_page_index)
    if view_mode == "actual_size":
        return f"[{page_xref} 0 R /XYZ null null 1]"
    if view_mode == "fit_width":
        return f"[{page_xref} 0 R /FitH null]"
    if view_mode == "fit_visible_width":
        return f"[{page_xref} 0 R /FitBH null]"
    return f"[{page_xref} 0 R /Fit]"


def clean_auto_name_text(text: str, limit: int = 20) -> str:
    cleaned = "".join(
        char
        for char in text
        if not char.isspace()
        and unicodedata.category(char)[0] not in {"C", "Z"}
    )
    return cleaned[:limit]


def rect_intersects(left: fitz.Rect, right: fitz.Rect) -> bool:
    intersection = fitz.Rect(left) & fitz.Rect(right)
    return not intersection.is_empty and intersection.width > 0.1 and intersection.height > 0.1


def automatic_rule_name(
    page: fitz.Page,
    rect: fitz.Rect,
    page_number: int,
    rules: list[LinkRule],
) -> str:
    """文字、画像、フォーム、注釈、図形、連番の順で設定名を作る。"""
    try:
        text_data = page.get_text("dict", clip=rect, sort=True)
        pieces: list[str] = []
        for block in text_data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    pieces.append(str(span.get("text", "")))
        extracted = clean_auto_name_text("".join(pieces))
        if extracted:
            return f"P{page_number} {extracted}"
    except Exception:
        pass

    try:
        if any(
            rect_intersects(rect, fitz.Rect(info["bbox"]))
            for info in page.get_image_info(xrefs=True)
            if info.get("bbox") is not None
        ):
            return f"P{page_number} 画像"
    except Exception:
        pass

    try:
        widgets = page.widgets() or []
        if any(rect_intersects(rect, widget.rect) for widget in widgets):
            return f"P{page_number} フォーム"
    except Exception:
        pass

    try:
        annotations = page.annots() or []
        if any(rect_intersects(rect, annotation.rect) for annotation in annotations):
            return f"P{page_number} 注釈"
    except Exception:
        pass

    try:
        if any(
            drawing.get("rect") is not None
            and rect_intersects(rect, fitz.Rect(drawing["rect"]))
            for drawing in page.get_drawings()
        ):
            return f"P{page_number} 図形"
    except Exception:
        pass

    pattern = re.compile(rf"^P{page_number}\s+(\d+)$")
    used_numbers = [
        int(match.group(1))
        for rule in rules
        if (match := pattern.match(rule.name.strip())) is not None
    ]
    serial = max(used_numbers, default=0) + 1
    return f"P{page_number} {serial:02d}"


def _rect_distance(left: fitz.Rect, right: fitz.Rect) -> float:
    return sum(
        abs(a - b)
        for a, b in zip(
            (left.x0, left.y0, left.x1, left.y1),
            (right.x0, right.y0, right.x1, right.y1),
        )
    )


def _safe_page_get_links(page: fitz.Page) -> list[dict]:
    """空または不正なActionでMuPDFの高水準解析が失敗しても列挙を継続する。"""
    try:
        return [dict(link) for link in page.get_links()]
    except Exception:
        return []


def _link_annotation_xrefs(document: fitz.Document, page: fitz.Page) -> list[int]:
    result: list[int] = []
    try:
        annotation_xrefs = page.annot_xrefs()
    except Exception:
        annotation_xrefs = []
    for item in annotation_xrefs:
        try:
            xref = int(item[0])
            value_type, value = document.xref_get_key(xref, "Subtype")
            if value_type == "name" and value == "/Link":
                result.append(xref)
        except (TypeError, ValueError, RuntimeError):
            continue
    return result


def _annotation_rect(
    document: fitz.Document, page: fitz.Page, xref: int
) -> fitz.Rect | None:
    try:
        value_type, value = document.xref_get_key(xref, "Rect")
        if value_type != "array":
            return None
        numbers = [
            float(item)
            for item in re.findall(
                r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?", value
            )[:4]
        ]
        if len(numbers) != 4:
            return None
        return fitz.Rect(numbers) * page.transformation_matrix
    except (TypeError, ValueError, RuntimeError):
        return None


def _appearance_from_xref(
    document: fitz.Document, xref: int
) -> tuple[float, str, str]:
    width = 0.0
    style = "solid"
    color = "#FF0000"
    style_by_pdf_value = {"S": "solid", "D": "dashed", "U": "underline"}
    try:
        width_type, width_value = document.xref_get_key(xref, "BS/W")
        if width_type in {"int", "real"}:
            width = float(width_value)
        else:
            border_type, border_value = document.xref_get_key(xref, "Border")
            if border_type == "array":
                numbers = re.findall(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", border_value)
                if len(numbers) >= 3:
                    width = float(numbers[2])
        style_type, style_value = document.xref_get_key(xref, "BS/S")
        if style_type == "name":
            style = style_by_pdf_value.get(style_value.lstrip("/"), "solid")
        color_type, color_value = document.xref_get_key(xref, "C")
        if color_type == "array":
            components = [
                clamp(float(item), 0.0, 1.0)
                for item in re.findall(
                    r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", color_value
                )[:3]
            ]
            if len(components) == 3:
                color = "#" + "".join(
                    f"{round(component * 255):02X}" for component in components
                )
    except (TypeError, ValueError, RuntimeError):
        pass
    return width, color, style


def _page_link_records(
    document: fitz.Document, page_index: int
) -> list[dict]:
    """get_linksのxrefが0でも、Link列と注釈辞書から実xrefを対応付ける。"""
    page = document[page_index]
    raw_xrefs = _link_annotation_xrefs(document, page)
    raw_rects = {
        xref: rect
        for xref in raw_xrefs
        if (rect := _annotation_rect(document, page, xref)) is not None
    }
    first_rects: dict[int, fitz.Rect] = {}
    try:
        link_object = page.first_link
        while link_object is not None:
            xref = int(getattr(link_object, "xref", 0) or 0)
            if xref > 0:
                first_rects[xref] = fitz.Rect(link_object.rect)
            link_object = link_object.next
    except Exception:
        pass

    records: list[dict] = []
    used_xrefs: set[int] = set()
    get_links_entries = _safe_page_get_links(page)
    for entry in get_links_entries:
        link = dict(entry)
        source_rect = link.get("from")
        rect = fitz.Rect(source_rect) if source_rect is not None else None
        reported_xref = int(link.get("xref", 0) or 0)
        xref = reported_xref if reported_xref > 0 else 0
        if xref <= 0 and rect is not None:
            candidates: list[tuple[float, int]] = []
            for candidate in raw_xrefs:
                if candidate in used_xrefs:
                    continue
                # 未対応アクションが混在するとfirst_link側のxrefもずれる版がある。
                # PDF注釈辞書の /Rect を優先し、Linkオブジェクトは補助に限定する。
                candidate_rect = raw_rects.get(candidate) or first_rects.get(candidate)
                if candidate_rect is not None:
                    candidates.append((_rect_distance(rect, candidate_rect), candidate))
            if candidates:
                distance, candidate = min(candidates)
                if distance <= 2.0:
                    xref = candidate
        if xref > 0:
            used_xrefs.add(xref)
            if rect is None:
                rect = raw_rects.get(xref) or first_rects.get(xref)
        records.append(
            {
                "xref": xref,
                "rect": rect,
                "link": link,
                "appearance": (
                    _appearance_from_xref(document, xref)
                    if xref > 0
                    else (0.0, "#FF0000", "solid")
                ),
            }
        )

    # MuPDFが辞書へ返さない注釈も残す。対応可能なら取込対象にし、
    # 未対応なら青枠表示と診断ログの対象にする。
    for xref in raw_xrefs:
        if xref in used_xrefs:
            continue
        records.append(
            {
                "xref": xref,
                "rect": raw_rects.get(xref) or first_rects.get(xref),
                "link": {},
                "appearance": _appearance_from_xref(document, xref),
            }
        )
    del page
    return records


def existing_links_as_rules(
    document: fitz.Document,
    include_named_destinations: bool = True,
    diagnostics: list[str] | None = None,
) -> list[LinkRule]:
    """ページ移動リンクとURLリンクを、失敗を局所化して編集設定へ変換する。"""
    rules: list[LinkRule] = []
    mode_by_pdf_value = {value: mode for mode, value in HIGHLIGHT_MODE_PDF.items()}
    try:
        resolved_names = document.resolve_names() if include_named_destinations else {}
    except Exception as exc:
        resolved_names = {}
        if diagnostics is not None:
            diagnostics.append(
                "名前付き移動先の一覧を解決できませんでした。"
                f"数値ページとURLの取込は継続します | {type(exc).__name__}: {exc}"
            )
    for page_index in range(document.page_count):
        try:
            records = _page_link_records(document, page_index)
        except Exception as exc:
            if diagnostics is not None:
                diagnostics.append(
                    f"ページ={page_index + 1} | リンク注釈の列挙に失敗 | "
                    f"{type(exc).__name__}: {exc}"
                )
            continue
        supported_index = 0
        for record in records:
            link = record["link"]
            xref = int(record["xref"] or 0)
            rect = record["rect"]
            context = rectangle_context(rect, page_number=page_index + 1)
            if xref > 0:
                context += f" | xref={xref}"
            try:
                if xref <= 0 or rect is None:
                    if diagnostics is not None:
                        diagnostics.append(
                            f"{context or f'ページ={page_index + 1}'} | "
                            "xrefまたはリンク範囲を復元できないため青枠で保持"
                        )
                    continue
                kind = link.get("kind")
                source_named_destination = named_destination_name_from_source(
                    document, xref
                )
                source_named_destination_pdf: str | None = None
                named_page = (
                    resolved_names.get(source_named_destination)
                    if source_named_destination
                    else None
                )
                direct_page = direct_destination_page_from_source(document, xref)
                uri = str(link.get("uri", "")).strip() or uri_from_source(
                    document, xref
                )

                if source_named_destination and include_named_destinations and named_page:
                    link_type = "page"
                    target_page: str | int = source_named_destination
                    source_named_destination_pdf = named_destination_pdf_from_source(
                        document, xref
                    )
                    url = "https://"
                    target_text = (
                        f"{target_page} → {int(named_page.get('page', -1)) + 1}ページ"
                    )
                elif direct_page is not None:
                    link_type = "page"
                    target_page = str(direct_page + 1)
                    url = "https://"
                    target_text = f"{target_page}ページ"
                elif kind == fitz.LINK_GOTO and isinstance(link.get("page"), int) and int(link["page"]) >= 0:
                    link_type = "page"
                    target_page = str(int(link["page"]) + 1)
                    url = "https://"
                    target_text = f"{target_page}ページ"
                elif uri:
                    link_type = "url"
                    target_page = "1"
                    url = uri
                    target_text = "URL"
                else:
                    if diagnostics is not None:
                        diagnostics.append(
                            f"{context} | 種類={action_kind_from_source(document, xref)} | "
                            "ページ移動またはURLとして解釈できないため青枠で保持"
                        )
                    continue

                supported_index += 1
                border_width, border_color, border_style = record["appearance"]
                _, pdf_mode = _safe_xref_get_key(document, xref, "H")
                highlight_mode = mode_by_pdf_value.get(pdf_mode, "invert")
                rules.append(
                    LinkRule(
                        name=f"既存 P{page_index + 1}-{supported_index} {target_text}",
                        page_spec=str(page_index + 1),
                        x=rect.x0,
                        y=rect.y0,
                        width=rect.width,
                        height=rect.height,
                        link_type=link_type,
                        target_page=target_page,
                        url=url,
                        destination_view=destination_view_from_source(
                            document, xref, link
                        ),
                        border_visible=border_width > 0,
                        border_color=border_color,
                        border_width=border_width if border_width > 0 else 1.0,
                        border_style=border_style,
                        highlight_mode=highlight_mode,
                        enabled=True,
                        source_page=page_index,
                        source_xref=xref,
                        source_named_destination=source_named_destination,
                        source_named_destination_pdf=source_named_destination_pdf,
                    )
                )
            except Exception as exc:
                if diagnostics is not None:
                    diagnostics.append(
                        f"{context} | 取込失敗 | {type(exc).__name__}: {exc}"
                    )
    return rules


def existing_link_count(document: fitz.Document) -> int:
    """未対応アクションを含むLink注釈の総数を返す。"""
    total = 0
    for page_index in range(document.page_count):
        try:
            total += len(_page_link_records(document, page_index))
        except Exception:
            try:
                total += len(document[page_index].get_links())
            except Exception:
                continue
    return total


def validate_document_and_rules(
    document: fitz.Document,
    rules: list[LinkRule],
) -> ValidationReport:
    report = ValidationReport()
    if document.page_count < 1:
        report.errors.append("PDFにページがありません。")
        return report

    for page_index in range(document.page_count):
        widgets = document[page_index].widgets()
        if widgets and any(
            widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE for widget in widgets
        ):
            report.warnings.append(
                f"ページ={page_index + 1} | 電子署名欄があります。"
                "PDFを編集すると既存の署名が無効になる可能性があります。"
            )
            break

    enabled = [rule for rule in rules if rule.enabled and not rule.deleted]
    has_existing_changes = any(rule.source_xref is not None for rule in rules)
    if not enabled and not has_existing_changes:
        report.errors.append("有効なリンク設定がありません。")
        return report

    expanded: list[tuple[LinkRule, list[int]]] = []
    for rule in enabled:
        try:
            pages = parse_page_spec(rule.page_spec, document.page_count)
        except ValidationError as exc:
            report.errors.append(f"{rule_context(rule)} | {exc}")
            continue
        rect = rule.rect()
        if (
            not all(math.isfinite(v) for v in (rule.x, rule.y, rule.width, rule.height))
            or rule.width <= 0
            or rule.height <= 0
        ):
            report.errors.append(
                f"{rule_context(rule)} | X・Y・幅・高さは数値で、"
                "幅・高さは0より大きくしてください。"
            )
            continue
        if rule.link_type == "page":
            try:
                resolve_target_page(document, rule.target_page)
            except ValidationError as exc:
                report.errors.append(f"{rule_context(rule)} | {exc}")
        elif rule.link_type == "url":
            if not rule.url.strip():
                report.errors.append(f"{rule_context(rule)} | URLを入力してください。")
        else:
            report.errors.append(f"{rule_context(rule)} | 未対応のリンク種類です。")
        try:
            hex_to_pdf_color(rule.border_color)
        except ValidationError as exc:
            report.errors.append(f"{rule_context(rule)} | {exc}")
        if rule.border_width < 0:
            report.errors.append(f"{rule_context(rule)} | 枠幅は0以上にしてください。")
        if rule.highlight_mode not in HIGHLIGHT_MODE_PDF:
            report.errors.append(
                f"{rule_context(rule)} | 未対応のクリック時表示です。"
            )
        if rule.border_style not in BORDER_STYLE_PDF:
            report.errors.append(f"{rule_context(rule)} | 未対応の枠スタイルです。")
        if rule.destination_view not in DESTINATION_VIEW_LABELS:
            report.errors.append(
                f"{rule_context(rule)} | 未対応のリンク先表示方法です。"
            )
        expanded.append((rule, pages))
        report.insertion_count += len(pages)

    # 追加予定どうしの重なり
    for index, (left_rule, left_pages) in enumerate(expanded):
        left_set = set(left_pages)
        for right_rule, right_pages in expanded[index + 1 :]:
            common = left_set.intersection(right_pages)
            overlapping_pages = []
            for page_index in common:
                if rects_overlap(left_rule.rect(), right_rule.rect()):
                    overlapping_pages.append(page_index)
            if overlapping_pages:
                first_page = min(overlapping_pages) + 1
                report.warnings.append(
                    f"ページ={first_page} | 追加予定どうしが重なります | "
                    f"{rule_context(left_rule, first_page)} | "
                    f"{rule_context(right_rule, first_page)}"
                )

    # 既存リンクとの重なり
    replaced_existing = {
        (rule.source_page, rule.source_xref)
        for rule in rules
        if rule.source_page is not None and rule.source_xref is not None
    }
    warned: set[tuple[str, int]] = set()
    for rule, pages in expanded:
        for page_index in pages:
            try:
                existing_links = _page_link_records(document, page_index)
            except Exception:
                existing_links = []
            if any(
                link.get("rect") is not None
                and rects_overlap(rule.rect(), fitz.Rect(link["rect"]))
                for link in existing_links
                if (page_index, int(link.get("xref", 0))) not in replaced_existing
            ):
                key = (rule.rule_id, page_index)
                if key not in warned:
                    warned.add(key)
                    report.warnings.append(
                        f"ページ={page_index + 1} | 既存リンクと重なります | "
                        f"{rule_context(rule, page_index + 1)}"
                    )
    return report


def _set_link_appearance_and_destination(
    document: fitz.Document,
    xref: int,
    rule: LinkRule,
    target_index: int | None,
    named_destination_pdf: str | None,
) -> None:
    """挿入直後の実xrefへ外観を設定し、保存後のリンク再探索を不要にする。"""
    width = rule.border_width if rule.border_visible else 0.0
    style_code = BORDER_STYLE_PDF[rule.border_style]
    if rule.border_style == "dashed":
        document.xref_set_key(xref, "Border", f"[0 0 {width:g} [3 3]]")
        border_style_value = f"<< /W {width:g} /S {style_code} /D [3 3] >>"
    else:
        document.xref_set_key(xref, "Border", f"[0 0 {width:g}]")
        border_style_value = f"<< /W {width:g} /S {style_code} >>"
    document.xref_set_key(xref, "BS", border_style_value)
    if rule.border_visible:
        red, green, blue = hex_to_pdf_color(rule.border_color)
        document.xref_set_key(xref, "C", f"[{red:g} {green:g} {blue:g}]")
    document.xref_set_key(xref, "H", HIGHLIGHT_MODE_PDF[rule.highlight_mode])
    if rule.link_type == "page" and target_index is not None:
        destination = (
            named_destination_pdf
            if named_destination_pdf is not None
            else destination_array(document, target_index, rule.destination_view)
        )
        document.xref_set_key(
            xref, "A", f"<< /S /GoTo /D {destination} >>"
        )


def _appearance_verification_messages(
    output_path: Path,
    expected: list[tuple[int, int, LinkRule]],
) -> list[str]:
    messages: list[str] = []
    document = fitz.open(output_path)
    try:
        for page_index, xref, rule in expected:
            context = rule_context(rule, page_index + 1) + f" | xref={xref}"
            try:
                subtype_type, subtype = document.xref_get_key(xref, "Subtype")
                if subtype_type != "name" or subtype != "/Link":
                    messages.append(f"{context} | 保存後にLink注釈を確認できません。")
                    continue
                actual_width, actual_color, actual_style = _appearance_from_xref(
                    document, xref
                )
                expected_width = rule.border_width if rule.border_visible else 0.0
                if not math.isclose(actual_width, expected_width, abs_tol=0.01):
                    messages.append(
                        f"{context} | 枠幅の確認不一致: "
                        f"指定={expected_width:g}, 出力={actual_width:g}"
                    )
                if actual_style != rule.border_style:
                    messages.append(
                        f"{context} | 枠スタイルの確認不一致: "
                        f"指定={rule.border_style}, 出力={actual_style}"
                    )
                if rule.border_visible and actual_color.upper() != rule.border_color.upper():
                    messages.append(
                        f"{context} | 枠色の確認不一致: "
                        f"指定={rule.border_color}, 出力={actual_color}"
                    )
                _, highlight = document.xref_get_key(xref, "H")
                expected_highlight = HIGHLIGHT_MODE_PDF[rule.highlight_mode]
                if highlight != expected_highlight:
                    messages.append(
                        f"{context} | ハイライトの確認不一致: "
                        f"指定={expected_highlight}, 出力={highlight}"
                    )
            except Exception as exc:
                messages.append(
                    f"{context} | 外観の保存後確認に失敗 | "
                    f"{type(exc).__name__}: {exc}"
                )
    finally:
        document.close()
    return messages


def apply_rules_to_pdf(
    source_path: Path,
    output_path: Path,
    rules: list[LinkRule],
    named_destination_output: str = "preserve",
) -> tuple[int, list[str]]:
    require_pymupdf()
    if source_path.resolve() == output_path.resolve():
        raise ValidationError("元PDFと同じ場所には保存できません。別名を指定してください。")
    if named_destination_output not in NAMED_DESTINATION_OUTPUT_LABELS:
        raise ValidationError("名前付き移動先の出力方式が正しくありません。")

    document = fitz.open(source_path)
    warnings: list[str] = []
    expected_appearances: list[tuple[int, int, LinkRule]] = []
    applied = 0
    try:
        if document.needs_pass:
            raise ValidationError("パスワード保護されたPDFには対応していません。")
        report = validate_document_and_rules(document, rules)
        if not report.ok:
            raise ValidationError("\n".join(report.errors))
        warnings.extend(report.warnings)

        # 取り込んだ既存リンクは実xrefで直接削除する。get_links() がxref=0を
        # 返すPDFでも、注釈辞書から復元したsource_xrefなら削除できる。
        removed_refs: set[tuple[int, int]] = set()
        for rule in rules:
            if rule.source_page is None or rule.source_xref is None:
                continue
            ref = (rule.source_page, rule.source_xref)
            if ref in removed_refs or not (0 <= rule.source_page < document.page_count):
                continue
            page = document[rule.source_page]
            try:
                if rule.source_xref in _link_annotation_xrefs(document, page):
                    page.delete_link({"xref": rule.source_xref})
                else:
                    warnings.append(
                        f"{rule_context(rule, rule.source_page + 1)} | "
                        f"xref={rule.source_xref} | 元の既存リンクを確認できないため、"
                        "新規リンクとして処理"
                    )
            except Exception as exc:
                warnings.append(
                    f"{rule_context(rule, rule.source_page + 1)} | "
                    f"xref={rule.source_xref} | 既存リンクの削除に失敗 | "
                    f"{type(exc).__name__}: {exc}"
                )
            removed_refs.add(ref)
            del page

        for rule in rules:
            if not rule.enabled or rule.deleted:
                continue
            for page_index in parse_page_spec(rule.page_spec, document.page_count):
                page = document[page_index]
                display_rect = rule.rect()
                before_xrefs = set(_link_annotation_xrefs(document, page))
                link_dict: dict = {
                    "from": display_rect_to_insertion(page, display_rect)
                }
                target_index: int | None = None
                named_destination_pdf: str | None = None
                if rule.link_type == "page":
                    target_index = resolve_target_page(document, rule.target_page)
                    is_original_named_destination = (
                        rule.source_named_destination is not None
                        and str(rule.target_page).strip()
                        == rule.source_named_destination
                    )
                    if (
                        is_original_named_destination
                        or (
                            named_destination_output == "preserve"
                            and not str(rule.target_page).strip().isdigit()
                        )
                    ):
                        named_destination_pdf = named_destination_pdf_for_rule(rule)
                    if target_index == page_index:
                        target_point = display_point_to_insertion(
                            page, fitz.Point(0, 0)
                        )
                    else:
                        target = document[target_index]
                        target_point = display_point_to_insertion(
                            target, fitz.Point(0, 0)
                        )
                        del target
                    link_dict.update(
                        {
                            "kind": fitz.LINK_GOTO,
                            "page": target_index,
                            "to": target_point,
                            "zoom": 0.0,
                        }
                    )
                else:
                    link_dict.update(
                        {"kind": fitz.LINK_URI, "uri": rule.url.strip()}
                    )
                page.insert_link(link_dict)
                new_xrefs = set(_link_annotation_xrefs(document, page)) - before_xrefs
                if len(new_xrefs) == 1:
                    new_xref = new_xrefs.pop()
                    try:
                        _set_link_appearance_and_destination(
                            document,
                            new_xref,
                            rule,
                            target_index,
                            named_destination_pdf,
                        )
                        expected_appearances.append(
                            (page_index, new_xref, copy.deepcopy(rule))
                        )
                    except Exception as exc:
                        warnings.append(
                            f"{rule_context(rule, page_index + 1)} | xref={new_xref} | "
                            f"枠外観の設定に失敗 | {type(exc).__name__}: {exc}"
                        )
                else:
                    warnings.append(
                        f"{rule_context(rule, page_index + 1)} | "
                        "挿入直後のxrefを一意に取得できず、枠外観を確認できません。"
                    )
                applied += 1
                del page

        # xrefを維持するgarbage=0で一度だけ別名保存する。reload_page()、
        # 保存後の書換え、saveIncr()、一時ファイルからの置換は行わない。
        output_path.parent.mkdir(parents=True, exist_ok=True)
        document.save(
            output_path,
            garbage=0,
            deflate=True,
            raise_on_repair=True,
        )
    finally:
        if not document.is_closed:
            document.close()

    if expected_appearances:
        try:
            warnings.extend(
                _appearance_verification_messages(output_path, expected_appearances)
            )
        except Exception as exc:
            warnings.append(
                f"出力PDF={output_path} | 保存後の外観確認に失敗 | "
                f"{type(exc).__name__}: {exc}"
            )
    return applied, warnings


class PDFLinkManagerApp(tk.Tk):
    def __init__(self) -> None:
        require_pymupdf()
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(1080, 760)

        self.document: fitz.Document | None = None
        self.pdf_path: Path | None = None
        self.current_page = 0
        self.rules: list[LinkRule] = []
        self.existing_links_imported = False
        self.selected_rule_id: str | None = None
        self.preview_image: tk.PhotoImage | None = None
        self.preview_scale = 1.0
        self.page_origin = (0.0, 0.0)
        self.drag_mode: str | None = None
        self.drag_anchor = (0.0, 0.0)
        self.drag_original: fitz.Rect | None = None
        self.space_pan_active = False
        self.pan_dragging = False
        self.last_wheel_action = 0.0
        self.current_rect = fitz.Rect(0, 0, MM_TO_PT, MM_TO_PT)
        self.undo_stack: list[list[LinkRule]] = []
        self.redo_stack: list[list[LinkRule]] = []
        self.last_log_path: Path | None = None

        self.status_var = tk.StringVar(value="PDFを開いてください。")
        self.page_number_var = tk.StringVar(value="1")
        self.page_total_var = tk.StringVar(value="/ 0")
        self.unit_var = tk.StringVar(value="mm")
        self.show_existing_var = tk.BooleanVar(value=True)
        self.show_planned_var = tk.BooleanVar(value=True)
        self.snap_var = tk.BooleanVar(value=False)
        self.step_var = tk.StringVar(value="0.1")
        self.border_width_preset_var = tk.StringVar(value="標準")
        self.border_display_var = tk.StringVar(value="ボックスを非表示")
        self.border_style_var = tk.StringVar(value=BORDER_STYLE_LABELS["solid"])
        self.destination_view_var = tk.StringVar(
            value=DESTINATION_VIEW_LABELS["fit_page"]
        )
        self.named_destination_output_var = tk.StringVar(
            value=NAMED_DESTINATION_OUTPUT_LABELS["preserve"]
        )

        self.form_vars: dict[str, tk.Variable] = {
            "name": tk.StringVar(value=""),
            "page_spec": tk.StringVar(value="全ページ"),
            "x": tk.StringVar(value="0"),
            "y": tk.StringVar(value="0"),
            "width": tk.StringVar(value="1"),
            "height": tk.StringVar(value="1"),
            "link_type": tk.StringVar(value="page"),
            "target_page": tk.StringVar(value="1"),
            "url": tk.StringVar(value="https://"),
            "border_color": tk.StringVar(value="#FF0000"),
            "border_width": tk.StringVar(value="1.0"),
            "highlight_mode": tk.StringVar(
                value=HIGHLIGHT_MODE_LABELS["none"]
            ),
            "enabled": tk.BooleanVar(value=True),
        }

        self._build_menu()
        self._build_ui()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self.close_document)

    def save_diagnostics(
        self,
        operation: str,
        messages: Iterable[str],
        *,
        output_path: Path | None = None,
        exception: BaseException | None = None,
        source_path: Path | None = None,
    ) -> Path | None:
        try:
            self.last_log_path = write_diagnostic_log(
                operation,
                messages,
                source_path=source_path or self.pdf_path,
                output_path=output_path,
                exception=exception,
            )
        except OSError:
            self.last_log_path = None
        return self.last_log_path

    @staticmethod
    def diagnostic_message(count: int, log_path: Path | None) -> str:
        if log_path is None:
            return f"{count}件の詳細があります。ログを保存できませんでした。"
        return f"{count}件の詳細をエラーログへ保存しました。\n\n{log_path}"

    def _build_menu(self) -> None:
        shortcut_name = "Command" if sys.platform == "darwin" else "Ctrl"
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(
            label="PDFを開く...",
            command=self.open_pdf,
            accelerator=f"{shortcut_name}+O",
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="PDFへ適用して別名保存...",
            command=self.export_pdf,
            accelerator=f"{shortcut_name}+S",
        )
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.close_document)
        menu.add_cascade(label="ファイル", menu=file_menu)

        edit_menu = tk.Menu(menu, tearoff=False)
        edit_menu.add_command(
            label="元に戻す", command=self.undo, accelerator=f"{shortcut_name}+Z"
        )
        edit_menu.add_command(
            label="やり直す", command=self.redo, accelerator=f"{shortcut_name}+Y"
        )
        edit_menu.add_separator()
        edit_menu.add_command(label="設定を複製", command=self.duplicate_rule)
        edit_menu.add_command(label="設定を削除", command=self.delete_rule)
        menu.add_cascade(label="編集", menu=edit_menu)

        view_menu = tk.Menu(menu, tearoff=False)
        view_menu.add_checkbutton(
            label="既存リンクを表示",
            variable=self.show_existing_var,
            command=self.draw_overlays,
        )
        view_menu.add_checkbutton(
            label="追加予定リンクを表示",
            variable=self.show_planned_var,
            command=self.draw_overlays,
        )
        view_menu.add_separator()
        view_menu.add_command(label="拡大", command=lambda: self.change_zoom(1.2))
        view_menu.add_command(label="縮小", command=lambda: self.change_zoom(1 / 1.2))
        view_menu.add_command(label="ページ全体", command=self.fit_page)
        menu.add_cascade(label="表示", menu=view_menu)

        tools_menu = tk.Menu(menu, tearoff=False)
        tools_menu.add_command(label="設定を検査", command=self.validate_all)
        tools_menu.add_command(label="PDFへ適用", command=self.export_pdf)
        menu.add_cascade(label="ツール", menu=tools_menu)

        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="使い方", command=self.show_help)
        help_menu.add_command(label="バージョン情報", command=self.show_about)
        menu.add_cascade(label="ヘルプ", menu=help_menu)
        self.config(menu=menu)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.configure("Primary.TButton", padding=(10, 5), font=("", 10, "bold"))

        toolbar = ttk.Frame(self, padding=(8, 6))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="PDFを開く", command=self.open_pdf).pack(side="left")
        ttk.Button(toolbar, text="前ページ", command=lambda: self.change_page(-1)).pack(
            side="left", padx=(8, 2)
        )
        self.page_number_entry = ttk.Spinbox(
            toolbar,
            textvariable=self.page_number_var,
            from_=1,
            to=1,
            width=6,
            justify="right",
            command=self.go_to_page,
        )
        self.page_number_entry.pack(side="left", padx=(4, 1))
        self.page_number_entry.bind("<Return>", self.go_to_page)
        ttk.Label(toolbar, textvariable=self.page_total_var, width=7, anchor="w").pack(
            side="left", padx=(0, 2)
        )
        ttk.Button(toolbar, text="移動", command=self.go_to_page).pack(side="left")
        ttk.Button(toolbar, text="次ページ", command=lambda: self.change_page(1)).pack(
            side="left", padx=(6, 2)
        )
        ttk.Button(toolbar, text="縮小", command=lambda: self.change_zoom(1 / 1.2)).pack(
            side="left", padx=(12, 2)
        )
        ttk.Button(toolbar, text="拡大", command=lambda: self.change_zoom(1.2)).pack(
            side="left"
        )
        ttk.Button(toolbar, text="全体表示", command=self.fit_page).pack(
            side="left", padx=4
        )
        ttk.Button(
            toolbar,
            text="設定を検査",
            command=self.validate_all,
            style="Primary.TButton",
        ).pack(
            side="right", padx=4
        )
        ttk.Button(
            toolbar,
            text="PDFへ適用して保存",
            command=self.export_pdf,
            style="Primary.TButton",
        ).pack(
            side="right"
        )
        ttk.Combobox(
            toolbar,
            textvariable=self.named_destination_output_var,
            values=tuple(NAMED_DESTINATION_OUTPUT_LABELS.values()),
            state="readonly",
            width=16,
        ).pack(side="right", padx=(4, 2))
        ttk.Label(toolbar, text="名前付き出力").pack(side="right", padx=(8, 0))

        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        preview_frame = ttk.Frame(paned)
        self.settings_frame = ttk.Frame(paned, width=360)
        paned.add(preview_frame, weight=2)
        paned.add(self.settings_frame, weight=1)

        def set_initial_sash(attempt: int = 0) -> None:
            pane_width = paned.winfo_width()
            if pane_width < 600 and attempt < 20:
                self.after(100, lambda: set_initial_sash(attempt + 1))
                return
            if pane_width >= 600:
                paned.sashpos(0, int(pane_width * 0.62))

        self.after_idle(set_initial_sash)

        canvas_box = ttk.Frame(preview_frame)
        canvas_box.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(
            canvas_box,
            background="#454545",
            highlightthickness=0,
            cursor="crosshair",
            takefocus=True,
        )
        x_scroll = ttk.Scrollbar(canvas_box, orient="horizontal", command=self.canvas.xview)
        y_scroll = ttk.Scrollbar(canvas_box, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        canvas_box.rowconfigure(0, weight=1)
        canvas_box.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", self.canvas_press)
        self.canvas.bind("<B1-Motion>", self.canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.canvas_release)
        self.canvas.bind("<MouseWheel>", self.canvas_mouse_wheel)
        self.canvas.bind("<Control-MouseWheel>", self.canvas_mouse_wheel_zoom)
        self.canvas.bind("<Button-4>", self.canvas_mouse_wheel)
        self.canvas.bind("<Button-5>", self.canvas_mouse_wheel)
        if sys.platform == "darwin":
            self.canvas.bind("<Command-MouseWheel>", self.canvas_mouse_wheel_zoom)
        self.canvas.bind("<Configure>", lambda _event: self._on_canvas_configure())
        self.canvas.bind("<Left>", lambda event: self.keyboard_move(-1, 0, event))
        self.canvas.bind("<Right>", lambda event: self.keyboard_move(1, 0, event))
        self.canvas.bind("<Up>", lambda event: self.keyboard_move(0, -1, event))
        self.canvas.bind("<Down>", lambda event: self.keyboard_move(0, 1, event))

        hint = ttk.Label(
            preview_frame,
            text=(
                "範囲: ドラッグ  /  移動: 枠内ドラッグ・矢印キー  /  "
                "サイズ: ■をドラッグ\n"
                "Space+ドラッグ: 表示移動  /  ホイール: ページ移動  /  "
                "Ctrl+ホイール: カーソル位置を中心に拡大縮小"
            ),
            anchor="center",
            justify="center",
        )
        hint.pack(fill="x", pady=(4, 0))

        self.settings_frame.columnconfigure(0, weight=1)
        self.settings_frame.rowconfigure(1, weight=1)
        self.edit_section = ttk.Labelframe(
            self.settings_frame, text="リンク設定", padding=(8, 4)
        )
        self.list_section = ttk.Labelframe(
            self.settings_frame, text="登録予定一覧", padding=(6, 4)
        )
        self.edit_section.grid(row=0, column=0, sticky="ew")
        self.list_section.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        edit_body = ttk.Frame(self.edit_section)
        edit_body.pack(fill="both", expand=True)
        self._build_edit_tab(edit_body)
        self._build_list_tab(self.list_section)

        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(fill="x", side="bottom")

    def _build_edit_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        row = 0

        def label(text: str) -> None:
            nonlocal row
            ttk.Label(parent, text=text).grid(row=row, column=0, sticky="w", pady=3)

        label("設定名")
        ttk.Entry(parent, textvariable=self.form_vars["name"]).grid(
            row=row, column=1, columnspan=3, sticky="ew", pady=3
        )
        row += 1
        label("対象ページ")
        ttk.Entry(parent, textvariable=self.form_vars["page_spec"]).grid(
            row=row, column=1, columnspan=3, sticky="ew", pady=3
        )
        row += 1
        ttk.Label(
            parent, text="例: 全ページ / 1-5,8,10 / 奇数 / 偶数", foreground="#555555"
        ).grid(row=row, column=1, columnspan=3, sticky="w")
        row += 1

        separator = ttk.Separator(parent)
        separator.grid(row=row, column=0, columnspan=4, sticky="ew", pady=8)
        row += 1
        ttk.Label(parent, text="位置と大きさ", font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(parent, text="単位").grid(row=row, column=2, sticky="e")
        unit_box = ttk.Combobox(
            parent,
            textvariable=self.unit_var,
            values=("mm", "pt"),
            width=5,
            state="readonly",
        )
        unit_box.grid(row=row, column=3, sticky="e")
        unit_box.bind("<<ComboboxSelected>>", self.on_unit_changed)
        row += 1

        position_frame = ttk.Frame(parent)
        position_frame.grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=(2, 0)
        )
        for position_column in range(3):
            position_frame.columnconfigure(position_column, weight=1)

        xy_frame = ttk.Frame(position_frame)
        size_frame = ttk.Frame(position_frame)
        adjustment_frame = ttk.Frame(position_frame)
        xy_frame.columnconfigure(0, minsize=76)
        size_frame.columnconfigure(0, minsize=40)
        xy_frame.grid(row=0, column=0, sticky="w", padx=(0, 10))
        size_frame.grid(row=0, column=1, sticky="w", padx=10)
        adjustment_frame.grid(row=0, column=2, sticky="w", padx=(10, 0))

        def add_position_field(
            container: ttk.Frame,
            field_name: str,
            text: str,
            field_row: int,
        ) -> None:
            ttk.Label(container, text=text, anchor="e").grid(
                row=field_row, column=0, sticky="e", pady=3
            )
            spin = ttk.Spinbox(
                container,
                textvariable=self.form_vars[field_name],
                from_=-9999,
                to=9999,
                increment=0.1,
                width=9,
                justify="left",
                command=self.form_rect_changed,
            )
            spin.grid(
                row=field_row,
                column=1,
                sticky="w",
                padx=(6, 0),
                pady=3,
            )
            spin.bind("<KeyRelease>", lambda _event: self.form_rect_changed())
            spin.bind("<FocusOut>", lambda _event: self.form_rect_changed())

        add_position_field(xy_frame, "x", "X（左から）", 0)
        add_position_field(xy_frame, "y", "Y（上から）", 1)
        add_position_field(size_frame, "width", "幅", 0)
        add_position_field(size_frame, "height", "高さ", 1)
        ttk.Label(adjustment_frame, text="増減量", anchor="e").grid(
            row=0, column=0, sticky="e", padx=(0, 6), pady=3
        )
        ttk.Combobox(
            adjustment_frame,
            textvariable=self.step_var,
            values=("0.1", "0.5", "1", "5", "10"),
            width=6,
            state="readonly",
        ).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(
            adjustment_frame, text="整数へ吸着", variable=self.snap_var
        ).grid(
            row=1, column=0, columnspan=2, sticky="e", pady=3
        )
        row += 1

        ttk.Separator(parent).grid(row=row, column=0, columnspan=4, sticky="ew", pady=8)
        row += 1
        ttk.Label(parent, text="リンク先", font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky="w"
        )
        row += 1
        destination_frame = ttk.Frame(parent)
        destination_frame.grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=3
        )
        destination_frame.columnconfigure(5, weight=1)
        ttk.Label(destination_frame, text="種類").grid(
            row=0, column=0, sticky="e"
        )
        type_box = ttk.Combobox(
            destination_frame,
            textvariable=self.form_vars["link_type"],
            values=("page", "url"),
            state="readonly",
            width=8,
        )
        type_box.grid(row=0, column=1, sticky="w", padx=(4, 10))
        type_box.bind("<<ComboboxSelected>>", lambda _event: self.update_link_type_state())
        ttk.Label(destination_frame, text="PDF内ページ").grid(
            row=0, column=2, sticky="e"
        )
        self.target_entry = ttk.Entry(
            destination_frame,
            textvariable=self.form_vars["target_page"],
            width=16,
            justify="left",
        )
        self.target_entry.grid(row=0, column=3, sticky="w", padx=(4, 10))
        ttk.Label(destination_frame, text="URL").grid(
            row=0, column=4, sticky="e"
        )
        self.url_entry = ttk.Entry(
            destination_frame, textvariable=self.form_vars["url"], width=12
        )
        self.url_entry.grid(row=0, column=5, sticky="ew", padx=(4, 0))
        ttk.Label(destination_frame, text="表示方法").grid(
            row=1, column=0, sticky="e", pady=(4, 0)
        )
        self.destination_view_box = ttk.Combobox(
            destination_frame,
            textvariable=self.destination_view_var,
            values=tuple(DESTINATION_VIEW_LABELS.values()),
            state="readonly",
            width=18,
        )
        self.destination_view_box.grid(
            row=1, column=1, columnspan=3, sticky="w", padx=(4, 10), pady=(4, 0)
        )
        row += 1

        ttk.Separator(parent).grid(row=row, column=0, columnspan=4, sticky="ew", pady=8)
        row += 1
        ttk.Label(parent, text="枠の外観", font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky="w"
        )
        row += 1

        appearance_frame = ttk.Frame(parent)
        appearance_frame.grid(
            row=row, column=0, columnspan=4, sticky="ew", pady=3
        )
        for appearance_column in range(3):
            appearance_frame.columnconfigure(appearance_column, weight=1)

        color_frame = ttk.Frame(appearance_frame)
        width_frame = ttk.Frame(appearance_frame)
        display_frame = ttk.Frame(appearance_frame)
        style_frame = ttk.Frame(appearance_frame)
        highlight_frame = ttk.Frame(appearance_frame)
        color_frame.grid(row=0, column=0, sticky="w", padx=(0, 8))
        width_frame.grid(row=0, column=1, columnspan=2, sticky="w")
        display_frame.grid(row=1, column=0, sticky="w", pady=(6, 0))
        style_frame.grid(row=1, column=1, sticky="w", pady=(6, 0), padx=8)
        highlight_frame.grid(row=1, column=2, sticky="e", pady=(6, 0))

        ttk.Label(color_frame, text="枠色").grid(row=0, column=0, sticky="e")
        ttk.Entry(
            color_frame, textvariable=self.form_vars["border_color"], width=8
        ).grid(row=0, column=1, sticky="w", padx=(4, 4))
        ttk.Button(
            color_frame, text="選択", command=self.choose_color
        ).grid(row=0, column=2, sticky="w")

        ttk.Label(width_frame, text="枠幅").grid(row=0, column=0, sticky="e")
        border_width_entry = ttk.Entry(
            width_frame,
            textvariable=self.form_vars["border_width"],
            width=4,
            justify="right",
        )
        border_width_entry.grid(row=0, column=1, sticky="w", padx=(4, 4))
        border_width_entry.bind(
            "<KeyRelease>", lambda _event: self.sync_border_width_preset()
        )
        border_width_entry.bind(
            "<FocusOut>", lambda _event: self.sync_border_width_preset()
        )
        border_preset_box = ttk.Combobox(
            width_frame,
            textvariable=self.border_width_preset_var,
            values=tuple(BORDER_WIDTH_PRESETS),
            state="readonly",
            width=6,
        )
        border_preset_box.grid(row=0, column=2, sticky="w")
        border_preset_box.bind(
            "<<ComboboxSelected>>", self.apply_border_width_preset
        )

        ttk.Label(display_frame, text="リンクの種類").grid(
            row=0, column=0, sticky="e"
        )
        border_display_box = ttk.Combobox(
            display_frame,
            textvariable=self.border_display_var,
            values=BORDER_DISPLAY_LABELS,
            state="readonly",
            width=14,
        )
        border_display_box.grid(row=0, column=1, sticky="w", padx=(4, 0))
        border_display_box.bind(
            "<<ComboboxSelected>>", self.update_border_display_state
        )

        ttk.Label(style_frame, text="スタイル").grid(
            row=0, column=0, sticky="e"
        )
        self.border_style_box = ttk.Combobox(
            style_frame,
            textvariable=self.border_style_var,
            values=tuple(BORDER_STYLE_LABELS.values()),
            state="disabled",
            width=6,
        )
        self.border_style_box.grid(row=0, column=1, sticky="w", padx=(4, 0))

        ttk.Label(highlight_frame, text="ハイライト").grid(
            row=0, column=0, sticky="e"
        )
        ttk.Combobox(
            highlight_frame,
            textvariable=self.form_vars["highlight_mode"],
            values=tuple(HIGHLIGHT_MODE_LABELS.values()),
            state="readonly",
            width=10,
        ).grid(row=0, column=1, sticky="w", padx=(4, 0))
        row += 1
        ttk.Checkbutton(
            parent, text="この設定を有効にする", variable=self.form_vars["enabled"]
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 3))
        row += 1

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=row, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Button(button_frame, text="新規として追加", command=self.add_rule).pack(
            side="left"
        )
        ttk.Button(
            button_frame, text="選択設定を更新", command=self.update_selected_rule
        ).pack(side="left", padx=6)
        ttk.Button(button_frame, text="入力を初期化", command=self.reset_form).pack(
            side="right"
        )
        parent.rowconfigure(row + 1, weight=1)
        self.update_link_type_state()
        self.update_border_display_state()

    def _build_list_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        columns = ("enabled", "pages", "position", "type", "target")
        self.tree = ttk.Treeview(
            parent, columns=columns, show="tree headings", selectmode="browse"
        )
        self.tree.heading("#0", text="設定名")
        self.tree.heading("enabled", text="有効")
        self.tree.heading("pages", text="対象ページ")
        self.tree.heading(
            "position", text=f"X, Y, 幅×高さ ({self.unit_var.get()})"
        )
        self.tree.heading("type", text="種類")
        self.tree.heading("target", text="リンク先")
        self.tree.column("#0", width=95)
        self.tree.column("enabled", width=40, anchor="center")
        self.tree.column("pages", width=78)
        self.tree.column("position", width=130)
        self.tree.column("type", width=48, anchor="center")
        self.tree.column("target", width=110)
        y_scroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        self.tree.configure(
            yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set
        )
        self.tree.grid(row=1, column=0, sticky="nsew")
        y_scroll.grid(row=1, column=1, sticky="ns")
        x_scroll.grid(row=2, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", lambda _event: self.update_selected_rule())

        buttons = ttk.Frame(parent)
        buttons.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Button(buttons, text="複製", command=self.duplicate_rule).pack(side="left")
        ttk.Button(buttons, text="削除", command=self.delete_rule).pack(
            side="left", padx=(4, 0)
        )
        ttk.Button(buttons, text="有効切替", command=self.toggle_rule).pack(
            side="left", padx=(4, 0)
        )
        ttk.Button(
            buttons, text="登録予定の保存", command=self.save_settings
        ).pack(
            side="left", padx=(4, 0)
        )
        ttk.Button(
            buttons, text="登録予定の読み込み", command=self.load_settings
        ).pack(
            side="left", padx=(4, 0)
        )

    def _bind_shortcuts(self) -> None:
        self.bind_all("<Control-o>", lambda _event: self.open_pdf())
        self.bind_all("<Control-s>", lambda _event: self.export_pdf())
        self.bind_all("<Control-z>", lambda _event: self.undo())
        self.bind_all("<Control-y>", lambda _event: self.redo())
        if sys.platform == "darwin":
            self.bind_all("<Command-o>", lambda _event: self.open_pdf())
            self.bind_all("<Command-s>", lambda _event: self.export_pdf())
            self.bind_all("<Command-z>", lambda _event: self.undo())
            self.bind_all("<Command-y>", lambda _event: self.redo())
        self.bind_all("<KeyPress-space>", self.space_pan_key_press)
        self.bind_all("<KeyRelease-space>", self.space_pan_key_release)

    def close_document(self) -> None:
        if self.document is not None:
            self.document.close()
            self.document = None
        self.destroy()

    def open_pdf(self) -> None:
        filename = filedialog.askopenfilename(
            title="PDFを開く", filetypes=[("PDFファイル", "*.pdf"), ("すべて", "*.*")]
        )
        if not filename:
            return
        try:
            document = fitz.open(filename)
            if document.needs_pass:
                document.close()
                raise ValidationError("パスワード保護されたPDFには対応していません。")
            if document.page_count < 1:
                document.close()
                raise ValidationError("ページのないPDFです。")
        except Exception as exc:
            log_path = self.save_diagnostics(
                "PDFを開く",
                [f"PDFを開けません | {type(exc).__name__}: {exc}"],
                source_path=Path(filename),
                exception=exc,
            )
            messagebox.showerror(
                APP_NAME,
                "PDFを開けませんでした。\n\n"
                + self.diagnostic_message(1, log_path),
            )
            return
        import_diagnostics: list[str] = []
        try:
            link_count = existing_link_count(document)
        except Exception as exc:
            link_count = 0
            import_diagnostics.append(
                f"既存リンク件数の確認に失敗 | {type(exc).__name__}: {exc}"
            )
        import_existing = False
        if link_count:
            import_existing = messagebox.askyesno(
                "既存リンク",
                f"このPDFには既存リンクが{link_count}件あります。\n\n"
                "リンク情報を登録しますか？\n\n"
                "はい: 登録予定一覧へ入れて編集対象にします。\n"
                "いいえ: 青枠で表示し、加工・修正せずそのまま残します。",
            )
        imported_rules: list[LinkRule] = []
        if import_existing:
            try:
                imported_rules = existing_links_as_rules(
                    document, diagnostics=import_diagnostics
                )
            except Exception as exc:
                # 取込処理だけの失敗でPDF自体を閉じない。既存注釈は変更せず、
                # プレビュー上の青枠として利用を継続できるようにする。
                import_diagnostics.append(
                    f"既存リンク取込の全体処理に失敗 | "
                    f"{type(exc).__name__}: {exc}"
                )
        if self.document is not None:
            self.document.close()
        self.document = document
        self.pdf_path = Path(filename)
        self.current_page = 0
        self.existing_links_imported = import_existing
        self.rules = imported_rules
        self.selected_rule_id = None
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.refresh_tree()
        self.reset_form()
        self.after_idle(self.fit_page)
        if not link_count:
            link_status = "既存リンクはありません。"
        elif import_existing:
            link_status = (
                f"既存リンク{len(self.rules)}件を登録予定一覧へ読み込みました。"
            )
            if len(self.rules) < link_count:
                link_status += f" 未対応の{link_count - len(self.rules)}件は青枠で保持します。"
        else:
            link_status = f"既存リンク{link_count}件は青枠のまま保持します。"
        self.status_var.set(
            f"{self.pdf_path.name} — {self.document.page_count}ページを開きました。"
            f" {link_status}"
        )
        if import_diagnostics:
            log_path = self.save_diagnostics(
                "既存リンクの取込",
                import_diagnostics,
                source_path=self.pdf_path,
            )
            messagebox.showwarning(
                "既存リンクの取込結果",
                "PDFは開きました。取込できなかったリンクは変更せず保持します。\n\n"
                + self.diagnostic_message(len(import_diagnostics), log_path),
            )

    def render_current_page(
        self,
        *,
        page_origin: tuple[float, float] | None = None,
        view_origin: tuple[float, float] | None = None,
    ) -> None:
        if self.document is None:
            self.canvas.delete("all")
            return
        page = self.document[self.current_page]
        matrix = fitz.Matrix(self.preview_scale, self.preview_scale)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        self.preview_image = tk.PhotoImage(data=pixmap.tobytes("ppm"))
        margin = 20
        x, y = page_origin if page_origin is not None else (margin, margin)
        view_left, view_top = view_origin if view_origin is not None else (0.0, 0.0)
        self.page_origin = (x, y)
        self.canvas.delete("all")
        self.canvas.create_image(x, y, image=self.preview_image, anchor="nw", tags="page")
        viewport_width = max(1, self.canvas.winfo_width())
        viewport_height = max(1, self.canvas.winfo_height())
        scrollregion = (
            min(0.0, x - margin, view_left),
            min(0.0, y - margin, view_top),
            max(
                float(viewport_width),
                x + pixmap.width + margin,
                view_left + viewport_width,
            ),
            max(
                float(viewport_height),
                y + pixmap.height + margin,
                view_top + viewport_height,
            ),
        )
        self.canvas.configure(scrollregion=scrollregion)
        self.page_number_var.set(str(self.current_page + 1))
        self.page_total_var.set(f"/ {self.document.page_count}")
        self.page_number_entry.configure(to=self.document.page_count)
        self.draw_overlays()
        scroll_width = max(1.0, scrollregion[2] - scrollregion[0])
        scroll_height = max(1.0, scrollregion[3] - scrollregion[1])
        self.canvas.xview_moveto((view_left - scrollregion[0]) / scroll_width)
        self.canvas.yview_moveto((view_top - scrollregion[1]) / scroll_height)

    def fit_page(self) -> None:
        if self.document is None:
            return
        self.update_idletasks()
        page = self.document[self.current_page]
        available_width = max(100, self.canvas.winfo_width() - 50)
        available_height = max(100, self.canvas.winfo_height() - 50)
        self.preview_scale = clamp(
            min(
                available_width / page.rect.width,
                available_height / page.rect.height,
            ),
            0.15,
            4.0,
        )
        self.render_current_page()

    def _on_canvas_configure(self) -> None:
        if self.document is not None and self.preview_image is None:
            self.fit_page()

    def change_zoom(
        self,
        factor: float,
        anchor: tuple[float, float] | None = None,
    ) -> None:
        if self.document is None:
            return
        old_scale = self.preview_scale
        new_scale = clamp(old_scale * factor, 0.15, 5.0)
        if math.isclose(old_scale, new_scale, abs_tol=1e-9):
            return
        if self.preview_image is None:
            self.preview_scale = new_scale
            self.render_current_page()
            return
        if anchor is None:
            anchor = (
                self.canvas.winfo_width() / 2,
                self.canvas.winfo_height() / 2,
            )
        view_origin = (
            self.canvas.canvasx(0),
            self.canvas.canvasy(0),
        )
        content_anchor = (
            self.canvas.canvasx(anchor[0]),
            self.canvas.canvasy(anchor[1]),
        )
        new_page_origin = zoomed_page_origin(
            self.page_origin,
            content_anchor,
            old_scale,
            new_scale,
        )
        self.preview_scale = new_scale
        self.render_current_page(
            page_origin=new_page_origin,
            view_origin=view_origin,
        )

    def pointer_is_over_canvas(self) -> bool:
        try:
            pointer_x, pointer_y = self.winfo_pointerxy()
            return self.winfo_containing(pointer_x, pointer_y) is self.canvas
        except tk.TclError:
            return False

    def space_pan_key_press(self, _event: tk.Event) -> str | None:
        if self.document is None or not self.pointer_is_over_canvas():
            return None
        self.space_pan_active = True
        if not self.pan_dragging:
            self.canvas.configure(cursor="hand2")
        return "break"

    def space_pan_key_release(self, _event: tk.Event) -> str | None:
        if not self.space_pan_active:
            return None
        self.space_pan_active = False
        if not self.pan_dragging:
            self.canvas.configure(cursor="crosshair")
        return "break"

    def canvas_mouse_wheel_zoom(self, event: tk.Event) -> str:
        return self.canvas_mouse_wheel(event, force_zoom=True)

    def canvas_mouse_wheel(
        self, event: tk.Event, force_zoom: bool = False
    ) -> str:
        if self.document is None:
            return "break"
        direction = 0
        if getattr(event, "num", None) == 4:
            direction = 1
        elif getattr(event, "num", None) == 5:
            direction = -1
        else:
            delta = int(getattr(event, "delta", 0) or 0)
            direction = 1 if delta > 0 else -1 if delta < 0 else 0
        if direction == 0:
            return "break"

        zooming = force_zoom or bool(int(getattr(event, "state", 0)) & 0x0004)
        now = time.monotonic()
        minimum_interval = 0.06 if zooming else 0.16
        if now - self.last_wheel_action < minimum_interval:
            return "break"
        self.last_wheel_action = now
        if zooming:
            self.change_zoom(
                1.12 if direction > 0 else 1 / 1.12,
                anchor=(
                    float(getattr(event, "x", self.canvas.winfo_width() / 2)),
                    float(getattr(event, "y", self.canvas.winfo_height() / 2)),
                ),
            )
        else:
            self.change_page(-1 if direction > 0 else 1)
        return "break"

    def change_page(self, delta: int) -> None:
        if self.document is None:
            return
        new_page = clamp(
            self.current_page + delta, 0, self.document.page_count - 1
        )
        if int(new_page) != self.current_page:
            self.current_page = int(new_page)
            self.render_current_page()

    def go_to_page(self, _event: tk.Event | None = None) -> str | None:
        if self.document is None:
            return "break" if _event is not None else None
        try:
            requested = int(self.page_number_var.get())
        except ValueError:
            self.page_number_var.set(str(self.current_page + 1))
            self.bell()
            return "break" if _event is not None else None
        requested = int(clamp(requested, 1, self.document.page_count))
        self.current_page = requested - 1
        self.page_number_var.set(str(requested))
        self.render_current_page()
        return "break" if _event is not None else None

    def page_to_canvas(self, rect: fitz.Rect) -> tuple[float, float, float, float]:
        ox, oy = self.page_origin
        return (
            ox + rect.x0 * self.preview_scale,
            oy + rect.y0 * self.preview_scale,
            ox + rect.x1 * self.preview_scale,
            oy + rect.y1 * self.preview_scale,
        )

    def canvas_to_page(self, x: float, y: float) -> fitz.Point:
        ox, oy = self.page_origin
        canvas_x = self.canvas.canvasx(x)
        canvas_y = self.canvas.canvasy(y)
        return fitz.Point(
            (canvas_x - ox) / self.preview_scale,
            (canvas_y - oy) / self.preview_scale,
        )

    def draw_overlays(self) -> None:
        self.canvas.delete("overlay")
        if self.document is None:
            return
        page = self.document[self.current_page]
        if self.show_existing_var.get():
            editable_xrefs = {
                rule.source_xref
                for rule in self.rules
                if rule.source_page == self.current_page
                and rule.source_xref is not None
                and not rule.deleted
            }
            try:
                existing_records = _page_link_records(
                    self.document, self.current_page
                )
            except Exception:
                existing_records = []
            for link in existing_records:
                if link.get("rect") is None:
                    continue
                if int(link.get("xref", 0)) in editable_xrefs:
                    continue
                coords = self.page_to_canvas(fitz.Rect(link["rect"]))
                self.canvas.create_rectangle(
                    *coords,
                    outline="#00A7FF",
                    width=2,
                    dash=(5, 3),
                    tags="overlay",
                )

        if self.show_planned_var.get():
            for rule in self.rules:
                if not rule.enabled or rule.deleted:
                    continue
                try:
                    if self.current_page not in parse_page_spec(
                        rule.page_spec, self.document.page_count
                    ):
                        continue
                except ValidationError:
                    continue
                selected = rule.rule_id == self.selected_rule_id
                self.canvas.create_rectangle(
                    *self.page_to_canvas(rule.rect()),
                    outline="#FF7A00" if selected else "#F2C300",
                    width=3 if selected else 2,
                    tags="overlay",
                )

        # 現在編集中の範囲
        coords = self.page_to_canvas(self.current_rect)
        self.canvas.create_rectangle(
            *coords, outline="#FF3030", width=2, tags="overlay"
        )
        for x, y in (
            (coords[0], coords[1]),
            (coords[2], coords[1]),
            (coords[0], coords[3]),
            (coords[2], coords[3]),
        ):
            self.canvas.create_rectangle(
                x - HANDLE_RADIUS,
                y - HANDLE_RADIUS,
                x + HANDLE_RADIUS,
                y + HANDLE_RADIUS,
                fill="#FFFFFF",
                outline="#B00000",
                tags="overlay",
            )

    def _hit_test(self, point: fitz.Point) -> str:
        threshold = HANDLE_RADIUS / self.preview_scale * 1.8
        rect = self.current_rect
        corners = {
            "resize_nw": fitz.Point(rect.x0, rect.y0),
            "resize_ne": fitz.Point(rect.x1, rect.y0),
            "resize_sw": fitz.Point(rect.x0, rect.y1),
            "resize_se": fitz.Point(rect.x1, rect.y1),
        }
        for name, corner in corners.items():
            if abs(point.x - corner.x) <= threshold and abs(point.y - corner.y) <= threshold:
                return name
        if rect.contains(point):
            return "move"
        return "new"

    def _rule_at_preview_point(self, point: fitz.Point) -> LinkRule | None:
        """Return the topmost visible planned rule containing the page point."""
        if self.document is None or not self.show_planned_var.get():
            return None
        for rule in reversed(self.rules):
            if not rule.enabled or rule.deleted:
                continue
            try:
                pages = parse_page_spec(rule.page_spec, self.document.page_count)
            except ValidationError:
                continue
            if self.current_page in pages and rule.rect().contains(point):
                return rule
        return None

    def _select_rule_from_preview(self, point: fitz.Point) -> bool:
        rule = self._rule_at_preview_point(point)
        if rule is None:
            return False
        self.selected_rule_id = rule.rule_id
        if self.tree.exists(rule.rule_id):
            self.tree.selection_set(rule.rule_id)
            self.tree.focus(rule.rule_id)
            self.tree.see(rule.rule_id)
        self.load_rule_into_form(rule)
        self.status_var.set(f"「{rule.name}」を登録予定一覧から選択しました。")
        return True

    def canvas_press(self, event: tk.Event) -> None:
        if self.document is None:
            return
        self.canvas.focus_set()
        if self.space_pan_active:
            self.canvas.scan_mark(event.x, event.y)
            self.pan_dragging = True
            self.drag_mode = "pan"
            self.canvas.configure(cursor="fleur")
            return
        point = self.canvas_to_page(event.x, event.y)
        page_rect = self.document[self.current_page].rect
        if not page_rect.contains(point):
            return
        self._select_rule_from_preview(point)
        self.drag_mode = self._hit_test(point)
        self.drag_anchor = (point.x, point.y)
        self.drag_original = fitz.Rect(self.current_rect)
        if self.drag_mode == "new":
            self.current_rect = fitz.Rect(point.x, point.y, point.x + 0.1, point.y + 0.1)
        self.draw_overlays()

    def canvas_drag(self, event: tk.Event) -> None:
        if self.pan_dragging and self.drag_mode == "pan":
            self.canvas.scan_dragto(event.x, event.y, gain=1)
            return
        if self.document is None or self.drag_mode is None or self.drag_original is None:
            return
        point = self.canvas_to_page(event.x, event.y)
        page_rect = self.document[self.current_page].rect
        point.x = clamp(point.x, 0, page_rect.width)
        point.y = clamp(point.y, 0, page_rect.height)
        anchor_x, anchor_y = self.drag_anchor
        original = self.drag_original

        if self.drag_mode == "new":
            rect = fitz.Rect(anchor_x, anchor_y, point.x, point.y)
            rect.normalize()
        elif self.drag_mode == "move":
            dx, dy = point.x - anchor_x, point.y - anchor_y
            dx = clamp(dx, -original.x0, page_rect.width - original.x1)
            dy = clamp(dy, -original.y0, page_rect.height - original.y1)
            rect = fitz.Rect(original.x0 + dx, original.y0 + dy, original.x1 + dx, original.y1 + dy)
        else:
            rect = fitz.Rect(original)
            if self.drag_mode in {"resize_nw", "resize_sw"}:
                rect.x0 = min(point.x, rect.x1 - 0.1)
            else:
                rect.x1 = max(point.x, rect.x0 + 0.1)
            if self.drag_mode in {"resize_nw", "resize_ne"}:
                rect.y0 = min(point.y, rect.y1 - 0.1)
            else:
                rect.y1 = max(point.y, rect.y0 + 0.1)
        if self.snap_var.get():
            rect = fitz.Rect(*(round(value) for value in (rect.x0, rect.y0, rect.x1, rect.y1)))
        self.current_rect = rect
        self.rect_to_form()
        self.draw_overlays()

    def canvas_release(self, _event: tk.Event) -> None:
        if self.pan_dragging:
            self.pan_dragging = False
            self.drag_mode = None
            self.drag_original = None
            self.canvas.configure(
                cursor="hand2" if self.space_pan_active else "crosshair"
            )
            return
        self.drag_mode = None
        self.drag_original = None
        self.rect_to_form()

    def unit_factor(self) -> float:
        return MM_TO_PT if self.unit_var.get() == "mm" else 1.0

    def rect_to_form(self) -> None:
        factor = self.unit_factor()
        values = {
            "x": self.current_rect.x0 / factor,
            "y": self.current_rect.y0 / factor,
            "width": self.current_rect.width / factor,
            "height": self.current_rect.height / factor,
        }
        for key, value in values.items():
            self.form_vars[key].set(f"{value:.2f}")

    def form_rect_changed(self) -> None:
        try:
            factor = self.unit_factor()
            x = float(str(self.form_vars["x"].get())) * factor
            y = float(str(self.form_vars["y"].get())) * factor
            width = float(str(self.form_vars["width"].get())) * factor
            height = float(str(self.form_vars["height"].get())) * factor
            if width <= 0 or height <= 0:
                return
            self.current_rect = fitz.Rect(x, y, x + width, y + height)
            self.draw_overlays()
        except (ValueError, tk.TclError):
            return

    def on_unit_changed(self, _event: tk.Event | None = None) -> None:
        self.rect_to_form()
        if hasattr(self, "tree"):
            self.refresh_tree(select_id=self.selected_rule_id)

    def keyboard_move(self, dx: int, dy: int, event: tk.Event) -> str:
        if self.document is None:
            return "break"
        multiplier = 10 if (event.state & 0x0001) else 1
        step = float(self.step_var.get()) * self.unit_factor() * multiplier
        move_x = dx * step
        move_y = dy * step
        self.current_rect = fitz.Rect(
            self.current_rect.x0 + move_x,
            self.current_rect.y0 + move_y,
            self.current_rect.x1 + move_x,
            self.current_rect.y1 + move_y,
        )
        self.rect_to_form()
        self.draw_overlays()
        return "break"

    def update_link_type_state(self) -> None:
        page_mode = self.form_vars["link_type"].get() == "page"
        self.target_entry.configure(state="normal" if page_mode else "disabled")
        self.url_entry.configure(state="disabled" if page_mode else "normal")
        self.destination_view_box.configure(
            state="readonly" if page_mode else "disabled"
        )

    def update_border_display_state(self, _event: tk.Event | None = None) -> None:
        visible = self.border_display_var.get() == "ボックスを表示"
        self.border_style_box.configure(state="readonly" if visible else "disabled")

    def apply_border_width_preset(self, _event: tk.Event | None = None) -> None:
        preset = self.border_width_preset_var.get()
        if preset in BORDER_WIDTH_PRESETS:
            self.form_vars["border_width"].set(BORDER_WIDTH_PRESETS[preset])

    def sync_border_width_preset(self) -> None:
        try:
            current_width = float(str(self.form_vars["border_width"].get()))
        except ValueError:
            self.border_width_preset_var.set("")
            return
        for preset, value in BORDER_WIDTH_PRESETS.items():
            if math.isclose(current_width, float(value), abs_tol=0.001):
                self.border_width_preset_var.set(preset)
                return
        self.border_width_preset_var.set("")

    def choose_color(self) -> None:
        _, selected = colorchooser.askcolor(
            color=str(self.form_vars["border_color"].get()), title="枠色を選択"
        )
        if selected:
            self.form_vars["border_color"].set(selected.upper())

    def rule_from_form(self) -> LinkRule:
        self.form_rect_changed()
        try:
            border_width = float(str(self.form_vars["border_width"].get()))
        except ValueError as exc:
            raise ValidationError("枠幅は数値で入力してください。") from exc
        name = str(self.form_vars["name"].get()).strip()
        if not name:
            raise ValidationError("設定名を入力してください。")
        rule = LinkRule(
            name=name,
            page_spec=str(self.form_vars["page_spec"].get()).strip(),
            x=self.current_rect.x0,
            y=self.current_rect.y0,
            width=self.current_rect.width,
            height=self.current_rect.height,
            link_type=str(self.form_vars["link_type"].get()),
            target_page=str(self.form_vars["target_page"].get()).strip(),
            url=str(self.form_vars["url"].get()).strip(),
            destination_view=DESTINATION_VIEW_LABEL_TO_MODE.get(
                self.destination_view_var.get(), ""
            ),
            border_visible=self.border_display_var.get() == "ボックスを表示",
            border_color=str(self.form_vars["border_color"].get()).strip().upper(),
            border_width=border_width,
            border_style=BORDER_STYLE_LABEL_TO_MODE.get(
                self.border_style_var.get(), ""
            ),
            highlight_mode=HIGHLIGHT_LABEL_TO_MODE.get(
                str(self.form_vars["highlight_mode"].get()), ""
            ),
            enabled=bool(self.form_vars["enabled"].get()),
        )
        hex_to_pdf_color(rule.border_color)
        if rule.highlight_mode not in HIGHLIGHT_MODE_PDF:
            raise ValidationError("ハイライトを選択してください。")
        if rule.border_style not in BORDER_STYLE_PDF:
            raise ValidationError("枠のスタイルを選択してください。")
        if rule.destination_view not in DESTINATION_VIEW_LABELS:
            raise ValidationError("リンク先の表示方法を選択してください。")
        if rule.width <= 0 or rule.height <= 0:
            raise ValidationError("幅・高さは0より大きくしてください。")
        if self.document is not None:
            parse_page_spec(rule.page_spec, self.document.page_count)
            if rule.link_type == "page":
                resolve_target_page(self.document, rule.target_page)
        if rule.link_type == "url" and not rule.url:
            raise ValidationError("URLを入力してください。")
        return rule

    def push_undo(self) -> None:
        self.undo_stack.append(copy.deepcopy(self.rules))
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self) -> None:
        if not self.undo_stack:
            self.bell()
            return
        self.redo_stack.append(copy.deepcopy(self.rules))
        self.rules = self.undo_stack.pop()
        self.selected_rule_id = None
        self.refresh_tree()
        self.draw_overlays()
        self.status_var.set("直前の一覧変更を元に戻しました。")

    def redo(self) -> None:
        if not self.redo_stack:
            self.bell()
            return
        self.undo_stack.append(copy.deepcopy(self.rules))
        self.rules = self.redo_stack.pop()
        self.selected_rule_id = None
        self.refresh_tree()
        self.draw_overlays()
        self.status_var.set("一覧変更をやり直しました。")

    def add_rule(self) -> None:
        if self.document is None:
            messagebox.showwarning(APP_NAME, "先にPDFを開いてください。")
            return
        if not str(self.form_vars["name"].get()).strip():
            page = self.document[self.current_page]
            self.form_vars["name"].set(
                automatic_rule_name(
                    page,
                    self.current_rect,
                    self.current_page + 1,
                    self.rules,
                )
            )
            del page
        try:
            rule = self.rule_from_form()
        except ValidationError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.push_undo()
        self.rules.append(rule)
        self.selected_rule_id = rule.rule_id
        self.refresh_tree(select_id=rule.rule_id)
        self.draw_overlays()
        self.status_var.set(f"「{rule.name}」を登録しました。")

    def update_selected_rule(self) -> None:
        if not self.selected_rule_id:
            messagebox.showinfo(APP_NAME, "一覧から更新する設定を選択してください。")
            return
        try:
            replacement = self.rule_from_form()
        except ValidationError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        replacement.rule_id = self.selected_rule_id
        for index, rule in enumerate(self.rules):
            if rule.rule_id == self.selected_rule_id:
                replacement.source_page = rule.source_page
                replacement.source_xref = rule.source_xref
                replacement.source_named_destination = rule.source_named_destination
                replacement.source_named_destination_pdf = (
                    rule.source_named_destination_pdf
                )
                self.push_undo()
                self.rules[index] = replacement
                break
        self.refresh_tree(select_id=replacement.rule_id)
        self.draw_overlays()
        self.status_var.set(f"「{replacement.name}」を更新しました。")

    def delete_rule(self) -> None:
        if not self.selected_rule_id:
            return
        rule = next(
            (item for item in self.rules if item.rule_id == self.selected_rule_id), None
        )
        if rule is None:
            return
        question = (
            f"「{rule.name}」を別名出力PDFから削除しますか？\n"
            "元PDFのリンクは変更されません。"
            if rule.source_xref is not None
            else f"「{rule.name}」を削除しますか？"
        )
        if not messagebox.askyesno(APP_NAME, question):
            return
        self.push_undo()
        if rule.source_xref is not None:
            rule.deleted = True
            rule.enabled = False
        else:
            self.rules = [
                item for item in self.rules if item.rule_id != self.selected_rule_id
            ]
        self.selected_rule_id = None
        self.refresh_tree()
        self.draw_overlays()

    def duplicate_rule(self) -> None:
        rule = next(
            (item for item in self.rules if item.rule_id == self.selected_rule_id), None
        )
        if rule is None:
            return
        self.push_undo()
        duplicate = copy.deepcopy(rule)
        duplicate.rule_id = uuid.uuid4().hex
        duplicate.name = f"{rule.name} のコピー"
        duplicate.source_page = None
        duplicate.source_xref = None
        duplicate.deleted = False
        self.rules.append(duplicate)
        self.selected_rule_id = duplicate.rule_id
        self.refresh_tree(select_id=duplicate.rule_id)
        self.load_rule_into_form(duplicate)
        self.draw_overlays()

    def toggle_rule(self) -> None:
        rule = next(
            (item for item in self.rules if item.rule_id == self.selected_rule_id), None
        )
        if rule is None:
            return
        self.push_undo()
        rule.enabled = not rule.enabled
        self.refresh_tree(select_id=rule.rule_id)
        self.draw_overlays()

    def refresh_tree(self, select_id: str | None = None) -> None:
        factor = self.unit_factor()
        unit = self.unit_var.get()
        decimals = 2 if unit == "mm" else 1
        self.tree.heading("position", text=f"X, Y, 幅×高さ ({unit})")
        for item in self.tree.get_children():
            self.tree.delete(item)
        for rule in self.rules:
            if rule.deleted:
                continue
            if rule.link_type == "page":
                target_text = str(rule.target_page).strip()
                if target_text.isdigit():
                    target = f"{target_text}ページ"
                elif self.document is not None:
                    try:
                        resolved = resolve_target_page(self.document, target_text) + 1
                        target = f"{target_text} → {resolved}ページ"
                    except ValidationError:
                        target = target_text
                else:
                    target = target_text
            else:
                target = rule.url
            self.tree.insert(
                "",
                "end",
                iid=rule.rule_id,
                text=rule.name,
                values=(
                    "✓" if rule.enabled else "—",
                    rule.page_spec,
                    (
                        f"{rule.x / factor:.{decimals}f}, "
                        f"{rule.y / factor:.{decimals}f}, "
                        f"{rule.width / factor:.{decimals}f}×"
                        f"{rule.height / factor:.{decimals}f}"
                    ),
                    "ページ" if rule.link_type == "page" else "URL",
                    target,
                ),
            )
        if select_id and self.tree.exists(select_id):
            self.tree.selection_set(select_id)
            self.tree.focus(select_id)
            self.tree.see(select_id)

    def on_tree_select(self, _event: tk.Event) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self.selected_rule_id = selection[0]
        rule = next(
            (item for item in self.rules if item.rule_id == self.selected_rule_id), None
        )
        if rule:
            if self.document is not None:
                try:
                    pages = parse_page_spec(rule.page_spec, self.document.page_count)
                except ValidationError:
                    pages = []
                if pages and self.current_page not in pages:
                    self.current_page = pages[0]
                    self.render_current_page()
            self.load_rule_into_form(rule)
            self.draw_overlays()

    def load_rule_into_form(self, rule: LinkRule) -> None:
        self.current_rect = rule.rect()
        for key in (
            "name",
            "page_spec",
            "link_type",
            "target_page",
            "url",
            "border_color",
            "border_width",
        ):
            self.form_vars[key].set(getattr(rule, key))
        self.form_vars["highlight_mode"].set(
            HIGHLIGHT_MODE_LABELS.get(
                rule.highlight_mode, HIGHLIGHT_MODE_LABELS["none"]
            )
        )
        self.form_vars["enabled"].set(rule.enabled)
        self.destination_view_var.set(
            DESTINATION_VIEW_LABELS.get(
                rule.destination_view, DESTINATION_VIEW_LABELS["fit_page"]
            )
        )
        self.border_display_var.set(
            "ボックスを表示" if rule.border_visible else "ボックスを非表示"
        )
        self.border_style_var.set(
            BORDER_STYLE_LABELS.get(rule.border_style, BORDER_STYLE_LABELS["solid"])
        )
        self.sync_border_width_preset()
        self.rect_to_form()
        self.update_link_type_state()
        self.update_border_display_state()

    def reset_form(self) -> None:
        self.selected_rule_id = None
        if hasattr(self, "tree"):
            self.tree.selection_remove(self.tree.selection())
        self.current_rect = fitz.Rect(0, 0, MM_TO_PT, MM_TO_PT)
        defaults = LinkRule(name="")
        self.load_rule_into_form(defaults)
        self.selected_rule_id = None
        self.draw_overlays()

    def validate_all(
        self,
        show_success: bool = True,
        show_warnings: bool = True,
    ) -> ValidationReport | None:
        if self.document is None:
            messagebox.showwarning(APP_NAME, "先にPDFを開いてください。")
            return None
        report = validate_document_and_rules(self.document, self.rules)
        if report.errors:
            log_path = self.save_diagnostics(
                "設定の検査_エラー",
                (f"ERROR | {item}" for item in report.errors),
            )
            messagebox.showerror(
                "設定エラー",
                "設定を適用できない問題があります。\n\n"
                + self.diagnostic_message(len(report.errors), log_path),
            )
            return report
        if report.warnings and show_warnings:
            log_path = self.save_diagnostics(
                "設定の検査_警告",
                (f"WARNING | {item}" for item in report.warnings),
            )
            messagebox.showwarning(
                "検査結果",
                f"{report.insertion_count}件を適用できますが、警告があります。\n\n"
                + self.diagnostic_message(len(report.warnings), log_path),
            )
        elif show_success:
            messagebox.showinfo(
                "検査完了",
                f"問題はありません。\n{report.insertion_count}件のリンクを適用します。",
            )
        return report

    def export_pdf(self) -> None:
        if self.document is None or self.pdf_path is None:
            messagebox.showwarning(APP_NAME, "先にPDFを開いてください。")
            return
        report = self.validate_all(show_success=False, show_warnings=False)
        if report is None or not report.ok:
            return
        named_output_mode = NAMED_DESTINATION_OUTPUT_LABEL_TO_MODE.get(
            self.named_destination_output_var.get(), "preserve"
        )
        if report.warnings:
            log_path = self.save_diagnostics(
                "PDF適用前の警告",
                (f"WARNING | {item}" for item in report.warnings),
            )
            proceed = messagebox.askyesno(
                "重なりの警告",
                self.diagnostic_message(len(report.warnings), log_path)
                + "\n\n警告内容を保持したまま適用します。続行しますか？",
            )
            if not proceed:
                return
        initial = self.pdf_path.with_name(f"{self.pdf_path.stem}_linked.pdf")
        filename = filedialog.asksaveasfilename(
            title="リンク追加後のPDFを保存",
            defaultextension=".pdf",
            initialdir=str(initial.parent),
            initialfile=initial.name,
            filetypes=[("PDFファイル", "*.pdf")],
        )
        if not filename:
            return
        output = Path(filename)
        if output.resolve() == self.pdf_path.resolve():
            messagebox.showerror(
                APP_NAME, "元PDFは上書きしません。「_linked」など別名を指定してください。"
            )
            return
        if output.exists() and not messagebox.askyesno(
            APP_NAME, f"{output.name} は既に存在します。置き換えますか？"
        ):
            return
        self.status_var.set("別名PDFへリンクを適用しています...")
        self.update_idletasks()
        try:
            inserted, warnings = apply_rules_to_pdf(
                self.pdf_path,
                output,
                self.rules,
                named_output_mode,
            )
        except Exception as exc:
            error_text = str(exc)
            sharing_error = (
                isinstance(exc, PermissionError)
                or getattr(exc, "winerror", None) == 32
                or "WinError 32" in error_text
                or "別のプロセス" in error_text
                or "Permission denied" in error_text
            )
            if sharing_error:
                suffix = 1
                fallback = output.with_name(f"{output.stem}_{suffix}{output.suffix}")
                while fallback.exists():
                    suffix += 1
                    fallback = output.with_name(
                        f"{output.stem}_{suffix}{output.suffix}"
                    )
                try:
                    inserted, warnings = apply_rules_to_pdf(
                        self.pdf_path,
                        fallback,
                        self.rules,
                        named_output_mode,
                    )
                    warnings.insert(
                        0,
                        f"指定先が使用中だったため、{fallback.name}へ保存しました。",
                    )
                    output = fallback
                except Exception as fallback_exc:
                    self.status_var.set(
                        "PDFへの適用に失敗しました。元PDFは変更されていません。"
                    )
                    log_path = self.save_diagnostics(
                        "PDFへの適用_別名再試行失敗",
                        [
                            f"指定先={output} | {type(exc).__name__}: {exc}",
                            f"再試行先={fallback} | "
                            f"{type(fallback_exc).__name__}: {fallback_exc}",
                        ],
                        output_path=fallback,
                        exception=fallback_exc,
                    )
                    messagebox.showerror(
                        APP_NAME,
                        "指定先が使用中だったため別名でも試しましたが、"
                        "保存できませんでした。\n\n"
                        + self.diagnostic_message(2, log_path),
                    )
                    return
            else:
                self.status_var.set(
                    "PDFへの適用に失敗しました。元PDFは変更されていません。"
                )
                log_path = self.save_diagnostics(
                    "PDFへの適用_失敗",
                    [
                        f"出力先={output} | {type(exc).__name__}: {exc}",
                    ],
                    output_path=output,
                    exception=exc,
                )
                messagebox.showerror(
                    APP_NAME,
                    "PDFへ適用できませんでした。元PDFは変更されていません。\n\n"
                    + self.diagnostic_message(1, log_path),
                )
                return
        self.status_var.set(f"{output.name} へ {inserted}件のリンクを適用しました。")
        warnings = list(dict.fromkeys(warnings))
        extra_parts = [
            "名前付き移動先: " + NAMED_DESTINATION_OUTPUT_LABELS[named_output_mode]
        ]
        if warnings:
            log_path = self.save_diagnostics(
                "PDFへの適用_警告",
                (f"WARNING | {item}" for item in warnings),
                output_path=output,
            )
            extra_parts.append(self.diagnostic_message(len(warnings), log_path))
        extra = "\n\n" + "\n".join(extra_parts)
        messagebox.showinfo(
            "保存完了",
            f"{inserted}件のリンクを別名PDFへ適用しました。\n\n{output}{extra}",
        )

    def save_settings(self) -> None:
        if not self.rules:
            messagebox.showinfo(APP_NAME, "保存する登録予定がありません。")
            return
        filename = filedialog.asksaveasfilename(
            title="登録予定を保存",
            defaultextension=".json",
            filetypes=[("JSON設定", "*.json")],
            initialfile="pdf_link_settings.json",
        )
        if not filename:
            return
        payload = {
            "format": "pdf-link-manager-settings",
            "version": 5,
            "named_destination_output": NAMED_DESTINATION_OUTPUT_LABEL_TO_MODE.get(
                self.named_destination_output_var.get(), "preserve"
            ),
            "coordinate_unit": "pt",
            "source_pdf_name": self.pdf_path.name if self.pdf_path else None,
            "rules": [asdict(rule) for rule in self.rules],
        }
        temp_path = Path(filename).with_suffix(".json.tmp")
        try:
            temp_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            os.replace(temp_path, filename)
        except Exception as exc:
            if temp_path.exists():
                temp_path.unlink()
            messagebox.showerror(APP_NAME, f"登録予定を保存できませんでした。\n\n{exc}")
            return
        self.status_var.set(f"登録予定を保存しました: {Path(filename).name}")

    def load_settings(self) -> None:
        if self.document is None or self.pdf_path is None:
            messagebox.showwarning(APP_NAME, "先に登録予定を適用するPDFを開いてください。")
            return
        filename = filedialog.askopenfilename(
            title="登録予定を読み込む", filetypes=[("JSON設定", "*.json")]
        )
        if not filename:
            return
        try:
            payload = json.loads(Path(filename).read_text(encoding="utf-8"))
            if payload.get("format") != "pdf-link-manager-settings":
                raise ValidationError("このアプリの設定ファイルではありません。")
            loaded = [LinkRule.from_dict(item) for item in payload.get("rules", [])]
            if not loaded:
                raise ValidationError("登録予定が含まれていません。")
            if payload.get("source_pdf_name") != self.pdf_path.name:
                reusable: list[LinkRule] = []
                for rule in loaded:
                    if rule.deleted:
                        continue
                    rule.source_page = None
                    rule.source_xref = None
                    reusable.append(rule)
                loaded = reusable
                if not loaded:
                    raise ValidationError("再利用できる登録予定がありません。")
            output_mode = payload.get("named_destination_output")
            if output_mode in NAMED_DESTINATION_OUTPUT_LABELS:
                self.named_destination_output_var.set(
                    NAMED_DESTINATION_OUTPUT_LABELS[output_mode]
                )
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"登録予定を読み込めませんでした。\n\n{exc}")
            return
        self.push_undo()
        self.rules = loaded
        self.selected_rule_id = None
        self.refresh_tree()
        self.draw_overlays()
        self.status_var.set(f"{len(loaded)}件の登録予定を読み込みました。")

    def show_help(self) -> None:
        messagebox.showinfo(
            "使い方",
            "1. PDFを開きます。既存リンクがある場合は、登録予定一覧へ\n"
            "   読み込んで編集するか、青枠のまま保持するかを選択します。\n"
            "   ページ番号欄へ入力して「移動」を押すと指定ページへ移動できます。\n"
            "2. プレビューの空いている部分をドラッグしてリンク範囲を作ります。\n"
            "   登録予定の枠をクリックすると、一覧の該当項目を選択します。\n"
            "3. X・Y・幅・高さを数値で微調整します。\n"
            "4. 対象ページとリンク先を指定し、「新規として追加」を押します。\n"
            "   設定名が空欄なら、範囲内の文字・画像などから自動命名します。\n"
            "5. 読み込んだ既存リンクは一覧から選んで更新できます。\n"
            "6. 必要な設定をすべて登録したら、検査して別名保存します。\n\n"
            "プレビュー操作: Space+ドラッグで表示移動、ホイールでページ移動、\n"
            "Ctrl+ホイール（MacはCommand+ホイール）で、"
            "カーソル位置を中心に拡大縮小します。\n\n"
            "対象ページ: 全ページ、1-5,8,10、奇数、偶数\n"
            "赤枠: 現在編集中 / 黄枠: 編集可能リンク / 青破線: 対象外の既存リンク\n\n"
            "リンク範囲はページの外側にはみ出していても、その数値のまま適用します。\n\n"
            "リンクの種類、実線・破線・下線、ハイライトと、\n"
            "リンク先の全体表示・100%表示・幅表示を指定できます。\n\n"
            "保存前に、名前付き移動先を元の名前で保持するか、\n"
            "ページ番号へ変換するかを上部の一覧から選択できます。\n"
            "元の名前を保持する場合、表示方法は移動先の定義に従います。",
        )

    def show_about(self) -> None:
        messagebox.showinfo(
            "バージョン情報",
            f"{APP_NAME}\nバージョン {APP_VERSION}\n統合版\n\n"
            "Tkinter + PyMuPDF",
        )


def main() -> None:
    try:
        require_pymupdf()
    except ValidationError as exc:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(APP_NAME, str(exc), parent=root)
            root.destroy()
        except Exception:
            print(f"{APP_NAME}: {exc}", file=sys.stderr)
        return
    app = PDFLinkManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
