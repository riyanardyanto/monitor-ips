from __future__ import annotations

import asyncio
from datetime import datetime
from io import BytesIO
import json
from pathlib import Path
import sys
import tkinter as tk
import traceback
from tkinter import filedialog

import flet as ft
import qrcode
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from check_excel_completeness import run_checks


if getattr(sys, "frozen", False):
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    APP_DIR = Path(sys.executable).resolve().parent
else:
    RESOURCE_DIR = Path(__file__).resolve().parent
    APP_DIR = RESOURCE_DIR


LOG_FILE = APP_DIR / "ips-checker-error.log"
ICON_FILE = RESOURCE_DIR / "assets" / "app_icon.ico"
CONFIG_FILE = APP_DIR / "ips-checker-config.json"
DEFAULT_WINDOW_WIDTH = 1180
DEFAULT_WINDOW_HEIGHT = 820


def write_error_log(context: str, exc: BaseException) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = "\n".join(
        [
            f"[{timestamp}] {context}",
            f"Type: {type(exc).__name__}",
            f"Message: {exc}",
            "Traceback:",
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip(),
            "-" * 80,
            "",
        ]
    )
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as log_file:
        log_file.write(log_entry)
    return LOG_FILE


def sanitize_window_size(value: object, fallback: int) -> int:
    try:
        parsed_value = int(float(value))
    except (TypeError, ValueError):
        return fallback
    return parsed_value if parsed_value >= 600 else fallback


def load_app_config(default_target_path: Path, default_field_file: Path) -> dict[str, object]:
    config = {
        "target_path": str(default_target_path),
        "field_file": str(default_field_file),
        "sheet_mode": "auto",
        "sheet_name": None,
        "csv_output": str(APP_DIR / "report.csv"),
        "window_width": DEFAULT_WINDOW_WIDTH,
        "window_height": DEFAULT_WINDOW_HEIGHT,
    }

    if not CONFIG_FILE.exists():
        return config

    try:
        loaded_config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        write_error_log("Gagal membaca file config", exc)
        return config

    if not isinstance(loaded_config, dict):
        return config

    target_path = str(loaded_config.get("target_path") or config["target_path"])
    field_file = str(loaded_config.get("field_file") or config["field_file"])
    csv_output = str(loaded_config.get("csv_output") or config["csv_output"])
    sheet_mode = str(loaded_config.get("sheet_mode") or "auto").strip().lower()
    if sheet_mode not in {"auto", "custom"}:
        sheet_mode = "auto"

    sheet_name_raw = loaded_config.get("sheet_name")
    sheet_name = str(sheet_name_raw).strip() if sheet_name_raw else None
    if sheet_mode == "auto":
        sheet_name = None

    return {
        "target_path": target_path,
        "field_file": field_file,
        "sheet_mode": sheet_mode,
        "sheet_name": sheet_name,
        "csv_output": csv_output,
        "window_width": sanitize_window_size(loaded_config.get("window_width"), DEFAULT_WINDOW_WIDTH),
        "window_height": sanitize_window_size(loaded_config.get("window_height"), DEFAULT_WINDOW_HEIGHT),
    }


def resolve_initial_directory(path_value: str | None, fallback: Path) -> str:
    if not path_value:
        return str(fallback)

    candidate = Path(path_value)
    if not candidate.is_absolute():
        candidate = (APP_DIR / candidate).resolve()

    if candidate.is_file():
        candidate = candidate.parent

    if candidate.exists():
        return str(candidate)

    if candidate.parent.exists():
        return str(candidate.parent)

    return str(fallback)


def with_hidden_tk_root(action):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return action(root)
    finally:
        root.destroy()


def main(page: ft.Page) -> None:
    section_labels = {
        "1.1": "1.1. Initiate IPS",
        "1.2": "1.2. Refocus the Problem",
        "1.3": "1.3. Verify base condition",
        "1.4": "1.4. Restore Base Condition",
        "1.5": "1.5. Ask WHY WHY",
    }
    side_button_width = 190

    default_field_file = APP_DIR / "field_cell.txt"
    if not default_field_file.exists():
        default_field_file = RESOURCE_DIR / "field_cell.txt"

    default_target_path = APP_DIR / "new create ips"
    if not default_target_path.exists():
        default_target_path = APP_DIR

    persisted_config = load_app_config(default_target_path, default_field_file)

    page.title = "IPS Completeness Checker"
    page.window_width = int(persisted_config["window_width"])
    page.window_height = int(persisted_config["window_height"])
    page.padding = 24
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#f4f1e8"
    page.scroll = ft.ScrollMode.AUTO
    if ICON_FILE.exists():
        page.window.icon = str(ICON_FILE)
    page.theme = ft.Theme(
        color_scheme=ft.ColorScheme(
            primary="#0f5c4d",
            secondary="#c96f3b",
            surface="#fffaf0",
        ),
        visual_density=ft.VisualDensity.COMFORTABLE,
    )

    status_text = ft.Text(color="#6b4f3a")
    summary_text = ft.Text(size=16, weight=ft.FontWeight.W_600, color="#16302b")
    loading_indicator = ft.ProgressRing(width=22, height=22, stroke_width=3, color="#0f5c4d", visible=False)
    loading_text = ft.Text("Sedang memeriksa file Excel...", color="#0f5c4d", weight=ft.FontWeight.W_600, visible=False)
    loading_row = ft.Row(
        visible=False,
        spacing=10,
        vertical_alignment=ft.CrossAxisAlignment.CENTER,
        controls=[loading_indicator, loading_text],
    )
    app_state = {
        "target_path": str(persisted_config["target_path"]),
        "field_file": str(persisted_config["field_file"]),
        "sheet_mode": str(persisted_config["sheet_mode"]),
        "sheet_name": persisted_config["sheet_name"],
        "csv_output": str(persisted_config["csv_output"]),
        "window_width": int(persisted_config["window_width"]),
        "window_height": int(persisted_config["window_height"]),
        "is_loading": False,
        "is_picker_open": False,
    }
    table_state = {
        "results": [],
        "sort_column_index": 1,
        "sort_ascending": True,
    }
    results_table = ft.Column(
        spacing=0,
        expand=True,
        scroll=ft.ScrollMode.AUTO,
    )
    table_container = ft.Container(
        expand=True,
        padding=0,
        margin=0,
        bgcolor="#fffaf0",
        border=ft.Border.all(1, "#d9c9ab"),
        border_radius=18,
        clip_behavior=ft.ClipBehavior.HARD_EDGE,
        content=results_table,
    )
    detail_dialog = ft.AlertDialog(
        modal=True,
        bgcolor="#fffaf0",
        shape=ft.RoundedRectangleBorder(radius=18),
        title=ft.Text("Detail Completeness IPS", weight=ft.FontWeight.W_700, color="#16302b"),
        actions=[ft.TextButton("Tutup")],
        actions_alignment=ft.MainAxisAlignment.END,
        scrollable=True,
    )
    follow_up_dialog = ft.AlertDialog(
        modal=True,
        bgcolor="#fffaf0",
        shape=ft.RoundedRectangleBorder(radius=18),
        title=ft.Text("Follow Up Section 1.5", weight=ft.FontWeight.W_700, color="#16302b"),
        actions=[ft.TextButton("Tutup")],
        actions_alignment=ft.MainAxisAlignment.END,
        scrollable=True,
    )
    share_dialog = ft.AlertDialog(
        modal=True,
        bgcolor="#fffaf0",
        shape=ft.RoundedRectangleBorder(radius=18),
        title=ft.Text("Share Detail Completeness", weight=ft.FontWeight.W_700, color="#16302b"),
        actions=[ft.TextButton("Tutup")],
        actions_alignment=ft.MainAxisAlignment.END,
        scrollable=True,
    )
    info_dialog = ft.AlertDialog(
        modal=True,
        bgcolor="#fffaf0",
        shape=ft.RoundedRectangleBorder(radius=18),
        title=ft.Text("Info GUI", weight=ft.FontWeight.W_700, color="#16302b"),
        actions=[ft.TextButton("Tutup")],
        actions_alignment=ft.MainAxisAlignment.END,
        scrollable=True,
    )
    export_dialog = ft.AlertDialog(
        modal=True,
        bgcolor="#fffaf0",
        shape=ft.RoundedRectangleBorder(radius=18),
        title=ft.Text("Pilih Tipe Export", weight=ft.FontWeight.W_700, color="#16302b"),
        actions=[ft.TextButton("Tutup")],
        actions_alignment=ft.MainAxisAlignment.END,
        scrollable=True,
    )
    clipboard_service = ft.Clipboard()
    folder_button = ft.Button(
        content=ft.Text("Pilih Folder", weight=ft.FontWeight.W_600),
        icon=ft.Icons.FOLDER_OPEN_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
        ),
    )
    file_button = ft.Button(
        content=ft.Text("Pilih File", weight=ft.FontWeight.W_600),
        icon=ft.Icons.UPLOAD_FILE_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
        ),
    )
    csv_button = ft.Button(
        content=ft.Text("Simpan CSV", weight=ft.FontWeight.W_600),
        icon=ft.Icons.SAVE_AS_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
        ),
    )
    share_button = ft.Button(
        content=ft.Text("Export", weight=ft.FontWeight.W_600),
        icon=ft.Icons.DOWNLOAD_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
        ),
    )
    info_button = ft.Button(
        content=ft.Text("Info", weight=ft.FontWeight.W_600),
        icon=ft.Icons.INFO_OUTLINE_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
        ),
    )
    app_logo = ft.Container(
        width=side_button_width,
        padding=ft.Padding.symmetric(horizontal=14, vertical=14),
        bgcolor="#efe4ce",
        border=ft.Border.all(1, "#d9c9ab"),
        border_radius=18,
        content=ft.Column(
            spacing=10,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Container(
                    width=54,
                    height=54,
                    bgcolor="#0f5c4d",
                    border_radius=16,
                    content=ft.Stack(
                        controls=[
                            ft.Container(
                                left=0,
                                right=0,
                                top=0,
                                bottom=0,
                                content=ft.Text(
                                    "IPS",
                                    color="#fffaf0",
                                    weight=ft.FontWeight.W_700,
                                    text_align=ft.TextAlign.CENTER,
                                    size=15,
                                ),
                            ),
                            ft.Container(
                                right=4,
                                bottom=4,
                                width=14,
                                height=14,
                                bgcolor="#c96f3b",
                                border_radius=999,
                            ),
                        ]
                    ),
                ),
                ft.Column(
                    spacing=2,
                    tight=True,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Text("IPS Checker", weight=ft.FontWeight.W_700, color="#16302b", size=17, text_align=ft.TextAlign.CENTER),
                        ft.Text("Monitor completeness form IPS", color="#6b4f3a", size=11, text_align=ft.TextAlign.CENTER),
                    ],
                ),
            ],
        ),
    )
    def refresh_busy_state() -> None:
        is_loading = bool(app_state["is_loading"])
        is_picker_open = bool(app_state["is_picker_open"])
        is_busy = is_loading or is_picker_open
        loading_indicator.visible = is_loading
        loading_text.visible = is_loading
        loading_row.visible = is_loading
        folder_button.disabled = is_busy
        file_button.disabled = is_busy
        csv_button.disabled = is_busy
        share_button.disabled = is_busy
        info_button.disabled = is_busy

    def persist_config() -> None:
        config_payload = {
            "target_path": app_state["target_path"],
            "field_file": app_state["field_file"],
            "sheet_mode": app_state["sheet_mode"],
            "sheet_name": app_state["sheet_name"],
            "csv_output": app_state["csv_output"],
            "window_width": app_state["window_width"],
            "window_height": app_state["window_height"],
        }
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            CONFIG_FILE.write_text(json.dumps(config_payload, indent=2), encoding="utf-8")
        except Exception as exc:
            write_error_log("Gagal menyimpan file config", exc)

    def sync_window_size_from_page() -> None:
        width = page.window_width or page.width or app_state["window_width"]
        height = page.window_height or page.height or app_state["window_height"]
        app_state["window_width"] = sanitize_window_size(width, int(app_state["window_width"]))
        app_state["window_height"] = sanitize_window_size(height, int(app_state["window_height"]))

    def set_loading(is_loading: bool) -> None:
        app_state["is_loading"] = is_loading
        refresh_busy_state()

    def set_picker_open(is_picker_open: bool) -> None:
        app_state["is_picker_open"] = is_picker_open
        refresh_busy_state()

    def notify(message: str, error: bool = False) -> None:
        page.snack_bar = ft.SnackBar(
            content=ft.Text(message),
            bgcolor="#a63d40" if error else "#245a4a",
            open=True,
        )
        page.update()

    def pick_target_folder(_: ft.ControlEvent) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        set_picker_open(True)
        page.update()
        selected_path = with_hidden_tk_root(
            lambda _: filedialog.askdirectory(
                title="Pilih folder Excel",
                initialdir=resolve_initial_directory(str(app_state["target_path"]), APP_DIR),
            )
        )
        if selected_path:
            app_state["target_path"] = selected_path
            persist_config()
            set_picker_open(False)
            start_run_check(selected_path)
            return
        set_picker_open(False)
        page.update()

    def pick_target_file(_: ft.ControlEvent) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        set_picker_open(True)
        page.update()
        selected_path = with_hidden_tk_root(
            lambda _: filedialog.askopenfilename(
                title="Pilih file Excel",
                initialdir=resolve_initial_directory(str(app_state["target_path"]), APP_DIR),
                filetypes=[("Excel files", "*.xlsx")],
            )
        )
        if selected_path:
            app_state["target_path"] = selected_path
            persist_config()
            set_picker_open(False)
            start_run_check(selected_path)
            return
        set_picker_open(False)
        page.update()

    def pick_csv_output(_: ft.ControlEvent) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        set_picker_open(True)
        page.update()
        selected_path = with_hidden_tk_root(
            lambda _: filedialog.asksaveasfilename(
                title="Simpan hasil CSV",
                initialdir=resolve_initial_directory(str(app_state["csv_output"]), APP_DIR),
                initialfile=Path(str(app_state["csv_output"])).name,
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
            )
        )
        if selected_path:
            app_state["csv_output"] = selected_path
            persist_config()
        set_picker_open(False)
        page.update()

    def handle_page_resize(_: ft.ControlEvent) -> None:
        sync_window_size_from_page()
        persist_config()

    def handle_page_close(_: ft.ControlEvent) -> None:
        sync_window_size_from_page()
        persist_config()

    def build_completeness_cell(result: dict[str, object]) -> ft.Control:
        completeness = float(result["completeness_percentage"])
        if completeness >= 80:
            bar_color = "#245a4a"
        elif completeness >= 50:
            bar_color = "#c96f3b"
        else:
            bar_color = "#a63d40"

        return ft.Row(
            spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.ProgressBar(
                    value=completeness / 100,
                    width=120,
                    bar_height=8,
                    color=bar_color,
                    bgcolor="#e7dbc5",
                    border_radius=999,
                ),
                ft.Text(f"{completeness:.2f}%", weight=ft.FontWeight.W_600, color="#16302b"),
            ],
        )

    def close_detail_dialog(_: ft.ControlEvent) -> None:
        detail_dialog.open = False
        page.pop_dialog()
        page.update()

    def close_follow_up_dialog(_: ft.ControlEvent) -> None:
        follow_up_dialog.open = False
        page.pop_dialog()
        page.update()

    def close_share_dialog(_: ft.ControlEvent) -> None:
        share_dialog.open = False
        page.pop_dialog()
        page.update()

    def close_info_dialog(_: ft.ControlEvent) -> None:
        info_dialog.open = False
        page.pop_dialog()
        page.update()

    def close_export_dialog(_: ft.ControlEvent) -> None:
        export_dialog.open = False
        page.pop_dialog()
        page.update()

    async def copy_summary_to_clipboard(summary_text_value: str) -> None:
        try:
            await clipboard_service.set(summary_text_value)
            notify("Summary copied to clipboard.")
        except Exception as exc:
            notify(f"Gagal menyalin summary: {exc}", error=True)

    def build_detail_summary_text(result: dict[str, object]) -> str:
        status_label = "COMPLETE" if result["is_complete"] else "INCOMPLETE"
        section_completeness = result.get("section_completeness", {})
        summary_lines = [
            f"File: {result['file_name']}",
            f"Participant: {result.get('participant') or '-'}",
            f"Status: {status_label}",
            f"Completeness: {float(result['completeness_percentage']):.2f}%",
            "Per Section:",
        ]
        for section in sorted(section_completeness):
            percentage = float(section_completeness.get(section, 0.0))
            summary_lines.append(f"- {section_labels.get(str(section), str(section))}: {percentage:.1f}%")
        return "\n".join(summary_lines)

    def build_follow_up_summary_text(result: dict[str, object]) -> str:
        follow_up_rows = result.get("follow_up_rows", [])
        summary_lines = [
            f"File: {result['file_name']}",
            f"Participant: {result.get('participant') or '-'}",
            "Follow Up Section 1.5 - Ask WHY WHY",
        ]

        if not follow_up_rows:
            summary_lines.append("Belum ada data follow up.")
            return "\n".join(summary_lines)

        for index, row in enumerate(follow_up_rows, start=1):
            summary_lines.extend(
                [
                    f"{index}. Countermeasure: {row.get('countermeasure') or '-'}",
                    f"   Responsible: {row.get('responsible') or '-'}",
                    f"   Due Date: {row.get('due_date') or '-'}",
                ]
            )
        return "\n".join(summary_lines)

    def build_results_summary_text(incomplete_only: bool) -> str:
        if incomplete_only:
            results = [result for result in table_state["results"] if not result.get("is_complete")]
        else:
            results = list(table_state["results"])
        if not results:
            if incomplete_only:
                return "Belum ada file incomplete untuk dibagikan."
            return "Belum ada hasil pemeriksaan untuk dibagikan."

        title = "Daftar File INCOMPLETE IPS" if incomplete_only else "Daftar Hasil Pemeriksaan IPS"
        summary_lines = [
            title,
            f"{len(results)} file",
            "",
        ]
        for index, result in enumerate(results, start=1):
            status_label = "COMPLETE" if result["is_complete"] else "INCOMPLETE"
            completeness = float(result["completeness_percentage"])
            participant = result.get("participant") or "-"
            summary_lines.extend(
                [
                    f"{index}. {result['file_name']}",
                    f"   Participant: {participant}",
                    f"   Status: {status_label}",
                    f"   Completeness: {completeness:.2f}%",
                ]
            )
        return "\n".join(summary_lines)

    def export_results_summary_pdf(output_path: Path) -> None:
        results = list(table_state["results"])
        if not results:
            raise ValueError("Belum ada hasil pemeriksaan untuk dibuatkan PDF.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        document = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm,
        )
        styles = getSampleStyleSheet()
        story = [
            Paragraph("IPS Completeness Summary", styles["Title"]),
            Spacer(1, 0.3 * cm),
            Paragraph(summary_text.value or "Belum ada ringkasan hasil.", styles["Normal"]),
            Spacer(1, 0.4 * cm),
        ]

        table_rows = [["File Name", "Participant", "Status", "Completeness %"]]
        for result in results:
            status_label = "COMPLETE" if result["is_complete"] else "INCOMPLETE"
            table_rows.append(
                [
                    Paragraph(str(result["file_name"]), styles["BodyText"]),
                    Paragraph(str(result.get("participant") or "-"), styles["BodyText"]),
                    Paragraph(status_label, styles["BodyText"]),
                    Paragraph(f"{float(result['completeness_percentage']):.2f}%", styles["BodyText"]),
                ]
            )

        results_table_pdf = Table(table_rows, colWidths=[8.5 * cm, 3.5 * cm, 3.0 * cm, 3.0 * cm], repeatRows=1)
        results_table_pdf.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#efe4ce")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#16302b")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d9c9ab")),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#fffaf0")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(results_table_pdf)
        document.build(story)

    def export_results_summary_json(output_path: Path) -> None:
        results = list(table_state["results"])
        if not results:
            raise ValueError("Belum ada hasil pemeriksaan untuk dibuatkan JSON.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "summary": summary_text.value or "",
            "results_count": len(results),
            "results": results,
        }
        output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    def export_results_summary_jpg(output_path: Path) -> None:
        results = list(table_state["results"])
        if not results:
            raise ValueError("Belum ada hasil pemeriksaan untuk dibuatkan JPG.")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        base_font = ImageFont.load_default()
        title_font = ImageFont.load_default()
        image_width = 1600
        margin_x = 56
        top_margin = 44
        card_gap = 18
        row_gap = 6
        line_gap = 6

        sample_draw = ImageDraw.Draw(Image.new("RGB", (10, 10), "white"))

        def line_height_for(font: ImageFont.ImageFont) -> int:
            bbox = sample_draw.textbbox((0, 0), "Ag", font=font)
            return bbox[3] - bbox[1]

        def wrap_text_to_width(text: object, max_width: int, font: ImageFont.ImageFont) -> list[str]:
            normalized_text = str(text or "-").strip() or "-"
            wrapped_lines: list[str] = []
            for paragraph in normalized_text.splitlines() or [normalized_text]:
                words = paragraph.split()
                if not words:
                    wrapped_lines.append("")
                    continue
                current_line = words[0]
                for word in words[1:]:
                    candidate = f"{current_line} {word}"
                    if sample_draw.textlength(candidate, font=font) <= max_width:
                        current_line = candidate
                    else:
                        wrapped_lines.append(current_line)
                        current_line = word
                wrapped_lines.append(current_line)
            return wrapped_lines or ["-"]

        def draw_text_lines(
            draw: ImageDraw.ImageDraw,
            lines: list[str],
            x: int,
            y: int,
            font: ImageFont.ImageFont,
            fill: str,
            extra_gap: int = line_gap,
        ) -> int:
            current_y = y
            text_line_height = line_height_for(font)
            for line in lines:
                draw.text((x, current_y), line, fill=fill, font=font)
                current_y += text_line_height + extra_gap
            return current_y

        def draw_card(
            draw: ImageDraw.ImageDraw,
            left: int,
            top: int,
            width: int,
            height: int,
            title: str,
            value: str,
            accent_color: str,
        ) -> None:
            draw.rounded_rectangle((left, top, left + width, top + height), radius=24, fill="#fffaf0", outline="#d9c9ab", width=2)
            draw.rounded_rectangle((left + 18, top + 18, left + 58, top + 58), radius=14, fill=accent_color)
            draw.text((left + 78, top + 18), title, fill="#6b4f3a", font=base_font)
            draw.text((left + 78, top + 52), value, fill="#16302b", font=title_font)

        complete_count = sum(1 for result in results if result["is_complete"])
        incomplete_count = len(results) - complete_count
        average_completeness = sum(float(result["completeness_percentage"]) for result in results) / len(results)

        summary_value = summary_text.value or (
            f"{len(results)} file diperiksa | {incomplete_count} file incomplete | {complete_count} file complete"
        )
        summary_lines = wrap_text_to_width(summary_value, image_width - (margin_x * 2) - 40, base_font)

        columns = [
            ("No", 70),
            ("File Name", 620),
            ("Participant", 270),
            ("Status", 220),
            ("Completeness", 220),
        ]
        table_left = margin_x
        table_width = sum(width for _, width in columns)
        header_height = 170
        cards_top = top_margin + header_height - 18
        card_width = (table_width - (card_gap * 3)) // 4
        card_height = 108
        table_top = cards_top + card_height + 32
        body_line_height = line_height_for(base_font)

        prepared_rows: list[dict[str, object]] = []
        for index, result in enumerate(results, start=1):
            status_label = "COMPLETE" if result["is_complete"] else "INCOMPLETE"
            completeness_label = f"{float(result['completeness_percentage']):.2f}%"
            row_cells = [
                wrap_text_to_width(str(index), columns[0][1] - 24, base_font),
                wrap_text_to_width(result["file_name"], columns[1][1] - 24, base_font),
                wrap_text_to_width(result.get("participant") or "-", columns[2][1] - 24, base_font),
                wrap_text_to_width(status_label, columns[3][1] - 24, base_font),
                wrap_text_to_width(completeness_label, columns[4][1] - 24, base_font),
            ]
            max_lines = max(len(cell_lines) for cell_lines in row_cells)
            row_height = max(44, 18 + max_lines * (body_line_height + line_gap))
            prepared_rows.append(
                {
                    "cells": row_cells,
                    "status": status_label,
                    "row_height": row_height,
                }
            )

        table_header_height = 48
        footer_height = 48
        image_height = table_top + table_header_height + sum(int(row["row_height"]) + row_gap for row in prepared_rows) + footer_height + 30

        image = Image.new("RGB", (image_width, max(900, image_height)), "#f4f1e8")
        draw = ImageDraw.Draw(image)

        draw.rounded_rectangle((24, 24, image_width - 24, image_height - 24), radius=32, fill="#f7efe0")
        draw.rounded_rectangle((margin_x, top_margin, image_width - margin_x, top_margin + header_height), radius=32, fill="#efe4ce")
        draw.rounded_rectangle((margin_x + 24, top_margin + 24, margin_x + 92, top_margin + 92), radius=20, fill="#0f5c4d")
        draw.ellipse((margin_x + 70, top_margin + 62, margin_x + 94, top_margin + 86), fill="#c96f3b")
        draw.text((margin_x + 118, top_margin + 26), "IPS Completeness Summary", fill="#16302b", font=title_font)
        draw.text((margin_x + 118, top_margin + 58), f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", fill="#6b4f3a", font=base_font)
        draw_text_lines(draw, summary_lines, margin_x + 118, top_margin + 88, base_font, "#4b5d58", extra_gap=4)

        draw_card(draw, table_left, cards_top, card_width, card_height, "Total File", str(len(results)), "#0f5c4d")
        draw_card(draw, table_left + card_width + card_gap, cards_top, card_width, card_height, "Complete", str(complete_count), "#245a4a")
        draw_card(draw, table_left + (card_width + card_gap) * 2, cards_top, card_width, card_height, "Incomplete", str(incomplete_count), "#a63d40")
        draw_card(draw, table_left + (card_width + card_gap) * 3, cards_top, card_width, card_height, "Avg Completeness", f"{average_completeness:.1f}%", "#c96f3b")

        draw.rounded_rectangle((table_left, table_top, table_left + table_width, table_top + table_header_height), radius=20, fill="#16302b")
        current_x = table_left
        for column_name, column_width in columns:
            draw.text((current_x + 14, table_top + 15), column_name, fill="#fffaf0", font=base_font)
            current_x += column_width

        current_y = table_top + table_header_height + 10
        for row_index, row in enumerate(prepared_rows):
            row_height = int(row["row_height"])
            status_label = str(row["status"])
            status_bg = "#e3f1ea" if status_label == "COMPLETE" else "#f7dfdf"
            status_fg = "#245a4a" if status_label == "COMPLETE" else "#a63d40"
            row_bg = "#fffaf0" if row_index % 2 == 0 else "#fcf5e8"
            draw.rounded_rectangle((table_left, current_y, table_left + table_width, current_y + row_height), radius=18, fill=row_bg, outline="#e0d2b7")

            current_x = table_left
            for cell_index, ((_, column_width), cell_lines) in enumerate(zip(columns, row["cells"])):
                cell_x = current_x + 14
                cell_y = current_y + 12
                if cell_index == 3:
                    badge_width = min(column_width - 28, max(110, int(sample_draw.textlength(status_label, font=base_font)) + 38))
                    badge_height = 30
                    draw.rounded_rectangle(
                        (cell_x, cell_y, cell_x + badge_width, cell_y + badge_height),
                        radius=15,
                        fill=status_bg,
                    )
                    draw.text((cell_x + 14, cell_y + 8), status_label, fill=status_fg, font=base_font)
                else:
                    draw_text_lines(draw, list(cell_lines), cell_x, cell_y, base_font, "#16302b", extra_gap=4)
                current_x += column_width

            current_y += row_height + row_gap

        draw.text((table_left, current_y + 14), "Generated by IPS Checker", fill="#9a8874", font=base_font)
        image.save(output_path, format="JPEG", quality=95)

    def request_export_path(dialog_title: str, extension: str) -> Path | None:
        set_picker_open(True)
        page.update()
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        default_name = f"ips-completeness-summary-{timestamp}.{extension}"
        selected_path = with_hidden_tk_root(
            lambda _: filedialog.asksaveasfilename(
                title=dialog_title,
                initialdir=resolve_initial_directory(str(app_state["csv_output"]), APP_DIR),
                initialfile=default_name,
                defaultextension=f".{extension}",
                filetypes=[(f"{extension.upper()} files", f"*.{extension}")],
            )
        )
        set_picker_open(False)
        page.update()
        if not selected_path:
            return None
        return Path(selected_path)

    def export_results_by_type(export_type: str) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        if not table_state["results"]:
            notify("Belum ada hasil pemeriksaan untuk diexport.", error=True)
            return

        export_type = export_type.lower()
        export_map = {
            "pdf": ("Simpan summary PDF", "pdf", export_results_summary_pdf),
            "jpg": ("Simpan summary JPG", "jpg", export_results_summary_jpg),
            "json": ("Simpan summary JSON", "json", export_results_summary_json),
        }
        if export_type not in export_map:
            notify(f"Tipe export tidak dikenali: {export_type}", error=True)
            return

        dialog_title, extension, exporter = export_map[export_type]
        output_path = request_export_path(dialog_title, extension)
        if output_path is None:
            return

        try:
            exporter(output_path)
        except Exception as exc:
            log_path = write_error_log(f"Export {export_type.upper()} summary gagal", exc)
            notify(str(exc), error=True)
            notify(f"Log error tersimpan di: {log_path.name}", error=True)
            return

        notify(f"Summary {export_type.upper()} berhasil dibuat.")

    def handle_export_option(export_type: str) -> None:
        if getattr(export_dialog, "open", False):
            export_dialog.open = False
            page.pop_dialog()
            page.update()
        export_results_by_type(export_type)

    def show_export_dialog(_: ft.ControlEvent) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        if not table_state["results"]:
            notify("Belum ada hasil pemeriksaan untuk diexport.", error=True)
            return

        export_dialog.content = ft.Container(
            width=460,
            content=ft.Column(
                spacing=12,
                tight=True,
                controls=[
                    ft.Text(
                        "Pilih format file export untuk ringkasan hasil pemeriksaan.",
                        color="#4b5d58",
                    ),
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Button(
                                content=ft.Text("PDF", weight=ft.FontWeight.W_600),
                                icon=ft.Icons.PICTURE_AS_PDF_ROUNDED,
                                width=130,
                                style=ft.ButtonStyle(bgcolor="#0f5c4d", color="#fffaf0"),
                                on_click=lambda _: handle_export_option("pdf"),
                            ),
                            ft.Button(
                                content=ft.Text("JPG", weight=ft.FontWeight.W_600),
                                icon=ft.Icons.IMAGE_ROUNDED,
                                width=130,
                                style=ft.ButtonStyle(bgcolor="#c96f3b", color="#fffaf0"),
                                on_click=lambda _: handle_export_option("jpg"),
                            ),
                            ft.Button(
                                content=ft.Text("JSON", weight=ft.FontWeight.W_600),
                                icon=ft.Icons.DATA_OBJECT_ROUNDED,
                                width=130,
                                style=ft.ButtonStyle(bgcolor="#6b4f3a", color="#fffaf0"),
                                on_click=lambda _: handle_export_option("json"),
                            ),
                        ],
                    ),
                ],
            ),
        )
        export_dialog.actions = [ft.TextButton("Tutup", on_click=close_export_dialog)]
        if getattr(export_dialog, "open", False):
            page.update()
            return
        export_dialog.open = True
        page.show_dialog(export_dialog)

    def build_qr_image_bytes(text: str) -> bytes:
        qr = qrcode.QRCode(border=2, box_size=8)
        qr.add_data(text)
        qr.make(fit=True)
        image = qr.make_image(fill_color="black", back_color="white")
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        return buffer.getvalue()

    def show_share_dialog(result: dict[str, object]) -> None:
        summary_text_value = build_detail_summary_text(result)
        qr_image_bytes = build_qr_image_bytes(summary_text_value)

        share_dialog.title = ft.Text("Share Detail Completeness", weight=ft.FontWeight.W_700, color="#16302b")

        share_dialog.content = ft.Container(
            width=520,
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=12,
                tight=True,
                controls=[
                    ft.Text(result["file_name"], weight=ft.FontWeight.W_700, color="#16302b", text_align=ft.TextAlign.CENTER),
                    ft.Image(src=qr_image_bytes, width=260, height=260, fit=ft.BoxFit.CONTAIN),
                    ft.Container(
                        bgcolor="#f7efe0",
                        border=ft.Border.all(1, "#d9c9ab"),
                        border_radius=12,
                        padding=12,
                        content=ft.Text(summary_text_value, color="#4b5d58", selectable=True, size=12),
                    ),
                ],
            ),
        )
        share_dialog.actions = [
            ft.TextButton(
                "Copy Summary",
                icon=ft.Icons.CONTENT_COPY_ROUNDED,
                on_click=lambda _: page.run_task(copy_summary_to_clipboard, summary_text_value),
            ),
            ft.TextButton("Tutup", on_click=close_share_dialog),
        ]

        if getattr(detail_dialog, "open", False):
            detail_dialog.open = False
            page.pop_dialog()

        if getattr(share_dialog, "open", False):
            page.update()
            return

        share_dialog.open = True
        page.show_dialog(share_dialog)

    def show_follow_up_share_dialog(result: dict[str, object]) -> None:
        summary_text_value = build_follow_up_summary_text(result)
        qr_image_bytes = build_qr_image_bytes(summary_text_value)

        share_dialog.title = ft.Text("Share Follow Up Section 1.5", weight=ft.FontWeight.W_700, color="#16302b")
        share_dialog.content = ft.Container(
            width=520,
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=12,
                tight=True,
                controls=[
                    ft.Text(result["file_name"], weight=ft.FontWeight.W_700, color="#16302b", text_align=ft.TextAlign.CENTER),
                    ft.Image(src=qr_image_bytes, width=260, height=260, fit=ft.BoxFit.CONTAIN),
                    ft.Container(
                        bgcolor="#f7efe0",
                        border=ft.Border.all(1, "#d9c9ab"),
                        border_radius=12,
                        padding=12,
                        content=ft.Text(summary_text_value, color="#4b5d58", selectable=True, size=12),
                    ),
                ],
            ),
        )
        share_dialog.actions = [
            ft.TextButton(
                "Copy Summary",
                icon=ft.Icons.CONTENT_COPY_ROUNDED,
                on_click=lambda _: page.run_task(copy_summary_to_clipboard, summary_text_value),
            ),
            ft.TextButton("Tutup", on_click=close_share_dialog),
        ]

        if getattr(follow_up_dialog, "open", False):
            follow_up_dialog.open = False
            page.pop_dialog()

        if getattr(share_dialog, "open", False):
            page.update()
            return

        share_dialog.open = True
        page.show_dialog(share_dialog)

    def show_results_share_dialog(_: ft.ControlEvent) -> None:
        incomplete_only = False
        if incomplete_only:
            filtered_results = [result for result in table_state["results"] if not result.get("is_complete")]
            empty_message = "Belum ada file incomplete untuk dibagikan."
            dialog_title = "Share Daftar File INCOMPLETE"
            heading_text = "Daftar File INCOMPLETE"
        else:
            filtered_results = list(table_state["results"])
            empty_message = "Belum ada hasil pemeriksaan untuk dibagikan."
            dialog_title = "Share Daftar Hasil Pemeriksaan"
            heading_text = "Daftar Hasil Pemeriksaan"

        if not filtered_results:
            notify(empty_message, error=True)
            return

        summary_text_value = build_results_summary_text(incomplete_only)
        qr_image_bytes = build_qr_image_bytes(summary_text_value)

        share_dialog.title = ft.Text(dialog_title, weight=ft.FontWeight.W_700, color="#16302b")
        share_dialog.content = ft.Container(
            width=520,
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=12,
                tight=True,
                controls=[
                    ft.Text(heading_text, weight=ft.FontWeight.W_700, color="#16302b", text_align=ft.TextAlign.CENTER),
                    ft.Image(src=qr_image_bytes, width=260, height=260, fit=ft.BoxFit.CONTAIN),
                    ft.Container(
                        bgcolor="#f7efe0",
                        border=ft.Border.all(1, "#d9c9ab"),
                        border_radius=12,
                        padding=12,
                        content=ft.Text(summary_text_value, color="#4b5d58", selectable=True, size=12),
                    ),
                ],
            ),
        )
        share_dialog.actions = [
            ft.TextButton(
                "Copy Summary",
                icon=ft.Icons.CONTENT_COPY_ROUNDED,
                on_click=lambda _: page.run_task(copy_summary_to_clipboard, summary_text_value),
            ),
            ft.TextButton("Tutup", on_click=close_share_dialog),
        ]

        if getattr(detail_dialog, "open", False):
            detail_dialog.open = False
            page.pop_dialog()
        if getattr(follow_up_dialog, "open", False):
            follow_up_dialog.open = False
            page.pop_dialog()

        if getattr(share_dialog, "open", False):
            page.update()
            return

        share_dialog.open = True
        page.show_dialog(share_dialog)

    def show_info_dialog(_: ft.ControlEvent) -> None:
        def info_bullet(icon, title: str, description: str) -> ft.Control:
            return ft.Row(
                spacing=10,
                vertical_alignment=ft.CrossAxisAlignment.START,
                controls=[
                    ft.Icon(icon, size=18, color="#0f5c4d"),
                    ft.Column(
                        spacing=2,
                        tight=True,
                        expand=True,
                        controls=[
                            ft.Text(title, weight=ft.FontWeight.W_700, color="#16302b", size=13),
                            ft.Text(description, color="#4b5d58", size=12),
                        ],
                    ),
                ],
            )

        info_dialog.content = ft.Container(
            width=620,
            content=ft.Column(
                spacing=12,
                tight=True,
                controls=[
                    ft.Text("Informasi GUI", weight=ft.FontWeight.W_700, color="#16302b"),
                    ft.Text(
                        "Aplikasi ini dipakai untuk memeriksa completeness form IPS dari satu file Excel atau satu folder file Excel.",
                        color="#4b5d58",
                    ),
                    ft.Divider(color="#d9c9ab"),
                    ft.Text("Cara menggunakan", weight=ft.FontWeight.W_700, color="#16302b"),
                    info_bullet(ft.Icons.FOLDER_OPEN_ROUNDED, "Pilih sumber file", "Gunakan Pilih Folder untuk banyak file atau Pilih File untuk satu dokumen Excel."),
                    info_bullet(ft.Icons.PLAY_ARROW_ROUNDED, "Pemeriksaan otomatis", "Setelah file atau folder dipilih, proses check berjalan otomatis dan hasil muncul di panel kanan."),
                    info_bullet(ft.Icons.VISIBILITY_ROUNDED, "Lihat detail", "Icon mata pada kolom Action membuka detail completeness per file."),
                    info_bullet(ft.Icons.ASSIGNMENT_TURNED_IN_ROUNDED, "Lihat follow up", "Icon follow up menampilkan data section 1.5 yang terisi."),
                    info_bullet(ft.Icons.DOWNLOAD_ROUNDED, "Export hasil", "Gunakan tombol Export untuk memilih ringkasan PDF, JPG, atau JSON, dan Simpan CSV untuk lokasi file CSV."),
                    ft.Divider(color="#d9c9ab"),
                    ft.Text("Dokumen IPS dianggap complete jika", weight=ft.FontWeight.W_700, color="#16302b"),
                    info_bullet(ft.Icons.CHECK_CIRCLE_OUTLINE_ROUNDED, "Section 1.1 dan 1.2 valid", "Semua field wajib terisi, dengan trigger 1.1 cukup salah satu dari C6, C8, C10, atau C12."),
                    info_bullet(ft.Icons.RULE_FOLDER_OUTLINED, "Section 1.3 valid", "Setiap item hanya boleh punya satu nilai OK, NOK, atau NA, dan cell lain harus kosong."),
                    info_bullet(ft.Icons.TABLE_ROWS_ROUNDED, "Section 1.4 valid", "Minimal satu row penuh pada 66, 68, 70, 72, 74, 76, atau 78. Row terisi lain tidak boleh parsial."),
                    info_bullet(ft.Icons.TABLE_ROWS_ROUNDED, "Section 1.5 valid", "Minimal satu row penuh pada 88, 97, 106, 115, 124, atau 133. Row terisi lain tidak boleh parsial."),
                    info_bullet(ft.Icons.TASK_ALT_ROUNDED, "Status COMPLETE", "File dinyatakan complete jika semua rule section terpenuhi tanpa field atau row rule yang gagal."),
                ],
            ),
        )
        info_dialog.actions = [ft.TextButton("Tutup", on_click=close_info_dialog)]
        if getattr(info_dialog, "open", False):
            page.update()
            return
        info_dialog.open = True
        page.show_dialog(info_dialog)

    def show_result_details(result: dict[str, object]) -> None:
        status_label = "COMPLETE" if result["is_complete"] else "INCOMPLETE"
        section_completeness = result.get("section_completeness", {})
        total_fields_by_section = result.get("total_fields_by_section", {})
        missing_fields_by_section = result.get("missing_fields_by_section", {})

        section_table_rows = [
            ft.Container(
                bgcolor="#efe4ce",
                padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                content=ft.Row(
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Container(expand=True, content=ft.Text("Section", weight=ft.FontWeight.W_700, color="#16302b", size=12)),
                        ft.Container(width=90, content=ft.Text("%", weight=ft.FontWeight.W_700, color="#16302b", size=12)),
                    ],
                ),
            )
        ]
        for index, section in enumerate(sorted(section_completeness)):
            percentage = float(section_completeness.get(section, 0.0))
            section_table_rows.append(
                ft.Container(
                    bgcolor="#fffaf0" if index % 2 == 0 else "#fcf5e8",
                    border=ft.Border(bottom=ft.BorderSide(1, "#eadfcb")),
                    padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                    content=ft.Row(
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Container(
                                expand=True,
                                content=ft.Text(
                                    section_labels.get(str(section), str(section)),
                                    weight=ft.FontWeight.W_600,
                                    color="#16302b",
                                    size=12,
                                ),
                            ),
                            ft.Container(width=90, content=ft.Text(f"{percentage:.1f}%", color="#16302b", weight=ft.FontWeight.W_600, size=12)),
                        ],
                    ),
                )
            )

        detail_controls: list[ft.Control] = [
            ft.Text(f"File: {result['file_name']}", color="#16302b", weight=ft.FontWeight.W_700),
            ft.Text(f"Participant: {result.get('participant') or '-'}", color="#4b5d58"),
            ft.Text(f"Status: {status_label}", color="#4b5d58"),
            ft.Text(f"Completeness: {result['completeness_percentage']:.2f}%", color="#4b5d58"),
            ft.Text("Per Section", weight=ft.FontWeight.W_700, color="#16302b"),
            ft.Container(
                border=ft.Border.all(1, "#d9c9ab"),
                border_radius=14,
                clip_behavior=ft.ClipBehavior.HARD_EDGE,
                content=ft.Column(spacing=0, controls=section_table_rows),
            ),
        ]

        detail_dialog.content = ft.Container(
            width=760,
            content=ft.Column(spacing=10, tight=True, controls=detail_controls),
        )
        detail_dialog.actions = [
            ft.TextButton("Share", icon=ft.Icons.SHARE_ROUNDED, on_click=lambda _, item=result: show_share_dialog(item)),
            ft.TextButton("Tutup", on_click=close_detail_dialog),
        ]
        if getattr(detail_dialog, "open", False):
            page.update()
            return
        detail_dialog.open = True
        page.show_dialog(detail_dialog)

    def show_follow_up_details(result: dict[str, object]) -> None:
        follow_up_rows = result.get("follow_up_rows", [])
        follow_up_table_rows = [
            ft.Container(
                bgcolor="#efe4ce",
                padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                content=ft.Row(
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Container(expand=2, content=ft.Text("Countermeasure", weight=ft.FontWeight.W_700, color="#16302b", size=12)),
                        ft.Container(expand=1, content=ft.Text("Responsible", weight=ft.FontWeight.W_700, color="#16302b", size=12)),
                        ft.Container(width=120, content=ft.Text("Due Date", weight=ft.FontWeight.W_700, color="#16302b", size=12)),
                    ],
                ),
            )
        ]
        if follow_up_rows:
            for index, row in enumerate(follow_up_rows):
                follow_up_table_rows.append(
                    ft.Container(
                        bgcolor="#fffaf0" if index % 2 == 0 else "#fcf5e8",
                        border=ft.Border(bottom=ft.BorderSide(1, "#eadfcb")),
                        padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                        content=ft.Row(
                            vertical_alignment=ft.CrossAxisAlignment.START,
                            controls=[
                                ft.Container(expand=2, content=ft.Text(str(row.get("countermeasure") or "-"), color="#16302b", size=12)),
                                ft.Container(expand=1, content=ft.Text(str(row.get("responsible") or "-"), color="#4b5d58", size=12)),
                                ft.Container(width=120, content=ft.Text(str(row.get("due_date") or "-"), color="#4b5d58", size=12)),
                            ],
                        ),
                    )
                )
        else:
            follow_up_table_rows.append(
                ft.Container(
                    padding=16,
                    content=ft.Text("Belum ada data follow up pada section 1.5.", color="#6b4f3a"),
                )
            )

        follow_up_dialog.content = ft.Container(
            width=760,
            content=ft.Column(
                spacing=10,
                tight=True,
                controls=[
                    ft.Text(f"File: {result['file_name']}", color="#16302b", weight=ft.FontWeight.W_700),
                    ft.Text("Section 1.5 - Ask WHY WHY", color="#4b5d58"),
                    ft.Container(
                        border=ft.Border.all(1, "#d9c9ab"),
                        border_radius=14,
                        clip_behavior=ft.ClipBehavior.HARD_EDGE,
                        content=ft.Column(spacing=0, controls=follow_up_table_rows),
                    ),
                ],
            ),
        )
        follow_up_dialog.actions = [
            ft.TextButton("Share", icon=ft.Icons.SHARE_ROUNDED, on_click=lambda _, item=result: show_follow_up_share_dialog(item)),
            ft.TextButton("Tutup", on_click=close_follow_up_dialog),
        ]
        if getattr(follow_up_dialog, "open", False):
            page.update()
            return
        follow_up_dialog.open = True
        page.show_dialog(follow_up_dialog)

    def sort_results(column_index: int, ascending: bool) -> None:
        table_state["sort_column_index"] = column_index
        table_state["sort_ascending"] = ascending

        results = list(table_state["results"])
        if column_index == 0:
            results.sort(key=lambda item: str(item["file_name"]).casefold(), reverse=not ascending)
        elif column_index == 1:
            results.sort(key=lambda item: float(item["completeness_percentage"]), reverse=not ascending)

        table_state["results"] = results
        refresh_results_table()

    def toggle_sort(column_index: int) -> None:
        if table_state["sort_column_index"] == column_index:
            sort_results(column_index, not table_state["sort_ascending"])
            return
        sort_results(column_index, True)

    def build_header_cell(label: str, column_index: int | None = None) -> ft.Control:
        sort_suffix = ""
        if column_index is not None and table_state["sort_column_index"] == column_index:
            sort_suffix = "  ^" if table_state["sort_ascending"] else "  v"

        label_text = ft.Text(
            f"{label}{sort_suffix}",
            weight=ft.FontWeight.W_700,
            color="#16302b",
        )
        if column_index is None:
            return label_text
        return ft.TextButton(
            content=label_text,
            style=ft.ButtonStyle(
                padding=0,
                color="#16302b",
                overlay_color="#00000000",
            ),
            on_click=lambda _: toggle_sort(column_index),
        )

    def build_result_row(result: dict[str, object], row_index: int) -> ft.Control:
        row_bgcolor = "#fffaf0" if row_index % 2 == 0 else "#fcf5e8"
        return ft.Container(
            bgcolor=row_bgcolor,
            border=ft.Border(bottom=ft.BorderSide(1, "#eadfcb")),
            padding=ft.Padding.symmetric(horizontal=16, vertical=14),
            content=ft.Row(
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Container(
                        expand=True,
                        content=ft.Text(
                            str(result["file_name"]),
                            color="#16302b",
                            weight=ft.FontWeight.W_600,
                            overflow=ft.TextOverflow.ELLIPSIS,
                            max_lines=1,
                        ),
                    ),
                    ft.Container(
                        width=190,
                        content=build_completeness_cell(result),
                    ),
                    ft.Container(
                        width=96,
                        content=ft.Row(
                            spacing=0,
                            alignment=ft.MainAxisAlignment.CENTER,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            controls=[
                                ft.IconButton(
                                    icon=ft.Icons.VISIBILITY_ROUNDED,
                                    icon_color="#0f5c4d",
                                    icon_size=18,
                                    tooltip="Lihat detail",
                                    on_click=lambda _, item=result: show_result_details(item),
                                ),
                                ft.IconButton(
                                    icon=ft.Icons.ASSIGNMENT_TURNED_IN_ROUNDED,
                                    icon_color="#0f5c4d",
                                    icon_size=18,
                                    tooltip="Lihat follow up",
                                    on_click=lambda _, item=result: show_follow_up_details(item),
                                ),
                            ],
                        ),
                    ),
                ],
            ),
        )

    def refresh_results_table() -> None:
        header_row = ft.Container(
            bgcolor="#efe4ce",
            padding=ft.Padding.symmetric(horizontal=16, vertical=14),
            content=ft.Row(
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Container(expand=True, content=build_header_cell("File Name", 0)),
                    ft.Container(width=190, content=build_header_cell("Completeness %", 1)),
                    ft.Container(width=96, content=build_header_cell("Action")),
                ],
            ),
        )
        result_rows = [
            build_result_row(result, index)
            for index, result in enumerate(table_state["results"])
        ]
        if not result_rows:
            result_rows = [
                ft.Container(
                    padding=24,
                    content=ft.Text("Belum ada hasil pemeriksaan.", color="#6b4f3a"),
                )
            ]
        results_table.controls = [header_row, *result_rows]

    refresh_results_table()

    def start_run_check(selected_target: str | None = None) -> None:
        if selected_target:
            app_state["target_path"] = selected_target
        persist_config()
        set_loading(True)
        status_text.value = "Sedang memeriksa file Excel..."
        table_state["results"] = []
        refresh_results_table()
        summary_text.value = ""
        page.update()
        page.run_task(run_check, selected_target)

    async def run_check(selected_target: str | None = None) -> None:

        try:
            target_path = selected_target or app_state["target_path"]
            results = await asyncio.to_thread(
                run_checks,
                target_path=target_path,
                field_file=app_state["field_file"],
                sheet_name=app_state["sheet_name"],
                csv_output=app_state["csv_output"],
                base_dir=APP_DIR,
            )
        except Exception as exc:
            log_path = write_error_log("Pemeriksaan gagal", exc)
            status_text.value = "Pemeriksaan gagal."
            set_loading(False)
            page.update()
            notify(str(exc), error=True)
            notify(f"Log error tersimpan di: {log_path.name}", error=True)
            return

        incomplete_count = sum(1 for result in results if not result["is_complete"])
        for result in results:
            result["section_summary"] = " | ".join(
                f"Section {section}: {percentage:.2f}% ({result['total_fields_by_section'].get(section, 0) - result['missing_fields_by_section'].get(section, 0)}/{result['total_fields_by_section'].get(section, 0)})"
                for section, percentage in sorted(result.get("section_completeness", {}).items())
            )
        summary_text.value = (
            f"{len(results)} file diperiksa | {incomplete_count} file incomplete | "
            f"{len(results) - incomplete_count} file complete"
        )
        status_text.value = "Pemeriksaan selesai."
        table_state["results"] = list(results)
        sort_results(table_state["sort_column_index"], table_state["sort_ascending"])
        set_loading(False)
        page.update()
        notify("Pemeriksaan selesai.")

    form_panel = ft.Container(
        expand=True,
        bgcolor="#fffaf0",
        border_radius=24,
        padding=24,
        content=ft.Column(
            expand=True,
            spacing=18,
            controls=[
                app_logo,
                ft.Text(
                    "Pilih folder atau file Excel. Pemeriksaan akan langsung berjalan otomatis.",
                    color="#6b4f3a",
                ),
                folder_button,
                file_button,
                csv_button,
                share_button,
                info_button,
                ft.Container(expand=True),
                ft.Divider(color="#e6dac6", height=1),
                ft.Container(
                    width=side_button_width,
                    padding=ft.Padding.only(top=4),
                    content=ft.Text(
                        "© 2026 rardyant",
                        size=10,
                        color="#9a8874",
                        text_align=ft.TextAlign.CENTER,
                    ),
                ),
            ],
        ),
    )

    result_panel = ft.Container(
        expand=True,
        bgcolor="#e9dfc9",
        border_radius=24,
        padding=24,
        content=ft.Column(
            expand=True,
            spacing=16,
            controls=[
                ft.Text("Hasil Pemeriksaan", size=24, weight=ft.FontWeight.W_700, color="#16302b"),
                summary_text,
                loading_row,
                status_text,
                ft.Divider(color="#c6b89b"),
                table_container,
            ],
        ),
    )

    folder_button.on_click = pick_target_folder
    file_button.on_click = pick_target_file
    csv_button.on_click = pick_csv_output
    share_button.on_click = show_export_dialog
    info_button.on_click = show_info_dialog
    page.on_resize = handle_page_resize
    page.on_close = handle_page_close
    sync_window_size_from_page()
    persist_config()

    page.add(
        ft.ResponsiveRow(
            columns=12,
            controls=[
                ft.Container(col={"xs": 12, "lg": 2}, content=form_panel),
                ft.Container(col={"xs": 12, "lg": 10}, content=result_panel),
            ],
        )
    )
    page.services.append(clipboard_service)


if __name__ == "__main__":
    try:
        ft.run(main)
    except Exception as exc:
        log_path = write_error_log("Aplikasi gagal dijalankan", exc)
        raise RuntimeError(f"Aplikasi gagal dijalankan. Lihat log: {log_path}") from exc
