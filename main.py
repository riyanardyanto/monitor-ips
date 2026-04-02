from __future__ import annotations

import asyncio
from datetime import datetime
from io import BytesIO
from pathlib import Path
import sys
import tkinter as tk
import traceback
from tkinter import filedialog

import flet as ft
import qrcode
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

    page.title = "IPS Completeness Checker"
    page.window_width = 1180
    page.window_height = 820
    page.padding = 24
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#f4f1e8"
    page.scroll = ft.ScrollMode.AUTO
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
        "target_path": str(default_target_path),
        "field_file": str(default_field_file),
        "sheet_name": None,
        "csv_output": str(APP_DIR / "report.csv"),
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
        content=ft.Text("Export PDF", weight=ft.FontWeight.W_600),
        icon=ft.Icons.DOWNLOAD_ROUNDED,
        width=side_button_width,
        style=ft.ButtonStyle(
            bgcolor="#e9dfc9",
            color="#16302b",
            padding=ft.Padding.symmetric(horizontal=18, vertical=16),
            shape=ft.RoundedRectangleBorder(radius=14),
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
                initialdir=str(APP_DIR),
            )
        )
        if selected_path:
            app_state["target_path"] = selected_path
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
                initialdir=str(APP_DIR),
                filetypes=[("Excel files", "*.xlsx")],
            )
        )
        if selected_path:
            app_state["target_path"] = selected_path
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
                initialdir=str(APP_DIR),
                initialfile="report.csv",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
            )
        )
        if selected_path:
            app_state["csv_output"] = selected_path
        set_picker_open(False)
        page.update()

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

    def export_results_summary_pdf_action(_: ft.ControlEvent) -> None:
        if app_state["is_picker_open"] or app_state["is_loading"]:
            return
        if not table_state["results"]:
            notify("Belum ada hasil pemeriksaan untuk dibuatkan PDF.", error=True)
            return

        set_picker_open(True)
        page.update()
        default_pdf_name = f"ips-completeness-summary-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
        selected_path = with_hidden_tk_root(
            lambda _: filedialog.asksaveasfilename(
                title="Simpan summary PDF",
                initialdir=str(APP_DIR),
                initialfile=default_pdf_name,
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
            )
        )
        set_picker_open(False)
        page.update()

        if not selected_path:
            return

        try:
            export_results_summary_pdf(Path(selected_path))
        except Exception as exc:
            log_path = write_error_log("Export PDF summary gagal", exc)
            notify(str(exc), error=True)
            notify(f"Log error tersimpan di: {log_path.name}", error=True)
            return

        notify("Summary PDF berhasil dibuat.")

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
                ft.Text("IPS Form Completeness", size=30, weight=ft.FontWeight.W_700, color="#16302b"),
                ft.Text(
                    "Pilih folder atau file Excel. Pemeriksaan akan langsung berjalan otomatis.",
                    color="#6b4f3a",
                ),
                folder_button,
                file_button,
                csv_button,
                share_button,
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
    share_button.on_click = export_results_summary_pdf_action

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
