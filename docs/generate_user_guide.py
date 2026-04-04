from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from svglib.svglib import svg2rlg

ROOT = Path(__file__).resolve().parent
OUTPUT_FILE = ROOT / "ips-checker-user-guide.pdf"
MOCKUP_FILE = ROOT / "ips-checker-mockup.svg"
PDF_MOCKUP_FILE = ROOT / "ips-checker-mockup-pdf.svg"


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="GuideTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=30,
            textColor=colors.HexColor("#16302b"),
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideHeading",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=15,
            leading=19,
            textColor=colors.HexColor("#16302b"),
            spaceBefore=10,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#4b5d58"),
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideBullet",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#4b5d58"),
            leftIndent=12,
            bulletIndent=0,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="GuideCaption",
            parent=styles["BodyText"],
            fontName="Helvetica-Oblique",
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#6b4f3a"),
            spaceAfter=8,
        )
    )
    return styles


def build_mockup_table(styles):
    cells = [
        [
            Paragraph("<b>Area</b>", styles["GuideBody"]),
            Paragraph("<b>Fungsi</b>", styles["GuideBody"]),
        ],
        [
            Paragraph("Header / AppBar", styles["GuideBody"]),
            Paragraph(
                "Menampilkan judul view aktif, tombol menu untuk pindah view, dan tombol Export.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Side Panel", styles["GuideBody"]),
            Paragraph(
                "Berisi tombol Pilih Folder, Pilih File, Buat IPS, dan Info untuk memulai alur kerja.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Tabel Hasil", styles["GuideBody"]),
            Paragraph(
                "Menampilkan nama file, completeness persen, tombol lihat detail, dan open file asli.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("View Follow Up", styles["GuideBody"]),
            Paragraph(
                "Menampilkan database follow up lokal otomatis lengkap dengan status, Standard ID, edit, dan hapus data.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Output Export", styles["GuideBody"]),
            Paragraph(
                "Hasil pemeriksaan: PDF, JPG, QR Code. Follow up: PDF, JPG infografis, Excel.",
                styles["GuideBody"],
            ),
        ],
    ]
    table = Table(cells, colWidths=[4.5 * cm, 11.5 * cm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#efe4ce")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#16302b")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d9c9ab")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#fffaf0")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return table


def build_export_table(styles):
    cells = [
        [
            Paragraph("<b>View Aktif</b>", styles["GuideBody"]),
            Paragraph("<b>Format Export</b>", styles["GuideBody"]),
            Paragraph("<b>Kegunaan</b>", styles["GuideBody"]),
        ],
        [
            Paragraph("Hasil Pemeriksaan", styles["GuideBody"]),
            Paragraph("PDF / JPG / QR Code", styles["GuideBody"]),
            Paragraph(
                "Ringkasan hasil check untuk laporan, distribusi, dan share cepat.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Follow Up Countermeasure", styles["GuideBody"]),
            Paragraph("PDF / JPG Infografis / Excel", styles["GuideBody"]),
            Paragraph(
                "Dokumentasi follow up, presentasi visual, atau olah data lanjutan di Excel.",
                styles["GuideBody"],
            ),
        ],
    ]
    table = Table(cells, colWidths=[4.2 * cm, 4.5 * cm, 7.3 * cm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16302b")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#fffaf0")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d9c9ab")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#fffaf0")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return table


def build_mockup_visual(max_width: float = 17.0 * cm):
    drawing = svg2rlg(str(PDF_MOCKUP_FILE))

    if drawing.width > max_width:
        scale_factor = max_width / drawing.width
        drawing.scale(scale_factor, scale_factor)
        drawing.width *= scale_factor
        drawing.height *= scale_factor

    return drawing


def build_story():
    styles = build_styles()
    story = []

    story.append(Paragraph("User Guide IPS Checker", styles["GuideTitle"]))
    story.append(
        Paragraph(
            "Panduan singkat penggunaan aplikasi desktop untuk memeriksa completeness form IPS dan mengelola follow up countermeasure.",
            styles["GuideBody"],
        )
    )
    story.append(
        Paragraph(
            "File pendamping mockup tersedia di folder docs: ips-checker-mockup.svg untuk versi utama dan ips-checker-mockup-pdf.svg untuk versi PDF-safe.",
            styles["GuideCaption"],
        )
    )

    story.append(Paragraph("Gambaran Layar", styles["GuideHeading"]))
    story.append(
        Paragraph(
            "Tampilan aplikasi terdiri dari side panel di kiri untuk memilih sumber file dan panel utama di kanan untuk tabel hasil, follow up, serta tombol export pada AppBar.",
            styles["GuideBody"],
        )
    )
    story.append(build_mockup_visual())
    story.append(
        Paragraph(
            "Mockup visual di bawah ini menampilkan posisi tombol Buat IPS secara eksplisit pada side panel agar konsisten dengan aplikasi saat ini.",
            styles["GuideCaption"],
        )
    )
    story.append(build_mockup_table(styles))

    story.append(Spacer(1, 0.35 * cm))
    story.append(Paragraph("Alur Penggunaan", styles["GuideHeading"]))
    for bullet in [
        "Pilih sumber file melalui tombol <b>Pilih Folder</b> untuk banyak file atau <b>Pilih File</b> untuk satu dokumen Excel.",
        "Gunakan tombol <b>Buat IPS</b> di side panel jika Anda ingin membuka halaman generator IPS dan menyiapkan dokumen IPS baru.",
        "Setelah dipilih, pemeriksaan berjalan otomatis dan hasil tampil di panel kanan tanpa langkah tambahan.",
        "Tinjau tabel hasil pemeriksaan, buka detail dengan icon mata, atau buka file asli dengan icon open file.",
        "Gunakan tombol menu pada AppBar untuk pindah ke view <b>Follow Up Countermeasure</b>.",
        "Di view follow up, Anda dapat melihat detail, mengubah status Open atau Close, mengisi Standard ID saat Close, serta menghapus data yang tidak diperlukan.",
        "Gunakan tombol <b>Export</b> di AppBar. Format export menyesuaikan view yang sedang aktif.",
    ]:
        story.append(Paragraph(bullet, styles["GuideBullet"], bulletText="•"))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Format Export", styles["GuideHeading"]))
    story.append(build_export_table(styles))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Aturan Completeness IPS", styles["GuideHeading"]))
    for bullet in [
        "Section 1.1 dan 1.2 dinilai valid bila semua field wajib terisi. Untuk trigger 1.1, cukup salah satu cell pemicu yang terisi.",
        "Section 1.3 valid bila setiap item hanya memiliki satu nilai OK, NOK, atau NA dan cell lain tetap kosong.",
        "Section 1.4 valid bila minimal satu row penuh terisi dan row terisi lain tidak parsial.",
        "Section 1.5 valid bila minimal satu row penuh terisi dan row terisi lain tidak parsial.",
        "Status COMPLETE hanya diberikan bila semua rule section terpenuhi.",
    ]:
        story.append(Paragraph(bullet, styles["GuideBullet"], bulletText="•"))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Tips Penggunaan", styles["GuideHeading"]))
    for bullet in [
        "Gunakan export JPG pada view follow up bila Anda membutuhkan infografis untuk presentasi cepat.",
        "Gunakan export Excel pada view follow up bila data perlu diedit atau dianalisis lebih lanjut.",
        "Gunakan tombol Buat IPS bila Anda perlu membuat dokumen IPS baru sebelum menjalankan proses pengecekan.",
        "Gunakan dialog Info di aplikasi untuk ringkasan alur penggunaan tanpa membuka dokumen ini.",
    ]:
        story.append(Paragraph(bullet, styles["GuideBullet"], bulletText="•"))

    return story


def main() -> None:
    document = SimpleDocTemplate(
        str(OUTPUT_FILE),
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.3 * cm,
    )
    document.build(build_story())
    print(f"Generated: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
