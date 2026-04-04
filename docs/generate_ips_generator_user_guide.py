from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from svglib.svglib import svg2rlg

ROOT = Path(__file__).resolve().parent
OUTPUT_FILE = ROOT / "ips-generator-user-guide.pdf"
MOCKUP_FILE = ROOT / "ips-generator-mockup-pdf.svg"


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


def build_mockup_visual(max_width: float = 17.2 * cm):
    drawing = svg2rlg(str(MOCKUP_FILE))
    if drawing.width > max_width:
        scale_factor = max_width / drawing.width
        drawing.scale(scale_factor, scale_factor)
        drawing.width *= scale_factor
        drawing.height *= scale_factor
    return drawing


def build_component_table(styles):
    cells = [
        [
            Paragraph("<b>Area</b>", styles["GuideBody"]),
            Paragraph("<b>Fungsi</b>", styles["GuideBody"]),
        ],
        [
            Paragraph("Hero", styles["GuideBody"]),
            Paragraph(
                "Menjelaskan tujuan IPS Generator dan menyediakan tombol untuk menampilkan Quick Guide.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Quick Guide", styles["GuideBody"]),
            Paragraph(
                "Kartu panduan interaktif yang dapat mengarahkan pengguna ke area form, trigger, dan tombol unduh.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Form Input", styles["GuideBody"]),
            Paragraph(
                "Area utama untuk mengisi Tanggal, Line, Technology, Trigger, dan Problem Description ke template IPS Excel.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Status Panel", styles["GuideBody"]),
            Paragraph(
                "Menampilkan status proses dan tombol Unduh hasil yang aktif setelah field wajib lengkap.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Footer", styles["GuideBody"]),
            Paragraph(
                "Menampilkan versi IPS Generator dan versi template workbook yang digunakan.",
                styles["GuideBody"],
            ),
        ],
    ]
    table = Table(cells, colWidths=[4.4 * cm, 11.6 * cm], repeatRows=1)
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


def build_field_table(styles):
    cells = [
        [
            Paragraph("<b>Field</b>", styles["GuideBody"]),
            Paragraph("<b>Cell Excel</b>", styles["GuideBody"]),
            Paragraph("<b>Wajib</b>", styles["GuideBody"]),
            Paragraph("<b>Catatan</b>", styles["GuideBody"]),
        ],
        [
            Paragraph("Tanggal", styles["GuideBody"]),
            Paragraph("P8", styles["GuideBody"]),
            Paragraph("Ya", styles["GuideBody"]),
            Paragraph(
                "Nilai tanggal ikut dipakai untuk nama file dan sequence Document ID.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Line", styles["GuideBody"]),
            Paragraph("P6", styles["GuideBody"]),
            Paragraph("Ya", styles["GuideBody"]),
            Paragraph(
                "Pilih dari daftar line preset. Prefix line dipakai pada nama file dan Document ID.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Technology", styles["GuideBody"]),
            Paragraph("Q14", styles["GuideBody"]),
            Paragraph("Ya", styles["GuideBody"]),
            Paragraph(
                "Dipilih dari daftar technology preset dan dipakai sebagai bagian nama file serta Document ID.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Trigger", styles["GuideBody"]),
            Paragraph("C6 / C8 / C10 / C12", styles["GuideBody"]),
            Paragraph("Tidak", styles["GuideBody"]),
            Paragraph(
                "Checkbox akan menulis Q, D, C, atau O ke cell trigger sesuai pilihan pengguna.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Problem Description", styles["GuideBody"]),
            Paragraph("B19", styles["GuideBody"]),
            Paragraph("Ya", styles["GuideBody"]),
            Paragraph(
                "Teks dipaksa uppercase sebelum diunduh dan menjadi bagian nama file output.",
                styles["GuideBody"],
            ),
        ],
        [
            Paragraph("Document ID", styles["GuideBody"]),
            Paragraph("W6", styles["GuideBody"]),
            Paragraph("Otomatis", styles["GuideBody"]),
            Paragraph(
                "Dibuat otomatis dengan format IPS-LINE-TECH-YYYYMMDD-SEQ dan sequence disimpan di localStorage browser.",
                styles["GuideBody"],
            ),
        ],
    ]
    table = Table(
        cells, colWidths=[3.5 * cm, 2.7 * cm, 1.8 * cm, 8.0 * cm], repeatRows=1
    )
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


def build_story():
    styles = build_styles()
    story = []

    story.append(Paragraph("User Guide IPS Generator", styles["GuideTitle"]))
    story.append(
        Paragraph(
            "Panduan singkat penggunaan halaman assets/ips.html untuk membuat form IPS baru dan mengunduh workbook Excel yang sudah terisi otomatis.",
            styles["GuideBody"],
        )
    )
    story.append(
        Paragraph(
            "Mockup visual PDF-safe tersedia di folder docs dengan nama ips-generator-mockup-pdf.svg.",
            styles["GuideCaption"],
        )
    )

    story.append(Paragraph("Gambaran Halaman", styles["GuideHeading"]))
    story.append(
        Paragraph(
            "IPS Generator memuat template workbook default, menyediakan Quick Guide interaktif, lalu menulis field utama ke template Excel saat pengguna menekan tombol Unduh hasil.",
            styles["GuideBody"],
        )
    )
    story.append(build_mockup_visual())
    story.append(
        Paragraph(
            "Mockup ini merangkum hero section, area Quick Guide, form utama, dan panel status unduhan pada halaman IPS Generator.",
            styles["GuideCaption"],
        )
    )
    story.append(build_component_table(styles))

    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Alur Pembuatan IPS Baru", styles["GuideHeading"]))
    for bullet in [
        "Buka halaman <b>IPS Generator</b> dari tombol <b>Buat IPS</b> pada aplikasi utama atau langsung melalui file <b>assets/ips.html</b>.",
        "Jika perlu panduan cepat di browser, klik tombol <b>Tampilkan Quick Guide</b> untuk melihat kartu langkah visual.",
        "Isi field wajib secara berurutan: <b>Tanggal</b>, <b>Line</b>, <b>Technology</b>, lalu <b>Problem Description</b>.",
        "Pilih satu atau lebih <b>Trigger</b> yang sesuai. Trigger akan ditulis ke cell C6, C8, C10, atau C12 sesuai kategori yang dicentang.",
        "Saat Problem Description diisi, sistem akan menjaga nilainya dalam format <b>uppercase</b> sebelum workbook diunduh.",
        "Tombol <b>Unduh hasil</b> akan aktif setelah seluruh field wajib lengkap. Sistem lalu menulis semua perubahan ke workbook template.",
        "Saat unduhan dibuat, sistem juga menghasilkan <b>Document ID</b> otomatis di cell <b>W6</b> dan nama file output berdasarkan tanggal, line, technology, serta judul masalah.",
        "Buka workbook hasil unduhan untuk melanjutkan evidence, why analysis, dan proses investigasi internal lainnya.",
    ]:
        story.append(Paragraph(bullet, styles["GuideBullet"], bulletText="•"))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Pemetaan Field", styles["GuideHeading"]))
    story.append(build_field_table(styles))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Perilaku Otomatis", styles["GuideHeading"]))
    for bullet in [
        "Template workbook default dimuat otomatis saat halaman dibuka, sehingga pengguna tidak perlu memilih file template secara manual.",
        "Nama file output dibangun dari pola tanggal, literal IPS dan SPM, prefix line, technology, lalu Problem Description yang sudah disanitasi.",
        "Sequence Document ID disimpan per tanggal melalui localStorage browser dengan prefix penyimpanan <b>ipsDocumentSequence</b>.",
        "Worksheet target utama adalah <b>IPS V4.1</b>. Sistem juga memaksa recalculation workbook sebelum file diunduh.",
    ]:
        story.append(Paragraph(bullet, styles["GuideBullet"], bulletText="•"))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Tips Penggunaan", styles["GuideHeading"]))
    for bullet in [
        "Jika tombol Unduh hasil masih nonaktif, periksa kembali field wajib: Tanggal, Line, Technology, dan Problem Description.",
        "Gunakan Quick Guide bila ingin lompat cepat ke area field tertentu karena setiap kartu panduan dapat men-scroll ke target input terkait.",
        "Periksa file Excel hasil unduhan untuk memastikan cell B19, P8, P6, Q14, dan W6 sudah terisi sesuai kebutuhan sebelum didistribusikan.",
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
