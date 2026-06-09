# -*- coding: utf-8 -*-
"""
Generator Bahan Ajar PowerPoint
Mata Kuliah: Seminar Proposal Penelitian (Tesis)
Program Studi: Magister Manajemen Pendidikan (S2)
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

# ----------------------------------------------------------------------------
# PALET WARNA & FONT
# ----------------------------------------------------------------------------
NAVY   = RGBColor(0x0F, 0x2A, 0x4A)   # biru tua akademik
BLUE   = RGBColor(0x1B, 0x5E, 0x8C)   # biru
TEAL   = RGBColor(0x2A, 0x9D, 0x8F)   # teal
AMBER  = RGBColor(0xE9, 0xA2, 0x3B)   # emas/amber (aksen)
CORAL  = RGBColor(0xE7, 0x6F, 0x51)   # aksen hangat
LIGHT  = RGBColor(0xF5, 0xF8, 0xFC)   # latar terang
CARD   = RGBColor(0xFF, 0xFF, 0xFF)   # kartu putih
GREY   = RGBColor(0x5B, 0x6B, 0x7B)   # abu teks sekunder
TEXT   = RGBColor(0x1C, 0x2B, 0x3A)   # teks utama
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
SOFT   = RGBColor(0xE7, 0xEF, 0xF7)   # biru sangat muda

FONT_H = "Georgia"          # judul (serif elegan)
FONT_B = "Calibri"          # isi (sans modern)

IMG = "assets/img"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ----------------------------------------------------------------------------
# HELPER
# ----------------------------------------------------------------------------
def slide():
    return prs.slides.add_slide(BLANK)

def bg(s, color):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color

def no_line(shp):
    shp.line.fill.background()

def rect(s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=None, line_w=None):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    if line is None:
        no_line(sp)
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp

def grad(s, x, y, w, h, c1, c2, angle=45):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    no_line(sp)
    f = sp.fill
    f.gradient()
    stops = f.gradient_stops
    stops[0].position = 0.0
    stops[0].color.rgb = c1
    stops[1].position = 1.0
    stops[1].color.rgb = c2
    try:
        f.gradient_angle = angle
    except Exception:
        pass
    return sp

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=6, line_spacing=1.0, wrap=True):
    """runs: list of paragraphs; each paragraph = list of (text, size, color, bold, italic, font)"""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (t, sz, col, bold, ital, fnt) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col
            r.font.bold = bold; r.font.italic = ital
            r.font.name = fnt
    return tb

def R(t, sz, col=TEXT, bold=False, ital=False, fnt=FONT_B):
    return (t, sz, col, bold, ital, fnt)

def pic_cover(s, path, x, y, w, h):
    """Tambahkan gambar dengan crop center agar mengisi area w x h (object-fit cover)."""
    im = Image.open(path)
    iw, ih = im.size
    tw, th = w / h, iw / ih  # target ratio vs image ratio
    target = w / h
    src = iw / ih
    pic = s.shapes.add_picture(path, x, y, w, h)
    if src > target:
        # gambar lebih lebar -> crop kiri/kanan
        new = target / src
        crop = (1 - new) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        new = src / target
        crop = (1 - new) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    return pic

def footer(s, idx, total=18):
    bar = rect(s, 0, SH - Inches(0.32), SW, Inches(0.32), NAVY)
    txt(s, Inches(0.45), SH - Inches(0.32), Inches(9), Inches(0.32),
        [[R("Seminar Proposal Penelitian (Tesis)  •  Magister Manajemen Pendidikan",
            9, RGBColor(0xC7,0xD6,0xE6), False, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, SW - Inches(1.6), SH - Inches(0.32), Inches(1.15), Inches(0.32),
        [[R(f"{idx:02d} / {total}", 9, AMBER, True, False, FONT_B)]],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def header(s, kicker, title, idx):
    """Header standar konten: bar atas + nomor + kicker + judul."""
    rect(s, 0, 0, SW, Inches(1.28), NAVY)
    rect(s, 0, Inches(1.28), SW, Inches(0.07), AMBER)
    # badge nomor
    badge = rect(s, Inches(0.5), Inches(0.30), Inches(0.72), Inches(0.72), AMBER,
                 shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.5), Inches(0.30), Inches(0.72), Inches(0.72),
        [[R(f"{idx:02d}", 22, NAVY, True, False, FONT_H)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.42), Inches(0.26), Inches(11), Inches(0.34),
        [[R(kicker.upper(), 12, AMBER, True, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.42), Inches(0.56), Inches(11.4), Inches(0.68),
        [[R(title, 26, WHITE, True, False, FONT_H)]],
        anchor=MSO_ANCHOR.MIDDLE)

def ref_note(s, text):
    """Catatan referensi kecil di kaki slide konten."""
    txt(s, Inches(0.5), SH - Inches(0.62), Inches(12.3), Inches(0.28),
        [[R("Referensi: ", 8.5, GREY, True, True, FONT_B), R(text, 8.5, GREY, False, True, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE)

def card(s, x, y, w, h, fill=CARD, accent=None):
    c = rect(s, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try:
        c.adjustments[0] = 0.06
    except Exception:
        pass
    # soft shadow
    sh = c.shadow
    sh.inherit = False
    if accent:
        rect(s, x, y, Inches(0.10), h, accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    return c

def bullets(s, x, y, w, h, items, size=15, gap=10, color=TEXT, marker_col=TEAL,
            line_spacing=1.05):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.space_before = Pt(0); p.line_spacing = line_spacing
        if isinstance(it, tuple):
            head, body = it
        else:
            head, body = None, it
        rb = p.add_run(); rb.text = "▸  "
        rb.font.size = Pt(size); rb.font.color.rgb = marker_col; rb.font.bold = True; rb.font.name = FONT_B
        if head:
            rh = p.add_run(); rh.text = head + ": "
            rh.font.size = Pt(size); rh.font.color.rgb = NAVY; rh.font.bold = True; rh.font.name = FONT_B
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(size); rt.font.color.rgb = color; rt.font.name = FONT_B
    return tb

slides_done = 0


# ============================================================================
# SLIDE 1 — COVER
# ============================================================================
s = slide()
bg(s, NAVY)
# panel gambar kanan
pic_cover(s, f"{IMG}/cover.jpg", Inches(7.5), 0, Inches(5.833), SH)
# blok teks kiri (menutup tepi gambar agar menyatu)
rect(s, 0, 0, Inches(7.7), SH, NAVY)
rect(s, 0, Inches(2.55), Inches(1.7), Inches(0.10), AMBER)
txt(s, Inches(0.7), Inches(0.55), Inches(6.5), Inches(0.4),
    [[R("BAHAN AJAR  •  PERTEMUAN INTI", 13, AMBER, True, False, FONT_B)]])
txt(s, Inches(0.7), Inches(1.05), Inches(6.7), Inches(1.6),
    [[R("Seminar Proposal", 40, WHITE, True, False, FONT_H)],
     [R("Penelitian (Tesis)", 40, WHITE, True, False, FONT_H)]],
    line_spacing=1.0, space_after=2)
txt(s, Inches(0.7), Inches(2.85), Inches(6.6), Inches(1.2),
    [[R("Merancang Proposal Tesis yang Logis, Sistematis, dan Layak Diteliti",
        16, SOFT, False, True, FONT_B)]], line_spacing=1.1)
# meta program
txt(s, Inches(0.7), Inches(4.15), Inches(6.6), Inches(1.5),
    [[R("Program Studi", 11, AMBER, True, False, FONT_B)],
     [R("Magister Manajemen Pendidikan (S2)", 17, WHITE, True, False, FONT_B)],
     [R("Jenjang Pascasarjana  |  18 Slide Materi + Referensi", 12, SOFT, False, False, FONT_B)]],
    space_after=4)
# footer kecil
txt(s, Inches(0.7), Inches(6.65), Inches(6.6), Inches(0.5),
    [[R("Disusun untuk pembelajaran mahasiswa pascasarjana", 11, RGBColor(0x9F,0xB4,0xCB), False, True, FONT_B)]])

# ============================================================================
# SLIDE 2 — TUJUAN PEMBELAJARAN
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Pengantar", "Tujuan Pembelajaran", 2)
intro = ("Setelah mengikuti sesi ini, mahasiswa diharapkan mampu memahami hakikat seminar "
         "proposal dan menyusun proposal tesis yang berkualitas pada bidang manajemen pendidikan.")
txt(s, Inches(0.6), Inches(1.55), Inches(7.1), Inches(0.9),
    [[R(intro, 14, GREY, False, True, FONT_B)]], line_spacing=1.15)

objs = [
    ("Menjelaskan", "hakikat, tujuan, dan kedudukan seminar proposal dalam penelitian tesis."),
    ("Mengidentifikasi", "komponen utama proposal: pendahuluan, kajian pustaka, dan metodologi."),
    ("Merumuskan", "masalah penelitian yang tajam berbasis kesenjangan (research gap)."),
    ("Menyusun", "kerangka berpikir dan memilih metode penelitian yang tepat."),
    ("Mempresentasikan", "proposal secara meyakinkan dan menjunjung etika akademik."),
]
y = Inches(2.55)
for i, (h, b) in enumerate(objs):
    cy = y + Inches(0.92) * i
    card(s, Inches(0.6), cy, Inches(7.1), Inches(0.8), CARD, accent=TEAL)
    num = rect(s, Inches(0.78), cy + Inches(0.16), Inches(0.48), Inches(0.48), TEAL, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.78), cy + Inches(0.16), Inches(0.48), Inches(0.48),
        [[R(str(i+1), 16, WHITE, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.45), cy + Inches(0.06), Inches(6.1), Inches(0.7),
        [[R(h + " ", 14, NAVY, True, False, FONT_B), R(b, 13, TEXT, False, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
# gambar kanan
pic_cover(s, f"{IMG}/students.jpg", Inches(8.0), Inches(1.7), Inches(4.73), Inches(4.9))
rect(s, Inches(8.0), Inches(1.7), Inches(4.73), Inches(0.12), AMBER)
footer(s, 2)

# ============================================================================
# SLIDE 3 — APA ITU SEMINAR PROPOSAL?
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Konsep Dasar", "Apa Itu Seminar Proposal?", 3)
# definisi besar
card(s, Inches(0.6), Inches(1.6), Inches(7.4), Inches(1.65), NAVY)
txt(s, Inches(0.95), Inches(1.78), Inches(6.8), Inches(1.3),
    [[R("\u201CForum akademik untuk menguji kelayakan rencana penelitian", 16, WHITE, False, True, FONT_B)],
     [R("sebelum penelitian tesis dilaksanakan.\u201D", 16, WHITE, False, True, FONT_B)],
     [R("Mahasiswa mempertanggungjawabkan ide, fokus masalah, dan metode di hadapan dosen penguji.", 12.5, SOFT, False, False, FONT_B)]],
    line_spacing=1.12, space_after=4)
# 4 fungsi
funcs = [
    ("Menguji Kelayakan", "Memastikan topik penting, fokus, dan dapat diteliti.", TEAL),
    ("Memberi Masukan", "Penguji menajamkan masalah, teori, dan metode.", BLUE),
    ("Melatih Argumentasi", "Mahasiswa berlatih mempertahankan gagasan ilmiah.", AMBER),
    ("Pintu Gerbang", "Syarat sebelum melanjutkan ke penelitian penuh.", CORAL),
]
gx = Inches(0.6); gw = Inches(3.6); gh = Inches(1.55); gap = Inches(0.2)
for i, (h, b, c) in enumerate(funcs):
    col = i % 2; row = i // 2
    x = gx + (gw + gap) * col
    yy = Inches(3.55) + (gh + gap) * row
    card(s, x, yy, gw, gh, CARD, accent=c)
    txt(s, x + Inches(0.32), yy + Inches(0.2), gw - Inches(0.5), Inches(0.4),
        [[R(h, 15, NAVY, True, False, FONT_B)]])
    txt(s, x + Inches(0.32), yy + Inches(0.66), gw - Inches(0.5), Inches(0.8),
        [[R(b, 12.5, TEXT, False, False, FONT_B)]], line_spacing=1.05)
# gambar kanan
pic_cover(s, f"{IMG}/presentation.jpg", Inches(8.25), Inches(1.6), Inches(4.48), Inches(5.0))
rect(s, Inches(8.25), Inches(1.6), Inches(4.48), Inches(0.12), AMBER)
ref_note(s, "Creswell & Creswell (2018); Sugiyono (2019).")
footer(s, 3)


# ============================================================================
# SLIDE 4 — KEDUDUKAN PROPOSAL DALAM PENELITIAN TESIS (alur)
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Peta Jalan", "Kedudukan Proposal dalam Penelitian Tesis", 4)
txt(s, Inches(0.6), Inches(1.55), Inches(12), Inches(0.5),
    [[R("Proposal adalah ", 14, GREY, False, False, FONT_B),
      R("fondasi", 14, TEAL, True, False, FONT_B),
      R(" dari keseluruhan perjalanan tesis. Kualitas proposal menentukan kelancaran tahap berikutnya.",
        14, GREY, False, False, FONT_B)]], line_spacing=1.1)
steps = [
    ("1", "Ide & Topik", "Menemukan isu manajemen pendidikan yang penting", TEAL),
    ("2", "Proposal & Seminar", "Bab I–III diuji kelayakannya (FOKUS SESI INI)", AMBER),
    ("3", "Penelitian Lapangan", "Pengumpulan & analisis data", BLUE),
    ("4", "Penulisan Tesis", "Hasil, pembahasan, simpulan", CORAL),
    ("5", "Sidang/Ujian", "Pertahankan tesis & publikasi", NAVY),
]
n = len(steps)
total_w = Inches(12.1)
bw = Inches(2.18); arrow = Inches(0.28)
x0 = Inches(0.6); y0 = Inches(2.55)
for i, (num, h, b, c) in enumerate(steps):
    x = x0 + (bw + arrow) * i
    highlight = (i == 1)
    box = card(s, x, y0, bw, Inches(2.5), (SOFT if highlight else CARD), accent=c)
    circ = rect(s, x + bw/2 - Inches(0.42), y0 + Inches(0.28), Inches(0.84), Inches(0.84), c, shape=MSO_SHAPE.OVAL)
    txt(s, x + bw/2 - Inches(0.42), y0 + Inches(0.28), Inches(0.84), Inches(0.84),
        [[R(num, 26, WHITE, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.12), y0 + Inches(1.25), bw - Inches(0.24), Inches(0.55),
        [[R(h, 14, NAVY, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.16), y0 + Inches(1.72), bw - Inches(0.32), Inches(0.7),
        [[R(b, 11, TEXT, False, False, FONT_B)]], align=PP_ALIGN.CENTER, line_spacing=1.0)
    if i < n - 1:
        ar = rect(s, x + bw - Inches(0.02), y0 + Inches(1.0), arrow + Inches(0.06), Inches(0.5),
                  GREY, shape=MSO_SHAPE.CHEVRON)
# pesan kunci
card(s, Inches(0.6), Inches(5.45), Inches(12.1), Inches(1.05), NAVY)
txt(s, Inches(0.95), Inches(5.6), Inches(11.4), Inches(0.8),
    [[R("Pesan Kunci:  ", 14, AMBER, True, False, FONT_B),
      R("Proposal yang matang menghemat waktu, mengurangi revisi besar, dan memperbesar peluang tesis selesai tepat waktu.",
        14, WHITE, False, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
ref_note(s, "Gall, Gall, & Borg (2007); Emzir (2015).")
footer(s, 4)

# ============================================================================
# SLIDE 5 — STRUKTUR / ANATOMI PROPOSAL TESIS
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Kerangka", "Anatomi Proposal Tesis: Tiga Bab Utama", 5)
cols = [
    ("BAB I", "PENDAHULUAN", TEAL,
     ["Latar belakang masalah", "Identifikasi & batasan masalah", "Rumusan masalah", "Tujuan penelitian", "Manfaat penelitian"]),
    ("BAB II", "KAJIAN PUSTAKA", BLUE,
     ["Kajian teori & konsep", "Penelitian relevan terdahulu", "Kerangka berpikir", "Hipotesis (bila kuantitatif)"]),
    ("BAB III", "METODE PENELITIAN", CORAL,
     ["Pendekatan & jenis penelitian", "Populasi, sampel/subjek", "Teknik pengumpulan data", "Instrumen & keabsahan data", "Teknik analisis data"]),
]
cw = Inches(3.95); gap = Inches(0.18); x0 = Inches(0.6); y0 = Inches(1.7); ch = Inches(4.55)
for i, (tag, title, c, items) in enumerate(cols):
    x = x0 + (cw + gap) * i
    card(s, x, y0, cw, ch, CARD)
    head = rect(s, x, y0, cw, Inches(1.0), c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: head.adjustments[0] = 0.08
    except Exception: pass
    rect(s, x, y0 + Inches(0.5), cw, Inches(0.5), c)  # ratakan bawah header
    txt(s, x, y0 + Inches(0.12), cw, Inches(0.42),
        [[R(tag, 16, WHITE, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x, y0 + Inches(0.52), cw, Inches(0.42),
        [[R(title, 14, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, x + Inches(0.35), y0 + Inches(1.2), cw - Inches(0.6), ch - Inches(1.4),
            items, size=13, gap=9, marker_col=c)
txt(s, Inches(0.6), Inches(6.45), Inches(12), Inches(0.4),
    [[R("Catatan: ", 11.5, NAVY, True, True, FONT_B),
      R("susunan bab dapat menyesuaikan pedoman penulisan tesis tiap perguruan tinggi.", 11.5, GREY, False, True, FONT_B)]])
ref_note(s, "Sugiyono (2019); Fraenkel, Wallen, & Hyun (2019).")
footer(s, 5)

# ============================================================================
# SLIDE 6 — BAB I PENDAHULUAN
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Bab I", "Pendahuluan: Membangun Argumen Masalah", 6)
# kiri: piramida latar belakang
txt(s, Inches(0.6), Inches(1.5), Inches(6.7), Inches(0.45),
    [[R("Pola Latar Belakang (Piramida Terbalik)", 15, NAVY, True, False, FONT_B)]])
funnel = [
    ("Kondisi ideal / harapan (das Sollen)", TEAL, Inches(6.4)),
    ("Kondisi nyata / fakta lapangan (das Sein)", BLUE, Inches(5.3)),
    ("Kesenjangan & dampak masalah", AMBER, Inches(4.2)),
    ("Fokus masalah yang diteliti", CORAL, Inches(3.1)),
]
fy = Inches(2.05)
for i, (t, c, w) in enumerate(funnel):
    x = Inches(0.6) + (Inches(6.4) - w) / 2
    b = rect(s, x, fy + Inches(0.92)*i, w, Inches(0.78), c, shape=MSO_SHAPE.TRAPEZOID)
    b.rotation = 180
    txt(s, Inches(0.6), fy + Inches(0.92)*i, Inches(6.4), Inches(0.78),
        [[R(t, 12.5, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# kanan: komponen
card(s, Inches(7.5), Inches(1.5), Inches(5.25), Inches(5.0), CARD, accent=NAVY)
txt(s, Inches(7.85), Inches(1.68), Inches(4.6), Inches(0.45),
    [[R("Komponen Inti Bab I", 15, NAVY, True, False, FONT_B)]])
bullets(s, Inches(7.85), Inches(2.2), Inches(4.6), Inches(4.1), [
    ("Latar Belakang", "data, fakta, & teori yang menunjukkan masalah."),
    ("Identifikasi Masalah", "memetakan semua kemungkinan persoalan."),
    ("Batasan Masalah", "mempersempit agar fokus & terjangkau."),
    ("Rumusan Masalah", "pertanyaan penelitian yang dijawab."),
    ("Tujuan", "selaras dengan rumusan masalah."),
    ("Manfaat", "kontribusi teoretis & praktis."),
], size=13, gap=10, marker_col=TEAL)
ref_note(s, "Sugiyono (2019); Sekaran & Bougie (2016).")
footer(s, 6)

# ============================================================================
# SLIDE 7 — MERUMUSKAN MASALAH (FINER + research gap)
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Inti Proposal", "Merumuskan Masalah Penelitian yang Tajam", 7)
# kiri gambar + research gap
pic_cover(s, f"{IMG}/magnify.jpg", Inches(0.6), Inches(1.65), Inches(4.0), Inches(2.5))
rect(s, Inches(0.6), Inches(1.65), Inches(4.0), Inches(0.12), AMBER)
card(s, Inches(0.6), Inches(4.35), Inches(4.0), Inches(2.15), NAVY)
txt(s, Inches(0.85), Inches(4.5), Inches(3.5), Inches(0.5),
    [[R("Research Gap", 15, AMBER, True, False, FONT_B)]])
txt(s, Inches(0.85), Inches(4.95), Inches(3.5), Inches(1.5),
    [[R("Kesenjangan antara apa yang sudah diketahui (teori/riset) dan apa yang belum terjawab. Inilah ruang yang Anda isi.",
        12.5, WHITE, False, False, FONT_B)]], line_spacing=1.12)
# kanan FINER
txt(s, Inches(4.95), Inches(1.55), Inches(7.8), Inches(0.45),
    [[R("Kriteria Masalah yang Baik — ", 15, NAVY, True, False, FONT_B),
      R("FINER", 15, CORAL, True, False, FONT_B)]])
finer = [
    ("F", "Feasible", "Layak dikerjakan: waktu, dana, akses data memadai.", TEAL),
    ("I", "Interesting", "Menarik bagi peneliti & komunitas akademik.", BLUE),
    ("N", "Novel", "Baru / memberi kontribusi, bukan pengulangan.", AMBER),
    ("E", "Ethical", "Etis & tidak merugikan subjek penelitian.", CORAL),
    ("R", "Relevant", "Relevan bagi ilmu & praktik manajemen pendidikan.", NAVY),
]
fy = Inches(2.1)
for i, (L, name, desc, c) in enumerate(finer):
    yy = fy + Inches(0.82)*i
    card(s, Inches(4.95), yy, Inches(7.8), Inches(0.72), CARD, accent=c)
    sq = rect(s, Inches(5.1), yy + Inches(0.12), Inches(0.48), Inches(0.48), c)
    txt(s, Inches(5.1), yy + Inches(0.12), Inches(0.48), Inches(0.48),
        [[R(L, 20, WHITE, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(5.75), yy + Inches(0.04), Inches(6.9), Inches(0.64),
        [[R(name + " — ", 13.5, NAVY, True, False, FONT_B), R(desc, 12.5, TEXT, False, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE)
ref_note(s, "Hulley et al. (2013) — kriteria FINER; Creswell & Creswell (2018).")
footer(s, 7)


# ============================================================================
# SLIDE 8 — BAB II KAJIAN PUSTAKA
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Bab II", "Kajian Pustaka: Fondasi Teoretis", 8)
txt(s, Inches(0.6), Inches(1.5), Inches(7.2), Inches(0.85),
    [[R("Kajian pustaka bukan sekadar ringkasan buku, melainkan ", 14, GREY, False, False, FONT_B),
      R("argumentasi terstruktur", 14, TEAL, True, False, FONT_B),
      R(" yang memposisikan penelitian Anda di antara karya ilmiah lain.", 14, GREY, False, False, FONT_B)]],
    line_spacing=1.15)
items = [
    ("Kajian Teori", "Definisi & teori utama variabel/konsep penelitian."),
    ("Penelitian Relevan", "Temuan riset terdahulu + posisi kebaruan Anda."),
    ("Sintesis Kritis", "Membandingkan, bukan menumpuk kutipan."),
    ("Sumber Kredibel", "Jurnal terindeks, buku rujukan, data resmi (utamakan mutakhir)."),
]
for i, (h, b) in enumerate(items):
    yy = Inches(2.45) + Inches(1.0)*i
    card(s, Inches(0.6), yy, Inches(7.1), Inches(0.88), CARD, accent=BLUE)
    ic = rect(s, Inches(0.8), yy + Inches(0.19), Inches(0.5), Inches(0.5), BLUE, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.8), yy + Inches(0.19), Inches(0.5), Inches(0.5),
        [[R("\u2713", 18, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.5), yy + Inches(0.1), Inches(6.0), Inches(0.7),
        [[R(h + ": ", 14, NAVY, True, False, FONT_B), R(b, 13, TEXT, False, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
# kanan: gambar buku + tips piramida sumber
pic_cover(s, f"{IMG}/library.jpg", Inches(8.0), Inches(1.5), Inches(4.73), Inches(2.7))
rect(s, Inches(8.0), Inches(1.5), Inches(4.73), Inches(0.12), AMBER)
card(s, Inches(8.0), Inches(4.4), Inches(4.73), Inches(2.1), NAVY)
txt(s, Inches(8.3), Inches(4.55), Inches(4.2), Inches(0.4),
    [[R("Hindari!", 14, AMBER, True, False, FONT_B)]])
bullets(s, Inches(8.3), Inches(5.0), Inches(4.2), Inches(1.4), [
    "Sumber tidak ilmiah / blog tak jelas",
    "Kutipan tanpa analisis",
    "Teori usang yang tak relevan",
], size=12, gap=7, color=WHITE, marker_col=AMBER)
ref_note(s, "Bryman (2016); Creswell & Creswell (2018).")
footer(s, 8)

# ============================================================================
# SLIDE 9 — KERANGKA BERPIKIR & HIPOTESIS
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Bab II (lanjutan)", "Kerangka Berpikir & Hipotesis", 9)
# kiri penjelasan
card(s, Inches(0.6), Inches(1.6), Inches(5.3), Inches(2.25), CARD, accent=TEAL)
txt(s, Inches(0.9), Inches(1.78), Inches(4.7), Inches(0.45),
    [[R("Kerangka Berpikir", 16, NAVY, True, False, FONT_B)]])
txt(s, Inches(0.9), Inches(2.28), Inches(4.7), Inches(1.5),
    [[R("Alur logika yang menghubungkan teori, variabel, dan masalah menjadi satu kesatuan argumen. Sering disajikan sebagai bagan/diagram.",
        13, TEXT, False, False, FONT_B)]], line_spacing=1.15)
card(s, Inches(0.6), Inches(4.0), Inches(5.3), Inches(2.5), CARD, accent=CORAL)
txt(s, Inches(0.9), Inches(4.18), Inches(4.7), Inches(0.45),
    [[R("Hipotesis", 16, NAVY, True, False, FONT_B)]])
txt(s, Inches(0.9), Inches(4.66), Inches(4.7), Inches(1.7),
    [[R("Jawaban sementara atas rumusan masalah; ", 13, TEXT, False, False, FONT_B),
      R("wajib pada penelitian kuantitatif", 13, CORAL, True, False, FONT_B),
      R(", dapat tidak ada pada penelitian kualitatif. Harus dapat diuji secara empiris.", 13, TEXT, False, False, FONT_B)]],
    line_spacing=1.15)
# kanan: diagram contoh kerangka berpikir
txt(s, Inches(6.2), Inches(1.55), Inches(6.5), Inches(0.4),
    [[R("Contoh Bagan Kerangka Berpikir", 14, NAVY, True, False, FONT_B)]])
def node(x, y, w, h, text, c, fs=12):
    nb = rect(s, x, y, w, h, c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: nb.adjustments[0]=0.15
    except Exception: pass
    txt(s, x, y, w, h, [[R(text, fs, WHITE, True, False, FONT_B)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return nb
# X1, X2 -> Y
node(Inches(6.2), Inches(2.2), Inches(2.7), Inches(0.95), "Kepemimpinan Kepala Sekolah (X1)", TEAL, 12)
node(Inches(6.2), Inches(3.45), Inches(2.7), Inches(0.95), "Budaya Organisasi (X2)", BLUE, 12)
node(Inches(10.0), Inches(2.85), Inches(2.7), Inches(0.95), "Kinerja Guru (Y)", CORAL, 13)
# panah
rect(s, Inches(8.9), Inches(2.55), Inches(1.1), Inches(0.18), GREY, shape=MSO_SHAPE.RIGHT_ARROW)
rect(s, Inches(8.9), Inches(3.85), Inches(1.1), Inches(0.18), GREY, shape=MSO_SHAPE.RIGHT_ARROW)
card(s, Inches(6.2), Inches(4.85), Inches(6.55), Inches(1.5), SOFT)
txt(s, Inches(6.5), Inches(5.0), Inches(6.0), Inches(1.25),
    [[R("Hipotesis (H1): ", 13, NAVY, True, False, FONT_B),
      R("Terdapat pengaruh positif kepemimpinan kepala sekolah dan budaya organisasi terhadap kinerja guru.",
        13, TEXT, False, False, FONT_B)]], line_spacing=1.15, anchor=MSO_ANCHOR.MIDDLE)
ref_note(s, "Sugiyono (2019); Sekaran & Bougie (2016).")
footer(s, 9)


# ============================================================================
# SLIDE 10 — BAB III METODE PENELITIAN (3 pendekatan)
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Bab III", "Metode Penelitian: Memilih Pendekatan yang Tepat", 10)
txt(s, Inches(0.6), Inches(1.5), Inches(12), Inches(0.45),
    [[R("Metode mengikuti masalah — bukan sebaliknya. Pilih pendekatan sesuai pertanyaan penelitian Anda.",
        14, GREY, False, True, FONT_B)]])
approaches = [
    ("KUANTITATIF", TEAL, "Menguji teori / hubungan antar variabel dengan angka & statistik.",
     ["Pertanyaan: seberapa besar / adakah pengaruh", "Data: angka, kuesioner, skala", "Analisis: statistik (uji-t, regresi, SEM)", "Hasil: generalisasi"]),
    ("KUALITATIF", BLUE, "Memahami makna, proses, & fenomena secara mendalam.",
     ["Pertanyaan: bagaimana / mengapa", "Data: wawancara, observasi, dokumen", "Analisis: tematik / naratif", "Hasil: pemahaman kontekstual"]),
    ("CAMPURAN", CORAL, "Menggabungkan kekuatan kuantitatif & kualitatif (mixed methods).",
     ["Pertanyaan: kompleks & berlapis", "Data: angka + narasi", "Analisis: terintegrasi", "Hasil: gambaran menyeluruh"]),
]
cw = Inches(3.95); gap = Inches(0.18); x0 = Inches(0.6); y0 = Inches(2.1); ch = Inches(4.3)
for i, (name, c, sub, items) in enumerate(approaches):
    x = x0 + (cw + gap) * i
    card(s, x, y0, cw, ch, CARD)
    head = rect(s, x, y0, cw, Inches(0.75), c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: head.adjustments[0]=0.1
    except Exception: pass
    rect(s, x, y0 + Inches(0.4), cw, Inches(0.35), c)
    txt(s, x, y0, cw, Inches(0.75), [[R(name, 16, WHITE, True, False, FONT_H)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.3), y0 + Inches(0.9), cw - Inches(0.6), Inches(0.95),
        [[R(sub, 12.5, GREY, False, True, FONT_B)]], line_spacing=1.1)
    bullets(s, x + Inches(0.32), y0 + Inches(1.95), cw - Inches(0.55), ch - Inches(2.1),
            items, size=12, gap=8, marker_col=c)
ref_note(s, "Creswell & Creswell (2018); Sugiyono (2019); Miles, Huberman, & Saldaña (2014).")
footer(s, 10)

# ============================================================================
# SLIDE 11 — POPULASI, SAMPEL, INSTRUMEN, KEABSAHAN
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Bab III (lanjutan)", "Subjek, Instrumen, & Keabsahan Data", 11)
quad = [
    ("Populasi & Sampel", TEAL, "Seluruh subjek vs bagian yang mewakili. Tentukan teknik sampling (acak/purposif) & ukuran sampel yang memadai."),
    ("Teknik Pengumpulan Data", BLUE, "Angket, wawancara, observasi, dokumentasi — disesuaikan dengan jenis data yang dibutuhkan."),
    ("Instrumen Penelitian", AMBER, "Alat ukur (kuesioner/pedoman). Sertakan kisi-kisi & validasi ahli sebelum digunakan."),
    ("Validitas & Reliabilitas", CORAL, "Kuantitatif: uji validitas & reliabilitas. Kualitatif: triangulasi, member check, kredibilitas."),
]
cw = Inches(5.85); gap = Inches(0.35); x0 = Inches(0.6); y0 = Inches(1.7); chh = Inches(2.25)
for i, (h, c, b) in enumerate(quad):
    col = i % 2; row = i // 2
    x = x0 + (cw + gap) * col
    yy = y0 + (chh + Inches(0.25)) * row
    card(s, x, yy, cw, chh, CARD, accent=c)
    ic = rect(s, x + Inches(0.3), yy + Inches(0.3), Inches(0.7), Inches(0.7), c, shape=MSO_SHAPE.OVAL)
    txt(s, x + Inches(0.3), yy + Inches(0.3), Inches(0.7), Inches(0.7),
        [[R(str(i+1), 26, WHITE, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(1.2), yy + Inches(0.35), cw - Inches(1.5), Inches(0.6),
        [[R(h, 16, NAVY, True, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(0.32), yy + Inches(1.15), cw - Inches(0.65), Inches(1.0),
        [[R(b, 13, TEXT, False, False, FONT_B)]], line_spacing=1.15)
ref_note(s, "Sekaran & Bougie (2016); Fraenkel, Wallen, & Hyun (2019); Miles, Huberman, & Saldaña (2014).")
footer(s, 11)

# ============================================================================
# SLIDE 12 — KONTEKS MANAJEMEN PENDIDIKAN (contoh topik)
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Aplikasi Bidang", "Topik Relevan dalam Manajemen Pendidikan", 12)
txt(s, Inches(0.6), Inches(1.5), Inches(7.0), Inches(0.8),
    [[R("Pilih topik yang dekat dengan praktik kepemimpinan, kebijakan, dan mutu pendidikan — agar hasil tesis bermanfaat nyata.",
        14, GREY, False, True, FONT_B)]], line_spacing=1.15)
topics = [
    ("Kepemimpinan & SDM", "Kepemimpinan kepala sekolah, kinerja & profesionalisme guru.", TEAL),
    ("Mutu & Akreditasi", "Manajemen mutu terpadu (TQM), penjaminan mutu, akreditasi.", BLUE),
    ("Kebijakan & Pembiayaan", "Implementasi kebijakan, manajemen pembiayaan pendidikan.", AMBER),
    ("Budaya & Iklim", "Budaya organisasi sekolah, iklim kerja, motivasi.", CORAL),
    ("Inovasi & Digital", "Transformasi digital, manajemen perubahan, kurikulum.", TEAL),
    ("Layanan & Kepuasan", "Manajemen layanan, kepuasan stakeholder pendidikan.", BLUE),
]
cw = Inches(3.78); gap = Inches(0.18); x0 = Inches(0.6); y0 = Inches(2.45); chh = Inches(1.25)
for i, (h, b, c) in enumerate(topics):
    col = i % 2; row = i // 2
    # 2 kolom kiri (4 kartu) + memanfaatkan ruang; gunakan 2 kolom x 3 baris di kiri 7.7"
    x = x0 + (cw + gap) * col
    yy = y0 + (chh + Inches(0.15)) * row
    card(s, x, yy, cw, chh, CARD, accent=c)
    txt(s, x + Inches(0.3), yy + Inches(0.16), cw - Inches(0.5), Inches(0.4),
        [[R(h, 14, NAVY, True, False, FONT_B)]])
    txt(s, x + Inches(0.3), yy + Inches(0.6), cw - Inches(0.5), Inches(0.6),
        [[R(b, 11.5, TEXT, False, False, FONT_B)]], line_spacing=1.05)
# gambar kanan
pic_cover(s, f"{IMG}/teacher.jpg", Inches(8.35), Inches(2.45), Inches(4.38), Inches(4.05))
rect(s, Inches(8.35), Inches(2.45), Inches(4.38), Inches(0.12), AMBER)
ref_note(s, "Bush (2011); Hoy & Miskel (2013) sebagai rujukan teori manajemen pendidikan.")
footer(s, 12)

# ============================================================================
# SLIDE 13 — ETIKA PENELITIAN & ANTI-PLAGIARISME
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Integritas", "Etika Penelitian & Anti-Plagiarisme", 13)
left = [
    ("Informed Consent", "Subjek paham tujuan & menyetujui keikutsertaan secara sukarela."),
    ("Kerahasiaan", "Lindungi identitas & data pribadi responden."),
    ("Kejujuran Data", "Dilarang memanipulasi, memalsukan, atau merekayasa data."),
    ("Sitasi yang Benar", "Setiap kutipan & ide orang lain wajib dirujuk dengan tepat."),
]
txt(s, Inches(0.6), Inches(1.55), Inches(6.6), Inches(0.4),
    [[R("Prinsip Etika yang Wajib Dijaga", 15, NAVY, True, False, FONT_B)]])
for i, (h, b) in enumerate(left):
    yy = Inches(2.1) + Inches(1.05)*i
    card(s, Inches(0.6), yy, Inches(6.7), Inches(0.92), CARD, accent=TEAL)
    ic = rect(s, Inches(0.8), yy + Inches(0.21), Inches(0.5), Inches(0.5), TEAL, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.8), yy + Inches(0.21), Inches(0.5), Inches(0.5),
        [[R("\u2713", 18, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.5), yy + Inches(0.1), Inches(5.6), Inches(0.74),
        [[R(h + ": ", 14, NAVY, True, False, FONT_B), R(b, 12.5, TEXT, False, False, FONT_B)]],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
# kanan: kotak peringatan plagiarisme
card(s, Inches(7.6), Inches(1.55), Inches(5.15), Inches(4.95), NAVY)
warn = rect(s, Inches(7.85), Inches(1.8), Inches(0.9), Inches(0.9), AMBER, shape=MSO_SHAPE.OVAL)
txt(s, Inches(7.85), Inches(1.8), Inches(0.9), Inches(0.9),
    [[R("!", 36, NAVY, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(8.95), Inches(1.95), Inches(3.6), Inches(0.7),
    [[R("Plagiarisme = ", 16, AMBER, True, False, FONT_B), R("Pelanggaran Berat", 16, WHITE, True, False, FONT_B)]],
    anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
bullets(s, Inches(7.9), Inches(2.95), Inches(4.6), Inches(3.4), [
    "Parafrase dengan kalimat sendiri + tetap merujuk sumber",
    "Gunakan manajer referensi (Mendeley/Zotero)",
    "Konsisten memakai satu gaya sitasi (APA 7th)",
    "Cek kemiripan (Turnitin) sebelum seminar",
    "Sanksi: pembatalan, hukum, & reputasi akademik",
], size=12.5, gap=11, color=WHITE, marker_col=AMBER)
ref_note(s, "Resnik (2011); pedoman integritas akademik & gaya sitasi APA (2020).")
footer(s, 13)


# ============================================================================
# SLIDE 14 — TEKNIK PRESENTASI SEMINAR YANG EFEKTIF
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Keterampilan", "Teknik Presentasi Seminar yang Efektif", 14)
# kiri gambar
pic_cover(s, f"{IMG}/study.jpg", Inches(0.6), Inches(1.6), Inches(3.7), Inches(4.9))
rect(s, Inches(0.6), Inches(1.6), Inches(3.7), Inches(0.12), AMBER)
# kanan: tahapan
phases = [
    ("Sebelum", TEAL, ["Kuasai isi proposal, bukan hafalan slide", "Siapkan slide ringkas (1 ide/slide)", "Latihan & kelola waktu (mis. 15 menit)"]),
    ("Saat Presentasi", BLUE, ["Buka dengan masalah & urgensi", "Tampilkan alur logis Bab I–III", "Bahasa tubuh & kontak mata percaya diri"]),
    ("Sesi Tanya Jawab", CORAL, ["Dengarkan penuh, catat pertanyaan", "Jawab jujur; akui bila perlu kaji ulang", "Sikap terbuka pada masukan penguji"]),
]
x = Inches(4.6); w = Inches(8.15); y0 = Inches(1.6); hh = Inches(1.55)
for i, (h, c, items) in enumerate(phases):
    yy = y0 + (hh + Inches(0.12))*i
    card(s, x, yy, w, hh, CARD, accent=c)
    tag = rect(s, x, yy, Inches(2.0), hh, c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: tag.adjustments[0]=0.08
    except Exception: pass
    rect(s, x + Inches(1.0), yy, Inches(1.0), hh, c)
    txt(s, x, yy, Inches(2.0), hh, [[R(h, 15, WHITE, True, False, FONT_H)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, x + Inches(2.25), yy + Inches(0.16), w - Inches(2.5), hh - Inches(0.3),
            items, size=12.5, gap=5, marker_col=c)
ref_note(s, "Bryman (2016); panduan presentasi ilmiah pascasarjana.")
footer(s, 14)

# ============================================================================
# SLIDE 15 — KESALAHAN UMUM & TIPS SUKSES
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Refleksi", "Kesalahan Umum vs Tips Sukses", 15)
# kolom kiri: kesalahan
card(s, Inches(0.6), Inches(1.6), Inches(5.9), Inches(4.9), CARD)
rect(s, Inches(0.6), Inches(1.6), Inches(5.9), Inches(0.85), CORAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(0.6), Inches(2.1), Inches(5.9), Inches(0.35), CORAL)
txt(s, Inches(0.6), Inches(1.6), Inches(5.9), Inches(0.85),
    [[R("\u2715   Kesalahan yang Sering Terjadi", 16, WHITE, True, False, FONT_B)]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(0.95), Inches(2.7), Inches(5.25), Inches(3.6), [
    "Latar belakang tidak menunjukkan masalah nyata",
    "Rumusan masalah & tujuan tidak sinkron",
    "Kajian pustaka hanya menumpuk kutipan",
    "Metode tidak sesuai dengan masalah",
    "Topik terlalu luas / tidak terjangkau",
    "Sumber rujukan usang & tidak kredibel",
], size=13.5, gap=12, marker_col=CORAL)
# kolom kanan: tips
card(s, Inches(6.85), Inches(1.6), Inches(5.9), Inches(4.9), CARD)
rect(s, Inches(6.85), Inches(1.6), Inches(5.9), Inches(0.85), TEAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(6.85), Inches(2.1), Inches(5.9), Inches(0.35), TEAL)
txt(s, Inches(6.85), Inches(1.6), Inches(5.9), Inches(0.85),
    [[R("\u2713   Kunci Proposal yang Sukses", 16, WHITE, True, False, FONT_B)]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(7.2), Inches(2.7), Inches(5.25), Inches(3.6), [
    "Mulai dari masalah yang Anda kuasai & pedulikan",
    "Jaga benang merah: masalah → tujuan → metode",
    "Sintesis kritis, bukan sekadar mengutip",
    "Pilih metode yang realistis & sesuai data",
    "Persempit fokus agar tuntas & mendalam",
    "Konsultasi rutin dengan dosen pembimbing",
], size=13.5, gap=12, marker_col=TEAL)
footer(s, 15)

# ============================================================================
# SLIDE 16 — RUBRIK PENILAIAN SEMINAR PROPOSAL
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Asesmen", "Aspek Penilaian Seminar Proposal", 16)
txt(s, Inches(0.6), Inches(1.5), Inches(12), Inches(0.45),
    [[R("Gunakan aspek-aspek berikut sebagai daftar periksa (checklist) kesiapan Anda.",
        14, GREY, False, True, FONT_B)]])
rows = [
    ("Latar Belakang & Masalah", "Kejelasan masalah, urgensi, & kesenjangan riset", "20%", TEAL),
    ("Kajian Pustaka & Teori", "Kecukupan, kemutakhiran, & sintesis kritis sumber", "20%", BLUE),
    ("Metodologi Penelitian", "Ketepatan metode, instrumen, & keabsahan data", "25%", CORAL),
    ("Sistematika & Tata Tulis", "Logika, konsistensi, sitasi, & bahasa ilmiah", "15%", AMBER),
    ("Presentasi & Penguasaan", "Penyampaian, argumentasi, & respons tanya jawab", "20%", NAVY),
]
# header tabel
hx = Inches(0.6); hy = Inches(2.15); tw = Inches(12.15)
hb = rect(s, hx, hy, tw, Inches(0.6), NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
try: hb.adjustments[0]=0.12
except Exception: pass
txt(s, hx + Inches(0.3), hy, Inches(4.2), Inches(0.6), [[R("ASPEK", 13, WHITE, True, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, hx + Inches(4.6), hy, Inches(6.0), Inches(0.6), [[R("INDIKATOR", 13, WHITE, True, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, hx + Inches(10.9), hy, Inches(1.1), Inches(0.6), [[R("BOBOT", 13, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
ry = hy + Inches(0.72)
for i, (asp, ind, bob, c) in enumerate(rows):
    yy = ry + Inches(0.78)*i
    rowbg = CARD if i % 2 == 0 else SOFT
    rb = rect(s, hx, yy, tw, Inches(0.68), rowbg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: rb.adjustments[0]=0.1
    except Exception: pass
    rect(s, hx, yy, Inches(0.1), Inches(0.68), c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s, hx + Inches(0.3), yy, Inches(4.2), Inches(0.68), [[R(asp, 13, NAVY, True, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, hx + Inches(4.6), yy, Inches(6.1), Inches(0.68), [[R(ind, 12, TEXT, False, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    bb = rect(s, hx + Inches(10.95), yy + Inches(0.12), Inches(1.0), Inches(0.44), c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    try: bb.adjustments[0]=0.3
    except Exception: pass
    txt(s, hx + Inches(10.95), yy + Inches(0.12), Inches(1.0), Inches(0.44), [[R(bob, 13, WHITE, True, False, FONT_B)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
ref_note(s, "Disarikan dari pedoman penilaian seminar proposal pascasarjana & Gall, Gall, & Borg (2007).")
footer(s, 16)


# ============================================================================
# SLIDE 17 — RANGKUMAN
# ============================================================================
s = slide()
bg(s, NAVY)
# aksen visual
pic_cover(s, f"{IMG}/success.jpg", Inches(8.4), 0, Inches(4.933), SH)
rect(s, Inches(8.4), 0, Inches(0.12), SH, AMBER)
rect(s, 0, 0, Inches(8.5), SH, NAVY)
rect(s, Inches(0.6), Inches(0.7), Inches(1.5), Inches(0.10), AMBER)
txt(s, Inches(0.6), Inches(0.95), Inches(7.4), Inches(0.9),
    [[R("Rangkuman", 34, WHITE, True, False, FONT_H)]])
txt(s, Inches(0.6), Inches(1.75), Inches(7.5), Inches(0.5),
    [[R("Lima hal yang harus Anda bawa pulang dari sesi ini:", 14, SOFT, False, True, FONT_B)]])
points = [
    "Seminar proposal menguji kelayakan rencana penelitian tesis Anda.",
    "Jaga benang merah: masalah → rumusan → tujuan → metode.",
    "Masalah yang baik memenuhi kriteria FINER & berbasis research gap.",
    "Pilih metode sesuai pertanyaan; perhatikan keabsahan data.",
    "Junjung etika akademik & presentasikan dengan percaya diri.",
]
for i, p in enumerate(points):
    yy = Inches(2.45) + Inches(0.82)*i
    num = rect(s, Inches(0.6), yy, Inches(0.55), Inches(0.55), AMBER, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(0.6), yy, Inches(0.55), Inches(0.55),
        [[R(str(i+1), 18, NAVY, True, False, FONT_H)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.35), yy - Inches(0.02), Inches(6.7), Inches(0.7),
        [[R(p, 14.5, WHITE, False, False, FONT_B)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
txt(s, Inches(0.6), Inches(6.7), Inches(7.5), Inches(0.5),
    [[R("\u201CProposal yang matang adalah separuh dari tesis yang selesai.\u201D", 13, AMBER, False, True, FONT_B)]])
footer(s, 17)

# ============================================================================
# SLIDE 18 — REFERENSI
# ============================================================================
s = slide()
bg(s, LIGHT)
header(s, "Sumber Rujukan", "Referensi Pendukung", 18)
txt(s, Inches(0.6), Inches(1.5), Inches(12), Inches(0.4),
    [[R("Disusun mengikuti gaya American Psychological Association (APA 7th edition).",
        12.5, GREY, False, True, FONT_B)]])
refs = [
    "Bryman, A. (2016). Social research methods (5th ed.). Oxford University Press.",
    "Bush, T. (2011). Theories of educational leadership and management (4th ed.). SAGE.",
    "Creswell, J. W., & Creswell, J. D. (2018). Research design: Qualitative, quantitative, and mixed methods approaches (5th ed.). SAGE.",
    "Emzir. (2015). Metodologi penelitian pendidikan: Kuantitatif dan kualitatif. Rajawali Pers.",
    "Fraenkel, J. R., Wallen, N. E., & Hyun, H. H. (2019). How to design and evaluate research in education (10th ed.). McGraw-Hill.",
    "Gall, M. D., Gall, J. P., & Borg, W. R. (2007). Educational research: An introduction (8th ed.). Pearson.",
    "Hoy, W. K., & Miskel, C. G. (2013). Educational administration: Theory, research, and practice (9th ed.). McGraw-Hill.",
    "Hulley, S. B., Cummings, S. R., Browner, W. S., Grady, D. G., & Newman, T. B. (2013). Designing clinical research (4th ed.). Lippincott Williams & Wilkins.",
    "Miles, M. B., Huberman, A. M., & Saldaña, J. (2014). Qualitative data analysis: A methods sourcebook (3rd ed.). SAGE.",
    "Sekaran, U., & Bougie, R. (2016). Research methods for business: A skill-building approach (7th ed.). Wiley.",
    "Sugiyono. (2019). Metode penelitian kuantitatif, kualitatif, dan R&D. Alfabeta.",
]
# dua kolom
col1 = refs[:6]; col2 = refs[6:]
def ref_col(x, w, items, start):
    tb = s.shapes.add_textbox(x, Inches(2.05), w, Inches(4.5))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10); p.line_spacing = 1.05
        rn = p.add_run(); rn.text = f"{start+i}. "
        rn.font.size = Pt(11.5); rn.font.bold = True; rn.font.color.rgb = TEAL; rn.font.name = FONT_B
        rt = p.add_run(); rt.text = it
        rt.font.size = Pt(11.5); rt.font.color.rgb = TEXT; rt.font.name = FONT_B
card(s, Inches(0.55), Inches(1.9), Inches(6.05), Inches(4.5), CARD, accent=TEAL)
card(s, Inches(6.75), Inches(1.9), Inches(6.0), Inches(4.5), CARD, accent=BLUE)
ref_col(Inches(0.85), Inches(5.5), col1, 1)
ref_col(Inches(7.05), Inches(5.5), col2, 7)
footer(s, 18)

# ----------------------------------------------------------------------------
out = "Bahan_Ajar_Seminar_Proposal_Tesis.pptx"
prs.save(out)
print("Tersimpan:", out, "| Jumlah slide:", len(prs.slides._sldIdLst))
