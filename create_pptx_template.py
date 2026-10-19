#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пересборка презентации в визуале шаблона FOSTER_Club_презентация.pptx
- Сохраняет всю финансовую информацию из исходной задачи
- Визуал 1:1 как в шаблоне: цвета, шрифты, сетка, карточки, футеры
- Использует подлинные фото Foster Club из галереи Яндекс Карт (извлечены из шаблона)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import os

# Template colors (точно из шаблона)
GOLD = RGBColor(0xC9, 0xA2, 0x5E)      # C9A25E
GOLD_LIGHT = RGBColor(0xE5, 0xC8, 0x88) # E5C888
BEIGE = RGBColor(0xF5, 0xF1, 0xE8)     # F5F1E8
BEIGE2 = RGBColor(0xDD, 0xD8, 0xCC)    # DDD8CC
BEIGE3 = RGBColor(0xE2, 0xDE, 0xD4)    # E2DED4
MUTED = RGBColor(0x7E, 0x7A, 0x72)     # 7E7A72
DARK = RGBColor(0x14, 0x14, 0x19)      # 141419
DARKER = RGBColor(0x09, 0x09, 0x0B)    # 09090B
BURGUNDY = RGBColor(0x7A, 0x1F, 0x2B)   # 7A1F2B
WHITE = RGBColor(0xFF,0xFF,0xFF)
BLACK = RGBColor(0x00,0x00,0x00)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
# blank layout
BLANK = prs.slide_layouts[6]

# paths to extracted images
IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg",
    "logo": "/tmp/foster_images/slide01_img2_Picture_3.png",
    "about": "/tmp/foster_images/slide02_img3_Picture_5.jpg",
    "concept_bg": "/tmp/foster_images/slide03_img4_Picture_1.jpg",
    "context": "/tmp/foster_images/slide04_img5_Picture_5.jpg",
    "entrance": "/tmp/foster_images/slide05_img6_Picture_5.jpg",
    "hall1": "/tmp/foster_images/slide06_img7_Picture_5.jpg",
    "hall2": "/tmp/foster_images/slide06_img8_Picture_6.jpg",
    "dj_bar": "/tmp/foster_images/slide07_img9_Picture_5.jpg",
    "vip": "/tmp/foster_images/slide07_img10_Picture_6.jpg",
    "bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "loft1": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "loft2": "/tmp/foster_images/slide09_img13_Picture_6.jpg",
    "timeline": "/tmp/foster_images/slide11_img18_Picture_5.jpg",
    "audience": "/tmp/foster_images/slide12_img19_Picture_5.jpg",
    "final": "/tmp/foster_images/slide13_img20_Picture_1.jpg",
}

def add_bg_image(slide, img_path, brightness=0.55):
    # full slide background image
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        # dark overlay for text readability
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        shape.fill.solid()
        # dark overlay with transparency via RGB with alpha simulation (use solid dark with 45% opacity effect by choosing dark)
        shape.fill.fore_color.rgb = RGBColor(0x0A,0x0A,0x0F)
        shape.line.fill.background()
        # python-pptx doesn't support transparency easily, we simulate by using dark fill with low brightness image already dark
        # We'll set transparency by not adding overlay for some slides, but for cover we add 35% dark
        return shape
    return None

def set_text_box(slide, left, top, width, height, text, font_name="Arial", font_size=10, color=BEIGE, bold=False, italic=False, alignment=PP_ALIGN.LEFT, line_spacing=1.0, letter_spacing=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.italic = italic
    p.alignment = alignment
    p.line_spacing = line_spacing
    return txBox

def add_paragraph(text_frame, text, font_name="Arial", font_size=10, color=BEIGE, bold=False, italic=False, alignment=PP_ALIGN.LEFT):
    p = text_frame.add_paragraph()
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.italic = italic
    p.alignment = alignment
    p.line_spacing = 1.0
    p.space_before = Pt(2)
    p.space_after = Pt(2)
    return p

def add_gold_line(slide, left=Inches(0.57), top=Inches(1.42), width=Inches(0.70), height=Inches(0.05)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    return shape

def add_card(slide, left, top, width, height, fill=DARK, radius=True, header=None, body=None, header_color=GOLD_LIGHT, body_color=BEIGE3):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if radius:
        shape.adjustments[0] = 0.08
    if header:
        tb = slide.shapes.add_textbox(left + Inches(0.22), top + Inches(0.16), width - Inches(0.44), Inches(0.30))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = header
        p.font.name = "Arial"
        p.font.size = Pt(7.5)
        p.font.color.rgb = header_color
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT
    if body:
        tb = slide.shapes.add_textbox(left + Inches(0.22), top + Inches(0.52), width - Inches(0.44), height - Inches(0.70))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = body
        p.font.name = "Arial"
        p.font.size = Pt(7)
        p.font.color.rgb = body_color
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.15
    return shape

def add_footer(slide, number="01 / 10"):
    # left footer
    set_text_box(slide, Inches(0.55), Inches(7.12), Inches(2), Inches(0.25), "Фото: галерея Foster Club — Яндекс Карты", font_size=7.5, color=MUTED, font_name="Arial")
    # right footer
    set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), f"FOSTER • {number.split('/')[0].strip()}", font_size=7.5, color=MUTED, bold=True, alignment=PP_ALIGN.RIGHT, font_name="Arial")

def add_header_block(slide, spaced_title, main_title, slide_no="01 / 10"):
    # spaced caps
    set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), spaced_title, font_size=10, color=GOLD, bold=True, font_name="Arial")
    # main title
    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(7.0), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = main_title
    p.font.name = "Arial"
    p.font.size = Pt(24)
    p.font.color.rgb = BEIGE
    p.font.bold = True
    p.line_spacing = 0.95
    add_gold_line(slide)
    # slide number
    set_text_box(slide, Inches(11.70), Inches(0.32), Inches(1.10), Inches(0.30), slide_no, font_size=9, color=MUTED, alignment=PP_ALIGN.RIGHT, font_name="Arial")

# ================= SLIDE 1 - COVER =================
slide = prs.slides.add_slide(BLANK)
# Background image full
add_bg_image(slide, IMG["cover"])
# dark overlay is added inside; adjust to 45% dark by changing overlay color opacity simulation
# Top beige rectangle (like original)
rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(0.45), Inches(2.45), Inches(0.98))
rect.fill.solid()
rect.fill.fore_color.rgb = BEIGE
rect.line.fill.background()
rect.adjustments[0]=0.08
# Logo inside rectangle (Foster logo)
if os.path.exists(IMG["logo"]):
    slide.shapes.add_picture(IMG["logo"], Inches(0.70), Inches(0.60), width=Inches(2.15), height=Inches(0.68))
# Top right compact info
set_text_box(slide, Inches(9.5), Inches(0.45), Inches(3.30), Inches(0.30), "Mriya Resort & SPA  •  с. Оползневое, Ялта", font_size=7, color=WHITE, alignment=PP_ALIGN.RIGHT, font_name="Arial")
set_text_box(slide, Inches(9.5), Inches(0.68), Inches(3.30), Inches(0.25), "5 декабря 2026  •  Суббота  •  300–350 гостей", font_size=7, color=GOLD, alignment=PP_ALIGN.RIGHT, font_name="Arial", bold=True)

# Main title large - copy original's two-line concept but with our headliners
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(8.5), Inches(1.6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "GLEB FILIPCHENKOW"
p.font.name = "Arial"
p.font.size = Pt(34)
p.font.color.rgb = WHITE
p.font.bold = True
p.line_spacing = 0.9
p = tf.add_paragraph()
p.text = "& NOEMY ABRANTES"
p.font.name = "Arial"
p.font.size = Pt(34)
p.font.color.rgb = GOLD
p.font.bold = True
p.line_spacing = 0.9
p.space_before = Pt(4)

add_gold_line(slide, Inches(0.57), Inches(3.55), Inches(0.85), Inches(0.06))

# subtitle under gold line
set_text_box(slide, Inches(0.55), Inches(3.75), Inches(8.2), Inches(0.45), "Первый привоз в Крым  •  Хит «Mallorca» (лето 2025)  •  Mallorca • Garota • HOJE", font_size=9, color=BEIGE2, font_name="Arial")

# bottom left description card semi-transparent dark
card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(4.40), Inches(6.0), Inches(1.55))
card.fill.solid()
card.fill.fore_color.rgb = DARK
card.line.fill.background()
card.adjustments[0]=0.06
set_text_box(slide, Inches(0.75), Inches(4.55), Inches(5.6), Inches(0.25), "◆  Ночь, которую не нужно искать за пределами курорта", font_size=8, color=GOLD, bold=True, font_name="Arial")
tb = slide.shapes.add_textbox(Inches(0.75), Inches(4.85), Inches(5.6), Inches(1.0))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Привоз планировался на лето — сорвался из-за топливного кризиса. Публика ждёт. 5 декабря — единственное свободное окно в графике артистов. Имя продаёт само событие."
p.font.name="Arial"
p.font.size=Pt(8)
p.font.color.rgb=BEIGE3
p.line_spacing=1.25

# 3 mini stats at bottom like original's cards but dark
for i, (num, lbl) in enumerate([("300–350","ГОстей / до 60 мест зал"), ("63%","Рентабельность Foster"), ("483 549 ₽","Ваша чистая прибыль")]):
    left = Inches(0.55 + i*2.05)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(6.20), Inches(1.90), Inches(0.75))
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK
    shape.line.fill.background()
    shape.adjustments[0]=0.08
    set_text_box(slide, left+Inches(0.12), Inches(6.28), Inches(1.66), Inches(0.25), num, font_size=10, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER, font_name="Arial")
    set_text_box(slide, left+Inches(0.12), Inches(6.55), Inches(1.66), Inches(0.20), lbl, font_size=6, color=GOLD_LIGHT, alignment=PP_ALIGN.CENTER, font_name="Arial", bold=True)

set_text_box(slide, Inches(0.55), Inches(7.12), Inches(5), Inches(0.25), "Пересборка презентации на подлинных фотографиях клуба из галереи Яндекс Карт", font_size=7.5, color=MUTED, font_name="Arial")
set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), "01 / 10", font_size=9, color=WHITE, bold=False, alignment=PP_ALIGN.RIGHT, font_name="Arial")

# ================= SLIDE 2 - ПОЗИЦИОНИРОВАНИЕ / О ПРОЕКТЕ =================
slide = prs.slides.add_slide(BLANK)
# white background (default) - ensure clean
add_header_block(slide, "П О З И Ц И О Н И Р О В А Н И Е", "Камерная вечеринка внутри курорта", slide_no="02 / 10")
# image right
if os.path.exists(IMG["about"]):
    slide.shapes.add_picture(IMG["about"], Inches(7.00), Inches(1.75), width=Inches(5.78), height=Inches(3.50))
# bullets left
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(6.15), Inches(3.40))
tf = tb.text_frame
tf.word_wrap = True
bullets = [
    ("Формат:", "привоз хедлайнеров 300–350 гостей, night & karaoke-шоу в составе Mriya Resort & SPA — вечерняя «гостиная» курорта."),
    ("Почему сейчас:", "летний привоз сорвался из-за топливного кризиса; публика ждёт; 5 декабря — единственное окно."),
    ("Масштаб:", "камерный зал Foster — до 60 мест (по данным 2ГИС) + танцпол; 300–350 гостей — полный SOLD OUT."),
    ("Атмосфера:", "роскошный лофт встречает Голливуд — гость курорта становится артистом."),
]
for label, body in bullets:
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    # diamond
    run = p.add_run()
    run.text = "◆  "
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = GOLD
    run.font.bold = True
    run = p.add_run()
    run.text = label + " "
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = BEIGE
    run.font.bold = True
    run = p.add_run()
    run.text = body
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x33,0x33,0x3A) if False else RGBColor(0x3A,0x3A,0x42) # dark for white bg
    run.font.bold = False
# But bullets need dark text on white bg, so override to dark grey
# Re-color: we used dark grey 3A3A42 - actually use DARK for white bg slide
# Let's keep as above but adjust: on white bg, body should be dark (2D2D32)
# Fix: re-assign
for para in tf.paragraphs:
    for run in para.runs:
        if run.text and "—" in run.text or "привоз" in run.text or "летний" in run.text:
            run.font.color.rgb = RGBColor(0x2A,0x2A,0x30)

# 3 cards bottom (replicate original's 3 rounded dark cards)
add_card(slide, Inches(0.55), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ФОРМАТ", body="Ночной привоз хедлайнеров •\nDJ-бар • танцпол • караоке •\nVIP-зал • шоу. 300–350 гостей.")
add_card(slide, Inches(4.70), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="РОЛИ СТОРОН", body="FOSTER — площадка и бар\nКАКТУСА — организация и продажи\nMRIYA — проживание (без скидки)")
add_card(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(1.35), fill=DARK, header="ДАТА И ЦЕНА", body="Суббота, 5 декабря 2026\nБилеты 3 000 ₽ → 3 500 ₽ → 4 000 ₽\nСтарт продаж 01.11")
add_footer(slide, "02 / 10")

# ================= SLIDE 3 - КОНЦЕПЦИЯ В ОДНОЙ ФРАЗЕ (переосмыслена под прибыль) =================
slide = prs.slides.add_slide(BLANK)
# full bg image
add_bg_image(slide, IMG["concept_bg"])
# remove dark overlay to show image well, then add dark cards as in original slide 03
# But original slide 03 had dark cards with transparency; we add overlay shape with 70% opacity simulation (use semi-dark)
# Top header
set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), "К О Н Ц Е П Ц И Я   С Д Е Л К И", font_size=10, color=GOLD, bold=True, font_name="Arial")
tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(10), Inches(0.95))
tf = tb.text_frame
tf.word_wrap=True
p=tf.paragraphs[0]
p.text="«Ваша прибыль — 483 549 ₽ за вечер»"
p.font.name="Arial"
p.font.size=Pt(24)
p.font.color.rgb=BEIGE
p.font.bold=True
p.line_spacing=0.95
add_gold_line(slide)
set_text_box(slide, Inches(11.70), Inches(0.32), Inches(1.10), Inches(0.30), "03 / 10", font_size=9, color=MUTED, alignment=PP_ALIGN.RIGHT, font_name="Arial")
# 3 cards dark as original: 1•ЛОФТ etc -> repurpose
add_card(slide, Inches(0.55), Inches(1.85), Inches(3.95), Inches(1.35), fill=DARKER, header="1 • ПРИБЫЛЬ", body="768 025 ₽ выручка — 284 476 ₽\nрасходы = 483 549 ₽ чистыми.\nРентабельность ~63% за ночь.")
add_card(slide, Inches(4.70), Inches(1.85), Inches(3.95), Inches(1.35), fill=DARKER, header="2 • ГАРАНТИЯ", body="330 000 ₽ живыми деньгами\n(180К коктейли + 50К столы +\n70К VIP) — фикс в договоре.")
add_card(slide, Inches(8.85), Inches(1.85), Inches(3.93), Inches(1.35), fill=DARKER, header="3 • НУЛЬ РИСКА", body="0 ₽ за аренду, артистов,\nоборудование — только\nоперационка клуба.")
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(9.5), Inches(0.25), "Дизайн-концепция: «роскошный лофт» — студия Елены Крыловой  •  Фото: красный лаунж Foster Club, Яндекс Карты", font_size=7.5, color=MUTED, font_name="Arial")
set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), "FOSTER • 03", font_size=7.5, color=MUTED, bold=True, alignment=PP_ALIGN.RIGHT, font_name="Arial")


# ================= SLIDE 4 - ЗАТРАТЫ FOSTER =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "З А Т Р А Т Ы   К Л У Б А", "Ваши затраты минимальны", slide_no="04 / 10")
# left bullets about costs? Instead create table-like cards left, image right context image
# Image right: context terrace
if os.path.exists(IMG["context"]):
    slide.shapes.add_picture(IMG["context"], Inches(7.00), Inches(1.75), width=Inches(5.85), height=Inches(3.50))
# Top left bullet summary
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(6.15), Inches(2.0))
tf = tb.text_frame
tf.word_wrap=True
bullets2 = [
    ("Вы предоставляете", "площадку и оборудование бесплатно — только операционка."),
    ("Все прямые расходы", "покрываются баром и кухней, фактически — нулевые инвестиции."),
    ("Сравнение масштаба:", "ваши 284 476 ₽ против 1 120 040 ₽ у Кактусы — в 4 раза меньше."),
]
for label, body in bullets2:
    p = tf.add_paragraph()
    p.space_before=Pt(5)
    r=p.add_run()
    r.text="◆  "
    r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=GOLD; r.font.bold=True
    r=p.add_run()
    r.text=label+" "
    r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run()
    r.text=body
    r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=RGBColor(0x2A,0x2A,0x30)

# bottom 4 small cards for cost breakdown (instead of 3, make 4)
costs = [
    ("ПЕРСОНАЛ", "165 470 ₽\nофицианты, бармены,\nповара, менеджмент", DARK),
    ("БАР", "92 006 ₽\nфудкост 25% от\n368 025 ₽", DARK),
    ("КУХНЯ", "27 000 ₽\nфудкост 27% от\n100 000 ₽", DARK),
    ("ИТОГО", "284 476 ₽\nвсе расходы клуба\nпокрываются", BURGUNDY),
]
for i, (h,b,fill) in enumerate(costs):
    left = Inches(0.55 + i*3.25)
    add_card(slide, left, Inches(5.45), Inches(3.00), Inches(1.35), fill=fill, header=h, body=b)
    # make header larger for ITGO
    # if fill is burgundy, header is white?
    # adjust

set_text_box(slide, Inches(0.55), Inches(7.12), Inches(6), Inches(0.25), "Терраса Mriya Resort & SPA — вид на Чёрное море. Концепцию в форме лотоса создал Норман Фостер.", font_size=7.5, color=MUTED, font_name="Arial")
add_footer(slide, "04 / 10")

# ================= SLIDE 5 - ДОХОД FOSTER =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Ч Т О   З А Р А Б О Т А Е Т   К Л У Б", "Чистая прибыль — 483 549 ₽", slide_no="05 / 10")
# left text with highlights
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(6.15), Inches(3.50))
tf = tb.text_frame
tf.word_wrap=True
# Add paragraph with colored emphasis for guaranteed
rows_rev = [
    ("Бар", "368 025 ₽", "валовая выручка"),
    ("Кухня", "100 000 ₽", "валовая выручка"),
    ("◆ Welcome-коктейли", "180 000 ₽", "300 × 600 ₽ компенсация → вам"),
    ("◆ Депозиты столов", "50 000 ₽", "4×5 000 + 3×10 000 → вам"),
    ("◆ VIP №1 + №2", "70 000 ₽", "40 000 + 30 000 → вам"),
    ("Итого выручка", "768 025 ₽", "+ кальяны от Макао"),
]
for name, price, note in rows_rev:
    p = tf.add_paragraph()
    p.space_before=Pt(4)
    is_highlight = "◆" in name or "Итого" in name
    r = p.add_run()
    r.text = name + "  "
    r.font.name="Arial"; r.font.size=Pt(8.5); r.font.color.rgb=GOLD if is_highlight else DARK; r.font.bold=True
    r = p.add_run()
    r.text = price + "  "
    r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=DARK; r.font.bold=True
    r = p.add_run()
    r.text = "— " + note
    r.font.name="Arial"; r.font.size=Pt(7); r.font.color.rgb=MUTED

# image right - entrance brick
if os.path.exists(IMG["entrance"]):
    slide.shapes.add_picture(IMG["entrance"], Inches(7.00), Inches(1.75), width=Inches(5.78), height=Inches(3.50))

# bottom 3 cards with profit visualization like original's 3 cards but with numbers
add_card(slide, Inches(0.55), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ВАША ВЫРУЧКА", body="768 025 ₽\nБар 368К + Кухня 100К +\n330К гарантии + кальяны")
add_card(slide, Inches(4.70), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ВАШИ РАСХОДЫ", body="284 476 ₽\nПерсонал 165 470 ₽ +\nБар 92 006 ₽ + Кухня 27К")
add_card(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(1.35), fill=BURGUNDY, header="ЧИСТАЯ ПРИБЫЛЬ", body="483 549 ₽\nРентабельность ~63%\nдо учёта кальянов")
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(9.5), Inches(0.25), "«Улица» внутри курорта: кирпичный фасад и брусчатка создают эффект ночного города под крышей отеля.", font_size=7.5, color=MUTED, font_name="Arial")
add_footer(slide, "05 / 10")

# ================= SLIDE 6 - ХЕДЛАЙНЕРЫ =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Х Е Д Л А Й Н Е Р Ы", "Имена, которые продают", slide_no="06 / 10")
# Two images top like slide 06 (two images side by side)
if os.path.exists(IMG["hall1"]):
    slide.shapes.add_picture(IMG["hall1"], Inches(0.55), Inches(1.75), width=Inches(8.10), height=Inches(3.60))
if os.path.exists(IMG["hall2"]):
    slide.shapes.add_picture(IMG["hall2"], Inches(8.85), Inches(1.75), width=Inches(3.93), height=Inches(3.60))
# Small captions under images (like template)
set_text_box(slide, Inches(0.55), Inches(5.45), Inches(8.10), Inches(0.25), "Потолочная световая конструкция — Helen Design / Большой LED-экран и сцена Foster Club", font_size=7.5, color=MUTED, font_name="Arial", alignment=PP_ALIGN.CENTER)
set_text_box(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(0.25), "Зеркальный бар Foster — фото Яндекс Карты", font_size=7.5, color=MUTED, font_name="Arial", alignment=PP_ALIGN.CENTER)

# Left text about Gleb
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(4.0), Inches(3.20))
# Actually images cover this area, so we need overlay text? Instead we will add text below images in cards
# Use 2 cards at bottom for Gleb and Noemy plus one burgundy highlight
add_card(slide, Inches(0.55), Inches(5.85), Inches(3.95), Inches(1.20), fill=DARK, header="GLEB FILIPCHENKOW", body="Соавтор хита «Mallorca» (лето 2025).\nВзрыв весной 2026 без бюджета —\nсотни тысяч просмотров. Далее —\nрелиз с Ольгой Серябкиной.")
add_card(slide, Inches(4.70), Inches(5.85), Inches(3.95), Inches(1.20), fill=DARK, header="NOEMY ABRANTES", body="Соавтор «Garota» и «HOJE».\nЯркий европейский вокал —\nидеален для ночной\nатмосферы Foster.")
add_card(slide, Inches(8.85), Inches(5.85), Inches(3.93), Inches(1.20), fill=BURGUNDY, header="ПОЧЕМУ ПРОДАСТСЯ", body="Первый привоз в Крым •\nОтложенный спрос с лета •\n5 декабря — единственное окно\nНоябрь у артистов расписан.")
# But this overlaps previous? Adjust: Remove previous image captions and cards are at 5.85
# Need to adjust footer
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(9.5), Inches(0.25), "Фото: главный зал, сцена и бар — галерея Foster Club, Яндекс Карты / Потолок Helen Design", font_size=7.5, color=MUTED, font_name="Arial")
add_footer(slide, "06 / 10")

# ================= SLIDE 7 - ИНВЕСТИЦИИ КАКТУСЫ =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "И Н В Е С Т И Ц И И   О Р Г А Н И З А Т О Р А", "Кактуса берёт весь риск", slide_no="07 / 10")
# image like slide 07 (one big left, one small right)
if os.path.exists(IMG["dj_bar"]):
    slide.shapes.add_picture(IMG["dj_bar"], Inches(0.55), Inches(1.75), width=Inches(7.90), height=Inches(3.50))
if os.path.exists(IMG["vip"]):
    slide.shapes.add_picture(IMG["vip"], Inches(8.70), Inches(1.75), width=Inches(4.10), height=Inches(2.70))
# bottom cards - 2 cards plus burgundy
add_card(slide, Inches(0.55), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ГОНОРАР И ЛОГИСТИКА", body="550 000 ₽ гонорар +\n80 000 ₽ перелёт/трансфер +\n30 000 ₽ техрайдер")
add_card(slide, Inches(4.70), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ПРОДАКШН И ПРОДВИЖЕНИЕ", body="75К местные артисты +\n30К видео + 25К фото +\n100К афиши/блогеры")
add_card(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(1.35), fill=BURGUNDY, header="ВАШИ 200 040 ₽ ГАРАНТ", body="180 000 ₽ коктейли +\n20 040 ₽ шаттл — фикс\nнезависимо от явки")
# Additional small note card above cards? Original slide 07 had a small card right under image: VIP-зал etc. We can add
# add bullet left text? Already covered via cards, but add footer
set_text_box(slide, Inches(8.70), Inches(4.60), Inches(4.10), Inches(0.70), "Фото: DJ-бар и VIP-зал Foster Club — Яндекс Карты. Бар — сердце ночного «лофта».", font_size=7.5, color=MUTED, font_name="Arial", alignment=PP_ALIGN.CENTER)
add_footer(slide, "07 / 10")

# ================= SLIDE 8 - СЦЕНАРИИ (timeline style like slide 11) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "С Ц Е Н А Р И И   П Р О Д А Ж", "Три сценария — один вывод", slide_no="08 / 10")
# Top image timeline style
if os.path.exists(IMG["timeline"]):
    slide.shapes.add_picture(IMG["timeline"], Inches(0.55), Inches(1.75), width=Inches(12.23), height=Inches(2.35))
set_text_box(slide, Inches(0.55), Inches(4.20), Inches(12.23), Inches(0.25), "Тайминг — концептуальный. Продажи билетов: 01.11–07.11 / 08.11–14.11 / с 15.11. Фото: галерея Foster Club — Яндекс Карты", font_size=7.5, color=MUTED, font_name="Arial", alignment=PP_ALIGN.CENTER)
# 3 scenario cards like 5 acts in template slide 11 (5 cards), we make 3
scenarios = [
    ("250 ЧЕЛ.", "Пессимистичный", "3 300 ₽ средн.\n825 000 ₽ выручка\n–295 040 ₽ убыток", DARK),
    ("300 ЧЕЛ.", "Базовый", "3 400 ₽ средн.\n1 020 000 ₽ выручка\n–100 040 ₽ убыток", DARK),
    ("350 ЧЕЛ.", "SOLD OUT ✓", "3 400 ₽ средн.\n1 190 000 ₽ выручка\n+69 960 ₽ прибыль", BURGUNDY),
]
for i, (t1,t2,body,fill) in enumerate(scenarios):
    left = Inches(0.55 + i*4.20)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(4.65), Inches(3.95), Inches(2.05))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.adjustments[0]=0.08
    # time header like 20:00
    set_text_box(slide, left+Inches(0.20), Inches(4.80), Inches(3.55), Inches(0.30), t1, font_size=11, color=GOLD_LIGHT if fill!=BURGUNDY else WHITE, bold=True, alignment=PP_ALIGN.CENTER, font_name="Arial")
    set_text_box(slide, left+Inches(0.20), Inches(5.15), Inches(3.55), Inches(0.25), t2, font_size=8, color=BEIGE, bold=True, alignment=PP_ALIGN.CENTER, font_name="Arial")
    tb = slide.shapes.add_textbox(left+Inches(0.20), Inches(5.50), Inches(3.55), Inches(0.90))
    tf = tb.text_frame
    tf.word_wrap=True
    p = tf.paragraphs[0]
    p.text = body
    p.font.name="Arial"
    p.font.size=Pt(8)
    p.font.color.rgb=BEIGE3 if fill!=BURGUNDY else WHITE
    p.alignment=PP_ALIGN.CENTER
    p.line_spacing=1.1

add_footer(slide, "08 / 10")

# ================= SLIDE 9 - АУДИТОРИЯ / РЕЗЮМЕ (like slide 12) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Р Е З Ю М Е", "Ваша сделка на 483 500 ₽", slide_no="09 / 10")
# left image like slide12 (small left)
if os.path.exists(IMG["audience"]):
    slide.shapes.add_picture(IMG["audience"], Inches(0.55), Inches(1.75), width=Inches(2.70), height=Inches(5.00))
# 4 cards right side (2x2 grid) like slide12's 4 cards
cards = [
    ("180 000 ₽ • КОКТЕЙЛИ", "300 × 600 ₽ — фикс,\nнезависимо от явки", DARK),
    ("20 040 ₽ • ШАТТЛ", "Охрана, клининг, хостес\nВелком-центра — фикс", DARK),
    ("120 000 ₽ • VIP И СТОЛЫ", "VIP №1 40К + №2 30К +\nстолы 50К — ваша зона", DARK),
    ("КАЛЬЯНЫ • МАКАО", "Доход по факту продаж\nполностью ваш", DARK),
]
for i, (h,b,fill) in enumerate(cards):
    row = i // 2
    col = i % 2
    left = Inches(3.55 + col*4.85)
    top = Inches(1.75 + row*1.75)
    add_card(slide, left, top, Inches(4.65), Inches(1.45), fill=fill, header=h, body=b)
# burgundy highlight spanning bottom
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.55), Inches(5.45), Inches(9.50), Inches(1.00))
shape.fill.solid()
shape.fill.fore_color.rgb = BURGUNDY
shape.line.fill.background()
shape.adjustments[0]=0.08
tb = slide.shapes.add_textbox(Inches(3.75), Inches(5.60), Inches(9.10), Inches(0.70))
tf = tb.text_frame
tf.word_wrap=True
p=tf.paragraphs[0]
p.text="Почему работает: гибрид «танцы + караоке + лаунж» удерживает компанию с разными настроениями в одном месте, а ваш риск — в 4 раза ниже, чем у организатора."
p.font.name="Arial"
p.font.size=Pt(8)
p.font.color.rgb=WHITE
p.line_spacing=1.2

set_text_box(slide, Inches(0.55), Inches(6.85), Inches(2.70), Inches(0.25), "Фото: аудитория Foster Club — Яндекс Карты", font_size=7.5, color=MUTED, font_name="Arial", alignment=PP_ALIGN.CENTER)
add_footer(slide, "09 / 10")

# ================= SLIDE 10 - ИТОГ (like slide 13 full image) =================
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["final"])
# center text like slide13
set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), "И Т О Г", font_size=10, color=GOLD, bold=True, font_name="Arial")
tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(9), Inches(1.2))
tf = tb.text_frame
tf.word_wrap=True
p=tf.paragraphs[0]
p.text="Foster — ночь, которую"
p.font.name="Arial"; p.font.size=Pt(24); p.font.color.rgb=BEIGE; p.font.bold=True; p.line_spacing=0.95
p=tf.add_paragraph()
p.text="не нужно искать"
p.font.name="Arial"; p.font.size=Pt(24); p.font.color.rgb=BEIGE; p.font.bold=True; p.line_spacing=0.95
p=tf.add_paragraph()
p.text="за пределами курорта"
p.font.name="Arial"; p.font.size=Pt(24); p.font.color.rgb=GOLD; p.font.bold=True; p.line_spacing=0.95

tb = slide.shapes.add_textbox(Inches(0.55), Inches(2.05), Inches(7), Inches(0.70))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]
p.text="Роскошный лофт, голливудский VIP, караоке-сцена и зеркальный свет — камерный театр одной ночи внутри Mriya. Ваша прибыль 483 549 ₽ ждёт 5 декабря."
p.font.name="Arial"; p.font.size=Pt(9); p.font.color.rgb=BEIGE3; p.line_spacing=1.25

# contact card dark
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.20), Inches(7.5), Inches(1.35))
shape.fill.solid()
shape.fill.fore_color.rgb = DARKER
shape.line.fill.background()
shape.adjustments[0]=0.08
set_text_box(slide, Inches(0.75), Inches(5.35), Inches(7.10), Inches(0.25), "КОНТАКТЫ И ТОЧКА НА КАРТЕ", font_size=7.5, color=GOLD_LIGHT, bold=True, font_name="Arial")
tb = slide.shapes.add_textbox(Inches(0.75), Inches(5.65), Inches(7.10), Inches(0.70))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]
p.text="Foster Club, Mriya Resort & SPA | Ялта, с. Оползневое, ул. Генерала Острякова, 9 | 300–350 гостей | Билеты 3 000–4 000 ₽ | Организатор: Кактуса"
p.font.name="Arial"; p.font.size=Pt(8); p.font.color.rgb=BEIGE3; p.line_spacing=1.15
# CTA line bottom
set_text_box(slide, Inches(0.55), Inches(6.80), Inches(12), Inches(0.25), "Перед публичным использованием проверьте актуальный статус и график — чек-лист на следующем слайде. Презентация пересобрана в vecherinka.", font_size=7.5, color=MUTED, font_name="Arial")
set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), "10 / 10", font_size=9, color=WHITE, alignment=PP_ALIGN.RIGHT, font_name="Arial")
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(5), Inches(0.25), "Пересборка на подлинных фото из галереи Яндекс Карт", font_size=7.5, color=MUTED, font_name="Arial")

# ================= SAVE =================
out_pptx = "/home/user/Foster_Gleb_Noemy_Presentation.pptx"
prs.save(out_pptx)
print(f"Saved PPTX {out_pptx} slides={len(prs.slides)}")
