#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Чистая пересборка: только финансы Foster (доходы/затраты/прибыль) без проживания.
Убраны визуалы слайдов 4,5,12,13,14 из шаблона.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import os

GOLD = RGBColor(0xC9, 0xA2, 0x5E)
GOLD_LIGHT = RGBColor(0xE5, 0xC8, 0x88)
BEIGE = RGBColor(0xF5, 0xF1, 0xE8)
BEIGE2 = RGBColor(0xDD, 0xD8, 0xCC)
BEIGE3 = RGBColor(0xE2, 0xDE, 0xD4)
MUTED = RGBColor(0x7E, 0x7A, 0x72)
DARK = RGBColor(0x14, 0x14, 0x19)
DARKER = RGBColor(0x09, 0x09, 0x0B)
BURGUNDY = RGBColor(0x7A, 0x1F, 0x2B)
WHITE = RGBColor(0xFF,0xFF,0xFF)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg",
    "about": "/tmp/foster_images/slide02_img3_Picture_5.jpg",
    "concept_bg": "/tmp/foster_images/slide03_img4_Picture_1.jpg",
    # REMOVED: context (slide04), entrance (slide05), audience (slide12), final (slide13) — не используем
    # Замены:
    "costs": "/tmp/foster_images/slide06_img7_Picture_5.jpg", # hall 8.1x3.6
    "costs2": "/tmp/foster_images/slide06_img8_Picture_6.jpg",
    "income": "/tmp/foster_images/slide07_img9_Picture_5.jpg", # DJ bar
    "income_vip": "/tmp/foster_images/slide07_img10_Picture_6.jpg",
    "hall1": "/tmp/foster_images/slide06_img7_Picture_5.jpg",
    "hall2": "/tmp/foster_images/slide06_img8_Picture_6.jpg",
    "dj_bar": "/tmp/foster_images/slide07_img9_Picture_5.jpg",
    "vip": "/tmp/foster_images/slide07_img10_Picture_6.jpg",
    "bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "loft1": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "loft2": "/tmp/foster_images/slide09_img13_Picture_6.jpg",
    "timeline": "/tmp/foster_images/slide11_img18_Picture_5.jpg",
    # альтернативы для запрещённых:
    "alt_bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "alt_loft": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "alt_multi1": "/tmp/foster_images/slide10_img14_Picture_5.jpg",
    "alt_multi2": "/tmp/foster_images/slide10_img15_Picture_6.jpg",
}

def add_bg_image(slide, img_path, darken=0.45):
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x0A,0x0A,0x0F)
        shape.line.fill.background()
        return shape
    return None

def set_text_box(slide, left, top, width, height, text, font_name="Arial", font_size=10, color=BEIGE, bold=False, italic=False, alignment=PP_ALIGN.LEFT):
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
    p.line_spacing = 1.0
    return txBox

def add_gold_line(slide, left=Inches(0.57), top=Inches(1.42), width=Inches(0.70), height=Inches(0.05)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    return shape

def add_card(slide, left, top, width, height, fill=DARK, header=None, body=None, header_color=GOLD_LIGHT, body_color=BEIGE3, radius=True):
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

def add_header_block(slide, spaced_title, main_title, slide_no="01 / 08"):
    set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), spaced_title, font_size=10, color=GOLD, bold=True)
    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(7.5), Inches(0.95))
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
    set_text_box(slide, Inches(11.70), Inches(0.32), Inches(1.10), Inches(0.30), slide_no, font_size=9, color=MUTED, alignment=PP_ALIGN.RIGHT)
    # footer (убраны упоминания Mriya, оставлен только фото-кредит без привязки)
    set_text_box(slide, Inches(0.55), Inches(7.12), Inches(6), Inches(0.25), "Фото: галерея Foster Club — Яндекс Карты", font_size=7.5, color=MUTED)
    set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), f"FOSTER • {slide_no.split('/')[0].strip()}", font_size=7.5, color=MUTED, bold=True, alignment=PP_ALIGN.RIGHT)

# ================= SLIDE 1 COVER =================
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["cover"])
# Top beige block with FOSTER
rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(0.45), Inches(2.45), Inches(0.98))
rect.fill.solid()
rect.fill.fore_color.rgb = BEIGE
rect.line.fill.background()
rect.adjustments[0]=0.08
if os.path.exists("/tmp/foster_images/slide01_img2_Picture_3.png"):
    slide.shapes.add_picture("/tmp/foster_images/slide01_img2_Picture_3.png", Inches(0.70), Inches(0.60), width=Inches(2.15), height=Inches(0.68))
else:
    set_text_box(slide, Inches(0.70), Inches(0.65), Inches(2.15), Inches(0.30), "FOSTER", font_size=12, color=DARK, bold=True, alignment=PP_ALIGN.CENTER)
    set_text_box(slide, Inches(0.70), Inches(0.90), Inches(2.15), Inches(0.20), "NIGHT & KARAOKE CLUB", font_size=6, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)
set_text_box(slide, Inches(9.2), Inches(0.45), Inches(3.60), Inches(0.30), "Ночной клуб Foster  •  Крым", font_size=7, color=WHITE, alignment=PP_ALIGN.RIGHT)
set_text_box(slide, Inches(9.2), Inches(0.68), Inches(3.60), Inches(0.25), "5 декабря 2026  •  300–350 гостей  •  SOLD OUT цель", font_size=7, color=GOLD, bold=True, alignment=PP_ALIGN.RIGHT)
# Main title
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(8.5), Inches(1.6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "GLEB FILIPCHENKOW"
p.font.name = "Arial"; p.font.size = Pt(32); p.font.color.rgb = WHITE; p.font.bold = True; p.line_spacing = 0.9
p = tf.add_paragraph()
p.text = "& NOEMY ABRANTES"
p.font.name = "Arial"; p.font.size = Pt(32); p.font.color.rgb = GOLD; p.font.bold = True; p.line_spacing = 0.9; p.space_before=Pt(4)
add_gold_line(slide, Inches(0.57), Inches(3.55), Inches(0.85), Inches(0.06))
set_text_box(slide, Inches(0.55), Inches(3.75), Inches(8.2), Inches(0.45), "Первый привоз в Крым  •  «Mallorca» — хит лета 2025  •  Garota • HOJE", font_size=9, color=BEIGE2)
# dark card
card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(4.40), Inches(6.0), Inches(1.55))
card.fill.solid()
card.fill.fore_color.rgb = DARK
card.line.fill.background()
card.adjustments[0]=0.06
set_text_box(slide, Inches(0.75), Inches(4.55), Inches(5.6), Inches(0.25), "◆  Ночь, которую не нужно искать за пределами клуба", font_size=8, color=GOLD, bold=True)
tb = slide.shapes.add_textbox(Inches(0.75), Inches(4.85), Inches(5.6), Inches(1.0))
tf = tb.text_frame
tf.word_wrap = True
p=tf.paragraphs[0]
p.text="Событие продаёт имя: вирусный трек «Mallorca» без бюджета, сотни тысяч просмотров. Привоз ждали с лета — 5 декабря единственное окно."
p.font.name="Arial"; p.font.size=Pt(8); p.font.color.rgb=BEIGE3; p.line_spacing=1.25
# 3 stats
for i, (num, lbl) in enumerate([("768 025 ₽","Выручка + кальяны"), ("284 476 ₽","Ваши затраты"), ("483 549 ₽","Чистая прибыль 63%")]):
    left = Inches(0.55 + i*2.05)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(6.20), Inches(1.90), Inches(0.75))
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK
    shape.line.fill.background()
    shape.adjustments[0]=0.08
    set_text_box(slide, left+Inches(0.12), Inches(6.28), Inches(1.66), Inches(0.25), num, font_size=9, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    set_text_box(slide, left+Inches(0.12), Inches(6.55), Inches(1.66), Inches(0.20), lbl, font_size=6, color=GOLD_LIGHT, bold=True, alignment=PP_ALIGN.CENTER)
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(6), Inches(0.25), "Пересборка на подлинных фото Foster Club из галереи Яндекс Карт", font_size=7.5, color=MUTED)
set_text_box(slide, Inches(11.30), Inches(7.12), Inches(1.50), Inches(0.25), "01 / 08", font_size=9, color=WHITE, alignment=PP_ALIGN.RIGHT)

# ================= SLIDE 2 ПОЗИЦИОНИРОВАНИЕ (без Mriya) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "П О З И Ц И О Н И Р О В А Н И Е", "Камерный премиум-клуб", slide_no="02 / 08")
if os.path.exists(IMG["about"]):
    slide.shapes.add_picture(IMG["about"], Inches(7.00), Inches(1.75), width=Inches(5.78), height=Inches(3.50))
tb = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(6.15), Inches(3.40))
tf = tb.text_frame
tf.word_wrap = True
bullets = [
    ("Формат:", "ночной и караоке-клуб — вечерняя «гостиная» с танцполом, сценой, VIP-залами."),
    ("Гости:", "300–350 человек — полный SOLD OUT камерного зала (до 60 мест за столами + танцпол)."),
    ("Почему сейчас:", "летний привоз сорвался из-за топливного кризиса; 5 декабря — единственное окно артистов."),
    ("Итог для вас:", "вы предоставляете только площадку и бар — без вложений в артистов и без проживания."),
]
for label, body in bullets:
    p = tf.add_paragraph()
    p.space_before=Pt(6)
    r=p.add_run(); r.text="◆  "; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=GOLD; r.font.bold=True
    r=p.add_run(); r.text=label+" "; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=body; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=RGBColor(0x2A,0x2A,0x30)
# 3 карточки без упоминания проживания
add_card(slide, Inches(0.55), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ПЛОЩАДКА", body="FOSTER — клуб, бар,\nкухня, персонал,\nоборудование и VIP-залы")
add_card(slide, Inches(4.70), Inches(5.45), Inches(3.95), Inches(1.35), fill=DARK, header="ОРГАНИЗАТОР", body="КАКТУСА — гонорары,\nлогистика, техрайдер,\nпродвижение, продажи")
add_card(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(1.35), fill=DARK, header="ДАТА И БИЛЕТЫ", body="Суббота, 5 декабря 2026\n3 000 ₽ → 3 500 ₽ → 4 000 ₽\nЦель: 350 SOLD OUT")

# ================= SLIDE 3 КОНЦЕПЦИЯ =================
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["concept_bg"])
set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), "К О Н Ц Е П Ц И Я   С Д Е Л К И", font_size=10, color=GOLD, bold=True)
tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(10), Inches(0.95))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="«Чистые финансы — без скрытых затрат»"; p.font.name="Arial"; p.font.size=Pt(24); p.font.color.rgb=BEIGE; p.font.bold=True; p.line_spacing=0.95
add_gold_line(slide)
set_text_box(slide, Inches(11.70), Inches(0.32), Inches(1.10), Inches(0.30), "03 / 08", font_size=9, color=MUTED, alignment=PP_ALIGN.RIGHT)
add_card(slide, Inches(0.55), Inches(1.85), Inches(3.95), Inches(1.35), fill=DARKER, header="1 • ДОХОДЫ", body="768 025 ₽ + кальяны\nБар 368К + Кухня 100К +\n180К коктейли + 120К столы/VIP")
add_card(slide, Inches(4.70), Inches(1.85), Inches(3.95), Inches(1.35), fill=DARKER, header="2 • ЗАТРАТЫ", body="284 476 ₽\nПерсонал 165К + Бар 92К +\nКухня 27К — всё покрывается")
add_card(slide, Inches(8.85), Inches(1.85), Inches(3.93), Inches(1.35), fill=DARKER, header="3 • ПРИБЫЛЬ", body="483 549 ₽ чистыми\nРентабельность ~63%\nБез проживания — ноль рисков")
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(9.5), Inches(0.25), "Проживание артистов полностью исключено из модели — никаких скидок и обязательств. Фото: красный лаунж Foster Club.", font_size=7.5, color=MUTED)

# ================= SLIDE 4 НАШИ ДОХОДЫ (чистые) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Н А Ш И   Д О Х О Д Ы", "768 025 ₽ + кальяны", slide_no="04 / 08")
# Используем альтернативные фото вместо запрещённых (бар и лофт)
if os.path.exists(IMG["alt_bar"]):
    slide.shapes.add_picture(IMG["alt_bar"], Inches(0.55), Inches(1.75), width=Inches(3.90), height=Inches(3.50))
    set_text_box(slide, Inches(0.55), Inches(5.35), Inches(3.90), Inches(0.25), "Бар Foster — фото Яндекс Карты", font_size=7, color=MUTED, alignment=PP_ALIGN.CENTER)
if os.path.exists(IMG["loft1"]):
    slide.shapes.add_picture(IMG["loft1"], Inches(4.65), Inches(1.75), width=Inches(3.90), height=Inches(3.50))
    set_text_box(slide, Inches(4.65), Inches(5.35), Inches(3.90), Inches(0.25), "Лаунж-зона — фото Foster Club", font_size=7, color=MUTED, alignment=PP_ALIGN.CENTER)
# Таблица доходов справа (вместо третьего фото)
incomes = [
    ("Бар", "368 025 ₽", "валовая выручка"),
    ("Кухня", "100 000 ₽", "валовая выручка"),
    ("◆ Коктейли", "180 000 ₽", "300 × 600 ₽ → вам"),
    ("◆ Столы", "50 000 ₽", "4×5 000 + 3×10 000"),
    ("◆ VIP №1", "40 000 ₽", "депозит → вам"),
    ("◆ VIP №2", "30 000 ₽", "депозит → вам"),
]
tb = slide.shapes.add_textbox(Inches(8.75), Inches(1.75), Inches(4.00), Inches(3.50))
tf = tb.text_frame; tf.word_wrap=True
for name, price, note in incomes:
    p=tf.add_paragraph(); p.space_before=Pt(3)
    is_hl = "◆" in name
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(8); r.font.color.rgb=GOLD if is_hl else DARK; r.font.bold=True
    r=p.add_run(); r.text=price+"  "; r.font.name="Arial"; r.font.size=Pt(8.5); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text="— "+note; r.font.name="Arial"; r.font.size=Pt(6); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(8)
r=p.add_run(); r.text="Итого: 768 025 ₽ + кальяны от Макао"; r.font.name="Arial"; r.font.size=Pt(8); r.font.color.rgb=BURGUNDY; r.font.bold=True
# bottom card highlight
add_card(slide, Inches(0.55), Inches(5.80), Inches(12.23), Inches(1.00), fill=BURGUNDY, header="ГАРАНТИРОВАНО ДОГОВОРОМ", body="330 000 ₽ (180К коктейли + 50К столы + 70К VIP) — фикс, независимо от явки. Кальяны от Макао — по факту продаж, полностью ваши.")

# ================= SLIDE 5 НАШИ ЗАТРАТЫ (чистые) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Н А Ш И   З А Т Р А Т Ы", "Всего 284 476 ₽", slide_no="05 / 08")
# Левая картинка - используем alternative multi (не запрещённые)
if os.path.exists(IMG["alt_multi1"]):
    slide.shapes.add_picture(IMG["alt_multi1"], Inches(0.55), Inches(1.75), width=Inches(2.70), height=Inches(3.50))
if os.path.exists(IMG["alt_multi2"]):
    slide.shapes.add_picture(IMG["alt_multi2"], Inches(3.40), Inches(1.75), width=Inches(3.80), height=Inches(3.50))
# Правая таблица затрат
costs = [
    ("Персонал клуба", "165 470 ₽", "официанты, бармены, повара"),
    ("Себестоимость бара", "92 006 ₽", "25% от 368 025 ₽"),
    ("Себестоимость кухни", "27 000 ₽", "27% от 100 000 ₽"),
]
tb = slide.shapes.add_textbox(Inches(7.50), Inches(1.85), Inches(5.20), Inches(3.40))
tf=tb.text_frame; tf.word_wrap=True
for name, price, note in costs:
    p=tf.add_paragraph(); p.space_before=Pt(8)
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=price; r.font.name="Arial"; r.font.size=Pt(10); r.font.color.rgb=DARK; r.font.bold=True
    p=tf.add_paragraph()
    r=p.add_run(); r.text=note; r.font.name="Arial"; r.font.size=Pt(7); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(10)
r=p.add_run(); r.text="Итого затраты: 284 476 ₽"; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=BURGUNDY; r.font.bold=True
# bottom 3 cards: breakdown + comparison
add_card(slide, Inches(0.55), Inches(5.60), Inches(3.95), Inches(1.20), fill=DARK, header="ПЕРСОНАЛ 165К", body="Самая крупная статья —\nно покрывается баром\nи кухней.")
add_card(slide, Inches(4.70), Inches(5.60), Inches(3.95), Inches(1.20), fill=DARK, header="ФУДКОСТ", body="Бар 25% + Кухня 27% —\nнормальные ресторанные\nпоказатели.")
add_card(slide, Inches(8.85), Inches(5.60), Inches(3.93), Inches(1.20), fill=DARK, header="0 ₽ ПРОЖИВАНИЕ", body="Полностью исключено:\nни скидок, ни затрат,\nни обязательств.")

# ================= SLIDE 6 ФИНАНСОВЫЙ РЕЗУЛЬТАТ =================
slide = prs.slides.add_slide(BLANK)
# Full image background for impact (use cover again, allowed)
add_bg_image(slide, IMG["cover"])
set_text_box(slide, Inches(0.55), Inches(0.32), Inches(9.5), Inches(0.30), "Ф И Н А Н С О В Ы Й   Р Е З У Л Ь Т А Т", font_size=10, color=GOLD, bold=True)
tb=slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(8), Inches(0.95))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Чистая прибыль — 483 549 ₽"; p.font.name="Arial"; p.font.size=Pt(26); p.font.color.rgb=BEIGE; p.font.bold=True
add_gold_line(slide)
set_text_box(slide, Inches(11.70), Inches(0.32), Inches(1.10), Inches(0.30), "06 / 08", font_size=9, color=MUTED, alignment=PP_ALIGN.RIGHT)
# 3 big cards dark on image
add_card(slide, Inches(0.55), Inches(1.85), Inches(3.95), Inches(1.60), fill=DARKER, header="ДОХОДЫ", body="768 025 ₽\nБар 368 025 + Кухня 100 000 +\nКоктейли 180К + Столы 50К +\nVIP 70К + кальяны")
add_card(slide, Inches(4.70), Inches(1.85), Inches(3.95), Inches(1.60), fill=DARKER, header="ЗАТРАТЫ", body="284 476 ₽\nПерсонал 165 470 +\nБар 92 006 + Кухня 27 000\nБез проживания — 0 ₽")
add_card(slide, Inches(8.85), Inches(1.85), Inches(3.93), Inches(1.60), fill=BURGUNDY, header="ПРИБЫЛЬ 63%", body="483 549 ₽\n768 025 – 284 476\n+ кальяны от Макао\nРентабельность ~63%")
# footer highlight
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.80), Inches(12.23), Inches(0.95))
shape.fill.solid(); shape.fill.fore_color.rgb = BEIGE; shape.line.fill.background(); shape.adjustments[0]=0.08
set_text_box(slide, Inches(0.75), Inches(5.95), Inches(11.8), Inches(0.25), "◆  Что зафиксировать в договоре", font_size=8, color=BURGUNDY, bold=True)
tb = slide.shapes.add_textbox(Inches(0.75), Inches(6.20), Inches(11.8), Inches(0.45))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="180 000 ₽ за коктейли (фикс)  •  20 040 ₽ за шаттл (охрана/клининг/хостес)  •  330 000 ₽ гарант до старта  •  VIP и столы — только ваша зона  •  Оплата 1–3 рабочих дня"; p.font.name="Arial"; p.font.size=Pt(7); p.font.color.rgb=DARK; p.line_spacing=1.2
set_text_box(slide, Inches(0.55), Inches(7.12), Inches(9), Inches(0.25), "Проживание в Mriya Resort исключено — никаких обязательств с вашей стороны.", font_size=7.5, color=MUTED)

# ================= SLIDE 7 ХЕДЛАЙНЕРЫ (сохранён, но без упоминания проживания) =================
slide = prs.slides.add_slide(BLANK)
add_header_block(slide, "Х Е Д Л А Й Н Е Р Ы", "Имена, которые продают", slide_no="07 / 08")
if os.path.exists(IMG["hall1"]):
    slide.shapes.add_picture(IMG["hall1"], Inches(0.55), Inches(1.75), width=Inches(8.10), height=Inches(3.60))
if os.path.exists(IMG["hall2"]):
    slide.shapes.add_picture(IMG["hall2"], Inches(8.85), Inches(1.75), width=Inches(3.93), height=Inches(3.60))
set_text_box(slide, Inches(0.55), Inches(5.45), Inches(8.10), Inches(0.25), "Сцена Foster Club — большой LED-экран и свет Helen Design", font_size=7, color=MUTED, alignment=PP_ALIGN.CENTER)
set_text_box(slide, Inches(8.85), Inches(5.45), Inches(3.93), Inches(0.25), "Зеркальный бар — Яндекс Карты", font_size=7, color=MUTED, alignment=PP_ALIGN.CENTER)
add_card(slide, Inches(0.55), Inches(5.85), Inches(3.95), Inches(1.20), fill=DARK, header="GLEB FILIPCHENKOW", body="Хит «Mallorca» (лето 2025),\nвзрыв весной 2026 без\nбюджета. Далее — релиз с\nОльгой Серябкиной.")
add_card(slide, Inches(4.70), Inches(5.85), Inches(3.95), Inches(1.20), fill=DARK, header="NOEMY ABRANTES", body="«Garota» и «HOJE».\nЕвропейский вокал для\nночной атмосферы Foster.\nПервый привоз в Крым.")
add_card(slide, Inches(8.85), Inches(5.85), Inches(3.93), Inches(1.20), fill=BURGUNDY, header="ПОЧЕМУ ПРОДАСТСЯ", body="Имя продаёт само событие.\nОтложенный спрос с лета.\n5 декабря — единственное окно.")

# ================= SLIDE 8 ИТОГ (альтернатива запрещённому слайду 13) =================
slide = prs.slides.add_slide(BLANK)
# White clean, no full image (since 13 banned with full image, we use bar image alternative left)
if os.path.exists(IMG["alt_bar"]):
    slide.shapes.add_picture(IMG["alt_bar"], Inches(0.55), Inches(1.75), width=Inches(4.20), height=Inches(4.00))
    set_text_box(slide, Inches(0.55), Inches(5.85), Inches(4.20), Inches(0.25), "Лаунж Foster Club — фото Яндекс Карты", font_size=7, color=MUTED, alignment=PP_ALIGN.CENTER)
add_header_block(slide, "И Т О Г", "Ваша сделка на 483 500 ₽", slide_no="08 / 08")
tb = slide.shapes.add_textbox(Inches(5.15), Inches(1.85), Inches(7.50), Inches(1.0))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Минимальный риск — максимальная отдача."; p.font.name="Arial"; p.font.size=Pt(12); p.font.color.rgb=DARK; p.font.bold=True
p=tf.add_paragraph(); p.text="Вы не платите за артистов и проживание — только операционка, которая окупается баром."; p.font.name="Arial"; p.font.size=Pt(8); p.font.color.rgb=RGBColor(0x4A,0x4A,0x52)
# 2 cards right
add_card(slide, Inches(5.15), Inches(3.10), Inches(3.70), Inches(1.35), fill=DARK, header="ЧТО ФИКСИРОВАТЬ", body="180 000 ₽ коктейли (фикс)\n50 000 ₽ столы + 70 000 ₽ VIP\n20 040 ₽ шаттл")
add_card(slide, Inches(9.05), Inches(3.10), Inches(3.70), Inches(1.35), fill=BURGUNDY, header="СРОК ОПЛАТЫ", body="1–3 рабочих дня\nпосле мероприятия\nЖивыми деньгами")
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.15), Inches(4.70), Inches(7.50), Inches(0.90))
shape.fill.solid(); shape.fill.fore_color.rgb=BEIGE; shape.line.fill.background(); shape.adjustments[0]=0.08
set_text_box(slide, Inches(5.35), Inches(4.85), Inches(7.10), Inches(0.25), "Следующий шаг", font_size=7.5, color=BURGUNDY, bold=True)
tb = slide.shapes.add_textbox(Inches(5.35), Inches(5.10), Inches(7.10), Inches(0.40))
tf=tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Согласовать договор → запустить продажи с 01.11 → 5 декабря SOLD OUT 350 гостей."; p.font.name="Arial"; p.font.size=Pt(7.5); p.font.color.rgb=DARK

out = "/home/user/Foster_Gleb_Noemy_Presentation.pptx"
prs.save(out)
print(f"Saved PPTX {out} slides={len(prs.slides)}")
