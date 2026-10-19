#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# Brand colors inspired by Mriya Resort
NAVY = RGBColor(0x0A, 0x1A, 0x2F)  # #0A1A2F deep navy
NAVY2 = RGBColor(0x12, 0x2E, 0x4B) # lighter navy
NAVY3 = RGBColor(0x1A, 0x3A, 0x5C)
GOLD = RGBColor(0xC9, 0xA9, 0x6A)  # #C9A96A
GOLD_LIGHT = RGBColor(0xE8, 0xD5, 0xB5)
GOLD_DARK = RGBColor(0xA8, 0x8A, 0x4A)
BEIGE = RGBColor(0xF6, 0xF1, 0xE7)  # #F6F1E7
BEIGE2 = RGBColor(0xED, 0xE6, 0xD6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CREAM = RGBColor(0xFF, 0xFD, 0xF7)
MUTED = RGBColor(0x8A, 0x9A, 0xB0)
LIGHT_GRAY = RGBColor(0xE5, 0xE9, 0xEF)
DARK_GOLD = RGBColor(0x9C, 0x85, 0x4A)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
# remove default layouts margins

def hex2rgb(h):
    h=h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill, line=None, radius=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    # adjust corner radius
    if radius:
        shape.adjustments[0] = 0.08
    return shape

def add_circle(slide, left, top, dia, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, dia, dia)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape

def add_text(slide, left, top, width, height, text, size=12, color=NAVY, bold=False, align=PP_ALIGN.LEFT, font_name="Calibri", line_spacing=1.0, all_caps=False, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text.upper() if all_caps else text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.font.italic = italic
    p.alignment = align
    p.line_spacing = line_spacing
    return txBox

def add_para(text_frame, text, size=12, color=NAVY, bold=False, align=PP_ALIGN.LEFT, font_name="Calibri", italic=False, space_before=Pt(0), space_after=Pt(4), all_caps=False):
    p = text_frame.add_paragraph()
    p.text = text.upper() if all_caps else text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.font.italic = italic
    p.alignment = align
    p.space_before = space_before
    p.space_after = space_after
    p.line_spacing = 1.1
    return p

# ========== SLIDE 1 - COVER ==========
slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank
set_bg(slide, NAVY)
# decorative gold line top
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), GOLD)
# left content
# small eyebrow
add_text(slide, Inches(0.8), Inches(0.6), Inches(3), Inches(0.3), "FOSTER  ×  КАКТУСА  ×  MRIYA RESORT", size=7.5, color=GOLD_LIGHT, bold=True, font_name="Calibri")
# add tiny line under eyebrow
add_rect(slide, Inches(0.8), Inches(0.9), Inches(1.2), Inches(0.02), GOLD)

add_text(slide, Inches(0.8), Inches(1.15), Inches(6.8), Inches(0.35), "ПРЕЗЕНТАЦИЯ ПРОЕКТА", size=10, color=GOLD, bold=False, font_name="Calibri")
# Main title
tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), Inches(7), Inches(1.6))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "GLEB FILIPCHENKOW"
p.font.size = Pt(44)
p.font.color.rgb = WHITE
p.font.bold = True
p.font.name = "Calibri"
p.line_spacing = 0.9
p = tf.add_paragraph()
p.text = "& NOEMY ABRANTES"
p.font.size = Pt(44)
p.font.color.rgb = GOLD
p.font.bold = True
p.font.name = "Calibri"
p.line_spacing = 0.9
p.space_before = Pt(2)
# subtitle divider
add_rect(slide, Inches(0.8), Inches(3.25), Inches(1.5), Inches(0.04), GOLD)
add_text(slide, Inches(0.8), Inches(3.45), Inches(7), Inches(0.5), "Первый привоз в Крым  •  Ночной клуб Foster  •  5 декабря 2026", size=11, color=BEIGE, bold=False)

# Key metrics row - 3 cards
for i, (num, label) in enumerate([("300–350", "ГОСТЕЙ"), ("63%", "РЕНТАБЕЛЬНОСТЬ"), ("483 549 ₽", "ВАША ПРИБЫЛЬ")]):
    left = Inches(0.8 + i*2.2)
    card = add_rect(slide, left, Inches(4.35), Inches(2.0), Inches(1.05), RGBColor(0x13,0x2B,0x4A), line=RGBColor(0x2A,0x4A,0x6E))
    # gold top border
    add_rect(slide, left, Inches(4.35), Inches(2.0), Inches(0.04), GOLD)
    add_text(slide, left+Inches(0.15), Inches(4.5), Inches(1.7), Inches(0.35), num, size=18, color=WHITE, bold=True, align=PP_ALIGN.LEFT)
    add_text(slide, left+Inches(0.15), Inches(4.9), Inches(1.7), Inches(0.2), label, size=7, color=GOLD_LIGHT, bold=True)
# bottom badge
add_rect(slide, Inches(0.8), Inches(6.45), Inches(3.8), Inches(0.55), WHITE)
add_text(slide, Inches(0.95), Inches(6.58), Inches(3.5), Inches(0.2), "ДАТА: СУББОТА, 5 ДЕКАБРЯ 2026  •  ФОРМАТ: ПРИВОЗ ХЕДЛАЙНЕРОВ", size=7, color=NAVY, bold=True)

# RIGHT visual - image placeholder with gradient card
# big card
card = add_rect(slide, Inches(7.8), Inches(0.6), Inches(4.7), Inches(6.3), WHITE, radius=True)
# image area (top)
img_shape = add_rect(slide, Inches(7.8), Inches(0.6), Inches(4.7), Inches(3.8), NAVY2)
# placeholder inner - concentric circles to mimic stage
add_circle(slide, Inches(9.15), Inches(1.4), Inches(1.9), RGBColor(0x1E,0x3D,0x5E))
add_circle(slide, Inches(9.45), Inches(1.7), Inches(1.3), RGBColor(0x2A,0x4F,0x7A))
add_circle(slide, Inches(9.8), Inches(2.05), Inches(0.6), GOLD)
add_text(slide, Inches(7.95), Inches(1.55), Inches(4.4), Inches(0.3), "✦  HEADLINERS NIGHT  ✦", size=7, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.1), Inches(3.45), Inches(4.1), Inches(0.4), "GLEB FILIPCHENKOW & NOEMY", size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.1), Inches(3.8), Inches(4.1), Inches(0.25), "Mallorca  •  Garota  •  HOJE", size=7, color=GOLD_LIGHT, align=PP_ALIGN.CENTER)

# bottom part of card - info
add_text(slide, Inches(8.15), Inches(4.7), Inches(4.0), Inches(0.25), "ПЛОЩАДКА", size=7, color=MUTED, bold=True)
add_text(slide, Inches(8.15), Inches(4.95), Inches(2.0), Inches(0.3), "FOSTER", size=13, color=NAVY, bold=True)
add_rect(slide, Inches(10.2), Inches(4.85), Inches(1.9), Inches(0.45), NAVY)
add_text(slide, Inches(10.25), Inches(4.92), Inches(1.8), Inches(0.3), "КРЫМ  •  300+ ГОСТЕЙ", size=7, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.15), Inches(5.45), Inches(4.0), Inches(0.2), "ОРГАНИЗАТОР", size=7, color=MUTED, bold=True)
add_text(slide, Inches(8.15), Inches(5.65), Inches(4.0), Inches(0.25), "КАКТУСА — логистика, продакшн, продажи", size=8, color=NAVY2)
add_rect(slide, Inches(8.15), Inches(6.05), Inches(4.0), Inches(0.02), LIGHT_GRAY)
add_text(slide, Inches(8.15), Inches(6.2), Inches(2.2), Inches(0.25), "БИЛЕТЫ ОТ 3 000 ₽", size=7, color=NAVY, bold=True)
add_text(slide, Inches(10.3), Inches(6.2), Inches(1.6), Inches(0.25), "SOLD OUT 350 ЧЕЛ.", size=7, color=GOLD_DARK, bold=True, align=PP_ALIGN.RIGHT)

# footer logos like
add_text(slide, Inches(7.8), Inches(7.05), Inches(4.7), Inches(0.2), "mriya resort  •  foster club  •  кактуса", size=6, color=RGBColor(0x4A,0x5A,0x70), align=PP_ALIGN.CENTER)

# ========== SLIDE 2 - ABOUT / KEY NUMBERS ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, BEIGE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
# left navy panel
add_rect(slide, Inches(0), Inches(0.06), Inches(4.6), Inches(7.44), NAVY)
add_text(slide, Inches(0.6), Inches(0.5), Inches(3.4), Inches(0.25), "01  —  О ПРОЕКТЕ", size=7, color=GOLD, bold=True)
add_text(slide, Inches(0.6), Inches(0.85), Inches(3.6), Inches(0.6), "Вечеринка,\nкоторая уже продана", size=22, color=WHITE, bold=True, line_spacing=0.95)
add_rect(slide, Inches(0.6), Inches(1.7), Inches(0.8), Inches(0.03), GOLD)
add_text(slide, Inches(0.6), Inches(1.9), Inches(3.5), Inches(1.0), "Привоз состоялся бы летом — сорвался из-за топливного кризиса. Публика ждёт. 5 декабря — единственное окно у артистов. Имя продаёт само событие.", size=9, color=BEIGE, line_spacing=1.3)
# stats in navy panel
stats = [("300–350", "гостей\nплан загрузки"), ("3 000–4 000 ₽", "цена билета\nдинамическое повышение"), ("SOLD OUT", "цель кактусы\nих мотивация = ваша защита")]
for i, (n,l) in enumerate(stats):
    top = Inches(3.2 + i*1.15)
    add_rect(slide, Inches(0.6), top, Inches(3.4), Inches(0.85), RGBColor(0x11,0x28,0x45), line=RGBColor(0x1E,0x3A,0x5A))
    add_text(slide, Inches(0.75), top+Inches(0.12), Inches(1.6), Inches(0.35), n, size=13, color=GOLD_LIGHT, bold=True)
    add_text(slide, Inches(2.1), top+Inches(0.12), Inches(1.7), Inches(0.6), l, size=7, color=BEIGE)

# right side - roles table
add_text(slide, Inches(5.1), Inches(0.55), Inches(7.5), Inches(0.3), "РОЛИ СТОРОН — КТО ЗА ЧТО ОТВЕЧАЕТ", size=8, color=NAVY, bold=True)
add_rect(slide, Inches(5.1), Inches(0.85), Inches(0.6), Inches(0.03), GOLD)
# 3 columns
roles = [
    ("FOSTER", "ПЛОЩАДКА", "Предоставляете клуб,\nбар, кухню, персонал\nи инфраструктуру", "0 ₽ аренды\n0 ₽ за артистов", NAVY),
    ("КАКТУСА", "ОРГАНИЗАТОР", "Берёт все риски:\nгонорары, логистика,\nпродакшн, продвижение", "~1,12 млн ₽\nинвестиций", GOLD),
    ("MRIYA", "ПАРТНЁР", "Проживание артистов\nбез скидки —\nчистый расчёт", "Без нагрузки\nна ваш бюджет", NAVY2),
]
for i, (t1,t2,desc,price,col) in enumerate(roles):
    left = Inches(5.1 + i*2.55)
    card = add_rect(slide, left, Inches(1.15), Inches(2.35), Inches(2.6), WHITE, radius=True)
    add_rect(slide, left, Inches(1.15), Inches(2.35), Inches(0.06), col)
    add_text(slide, left+Inches(0.2), Inches(1.35), Inches(1.95), Inches(0.2), t1, size=9, color=col, bold=True)
    add_text(slide, left+Inches(0.2), Inches(1.6), Inches(1.95), Inches(0.15), t2, size=6, color=MUTED, bold=True)
    add_rect(slide, left+Inches(0.2), Inches(1.85), Inches(1.95), Inches(0.02), LIGHT_GRAY)
    add_text(slide, left+Inches(0.2), Inches(2.0), Inches(1.95), Inches(0.6), desc, size=7.5, color=NAVY2, line_spacing=1.25)
    add_rect(slide, left+Inches(0.2), Inches(2.7), Inches(1.95), Inches(0.45), BEIGE)
    add_text(slide, left+Inches(0.2), Inches(2.78), Inches(1.95), Inches(0.3), price, size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)

# bottom big highlight bar
add_rect(slide, Inches(5.1), Inches(4.15), Inches(7.7), Inches(1.1), WHITE, radius=True)
add_rect(slide, Inches(5.1), Inches(4.15), Inches(0.08), Inches(1.1), GOLD)
add_text(slide, Inches(5.4), Inches(4.35), Inches(1.2), Inches(0.25), "ВАШЕ ПРЕИМУЩЕСТВО", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(5.4), Inches(4.6), Inches(3.5), Inches(0.35), "Вы не инвестируете в артистов.", size=10, color=NAVY, bold=True)
add_text(slide, Inches(5.4), Inches(4.95), Inches(3.5), Inches(0.3), "Только операционка клуба, которая окупается баром.", size=7.5, color=MUTED)
add_rect(slide, Inches(9.0), Inches(4.3), Inches(3.7), Inches(0.8), NAVY)
add_text(slide, Inches(9.1), Inches(4.4), Inches(3.5), Inches(0.2), "ФИКСИРОВАННЫЙ ДОХОД ДО СТАРТА ВЕЧЕРА", size=6, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(9.1), Inches(4.65), Inches(3.5), Inches(0.3), "330 000 ₽  живыми деньгами", size=11, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# timeline
add_text(slide, Inches(5.1), Inches(5.6), Inches(7.7), Inches(0.25), "ТАЙМЛАЙН ПРОДАЖ БИЛЕТОВ", size=7, color=MUTED, bold=True)
for i, (d,p,c) in enumerate([("01–07 НОЯБРЯ","3 000 ₽","Старт"),("08–14 НОЯБРЯ","3 500 ₽","Рост"),("С 15 НОЯБРЯ","4 000 ₽","Финал")] ):
    left = Inches(5.1 + i*2.58)
    add_rect(slide, left, Inches(5.9), Inches(2.45), Inches(0.75), WHITE)
    if i==2:
        add_rect(slide, left, Inches(5.9), Inches(2.45), Inches(0.75), NAVY)
        add_text(slide, left+Inches(0.15), Inches(5.98), Inches(2.15), Inches(0.15), d, size=6, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, left+Inches(0.15), Inches(6.18), Inches(2.15), Inches(0.2), p, size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, left+Inches(0.15), Inches(6.42), Inches(2.15), Inches(0.15), c, size=6, color=BEIGE, align=PP_ALIGN.CENTER)
    else:
        add_text(slide, left+Inches(0.15), Inches(5.98), Inches(2.15), Inches(0.15), d, size=6, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, left+Inches(0.15), Inches(6.18), Inches(2.15), Inches(0.2), p, size=10, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, left+Inches(0.15), Inches(6.42), Inches(2.15), Inches(0.15), c, size=6, color=MUTED, align=PP_ALIGN.CENTER)
    if i<2:
        add_text(slide, left+Inches(2.45), Inches(6.15), Inches(0.2), Inches(0.2), "›", size=11, color=GOLD, align=PP_ALIGN.CENTER)


# ========== SLIDE 3 - ЗАТРАТЫ FOSTER ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(2.5), Inches(0.2), "02  —  ЗАТРАТЫ FOSTER", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(6), Inches(0.5), "Ваши затраты минимальны", size=24, color=NAVY, bold=True)
add_text(slide, Inches(0.7), Inches(1.15), Inches(6), Inches(0.3), "Площадка и оборудование — бесплатно. Только операционка.", size=9, color=MUTED)

# left table card
card = add_rect(slide, Inches(0.7), Inches(1.65), Inches(6.2), Inches(4.9), WHITE, line=LIGHT_GRAY)
# header
add_rect(slide, Inches(0.7), Inches(1.65), Inches(6.2), Inches(0.55), NAVY)
add_text(slide, Inches(0.9), Inches(1.78), Inches(3.5), Inches(0.3), "СТАТЬЯ РАСХОДОВ", size=7, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(5.0), Inches(1.78), Inches(1.5), Inches(0.3), "СУММА", size=7, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.RIGHT)
rows = [
    ("Персонал клуба", "официанты, бармены, повара, менеджмент", "165 470 ₽"),
    ("Бар", "фудкост 25% от выручки 368 025 ₽", "92 006 ₽"),
    ("Кухня", "фудкост 27% от выручки 100 000 ₽", "27 000 ₽"),
    ("Скидка на проживание", "отменена по договорённости", "0 ₽"),
]
for i, (t1,t2,price) in enumerate(rows):
    y = Inches(2.2 + i*0.85)
    if i%2==0:
        add_rect(slide, Inches(0.7), y, Inches(6.2), Inches(0.85), BEIGE)
    # dot
    add_circle(slide, Inches(0.9), y+Inches(0.32), Inches(0.08), GOLD if i<3 else MUTED)
    add_text(slide, Inches(1.15), y+Inches(0.12), Inches(3.2), Inches(0.2), t1, size=8, color=NAVY, bold=True)
    add_text(slide, Inches(1.15), y+Inches(0.33), Inches(3.2), Inches(0.2), t2, size=6.5, color=MUTED)
    add_text(slide, Inches(4.7), y+Inches(0.22), Inches(1.65), Inches(0.3), price, size=10, color=NAVY, bold=True, align=PP_ALIGN.RIGHT)
    if i<3:
        add_rect(slide, Inches(0.9), y+Inches(0.72), Inches(5.8), Inches(0.01), LIGHT_GRAY)
# total row
add_rect(slide, Inches(0.7), Inches(5.65), Inches(6.2), Inches(0.9), NAVY)
add_text(slide, Inches(0.9), Inches(5.85), Inches(3.0), Inches(0.2), "ИТОГО ЗАТРАТЫ КЛУБА", size=8, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(0.9), Inches(6.08), Inches(3.0), Inches(0.15), "полностью покрываются доходами", size=6.5, color=BEIGE)
add_text(slide, Inches(4.3), Inches(5.85), Inches(2.3), Inches(0.4), "284 476 ₽", size=18, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)

# right side - visual cards
# big profit teaser
add_rect(slide, Inches(7.2), Inches(1.65), Inches(5.4), Inches(2.1), BEIGE, radius=True)
add_text(slide, Inches(7.5), Inches(1.85), Inches(4.8), Inches(0.2), "ЧТО ЭТО ЗНАЧИТ НА ПРАКТИКЕ", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(7.5), Inches(2.1), Inches(4.8), Inches(0.35), "Вы не платите ни за аренду, ни за артистов.", size=10, color=NAVY, bold=True)
add_text(slide, Inches(7.5), Inches(2.5), Inches(4.8), Inches(0.5), "Ваши 284 тыс. ₽ — это только себестоимость бара и кухни, которая уже заложена в выручку. Фактически — нулевые инвестиции.", size=7.5, color=MUTED, line_spacing=1.3)
# 3 mini stats
for i, (num, label, sub) in enumerate([("0 ₽","Аренда","площадка ваша"),("0 ₽","Гонорар","платит Кактуса"),("0 ₽","Оборудование","уже на месте")] ):
    left = Inches(7.2 + i*1.85)
    add_rect(slide, left, Inches(3.95), Inches(1.65), Inches(1.0), WHITE, line=LIGHT_GRAY)
    add_text(slide, left+Inches(0.15), Inches(4.08), Inches(1.35), Inches(0.25), num, size=12, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.15), Inches(4.38), Inches(1.35), Inches(0.15), label, size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.15), Inches(4.58), Inches(1.35), Inches(0.15), sub, size=6, color=MUTED, align=PP_ALIGN.CENTER)
# comparison bar
add_rect(slide, Inches(7.2), Inches(5.25), Inches(5.4), Inches(1.3), NAVY2, radius=True)
add_text(slide, Inches(7.5), Inches(5.4), Inches(4.8), Inches(0.2), "СРАВНЕНИЕ МАСШТАБА", size=6, color=GOLD_LIGHT, bold=True)
# bar 1
add_text(slide, Inches(7.5), Inches(5.7), Inches(1.8), Inches(0.15), "ВАШИ ЗАТРАТЫ", size=6, color=BEIGE, bold=True)
add_rect(slide, Inches(7.5), Inches(5.92), Inches(1.8), Inches(0.28), GOLD)
add_text(slide, Inches(7.6), Inches(5.95), Inches(1.6), Inches(0.2), "284 тыс.", size=7, color=NAVY, bold=True)
# bar 2
add_text(slide, Inches(7.5), Inches(6.25), Inches(1.8), Inches(0.15), "ЗАТРАТЫ КАКТУСЫ", size=6, color=BEIGE, bold=True)
add_rect(slide, Inches(7.5), Inches(6.45), Inches(4.8), Inches(0.28), WHITE)
add_text(slide, Inches(7.6), Inches(6.48), Inches(2.0), Inches(0.2), "1 120 тыс.  •  в 4 раза больше", size=7, color=NAVY, bold=True)
add_text(slide, Inches(7.2), Inches(6.85), Inches(5.4), Inches(0.2), "Весь финансовый риск — на организаторе, не на вас.", size=6.5, color=GOLD_LIGHT, align=PP_ALIGN.CENTER, italic=True)


# ========== SLIDE 4 - ДОХОД FOSTER ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, NAVY)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(3), Inches(0.2), "03  —  ЧТО ЗАРАБОТАЕТ FOSTER", size=7, color=GOLD, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(7), Inches(0.5), "Ваша чистая прибыль — 483 549 ₽", size=24, color=WHITE, bold=True)
add_rect(slide, Inches(0.7), Inches(1.15), Inches(0.8), Inches(0.04), GOLD)
add_text(slide, Inches(1.7), Inches(1.12), Inches(5), Inches(0.2), "Рентабельность ~63% за один вечер", size=9, color=BEIGE)

# left revenue table
add_rect(slide, Inches(0.7), Inches(1.55), Inches(6.0), Inches(5.3), WHITE, radius=True)
add_rect(slide, Inches(0.7), Inches(1.55), Inches(6.0), Inches(0.5), BEIGE)
add_text(slide, Inches(0.9), Inches(1.67), Inches(3), Inches(0.2), "СТАТЬЯ ДОХОДА", size=7, color=NAVY, bold=True)
add_text(slide, Inches(4.6), Inches(1.67), Inches(1.7), Inches(0.2), "СУММА", size=7, color=NAVY, bold=True, align=PP_ALIGN.RIGHT)
revenues = [
    ("Бар", "валовая выручка", "368 025 ₽", False),
    ("Кухня", "валовая выручка", "100 000 ₽", False),
    ("Коктейли Welcome", "300 × 600 ₽ компенсация", "180 000 ₽", True),
    ("Депозиты столов", "4×5 000 + 3×10 000", "50 000 ₽", True),
    ("VIP №1", "депозит", "40 000 ₽", True),
    ("VIP №2", "депозит", "30 000 ₽", True),
]
for i,(t1,t2,price,highlight) in enumerate(revenues):
    y = Inches(2.05 + i*0.55)
    if highlight:
        add_rect(slide, Inches(0.7), y, Inches(6.0), Inches(0.55), RGBColor(0xFF,0xF7,0xE0))
        add_rect(slide, Inches(0.7), y, Inches(0.06), Inches(0.55), GOLD)
    add_text(slide, Inches(0.9), y+Inches(0.08), Inches(2.2), Inches(0.2), t1, size=8, color=NAVY, bold=True)
    add_text(slide, Inches(2.6), y+Inches(0.1), Inches(1.8), Inches(0.2), t2, size=6.5, color=MUTED)
    add_text(slide, Inches(4.6), y+Inches(0.08), Inches(1.7), Inches(0.25), price, size=9, color=NAVY if not highlight else NAVY, bold=True, align=PP_ALIGN.RIGHT)
    if highlight:
        add_circle(slide, Inches(6.35), y+Inches(0.18), Inches(0.12), GOLD)
        add_text(slide, Inches(6.35), y+Inches(0.16), Inches(0.12), Inches(0.15), "✓", size=6, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
# total revenue
add_rect(slide, Inches(0.7), Inches(5.45), Inches(6.0), Inches(0.7), NAVY2)
add_text(slide, Inches(0.9), Inches(5.58), Inches(3.0), Inches(0.2), "ИТОГО ВЫРУЧКА", size=7, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(0.9), Inches(5.78), Inches(3.0), Inches(0.15), "+ кальяны от Макао по факту", size=6, color=BEIGE)
add_text(slide, Inches(4.2), Inches(5.6), Inches(2.1), Inches(0.3), "768 025 ₽", size=14, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
add_text(slide, Inches(4.2), Inches(5.92), Inches(2.1), Inches(0.15), "+ КАЛЬЯНЫ", size=6, color=GOLD_LIGHT, align=PP_ALIGN.RIGHT)

# right side - profit visualization
add_text(slide, Inches(7.1), Inches(1.55), Inches(5.5), Inches(0.25), "ФИНАНСОВЫЙ РЕЗУЛЬТАТ ЗА ВЕЧЕР", size=7, color=GOLD, bold=True)
# big profit card
add_rect(slide, Inches(7.1), Inches(1.95), Inches(5.5), Inches(2.15), WHITE, radius=True)
add_rect(slide, Inches(7.1), Inches(1.95), Inches(5.5), Inches(0.06), GOLD)
add_text(slide, Inches(7.4), Inches(2.15), Inches(2.5), Inches(0.2), "ЧИСТАЯ ПРИБЫЛЬ", size=7, color=MUTED, bold=True)
add_text(slide, Inches(7.4), Inches(2.4), Inches(3.0), Inches(0.4), "483 549 ₽", size=22, color=NAVY, bold=True)
add_text(slide, Inches(7.4), Inches(2.85), Inches(3.0), Inches(0.2), "768 025 – 284 476 = ваш плюс", size=7, color=MUTED)
# badge 63%
add_rect(slide, Inches(10.5), Inches(2.15), Inches(1.7), Inches(0.9), NAVY)
add_text(slide, Inches(10.6), Inches(2.25), Inches(1.5), Inches(0.25), "~63%", size=16, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(10.6), Inches(2.6), Inches(1.5), Inches(0.15), "РЕНТАБЕЛЬНОСТЬ", size=6, color=WHITE, align=PP_ALIGN.CENTER, bold=True)

# waterfall like bar
add_text(slide, Inches(7.4), Inches(3.25), Inches(4.9), Inches(0.15), "КАК ФОРМИРУЕТСЯ ПРИБЫЛЬ", size=6, color=MUTED, bold=True)
# draw horizontal bars
# revenue bar full
add_rect(slide, Inches(7.4), Inches(3.5), Inches(4.9), Inches(0.32), BEIGE)
add_rect(slide, Inches(7.4), Inches(3.5), Inches(4.9), Inches(0.32), RGBColor(0xC9,0xA9,0x6A))
add_text(slide, Inches(7.5), Inches(3.53), Inches(4.7), Inches(0.25), "ВЫРУЧКА  768 025 ₽", size=7, color=NAVY, bold=True)
# costs bar shorter
add_rect(slide, Inches(7.4), Inches(3.9), Inches(1.8), Inches(0.28), WHITE, line=LIGHT_GRAY)
add_text(slide, Inches(7.5), Inches(3.93), Inches(1.6), Inches(0.2), "РАСХОДЫ 284 476 ₽", size=6.5, color=MUTED, bold=True)
# arrow down
add_text(slide, Inches(7.4), Inches(4.22), Inches(4.9), Inches(0.2), "▼  ЧИСТАЯ ПРИБЫЛЬ 483 549 ₽  •  остается у вас", size=7, color=GOLD, bold=True)

# guaranteed block
add_rect(slide, Inches(7.1), Inches(4.65), Inches(5.5), Inches(1.35), RGBColor(0x11,0x28,0x45), radius=True)
add_rect(slide, Inches(7.1), Inches(4.65), Inches(5.5), Inches(0.04), GOLD)
add_text(slide, Inches(7.4), Inches(4.85), Inches(5.0), Inches(0.2), "⚡  ГАРАНТИРОВАННЫЕ ДЕНЬГИ ЕЩЁ ДО БАРА", size=7, color=GOLD, bold=True)
add_text(slide, Inches(7.4), Inches(5.1), Inches(3.0), Inches(0.6), "330 000 ₽ живыми деньгами:\nкоктейли + столы + VIP — не зависят от явки при правильном договоре.", size=7.5, color=BEIGE, line_spacing=1.2)
add_rect(slide, Inches(10.2), Inches(4.95), Inches(2.0), Inches(0.7), GOLD)
add_text(slide, Inches(10.3), Inches(5.05), Inches(1.8), Inches(0.15), "180К + 50К + 70К", size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(10.3), Inches(5.28), Inches(1.8), Inches(0.2), "330 000 ₽", size=11, color=NAVY, bold=True, align=PP_ALIGN.CENTER)

add_rect(slide, Inches(7.1), Inches(6.2), Inches(5.5), Inches(0.65), BEIGE)
add_text(slide, Inches(7.3), Inches(6.3), Inches(0.35), Inches(0.4), "💡", size=12, align=PP_ALIGN.CENTER)
add_text(slide, Inches(7.75), Inches(6.3), Inches(4.8), Inches(0.4), "Даже при 200 гостях вы в плюсе — бар и кухня уже покрывают все расходы.", size=7.5, color=NAVY, bold=True)


# ========== SLIDE 5 - HEADLINERS ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(3), Inches(0.2), "04  —  ХЕДЛАЙНЕРЫ", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(7), Inches(0.5), "Имена, которые продают", size=24, color=NAVY, bold=True)
add_text(slide, Inches(0.7), Inches(1.15), Inches(7), Inches(0.25), "Вирусный хит без бюджета • Первый приезд в Крым • Ожидание с лета", size=9, color=MUTED)

# left card - Gleb
add_rect(slide, Inches(0.7), Inches(1.6), Inches(6.2), Inches(3.1), BEIGE, radius=True)
# image side
add_rect(slide, Inches(0.7), Inches(1.6), Inches(2.4), Inches(3.1), NAVY)
add_circle(slide, Inches(1.0), Inches(1.95), Inches(1.55), RGBColor(0x1E,0x3D,0x5E))
add_text(slide, Inches(0.85), Inches(2.45), Inches(2.1), Inches(0.25), "G F", size=18, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(0.8), Inches(3.7), Inches(2.2), Inches(0.15), "ПРОДЮСЕР  •  MALLORCA", size=6, color=GOLD_LIGHT, align=PP_ALIGN.CENTER, bold=True)
add_text(slide, Inches(0.85), Inches(3.95), Inches(2.1), Inches(0.2), "Фото: промо артиста", size=6, color=MUTED, align=PP_ALIGN.CENTER, italic=True)
# text side
add_text(slide, Inches(3.3), Inches(1.8), Inches(3.4), Inches(0.25), "GLEB FILIPCHENKOW", size=11, color=NAVY, bold=True)
add_rect(slide, Inches(3.3), Inches(2.08), Inches(0.6), Inches(0.03), GOLD)
add_text(slide, Inches(3.3), Inches(2.2), Inches(3.4), Inches(0.9), "Соавтор вирусного трека «Mallorca» — лето 2025. Взрыв весной 2026 без бюджета: трек звучал на всех московских вечеринках, видео — сотни тысяч просмотров.", size=7.5, color=NAVY2, line_spacing=1.25)
add_text(slide, Inches(3.3), Inches(3.15), Inches(3.4), Inches(0.6), "Дальше — релиз с Ольгой Серябкиной. Карьера на взлёте, интерес к артисту максимальный.", size=7, color=MUTED)
add_rect(slide, Inches(3.3), Inches(3.85), Inches(3.4), Inches(0.45), WHITE)
add_text(slide, Inches(3.4), Inches(3.92), Inches(3.2), Inches(0.3), "✦  Трек Mallorca — главный хит вечеринок Москвы 2026", size=6.5, color=NAVY, bold=True)

# right card - Noemy
add_rect(slide, Inches(7.1), Inches(1.6), Inches(5.5), Inches(3.1), WHITE, line=LIGHT_GRAY, radius=True)
add_rect(slide, Inches(7.1), Inches(1.6), Inches(2.0), Inches(3.1), NAVY2)
add_circle(slide, Inches(7.35), Inches(1.95), Inches(1.45), NAVY)
add_text(slide, Inches(7.35), Inches(2.45), Inches(1.45), Inches(0.25), "N A", size=16, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(7.2), Inches(3.7), Inches(1.8), Inches(0.15), "ВОКАЛ  •  GAROTA / HOJE", size=6, color=GOLD_LIGHT, align=PP_ALIGN.CENTER, bold=True)
add_text(slide, Inches(9.3), Inches(1.8), Inches(2.9), Inches(0.25), "NOEMY ABRANTES", size=11, color=NAVY, bold=True)
add_rect(slide, Inches(9.3), Inches(2.08), Inches(0.6), Inches(0.03), GOLD)
add_text(slide, Inches(9.3), Inches(2.2), Inches(2.9), Inches(0.9), "Соавтор релизов «Garota» и «HOJE». Яркий вокал, европейское звучание, идеально для ночной атмосферы Foster.", size=7.5, color=NAVY2, line_spacing=1.25)
add_rect(slide, Inches(9.3), Inches(3.2), Inches(2.9), Inches(0.65), BEIGE)
add_text(slide, Inches(9.45), Inches(3.3), Inches(2.6), Inches(0.15), "ПОЧЕМУ ДУЭТ РАБОТАЕТ", size=6, color=GOLD_DARK, bold=True)
add_text(slide, Inches(9.45), Inches(3.5), Inches(2.6), Inches(0.3), "Мужской продюсерский грув + женский вокал = полный танцпол", size=6.5, color=NAVY)

# bottom - why event
add_text(slide, Inches(0.7), Inches(5.0), Inches(12), Inches(0.25), "ПОЧЕМУ ЭТО СОБЫТИЕ ПРОДАСТСЯ", size=7, color=NAVY, bold=True)
# 3 reasons
reasons = [
    ("01", "ПЕРВЫЙ ПРИВОЗ", "В Крыму не выступали\nникогда — эффект\nэксклюзива", GOLD),
    ("02", "ОТЛОЖЕННЫЙ СПРОС", "Летний привоз сорвался\nиз-за топливного кризиса\nпублика ждёт", NAVY),
    ("03", "ОКНО В ГРАФИКЕ", "Ноябрь расписан,\n5 декабря — единственная\nсвободная дата", NAVY2),
    ("04", "ВИРАЛЬНОСТЬ", "Сотни тысяч просмотров\nбез вложений\nсарафан уже работает", GOLD_DARK),
]
for i,(num,title,desc,col) in enumerate(reasons):
    left = Inches(0.7 + i*3.15)
    add_rect(slide, left, Inches(5.3), Inches(2.95), Inches(1.5), WHITE, line=LIGHT_GRAY)
    add_text(slide, left+Inches(0.15), Inches(5.4), Inches(0.6), Inches(0.25), num, size=14, color=col, bold=True)
    add_text(slide, left+Inches(0.75), Inches(5.48), Inches(2.0), Inches(0.2), title, size=7, color=NAVY, bold=True)
    add_rect(slide, left+Inches(0.15), Inches(5.75), Inches(2.65), Inches(0.02), LIGHT_GRAY)
    add_text(slide, left+Inches(0.15), Inches(5.88), Inches(2.65), Inches(0.6), desc, size=7, color=MUTED, line_spacing=1.2)
# urgency bar
add_rect(slide, Inches(0.7), Inches(7.0), Inches(12.0), Inches(0.5), NAVY)
add_text(slide, Inches(0.9), Inches(7.12), Inches(6), Inches(0.25), "⚡  ИМЯ ПРОДАЁТ САМО СОБЫТИЕ — ВАМ ОСТАЁТСЯ ТОЛЬКО ПРИНЯТЬ ГОСТЕЙ", size=7.5, color=GOLD_LIGHT, bold=True)
add_rect(slide, Inches(9.5), Inches(7.08), Inches(3.1), Inches(0.32), GOLD)
add_text(slide, Inches(9.55), Inches(7.12), Inches(3.0), Inches(0.22), "350 БИЛЕТОВ  •  АЖИОТАЖ ОБЕСПЕЧЕН", size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)


# ========== SLIDE 6 - ЗАТРАТЫ КАКТУСЫ ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, BEIGE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(3), Inches(0.2), "05  —  ЗОНА ОТВЕТСТВЕННОСТИ КАКТУСЫ", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(7), Inches(0.5), "Организатор берёт весь риск", size=24, color=NAVY, bold=True)
add_text(slide, Inches(0.7), Inches(1.15), Inches(7), Inches(0.25), "Вы не участвуете в этих расходах — они мотивируют Кактусу продавать", size=9, color=MUTED)

# table
rows = [
    ("Гонорар хедлайнеров", "550 000 ₽", "фикс", NAVY),
    ("Перелёт / поезд / трансфер", "~ 80 000 ₽", "допущение", MUTED),
    ("Технический райдер", "~ 30 000 ₽", "сверх клуба", MUTED),
    ("Местные артисты", "75 000 ₽", "среднее", MUTED),
    ("Видеосъёмка и монтаж", "30 000 ₽", "", MUTED),
    ("Фотоотчёт", "25 000 ₽", "", MUTED),
    ("Продвижение, блогеры, афиши", "~ 100 000 ₽", "допущение", MUTED),
    ("Внешняя охрана", "~ 30 000 ₽", "сверх шаттла", MUTED),
    ("Компенсация коктейлей", "180 000 ₽", "→ вам", GOLD),
    ("Оплата шаттла", "20 040 ₽", "→ вам", GOLD),
]
card = add_rect(slide, Inches(0.7), Inches(1.55), Inches(7.0), Inches(4.85), WHITE, line=LIGHT_GRAY, radius=True)
add_rect(slide, Inches(0.7), Inches(1.55), Inches(7.0), Inches(0.5), NAVY)
add_text(slide, Inches(0.9), Inches(1.68), Inches(3.5), Inches(0.2), "СТАТЬЯ РАСХОДОВ", size=7, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(5.2), Inches(1.68), Inches(1.5), Inches(0.2), "СУММА", size=7, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.RIGHT)
for i,(name,price,note,col) in enumerate(rows):
    y = Inches(2.05 + i*0.41)
    if i in [8,9]:
        add_rect(slide, Inches(0.7), y, Inches(7.0), Inches(0.41), RGBColor(0xFF,0xF7,0xE0))
        add_rect(slide, Inches(0.7), y, Inches(0.06), Inches(0.41), GOLD)
    elif i%2==0:
        add_rect(slide, Inches(0.7), y, Inches(7.0), Inches(0.41), RGBColor(0xF9,0xF6,0xF0))
    add_text(slide, Inches(0.95), y+Inches(0.08), Inches(3.5), Inches(0.25), name, size=7, color=NAVY, bold=True if i in [0,8,9] else False)
    if note:
        add_text(slide, Inches(4.2), y+Inches(0.08), Inches(1.2), Inches(0.25), note, size=6, color=MUTED, align=PP_ALIGN.RIGHT, italic=True)
    add_text(slide, Inches(5.5), y+Inches(0.08), Inches(1.3), Inches(0.25), price, size=7.5, color=col, bold=True, align=PP_ALIGN.RIGHT)

add_rect(slide, Inches(0.7), Inches(6.45), Inches(7.0), Inches(0.65), NAVY)
add_text(slide, Inches(0.9), Inches(6.58), Inches(3.5), Inches(0.15), "ИТОГО ИНВЕСТИЦИИ КАКТУСЫ", size=7, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(0.9), Inches(6.78), Inches(3.5), Inches(0.15), "все риски на организаторе", size=6, color=BEIGE)
add_text(slide, Inches(5.0), Inches(6.58), Inches(1.9), Inches(0.35), "~ 1 120 040 ₽", size=14, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)

# right side - insight
add_rect(slide, Inches(8.0), Inches(1.55), Inches(4.6), Inches(2.3), NAVY, radius=True)
add_text(slide, Inches(8.3), Inches(1.75), Inches(4.0), Inches(0.2), "ЧТО ЭТО ЗНАЧИТ ДЛЯ ВАС", size=7, color=GOLD, bold=True)
add_text(slide, Inches(8.3), Inches(2.0), Inches(4.0), Inches(0.6), "Кактуса вложит\nбольше миллиона", size=16, color=WHITE, bold=True, line_spacing=0.95)
add_rect(slide, Inches(8.3), Inches(2.7), Inches(0.6), Inches(0.03), GOLD)
add_text(slide, Inches(8.3), Inches(2.85), Inches(4.0), Inches(0.6), "Чтобы окупиться, им нужно продать\nкаждый билет. Они будут рвать рынок —\nа вы просто получаете гостей.", size=7.5, color=BEIGE, line_spacing=1.25)

add_rect(slide, Inches(8.0), Inches(4.05), Inches(4.6), Inches(1.05), WHITE, radius=True)
add_text(slide, Inches(8.3), Inches(4.2), Inches(1.0), Inches(0.7), "💰", size=18, align=PP_ALIGN.CENTER)
add_text(slide, Inches(9.3), Inches(4.2), Inches(2.8), Inches(0.2), "ВАШИ 200 040 ₽ ГАРАНТИРОВАНЫ", size=7, color=NAVY, bold=True)
add_text(slide, Inches(9.3), Inches(4.45), Inches(2.8), Inches(0.35), "180 000 ₽ за коктейли + 20 040 ₽ за шаттл — фикс в договоре", size=7, color=MUTED)
add_rect(slide, Inches(8.3), Inches(4.85), Inches(4.0), Inches(0.04), LIGHT_GRAY)
add_text(slide, Inches(8.3), Inches(4.95), Inches(4.0), Inches(0.15), "Независимо от продаж билетов — это ваши деньги", size=6.5, color=GOLD_DARK, italic=True)

add_rect(slide, Inches(8.0), Inches(5.3), Inches(4.6), Inches(1.35), BEIGE, radius=True)
add_text(slide, Inches(8.3), Inches(5.5), Inches(4.0), Inches(0.2), "СРАВНЕНИЕ РИСКОВ", size=7, color=GOLD_DARK, bold=True)
# Foster risk bar small
add_text(slide, Inches(8.3), Inches(5.8), Inches(1.6), Inches(0.15), "FOSTER", size=6, color=NAVY, bold=True)
add_rect(slide, Inches(8.3), Inches(6.0), Inches(1.2), Inches(0.22), NAVY)
add_text(slide, Inches(8.35), Inches(6.03), Inches(1.1), Inches(0.15), "284К", size=6, color=WHITE, bold=True)
add_text(slide, Inches(9.8), Inches(6.0), Inches(2.2), Inches(0.22), "← минимальный", size=6, color=MUTED)
# Kaktus bar large
add_text(slide, Inches(8.3), Inches(6.28), Inches(1.6), Inches(0.15), "КАКТУСА", size=6, color=NAVY, bold=True)
add_rect(slide, Inches(8.3), Inches(6.45), Inches(3.8), Inches(0.22), GOLD)
add_text(slide, Inches(8.35), Inches(6.48), Inches(2.0), Inches(0.15), "1 120К  •  ×4", size=6, color=NAVY, bold=True)


# ========== SLIDE 7 - ПОТЕНЦИАЛЬНЫЙ ДОХОД КАКТУСЫ ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(4), Inches(0.2), "06  —  ПОТЕНЦИАЛ ОРГАНИЗАТОРА", size=7, color=GOLD_DARK, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(8), Inches(0.5), "Кактуса в зоне риска — ваша защита", size=22, color=NAVY, bold=True)
add_text(slide, Inches(0.7), Inches(1.15), Inches(8), Inches(0.25), "Они выходят в плюс только при SOLD OUT. Значит — будут продавать до последнего билета.", size=8, color=MUTED)

# 3 scenario cards
scenarios = [
    ("ПЕССИМИСТИЧНЫЙ", "250 чел.", "3 300 ₽", "825 000 ₽", "– 295 040 ₽", "Убыток", RGBColor(0xE0,0xE0,0xE0), MUTED, NAVY),
    ("БАЗОВЫЙ", "300 чел.", "3 400 ₽", "1 020 000 ₽", "– 100 040 ₽", "Убыток", RGBColor(0xED,0xE6,0xD6), NAVY2, NAVY),
    ("ОПТИМИСТИЧНЫЙ", "350 чел.", "3 400 ₽", "1 190 000 ₽", "+ 69 960 ₽", "Прибыль • SOLD OUT", GOLD, WHITE, WHITE),
]
for i,(title, guests, avg, rev, result, status, bg, col_text, col_bg) in enumerate(scenarios):
    left = Inches(0.7 + i*4.2)
    # card
    is_best = i==2
    card_col = NAVY if is_best else WHITE
    border = GOLD if is_best else LIGHT_GRAY
    add_rect(slide, left, Inches(1.55), Inches(3.95), Inches(3.8), card_col, line=border, radius=True)
    if is_best:
        add_rect(slide, left, Inches(1.55), Inches(3.95), Inches(0.5), GOLD)
        add_text(slide, left+Inches(0.2), Inches(1.62), Inches(3.5), Inches(0.35), "✦  SOLD OUT  •  ЦЕЛЬ", size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, left+Inches(0.2), Inches(2.1), Inches(3.5), Inches(0.2), title, size=7, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
    else:
        add_text(slide, left+Inches(0.2), Inches(1.75), Inches(3.5), Inches(0.2), title, size=7, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    # guests large
    add_text(slide, left+Inches(0.2), Inches(2.4), Inches(3.5), Inches(0.35), guests, size=18, color=NAVY if not is_best else WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.2), Inches(2.78), Inches(3.5), Inches(0.15), f"средняя цена {avg}", size=6, color=MUTED if not is_best else BEIGE, align=PP_ALIGN.CENTER)
    add_rect(slide, left+Inches(0.5), Inches(3.05), Inches(2.95), Inches(0.02), LIGHT_GRAY if not is_best else RGBColor(0x2A,0x4A,0x6E))
    # revenue
    add_text(slide, left+Inches(0.2), Inches(3.2), Inches(1.7), Inches(0.15), "ВЫРУЧКА", size=6, color=MUTED if not is_best else GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(2.0), Inches(3.2), Inches(1.7), Inches(0.15), "РЕЗУЛЬТАТ", size=6, color=MUTED if not is_best else GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.2), Inches(3.4), Inches(1.7), Inches(0.25), rev, size=9, color=NAVY if not is_best else WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(2.0), Inches(3.4), Inches(1.7), Inches(0.25), result, size=9, color=RGBColor(0xC0,0x39,0x2B) if "–" in result else (RGBColor(0x1E,0x8A,0x49) if is_best else RGBColor(0xC0,0x39,0x2B)), bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.2), Inches(3.75), Inches(3.5), Inches(0.2), status, size=6, color=RGBColor(0xC0,0x39,0x2B) if "Убыток" in status else RGBColor(0x1E,0x8A,0x49), bold=True, align=PP_ALIGN.CENTER, italic=True)
    # bar visualization
    bar_w = [Inches(1.8), Inches(2.4), Inches(3.0)][i]
    add_rect(slide, left+Inches(0.5), Inches(4.2), bar_w, Inches(0.28), LIGHT_GRAY if not is_best else GOLD)
    add_rect(slide, left+Inches(0.5), Inches(4.2), Inches(2.7), Inches(0.28), RGBColor(0xFF,0x00,0x00) if False else (GOLD if is_best else RGBColor(0xC9,0xA9,0x6A))) # just gold
    # threshold line
    add_rect(slide, left+Inches(0.5)+Inches(2.7), Inches(4.0), Inches(0.02), Inches(0.65), RGBColor(0xC0,0x39,0x2B))
    add_text(slide, left+Inches(0.5)+Inches(2.4), Inches(4.55), Inches(0.6), Inches(0.15), "1.12М", size=5, color=RGBColor(0xC0,0x39,0x2B), align=PP_ALIGN.CENTER)

add_rect(slide, Inches(0.7), Inches(5.6), Inches(12.0), Inches(0.9), BEIGE, radius=True)
add_text(slide, Inches(1.0), Inches(5.75), Inches(0.5), Inches(0.6), "💡", size=16, align=PP_ALIGN.CENTER)
add_text(slide, Inches(1.6), Inches(5.75), Inches(7.0), Inches(0.25), "ВЫВОД ДЛЯ ВАС: КАКТУСА МАКСИМАЛЬНО МОТИВИРОВАНА", size=8, color=NAVY, bold=True)
add_text(slide, Inches(1.6), Inches(6.0), Inches(7.0), Inches(0.3), "Их прибыль только при 350 гостях. Они будут продавать агрессивно — а ваши 330 000 ₽ зафиксированы в договоре и не зависят от их результата.", size=7, color=MUTED, line_spacing=1.2)
add_rect(slide, Inches(9.2), Inches(5.8), Inches(3.1), Inches(0.5), NAVY)
add_text(slide, Inches(9.3), Inches(5.9), Inches(2.9), Inches(0.3), "ИХ РИСК  →  ВАША ВЫГОДА", size=7, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)


# ========== SLIDE 8 - РЕЗЮМЕ ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, NAVY)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
add_text(slide, Inches(0.7), Inches(0.35), Inches(3), Inches(0.2), "07  —  РЕЗЮМЕ", size=7, color=GOLD, bold=True)
add_text(slide, Inches(0.7), Inches(0.6), Inches(7), Inches(0.5), "Ваша сделка на 483 500 ₽", size=24, color=WHITE, bold=True)
add_text(slide, Inches(0.7), Inches(1.15), Inches(7), Inches(0.25), "Минимальный риск • Максимальная отдача за один вечер", size=9, color=BEIGE)

# 3 benefits large
benefits = [
    ("483 549 ₽", "ЧИСТАЯ ПРИБЫЛЬ", "за вечер + кальяны", "63% рентабельность"),
    ("330 000 ₽", "ГАРАНТИРОВАНО", "до старта вечеринки", "фикс в договоре"),
    ("0 ₽", "ВАШИ РИСКИ", "не платите за артистов", "только операционка"),
]
for i,(num,label,sub,foot) in enumerate(benefits):
    left = Inches(0.7 + i*2.5)
    add_rect(slide, left, Inches(1.6), Inches(2.3), Inches(1.6), WHITE, radius=True)
    add_rect(slide, left, Inches(1.6), Inches(2.3), Inches(0.05), GOLD)
    add_text(slide, left+Inches(0.15), Inches(1.75), Inches(2.0), Inches(0.35), num, size=16, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.15), Inches(2.12), Inches(2.0), Inches(0.15), label, size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, left+Inches(0.15), Inches(2.32), Inches(2.0), Inches(0.15), sub, size=6, color=MUTED, align=PP_ALIGN.CENTER)
    add_rect(slide, left+Inches(0.15), Inches(2.55), Inches(2.0), Inches(0.02), LIGHT_GRAY)
    add_text(slide, left+Inches(0.15), Inches(2.65), Inches(2.0), Inches(0.2), foot, size=6, color=GOLD_DARK, bold=True, align=PP_ALIGN.CENTER)

# checklist card
add_rect(slide, Inches(0.7), Inches(3.45), Inches(7.8), Inches(3.6), WHITE, radius=True)
add_rect(slide, Inches(0.7), Inches(3.45), Inches(7.8), Inches(0.5), BEIGE)
add_text(slide, Inches(0.9), Inches(3.58), Inches(4), Inches(0.2), "✓  ЧТО ЗАФИКСИРОВАТЬ В ДОГОВОРЕ", size=8, color=NAVY, bold=True)
add_text(slide, Inches(5.2), Inches(3.58), Inches(2.9), Inches(0.2), "СРОК ОПЛАТЫ: 1–3 РАБОЧИХ ДНЯ", size=6, color=GOLD_DARK, bold=True, align=PP_ALIGN.RIGHT)
checks = [
    ("180 000 ₽", "Компенсация за коктейли", "300 × 600 ₽ — фикс, независимо от явки"),
    ("20 040 ₽", "Компенсация за шаттл", "охрана, клининг, хостес Велком-центра"),
    ("120 000 ₽", "VIP и столы — ваша зона", "VIP №1 40К + VIP №2 30К + столы 50К"),
    ("Кальяны", "Доход от Макао", "по факту продаж — полностью ваш"),
]
for i,(price,title,desc) in enumerate(checks):
    y = Inches(4.05 + i*0.62)
    add_rect(slide, Inches(0.9), y, Inches(0.35), Inches(0.35), NAVY if i<3 else GOLD)
    add_text(slide, Inches(0.92), y+Inches(0.05), Inches(0.31), Inches(0.25), "✓", size=8, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1.4), y+Inches(0.02), Inches(1.5), Inches(0.2), price, size=8, color=NAVY, bold=True)
    add_text(slide, Inches(2.8), y+Inches(0.02), Inches(4.5), Inches(0.2), title, size=7.5, color=NAVY, bold=True)
    add_text(slide, Inches(2.8), y+Inches(0.22), Inches(4.5), Inches(0.15), desc, size=6.5, color=MUTED)
    if i<3:
        add_rect(slide, Inches(0.9), y+Inches(0.5), Inches(7.4), Inches(0.01), LIGHT_GRAY)
add_rect(slide, Inches(0.9), Inches(6.7), Inches(7.4), Inches(0.32), RGBColor(0xFF,0xF7,0xE0))
add_text(slide, Inches(1.0), Inches(6.75), Inches(7.2), Inches(0.2), "⚠️  Важно: зафиксируйте суммы как фиксированные обязательства Кактусы — их риски не должны переходить на вас.", size=6.5, color=NAVY, bold=True)

# right panel - next steps
add_rect(slide, Inches(8.8), Inches(1.6), Inches(3.85), Inches(5.45), BEIGE, radius=True)
add_text(slide, Inches(9.1), Inches(1.85), Inches(3.25), Inches(0.2), "СЛЕДУЮЩИЕ ШАГИ", size=7, color=GOLD_DARK, bold=True)
steps = [
    ("1", "СОГЛАСОВАТЬ", "Утвердить дату и условия\nс Кактуса и Mriya"),
    ("2", "ПОДПИСАТЬ", "Договор с фиксацией\n330 000 ₽ и сроков оплаты"),
    ("3", "ЗАПУСТИТЬ", "Продажи билетов с 01.11\nпродвижение и афиши"),
    ("4", "ПРОВЕСТИ", "5 декабря — вечеринка\nна 300+ гостей"),
]
for i,(num,title,desc) in enumerate(steps):
    y = Inches(2.2 + i*0.95)
    add_circle(slide, Inches(9.1), y, Inches(0.32), NAVY)
    add_text(slide, Inches(9.1), y+Inches(0.02), Inches(0.32), Inches(0.28), num, size=8, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(9.6), y, Inches(2.7), Inches(0.18), title, size=7, color=NAVY, bold=True)
    add_text(slide, Inches(9.6), y+Inches(0.2), Inches(2.7), Inches(0.35), desc, size=6.5, color=MUTED, line_spacing=1.2)
    if i<3:
        add_rect(slide, Inches(9.25), y+Inches(0.42), Inches(0.02), Inches(0.45), GOLD)
# CTA button
add_rect(slide, Inches(9.1), Inches(6.2), Inches(3.25), Inches(0.55), NAVY)
add_text(slide, Inches(9.2), Inches(6.32), Inches(3.05), Inches(0.3), "ГОТОВЫ ОБСУДИТЬ?  →", size=8, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(9.1), Inches(6.8), Inches(3.25), Inches(0.2), "Контакты организатора Кактуса — по запросу", size=6, color=MUTED, align=PP_ALIGN.CENTER, italic=True)


# ========== SLIDE 9 - CONTACTS / FINAL ==========
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_rect(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.06), GOLD)
# left
add_rect(slide, Inches(0), Inches(0.06), Inches(7.5), Inches(7.44), NAVY)
add_text(slide, Inches(0.7), Inches(0.6), Inches(6), Inches(0.2), "G L E B  F I L I P C H E N K O W  ×  N O E M Y  A B R A N T E S", size=7, color=GOLD, bold=True)
add_text(slide, Inches(0.7), Inches(0.85), Inches(6), Inches(0.7), "Сделаем\n5 декабря\nлегендой Foster", size=30, color=WHITE, bold=True, line_spacing=0.9)
add_rect(slide, Inches(0.7), Inches(1.9), Inches(0.8), Inches(0.04), GOLD)
add_text(slide, Inches(0.7), Inches(2.1), Inches(6), Inches(0.6), "Ночной клуб Foster ждёт своих 300 гостей.\nВаша прибыль — 483 549 ₽ за один вечер.", size=9, color=BEIGE, line_spacing=1.3)
# contacts
contacts = [
    ("📅", "ДАТА", "Суббота, 5 декабря 2026"),
    ("📍", "МЕСТО", "Foster Nightclub, Крым"),
    ("🎟️", "БИЛЕТЫ", "3 000 → 3 500 → 4 000 ₽"),
]
for i,(icon,label,val) in enumerate(contacts):
    y = Inches(3.0 + i*0.7)
    add_circle(slide, Inches(0.7), y, Inches(0.35), RGBColor(0x13,0x2B,0x4A))
    add_text(slide, Inches(0.7), y+Inches(0.05), Inches(0.35), Inches(0.25), icon, size=8, color=GOLD, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1.25), y+Inches(0.02), Inches(1.2), Inches(0.15), label, size=6, color=GOLD_LIGHT, bold=True)
    add_text(slide, Inches(1.25), y+Inches(0.18), Inches(3.5), Inches(0.2), val, size=8, color=WHITE, bold=True)
# CTA
add_rect(slide, Inches(0.7), Inches(5.5), Inches(3.0), Inches(0.5), GOLD)
add_text(slide, Inches(0.8), Inches(5.62), Inches(2.8), Inches(0.25), "ПОДТВЕРДИТЬ УЧАСТИЕ →", size=8, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(3.9), Inches(5.62), Inches(3.0), Inches(0.25), "foster.club  •  кактуса  •  mriya resort", size=6, color=MUTED, italic=True)
add_text(slide, Inches(0.7), Inches(6.3), Inches(6), Inches(0.2), "Презентация подготовлена для внутренней отчётности и переговоров с партнёрами", size=6, color=RGBColor(0x5A,0x6A,0x80), italic=True)

# right - summary numbers
add_text(slide, Inches(8.0), Inches(0.5), Inches(4.6), Inches(0.2), "КЛЮЧЕВЫЕ ЦИФРЫ НА ОДНОМ ЭКРАНЕ", size=7, color=NAVY, bold=True)
# waterfall card
add_rect(slide, Inches(8.0), Inches(0.85), Inches(4.6), Inches(2.2), BEIGE, radius=True)
# row 1
add_text(slide, Inches(8.3), Inches(1.0), Inches(2.5), Inches(0.15), "ВАША ВЫРУЧКА", size=6, color=MUTED, bold=True)
add_text(slide, Inches(10.2), Inches(1.0), Inches(1.9), Inches(0.15), "768 025 ₽", size=9, color=NAVY, bold=True, align=PP_ALIGN.RIGHT)
add_rect(slide, Inches(8.3), Inches(1.25), Inches(4.0), Inches(0.08), GOLD)
# row2
add_text(slide, Inches(8.3), Inches(1.45), Inches(2.5), Inches(0.15), "ВАШИ РАСХОДЫ", size=6, color=MUTED, bold=True)
add_text(slide, Inches(10.2), Inches(1.45), Inches(1.9), Inches(0.15), "– 284 476 ₽", size=9, color=NAVY, bold=True, align=PP_ALIGN.RIGHT)
add_rect(slide, Inches(8.3), Inches(1.7), Inches(1.5), Inches(0.08), NAVY)
# row3
add_rect(slide, Inches(8.3), Inches(1.95), Inches(4.0), Inches(0.65), NAVY)
add_text(slide, Inches(8.45), Inches(2.05), Inches(2.0), Inches(0.15), "ЧИСТАЯ ПРИБЫЛЬ", size=6, color=GOLD_LIGHT, bold=True)
add_text(slide, Inches(8.45), Inches(2.25), Inches(2.0), Inches(0.25), "483 549 ₽", size=13, color=WHITE, bold=True)
add_text(slide, Inches(10.5), Inches(2.05), Inches(1.6), Inches(0.5), "~63%\nрентаб.", size=7, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
# guarantee
add_rect(slide, Inches(8.0), Inches(3.25), Inches(4.6), Inches(1.0), WHITE, line=GOLD, radius=True)
add_text(slide, Inches(8.3), Inches(3.4), Inches(4.0), Inches(0.15), "ГАРАНТИРОВАНО ДОГОВОРОМ", size=6, color=GOLD_DARK, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.3), Inches(3.62), Inches(4.0), Inches(0.25), "330 000 ₽", size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.3), Inches(3.9), Inches(4.0), Inches(0.15), "180К коктейли + 50К столы + 70К VIP", size=6, color=MUTED, align=PP_ALIGN.CENTER)
# risk compare
add_rect(slide, Inches(8.0), Inches(4.45), Inches(4.6), Inches(0.9), WHITE, line=LIGHT_GRAY, radius=True)
add_text(slide, Inches(8.3), Inches(4.6), Inches(4.0), Inches(0.15), "БАЛАНС РИСКОВ", size=6, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(8.3), Inches(4.8), Inches(2.0), Inches(0.2), "FOSTER  284К", size=7, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(10.4), Inches(4.8), Inches(2.0), Inches(0.2), "КАКТУСА  1 120К", size=7, color=GOLD_DARK, bold=True, align=PP_ALIGN.CENTER)
add_rect(slide, Inches(8.6), Inches(5.05), Inches(1.0), Inches(0.18), NAVY)
add_rect(slide, Inches(10.4), Inches(5.05), Inches(1.7), Inches(0.18), GOLD)
add_text(slide, Inches(8.3), Inches(5.3), Inches(4.0), Inches(0.15), "Ваш риск в 4 раза ниже — их мотивация выше", size=6, color=MUTED, italic=True, align=PP_ALIGN.CENTER)

# footer
add_rect(slide, Inches(8.0), Inches(6.6), Inches(4.6), Inches(0.5), NAVY)
add_text(slide, Inches(8.1), Inches(6.7), Inches(4.4), Inches(0.3), "ГОТОВЫ ФИКСИРОВАТЬ ДАТУ?  •  5 ДЕКАБРЯ 2026", size=7, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# Save
prs.save('/home/user/Foster_Gleb_Noemy_Presentation.pptx')
print("PPTX saved")
