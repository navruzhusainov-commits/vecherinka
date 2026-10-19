#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пересборка с увеличенным читаемым шрифтом + новый текст: концепция "Чёрным по красному" + финмодели
Убраны слайды 4,5,12,13,14 визуалов
Шрифты увеличены на ~30-40% для читаемости
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import os

# Colors template
GOLD = RGBColor(0xC9, 0xA2, 0x5E)
GOLD_LIGHT = RGBColor(0xE5, 0xC8, 0x88)
GOLD_MID = RGBColor(0xD4, 0xB0, 0x5E)
BEIGE = RGBColor(0xF5, 0xF1, 0xE8)
BEIGE2 = RGBColor(0xDD, 0xD8, 0xCC)
BEIGE3 = RGBColor(0xE2, 0xDE, 0xD4)
MUTED = RGBColor(0x7E, 0x7A, 0x72)
MUTED2 = RGBColor(0x9A, 0x96, 0x8E)
DARK = RGBColor(0x14, 0x14, 0x19)
DARKER = RGBColor(0x09, 0x09, 0x0B)
BURGUNDY = RGBColor(0x7A, 0x1F, 0x2B)
WHITE = RGBColor(0xFF,0xFF,0xFF)
RED_ACCENT = RGBColor(0xC9, 0x1A, 0x2E) # для концепции чёрным по красному

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg",
    "about": "/tmp/foster_images/slide02_img3_Picture_5.jpg",
    "concept_bg": "/tmp/foster_images/slide03_img4_Picture_1.jpg",
    "hall1": "/tmp/foster_images/slide06_img7_Picture_5.jpg",
    "hall2": "/tmp/foster_images/slide06_img8_Picture_6.jpg",
    "dj_bar": "/tmp/foster_images/slide07_img9_Picture_5.jpg",
    "vip": "/tmp/foster_images/slide07_img10_Picture_6.jpg",
    "bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "loft1": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "timeline": "/tmp/foster_images/slide11_img18_Picture_5.jpg",
    "multi1": "/tmp/foster_images/slide10_img14_Picture_5.jpg",
    "multi2": "/tmp/foster_images/slide10_img15_Picture_6.jpg",
}

def add_bg_image(slide, path, darken=0.45):
    if os.path.exists(path):
        slide.shapes.add_picture(path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        s.fill.solid(); s.fill.fore_color.rgb = RGBColor(0x0A,0x0A,0x0F); s.line.fill.background()
        return s
    return None

def tb(slide, l, t, w, h, txt, size=10, color=BEIGE, bold=False, align=PP_ALIGN.LEFT, font="Arial", line=1.15):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.name = font
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.alignment = align
    p.line_spacing = line
    return box

def add_gold_line(slide, l=Inches(0.57), t=Inches(1.42), w=Inches(0.85), h=Inches(0.06)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = GOLD; s.line.fill.background()
    return s
def add_gold_line_alt(slide, t):
    return add_gold_line(slide, t=t)

def card(slide, l, t, w, h, fill=DARK, header=None, body=None, header_color=GOLD_LIGHT, body_color=BEIGE3, header_size=9, body_size=8.5):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background(); s.adjustments[0]=0.06
    if header:
        box = slide.shapes.add_textbox(l+Inches(0.22), t+Inches(0.14), w-Inches(0.44), Inches(0.32))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=header; p.font.name="Arial"; p.font.size=Pt(header_size); p.font.color.rgb=header_color; p.font.bold=True
    if body:
        box = slide.shapes.add_textbox(l+Inches(0.22), t+Inches(0.50), w-Inches(0.44), h-Inches(0.68))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=body; p.font.name="Arial"; p.font.size=Pt(body_size); p.font.color.rgb=body_color; p.line_spacing=1.25
    return s

def header_block(slide, spaced, title, num="01 / 09"):
    tb(slide, Inches(0.55), Inches(0.28), Inches(9.5), Inches(0.35), spaced, size=11, color=GOLD, bold=True)
    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(8.5), Inches(1.0))
    tf=box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
    p.text=title; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=BEIGE; p.font.bold=True; p.line_spacing=0.95
    add_gold_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06))
    tb(slide, Inches(11.50), Inches(0.28), Inches(1.30), Inches(0.30), num, size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    tb(slide, Inches(0.55), Inches(7.05), Inches(6), Inches(0.25), "Фото: галерея Foster Club — Яндекс Карты", size=8, color=MUTED)
    tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), f"FOSTER • {num.split('/')[0].strip()}", size=8, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)

# ===== SLIDE 1 COVER — концепция Чёрным по красному =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["cover"], darken=0.50)
# top badge
r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(0.35), Inches(2.80), Inches(0.95))
r.fill.solid(); r.fill.fore_color.rgb=BEIGE; r.line.fill.background(); r.adjustments[0]=0.08
tb(slide, Inches(0.70), Inches(0.45), Inches(2.50), Inches(0.25), "FOSTER", size=13, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.70), Inches(0.70), Inches(2.50), Inches(0.20), "NIGHT & KARAOKE CLUB", size=7, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.70), Inches(0.88), Inches(2.50), Inches(0.20), "MRiya Resort • Ялта • Крым", size=6, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(9.0), Inches(0.35), Inches(3.80), Inches(0.30), "5 декабря 2026  •  Суббота  •  22:00–06:00", size=8, color=WHITE, align=PP_ALIGN.RIGHT, bold=True)
tb(slide, Inches(9.0), Inches(0.62), Inches(3.80), Inches(0.25), "300–350 гостей  •  Дресс-код: Total Black + Red", size=7, color=GOLD_LIGHT, align=PP_ALIGN.RIGHT)
# title
box = slide.shapes.add_textbox(Inches(0.55), Inches(1.65), Inches(9), Inches(1.8))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="GLEB FILIPCHENKOW"; p.font.name="Arial"; p.font.size=Pt(36); p.font.color.rgb=WHITE; p.font.bold=True; p.line_spacing=0.9
p=tf.add_paragraph(); p.text="& NOEMY ABRANTES"; p.font.name="Arial"; p.font.size=Pt(36); p.font.color.rgb=GOLD; p.font.bold=True; p.line_spacing=0.9; p.space_before=Pt(4)
p=tf.add_paragraph(); p.text="КОНЦЕПЦИЯ «ЧЁРНЫМ ПО КРАСНОМУ»  •  Первый привоз в Крым"; p.font.name="Arial"; p.font.size=Pt(11); p.font.color.rgb=GOLD_LIGHT; p.font.bold=True; p.space_before=Pt(6)
add_gold_line(slide, Inches(0.57), Inches(3.75), Inches(1.0), Inches(0.06))
# dark card with concept teaser
s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(4.00), Inches(6.2), Inches(1.85))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(4.15), Inches(5.8), Inches(0.30), "◆  Клуб в красном свете и бархате", size=10, color=GOLD, bold=True)
box = slide.shapes.add_textbox(Inches(0.75), Inches(4.50), Inches(5.8), Inches(1.15))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Гости — total black с красной деталью. Минимум декора, максимум эффекта: красный свет, бархат, фото-зона. Перформеры — танцовщица и go-go в красном."; p.font.name="Arial"; p.font.size=Pt(9.5); p.font.color.rgb=BEIGE3; p.line_spacing=1.3
# stats bottom larger
for i,(num,lbl) in enumerate([("483 549 ₽","Ваша прибыль"),("63%","Рентабельность"),("330 000 ₽","Гарант")]):
    l = Inches(0.55 + i*2.15)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(6.15), Inches(1.95), Inches(0.90))
    s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.10), Inches(6.25), Inches(1.75), Inches(0.30), num, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.10), Inches(6.60), Inches(1.75), Inches(0.20), lbl, size=7.5, color=GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(7.05), Inches(7), Inches(0.25), "Пересборка на подлинных фото Foster Club • Шрифт увеличен для читаемости", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), "01 / 09", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 2 ХЕДЛАЙНЕР =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Х Е Д Л А Й Н Е Р", "Gleb Filipchenkow & Noemy", num="02 / 09")
# Use hall images
slide.shapes.add_picture(IMG["hall1"], Inches(7.00), Inches(1.80), width=Inches(5.78), height=Inches(3.40))
# left bullets larger
box = slide.shapes.add_textbox(Inches(0.55), Inches(1.85), Inches(6.15), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
bullets=[("Глеб —", "продюсер и соавтор вирусного трека «Mallorca» (взрыв в 2026). Первый приезд в Крым."),
         ("Noemy —", "соавтор релизов «Garota» и «HOJE», европейский вокал."),
         ("Почему продаёт:", "имя уже на слуху, отложенный спрос с лета, 5 декабря — единственное окно."),
         ("Концепция:", "«Чёрным по красному» — тотальный дресс-код и красный свет.")]
for label,body in bullets:
    p=tf.add_paragraph(); p.space_before=Pt(8)
    r=p.add_run(); r.text="◆  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=GOLD; r.font.bold=True
    r=p.add_run(); r.text=label+" "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=body; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RGBColor(0x2A,0x2A,0x30)
# cards bottom larger
card(slide, Inches(0.55), Inches(5.50), Inches(3.95), Inches(1.40), fill=DARK, header="MALLORCA", body="Вирусный хит без бюджета:\nсотни тысяч просмотров,\nмосковские вечеринки 2026.\nДалее — трек с Серябкиной.", header_size=10, body_size=9)
card(slide, Inches(4.70), Inches(5.50), Inches(3.95), Inches(1.40), fill=DARK, header="ПЕРВЫЙ ПРИВОЗ", body="В Крыму не выступали.\nЭффект эксклюзива +\nожидание с лета (сорвался\nиз-за топливного кризиса).", header_size=10, body_size=9)
card(slide, Inches(8.85), Inches(5.50), Inches(3.93), Inches(1.40), fill=BURGUNDY, header="01:00 БЛЭКАУТ", body="Выход хедлайнеров —\nблэкаут и вспышка\nкрасного света. Кульминация\nвечеринки.", header_size=10, body_size=9)

# ===== SLIDE 3 КОНЦЕПЦИЯ ДЕТАЛЬНО =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["concept_bg"], darken=0.50)
tb(slide, Inches(0.55), Inches(0.28), Inches(10), Inches(0.35), "К О Н Ц Е П Ц И Я   « Ч Ё Р Н Ы М   П О   К Р А С Н О М У »", size=11, color=GOLD, bold=True)
box=slide.shapes.add_textbox(Inches(0.55), Inches(0.65), Inches(9), Inches(0.90))
tf=box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]; p.text="Total black + красная деталь"; p.font.name="Arial"; p.font.size=Pt(26); p.font.color.rgb=BEIGE; p.font.bold=True
add_gold_line(slide, Inches(0.57), Inches(1.30), Inches(0.85), Inches(0.06))
tb(slide, Inches(11.50), Inches(0.28), Inches(1.30), Inches(0.35), "03 / 09", size=10, color=MUTED, align=PP_ALIGN.RIGHT)
# 3 cards larger
card(slide, Inches(0.55), Inches(1.75), Inches(3.95), Inches(1.65), fill=DARKER, header="ИДЕЯ", body="Клуб погружается в красный\nсвет и бархат. Гости — total\nblack с яркой красной\nдеталью (аксессуар, помада).", header_size=10, body_size=9.5)
card(slide, Inches(4.70), Inches(1.75), Inches(3.95), Inches(1.65), fill=DARKER, header="ОФОРМЛЕНИЕ", body="Минимум декора — максимум\nэффекта: красный свет,\nкрасный бархат, фото-зона.\nМощно и недорого.", header_size=10, body_size=9.5)
card(slide, Inches(8.85), Inches(1.75), Inches(3.93), Inches(1.65), fill=DARKER, header="ПЕРФОРМЕРЫ", body="Танцовщица и go-go\nв красном. Пластика +\nсвет = идеальный контент\nдля соцсетей.", header_size=10, body_size=9.5)
# тайминг timeline as cards
timings = [("22:00","Сбор + warm-up","DJ, аперитив"),("23:30","Перформанс","Танец в красном"),("01:00","ХЕДЛАЙНЕРЫ","Блэкаут + красный"),("03:00","Afterparty","DJ-сет до 06:00")]
for i,(t,title,desc) in enumerate(timings):
    l = Inches(0.55 + i*3.25)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(5.35), Inches(3.00), Inches(1.25))
    s.fill.solid(); s.fill.fore_color.rgb = BURGUNDY if "01:00" in t else DARK; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.20), Inches(5.45), Inches(2.60), Inches(0.30), t, size=13, color=GOLD_LIGHT if "01:00" not in t else WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.80), Inches(2.60), Inches(0.25), title, size=9, color=BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(6.10), Inches(2.60), Inches(0.25), desc, size=8, color=BEIGE3, align=PP_ALIGN.CENTER)

# ===== SLIDE 4 КАКТУСА — ПОТЕНЦИАЛЬНЫЙ ДОХОД =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  1", "Доход от билетов", num="04 / 09")
# image left bar
slide.shapes.add_picture(IMG["bar"], Inches(0.55), Inches(1.80), width=Inches(5.20), height=Inches(3.20))
tb(slide, Inches(0.55), Inches(5.15), Inches(5.20), Inches(0.25), "Foster Club — бар и танцпол, Яндекс Карты", size=8, color=MUTED, align=PP_ALIGN.CENTER)
# right text larger
box = slide.shapes.add_textbox(Inches(6.20), Inches(1.80), Inches(6.50), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
items = [("Источник:", "продажа билетов 300–350 гостей."),("Стратегия:", "3 000 ₽ (старт) → 3 500 ₽ → 4 000 ₽ в день события."),("Базовый прогноз:", "300 чел. × 3 400 ₽ средний чек = 1 020 000 ₽ выручки.")]
for label,body in items:
    p=tf.add_paragraph(); p.space_before=Pt(8)
    r=p.add_run(); r.text="◆  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=GOLD; r.font.bold=True
    r=p.add_run(); r.text=label+" "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=body; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RGBColor(0x2A,0x2A,0x30)
# highlight card
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.20), Inches(5.00), Inches(6.50), Inches(0.95))
s.fill.solid(); s.fill.fore_color.rgb=BURGUNDY; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(6.40), Inches(5.15), Inches(6.10), Inches(0.30), "Базовый сценарий — 1 020 000 ₽ выручки", size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(6.40), Inches(5.45), Inches(6.10), Inches(0.40), "При SOLD OUT 350 чел. × 3 400 ₽ = 1 190 000 ₽ → прибыль ~70 000 ₽", size=8, color=GOLD_LIGHT, align=PP_ALIGN.CENTER)

# ===== SLIDE 5 КАКТУСА — ЗАТРАТЫ =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  2", "Затраты организатора", num="05 / 09")
# Use two images top (hall)
slide.shapes.add_picture(IMG["hall1"], Inches(0.55), Inches(1.80), width=Inches(6.20), height=Inches(2.80))
slide.shapes.add_picture(IMG["hall2"], Inches(6.95), Inches(1.80), width=Inches(5.83), height=Inches(2.80))
tb(slide, Inches(0.55), Inches(4.75), Inches(6.20), Inches(0.25), "Сцена и бар Foster — фото Яндекс Карты", size=8, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(6.95), Inches(4.75), Inches(5.83), Inches(0.25), "Зеркальный бар — фото Яндекс Карты", size=8, color=MUTED, align=PP_ALIGN.CENTER)
# costs list as cards larger
costs = [("Гонорар", "550 000 ₽", BURGUNDY),("Коктейли", "180 000 ₽", GOLD),("Шаттл", "20 040 ₽", GOLD),("Логистика", "~80 000 ₽", DARK),("Промо", "~100 000 ₽", DARK),("Артисты", "75 000 ₽", DARK),("Тех/Видео/Фото", "85 000 ₽", DARK),("Охрана", "30 000 ₽", DARK)]
# Create 4 columns of 2 cards each? But need larger readable - do 4 cards per row
for i,(name,price,fill) in enumerate(costs):
    row = i // 4
    col = i % 4
    l = Inches(0.55 + col*3.25)
    t = Inches(5.15 + row*1.20)
    is_highlight = fill in [BURGUNDY, GOLD]
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, Inches(3.00), Inches(1.00))
    s.fill.solid(); s.fill.fore_color.rgb=fill; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.15), t+Inches(0.12), Inches(2.70), Inches(0.25), name, size=8.5, color=WHITE if fill==BURGUNDY else (DARK if fill==GOLD else GOLD_LIGHT), bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), t+Inches(0.40), Inches(2.70), Inches(0.30), price, size=9.5, color=WHITE if fill!=GOLD else DARK, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), t+Inches(0.70), Inches(2.70), Inches(0.20), "фикс → вам" if name in ["Коктейли","Шаттл"] else "", size=6.5, color=WHITE if fill==BURGUNDY else MUTED, align=PP_ALIGN.CENTER)
# total
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(6.65), Inches(12.23), Inches(0.50))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.75), Inches(11.8), Inches(0.30), "ИТОГО ЗАТРАТЫ КАКТУСЫ:  ~1 120 040 ₽", size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# ===== SLIDE 6 КАКТУСА — РЕЗУЛЬТАТ =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  3", "Результат: риск на них", num="06 / 09")
slide.shapes.add_picture(IMG["timeline"], Inches(0.55), Inches(1.80), width=Inches(12.23), height=Inches(2.40))
tb(slide, Inches(0.55), Inches(4.35), Inches(12.23), Inches(0.25), "Тайминг продаж билетов — как в шаблоне Foster Club", size=8, color=MUTED, align=PP_ALIGN.CENTER)
# 3 scenario cards larger
scenarios = [("250 чел.","Пессимистичный","825 000 ₽","–295 040 ₽",DARK),("300 чел.","Базовый","1 020 000 ₽","–100 040 ₽",DARK),("350 чел.","SOLD OUT","1 190 000 ₽","+69 960 ₽",BURGUNDY)]
for i,(cnt, title, rev, res, fill) in enumerate(scenarios):
    l = Inches(0.55 + i*4.20)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(4.80), Inches(3.95), Inches(2.00))
    s.fill.solid(); s.fill.fore_color.rgb=fill; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.20), Inches(4.95), Inches(3.55), Inches(0.35), cnt, size=15, color=WHITE if fill==BURGUNDY else GOLD_LIGHT, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.35), Inches(3.55), Inches(0.25), title, size=9, color=BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.70), Inches(1.77), Inches(0.20), "Выручка", size=7, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(2.17), Inches(5.70), Inches(1.77), Inches(0.20), "Итог", size=7, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.95), Inches(1.77), Inches(0.30), rev, size=10, color=WHITE if fill==BURGUNDY else BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(2.17), Inches(5.95), Inches(1.77), Inches(0.30), res, size=10, color=RGBColor(0xFF,0x8A,0x90) if "–" in res else RGBColor(0x7E,0xD3,0xA0), bold=True, align=PP_ALIGN.CENTER)
# highlight
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(6.05), Inches(12.23), Inches(0.65))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.15), Inches(11.8), Inches(0.30), "💡 Вывод:", size=10, color=BURGUNDY, bold=True)
tb(slide, Inches(0.75), Inches(6.45), Inches(11.8), Inches(0.25), "Кактуса выходит в плюс только при SOLD OUT — значит, будет продавать до последнего билета. Ваши 200 040 ₽ зафиксированы в договоре.", size=9, color=DARK)

# ===== SLIDE 7 FOSTER — ЗАТРАТЫ =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   F O S T E R  •  1", "Наши затраты — 284 476 ₽", num="07 / 09")
slide.shapes.add_picture(IMG["multi1"], Inches(0.55), Inches(1.80), width=Inches(2.80), height=Inches(3.40))
slide.shapes.add_picture(IMG["multi2"], Inches(3.50), Inches(1.80), width=Inches(3.90), height=Inches(3.40))
tb(slide, Inches(0.55), Inches(5.35), Inches(2.80), Inches(0.25), "Танцпол Foster", size=8, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(3.50), Inches(5.35), Inches(3.90), Inches(0.25), "VIP-зона — фото Яндекс Карты", size=8, color=MUTED, align=PP_ALIGN.CENTER)
# table larger
costs = [("Персонал клуба","165 470 ₽","официанты, бармены, повара"),("Себестоимость бара","92 006 ₽","25% от 368 025 ₽"),("Себестоимость кухни","27 000 ₽","27% от 100 000 ₽")]
box = slide.shapes.add_textbox(Inches(7.60), Inches(1.85), Inches(5.20), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
for name,price,note in costs:
    p=tf.add_paragraph(); p.space_before=Pt(10)
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=price; r.font.name="Arial"; r.font.size=Pt(13); r.font.color.rgb=DARK; r.font.bold=True
    p=tf.add_paragraph()
    r=p.add_run(); r.text=note; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(14)
r=p.add_run(); r.text="Итого затраты:  284 476 ₽"; r.font.name="Arial"; r.font.size=Pt(12); r.font.color.rgb=BURGUNDY; r.font.bold=True
# cards bottom larger
card(slide, Inches(0.55), Inches(5.80), Inches(3.95), Inches(1.20), fill=DARK, header="ПЕРСОНАЛ 58%", body="165 470 ₽ — официанты,\nбармены, повара, менеджмент.\nКрупнейшая статья.", header_size=10, body_size=9)
card(slide, Inches(4.70), Inches(5.80), Inches(3.95), Inches(1.20), fill=DARK, header="ФУДКОСТ 42%", body="Бар 25% + Кухня 27% —\nнормальные показатели.\nПокрываются выручкой.", header_size=10, body_size=9)
card(slide, Inches(8.85), Inches(5.80), Inches(3.93), Inches(1.20), fill=DARK, header="0 ₽ ПРОЖИВАНИЕ", body="Полностью исключено.\nНи скидок, ни затрат.\nАренда — 0 ₽.", header_size=10, body_size=9)

# ===== SLIDE 8 FOSTER — ДОХОДЫ =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   F O S T E R  •  2", "Наши доходы — 768 025 ₽", num="08 / 09")
# two images left
slide.shapes.add_picture(IMG["bar"], Inches(0.55), Inches(1.80), width=Inches(3.90), height=Inches(3.30))
tb(slide, Inches(0.55), Inches(5.25), Inches(3.90), Inches(0.25), "Бар Foster — Яндекс Карты", size=8, color=MUTED, align=PP_ALIGN.CENTER)
slide.shapes.add_picture(IMG["loft1"], Inches(4.65), Inches(1.80), width=Inches(3.90), height=Inches(3.30))
tb(slide, Inches(4.65), Inches(5.25), Inches(3.90), Inches(0.25), "Лаунж-зона — фото клуба", size=8, color=MUTED, align=PP_ALIGN.CENTER)
# incomes table larger
incomes = [("Бар","368 025 ₽","валовая выручка"),("Кухня","100 000 ₽","валовая выручка"),("◆ Коктейли","180 000 ₽","300 × 600 ₽ → вам"),("◆ Столы","50 000 ₽","4×5 000 + 3×10 000"),("◆ VIP №1","40 000 ₽","депозит"),("◆ VIP №2","30 000 ₽","депозит")]
box = slide.shapes.add_textbox(Inches(8.85), Inches(1.80), Inches(3.95), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
for name,price,note in incomes:
    p=tf.add_paragraph(); p.space_before=Pt(6)
    is_hl="◆" in name
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(10); r.font.color.rgb=GOLD if is_hl else DARK; r.font.bold=True
    r=p.add_run(); r.text=price+"  "; r.font.name="Arial"; r.font.size=Pt(10.5); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text="— "+note; r.font.name="Arial"; r.font.size=Pt(8); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(10)
r=p.add_run(); r.text="Итого: 768 025 ₽ + кальяны"; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=BURGUNDY; r.font.bold=True
# bottom highlight
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.70), Inches(12.23), Inches(1.10))
s.fill.solid(); s.fill.fore_color.rgb=BURGUNDY; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(5.85), Inches(11.8), Inches(0.30), "ГАРАНТИРОВАНО ДОГОВОРОМ — 330 000 ₽", size=11, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(6.20), Inches(11.8), Inches(0.45), "180 000 ₽ коктейли (фикс) + 50 000 ₽ столы + 70 000 ₽ VIP-залы — независимо от явки. Кальяны от Макао — по факту.", size=9, color=GOLD_LIGHT, align=PP_ALIGN.CENTER)

# ===== SLIDE 9 ФИНАЛ — ПРИБЫЛЬ =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["cover"], darken=0.52)
tb(slide, Inches(0.55), Inches(0.28), Inches(10), Inches(0.35), "Г Л А В Н Ы Й   И Т О Г", size=11, color=GOLD, bold=True)
box=slide.shapes.add_textbox(Inches(0.55), Inches(0.65), Inches(9), Inches(1.2))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Сделка безопасна и"; p.font.name="Arial"; p.font.size=Pt(30); p.font.color.rgb=BEIGE; p.font.bold=True
p=tf.add_paragraph(); p.text="высокорентабельна"; p.font.name="Arial"; p.font.size=Pt(30); p.font.color.rgb=GOLD; p.font.bold=True
add_gold_line(slide, Inches(0.57), Inches(1.65), Inches(1.0), Inches(0.06))
# profit card large
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.95), Inches(6.5), Inches(1.80))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(2.10), Inches(6.10), Inches(0.30), "ВАША ЧИСТАЯ ПРИБЫЛЬ", size=9, color=GOLD_LIGHT, bold=True)
tb(slide, Inches(0.75), Inches(2.45), Inches(3.50), Inches(0.50), "483 549 ₽", size=26, color=WHITE, bold=True)
tb(slide, Inches(4.50), Inches(2.50), Inches(1.80), Inches(0.50), "~63%\nрентабельность", size=9, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(3.05), Inches(6.10), Inches(0.25), "768 025 ₽ доходы – 284 476 ₽ затраты  + прибыль с кальянов", size=8, color=BEIGE3)
# right card with what to fix
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.30), Inches(1.95), Inches(5.30), Inches(1.80))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(7.50), Inches(2.10), Inches(4.90), Inches(0.30), "ЧТО ЗАФИКСИРОВАТЬ В ДОГОВОРЕ", size=9, color=BURGUNDY, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(7.50), Inches(2.45), Inches(4.90), Inches(1.00))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="✓  180 000 ₽ — компенсация коктейлей (фикс)\n✓  20 040 ₽ — оплата шаттла (фикс)\n✓  Оплата 1–3 рабочих дня после события"; p.font.name="Arial"; p.font.size=Pt(8.5); p.font.color.rgb=DARK; p.line_spacing=1.3
# bottom timeline recap? simpler: CTA
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.90), Inches(12.23), Inches(0.95))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.05), Inches(11.8), Inches(0.30), "Вы получаете почти полмиллиона рублей чистыми, не вкладываясь в артистов, аренду и оборудование.", size=9.5, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(6.40), Inches(11.8), Inches(0.30), "Задача — жёстко зафиксировать компенсации.  •  Организатор максимально мотивирован продать SOLD OUT.", size=8, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(7.05), Inches(9), Inches(0.25), "Презентация подготовлена для переговоров с Foster • Шрифт увеличен для удобства чтения", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), "09 / 09", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

out = "/home/user/Foster_Gleb_Noemy_Presentation.pptx"
prs.save(out)
print(f"Saved PPTX {out} slides={len(prs.slides)}")
