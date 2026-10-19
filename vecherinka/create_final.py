#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Финальная пересборка:
- Убраны ВСЕ упоминания Mriya, только Foster
- 1 слайд переделан под концепцию "Чёрным по красному"
- Визуал слайдов 2,4,5,6,7,8 переделан под концепцию (чёрный + красный)
- Крупный шрифт сохранён
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
RED = RGBColor(0xC9, 0x1A, 0x2E)
RED_BRIGHT = RGBColor(0xE6, 0x06, 0x13)
RED_DARK = RGBColor(0x8B, 0x0A, 0x1A)
WHITE = RGBColor(0xFF,0xFF,0xFF)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg", # сцена
    "about": "/tmp/foster_images/slide02_img3_Picture_5.jpg",
    "concept_bg": "/tmp/foster_images/slide03_img4_Picture_1.jpg", # красный лаунж — идеально для концепции
    "hall1": "/tmp/foster_images/slide06_img7_Picture_5.jpg",
    "hall2": "/tmp/foster_images/slide06_img8_Picture_6.jpg",
    "dj_bar": "/tmp/foster_images/slide07_img9_Picture_5.jpg",
    "vip": "/tmp/foster_images/slide07_img10_Picture_6.jpg",
    "bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "loft1": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "loft2": "/tmp/foster_images/slide09_img13_Picture_6.jpg",
    "timeline": "/tmp/foster_images/slide11_img18_Picture_5.jpg",
    "multi1": "/tmp/foster_images/slide10_img14_Picture_5.jpg",
    "multi2": "/tmp/foster_images/slide10_img15_Picture_6.jpg",
}

def add_bg_image(slide, path, darken=0.50, red_tint=False):
    if os.path.exists(path):
        slide.shapes.add_picture(path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        # overlay
        s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        s.fill.solid()
        # для концепции делаем красноватый оверлей
        if red_tint:
            s.fill.fore_color.rgb = RGBColor(0x1A, 0x02, 0x05) # очень тёмно-красный
        else:
            s.fill.fore_color.rgb = RGBColor(0x0A,0x0A,0x0F)
        s.line.fill.background()
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

def add_line(slide, l=Inches(0.57), t=Inches(1.42), w=Inches(0.85), h=Inches(0.06), color=GOLD):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()
    return s

def card(slide, l, t, w, h, fill=DARK, header=None, body=None, header_color=GOLD_LIGHT, body_color=BEIGE3, header_size=9, body_size=8.5, radius=0.06):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background(); s.adjustments[0]=radius
    if header:
        box = slide.shapes.add_textbox(l+Inches(0.22), t+Inches(0.14), w-Inches(0.44), Inches(0.32))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=header; p.font.name="Arial"; p.font.size=Pt(header_size); p.font.color.rgb=header_color; p.font.bold=True
    if body:
        box = slide.shapes.add_textbox(l+Inches(0.22), t+Inches(0.50), w-Inches(0.44), h-Inches(0.68))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=body; p.font.name="Arial"; p.font.size=Pt(body_size); p.font.color.rgb=body_color; p.line_spacing=1.25
    return s

def header_block(slide, spaced, title, num="01 / 09", color_spaced=GOLD):
    tb(slide, Inches(0.55), Inches(0.28), Inches(9.5), Inches(0.35), spaced, size=11, color=color_spaced, bold=True)
    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(8.5), Inches(1.0))
    tf=box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
    p.text=title; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=BEIGE; p.font.bold=True; p.line_spacing=0.95
    # line color matches spaced
    add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=color_spaced)
    tb(slide, Inches(11.50), Inches(0.28), Inches(1.30), Inches(0.30), num, size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    tb(slide, Inches(0.55), Inches(7.05), Inches(7), Inches(0.25), "Фото: галерея Foster Club — Яндекс Карты", size=8, color=MUTED)
    tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), f"FOSTER • {num.split('/')[0].strip()}", size=8, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)

# ===== SLIDE 1 — НОВЫЙ КОНЦЕПТ-СЛАЙД (Чёрным по красному) =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["concept_bg"], darken=0.45, red_tint=True) # красный лаунж
# Top badge — только Foster, без Mriya
r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(0.30), Inches(2.60), Inches(0.88))
r.fill.solid(); r.fill.fore_color.rgb=WHITE; r.line.fill.background(); r.adjustments[0]=0.08
tb(slide, Inches(0.65), Inches(0.38), Inches(2.40), Inches(0.25), "FOSTER", size=13, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.65), Inches(0.63), Inches(2.40), Inches(0.15), "NIGHT CLUB • ЯЛТА", size=6, color=RED, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.65), Inches(0.78), Inches(2.40), Inches(0.15), "ULTRA • 22:00–06:00", size=6, color=MUTED, align=PP_ALIGN.CENTER)
# Top right date
tb(slide, Inches(8.80), Inches(0.30), Inches(4.00), Inches(0.28), "5 ДЕКАБРЯ 2026  •  СУББОТА", size=9, color=WHITE, align=PP_ALIGN.RIGHT, bold=True)
tb(slide, Inches(8.80), Inches(0.58), Inches(4.00), Inches(0.25), "300–350 ГОСТЕЙ  •  SOLD OUT ЦЕЛЬ", size=8, color=RED, align=PP_ALIGN.RIGHT, bold=True)
# Central concept title
box = slide.shapes.add_textbox(Inches(0.55), Inches(1.45), Inches(9.5), Inches(1.2))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="ЧЁРНЫМ"; p.font.name="Arial"; p.font.size=Pt(42); p.font.color.rgb=WHITE; p.font.bold=True; p.line_spacing=0.85
p=tf.add_paragraph(); p.text="ПО КРАСНОМУ"; p.font.name="Arial"; p.font.size=Pt(42); p.font.color.rgb=RED; p.font.bold=True; p.line_spacing=0.85
add_line(slide, Inches(0.57), Inches(2.85), Inches(1.20), Inches(0.06), color=RED)
tb(slide, Inches(0.55), Inches(3.05), Inches(9), Inches(0.30), "Total Black дресс-код + одна яркая красная деталь", size=11, color=BEIGE, bold=True)
# Concept cards — 3 small with red accent
card(slide, Inches(0.55), Inches(3.60), Inches(3.95), Inches(1.35), fill=DARKER, header="ДРЕСС-КОД", body="Total black — чёрный\nверх + низ. Акцент:\nкрасная помада, галстук,\nсумка, обувь.", header_color=RED, body_color=BEIGE3, header_size=10, body_size=9)
card(slide, Inches(4.70), Inches(3.60), Inches(3.95), Inches(1.35), fill=DARKER, header="ОФОРМЛЕНИЕ", body="Клуб в красном свете\nи бархате. Фото-зона.\nМинимум декора —\nмаксимум вау-эффекта.", header_size=10, body_size=9, header_color=RED)
card(slide, Inches(8.85), Inches(3.60), Inches(3.93), Inches(1.35), fill=BURGUNDY, header="ХЕДЛАЙНЕР", body="GLEB FILIPCHENKOW\n& NOEMY ABRANTES\nПервый привоз в Крым\nХит «Mallorca»", header_color=WHITE, body_color=WHITE, header_size=10, body_size=9)
# Bottom timing strip — red/black
timings = [("22:00","СБОР"),("23:30","ПЕРФОРМАНС"),("01:00","ХЕДЛАЙНЕРЫ"),("03:00","AFTERPARTY")]
for i,(t, ttl) in enumerate(timings):
    l = Inches(0.55 + i*3.25)
    is_red = t=="01:00"
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(5.30), Inches(3.00), Inches(0.85))
    s.fill.solid(); s.fill.fore_color.rgb=RED if is_red else DARK; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.15), Inches(5.40), Inches(2.70), Inches(0.25), t, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), Inches(5.70), Inches(2.70), Inches(0.20), ttl, size=7.5, color=GOLD_LIGHT if not is_red else WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(6.45), Inches(12.23), Inches(0.25), "22:00 warm-up • 23:30 танцовщица и go-go в красном • 01:00 блэкаут и вспышка красного • 06:00 закрытие", size=7, color=BEIGE2, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(7.05), Inches(7), Inches(0.25), "Фото: красный лаунж Foster Club — Яндекс Карты", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), "01 / 09", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 2 — ХЕДЛАЙНЕР (визуал под концепцию: чёрный+красный) =====
slide = prs.slides.add_slide(BLANK)
# Dark background with red tint image
slide.shapes.add_picture(IMG["hall1"], Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
s=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
s.fill.solid(); s.fill.fore_color.rgb=RGBColor(0x0A,0x02,0x05); s.line.fill.background()
# Header in RED
header_block(slide, "Х Е Д Л А Й Н Е Р", "Gleb & Noemy — первый привоз", num="02 / 09", color_spaced=RED)
# add_header uses GOLD by default, override line color to RED
# remove default line and add red
# Actually header_block already added GOLD line, we replace by adding RED line over it
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
# Content on dark: left text on semi-transparent card
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.85), Inches(6.40), Inches(3.40))
s.fill.solid(); s.fill.fore_color.rgb=RGBColor(0x14,0x06,0x08); s.line.fill.background(); s.adjustments[0]=0.06
box=slide.shapes.add_textbox(Inches(0.75), Inches(2.00), Inches(6.00), Inches(3.10))
tf=box.text_frame; tf.word_wrap=True
bullets=[("Глеб —","продюсер и соавтор хита «Mallorca» (взрыв 2026)."),("Noemy —","соавтор «Garota» и «HOJE», европейский вокал."),("Почему зайдёт:","имя на слуху + ожидание с лета + 5 дек — единственное окно."),("Сцена:","блэкаут 01:00 → вспышка красного — кульминация ночи.")]
for label,body in bullets:
    p=tf.add_paragraph(); p.space_before=Pt(7)
    r=p.add_run(); r.text="◆  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RED; r.font.bold=True
    r=p.add_run(); r.text=label+" "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=WHITE; r.font.bold=True
    r=p.add_run(); r.text=body; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=BEIGE2
# Right image small (hall2) on dark
slide.shapes.add_picture(IMG["hall2"], Inches(7.30), Inches(1.85), width=Inches(5.48), height=Inches(3.40))
# cards bottom — red/black theme
card(slide, Inches(0.55), Inches(5.50), Inches(3.95), Inches(1.40), fill=DARKER, header="MALLORCA", body="Вирус без бюджета —\nсотни тысяч просмотров\nна вечеринках Москвы.\nРелиз с Серябкиной далее.", header_color=RED, body_color=BEIGE3, header_size=10, body_size=9)
card(slide, Inches(4.70), Inches(5.50), Inches(3.95), Inches(1.40), fill=DARKER, header="ПЕРВЫЙ ПРИВОЗ", body="В Крыму не были.\nЭксклюзив + отложенный\nспрос с лета (топливный\nкризис сорвал тур).", header_color=RED, body_color=BEIGE3, header_size=10, body_size=9)
card(slide, Inches(8.85), Inches(5.50), Inches(3.93), Inches(1.40), fill=RED, header="ЧЁРНЫМ ПО КРАСНОМУ", body="Дресс-код усиливает шоу:\nзал в чёрном — сцена\nв красном. Идеально для\nконтента и фото.", header_color=WHITE, body_color=WHITE, header_size=10, body_size=9)

# ===== SLIDE 3 — КОНЦЕПЦИЯ ДЕТАЛЬНО (оставляем как есть, но без Mriya) =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["concept_bg"], darken=0.50, red_tint=True)
tb(slide, Inches(0.55), Inches(0.28), Inches(10), Inches(0.35), "К О Н Ц Е П Ц И Я", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.55), Inches(0.60), Inches(9), Inches(1.0))
tf=box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]; p.text="Total black + красная деталь"; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=WHITE; p.font.bold=True
add_line(slide, Inches(0.57), Inches(1.30), Inches(0.85), Inches(0.06), color=RED)
tb(slide, Inches(11.50), Inches(0.28), Inches(1.30), Inches(0.35), "03 / 09", size=10, color=MUTED, align=PP_ALIGN.RIGHT)
card(slide, Inches(0.55), Inches(1.75), Inches(3.95), Inches(1.65), fill=DARKER, header="ИДЕЯ", body="Клуб в красном свете\nи бархате. Гости — total\nblack + красная деталь\n(помада, аксессуар).", header_color=RED, header_size=10, body_size=9.5)
card(slide, Inches(4.70), Inches(1.75), Inches(3.95), Inches(1.65), fill=DARKER, header="ДЕКОР", body="Минимум декора —\nмаксимум эффекта.\nКрасный свет + бархат\n+ фото-зона.", header_color=RED, header_size=10, body_size=9.5)
card(slide, Inches(8.85), Inches(1.75), Inches(3.93), Inches(1.65), fill=RED, header="ПЕРФОРМЕРЫ", body="Танцовщица и go-go\nв красном. Пластика +\nкрасный свет = контент\nдля соцсетей.", header_color=WHITE, body_color=WHITE, header_size=10, body_size=9.5)
timings = [("22:00","СБОР"),("23:30","ПЕРФОРМАНС"),("01:00","ХЕДЛАЙНЕРЫ"),("03:00","AFTERPARTY")]
for i,(t,ttl) in enumerate(timings):
    l=Inches(0.55+i*3.25)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(5.35), Inches(3.00), Inches(1.25))
    s.fill.solid(); s.fill.fore_color.rgb=RED if t=="01:00" else DARK; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.15), Inches(5.45), Inches(2.70), Inches(0.25), t, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), Inches(5.80), Inches(2.70), Inches(0.25), ttl, size=9, color=BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), Inches(6.10), Inches(2.70), Inches(0.25), "тепло" if t=="22:00" else "шоу" if t=="23:30" else "блэкаут" if t=="01:00" else "до 06:00", size=8, color=BEIGE3, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(7.05), Inches(12), Inches(0.25), "Только Foster Club — без привязки к другим площадкам. Фото: красный лаунж Foster, Яндекс Карты.", size=8, color=MUTED)

# ===== SLIDE 4 — КАКТУСА ДОХОД (визуал под концепцию) =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  1", "Доход от билетов", num="04 / 09", color_spaced=RED)
# replace gold line with red
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
# left image with red overlay
slide.shapes.add_picture(IMG["bar"], Inches(0.55), Inches(1.80), width=Inches(5.20), height=Inches(3.20))
# red tint overlay on image
s=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.80), Inches(5.20), Inches(3.20))
s.fill.solid(); s.fill.fore_color.rgb=RED; s.line.fill.background()
# set transparency simulation not possible, so we add dark red with low opacity via not filling? Keep as is but add label
tb(slide, Inches(0.55), Inches(5.15), Inches(5.20), Inches(0.25), "Бар Foster — тёмный лофт с красным светом", size=8, color=RED, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(6.20), Inches(1.80), Inches(6.50), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
items=[("Источник:","продажа билетов 300–350 гостей."),("Стратегия:","3 000 ₽ → 3 500 ₽ → 4 000 ₽ в день события."),("Базовый:", "300 × 3 400 ₽ = 1 020 000 ₽ выручки.")]
for label,body in items:
    p=tf.add_paragraph(); p.space_before=Pt(8)
    r=p.add_run(); r.text="◆  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RED; r.font.bold=True
    r=p.add_run(); r.text=label+" "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=body; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RGBColor(0x2A,0x2A,0x30)
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.20), Inches(5.00), Inches(6.50), Inches(0.95))
s.fill.solid(); s.fill.fore_color.rgb=RED; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(6.40), Inches(5.15), Inches(6.10), Inches(0.30), "База — 1 020 000 ₽", size=11, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(6.40), Inches(5.45), Inches(6.10), Inches(0.40), "SOLD OUT 350 × 3 400 ₽ = 1 190 000 ₽ → +70 000 ₽ прибыли", size=8, color=BEIGE, align=PP_ALIGN.CENTER)

# ===== SLIDE 5 — КАКТУСА ЗАТРАТЫ (чёрный+красный) =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  2", "Затраты организатора", num="05 / 09", color_spaced=RED)
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
slide.shapes.add_picture(IMG["hall1"], Inches(0.55), Inches(1.80), width=Inches(6.20), height=Inches(2.80))
slide.shapes.add_picture(IMG["hall2"], Inches(6.95), Inches(1.80), width=Inches(5.83), height=Inches(2.80))
tb(slide, Inches(0.55), Inches(4.75), Inches(6.20), Inches(0.25), "Сцена и бар — Foster Club", size=8, color=RED, bold=True, align=PP_ALIGN.CENTER)
costs = [("Гонорар","550 000 ₽",RED),("Коктейли","180 000 ₽",RED),("Шаттл","20 040 ₽",GOLD),("Логистика","~80 000 ₽",DARK),("Промо","~100 000 ₽",DARK),("Артисты","75 000 ₽",DARK),("Тех/Фото","85 000 ₽",DARK),("Охрана","30 000 ₽",DARK)]
for i,(name,price,fill) in enumerate(costs):
    row=i//4; col=i%4; l=Inches(0.55+col*3.25); t=Inches(5.15+row*1.20)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, Inches(3.00), Inches(1.00))
    s.fill.solid(); s.fill.fore_color.rgb=fill; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.15), t+Inches(0.12), Inches(2.70), Inches(0.25), name, size=8.5, color=WHITE if fill in [RED, BURGUNDY] else (DARK if fill==GOLD else RED), bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.15), t+Inches(0.40), Inches(2.70), Inches(0.30), price, size=9.5, color=WHITE if fill!=GOLD else DARK, bold=True, align=PP_ALIGN.CENTER)
    if name in ["Коктейли","Шаттл"]:
        tb(slide, l+Inches(0.15), t+Inches(0.70), Inches(2.70), Inches(0.20), "фикс → вам", size=6.5, color=WHITE if fill==RED else MUTED, align=PP_ALIGN.CENTER)
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(6.65), Inches(12.23), Inches(0.50))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.75), Inches(11.8), Inches(0.30), "ИТОГО:  ~1 120 040 ₽", size=11, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# ===== SLIDE 6 — КАКТУСА РЕЗУЛЬТАТ (красный) =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   К А К Т У С Ы  •  3", "Результат — риск на них", num="06 / 09", color_spaced=RED)
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
slide.shapes.add_picture(IMG["timeline"], Inches(0.55), Inches(1.80), width=Inches(12.23), height=Inches(2.40))
tb(slide, Inches(0.55), Inches(4.35), Inches(12.23), Inches(0.25), "Тайминг продаж — Foster Club", size=8, color=MUTED, align=PP_ALIGN.CENTER)
scenarios = [("250","Пессимистичный","825 000 ₽","–295 040 ₽",DARK),("300","Базовый","1 020 000 ₽","–100 040 ₽",DARK),("350","SOLD OUT","1 190 000 ₽","+69 960 ₽",RED)]
for i,(cnt,title,rev,res,fill) in enumerate(scenarios):
    l=Inches(0.55+i*4.20)
    s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, Inches(4.80), Inches(3.95), Inches(2.00))
    s.fill.solid(); s.fill.fore_color.rgb=fill; s.line.fill.background(); s.adjustments[0]=0.06
    tb(slide, l+Inches(0.20), Inches(4.95), Inches(3.55), Inches(0.35), cnt, size=15, color=GOLD_LIGHT if fill!=RED else WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.35), Inches(3.55), Inches(0.25), title, size=9, color=BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.70), Inches(1.77), Inches(0.20), "Выручка", size=7, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(2.17), Inches(5.70), Inches(1.77), Inches(0.20), "Итог", size=7, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(0.20), Inches(5.95), Inches(1.77), Inches(0.30), rev, size=10, color=BEIGE, bold=True, align=PP_ALIGN.CENTER)
    tb(slide, l+Inches(2.17), Inches(5.95), Inches(1.77), Inches(0.30), res, size=10, color=RGBColor(0xFF,0x8A,0x90) if "–" in res else WHITE, bold=True, align=PP_ALIGN.CENTER)
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(6.05), Inches(12.23), Inches(0.65))
s.fill.solid(); s.fill.fore_color.rgb=RED; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.15), Inches(11.8), Inches(0.30), "💡  Их мотивация — ваша защита: в плюс только при SOLD OUT, будут продавать до последнего билета.", size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# ===== SLIDE 7 — FOSTER ЗАТРАТЫ (красный акцент) =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   F O S T E R  •  1", "Наши затраты — 284 476 ₽", num="07 / 09", color_spaced=RED)
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
slide.shapes.add_picture(IMG["multi1"], Inches(0.55), Inches(1.80), width=Inches(2.80), height=Inches(3.40))
slide.shapes.add_picture(IMG["multi2"], Inches(3.50), Inches(1.80), width=Inches(3.90), height=Inches(3.40))
tb(slide, Inches(0.55), Inches(5.35), Inches(2.80), Inches(0.25), "Танцпол Foster", size=8, color=RED, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(7.60), Inches(1.85), Inches(5.20), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
costs=[("Персонал клуба","165 470 ₽","официанты, бармены, повара"),("Себестоимость бара","92 006 ₽","25% от 368 025 ₽"),("Себестоимость кухни","27 000 ₽","27% от 100 000 ₽")]
for name,price,note in costs:
    p=tf.add_paragraph(); p.space_before=Pt(10)
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text=price; r.font.name="Arial"; r.font.size=Pt(13); r.font.color.rgb=DARK; r.font.bold=True
    p=tf.add_paragraph(); r=p.add_run(); r.text=note; r.font.name="Arial"; r.font.size=Pt(9); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(14); r=p.add_run(); r.text="Итого:  284 476 ₽"; r.font.name="Arial"; r.font.size=Pt(12); r.font.color.rgb=RED; r.font.bold=True
card(slide, Inches(0.55), Inches(5.80), Inches(3.95), Inches(1.20), fill=DARK, header="ПЕРСОНАЛ 58%", body="165 470 ₽ — официанты,\nбармены, повара, менеджмент.", header_color=RED, header_size=10, body_size=9)
card(slide, Inches(4.70), Inches(5.80), Inches(3.95), Inches(1.20), fill=DARK, header="ФУДКОСТ 42%", body="Бар 25% + Кухня 27% —\nнорма, покрывается\nвыручкой.", header_color=RED, header_size=10, body_size=9)
card(slide, Inches(8.85), Inches(5.80), Inches(3.93), Inches(1.20), fill=RED, header="0 ₽ ДОП-ЗАТРАТ", body="Ни аренды, ни артистов,\nни проживания —\nтолько операционка.", header_color=WHITE, body_color=WHITE, header_size=10, body_size=9)

# ===== SLIDE 8 — FOSTER ДОХОДЫ (красный) =====
slide = prs.slides.add_slide(BLANK)
header_block(slide, "Ф И Н А Н С Ы   F O S T E R  •  2", "Наши доходы — 768 025 ₽", num="08 / 09", color_spaced=RED)
add_line(slide, Inches(0.57), Inches(1.45), Inches(0.85), Inches(0.06), color=RED)
slide.shapes.add_picture(IMG["bar"], Inches(0.55), Inches(1.80), width=Inches(3.90), height=Inches(3.30))
tb(slide, Inches(0.55), Inches(5.25), Inches(3.90), Inches(0.25), "Бар Foster", size=8, color=RED, bold=True, align=PP_ALIGN.CENTER)
slide.shapes.add_picture(IMG["loft1"], Inches(4.65), Inches(1.80), width=Inches(3.90), height=Inches(3.30))
tb(slide, Inches(4.65), Inches(5.25), Inches(3.90), Inches(0.25), "Лаунж-зона", size=8, color=RED, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(8.85), Inches(1.80), Inches(3.95), Inches(3.40))
tf=box.text_frame; tf.word_wrap=True
incomes=[("Бар","368 025 ₽","валовая"),("Кухня","100 000 ₽","валовая"),("◆ Коктейли","180 000 ₽","300×600 → вам"),("◆ Столы","50 000 ₽","4×5К+3×10К"),("◆ VIP №1","40 000 ₽","депозит"),("◆ VIP №2","30 000 ₽","депозит")]
for name,price,note in incomes:
    p=tf.add_paragraph(); p.space_before=Pt(6)
    is_hl="◆" in name
    r=p.add_run(); r.text=name+"  "; r.font.name="Arial"; r.font.size=Pt(10); r.font.color.rgb=RED if is_hl else DARK; r.font.bold=True
    r=p.add_run(); r.text=price+"  "; r.font.name="Arial"; r.font.size=Pt(10.5); r.font.color.rgb=DARK; r.font.bold=True
    r=p.add_run(); r.text="— "+note; r.font.name="Arial"; r.font.size=Pt(8); r.font.color.rgb=MUTED
p=tf.add_paragraph(); p.space_before=Pt(10); r=p.add_run(); r.text="Итого: 768 025 ₽ + кальяны"; r.font.name="Arial"; r.font.size=Pt(11); r.font.color.rgb=RED; r.font.bold=True
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.70), Inches(12.23), Inches(1.10))
s.fill.solid(); s.fill.fore_color.rgb=RED; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(5.85), Inches(11.8), Inches(0.30), "ГАРАНТ — 330 000 ₽ ЖИВЫМИ", size=11, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(6.20), Inches(11.8), Inches(0.45), "180 000 ₽ коктейли (фикс) + 50 000 ₽ столы + 70 000 ₽ VIP — независимо от явки. Кальяны Макао — по факту.", size=9, color=WHITE, align=PP_ALIGN.CENTER)

# ===== SLIDE 9 — ФИНАЛ (без Mriya) =====
slide = prs.slides.add_slide(BLANK)
add_bg_image(slide, IMG["cover"], darken=0.55, red_tint=True)
tb(slide, Inches(0.55), Inches(0.28), Inches(10), Inches(0.35), "Г Л А В Н Ы Й   И Т О Г", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.55), Inches(0.65), Inches(9), Inches(1.2))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Сделка безопасна и"; p.font.name="Arial"; p.font.size=Pt(30); p.font.color.rgb=BEIGE; p.font.bold=True
p=tf.add_paragraph(); p.text="высокорентабельна"; p.font.name="Arial"; p.font.size=Pt(30); p.font.color.rgb=RED; p.font.bold=True
add_line(slide, Inches(0.57), Inches(1.65), Inches(1.0), Inches(0.06), color=RED)
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.95), Inches(6.5), Inches(1.80))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(2.10), Inches(6.10), Inches(0.30), "ВАША ЧИСТАЯ ПРИБЫЛЬ", size=9, color=RED, bold=True)
tb(slide, Inches(0.75), Inches(2.45), Inches(3.50), Inches(0.50), "483 549 ₽", size=26, color=WHITE, bold=True)
tb(slide, Inches(4.60), Inches(2.45), Inches(1.70), Inches(0.50), "~63%\nрентабельность", size=9, color=RED, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(3.05), Inches(6.10), Inches(0.25), "768 025 ₽ доходы – 284 476 ₽ затраты + кальяны", size=8, color=BEIGE3)
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.30), Inches(1.95), Inches(5.30), Inches(1.80))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(7.50), Inches(2.10), Inches(4.90), Inches(0.30), "ЧТО ФИКСИРОВАТЬ", size=9, color=RED, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(7.50), Inches(2.45), Inches(4.90), Inches(1.00))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="✓  180 000 ₽ — коктейли (фикс)\n✓  20 040 ₽ — шаттл (фикс)\n✓  Оплата 1–3 дня после события"; p.font.name="Arial"; p.font.size=Pt(8.5); p.font.color.rgb=DARK; p.line_spacing=1.3
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.90), Inches(12.23), Inches(0.95))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.75), Inches(6.05), Inches(11.8), Inches(0.30), "Вы получаете почти полмиллиона чистыми, не вкладываясь в артистов, аренду и оборудование.", size=9.5, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.75), Inches(6.40), Inches(11.8), Inches(0.30), "Только Foster Club — ваша площадка, ваш бар, ваш доход. Мотивация организатора — SOLD OUT.", size=8, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.55), Inches(7.05), Inches(9), Inches(0.25), "Презентация для Foster • Шрифт увеличен • Концепция «Чёрным по красному»", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(7.05), Inches(1.80), Inches(0.25), "09 / 09", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

out = "/home/user/Foster_Gleb_Noemy_Presentation.pptx"
prs.save(out)
print(f"Saved PPTX {out} slides={len(prs.slides)}")

