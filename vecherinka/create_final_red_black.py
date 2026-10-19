#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Вся презентация в красно-чёрном на фоне фото клуба + меньше слайдов + больше шрифт
5 слайдов, каждый — full photo background + red-black palette + крупный шрифт 12-38pt
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import os

# Palette — красно-чёрный
RED = RGBColor(0xE6, 0x06, 0x13)      # яркий красный акцент
RED_DARK = RGBColor(0xC9, 0x1A, 0x2E)  # C91A2E — основной
BURGUNDY = RGBColor(0x7A, 0x1F, 0x2B)
BLACK = RGBColor(0x0A, 0x0A, 0x0F)
DARK = RGBColor(0x14, 0x14, 0x19)      # 141419
DARKER = RGBColor(0x09, 0x09, 0x0B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BEIGE = RGBColor(0xF5, 0xF1, 0xE8)
BEIGE2 = RGBColor(0xE8, 0xD5, 0xB5)
MUTED = RGBColor(0xB8, 0xB5, 0xAE)
GOLD = RGBColor(0xC9, 0xA2, 0x5E) # оставить для редких акцентов, но основной — RED

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg",
    "about": "/tmp/foster_images/slide02_img3_Picture_5.jpg",
    "concept_bg": "/tmp/foster_images/slide03_img4_Picture_1.jpg",
    "dj_bar": "/tmp/foster_images/slide07_img9_Picture_5.jpg",
    "bar": "/tmp/foster_images/slide08_img11_Picture_5.jpg",
    "hall1": "/tmp/foster_images/slide06_img7_Picture_5.jpg",
    "loft1": "/tmp/foster_images/slide09_img12_Picture_5.jpg",
    "timeline": "/tmp/foster_images/slide11_img18_Picture_5.jpg",
}

def add_bg(slide, path):
    slide.shapes.add_picture(path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
    # тёмно-красный оверлей 55%
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor(0x12, 0x02, 0x04)
    s.line.fill.background()
    # python-pptx нет прозрачности, поэтому используем тёмный, а фото уже затемнено
    return s

def tb(slide, l, t, w, h, txt, size=12, color=WHITE, bold=False, align=PP_ALIGN.LEFT, font="Arial", line=1.2):
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

def line_red(slide, l=Inches(0.60), t=Inches(1.45), w=Inches(1.0), h=Inches(0.07)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = RED; s.line.fill.background()
    return s

def card(slide, l, t, w, h, fill=DARK, header=None, body=None, header_color=RED, body_color=WHITE, header_size=11, body_size=10):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background(); s.adjustments[0]=0.05
    if header:
        box = slide.shapes.add_textbox(l+Inches(0.25), t+Inches(0.18), w-Inches(0.50), Inches(0.35))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=header; p.font.name="Arial"; p.font.size=Pt(header_size); p.font.color.rgb=header_color; p.font.bold=True
    if body:
        box = slide.shapes.add_textbox(l+Inches(0.25), t+Inches(0.58), w-Inches(0.50), h-Inches(0.75))
        tf = box.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]
        p.text=body; p.font.name="Arial"; p.font.size=Pt(body_size); p.font.color.rgb=body_color; p.line_spacing=1.30
    return s

# ===== SLIDE 1 — ТИТУЛ: ЧЁРНЫМ ПО КРАСНОМУ =====
slide = prs.slides.add_slide(BLANK)
add_bg(slide, IMG["cover"])
# Верхний бейдж — только Foster
r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.60), Inches(0.35), Inches(2.20), Inches(0.45))
r.fill.solid(); r.fill.fore_color.rgb=WHITE; r.line.fill.background(); r.adjustments[0]=0.12
tb(slide, Inches(0.65), Inches(0.38), Inches(2.10), Inches(0.40), "FOSTER • ЯЛТА", size=10, color=DARK, bold=True, align=PP_ALIGN.CENTER)
# Дата справа
tb(slide, Inches(8.0), Inches(0.35), Inches(4.70), Inches(0.30), "5 ДЕКАБРЯ 2026 • СУББОТА • 22:00–06:00", size=10, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
tb(slide, Inches(8.0), Inches(0.65), Inches(4.70), Inches(0.25), "300–350 ГОСТЕЙ • SOLD OUT ЦЕЛЬ", size=9, color=RED, bold=True, align=PP_ALIGN.RIGHT)
# Главный заголовок — крупно
box = slide.shapes.add_textbox(Inches(0.60), Inches(1.30), Inches(12.0), Inches(1.8))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="ЧЁРНЫМ ПО КРАСНОМУ"; p.font.name="Arial"; p.font.size=Pt(48); p.font.color.rgb=WHITE; p.font.bold=True; p.line_spacing=0.85
p=tf.add_paragraph(); p.text="GLEB FILIPCHENKOW & NOEMY ABRANTES"; p.font.name="Arial"; p.font.size=Pt(14); p.font.color.rgb=RED; p.font.bold=True; p.space_before=Pt(6)
line_red(slide, Inches(0.62), Inches(3.30), Inches(1.40), Inches(0.08))
tb(slide, Inches(0.60), Inches(3.50), Inches(12), Inches(0.35), "Первый привоз в Крым • Хит «Mallorca» • Total Black дресс-код с красной деталью", size=11, color=BEIGE, bold=True)
# 3 карточки внизу — крупно
card(slide, Inches(0.60), Inches(4.10), Inches(3.90), Inches(1.70), fill=DARK, header="ДРЕСС-КОД", body="Total black: чёрный верх +\nниз. Одна яркая красная\nдеталь — помада, галстук,\nсума, обувь, аксессуар.", header_color=RED, header_size=12, body_size=10)
card(slide, Inches(4.75), Inches(4.10), Inches(3.90), Inches(1.70), fill=DARK, header="ДЕКОР", body="Красный свет + красный\nбархат + фото-зона.\nМинимум декора — максимум\nэффекта. Бюджетно.", header_size=12, body_size=10, header_color=RED)
card(slide, Inches(8.90), Inches(4.10), Inches(3.90), Inches(1.70), fill=RED, header="ХЕДЛАЙНЕР 01:00", body="Блэкаут → вспышка красного.\nGleb & Noemy — кульминация\nночи. Взрывной момент для\nвидео и соцсетей.", header_size=12, body_size=10, header_color=WHITE, body_color=WHITE)
# тайминг лента
tb(slide, Inches(0.60), Inches(6.05), Inches(12.1), Inches(0.30), "22:00 СБОР • 23:30 ПЕРФОРМАНС (танцовщица и go-go в красном) • 01:00 ХЕДЛАЙНЕРЫ • 03:00 AFTERPARTY • 06:00 ЗАКРЫТИЕ", size=8, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.60), Inches(6.95), Inches(6), Inches(0.25), "Только Foster Club", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(6.95), Inches(1.70), Inches(0.25), "01 / 05", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 2 — ХЕДЛАЙНЕР =====
slide = prs.slides.add_slide(BLANK)
add_bg(slide, IMG["concept_bg"])
tb(slide, Inches(0.60), Inches(0.35), Inches(8), Inches(0.35), "Х Е Д Л А Й Н Е Р • ПЕРВЫЙ ПРИВОЗ", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.60), Inches(0.75), Inches(8), Inches(0.90))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Gleb & Noemy — имя продаёт"; p.font.name="Arial"; p.font.size=Pt(30); p.font.color.rgb=WHITE; p.font.bold=True
line_red(slide, Inches(0.62), Inches(1.45), Inches(1.0), Inches(0.08))
# Полупрозрачные карточки на фото — крупно
card(slide, Inches(0.60), Inches(1.75), Inches(6.00), Inches(1.25), fill=DARK, header="GLEB FILIPCHENKOW — MALLORCA", body="Продюсер хита «Mallorca» (лето 2025). Взрыв в 2026 без бюджета — сотни тысяч просмотров. Далее — трек с Серябкиной.", header_size=11, body_size=10, header_color=RED)
card(slide, Inches(0.60), Inches(3.15), Inches(6.00), Inches(1.25), fill=DARK, header="NOEMY ABRANTES — GAROTA / HOJE", body="Европейский вокал, соавтор «Garota» и «HOJE». Идеальна для красной ночи Foster — ярко, стильно, танцевально.", header_size=11, body_size=10, header_color=RED)
card(slide, Inches(0.60), Inches(4.55), Inches(6.00), Inches(1.35), fill=RED, header="ПОЧЕМУ SOLD OUT", body="Отложенный спрос с лета (сорвался из-за топливного кризиса). 5 декабря — единственное окно. Зал ждёт.", header_size=11, body_size=10, header_color=WHITE, body_color=WHITE)
# Правые фото-карточки маленькие для фона? Но у нас full bg, поэтому не нужны доп фото
tb(slide, Inches(7.20), Inches(1.75), Inches(5.30), Inches(4.15), "", size=10, color=WHITE)
# Добавим декоративные цифры
card(slide, Inches(7.20), Inches(2.00), Inches(5.50), Inches(1.10), fill=DARK, header="300–350", body="гостей — полный зал Foster", header_size=16, body_size=10, header_color=RED)
card(slide, Inches(7.20), Inches(3.30), Inches(5.50), Inches(1.10), fill=DARK, header="01:00", body="блэкаут + красный свет — главный момент", header_size=16, body_size=10, header_color=RED)
tb(slide, Inches(0.60), Inches(6.95), Inches(6), Inches(0.25), "Фото: красный лаунж Foster Club", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(6.95), Inches(1.70), Inches(0.25), "02 / 05", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 3 — ФИНАНСЫ КАКТУСЫ (объединённый, крупный) =====
slide = prs.slides.add_slide(BLANK)
add_bg(slide, IMG["dj_bar"])
tb(slide, Inches(0.60), Inches(0.35), Inches(9), Inches(0.35), "Ф И Н А Н С Ы   К А К Т У С Ы", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.60), Inches(0.75), Inches(9), Inches(0.90))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Инвестируют — рискуют — продают"; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=WHITE; p.font.bold=True
line_red(slide, Inches(0.62), Inches(1.45), Inches(1.0), Inches(0.08))
# Левая колонка — доход
card(slide, Inches(0.60), Inches(1.75), Inches(3.90), Inches(1.55), fill=DARK, header="ДОХОД", body="Билеты 300×3400 ₽ =\n1 020 000 ₽ (база)\nSOLD OUT 350×3400 =\n1 190 000 ₽", header_size=12, body_size=11)
card(slide, Inches(0.60), Inches(3.45), Inches(3.90), Inches(1.55), fill=DARK, header="ЦЕНЫ", body="3 000 ₽ старт →\n3 500 ₽ → 4 000 ₽\nв день события", header_size=12, body_size=11)
# Средняя — затраты
card(slide, Inches(4.70), Inches(1.75), Inches(3.90), Inches(1.55), fill=DARK, header="ЗАТРАТЫ 1.12 МЛН", body="Гонорар 550К + коктейли\n180К + шаттл 20К +\nлогистика 80К + промо 100К", header_size=12, body_size=10, header_color=WHITE)
card(slide, Inches(4.70), Inches(3.45), Inches(3.90), Inches(1.55), fill=DARK, header="ЕЩЁ ЗАТРАТЫ", body="Артисты 75К + тех 85К +\nохрана 30К + райдер\n= 1 120 040 ₽ итого", header_size=12, body_size=10, header_color=WHITE)
# Правая — результат
card(slide, Inches(8.90), Inches(1.75), Inches(3.90), Inches(1.55), fill=RED, header="БАЗА: –100 040 ₽", body="300 чел. — убыток\n1 020 000 – 1 120 040\nМотивация продавать!", header_size=11, body_size=11, header_color=WHITE, body_color=WHITE)
card(slide, Inches(8.90), Inches(3.45), Inches(3.90), Inches(1.55), fill=RED, header="SOLD OUT: +70 000 ₽", body="350 чел. — прибыль\n1 190 000 – 1 120 040\nВ плюс только так!", header_size=11, body_size=11, header_color=WHITE, body_color=WHITE)
# Нижняя плашка — вывод
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.60), Inches(5.30), Inches(12.10), Inches(0.90))
s.fill.solid(); s.fill.fore_color.rgb=WHITE; s.line.fill.background(); s.adjustments[0]=0.05
tb(slide, Inches(0.80), Inches(5.45), Inches(11.7), Inches(0.30), "ВЫВОД: Все риски на Кактусе — будут рвать продажи. Ваши 200 040 ₽ (180К коктейли + 20К шаттл) — фикс, независимо от явки.", size=10, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.60), Inches(6.95), Inches(7), Inches(0.25), "Фото: DJ-бар Foster Club", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(6.95), Inches(1.70), Inches(0.25), "03 / 05", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 4 — ФИНАНСЫ FOSTER (крупно, объединённый) =====
slide = prs.slides.add_slide(BLANK)
add_bg(slide, IMG["bar"])
tb(slide, Inches(0.60), Inches(0.35), Inches(9), Inches(0.35), "В А Ш И   Ф И Н А Н С Ы   •  FOSTER", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.60), Inches(0.75), Inches(9), Inches(0.90))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Ваша прибыль — 483 549 ₽"; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=WHITE; p.font.bold=True
line_red(slide, Inches(0.62), Inches(1.45), Inches(1.0), Inches(0.08))
# 3 большие карточки — затраты / доходы / прибыль
card(slide, Inches(0.60), Inches(1.80), Inches(3.90), Inches(1.85), fill=DARK, header="ЗАТРАТЫ 284 476 ₽", body="Персонал 165 470 ₽\nБар 92 006 ₽ (25% от 368К)\nКухня 27 000 ₽ (27% от 100К)\n0 ₽ — артисты/аренда/жильё", header_size=11, body_size=10)
card(slide, Inches(4.70), Inches(1.80), Inches(3.90), Inches(1.85), fill=DARK, header="ДОХОДЫ 768 025 ₽", body="Бар 368К + Кухня 100К\nКоктейли 180К (300×600)\nСтолы 50К + VIP 70К\n+ кальяны Макао", header_size=11, body_size=10)
card(slide, Inches(8.90), Inches(1.80), Inches(3.90), Inches(1.85), fill=RED, header="ПРИБЫЛЬ 483 549 ₽", body="768 025 – 284 476 =\n+483 549 ₽ чистыми\n~63% рентабельность\n+ кальяны", header_size=12, body_size=11, header_color=WHITE, body_color=WHITE)
# Нижний гарант — крупно
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.60), Inches(5.00), Inches(12.10), Inches(1.10))
s.fill.solid(); s.fill.fore_color.rgb=RED; s.line.fill.background(); s.adjustments[0]=0.06
tb(slide, Inches(0.80), Inches(5.15), Inches(11.7), Inches(0.35), "ГАРАНТИРОВАНО ДОГОВОРОМ — 330 000 ₽", size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.80), Inches(5.55), Inches(11.7), Inches(0.45), "180 000 ₽ компенсация коктейлей (300×600) + 50 000 ₽ депозиты столов + 70 000 ₽ VIP №1/№2 — фикс, независимо от явки. Даже при 200 гостях вы в плюсе.", size=9, color=WHITE, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.60), Inches(6.95), Inches(7), Inches(0.25), "Только Foster Club", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(6.95), Inches(1.70), Inches(0.25), "04 / 05", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

# ===== SLIDE 5 — ИТОГ =====
slide = prs.slides.add_slide(BLANK)
add_bg(slide, IMG["cover"])
tb(slide, Inches(0.60), Inches(0.35), Inches(9), Inches(0.35), "Г Л А В Н Ы Й   И Т О Г", size=11, color=RED, bold=True)
box=slide.shapes.add_textbox(Inches(0.60), Inches(0.75), Inches(9), Inches(1.0))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="Безопасно. Прибыльно."; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=WHITE; p.font.bold=True
p=tf.add_paragraph(); p.text="Полмиллиона за ночь."; p.font.name="Arial"; p.font.size=Pt(28); p.font.color.rgb=RED; p.font.bold=True
line_red(slide, Inches(0.62), Inches(1.50), Inches(1.0), Inches(0.08))
# Большая прибыль карточка
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.60), Inches(1.90), Inches(6.20), Inches(1.90))
s.fill.solid(); s.fill.fore_color.rgb=DARK; s.line.fill.background(); s.adjustments[0]=0.05
tb(slide, Inches(0.80), Inches(2.05), Inches(5.80), Inches(0.30), "ВАША ЧИСТАЯ ПРИБЫЛЬ ЗА ВЕЧЕР", size=10, color=RED, bold=True)
tb(slide, Inches(0.80), Inches(2.40), Inches(3.80), Inches(0.55), "483 549 ₽", size=28, color=WHITE, bold=True)
tb(slide, Inches(4.80), Inches(2.45), Inches(1.60), Inches(0.55), "63%\nрентаб.", size=12, color=RED, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.80), Inches(3.05), Inches(5.80), Inches(0.35), "768 025 ₽ доходы – 284 476 ₽ затраты  •  + кальяны Макао по факту", size=9, color=BEIGE2)
# Что фиксировать
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.10), Inches(1.90), Inches(5.40), Inches(1.90))
s.fill.solid(); s.fill.fore_color.rgb=BEIGE; s.line.fill.background(); s.adjustments[0]=0.05
tb(slide, Inches(7.30), Inches(2.05), Inches(5.00), Inches(0.30), "ЗАФИКСИРОВАТЬ В ДОГОВОРЕ", size=10, color=RED, bold=True, align=PP_ALIGN.CENTER)
box=slide.shapes.add_textbox(Inches(7.30), Inches(2.45), Inches(5.00), Inches(1.10))
tf=box.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.text="✓ 180 000 ₽ — коктейли (300×600, фикс)\n✓ 20 040 ₽ — шаттл (охрана/клининг/хостес)\n✓ Оплата 1–3 дня после события"; p.font.name="Arial"; p.font.size=Pt(9.5); p.font.color.rgb=DARK; p.line_spacing=1.30
# Нижний CTA
s=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.60), Inches(5.20), Inches(12.10), Inches(1.30))
s.fill.solid(); s.fill.fore_color.rgb=WHITE; s.line.fill.background(); s.adjustments[0]=0.05
tb(slide, Inches(0.80), Inches(5.35), Inches(11.7), Inches(0.35), "Вы не вкладываетесь в артистов, аренду и оборудование. Просто принимаете гостей и считаете прибыль.", size=11, color=DARK, bold=True, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.80), Inches(5.75), Inches(11.7), Inches(0.45), "Организатор мотивирован на SOLD OUT 350 чел. (только там его +70К). Ваши 330 000 ₽ гарантированы. Следующий шаг → согласовать договор → старт продаж 01.11 → 5 декабря.", size=9, color=MUTED, align=PP_ALIGN.CENTER)
tb(slide, Inches(0.60), Inches(6.95), Inches(7), Inches(0.25), "Только Foster Club • Фото: Foster Club Яндекс Карты", size=8, color=MUTED)
tb(slide, Inches(11.0), Inches(6.95), Inches(1.70), Inches(0.25), "05 / 05", size=10, color=WHITE, align=PP_ALIGN.RIGHT)

out = "/home/user/Foster_Gleb_Noemy_Presentation.pptx"
prs.save(out)
print(f"Saved PPTX {out} slides={len(prs.slides)}")
