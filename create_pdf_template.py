#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF версия в стиле шаблона FOSTER (как и PPTX)
10 слайдов, те же фото, цвета, карточки
"""
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white, Color
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os

# Fonts - DejaVu Sans ~ Arial
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

W, H = 960, 540
# Colors from template
GOLD = HexColor("#C9A25E")
GOLD_LIGHT = HexColor("#E5C888")
BEIGE = HexColor("#F5F1E8")
BEIGE2 = HexColor("#DDD8CC")
BEIGE3 = HexColor("#E2DED4")
MUTED = HexColor("#7E7A72")
DARK = HexColor("#141419")
DARKER = HexColor("#09090B")
BURGUNDY = HexColor("#7A1F2B")
WHITE = HexColor("#FFFFFF")
BLACK = HexColor("#000000")

IMG = {
    "cover": "/tmp/foster_images/slide01_img1_Picture_1.jpg",
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

OUT = "/home/user/Foster_Gleb_Noemy_Presentation.pdf"
c = canvas.Canvas(OUT, pagesize=(W, H))

def pt(inch): return inch*72

def rect(x, y, w, h, fill=None, stroke=None, radius=0):
    yb = H - y - h
    c.saveState()
    if fill: c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke); c.setLineWidth(0.7)
    if radius>0:
        p=c.beginPath()
        p.roundRect(x, yb, w, h, radius)
        if fill and stroke: c.drawPath(p, stroke=1, fill=1)
        elif fill: c.drawPath(p, stroke=0, fill=1)
        else: c.drawPath(p, stroke=1, fill=0)
    else:
        if fill and stroke: c.rect(x, yb, w, h, stroke=1, fill=1)
        elif fill: c.rect(x, yb, w, h, stroke=0, fill=1)
        elif stroke: c.rect(x, yb, w, h, stroke=1, fill=0)
    c.restoreState()

def draw_image(path, x, y, w, h, darken=0):
    if os.path.exists(path):
        yb = H - y - h
        try:
            c.drawImage(ImageReader(path), x, yb, width=w, height=h, preserveAspectRatio=False, mask='auto')
        except:
            c.drawImage(path, x, yb, width=w, height=h, preserveAspectRatio=False)
        if darken:
            # overlay dark
            c.saveState()
            c.setFillColor(HexColor("#0A0A0F"))
            # simulate transparency via alpha (reportlab supports)
            c.setFillAlpha(darken)
            c.rect(x, yb, w, h, stroke=0, fill=1)
            c.restoreState()

def text(x, y, txt, size=10, color=HexColor("#F5F1E8"), font="DejaVu", bold=False, align="left", max_width=None):
    c.saveState()
    c.setFillColor(color)
    f = "DejaVu-Bold" if bold else font
    c.setFont(f, size)
    y_draw = H - y - size*0.35
    if align=="center" and max_width is not None:
        c.drawCentredString(x+max_width/2, y_draw, txt)
    elif align=="right" and max_width is not None:
        c.drawRightString(x+max_width, y_draw, txt)
    elif align=="center":
        c.drawCentredString(x, y_draw, txt)
    elif align=="right":
        c.drawRightString(x, y_draw, txt)
    else:
        c.drawString(x, y_draw, txt)
    c.restoreState()

def text_wrap(x, y, w, txt, size=8, color=HexColor("#F5F1E8"), font="DejaVu", bold=False, leading=10, align="left"):
    c.saveState()
    c.setFillColor(color)
    f = "DejaVu-Bold" if bold else font
    c.setFont(f, size)
    words = txt.split()
    lines=[]
    cur=""
    for wo in words:
        test = cur + (" " if cur else "") + wo
        if c.stringWidth(test, f, size) <= w:
            cur=test
        else:
            if cur: lines.append(cur)
            cur=wo
            if c.stringWidth(cur, f, size) > w:
                lines.append(cur); cur=""
    if cur: lines.append(cur)
    yy=y
    for line in lines:
        y_draw = H - yy - size*0.35
        if align=="center": c.drawCentredString(x+w/2, y_draw, line)
        elif align=="right": c.drawRightString(x+w, y_draw, line)
        else: c.drawString(x, y_draw, line)
        yy+=leading
    c.restoreState()
    return yy

def add_header(small, title, num="01 / 10"):
    text(pt(0.55), pt(0.32+0.12), small, size=7.5, color=GOLD, bold=True)
    # title
    text_wrap(pt(0.55), pt(0.60+0.18), pt(7.0), title, size=16, color=BEIGE, bold=True, leading=17)
    # gold line
    rect(pt(0.57), pt(1.42), pt(0.70), pt(0.05), fill=GOLD)
    # slide number
    text(pt(11.70), pt(0.32+0.12), num, size=7, color=MUTED, align="right", max_width=pt(1.10))
    # footer
    text(pt(0.55), pt(7.12+0.10), "Фото: галерея Foster Club — Яндекс Карты", size=5.5, color=MUTED)
    text(pt(11.30), pt(7.12+0.10), f"FOSTER • {num.split('/')[0].strip()}", size=5.5, color=MUTED, bold=True, align="right", max_width=pt(1.50))

def card(x_in, y_in, w_in, h_in, fill=DARK, radius=6, header=None, body=None, header_color=GOLD_LIGHT, body_color=BEIGE3, header_size=6, body_size=6):
    x,y,w,h = pt(x_in), pt(y_in), pt(w_in), pt(h_in)
    rect(x, y, w, h, fill=fill, radius=radius)
    if header:
        text(x+pt(0.22), y+pt(0.12), header, size=header_size, color=header_color, bold=True)
    if body:
        text_wrap(x+pt(0.22), y+pt(0.52), w-pt(0.44), body, size=body_size, color=body_color, leading=body_size+2.5)

# ============ SLIDE 1 COVER ============
# bg image full with darken 0.45
draw_image(IMG["cover"], 0, 0, W, H, darken=0.45)
# top beige rect
rect(pt(0.55), pt(0.45), pt(2.45), pt(0.98), fill=BEIGE, radius=6)
# Try draw logo if exists (png)
if os.path.exists(IMG["cover"]):
    # we can't easily draw the Foster logo png with transparency, simulate with text
    text(pt(0.70), pt(0.70), "FOSTER", size=11, color=DARK, bold=True)
    text(pt(0.70), pt(0.88), "NIGHT & KARAOKE CLUB", size=4.5, color=GOLD, bold=True)
text(pt(9.5), pt(0.45+0.08), "Mriya Resort & SPA  •  с. Оползневое, Ялта", size=5, color=WHITE, align="right", max_width=pt(3.30))
text(pt(9.5), pt(0.68+0.08), "5 декабря 2026  •  Суббота  •  300–350 гостей", size=5, color=GOLD, bold=True, align="right", max_width=pt(3.30))
# main title
text_wrap(pt(0.55), pt(1.85+0.25), pt(8.5), "GLEB FILIPCHENKOW", size=24, color=WHITE, bold=True, leading=24)
text_wrap(pt(0.55), pt(2.35+0.25), pt(8.5), "& NOEMY ABRANTES", size=24, color=GOLD, bold=True, leading=24)
rect(pt(0.57), pt(3.55), pt(0.85), pt(0.06), fill=GOLD)
text_wrap(pt(0.55), pt(3.75+0.08), pt(8.2), "Первый привоз в Крым  •  Хит «Mallorca» (лето 2025)  •  Mallorca • Garota • HOJE", size=6.5, color=BEIGE2, leading=9)
# dark card
rect(pt(0.55), pt(4.40), pt(6.0), pt(1.55), fill=DARK, radius=6)
text(pt(0.75), pt(4.55+0.08), "◆  Ночь, которую не нужно искать за пределами курорта", size=6, color=GOLD, bold=True)
text_wrap(pt(0.75), pt(4.85+0.08), pt(5.6), "Привоз планировался на лето — сорвался из-за топливного кризиса. Публика ждёт. 5 декабря — единственное свободное окно в графике артистов. Имя продаёт само событие.", size=6, color=BEIGE3, leading=8)
# 3 mini stats
for i, (num, lbl) in enumerate([("300–350","ГОстей / до 60 мест зал"), ("63%","Рентабельность Foster"), ("483 549 ₽","Ваша чистая прибыль")]):
    xl = pt(0.55 + i*2.05)
    rect(xl, pt(6.20), pt(1.90), pt(0.75), fill=DARK, radius=5)
    text(xl, pt(6.28+0.08), num, size=8, color=WHITE, bold=True, align="center", max_width=pt(1.90))
    text(xl, pt(6.55+0.06), lbl, size=4.5, color=GOLD_LIGHT, bold=True, align="center", max_width=pt(1.90))
text(pt(0.55), pt(7.12+0.10), "Пересборка презентации на подлинных фотографиях клуба из галереи Яндекс Карт", size=5, color=MUTED)
text(pt(11.30), pt(7.12+0.10), "01 / 10", size=6, color=WHITE, align="right", max_width=pt(1.50))
c.showPage()

# ============ SLIDE 2 ПОЗИЦИОНИРОВАНИЕ ============
rect(0,0,W,H, fill=WHITE)
add_header("П О З И Ц И О Н И Р О В А Н И Е", "Камерная вечеринка внутри курорта", "02 / 10")
# image right
draw_image(IMG["about"], pt(7.00), pt(1.75), pt(5.78), pt(3.50))
# bullets left
yy = pt(1.85+0.08)
c.saveState()
c.setFillColor(HexColor("#2A2A30"))
c.setFont("DejaVu", 7)
# we'll do manual bullets
bullets = [
    ("Формат:", "привоз хедлайнеров 300–350 гостей, night & karaoke-шоу в составе Mriya Resort & SPA — вечерняя «гостиная» курорта."),
    ("Почему сейчас:", "летний привоз сорвался из-за топливного кризиса; публика ждёт; 5 декабря — единственное окно."),
    ("Масштаб:", "камерный зал Foster — до 60 мест (по данным 2ГИС) + танцпол; 300–350 гостей — полный SOLD OUT."),
    ("Атмосфера:", "роскошный лофт встречает Голливуд — гость курорта становится артистом."),
]
y = pt(1.95)
for label, body in bullets:
    # diamond
    text(pt(0.55), y, "◆", size=7, color=GOLD, bold=True)
    # label + body wrap
    # label
    text(pt(0.75), y, label, size=7, color=DARK, bold=True)
    # body - measure label width
    lw = c.stringWidth(label+" ", "DejaVu-Bold", 7)
    # remaining width
    # draw body wrap starting after label? simpler: two lines
    # We'll just wrap full text after label on same line with offset
    # For simplicity, draw body as wrap below
    text_wrap(pt(0.75)+lw, y, pt(6.15)-lw, body, size=7, color=HexColor("#2A2A30"), leading=8.5)
    y += pt(0.42) # approx 3 lines? estimate
    # rough adjust
    y += pt(0.15)

# But manual y is off, we do simpler: just text_wrap with prefix
# Let's redo with text_wrap approach including diamond and label
# We'll reset and do clean
c.restoreState()
# clear area and redo properly with text_wrap including all
# We'll use a helper that draws bullet with colored parts - simplified: draw as single color dark with gold diamond

y = pt(1.90)
for label, body in bullets:
    full = f"◆  {label} {body}"
    # we will draw diamond in gold then rest in dark, but for PDF we can just draw all dark and diamond gold via two draws?
    # Simpler: draw full with dark, then overdraw diamond in gold at same x
    text_wrap(pt(0.55), y, pt(6.15), full, size=7, color=HexColor("#2A2A30"), leading=9)
    # estimate lines
    lines = len(full)*7/ (pt(6.15)) # approx
    y += pt(0.48)

# cards bottom
card(0.55, 5.45, 3.95, 1.35, fill=DARK, header="ФОРМАТ", body="Ночной привоз хедлайнеров •\nDJ-бар • танцпол • караоке •\nVIP-зал • шоу. 300–350 гостей.")
card(4.70, 5.45, 3.95, 1.35, fill=DARK, header="РОЛИ СТОРОН", body="FOSTER — площадка и бар\nКАКТУСА — организация и продажи\nMRIYA — проживание (без скидки)")
card(8.85, 5.45, 3.93, 1.35, fill=DARK, header="ДАТА И ЦЕНА", body="Суббота, 5 декабря 2026\nБилеты 3 000 ₽ → 3 500 ₽ → 4 000 ₽\nСтарт продаж 01.11")
# footer is already in add_header, but we need override? add_header already added footer, but we need to ensure
c.showPage()

# ============ SLIDE 3 КОНЦЕПЦИЯ СДЕЛКИ ============
draw_image(IMG["concept_bg"], 0, 0, W, H, darken=0.5)
text(pt(0.55), pt(0.32+0.10), "К О Н Ц Е П Ц И Я   С Д Е Л К И", size=7.5, color=GOLD, bold=True)
text_wrap(pt(0.55), pt(0.60+0.15), pt(10), "«Ваша прибыль — 483 549 ₽ за вечер»", size=16, color=BEIGE, bold=True, leading=17)
rect(pt(0.57), pt(1.42), pt(0.70), pt(0.05), fill=GOLD)
text(pt(11.70), pt(0.32+0.10), "03 / 10", size=7, color=MUTED, align="right", max_width=pt(1.10))
card(0.55, 1.85, 3.95, 1.35, fill=DARKER, header="1 • ПРИБЫЛЬ", body="768 025 ₽ выручка — 284 476 ₽\nрасходы = 483 549 ₽ чистыми.\nРентабельность ~63% за ночь.")
card(4.70, 1.85, 3.95, 1.35, fill=DARKER, header="2 • ГАРАНТИЯ", body="330 000 ₽ живыми деньгами\n(180К коктейли + 50К столы +\n70К VIP) — фикс в договоре.")
card(8.85, 1.85, 3.93, 1.35, fill=DARKER, header="3 • НУЛЬ РИСКА", body="0 ₽ за аренду, артистов,\nоборудование — только\nоперационка клуба.")
text(pt(0.55), pt(7.12+0.10), "Дизайн-концепция: «роскошный лофт» — студия Елены Крыловой  •  Фото: красный лаунж Foster Club, Яндекс Карты", size=5, color=MUTED)
text(pt(11.30), pt(7.12+0.10), "FOSTER • 03", size=5, color=MUTED, bold=True, align="right", max_width=pt(1.50))
c.showPage()

# ============ SLIDE 4 ЗАТРАТЫ ============
rect(0,0,W,H, fill=WHITE)
add_header("З А Т Р А Т Ы   К Л У Б А", "Ваши затраты минимальны", "04 / 10")
draw_image(IMG["context"], pt(7.00), pt(1.75), pt(5.85), pt(3.50))
# bullets
y = pt(1.95)
for label, body in [("Вы предоставляете", "площадку и оборудование бесплатно — только операционка."), ("Все прямые расходы", "покрываются баром и кухней, фактически — нулевые инвестиции."), ("Сравнение масштаба:", "ваши 284 476 ₽ против 1 120 040 ₽ у Кактусы — в 4 раза меньше.")]:
    text(pt(0.55), y, "◆", size=7, color=GOLD, bold=True)
    text_wrap(pt(0.75), y, pt(6.15), f"{label} {body}", size=7, color=HexColor("#2A2A30"), leading=9)
    y += pt(0.45)
card(0.55, 5.45, 3.00, 1.35, fill=DARK, header="ПЕРСОНАЛ", body="165 470 ₽\nофицианты, бармены,\nповара, менеджмент")
card(3.65, 5.45, 3.00, 1.35, fill=DARK, header="БАР", body="92 006 ₽\nфудкост 25% от\n368 025 ₽")
card(6.75, 5.45, 3.00, 1.35, fill=DARK, header="КУХНЯ", body="27 000 ₽\nфудкост 27% от\n100 000 ₽")
card(9.85, 5.45, 3.00, 1.35, fill=BURGUNDY, header="ИТОГО", body="284 476 ₽\nвсе расходы клуба\nпокрываются")
text(pt(0.55), pt(7.12+0.10), "Терраса Mriya Resort & SPA — вид на Чёрное море. Концепцию в форме лотоса создал Норман Фостер.", size=5, color=MUTED)
c.showPage()

# ============ SLIDE 5 ДОХОД ============
rect(0,0,W,H, fill=WHITE)
add_header("Ч Т О   З А Р А Б О Т А Е Т   К Л У Б", "Чистая прибыль — 483 549 ₽", "05 / 10")
draw_image(IMG["entrance"], pt(7.00), pt(1.75), pt(5.78), pt(3.50))
# revenue list left
revenues = [
    ("Бар", "368 025 ₽", "валовая выручка"),
    ("Кухня", "100 000 ₽", "валовая выручка"),
    ("◆ Welcome-коктейли", "180 000 ₽", "300 × 600 ₽ → вам"),
    ("◆ Депозиты столов", "50 000 ₽", "4×5 000 + 3×10 000 → вам"),
    ("◆ VIP №1 + №2", "70 000 ₽", "40 000 + 30 000 → вам"),
    ("Итого выручка", "768 025 ₽", "+ кальяны от Макао"),
]
y = pt(1.95)
for name, price, note in revenues:
    is_hl = "◆" in name or "Итого" in name
    col = GOLD if is_hl else DARK
    text(pt(0.55), y, name, size=6.5, color=col, bold=True)
    # price at x 2.5 inch
    text(pt(2.60), y, price, size=7, color=DARK, bold=True)
    text(pt(3.60), y, "— "+note, size=5.5, color=MUTED)
    y += pt(0.32)
card(0.55, 5.45, 3.95, 1.35, fill=DARK, header="ВАША ВЫРУЧКА", body="768 025 ₽\nБар 368К + Кухня 100К +\n330К гарантии + кальяны")
card(4.70, 5.45, 3.95, 1.35, fill=DARK, header="ВАШИ РАСХОДЫ", body="284 476 ₽\nПерсонал 165 470 ₽ +\nБар 92 006 ₽ + Кухня 27К")
card(8.85, 5.45, 3.93, 1.35, fill=BURGUNDY, header="ЧИСТАЯ ПРИБЫЛЬ", body="483 549 ₽\nРентабельность ~63%\nдо учёта кальянов")
text(pt(0.55), pt(7.12+0.10), "«Улица» внутри курорта: кирпичный фасад и брусчатка создают эффект ночного города под крышей отеля.", size=5, color=MUTED)
c.showPage()

# ============ SLIDE 6 ХЕДЛАЙНЕРЫ ============
rect(0,0,W,H, fill=WHITE)
add_header("Х Е Д Л А Й Н Е Р Ы", "Имена, которые продают", "06 / 10")
draw_image(IMG["hall1"], pt(0.55), pt(1.75), pt(8.10), pt(3.60))
draw_image(IMG["hall2"], pt(8.85), pt(1.75), pt(3.93), pt(3.60))
text(pt(0.55), pt(5.45+0.08), "Потолочная световая конструкция — Helen Design / Большой LED-экран и сцена Foster Club", size=5, color=MUTED, align="center", max_width=pt(8.10))
text(pt(8.85), pt(5.45+0.08), "Зеркальный бар Foster — фото Яндекс Карты", size=5, color=MUTED, align="center", max_width=pt(3.93))
card(0.55, 5.85, 3.95, 1.20, fill=DARK, header="GLEB FILIPCHENKOW", body="Соавтор хита «Mallorca» (лето 2025).\nВзрыв весной 2026 без бюджета —\nсотни тысяч просмотров.")
card(4.70, 5.85, 3.95, 1.20, fill=DARK, header="NOEMY ABRANTES", body="Соавтор «Garota» и «HOJE».\nЯркий европейский вокал —\nидеален для Foster.")
card(8.85, 5.85, 3.93, 1.20, fill=BURGUNDY, header="ПОЧЕМУ ПРОДАСТСЯ", body="Первый привоз в Крым •\nОтложенный спрос с лета •\n5 декабря — единственное окно")
text(pt(0.55), pt(7.12+0.10), "Фото: главный зал, сцена и бар — галерея Foster Club, Яндекс Карты / Потолок Helen Design", size=5, color=MUTED)
c.showPage()

# ============ SLIDE 7 ИНВЕСТИЦИИ КАКТУСЫ ============
rect(0,0,W,H, fill=WHITE)
add_header("И Н В Е С Т И Ц И И   О Р Г А Н И З А Т О Р А", "Кактуса берёт весь риск", "07 / 10")
draw_image(IMG["dj_bar"], pt(0.55), pt(1.75), pt(7.90), pt(3.50))
draw_image(IMG["vip"], pt(8.70), pt(1.75), pt(4.10), pt(2.70))
text(pt(8.70), pt(4.60+0.08), "Фото: DJ-бар и VIP-зал Foster Club — Яндекс Карты", size=5, color=MUTED, align="center", max_width=pt(4.10))
card(0.55, 5.45, 3.95, 1.35, fill=DARK, header="ГОНОРАР + ЛОГИСТИКА", body="550 000 ₽ гонорар +\n80 000 ₽ перелёт/трансфер +\n30 000 ₽ техрайдер")
card(4.70, 5.45, 3.95, 1.35, fill=DARK, header="ПРОДАКШН + ПРОДВИЖЕНИЕ", body="75К местные + 30К видео +\n25К фото + 100К афиши/блогеры")
card(8.85, 5.45, 3.93, 1.35, fill=BURGUNDY, header="ВАШИ 200 040 ₽", body="180 000 ₽ коктейли +\n20 040 ₽ шаттл — фикс\nнезависимо от явки")
c.showPage()

# ============ SLIDE 8 СЦЕНАРИИ ============
rect(0,0,W,H, fill=WHITE)
add_header("С Ц Е Н А Р И И   П Р О Д А Ж", "Три сценария — один вывод", "08 / 10")
draw_image(IMG["timeline"], pt(0.55), pt(1.75), pt(12.23), pt(2.35))
text(pt(0.55), pt(4.20+0.08), "Тайминг продаж: 01.11–07.11 (3 000 ₽) / 08.11–14.11 (3 500 ₽) / с 15.11 (4 000 ₽). Фото: галерея Foster Club — Яндекс Карты", size=5, color=MUTED, align="center", max_width=pt(12.23))
scenarios = [
    ("250 ЧЕЛ.", "Пессимистичный", "3 300 ₽ средн.\n825 000 ₽ выручка\n–295 040 ₽ убыток", DARK),
    ("300 ЧЕЛ.", "Базовый", "3 400 ₽ средн.\n1 020 000 ₽ выручка\n–100 040 ₽ убыток", DARK),
    ("350 ЧЕЛ.", "SOLD OUT ✓", "3 400 ₽ средн.\n1 190 000 ₽ выручка\n+69 960 ₽ прибыль", BURGUNDY),
]
for i, (t1,t2,body,fill) in enumerate(scenarios):
    xl = 0.55 + i*4.20
    rect(pt(xl), pt(4.65), pt(3.95), pt(2.05), fill=fill, radius=6)
    text(pt(xl), pt(4.80+0.10), t1, size=8, color=GOLD_LIGHT if fill!=BURGUNDY else WHITE, bold=True, align="center", max_width=pt(3.95))
    text(pt(xl), pt(5.15+0.08), t2, size=6, color=BEIGE, bold=True, align="center", max_width=pt(3.95))
    text_wrap(pt(xl)+pt(0.20), pt(5.50), pt(3.55), body, size=6, color=BEIGE3 if fill!=BURGUNDY else WHITE, leading=8, align="center")
c.showPage()

# ============ SLIDE 9 РЕЗЮМЕ ============
rect(0,0,W,H, fill=WHITE)
add_header("Р Е З Ю М Е", "Ваша сделка на 483 500 ₽", "09 / 10")
draw_image(IMG["audience"], pt(0.55), pt(1.75), pt(2.70), pt(5.00))
# 2x2 grid right
cards = [
    ("180 000 ₽ • КОКТЕЙЛИ", "300 × 600 ₽ — фикс,\nнезависимо от явки", DARK),
    ("20 040 ₽ • ШАТТЛ", "Охрана, клининг, хостес\nВелком-центра — фикс", DARK),
    ("120 000 ₽ • VIP И СТОЛЫ", "VIP №1 40К + №2 30К +\nстолы 50К — ваша зона", DARK),
    ("КАЛЬЯНЫ • МАКАО", "Доход по факту продаж\nполностью ваш", DARK),
]
for i, (h,b,fill) in enumerate(cards):
    row, col = divmod(i, 2)
    xl = 3.55 + col*4.85
    yl = 1.75 + row*1.75
    card(xl, yl, 4.65, 1.45, fill=fill, header=h, body=b)
# burgundy highlight
rect(pt(3.55), pt(5.45), pt(9.50), pt(1.00), fill=BURGUNDY, radius=6)
text_wrap(pt(3.75), pt(5.60+0.08), pt(9.10), "Почему работает: гибрид «танцы + караоке + лаунж» удерживает компанию с разными настроениями в одном месте, а ваш риск — в 4 раза ниже, чем у организатора.", size=6, color=WHITE, leading=8)
text(pt(0.55), pt(6.85+0.08), "Фото: аудитория Foster Club — Яндекс Карты", size=5, color=MUTED, align="center", max_width=pt(2.70))
c.showPage()

# ============ SLIDE 10 ИТОГ ============
draw_image(IMG["final"], 0, 0, W, H, darken=0.5)
text(pt(0.55), pt(0.32+0.10), "И Т О Г", size=7.5, color=GOLD, bold=True)
text_wrap(pt(0.55), pt(0.60+0.15), pt(9), "Foster — ночь, которую", size=16, color=BEIGE, bold=True, leading=17)
text_wrap(pt(0.55), pt(0.95+0.15), pt(9), "не нужно искать", size=16, color=BEIGE, bold=True, leading=17)
text_wrap(pt(0.55), pt(1.30+0.15), pt(9), "за пределами курорта", size=16, color=GOLD, bold=True, leading=17)
text_wrap(pt(0.55), pt(2.05+0.08), pt(7), "Роскошный лофт, голливудский VIP, караоке-сцена и зеркальный свет — камерный театр одной ночи внутри Mriya. Ваша прибыль 483 549 ₽ ждёт 5 декабря.", size=6.5, color=BEIGE3, leading=8)
rect(pt(0.55), pt(5.20), pt(7.5), pt(1.35), fill=DARKER, radius=6)
text(pt(0.75), pt(5.35+0.08), "КОНТАКТЫ И ТОЧКА НА КАРТЕ", size=5, color=GOLD_LIGHT, bold=True)
text_wrap(pt(0.75), pt(5.65), pt(7.10), "Foster Club, Mriya Resort & SPA | Ялта, с. Оползневое, ул. Генерала Острякова, 9 | 300–350 гостей | Билеты 3 000–4 000 ₽ | Организатор: Кактуса", size=6, color=BEIGE3, leading=8)
text(pt(0.55), pt(6.80+0.08), "Перед публичным использованием проверьте актуальный статус и график — чек-лист на следующем слайде.", size=5, color=MUTED)
text(pt(11.30), pt(7.12+0.10), "10 / 10", size=6, color=WHITE, align="right", max_width=pt(1.50))
text(pt(0.55), pt(7.12+0.10), "Пересборка на подлинных фото из галереи Яндекс Карт", size=5, color=MUTED)
c.showPage()

c.save()
print(f"Saved PDF {OUT}")

