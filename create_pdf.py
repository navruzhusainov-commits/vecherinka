#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import os

# Register DejaVu for Cyrillic
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Serif', '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Serif-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf'))

# Colors
NAVY = HexColor("#0A1A2F")
NAVY2 = HexColor("#122E4B")
NAVY3 = HexColor("#1A3A5C")
GOLD = HexColor("#C9A96A")
GOLD_LIGHT = HexColor("#E8D5B5")
GOLD_DARK = HexColor("#A88A4A")
BEIGE = HexColor("#F6F1E7")
BEIGE2 = HexColor("#EDE6D6")
MUTED = HexColor("#8A9AB0")
LIGHT_GRAY = HexColor("#E5E9EF")
DARK_MUT = HexColor("#4A5A70")

# Slide size 16:9 HD-like: 960x540 pt (13.33x7.5 inch)
W, H = 960, 540  # points

OUT = "/home/user/Foster_Gleb_Noemy_Presentation.pdf"
c = canvas.Canvas(OUT, pagesize=(W, H))

def set_font(name, size):
    c.setFont(name, size)

def rect(x, y, w, h, fill=None, stroke=None, radius=0):
    # y is bottom in canvas, our design top-origin. Convert.
    # we will work in top-origin by translating: canvas y bottom = H - y - h
    yb = H - y - h
    if fill or stroke:
        c.saveState()
        if fill:
            c.setFillColor(fill)
        if stroke:
            c.setStrokeColor(stroke)
            c.setLineWidth(0.7)
        # rounded?
        if radius>0:
            p = c.beginPath()
            p.roundRect(x, yb, w, h, radius)
            if fill and stroke:
                c.setFillColor(fill); c.setStrokeColor(stroke); c.setLineWidth(0.7)
                c.drawPath(p, stroke=1, fill=1)
            elif fill:
                c.drawPath(p, stroke=0, fill=1)
            else:
                c.drawPath(p, stroke=1, fill=0)
        else:
            if fill and stroke:
                c.rect(x, yb, w, h, stroke=1, fill=1)
            elif fill:
                c.rect(x, yb, w, h, stroke=0, fill=1)
            else:
                c.rect(x, yb, w, h, stroke=1, fill=0)
        c.restoreState()

def text(x, y, txt, size=10, color=NAVY, font="DejaVu", align="left", bold=False, max_width=None):
    c.saveState()
    c.setFillColor(color)
    f = "DejaVu-Bold" if bold else font
    # handle bold+DejaVu mapping
    if bold and font=="DejaVu":
        f="DejaVu-Bold"
    c.setFont(f, size)
    # alignment
    if align=="center":
        if max_width:
            # center within max_width box starting at x
            tx = x + max_width/2
            c.drawCentredString(tx, H - y - size*0.3, txt)
        else:
            c.drawCentredString(x, H - y - size*0.3, txt)
    elif align=="right":
        if max_width:
            tx = x + max_width
            c.drawRightString(tx, H - y - size*0.3, txt)
        else:
            c.drawRightString(x, H - y - size*0.3, txt)
    else:
        c.drawString(x, H - y - size*0.3, txt)
    c.restoreState()

def text_wrap(x, y, w, txt, size=8, color=NAVY, font="DejaVu", bold=False, leading=10, align="left"):
    """Simple wrap: split by words, draw lines"""
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
            if cur:
                lines.append(cur)
            cur=wo
            # if single word too long, force
            if c.stringWidth(cur, f, size) > w:
                lines.append(cur)
                cur=""
    if cur:
        lines.append(cur)
    yy=y
    for line in lines:
        # alignment
        if align=="center":
            c.drawCentredString(x + w/2, H - yy - size*0.3, line)
        elif align=="right":
            c.drawRightString(x+w, H - yy - size*0.3, line)
        else:
            c.drawString(x, H - yy - size*0.3, line)
        yy+=leading
    c.restoreState()
    return yy

def circle(x,y,d,fill):
    yb = H - y - d
    c.setFillColor(fill)
    c.circle(x+d/2, yb+d/2, d/2, stroke=0, fill=1)

# =============== SLIDE 1 - COVER ===============
# bg navy
rect(0,0,W,H, fill=NAVY)
rect(0,0,W,5, fill=GOLD)
# eyebrow
text(55,28,"FOSTER  ×  КАКТУСА  ×  MRIYA RESORT", size=6, color=GOLD_LIGHT, bold=True)
rect(55,48,80,1.5, fill=GOLD)
text(55,62,"ПРЕЗЕНТАЦИЯ ПРОЕКТА", size=8, color=GOLD)
# title
c.setFillColor(white)
c.setFont("DejaVu-Bold", 28)
c.drawString(55, H- (85*1) - 10 , "GLEB FILIPCHENKOW") # manual y for top origin
# Actually use text function top y: y = distance from top
text(55,95,"GLEB FILIPCHENKOW", size=28, color=white, bold=True)
text(55,130,"& NOEMY ABRANTES", size=28, color=GOLD, bold=True)
rect(55,165,90,2.5, fill=GOLD)
text(55,185,"Первый привоз в Крым  •  Ночной клуб Foster  •  5 декабря 2026", size=8, color=BEIGE)
# 3 cards
cards = [("300–350","ГОСТЕЙ"), ("63%","РЕНТАБЕЛЬНОСТЬ"), ("483 549 ₽","ВАША ПРИБЫЛЬ")]
for i,(num,label) in enumerate(cards):
    xl=55 + i*145
    rect(xl,235,125,58, fill=HexColor("#132B4A"), stroke=HexColor("#2A4A6E"), radius=6)
    rect(xl,235,125,2.5, fill=GOLD, radius=0)
    text(xl+10,255,num, size=13, color=white, bold=True)
    text(xl+10,275,label, size=6, color=GOLD_LIGHT, bold=True)
rect(55,340,235,32, fill=white, radius=5)
text(65,355,"ДАТА: СУББОТА, 5 ДЕКАБРЯ 2026  •  ФОРМАТ: ПРИВОЗ ХЕДЛАЙНЕРОВ", size=6, color=NAVY, bold=True)

# right card
rect(530,30,360,340, fill=white, radius=10)
rect(530,30,360,200, fill=NAVY2, radius=0)
# circles stage
circle(650,95,95, fill=HexColor("#1E3D5E"))
circle(665,110,70, fill=HexColor("#2A4F7A"))
circle(685,130,32, fill=GOLD)
text(530,88,"✦  HEADLINERS NIGHT  ✦", size=6, color=GOLD_LIGHT, bold=True, align="center", max_width=360)
text(530,185,"GLEB FILIPCHENKOW & NOEMY", size=10, color=white, bold=True, align="center", max_width=360)
text(530,202,"Mallorca  •  Garota  •  HOJE", size=7, color=GOLD_LIGHT, align="center", max_width=360)
text(555,255,"ПЛОЩАДКА", size=6, color=MUTED, bold=True)
text(555,272,"FOSTER", size=13, color=NAVY, bold=True)
rect(680,262,135,24, fill=NAVY, radius=5)
text(680,278,"КРЫМ  •  300+ ГОСТЕЙ", size=6, color=white, bold=True, align="center", max_width=135)
text(555,300,"ОРГАНИЗАТОР", size=6, color=MUTED, bold=True)
text(555,317,"КАКТУСА — логистика, продакшн, продажи", size=7, color=NAVY2)
rect(555,334,320,1, fill=LIGHT_GRAY)
text(555,348,"БИЛЕТЫ ОТ 3 000 ₽", size=6, color=NAVY, bold=True)
text(680,348,"SOLD OUT 350 ЧЕЛ.", size=6, color=GOLD_DARK, bold=True, align="right", max_width=195)
text(530,380,"mriya resort  •  foster club  •  кактуса", size=6, color=DARK_MUT, align="center", max_width=360)

c.showPage()

# =============== SLIDE 2 - ABOUT ===============
rect(0,0,W,H, fill=BEIGE)
rect(0,0,W,4, fill=GOLD)
rect(0,4,330,536, fill=NAVY)
text(40,30,"01  —  О ПРОЕКТЕ", size=6, color=GOLD, bold=True)
text_wrap(40,50,230,"Вечеринка, которая уже продана", size=16, color=white, bold=True, leading=18)
rect(40,110,50,2, fill=GOLD)
text_wrap(40,120,230,"Привоз состоялся бы летом — сорвался из-за топливного кризиса. Публика ждёт. 5 декабря — единственное окно у артистов. Имя продаёт само событие.", size=7, color=BEIGE, leading=10)
stats=[("300–350","гостей  •  план загрузки"), ("3 000–4 000 ₽","цена билета  •  динамика"), ("SOLD OUT","цель кактусы  •  ваша защита")]
for i,(n,l) in enumerate(stats):
    yt=180 + i*65
    rect(40,yt,250,50, fill=HexColor("#112845"), stroke=HexColor("#1E3A5A"), radius=6)
    text(50,yt+13,n, size=11, color=GOLD_LIGHT, bold=True)
    text(125,yt+13,l, size=6, color=BEIGE)
    # second line of label
    # quick split
    parts=l.split("•")
    if len(parts)>1:
        text(125,yt+26,parts[1].strip(), size=5.5, color=GOLD_LIGHT)

text(360,30,"РОЛИ СТОРОН — КТО ЗА ЧТО ОТВЕЧАЕТ", size=7, color=NAVY, bold=True)
rect(360,42,35,2, fill=GOLD)
roles=[
    ("FOSTER","ПЛОЩАДКА","Предоставляете клуб,\nбар, кухню, персонал\nи инфраструктуру","0 ₽ аренды", NAVY),
    ("КАКТУСА","ОРГАНИЗАТОР","Берёт все риски:\nгонорары, логистика,\nпродакшн, продвижение","~1,12 млн ₽", GOLD),
    ("MRIYA","ПАРТНЁР","Проживание артистов\nбез скидки —\nчистый расчёт","Без нагрузки", NAVY2),
]
for i,(t1,t2,desc,price,col) in enumerate(roles):
    xl=360 + i*175
    rect(xl,55,160,135, fill=white, radius=6)
    rect(xl,55,160,4, fill=col, radius=0)
    text(xl+12,75,t1, size=8, color=col, bold=True)
    text(xl+12,88,t2, size=6, color=MUTED, bold=True)
    rect(xl+12,98,136,1, fill=LIGHT_GRAY)
    text_wrap(xl+12,108,136,desc, size=6.5, color=NAVY2, leading=8)
    rect(xl+12,152,136,24, fill=BEIGE, radius=4)
    text(xl+12,167,price, size=6.5, color=NAVY, bold=True, align="center", max_width=136)

rect(360,210,530,58, fill=white, radius=6)
rect(360,210,4,58, fill=GOLD)
text(375,225,"ВАШЕ ПРЕИМУЩЕСТВО", size=6, color=GOLD_DARK, bold=True)
text(375,243,"Вы не инвестируете в артистов.", size=9, color=NAVY, bold=True)
text(375,260,"Только операционка клуба, которая окупается баром.", size=6.5, color=MUTED)
rect(620,220,260,38, fill=NAVY, radius=5)
text(620,230,"ФИКСИРОВАННЫЙ ДОХОД ДО СТАРТА ВЕЧЕРА", size=5, color=GOLD_LIGHT, bold=True, align="center", max_width=260)
text(620,250,"330 000 ₽  живыми деньгами", size=10, color=white, bold=True, align="center", max_width=260)

text(360,285,"ТАЙМЛАЙН ПРОДАЖ БИЛЕТОВ", size=6, color=MUTED, bold=True)
tls=[("01–07 НОЯБРЯ","3 000 ₽","Старт"),("08–14 НОЯБРЯ","3 500 ₽","Рост"),("С 15 НОЯБРЯ","4 000 ₽","Финал")]
for i,(d,p,cc) in enumerate(tls):
    xl=360 + i*178
    if i==2:
        rect(xl,300,170,40, fill=NAVY, radius=5)
        text(xl+10,313,d, size=6, color=GOLD_LIGHT, bold=True, align="center", max_width=150)
        text(xl+10,328,p, size=10, color=white, bold=True, align="center", max_width=150)
        text(xl+10,342,cc, size=6, color=BEIGE, align="center", max_width=150)
    else:
        rect(xl,300,170,40, fill=white, radius=5)
        text(xl+10,313,d, size=6, color=MUTED, bold=True, align="center", max_width=150)
        text(xl+10,328,p, size=10, color=NAVY, bold=True, align="center", max_width=150)
        text(xl+10,342,cc, size=6, color=MUTED, align="center", max_width=150)

c.showPage()

# =============== SLIDE 3 - COSTS FOSTER ===============
rect(0,0,W,H, fill=white)
rect(0,0,W,4, fill=GOLD)
text(45,22,"02  —  ЗАТРАТЫ FOSTER", size=6, color=GOLD_DARK, bold=True)
text(45,38,"Ваши затраты минимальны", size=18, color=NAVY, bold=True)
text(45,62,"Площадка и оборудование — бесплатно. Только операционка.", size=7, color=MUTED)

rect(45,80,420,265, fill=white, stroke=LIGHT_GRAY, radius=6)
rect(45,80,420,28, fill=NAVY, radius=0)
text(60,98,"СТАТЬЯ РАСХОДОВ", size=6, color=GOLD_LIGHT, bold=True)
text(340,98,"СУММА", size=6, color=GOLD_LIGHT, bold=True, align="right", max_width=110)

rows=[
    ("Персонал клуба","официанты, бармены, повара, менеджмент","165 470 ₽"),
    ("Бар","фудкост 25% от выручки 368 025 ₽","92 006 ₽"),
    ("Кухня","фудкост 27% от выручки 100 000 ₽","27 000 ₽"),
    ("Скидка на проживание","отменена по договорённости","0 ₽"),
]
for i,(t1,t2,price) in enumerate(rows):
    yb=108+ i*48
    if i%2==0:
        rect(45,yb,420,48, fill=BEIGE)
    circle(62,yb+16,5, fill=GOLD if i<3 else MUTED)
    text(75,yb+13,t1, size=7, color=NAVY, bold=True)
    text(75,yb+26,t2, size=5.5, color=MUTED)
    text(340,yb+16,price, size=9, color=NAVY, bold=True, align="right", max_width=110)
    if i<3:
        rect(60,yb+40,400,0.7, fill=LIGHT_GRAY)

rect(45,305,420,48, fill=NAVY, radius=0)
# round bottom
text(60,322,"ИТОГО ЗАТРАТЫ КЛУБА", size=7, color=GOLD_LIGHT, bold=True)
text(60,338,"полностью покрываются доходами", size=5.5, color=BEIGE)
text(340,327,"284 476 ₽", size=14, color=white, bold=True, align="right", max_width=110)

rect(500,80,380,115, fill=BEIGE, radius=8)
text(520,98,"ЧТО ЭТО ЗНАЧИТ НА ПРАКТИКЕ", size=6, color=GOLD_DARK, bold=True)
text(520,115,"Вы не платите ни за аренду, ни за артистов.", size=8, color=NAVY, bold=True)
text_wrap(520,132,340,"Ваши 284 тыс. ₽ — это только себестоимость бара и кухни, которая уже заложена в выручку. Фактически — нулевые инвестиции.", size=6.5, color=MUTED, leading=9)

mini=[("0 ₽","Аренда","площадка ваша"),("0 ₽","Гонорар","платит Кактуса"),("0 ₽","Оборудование","уже на месте")]
for i,(num,label,sub) in enumerate(mini):
    xl=500 + i*130
    rect(xl,205,115,55, fill=white, stroke=LIGHT_GRAY, radius=5)
    text(xl+10,220,num, size=10, color=GOLD, bold=True, align="center", max_width=95)
    text(xl+10,236,label, size=6, color=NAVY, bold=True, align="center", max_width=95)
    text(xl+10,248,sub, size=5, color=MUTED, align="center", max_width=95)

rect(500,278,380,70, fill=NAVY2, radius=8)
text(520,290,"СРАВНЕНИЕ МАСШТАБА", size=5, color=GOLD_LIGHT, bold=True)
text(520,308,"ВАШИ ЗАТРАТЫ", size=5, color=BEIGE, bold=True)
rect(520,318,120,15, fill=GOLD, radius=3)
text(525,328,"284 тыс.", size=6, color=NAVY, bold=True)
text(520,340,"ЗАТРАТЫ КАКТУСЫ", size=5, color=BEIGE, bold=True)
rect(520,350,330,15, fill=white, radius=3)
text(525,360,"1 120 тыс.  •  в 4 раза больше", size=6, color=NAVY, bold=True)
text(500,390,"Весь финансовый риск — на организаторе, не на вас.", size=6, color=GOLD_LIGHT, align="center", max_width=380)

c.showPage()

# =============== SLIDE 4 - INCOME FOSTER ===============
rect(0,0,W,H, fill=NAVY)
rect(0,0,W,4, fill=GOLD)
text(45,22,"03  —  ЧТО ЗАРАБОТАЕТ FOSTER", size=6, color=GOLD, bold=True)
text(45,38,"Ваша чистая прибыль — 483 549 ₽", size=18, color=white, bold=True)
rect(45,60,50,2, fill=GOLD)
text(105,60,"Рентабельность ~63% за один вечер", size=7, color=BEIGE)

rect(45,80,410,285, fill=white, radius=8)
rect(45,80,410,26, fill=BEIGE, radius=0)
text(60,97,"СТАТЬЯ ДОХОДА", size=6, color=NAVY, bold=True)
text(315,97,"СУММА", size=6, color=NAVY, bold=True, align="right", max_width=110)

revenues=[
    ("Бар","валовая выручка","368 025 ₽",False),
    ("Кухня","валовая выручка","100 000 ₽",False),
    ("Коктейли Welcome","300 × 600 ₽ компенсация","180 000 ₽",True),
    ("Депозиты столов","4×5 000 + 3×10 000","50 000 ₽",True),
    ("VIP №1","депозит","40 000 ₽",True),
    ("VIP №2","депозит","30 000 ₽",True),
]
for i,(t1,t2,price,hl) in enumerate(revenues):
    yb=106 + i*32
    if hl:
        rect(45,yb,410,32, fill=HexColor("#FFF7E0"))
        rect(45,yb,4,32, fill=GOLD)
    text(60,yb+10,t1, size=7, color=NAVY, bold=True)
    text(175,yb+10,t2, size=5.5, color=MUTED)
    text(315,yb+10,price, size=8, color=NAVY, bold=True, align="right", max_width=110)
    if hl:
        circle(442,yb+10,7, fill=GOLD)
        # check mark
        c.setFillColor(white)
        c.setFont("DejaVu-Bold",5)
        c.drawCentredString(442+3.5, H - yb -13, "✓")

rect(45,300,410,38, fill=NAVY2, radius=0)
text(60,315,"ИТОГО ВЫРУЧКА", size=6, color=GOLD_LIGHT, bold=True)
text(60,330,"+ кальяны от Макао по факту", size=5, color=BEIGE)
text(290,318,"768 025 ₽", size=12, color=white, bold=True, align="right", max_width=150)
text(315,330,"+ КАЛЬЯНЫ", size=5, color=GOLD_LIGHT, align="right", max_width=110)

text(485,85,"ФИНАНСОВЫЙ РЕЗУЛЬТАТ ЗА ВЕЧЕР", size=6, color=GOLD, bold=True)
rect(485,98,380,115, fill=white, radius=8)
rect(485,98,380,4, fill=GOLD)
text(505,115,"ЧИСТАЯ ПРИБЫЛЬ", size=6, color=MUTED, bold=True)
text(505,135,"483 549 ₽", size=18, color=NAVY, bold=True)
text(505,155,"768 025 – 284 476 = ваш плюс", size=6, color=MUTED)
rect(730,115,115,48, fill=NAVY, radius=5)
text(730,128,"~63%", size=12, color=GOLD, bold=True, align="center", max_width=115)
text(730,148,"РЕНТАБЕЛЬНОСТЬ", size=5, color=white, bold=True, align="center", max_width=115)

text(505,178,"КАК ФОРМИРУЕТСЯ ПРИБЫЛЬ", size=5, color=MUTED, bold=True)
rect(505,188,340,17, fill=GOLD, radius=3)
text(515,199,"ВЫРУЧКА  768 025 ₽", size=6, color=NAVY, bold=True)
rect(505,210,125,15, fill=white, stroke=LIGHT_GRAY, radius=3)
text(515,220,"РАСХОДЫ 284 476 ₽", size=5.5, color=MUTED, bold=True)
text(505,235,"▼  ЧИСТАЯ ПРИБЫЛЬ 483 549 ₽  •  остается у вас", size=6, color=GOLD, bold=True)

rect(485,250,380,73, fill=HexColor("#112845"), radius=8)
rect(485,250,380,2, fill=GOLD)
text(505,265,"⚡  ГАРАНТИРОВАННЫЕ ДЕНЬГИ ЕЩЁ ДО БАРА", size=6, color=GOLD, bold=True)
text_wrap(505,280,190,"330 000 ₽ живыми деньгами: коктейли + столы + VIP — не зависят от явки при правильном договоре.", size=6.5, color=BEIGE, leading=8)
rect(705,268,135,38, fill=GOLD, radius=5)
text(705,280,"180К + 50К + 70К", size=6, color=NAVY, bold=True, align="center", max_width=135)
text(705,295,"330 000 ₽", size=9, color=NAVY, bold=True, align="center", max_width=135)

rect(485,335,380,35, fill=BEIGE, radius=6)
text(495,345,"💡", size=10, align="center", max_width=20)
text_wrap(525,345,320,"Даже при 200 гостях вы в плюсе — бар и кухня уже покрывают все расходы.", size=6.5, color=NAVY, bold=True, leading=8)

c.showPage()

# =============== SLIDE 5 - HEADLINERS ===============
rect(0,0,W,H, fill=white)
rect(0,0,W,4, fill=GOLD)
text(45,22,"04  —  ХЕДЛАЙНЕРЫ", size=6, color=GOLD_DARK, bold=True)
text(45,38,"Имена, которые продают", size=18, color=NAVY, bold=True)
text(45,62,"Вирусный хит без бюджета • Первый приезд в Крым • Ожидание с лета", size=7, color=MUTED)

rect(45,80,430,165, fill=BEIGE, radius=8)
rect(45,80,160,165, fill=NAVY, radius=0)
circle(75,105,80, fill=HexColor("#1E3D5E"))
text(75,145,"G F", size=13, color=GOLD, bold=True, align="center", max_width=80)
text(50,195,"ПРОДЮСЕР  •  MALLORCA", size=5, color=GOLD_LIGHT, bold=True, align="center", max_width=150)
text(50,210,"Фото: промо артиста", size=5, color=MUTED, align="center", max_width=150)
text(220,97,"GLEB FILIPCHENKOW", size=9, color=NAVY, bold=True)
rect(220,110,35,2, fill=GOLD)
text_wrap(220,118,210,"Соавтор вирусного трека «Mallorca» — лето 2025. Взрыв весной 2026 без бюджета: трек звучал на всех московских вечеринках, видео — сотни тысяч просмотров.", size=6.5, color=NAVY2, leading=8)
text_wrap(220,162,210,"Дальше — релиз с Ольгой Серябкиной. Карьера на взлёте, интерес максимальный.", size=6, color=MUTED, leading=8)
rect(220,195,210,24, fill=white, radius=4)
text(228,210,"✦  Трек Mallorca — главный хит вечеринок Москвы 2026", size=5.5, color=NAVY, bold=True)

rect(500,80,380,165, fill=white, stroke=LIGHT_GRAY, radius=8)
rect(500,80,135,165, fill=NAVY2, radius=0)
circle(535,105,75, fill=NAVY)
text(535,143,"N A", size=12, color=GOLD, bold=True, align="center", max_width=75)
text(505,195,"ВОКАЛ  •  GAROTA / HOJE", size=5, color=GOLD_LIGHT, bold=True, align="center", max_width=125)
text(650,97,"NOEMY ABRANTES", size=9, color=NAVY, bold=True)
rect(650,110,35,2, fill=GOLD)
text_wrap(650,118,210,"Соавтор релизов «Garota» и «HOJE». Яркий вокал, европейское звучание, идеально для ночной атмосферы Foster.", size=6.5, color=NAVY2, leading=8)
rect(650,168,210,35, fill=BEIGE, radius=4)
text(660,178,"ПОЧЕМУ ДУЭТ РАБОТАЕТ", size=5, color=GOLD_DARK, bold=True)
text_wrap(660,188,190,"Мужской грув + женский вокал = полный танцпол", size=5.5, color=NAVY, leading=7)

text(45,265,"ПОЧЕМУ ЭТО СОБЫТИЕ ПРОДАСТСЯ", size=6, color=NAVY, bold=True)
reasons=[
    ("01","ПЕРВЫЙ ПРИВОЗ","В Крыму не выступали\nникогда — эффект\nэксклюзива", GOLD),
    ("02","ОТЛОЖЕННЫЙ СПРОС","Летний привоз сорвался\nиз-за кризиса —\nпублика ждёт", NAVY),
    ("03","ОКНО В ГРАФИКЕ","Ноябрь расписан,\n5 декабря — единственная\nсвободная дата", NAVY2),
    ("04","ВИРАЛЬНОСТЬ","Сотни тысяч просмотров\nбез вложений\nсарафан работает", GOLD_DARK),
]
for i,(num,title,desc,col) in enumerate(reasons):
    xl=45 + i*225
    rect(xl,278,210,82, fill=white, stroke=LIGHT_GRAY, radius=6)
    text(xl+10,295,num, size=11, color=col, bold=True)
    text(xl+35,297,title, size=6, color=NAVY, bold=True)
    rect(xl+10,310,190,1, fill=LIGHT_GRAY)
    text_wrap(xl+10,320,190,desc, size=6, color=MUTED, leading=7)

rect(45,375,870,28, fill=NAVY, radius=5)
text(65,392,"⚡  ИМЯ ПРОДАЁТ САМО СОБЫТИЕ — ВАМ ОСТАЁТСЯ ТОЛЬКО ПРИНЯТЬ ГОСТЕЙ", size=6.5, color=GOLD_LIGHT, bold=True)
rect(680,380,215,18, fill=GOLD, radius=4)
text(680,392,"350 БИЛЕТОВ  •  АЖИОТАЖ ОБЕСПЕЧЕН", size=6, color=NAVY, bold=True, align="center", max_width=215)

c.showPage()

# =============== SLIDE 6 - COSTS KAKTUSA ===============
rect(0,0,W,H, fill=BEIGE)
rect(0,0,W,4, fill=GOLD)
text(45,22,"05  —  ЗОНА ОТВЕТСТВЕННОСТИ КАКТУСЫ", size=6, color=GOLD_DARK, bold=True)
text(45,38,"Организатор берёт весь риск", size=18, color=NAVY, bold=True)
text(45,62,"Вы не участвуете в этих расходах — они мотивируют Кактусу продавать", size=7, color=MUTED)

rect(45,80,480,265, fill=white, stroke=LIGHT_GRAY, radius=8)
rect(45,80,480,26, fill=NAVY, radius=0)
text(60,97,"СТАТЬЯ РАСХОДОВ", size=6, color=GOLD_LIGHT, bold=True)
text(360,97,"СУММА", size=6, color=GOLD_LIGHT, bold=True, align="right", max_width=150)
rows2=[
    ("Гонорар хедлайнеров","550 000 ₽","фикс"),
    ("Перелёт / поезд / трансфер","~ 80 000 ₽","допущение"),
    ("Технический райдер","~ 30 000 ₽","сверх клуба"),
    ("Местные артисты","75 000 ₽","среднее"),
    ("Видеосъёмка и монтаж","30 000 ₽",""),
    ("Фотоотчёт","25 000 ₽",""),
    ("Продвижение, блогеры, афиши","~ 100 000 ₽","допущение"),
    ("Внешняя охрана","~ 30 000 ₽","сверх шаттла"),
    ("Компенсация коктейлей","180 000 ₽","→ вам"),
    ("Оплата шаттла","20 040 ₽","→ вам"),
]
for i,(name,price,note) in enumerate(rows2):
    yb=106 + i*22
    if i in [8,9]:
        rect(45,yb,480,22, fill=HexColor("#FFF7E0"))
        rect(45,yb,4,22, fill=GOLD)
    elif i%2==0:
        rect(45,yb,480,22, fill=HexColor("#F9F6F0"))
    text(60,yb+9,name, size=6, color=NAVY, bold=True if i in [0,8,9] else False)
    if note:
        text(360,yb+9,note, size=5, color=MUTED, align="right", max_width=60)
    text(380,yb+9,price, size=6.5, color=GOLD if i in [8,9] else NAVY, bold=True, align="right", max_width=110)

rect(45,330,480,35, fill=NAVY, radius=0)
text(60,345,"ИТОГО ИНВЕСТИЦИИ КАКТУСЫ", size=6, color=GOLD_LIGHT, bold=True)
text(60,358,"все риски на организаторе", size=5, color=BEIGE)
text(360,350,"~ 1 120 040 ₽", size=12, color=white, bold=True, align="right", max_width=150)

rect(560,80,330,125, fill=NAVY, radius=8)
text(580,98,"ЧТО ЭТО ЗНАЧИТ ДЛЯ ВАС", size=6, color=GOLD, bold=True)
text_wrap(580,112,290,"Кактуса вложит больше миллиона", size=13, color=white, bold=True, leading=14)
rect(580,148,35,2, fill=GOLD)
text_wrap(580,158,290,"Чтобы окупиться, им нужно продать каждый билет. Они будут рвать рынок — а вы просто получаете гостей.", size=6.5, color=BEIGE, leading=8)

rect(560,220,330,58, fill=white, radius=8)
text(570,233,"💰", size=10, align="center", max_width=20)
text(595,233,"ВАШИ 200 040 ₽ ГАРАНТИРОВАНЫ", size=6, color=NAVY, bold=True)
text_wrap(595,248,260,"180 000 ₽ за коктейли + 20 040 ₽ за шаттл — фикс в договоре", size=6, color=MUTED, leading=7)
rect(580,268,290,1, fill=LIGHT_GRAY)
text(580,278,"Независимо от продаж билетов — это ваши деньги", size=5.5, color=GOLD_DARK)

rect(560,290,330,75, fill=BEIGE, radius=8)
text(580,305,"СРАВНЕНИЕ РИСКОВ", size=6, color=GOLD_DARK, bold=True)
text(580,323,"FOSTER", size=5, color=NAVY, bold=True)
rect(580,333,85,12, fill=NAVY, radius=3)
text(583,341,"284К", size=5, color=white, bold=True)
text(680,340,"← минимальный", size=5, color=MUTED)
text(580,352,"КАКТУСА", size=5, color=NAVY, bold=True)
rect(580,362,265,12, fill=GOLD, radius=3)
text(583,370,"1 120К  •  ×4", size=5, color=NAVY, bold=True)

c.showPage()

# =============== SLIDE 7 - SCENARIOS ===============
rect(0,0,W,H, fill=white)
rect(0,0,W,4, fill=GOLD)
text(45,22,"06  —  ПОТЕНЦИАЛ ОРГАНИЗАТОРА", size=6, color=GOLD_DARK, bold=True)
text(45,38,"Кактуса в зоне риска — ваша защита", size=16, color=NAVY, bold=True)
text(45,60,"Они выходят в плюс только при SOLD OUT. Значит — будут продавать до последнего билета.", size=7, color=MUTED)

scenarios=[
    ("ПЕССИМИСТИЧНЫЙ","250 чел.","3 300 ₽","825 000 ₽","– 295 040 ₽","Убыток", HexColor("#E0E0E0"), MUTED, NAVY, False),
    ("БАЗОВЫЙ","300 чел.","3 400 ₽","1 020 000 ₽","– 100 040 ₽","Убыток", HexColor("#EDE6D6"), NAVY2, NAVY, False),
    ("ОПТИМИСТИЧНЫЙ","350 чел.","3 400 ₽","1 190 000 ₽","+ 69 960 ₽","Прибыль • SOLD OUT", GOLD, white, white, True),
]
for i,(title, guests, avg, rev, result, status, bg, col_text, col_bg, is_best) in enumerate(scenarios):
    xl=45 + i*295
    card_col = NAVY if is_best else white
    border = GOLD if is_best else LIGHT_GRAY
    rect(xl,80,275,205, fill=card_col, stroke=border, radius=8)
    if is_best:
        rect(xl,80,275,26, fill=GOLD, radius=0)
        text(xl+20,97,"✦  SOLD OUT  •  ЦЕЛЬ", size=6, color=NAVY, bold=True, align="center", max_width=235)
        text(xl+20,120,title, size=6, color=GOLD_LIGHT, bold=True, align="center", max_width=235)
    else:
        text(xl+20,97,title, size=6, color=MUTED, bold=True, align="center", max_width=235)
    text(xl+20,138,guests, size=14, color=NAVY if not is_best else white, bold=True, align="center", max_width=235)
    text(xl+20,155,"средняя цена "+avg, size=5, color=MUTED if not is_best else BEIGE, align="center", max_width=235)
    rect(xl+35,168,205,1, fill=LIGHT_GRAY if not is_best else HexColor("#2A4A6E"))
    text(xl+20,183,"ВЫРУЧКА", size=5, color=MUTED if not is_best else GOLD_LIGHT, bold=True, align="center", max_width=117)
    text(xl+138,183,"РЕЗУЛЬТАТ", size=5, color=MUTED if not is_best else GOLD_LIGHT, bold=True, align="center", max_width=117)
    text(xl+20,198,rev, size=8, color=NAVY if not is_best else white, bold=True, align="center", max_width=117)
    res_col = HexColor("#C0392B") if "–" in result else HexColor("#1E8A49")
    text(xl+138,198,result, size=8, color=res_col if not is_best else HexColor("#1E8A49"), bold=True, align="center", max_width=117)
    status_col = HexColor("#C0392B") if "Убыток" in status else HexColor("#1E8A49")
    text(xl+20,215,status, size=5, color=status_col if not is_best else HexColor("#7ED3A0"), bold=True, align="center", max_width=235)
    # bar
    bar_w = [65,85,110][i]
    rect(xl+35,230,185,15, fill=LIGHT_GRAY, radius=3)
    rect(xl+35,230,bar_w,15, fill=GOLD if is_best else HexColor("#C9A96A"), radius=3)
    rect(xl+35+185,225,1.5,25, fill=HexColor("#C0392B"))
    text(xl+155,250,"1.12М", size=4.5, color=HexColor("#C0392B"))

rect(45,305,870,50, fill=BEIGE, radius=8)
text(70,320,"💡", size=11, align="center", max_width=20)
text(100,320,"ВЫВОД ДЛЯ ВАС: КАКТУСА МАКСИМАЛЬНО МОТИВИРОВАНА", size=7, color=NAVY, bold=True)
text_wrap(100,335,500,"Их прибыль только при 350 гостях. Они будут продавать агрессивно — а ваши 330 000 ₽ зафиксированы в договоре и не зависят от их результата.", size=6, color=MUTED, leading=7)
rect(650,315,215,28, fill=NAVY, radius=5)
text(650,332,"ИХ РИСК  →  ВАША ВЫГОДА", size=6, color=GOLD_LIGHT, bold=True, align="center", max_width=215)

c.showPage()

# =============== SLIDE 8 - RESUME ===============
rect(0,0,W,H, fill=NAVY)
rect(0,0,W,4, fill=GOLD)
text(45,22,"07  —  РЕЗЮМЕ", size=6, color=GOLD, bold=True)
text(45,38,"Ваша сделка на 483 500 ₽", size=18, color=white, bold=True)
text(45,60,"Минимальный риск • Максимальная отдача за один вечер", size=7, color=BEIGE)

benefits=[
    ("483 549 ₽","ЧИСТАЯ ПРИБЫЛЬ","за вечер + кальяны","63% рентабельность"),
    ("330 000 ₽","ГАРАНТИРОВАНО","до старта вечеринки","фикс в договоре"),
    ("0 ₽","ВАШИ РИСКИ","не платите за артистов","только операционка"),
]
for i,(num,label,sub,foot) in enumerate(benefits):
    xl=45 + i*165
    rect(xl,80,150,88, fill=white, radius=8)
    rect(xl,80,150,3, fill=GOLD)
    text(xl+10,100,num, size=11, color=NAVY, bold=True, align="center", max_width=130)
    text(xl+10,118,label, size=6, color=NAVY, bold=True, align="center", max_width=130)
    text(xl+10,130,sub, size=5, color=MUTED, align="center", max_width=130)
    rect(xl+10,142,130,1, fill=LIGHT_GRAY)
    text(xl+10,150,foot, size=5, color=GOLD_DARK, bold=True, align="center", max_width=130)

rect(45,185,520,195, fill=white, radius=8)
rect(45,185,520,26, fill=BEIGE)
text(60,202,"✓  ЧТО ЗАФИКСИРОВАТЬ В ДОГОВОРЕ", size=7, color=NAVY, bold=True)
text(350,202,"СРОК ОПЛАТЫ: 1–3 РАБОЧИХ ДНЯ", size=5, color=GOLD_DARK, bold=True, align="right", max_width=160)
checks=[
    ("180 000 ₽","Компенсация за коктейли","300 × 600 ₽ — фикс, независимо от явки"),
    ("20 040 ₽","Компенсация за шаттл","охрана, клининг, хостес Велком-центра"),
    ("120 000 ₽","VIP и столы — ваша зона","VIP №1 40К + VIP №2 30К + столы 50К"),
    ("Кальяны","Доход от Макао","по факту продаж — полностью ваш"),
]
for i,(price,title,desc) in enumerate(checks):
    yb=215 + i*35
    rect(60,yb,18,18, fill=NAVY if i<3 else GOLD, radius=3)
    c.setFillColor(white)
    c.setFont("DejaVu-Bold",6)
    c.drawCentredString(60+9, H - yb -11, "✓")
    text(85,yb+4,price, size=7, color=NAVY, bold=True)
    text(175,yb+4,title, size=6.5, color=NAVY, bold=True)
    text(175,yb+14,desc, size=5.5, color=MUTED)
    if i<3:
        rect(60,yb+30,490,0.7, fill=LIGHT_GRAY)
rect(60,360,490,17, fill=HexColor("#FFF7E0"), radius=4)
text(70,371,"⚠️  Важно: зафиксируйте суммы как фиксированные обязательства Кактусы — их риски не должны переходить на вас.", size=5.5, color=NAVY, bold=True)

rect(590,80,310,300, fill=BEIGE, radius=8)
text(610,98,"СЛЕДУЮЩИЕ ШАГИ", size=6, color=GOLD_DARK, bold=True)
steps=[
    ("1","СОГЛАСОВАТЬ","Утвердить дату и условия с Кактуса и Mriya"),
    ("2","ПОДПИСАТЬ","Договор с фиксацией 330 000 ₽ и сроков оплаты"),
    ("3","ЗАПУСТИТЬ","Продажи билетов с 01.11 продвижение и афиши"),
    ("4","ПРОВЕСТИ","5 декабря — вечеринка на 300+ гостей"),
]
for i,(num,title,desc) in enumerate(steps):
    yb=118 + i*55
    circle(610,yb,17, fill=NAVY)
    c.setFillColor(white)
    c.setFont("DejaVu-Bold",7)
    c.drawCentredString(610+8.5, H - yb -11, num)
    text(635,yb+2,title, size=6, color=NAVY, bold=True)
    text_wrap(635,yb+12,220,desc, size=5.5, color=MUTED, leading=7)
    if i<3:
        rect(618,yb+22,1.2,28, fill=GOLD)
rect(610,340,260,30, fill=NAVY, radius=5)
text(610,358,"ГОТОВЫ ОБСУДИТЬ?  →", size=7, color=white, bold=True, align="center", max_width=260)
text(590,378,"Контакты организатора Кактуса — по запросу", size=5, color=MUTED, align="center", max_width=310)

c.showPage()

# =============== SLIDE 9 - FINAL ===============
rect(0,0,W,H, fill=white)
rect(0,0,W,4, fill=GOLD)
rect(0,4,460,536, fill=NAVY)
text(45,32,"G L E B  F I L I P C H E N K O W  ×  N O E M Y  A B R A N T E S", size=6, color=GOLD, bold=True)
text_wrap(45,48,400,"Сделаем 5 декабря легендой Foster", size=22, color=white, bold=True, leading=24)
rect(45,120,50,2, fill=GOLD)
text_wrap(45,130,400,"Ночной клуб Foster ждёт своих 300 гостей. Ваша прибыль — 483 549 ₽ за один вечер.", size=7, color=BEIGE, leading=10)
contacts=[
    ("📅","ДАТА","Суббота, 5 декабря 2026"),
    ("📍","МЕСТО","Foster Nightclub, Крым"),
    ("🎟️","БИЛЕТЫ","3 000 → 3 500 → 4 000 ₽"),
]
for i,(icon,label,val) in enumerate(contacts):
    yb=170 + i*38
    circle(45,yb,19, fill=HexColor("#132B4A"))
    # icon as text
    text(45, yb+4, icon, size=7, color=GOLD, align="center", max_width=19)
    text(75,yb+4,label, size=5, color=GOLD_LIGHT, bold=True)
    text(75,yb+14,val, size=7, color=white, bold=True)
rect(45,300,195,28, fill=GOLD, radius=5)
text(45,317,"ПОДТВЕРДИТЬ УЧАСТИЕ →", size=7, color=NAVY, bold=True, align="center", max_width=195)
text(255,314,"foster.club  •  кактуса  •  mriya resort", size=5, color=MUTED)
text(45,350,"Презентация подготовлена для внутренней отчётности и переговоров с партнёрами", size=5, color=HexColor("#5A6A80"))

text(500,30,"КЛЮЧЕВЫЕ ЦИФРЫ НА ОДНОМ ЭКРАНЕ", size=6, color=NAVY, bold=True)
rect(500,42,330,120, fill=BEIGE, radius=8)
text(520,57,"ВАША ВЫРУЧКА", size=5, color=MUTED, bold=True)
text(710,57,"768 025 ₽", size=8, color=NAVY, bold=True, align="right", max_width=110)
rect(520,68,290,4, fill=GOLD, radius=2)
text(520,82,"ВАШИ РАСХОДЫ", size=5, color=MUTED, bold=True)
text(710,82,"– 284 476 ₽", size=8, color=NAVY, bold=True, align="right", max_width=110)
rect(520,93,105,4, fill=NAVY, radius=2)
rect(520,108,290,36, fill=NAVY, radius=5)
text(532,120,"ЧИСТАЯ ПРИБЫЛЬ", size=5, color=GOLD_LIGHT, bold=True)
text(532,135,"483 549 ₽", size=11, color=white, bold=True)
text(710,118,"~63%\nрентаб.", size=6, color=GOLD, bold=True, align="center", max_width=80)

rect(500,175,330,55, fill=white, stroke=GOLD, radius=8)
text(500,185,"ГАРАНТИРОВАНО ДОГОВОРОМ", size=5, color=GOLD_DARK, bold=True, align="center", max_width=330)
text(500,202,"330 000 ₽", size=12, color=NAVY, bold=True, align="center", max_width=330)
text(500,218,"180К коктейли + 50К столы + 70К VIP", size=5, color=MUTED, align="center", max_width=330)

rect(500,245,330,50, fill=white, stroke=LIGHT_GRAY, radius=8)
text(500,255,"БАЛАНС РИСКОВ", size=5, color=MUTED, bold=True, align="center", max_width=330)
text(510,268,"FOSTER  284К", size=6, color=NAVY, bold=True, align="center", max_width=140)
text(680,268,"КАКТУСА  1 120К", size=6, color=GOLD_DARK, bold=True, align="center", max_width=140)
rect(535,280,70,10, fill=NAVY, radius=3)
rect(695,280,105,10, fill=GOLD, radius=3)
text(500,298,"Ваш риск в 4 раза ниже — их мотивация выше", size=5, color=MUTED, align="center", max_width=330)

rect(500,360,330,28, fill=NAVY, radius=5)
text(500,377,"ГОТОВЫ ФИКСИРОВАТЬ ДАТУ?  •  5 ДЕКАБРЯ 2026", size=6, color=white, bold=True, align="center", max_width=330)

c.save()
print("PDF saved", OUT)

