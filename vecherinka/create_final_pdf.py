#!/usr/bin/env python3
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
W,H=960,540
RED=HexColor("#E60613"); RED_DARK=HexColor("#C91A2E"); DARK=HexColor("#141419"); DARKER=HexColor("#09090B"); BURGUNDY=HexColor("#7A1F2B"); WHITE=HexColor("#FFFFFF"); BEIGE=HexColor("#F5F1E8"); BEIGE2=HexColor("#E8D5B5"); MUTED=HexColor("#B8B5AE"); BEIGE3=HexColor("#E2DED4")
IMG={"cover":"/tmp/foster_images/slide01_img1_Picture_1.jpg","concept_bg":"/tmp/foster_images/slide03_img4_Picture_1.jpg","hall1":"/tmp/foster_images/slide06_img7_Picture_5.jpg","dj_bar":"/tmp/foster_images/slide07_img9_Picture_5.jpg","bar":"/tmp/foster_images/slide08_img11_Picture_5.jpg"}
OUT="/home/user/Foster_Gleb_Noemy_Presentation.pdf"
c=canvas.Canvas(OUT,pagesize=(W,H))
def pt(v): return v*72
def rect(x,y,w,h,fill=None,stroke=None,radius=0):
    yb=H-y-h; c.saveState()
    if fill: c.setFillColor(fill)
    if stroke: c.setStrokeColor(stroke); c.setLineWidth(0.7)
    if radius>0:
        p=c.beginPath(); p.roundRect(x,yb,w,h,radius)
        if fill and stroke: c.drawPath(p,stroke=1,fill=1)
        elif fill: c.drawPath(p,stroke=0,fill=1)
        else: c.drawPath(p,stroke=1,fill=0)
    else:
        if fill and stroke: c.rect(x,yb,w,h,stroke=1,fill=1)
        elif fill: c.rect(x,yb,w,h,stroke=0,fill=1)
        elif stroke: c.rect(x,yb,w,h,stroke=1,fill=0)
    c.restoreState()
def draw_image(path,x,y,w,h,darken=0.5):
    if os.path.exists(path):
        yb=H-y-h
        try: c.drawImage(ImageReader(path),x,yb,width=w,height=h,preserveAspectRatio=False,mask='auto')
        except: c.drawImage(path,x,yb,width=w,height=h,preserveAspectRatio=False)
        c.saveState(); c.setFillColor(HexColor("#120204")); c.setFillAlpha(darken); c.rect(x,yb,w,h,stroke=0,fill=1); c.restoreState()
def text(x,y,txt,size=10,color=WHITE,font="DejaVu",bold=False,align="left",max_width=None):
    c.saveState(); c.setFillColor(color); f="DejaVu-Bold" if bold else font; c.setFont(f,size)
    y_draw=H-y-size*0.35
    if align=="center" and max_width is not None: c.drawCentredString(x+max_width/2,y_draw,txt)
    elif align=="right" and max_width is not None: c.drawRightString(x+max_width,y_draw,txt)
    elif align=="center": c.drawCentredString(x,y_draw,txt)
    elif align=="right": c.drawRightString(x,y_draw,txt)
    else: c.drawString(x,y_draw,txt)
    c.restoreState()
def text_wrap(x,y,w,txt,size=10,color=WHITE,font="DejaVu",bold=False,leading=12,align="left"):
    c.saveState(); c.setFillColor(color); f="DejaVu-Bold" if bold else font; c.setFont(f,size)
    words=txt.split(); lines=[]; cur=""
    for wo in words:
        test=cur+(" " if cur else "")+wo
        if c.stringWidth(test,f,size) <= w: cur=test
        else:
            if cur: lines.append(cur)
            cur=wo
            if c.stringWidth(cur,f,size)>w: lines.append(cur); cur=""
    if cur: lines.append(cur)
    yy=y
    for line in lines:
        y_draw=H-yy-size*0.35
        if align=="center": c.drawCentredString(x+w/2,y_draw,line)
        elif align=="right": c.drawRightString(x+w,y_draw,line)
        else: c.drawString(x,y_draw,line)
        yy+=leading
    c.restoreState(); return yy

# SLIDE 1
draw_image(IMG["cover"],0,0,W,H,darken=0.55)
rect(pt(0.60),pt(0.35),pt(2.20),pt(0.45),fill=WHITE,radius=5)
text(pt(0.60),pt(0.38+0.10),"FOSTER • ЯЛТА",size=9,color=DARK,bold=True,align="center",max_width=pt(2.20))
text(pt(8.0),pt(0.35+0.12),"5 ДЕКАБРЯ 2026 • 22:00–06:00",size=8,color=WHITE,bold=True,align="right",max_width=pt(4.70))
text(pt(8.0),pt(0.65+0.10),"300–350 ГОСТЕЙ • Total Black + Red",size=7,color=RED,align="right",max_width=pt(4.70))
text_wrap(pt(0.60),pt(1.30+0.20),pt(12),"ЧЁРНЫМ ПО КРАСНОМУ",size=32,color=WHITE,bold=True,leading=32)
text_wrap(pt(0.60),pt(2.05+0.20),pt(12),"GLEB FILIPCHENKOW & NOEMY ABRANTES",size=11,color=RED,bold=True,leading=13)
rect(pt(0.62),pt(2.55),pt(1.40),pt(0.08),fill=RED)
text_wrap(pt(0.60),pt(2.75),pt(12),"Первый привоз в Крым • Хит «Mallorca» • Total Black дресс-код с красной деталью",size=9,color=BEIGE,bold=True,leading=11)
rect(pt(0.60),pt(3.35),pt(3.90),pt(1.55),fill=DARK,radius=6)
text(pt(0.60),pt(3.50+0.10),"ДРЕСС-КОД",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(0.80),pt(3.85),pt(3.50),"Total black + красная деталь — помада, галстук, сумка, обувь.",size=8,color=BEIGE3,leading=10,align="center")
rect(pt(4.75),pt(3.35),pt(3.90),pt(1.55),fill=DARK,radius=6)
text(pt(4.75),pt(3.50+0.10),"ДЕКОР",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(4.95),pt(3.85),pt(3.50),"Красный свет + бархат + фото-зона. Минимум декора — максимум эффекта.",size=8,color=BEIGE3,leading=10,align="center")
rect(pt(8.90),pt(3.35),pt(3.90),pt(1.55),fill=RED,radius=6)
text(pt(8.90),pt(3.50+0.10),"ХЕДЛАЙНЕР 01:00",size=10,color=WHITE,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(9.10),pt(3.85),pt(3.50),"Блэкаут → вспышка красного. Кульминация ночи для контента.",size=8,color=WHITE,leading=10,align="center")
text(pt(0.60),pt(5.15+0.10),"22:00 СБОР • 23:30 ПЕРФОРМАНС • 01:00 ХЕДЛАЙНЕРЫ • 03:00 AFTERPARTY",size=7,color=MUTED,align="center",max_width=pt(12.1))
rect(pt(0.60),pt(5.45),pt(12.10),pt(0.90),fill=DARK,radius=6)
text(pt(0.60),pt(5.60+0.10),"Глеб — Mallorca (2026 хит, далее Серябкина) • Noemy — Garota/HOJE • Первый привоз, ожидание с лета",size=7.5,color=BEIGE,align="center",max_width=pt(12.10))
text(pt(0.60),pt(6.95+0.10),"Только Foster Club",size=7,color=MUTED)
text(pt(11.0),pt(6.95+0.10),"01 / 05",size=8,color=WHITE,align="right",max_width=pt(1.70))
c.showPage()

# SLIDE 2 — ХЕДЛАЙНЕР + КОНЦЕПЦИЯ
draw_image(IMG["concept_bg"],0,0,W,H,darken=0.55)
text(pt(0.60),pt(0.35+0.12),"Х Е Д Л А Й Н Е Р • КОНЦЕПЦИЯ",size=9,color=RED,bold=True)
text_wrap(pt(0.60),pt(0.75+0.15),pt(12),"Gleb & Noemy — имя продаёт",size=20,color=WHITE,bold=True,leading=21)
rect(pt(0.62),pt(1.30),pt(1.0),pt(0.08),fill=RED)
rect(pt(0.60),pt(1.60),pt(6.00),pt(1.40),fill=DARK,radius=6)
text(pt(0.60),pt(1.75+0.10),"GLEB — MALLORCA",size=10,color=RED,bold=True,align="center",max_width=pt(6.00))
text_wrap(pt(0.80),pt(2.05),pt(5.60),"Продюсер хита без бюджета — сотни тысяч просмотров. Взрыв в 2026, далее трек с Серябкиной.",size=8.5,color=BEIGE3,leading=10,align="center")
rect(pt(0.60),pt(3.15),pt(6.00),pt(1.40),fill=DARK,radius=6)
text(pt(0.60),pt(3.30+0.10),"NOEMY — GAROTA / HOJE",size=10,color=RED,bold=True,align="center",max_width=pt(6.00))
text_wrap(pt(0.80),pt(3.60),pt(5.60),"Европейский вокал, соавтор «Garota» и «HOJE». Идеальна для красной ночи Foster.",size=8.5,color=BEIGE3,leading=10,align="center")
rect(pt(0.60),pt(4.70),pt(6.00),pt(1.10),fill=RED,radius=6)
text(pt(0.60),pt(4.80+0.10),"Почему SOLD OUT — отложенный спрос с лета, 5 декабря единственное окно",size=7.5,color=WHITE,bold=True,align="center",max_width=pt(6.00))
rect(pt(7.00),pt(1.60),pt(5.70),pt(1.15),fill=DARK,radius=6)
text(pt(7.00),pt(1.75+0.10),"300–350",size=14,color=RED,bold=True,align="center",max_width=pt(5.70))
text(pt(7.00),pt(2.10+0.08),"гостей — полный зал",size=8,color=BEIGE,align="center",max_width=pt(5.70))
rect(pt(7.00),pt(2.90),pt(5.70),pt(1.15),fill=DARK,radius=6)
text(pt(7.00),pt(3.05+0.10),"01:00",size=14,color=RED,bold=True,align="center",max_width=pt(5.70))
text(pt(7.00),pt(3.40+0.08),"блэкаут + красный — главный момент",size=8,color=BEIGE,align="center",max_width=pt(5.70))
rect(pt(7.00),pt(4.20),pt(5.70),pt(1.60),fill=DARK,radius=6)
text(pt(7.00),pt(4.35+0.10),"ТАЙМИНГ",size=9,color=RED,bold=True,align="center",max_width=pt(5.70))
text_wrap(pt(7.20),pt(4.65),pt(5.30),"22:00 сбор • 23:30 перформанс • 01:00 хедлайнеры • 03:00 afterparty • 06:00 закрытие",size=7,color=BEIGE3,leading=9,align="center")
text(pt(0.60),pt(6.95+0.10),"Фото: красный лаунж Foster Club",size=7,color=MUTED)
text(pt(11.0),pt(6.95+0.10),"02 / 05",size=8,color=WHITE,align="right",max_width=pt(1.70))
c.showPage()

# SLIDE 3 — КАКТУСА ФИНАНСЫ
draw_image(IMG["dj_bar"],0,0,W,H,darken=0.55)
text(pt(0.60),pt(0.35+0.12),"Ф И Н А Н С Ы   К А К Т У С Ы",size=9,color=RED,bold=True)
text_wrap(pt(0.60),pt(0.75+0.15),pt(12),"Инвестируют 1.12 млн — риск на них",size=18,color=WHITE,bold=True,leading=19)
rect(pt(0.62),pt(1.25),pt(1.0),pt(0.08),fill=RED)
rect(pt(0.60),pt(1.55),pt(3.90),pt(1.55),fill=DARK,radius=6)
text(pt(0.60),pt(1.70+0.10),"ДОХОД",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(0.80),pt(2.00),pt(3.50),"Билеты 300×3400 = 1 020 000 ₽ (база) • SOLD OUT 350×3400 = 1 190 000 ₽ • Цены 3000→3500→4000",size=7.5,color=BEIGE3,leading=9,align="center")
rect(pt(4.75),pt(1.55),pt(3.90),pt(1.55),fill=DARK,radius=6)
text(pt(4.75),pt(1.70+0.10),"ЗАТРАТЫ",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(4.95),pt(2.00),pt(3.50),"Гонорар 550К + коктейли 180К + шаттл 20К + логистика 80К + промо 100К + артисты 75К + тех 85К + охрана 30К = 1 120 040 ₽",size=7,color=BEIGE3,leading=8,align="center")
rect(pt(8.90),pt(1.55),pt(3.90),pt(1.55),fill=RED,radius=6)
text(pt(8.90),pt(1.70+0.10),"РЕЗУЛЬТАТ",size=10,color=WHITE,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(9.10),pt(2.00),pt(3.50),"База 300 чел: –100 040 ₽ убыток • Пессим 250: –295К • SOLD OUT 350: +69 960 ₽ прибыль — в плюс только так!",size=7.5,color=WHITE,leading=9,align="center")
rect(pt(0.60),pt(3.40),pt(12.10),pt(1.15),fill=WHITE,radius=6)
text(pt(0.60),pt(3.55+0.10),"ВЫВОД: Мотивация продавать до последнего билета — ваши 200 040 ₽ (180К коктейли + 20К шаттл) зафиксированы фиксом, независимо от явки.",size=8.5,color=DARK,bold=True,align="center",max_width=pt(12.10))
text_wrap(pt(0.80),pt(3.95),pt(11.7),"Гонорар 550К (фикс) • Логистика 80К • Промо 100К • Местные 75К • Тех 85К • Охрана 30К",size=7,color=MUTED,align="center")
rect(pt(0.60),pt(4.80),pt(3.90),pt(1.15),fill=DARK,radius=6)
text(pt(0.60),pt(4.95+0.10),"250 ЧЕЛ",size=11,color=RED,bold=True,align="center",max_width=pt(3.90))
text(pt(0.60),pt(5.30+0.08),"–295 040 ₽ • 825К выручка",size=7,color=BEIGE,align="center",max_width=pt(3.90))
rect(pt(4.75),pt(4.80),pt(3.90),pt(1.15),fill=DARK,radius=6)
text(pt(4.75),pt(4.95+0.10),"300 ЧЕЛ",size=11,color=RED,bold=True,align="center",max_width=pt(3.90))
text(pt(4.75),pt(5.30+0.08),"–100 040 ₽ • 1 020К выручка",size=7,color=BEIGE,align="center",max_width=pt(3.90))
rect(pt(8.90),pt(4.80),pt(3.90),pt(1.15),fill=RED,radius=6)
text(pt(8.90),pt(4.95+0.10),"350 ЧЕЛ",size=11,color=WHITE,bold=True,align="center",max_width=pt(3.90))
text(pt(8.90),pt(5.30+0.08),"+69 960 ₽ • 1 190К выручка",size=7,color=WHITE,align="center",max_width=pt(3.90))
text(pt(0.60),pt(6.95+0.10),"Фото: DJ-бар Foster Club",size=7,color=MUTED)
text(pt(11.0),pt(6.95+0.10),"03 / 05",size=8,color=WHITE,align="right",max_width=pt(1.70))
c.showPage()

# SLIDE 4 — FOSTER ФИНАНСЫ
draw_image(IMG["bar"],0,0,W,H,darken=0.55)
text(pt(0.60),pt(0.35+0.12),"В А Ш И   Ф И Н А Н С Ы  •  FOSTER",size=9,color=RED,bold=True)
text_wrap(pt(0.60),pt(0.75+0.15),pt(12),"Ваша прибыль — 483 549 ₽",size=20,color=WHITE,bold=True,leading=21)
rect(pt(0.62),pt(1.25),pt(1.0),pt(0.08),fill=RED)
rect(pt(0.60),pt(1.55),pt(3.90),pt(1.85),fill=DARK,radius=6)
text(pt(0.60),pt(1.70+0.10),"ЗАТРАТЫ 284 476 ₽",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(0.80),pt(2.05),pt(3.50),"Персонал 165 470 ₽\nБар 92 006 ₽ (25% от 368К)\nКухня 27 000 ₽ (27% от 100К)\n0 ₽ — артисты/аренда",size=8.5,color=BEIGE3,leading=10,align="center")
rect(pt(4.75),pt(1.55),pt(3.90),pt(1.85),fill=DARK,radius=6)
text(pt(4.75),pt(1.70+0.10),"ДОХОДЫ 768 025 ₽",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(4.95),pt(2.05),pt(3.50),"Бар 368К + Кухня 100К\nКоктейли 180К (300×600)\nСтолы 50К + VIP 70К\n+ кальяны Макао",size=8.5,color=BEIGE3,leading=10,align="center")
rect(pt(8.90),pt(1.55),pt(3.90),pt(1.85),fill=RED,radius=6)
text(pt(8.90),pt(1.70+0.10),"ПРИБЫЛЬ 483 549 ₽",size=10,color=WHITE,bold=True,align="center",max_width=pt(3.90))
text_wrap(pt(9.10),pt(2.05),pt(3.50),"768 025 – 284 476 =\n+483 549 ₽ чистыми\n~63% рентабельность\n+ кальяны",size=9,color=WHITE,leading=11,align="center")
rect(pt(0.60),pt(3.70),pt(12.10),pt(1.15),fill=RED,radius=6)
text(pt(0.60),pt(3.85+0.10),"ГАРАНТИРОВАНО ДОГОВОРОМ — 330 000 ₽",size=11,color=WHITE,bold=True,align="center",max_width=pt(12.10))
text_wrap(pt(0.80),pt(4.20),pt(11.7),"180 000 ₽ коктейли (300×600, фикс) + 50 000 ₽ депозиты столов + 70 000 ₽ VIP №1/№2 (40К+30К) — фикс, независимо от явки. Даже при 200 гостях вы в плюсе.",size=7.5,color=WHITE,leading=9,align="center")
rect(pt(0.60),pt(5.10),pt(3.90),pt(1.15),fill=DARK,radius=6)
text(pt(0.60),pt(5.25+0.10),"БАР 368К",size=11,color=RED,bold=True,align="center",max_width=pt(3.90))
text(pt(0.60),pt(5.60+0.08),"валовая выручка",size=7,color=BEIGE,align="center",max_width=pt(3.90))
rect(pt(4.75),pt(5.10),pt(3.90),pt(1.15),fill=DARK,radius=6)
text(pt(4.75),pt(5.25+0.10),"КУХНЯ 100К",size=11,color=RED,bold=True,align="center",max_width=pt(3.90))
text(pt(4.75),pt(5.60+0.08),"валовая выручка",size=7,color=BEIGE,align="center",max_width=pt(3.90))
rect(pt(8.90),pt(5.10),pt(3.90),pt(1.15),fill=DARK,radius=6)
text(pt(8.90),pt(5.25+0.10),"0 ₽ ДОП-ЗАТРАТ",size=10,color=RED,bold=True,align="center",max_width=pt(3.90))
text(pt(8.90),pt(5.60+0.08),"ни аренды, ни артистов",size=7,color=BEIGE,align="center",max_width=pt(3.90))
text(pt(0.60),pt(6.95+0.10),"Фото: бар Foster Club",size=7,color=MUTED)
text(pt(11.0),pt(6.95+0.10),"04 / 05",size=8,color=WHITE,align="right",max_width=pt(1.70))
c.showPage()

# SLIDE 5 — ИТОГ
draw_image(IMG["cover"],0,0,W,H,darken=0.60)
text(pt(0.60),pt(0.35+0.12),"Г Л А В Н Ы Й   И Т О Г",size=9,color=RED,bold=True)
text_wrap(pt(0.60),pt(0.75+0.15),pt(12),"Безопасно. Прибыльно.",size=22,color=WHITE,bold=True,leading=23)
text_wrap(pt(0.60),pt(1.25+0.15),pt(12),"Полмиллиона за ночь.",size=22,color=RED,bold=True,leading=23)
rect(pt(0.62),pt(1.75),pt(1.0),pt(0.08),fill=RED)
rect(pt(0.60),pt(2.00),pt(6.20),pt(1.90),fill=DARK,radius=6)
text(pt(0.80),pt(2.15+0.10),"ВАША ЧИСТАЯ ПРИБЫЛЬ",size=9,color=RED,bold=True)
text(pt(0.80),pt(2.50+0.12),"483 549 ₽",size=24,color=WHITE,bold=True)
text(pt(4.80),pt(2.55+0.10),"63%\nрентаб.",size=11,color=RED,bold=True,align="center",max_width=pt(1.60))
text(pt(0.80),pt(3.15+0.10),"768 025 ₽ доходы – 284 476 ₽ затраты + кальяны",size=8,color=BEIGE2)
rect(pt(7.10),pt(2.00),pt(5.40),pt(1.90),fill=WHITE,radius=6)
text(pt(7.10),pt(2.15+0.10),"ЗАФИКСИРОВАТЬ В ДОГОВОРЕ",size=9,color=RED,bold=True,align="center",max_width=pt(5.40))
text_wrap(pt(7.30),pt(2.50),pt(5.00),"✓ 180 000 ₽ — коктейли (300×600, фикс)\n✓ 20 040 ₽ — шаттл (фикс)\n✓ Оплата 1–3 дня после события",size=9,color=DARK,leading=11)
rect(pt(0.60),pt(4.20),pt(12.10),pt(1.45),fill=WHITE,radius=6)
text(pt(0.60),pt(4.35+0.12),"Вы не вкладываетесь в артистов, аренду и оборудование. Просто принимаете гостей и считаете прибыль.",size=9,color=DARK,bold=True,align="center",max_width=pt(12.10))
text_wrap(pt(0.80),pt(4.75),pt(11.7),"Организатор мотивирован на SOLD OUT 350 чел. (только там его +70К). Ваши 330 000 ₽ гарантированы. Следующий шаг → согласовать договор → старт продаж 01.11 → 5 декабря.",size=8,color=MUTED,leading=10,align="center")
text(pt(0.60),pt(6.95+0.10),"Только Foster Club • Красно-чёрный • Крупный шрифт для чтения",size=7,color=MUTED)
text(pt(11.0),pt(6.95+0.10),"05 / 05",size=8,color=WHITE,align="right",max_width=pt(1.70))
c.showPage()
c.save()
print(f"Saved PDF {OUT}")

