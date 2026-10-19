#!/usr/bin/env python3
# PDF для чистой версии (8 слайдов, без 4,5,12,13,14)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
W,H=960,540
GOLD=HexColor("#C9A25E"); GOLD_LIGHT=HexColor("#E5C888"); BEIGE=HexColor("#F5F1E8"); BEIGE2=HexColor("#DDD8CC"); BEIGE3=HexColor("#E2DED4"); MUTED=HexColor("#7E7A72"); DARK=HexColor("#141419"); DARKER=HexColor("#09090B"); BURGUNDY=HexColor("#7A1F2B"); WHITE=HexColor("#FFFFFF")
IMG={"cover":"/tmp/foster_images/slide01_img1_Picture_1.jpg","about":"/tmp/foster_images/slide02_img3_Picture_5.jpg","concept_bg":"/tmp/foster_images/slide03_img4_Picture_1.jpg","alt_bar":"/tmp/foster_images/slide08_img11_Picture_5.jpg","alt_loft":"/tmp/foster_images/slide09_img12_Picture_5.jpg","hall1":"/tmp/foster_images/slide06_img7_Picture_5.jpg","hall2":"/tmp/foster_images/slide06_img8_Picture_6.jpg","dj_bar":"/tmp/foster_images/slide07_img9_Picture_5.jpg","vip":"/tmp/foster_images/slide07_img10_Picture_6.jpg","timeline":"/tmp/foster_images/slide11_img18_Picture_5.jpg","multi1":"/tmp/foster_images/slide10_img14_Picture_5.jpg","multi2":"/tmp/foster_images/slide10_img15_Picture_6.jpg"}
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
def draw_image(path,x,y,w,h,darken=0):
    if os.path.exists(path):
        yb=H-y-h
        try: c.drawImage(ImageReader(path),x,yb,width=w,height=h,preserveAspectRatio=False,mask='auto')
        except: c.drawImage(path,x,yb,width=w,height=h,preserveAspectRatio=False)
        if darken:
            c.saveState(); c.setFillColor(HexColor("#0A0A0F")); c.setFillAlpha(darken); c.rect(x,yb,w,h,stroke=0,fill=1); c.restoreState()
def text(x,y,txt,size=10,color=BEIGE,font="DejaVu",bold=False,align="left",max_width=None):
    c.saveState(); c.setFillColor(color); f="DejaVu-Bold" if bold else font; c.setFont(f,size)
    y_draw=H-y-size*0.35
    if align=="center" and max_width is not None: c.drawCentredString(x+max_width/2,y_draw,txt)
    elif align=="right" and max_width is not None: c.drawRightString(x+max_width,y_draw,txt)
    elif align=="center": c.drawCentredString(x,y_draw,txt)
    elif align=="right": c.drawRightString(x,y_draw,txt)
    else: c.drawString(x,y_draw,txt)
    c.restoreState()
def text_wrap(x,y,w,txt,size=8,color=BEIGE,font="DejaVu",bold=False,leading=10,align="left"):
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
def card(x_in,y_in,w_in,h_in,fill=DARK,radius=6,header=None,body=None,header_color=GOLD_LIGHT,body_color=BEIGE3,header_size=6,body_size=6):
    x,y,w,h=pt(x_in),pt(y_in),pt(w_in),pt(h_in)
    rect(x,y,w,h,fill=fill,radius=radius)
    if header: text(x+pt(0.22),y+pt(0.12),header,size=header_size,color=header_color,bold=True)
    if body: text_wrap(x+pt(0.22),y+pt(0.52),w-pt(0.44),body,size=body_size,color=body_color,leading=body_size+2.5)
def add_header(small,title,num="01 / 08"):
    text(pt(0.55),pt(0.32+0.10),small,size=7.5,color=GOLD,bold=True)
    text_wrap(pt(0.55),pt(0.60+0.15),pt(7.5),title,size=16,color=BEIGE,bold=True,leading=17)
    rect(pt(0.57),pt(1.42),pt(0.70),pt(0.05),fill=GOLD)
    text(pt(11.70),pt(0.32+0.10),num,size=7,color=MUTED,align="right",max_width=pt(1.10))
    text(pt(0.55),pt(7.12+0.10),"Фото: галерея Foster Club — Яндекс Карты",size=5.5,color=MUTED)
    text(pt(11.30),pt(7.12+0.10),f"FOSTER • {num.split('/')[0].strip()}",size=5.5,color=MUTED,bold=True,align="right",max_width=pt(1.50))

# SLIDE 1 COVER
draw_image(IMG["cover"],0,0,W,H,darken=0.45)
rect(pt(0.55),pt(0.45),pt(2.45),pt(0.98),fill=BEIGE,radius=6)
text(pt(0.70),pt(0.70),"FOSTER",size=11,color=DARK,bold=True,align="center",max_width=pt(2.15))
text(pt(0.70),pt(0.88),"NIGHT & KARAOKE CLUB",size=4.5,color=GOLD,bold=True,align="center",max_width=pt(2.15))
text(pt(9.2),pt(0.45+0.08),"Ночной клуб Foster  •  Крым",size=5,color=WHITE,align="right",max_width=pt(3.6))
text(pt(9.2),pt(0.68+0.08),"5 декабря 2026  •  300–350 гостей  •  SOLD OUT цель",size=5,color=GOLD,bold=True,align="right",max_width=pt(3.6))
text_wrap(pt(0.55),pt(1.85+0.25),pt(8.5),"GLEB FILIPCHENKOW",size=24,color=WHITE,bold=True,leading=24)
text_wrap(pt(0.55),pt(2.35+0.25),pt(8.5),"& NOEMY ABRANTES",size=24,color=GOLD,bold=True,leading=24)
rect(pt(0.57),pt(3.55),pt(0.85),pt(0.06),fill=GOLD)
text_wrap(pt(0.55),pt(3.75+0.08),pt(8.2),"Первый привоз в Крым  •  «Mallorca» — хит лета 2025  •  Garota • HOJE",size=6.5,color=BEIGE2,leading=9)
rect(pt(0.55),pt(4.40),pt(6.0),pt(1.55),fill=DARK,radius=6)
text(pt(0.75),pt(4.55+0.08),"◆  Ночь, которую не нужно искать за пределами клуба",size=6,color=GOLD,bold=True)
text_wrap(pt(0.75),pt(4.85+0.08),pt(5.6),"Событие продаёт имя: вирусный трек «Mallorca» без бюджета, сотни тысяч просмотров. Привоз ждали с лета — 5 декабря единственное окно.",size=6,color=BEIGE3,leading=8)
for i,(num,lbl) in enumerate([("768 025 ₽","Выручка + кальяны"),("284 476 ₽","Ваши затраты"),("483 549 ₽","Чистая прибыль 63%")]):
    xl=pt(0.55+i*2.05)
    rect(xl,pt(6.20),pt(1.90),pt(0.75),fill=DARK,radius=5)
    text(xl,pt(6.28+0.08),num,size=8,color=WHITE,bold=True,align="center",max_width=pt(1.90))
    text(xl,pt(6.55+0.06),lbl,size=4.5,color=GOLD_LIGHT,bold=True,align="center",max_width=pt(1.90))
text(pt(0.55),pt(7.12+0.10),"Пересборка на подлинных фото Foster Club из галереи Яндекс Карт",size=5,color=MUTED)
text(pt(11.30),pt(7.12+0.10),"01 / 08",size=6,color=WHITE,align="right",max_width=pt(1.50))
c.showPage()

# SLIDE 2 ПОЗИЦИОНИРОВАНИЕ
rect(0,0,W,H,fill=WHITE)
add_header("П О З И Ц И О Н И Р О В А Н И Е","Камерный премиум-клуб","02 / 08")
draw_image(IMG["about"],pt(7.00),pt(1.75),pt(5.78),pt(3.50))
y=pt(1.95)
for label,body in [("Формат:","ночной и караоке-клуб — вечерняя «гостиная» с танцполом, сценой, VIP-залами."),("Гости:","300–350 человек — полный SOLD OUT камерного зала (до 60 мест за столами + танцпол)."),("Почему сейчас:","летний привоз сорвался из-за топливного кризиса; 5 декабря — единственное окно артистов."),("Итог для вас:","вы предоставляете только площадку и бар — без вложений в артистов и без проживания.")]:
    text(pt(0.55),y,"◆",size=7,color=GOLD,bold=True)
    text_wrap(pt(0.75),y,pt(6.15),f"{label} {body}",size=7,color=HexColor("#2A2A30"),leading=9)
    y+=pt(0.42)
card(0.55,5.45,3.95,1.35,fill=DARK,header="ПЛОЩАДКА",body="FOSTER — клуб, бар,\nкухня, персонал,\nоборудование и VIP-залы")
card(4.70,5.45,3.95,1.35,fill=DARK,header="ОРГАНИЗАТОР",body="КАКТУСА — гонорары,\nлогистика, техрайдер,\nпродвижение, продажи")
card(8.85,5.45,3.93,1.35,fill=DARK,header="ДАТА И БИЛЕТЫ",body="Суббота, 5 декабря 2026\n3 000 ₽ → 3 500 ₽ → 4 000 ₽\nЦель: 350 SOLD OUT")
c.showPage()

# SLIDE 3 КОНЦЕПЦИЯ
draw_image(IMG["concept_bg"],0,0,W,H,darken=0.5)
text(pt(0.55),pt(0.32+0.10),"К О Н Ц Е П Ц И Я   С Д Е Л К И",size=7.5,color=GOLD,bold=True)
text_wrap(pt(0.55),pt(0.60+0.15),pt(10),"«Чистые финансы — без скрытых затрат»",size=16,color=BEIGE,bold=True,leading=17)
rect(pt(0.57),pt(1.42),pt(0.70),pt(0.05),fill=GOLD)
text(pt(11.70),pt(0.32+0.10),"03 / 08",size=7,color=MUTED,align="right",max_width=pt(1.10))
card(0.55,1.85,3.95,1.35,fill=DARKER,header="1 • ДОХОДЫ",body="768 025 ₽ + кальяны\nБар 368К + Кухня 100К +\n180К коктейли + 120К столы/VIP")
card(4.70,1.85,3.95,1.35,fill=DARKER,header="2 • ЗАТРАТЫ",body="284 476 ₽\nПерсонал 165К + Бар 92К +\nКухня 27К — всё покрывается")
card(8.85,1.85,3.93,1.35,fill=DARKER,header="3 • ПРИБЫЛЬ",body="483 549 ₽ чистыми\nРентабельность ~63%\nБез проживания — ноль рисков")
text(pt(0.55),pt(7.12+0.10),"Проживание артистов полностью исключено — никаких скидок и обязательств. Фото: красный лаунж Foster Club.",size=5,color=MUTED)
text(pt(11.30),pt(7.12+0.10),"FOSTER • 03",size=5,color=MUTED,bold=True,align="right",max_width=pt(1.50))
c.showPage()

# SLIDE 4 ДОХОДЫ
rect(0,0,W,H,fill=WHITE)
add_header("Н А Ш И   Д О Х О Д Ы","768 025 ₽ + кальяны","04 / 08")
draw_image(IMG["alt_bar"],pt(0.55),pt(1.75),pt(3.90),pt(3.50))
text(pt(0.55),pt(5.35+0.08),"Бар Foster — фото Яндекс Карты",size=5,color=MUTED,align="center",max_width=pt(3.90))
draw_image(IMG["alt_loft"],pt(4.65),pt(1.75),pt(3.90),pt(3.50))
text(pt(4.65),pt(5.35+0.08),"Лаунж-зона — фото Foster Club",size=5,color=MUTED,align="center",max_width=pt(3.90))
y=pt(1.95)
for name,price,note in [("Бар","368 025 ₽","валовая выручка"),("Кухня","100 000 ₽","валовая выручка"),("◆ Коктейли","180 000 ₽","300 × 600 ₽ → вам"),("◆ Столы","50 000 ₽","4×5 000 + 3×10 000"),("◆ VIP №1","40 000 ₽","депозит → вам"),("◆ VIP №2","30 000 ₽","депозит → вам")]:
    is_hl="◆" in name
    col=GOLD if is_hl else DARK
    text(pt(8.75),y,name,size=6.5,color=col,bold=True)
    text(pt(10.30),y,price,size=7,color=DARK,bold=True)
    text(pt(11.50),y,"— "+note,size=5,color=MUTED)
    y+=pt(0.30)
text(pt(8.75),y+pt(0.10),"Итого: 768 025 ₽ + кальяны от Макао",size=6,color=BURGUNDY,bold=True)
rect(pt(0.55),pt(5.80),pt(12.23),pt(1.00),fill=BURGUNDY,radius=6)
text(pt(0.75),pt(5.95+0.08),"ГАРАНТИРОВАНО ДОГОВОРОМ",size=5,color=GOLD_LIGHT,bold=True,align="center",max_width=pt(12.23))
text(pt(0.75),pt(6.25+0.08),"330 000 ₽ (180К коктейли + 50К столы + 70К VIP) — фикс, независимо от явки. Кальяны от Макао — по факту продаж.",size=6,color=WHITE,align="center",max_width=pt(12.23))
c.showPage()

# SLIDE 5 ЗАТРАТЫ
rect(0,0,W,H,fill=WHITE)
add_header("Н А Ш И   З А Т Р А Т Ы","Всего 284 476 ₽","05 / 08")
draw_image(IMG["multi1"],pt(0.55),pt(1.75),pt(2.70),pt(3.50))
draw_image(IMG["multi2"],pt(3.40),pt(1.75),pt(3.80),pt(3.50))
y=pt(1.95)
for name,price,note in [("Персонал клуба","165 470 ₽","официанты, бармены, повара"),("Себестоимость бара","92 006 ₽","25% от 368 025 ₽"),("Себестоимость кухни","27 000 ₽","27% от 100 000 ₽")]:
    text(pt(7.50),y,name,size=7,color=DARK,bold=True)
    text(pt(7.50),y+pt(0.18),price,size=8,color=DARK,bold=True,align="right",max_width=pt(5.2))
    text(pt(7.50),y+pt(0.32),note,size=5,color=MUTED)
    y+=pt(0.45)
text(pt(7.50),y+pt(0.15),"Итого затраты: 284 476 ₽",size=7,color=BURGUNDY,bold=True)
card(0.55,5.60,3.95,1.20,fill=DARK,header="ПЕРСОНАЛ 165К",body="Самая крупная статья —\nно покрывается баром\nи кухней.")
card(4.70,5.60,3.95,1.20,fill=DARK,header="ФУДКОСТ",body="Бар 25% + Кухня 27% —\nнормальные ресторанные\nпоказатели.")
card(8.85,5.60,3.93,1.20,fill=DARK,header="0 ₽ ПРОЖИВАНИЕ",body="Полностью исключено:\nни скидок, ни затрат,\nни обязательств.")
c.showPage()

# SLIDE 6 ФИН РЕЗУЛЬТАТ
draw_image(IMG["cover"],0,0,W,H,darken=0.45)
text(pt(0.55),pt(0.32+0.10),"Ф И Н А Н С О В Ы Й   Р Е З У Л Ь Т А Т",size=7.5,color=GOLD,bold=True)
text_wrap(pt(0.55),pt(0.60+0.15),pt(8),"Чистая прибыль — 483 549 ₽",size=18,color=BEIGE,bold=True,leading=19)
rect(pt(0.57),pt(1.42),pt(0.70),pt(0.05),fill=GOLD)
text(pt(11.70),pt(0.32+0.10),"06 / 08",size=7,color=MUTED,align="right",max_width=pt(1.10))
card(0.55,1.85,3.95,1.60,fill=DARKER,header="ДОХОДЫ",body="768 025 ₽\nБар 368 025 + Кухня 100 000 +\nКоктейли 180К + Столы 50К +\nVIP 70К + кальяны")
card(4.70,1.85,3.95,1.60,fill=DARKER,header="ЗАТРАТЫ",body="284 476 ₽\nПерсонал 165 470 +\nБар 92 006 + Кухня 27 000\nБез проживания — 0 ₽")
card(8.85,1.85,3.93,1.60,fill=BURGUNDY,header="ПРИБЫЛЬ 63%",body="483 549 ₽\n768 025 – 284 476\n+ кальяны от Макао\nРентабельность ~63%")
rect(pt(0.55),pt(5.80),pt(12.23),pt(0.95),fill=BEIGE,radius=6)
text(pt(0.75),pt(5.95+0.08),"◆  Что зафиксировать в договоре",size=6,color=BURGUNDY,bold=True)
text_wrap(pt(0.75),pt(6.20),pt(11.8),"180 000 ₽ за коктейли (фикс)  •  20 040 ₽ за шаттл (охрана/клининг/хостес)  •  330 000 ₽ гарант до старта  •  VIP и столы — только ваша зона  •  Оплата 1–3 рабочих дня",size=6,color=DARK,leading=8)
text(pt(0.55),pt(7.12+0.10),"Проживание в Mriya Resort исключено — никаких обязательств с вашей стороны.",size=5,color=MUTED)
c.showPage()

# SLIDE 7 ХЕДЛАЙНЕРЫ
rect(0,0,W,H,fill=WHITE)
add_header("Х Е Д Л А Й Н Е Р Ы","Имена, которые продают","07 / 08")
draw_image(IMG["hall1"],pt(0.55),pt(1.75),pt(8.10),pt(3.60))
draw_image(IMG["hall2"],pt(8.85),pt(1.75),pt(3.93),pt(3.60))
text(pt(0.55),pt(5.45+0.08),"Сцена Foster Club — большой LED-экран и свет Helen Design",size=5,color=MUTED,align="center",max_width=pt(8.10))
text(pt(8.85),pt(5.45+0.08),"Зеркальный бар — Яндекс Карты",size=5,color=MUTED,align="center",max_width=pt(3.93))
card(0.55,5.85,3.95,1.20,fill=DARK,header="GLEB FILIPCHENKOW",body="Хит «Mallorca» (лето 2025),\nвзрыв весной 2026 без\nбюджета. Релиз с Серябкиной.")
card(4.70,5.85,3.95,1.20,fill=DARK,header="NOEMY ABRANTES",body="«Garota» и «HOJE».\nЕвропейский вокал для\nночной атмосферы Foster.")
card(8.85,5.85,3.93,1.20,fill=BURGUNDY,header="ПОЧЕМУ ПРОДАСТСЯ",body="Имя продаёт событие.\nОтложенный спрос с лета.\n5 декабря — единственное окно.")
c.showPage()

# SLIDE 8 ИТОГ
rect(0,0,W,H,fill=WHITE)
draw_image(IMG["alt_bar"],pt(0.55),pt(1.75),pt(4.20),pt(4.00))
text(pt(0.55),pt(5.85+0.08),"Лаунж Foster Club — фото Яндекс Карты",size=5,color=MUTED,align="center",max_width=pt(4.20))
add_header("И Т О Г","Ваша сделка на 483 500 ₽","08 / 08")
text_wrap(pt(5.15),pt(1.85+0.08),pt(7.50),"Минимальный риск — максимальная отдача.",size=10,color=DARK,bold=True,leading=11)
text_wrap(pt(5.15),pt(2.30),pt(7.50),"Вы не платите за артистов и проживание — только операционка, которая окупается баром.",size=6,color=HexColor("#4A4A52"),leading=8)
card(5.15,3.10,3.70,1.35,fill=DARK,header="ЧТО ФИКСИРОВАТЬ",body="180 000 ₽ коктейли (фикс)\n50 000 ₽ столы + 70 000 ₽ VIP\n20 040 ₽ шаттл")
card(9.05,3.10,3.70,1.35,fill=BURGUNDY,header="СРОК ОПЛАТЫ",body="1–3 рабочих дня\nпосле мероприятия\nЖивыми деньгами")
rect(pt(5.15),pt(4.70),pt(7.50),pt(0.90),fill=BEIGE,radius=6)
text(pt(5.35),pt(4.85+0.08),"Следующий шаг",size=5,color=BURGUNDY,bold=True)
text_wrap(pt(5.35),pt(5.10),pt(7.10),"Согласовать договор → запустить продажи с 01.11 → 5 декабря SOLD OUT 350 гостей.",size=6,color=DARK,leading=8)
c.showPage()
c.save()
print(f"Saved PDF {OUT}")

