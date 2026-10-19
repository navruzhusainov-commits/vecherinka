#!/usr/bin/env python3
# PDF крупным шрифтом — 9 слайдов, концепция Чёрным по красному
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
IMG={"cover":"/tmp/foster_images/slide01_img1_Picture_1.jpg","about":"/tmp/foster_images/slide02_img3_Picture_5.jpg","concept_bg":"/tmp/foster_images/slide03_img4_Picture_1.jpg","hall1":"/tmp/foster_images/slide06_img7_Picture_5.jpg","hall2":"/tmp/foster_images/slide06_img8_Picture_6.jpg","dj_bar":"/tmp/foster_images/slide07_img9_Picture_5.jpg","vip":"/tmp/foster_images/slide07_img10_Picture_6.jpg","bar":"/tmp/foster_images/slide08_img11_Picture_5.jpg","loft1":"/tmp/foster_images/slide09_img12_Picture_5.jpg","timeline":"/tmp/foster_images/slide11_img18_Picture_5.jpg","multi1":"/tmp/foster_images/slide10_img14_Picture_5.jpg","multi2":"/tmp/foster_images/slide10_img15_Picture_6.jpg"}
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
def card(x_in,y_in,w_in,h_in,fill=DARK,radius=6,header=None,body=None,header_color=GOLD_LIGHT,body_color=BEIGE3,header_size=9,body_size=8.5):
    x,y,w,h=pt(x_in),pt(y_in),pt(w_in),pt(h_in)
    rect(x,y,w,h,fill=fill,radius=radius)
    if header: text(x+pt(0.22),y+pt(0.14),header,size=header_size,color=header_color,bold=True)
    if body: text_wrap(x+pt(0.22),y+pt(0.50),w-pt(0.44),body,size=body_size,color=body_color,leading=body_size+2.5)
def add_header(small,title,num="01 / 09"):
    text(pt(0.55),pt(0.28+0.12),small,size=8.5,color=GOLD,bold=True)
    text_wrap(pt(0.55),pt(0.60+0.15),pt(8.5),title,size=18,color=BEIGE,bold=True,leading=19)
    rect(pt(0.57),pt(1.45),pt(0.85),pt(0.06),fill=GOLD)
    text(pt(11.50),pt(0.28+0.12),num,size=8,color=MUTED,align="right",max_width=pt(1.30))
    text(pt(0.55),pt(7.05+0.10),"Фото: галерея Foster Club — Яндекс Карты",size=6,color=MUTED)
    text(pt(11.0),pt(7.05+0.10),f"FOSTER • {num.split('/')[0].strip()}",size=6,color=MUTED,bold=True,align="right",max_width=pt(1.80))

# SLIDE 1 COVER
draw_image(IMG["cover"],0,0,W,H,darken=0.50)
rect(pt(0.55),pt(0.35),pt(2.80),pt(0.95),fill=BEIGE,radius=6)
text(pt(0.70),pt(0.45+0.10),"FOSTER",size=10,color=DARK,bold=True,align="center",max_width=pt(2.80))
text(pt(0.70),pt(0.70+0.10),"NIGHT & KARAOKE CLUB",size=5,color=GOLD,bold=True,align="center",max_width=pt(2.80))
text(pt(0.70),pt(0.88+0.08),"Mriya Resort • Ялта • Крым",size=5,color=MUTED,align="center",max_width=pt(2.80))
text(pt(9.0),pt(0.35+0.10),"5 декабря 2026  •  Суббота  •  22:00–06:00",size=6,color=WHITE,align="right",max_width=pt(3.8))
text(pt(9.0),pt(0.62+0.10),"300–350 гостей  •  Total Black + Red",size=6,color=GOLD_LIGHT,align="right",max_width=pt(3.8))
text_wrap(pt(0.55),pt(1.65+0.20),pt(9),"GLEB FILIPCHENKOW",size=26,color=WHITE,bold=True,leading=26)
text_wrap(pt(0.55),pt(2.20+0.20),pt(9),"& NOEMY ABRANTES",size=26,color=GOLD,bold=True,leading=26)
text(pt(0.55),pt(2.85+0.10),"КОНЦЕПЦИЯ «ЧЁРНЫМ ПО КРАСНОМУ»  •  Первый привоз в Крым",size=8,color=GOLD_LIGHT,bold=True)
rect(pt(0.57),pt(3.05),pt(1.0),pt(0.06),fill=GOLD)
rect(pt(0.55),pt(3.35),pt(6.2),pt(1.75),fill=DARK,radius=6)
text(pt(0.75),pt(3.50+0.10),"◆  Клуб в красном свете и бархате",size=8,color=GOLD,bold=True)
text_wrap(pt(0.75),pt(3.80+0.10),pt(5.8),"Гости — total black с красной деталью. Минимум декора, максимум эффекта: красный свет, бархат, фото-зона. Перформеры — танцовщица и go-go в красном.",size=7.5,color=BEIGE3,leading=10)
for i,(num,lbl) in enumerate([("483 549 ₽","Ваша прибыль"),("63%","Рентабельность"),("330 000 ₽","Гарант")]):
    xl=pt(0.55+i*2.15)
    rect(xl,pt(5.35),pt(1.95),pt(0.90),fill=DARK,radius=5)
    text(xl,pt(5.50+0.10),num,size=10,color=WHITE,bold=True,align="center",max_width=pt(1.95))
    text(xl,pt(5.80+0.08),lbl,size=6,color=GOLD_LIGHT,bold=True,align="center",max_width=pt(1.95))
text(pt(0.55),pt(7.05+0.10),"Пересборка на подлинных фото Foster Club • Шрифт увеличен",size=6,color=MUTED)
text(pt(11.0),pt(7.05+0.10),"01 / 09",size=8,color=WHITE,align="right",max_width=pt(1.8))
c.showPage()

# SLIDE 2 HEADLINER
rect(0,0,W,H,fill=WHITE)
add_header("Х Е Д Л А Й Н Е Р","Gleb Filipchenkow & Noemy","02 / 09")
draw_image(IMG["hall1"],pt(7.00),pt(1.80),pt(5.78),pt(3.40))
y=pt(1.95)
for label,body in [("Глеб —","продюсер и соавтор вирусного трека «Mallorca» (взрыв в 2026). Первый приезд в Крым."),("Noemy —","соавтор релизов «Garota» и «HOJE», европейский вокал."),("Почему продаёт:","имя уже на слуху, отложенный спрос с лета, 5 декабря — единственное окно."),("Концепция:","«Чёрным по красному» — тотальный дресс-код и красный свет.")]:
    text(pt(0.55),y,"◆",size=8,color=GOLD,bold=True)
    text_wrap(pt(0.75),y,pt(6.15),f"{label} {body}",size=8.5,color=HexColor("#2A2A30"),leading=10)
    y+=pt(0.45)
card(0.55,5.50,3.95,1.40,fill=DARK,header="MALLORCA",body="Вирусный хит без бюджета:\nсотни тысяч просмотров,\nмосковские вечеринки 2026.\nДалее — трек с Серябкиной.",header_size=9,body_size=8)
card(4.70,5.50,3.95,1.40,fill=DARK,header="ПЕРВЫЙ ПРИВОЗ",body="В Крыму не выступали.\nЭффект эксклюзива +\nожидание с лета (сорвался\nиз-за топливного кризиса).",header_size=9,body_size=8)
card(8.85,5.50,3.93,1.40,fill=BURGUNDY,header="01:00 БЛЭКАУТ",body="Выход хедлайнеров —\nблэкаут и вспышка\nкрасного света. Кульминация\nвечеринки.",header_size=9,body_size=8)
c.showPage()

# SLIDE 3 КОНЦЕПЦИЯ ДЕТАЛЬНО
draw_image(IMG["concept_bg"],0,0,W,H,darken=0.50)
text(pt(0.55),pt(0.28+0.12),"К О Н Ц Е П Ц И Я   « Ч Ё Р Н Ы М   П О   К Р А С Н О М У »",size=8.5,color=GOLD,bold=True)
text_wrap(pt(0.55),pt(0.65+0.12),pt(9),"Total black + красная деталь",size=18,color=BEIGE,bold=True,leading=19)
rect(pt(0.57),pt(1.30),pt(0.85),pt(0.06),fill=GOLD)
text(pt(11.50),pt(0.28+0.12),"03 / 09",size=8,color=MUTED,align="right",max_width=pt(1.30))
card(0.55,1.75,3.95,1.65,fill=DARKER,header="ИДЕЯ",body="Клуб погружается в красный\nсвет и бархат. Гости — total\nblack с яркой красной\nдеталью (аксессуар, помада).",header_size=9,body_size=8.5)
card(4.70,1.75,3.95,1.65,fill=DARKER,header="ОФОРМЛЕНИЕ",body="Минимум декора — максимум\nэффекта: красный свет,\nкрасный бархат, фото-зона.\nМощно и недорого.",header_size=9,body_size=8.5)
card(8.85,1.75,3.93,1.65,fill=DARKER,header="ПЕРФОРМЕРЫ",body="Танцовщица и go-go\nв красном. Пластика +\nсвет = идеальный контент\nдля соцсетей.",header_size=9,body_size=8.5)
for i,(t,title,desc) in enumerate([("22:00","Сбор + warm-up","DJ, аперитив"),("23:30","Перформанс","Танец в красном"),("01:00","ХЕДЛАЙНЕРЫ","Блэкаут + красный"),("03:00","Afterparty","DJ-сет до 06:00")]):
    l=0.55+i*3.25
    rect(pt(l),pt(5.35),pt(3.00),pt(1.25),fill=BURGUNDY if "01:00" in t else DARK,radius=6)
    text(pt(l),pt(5.45+0.10),t,size=10,color=GOLD_LIGHT if "01:00" not in t else WHITE,bold=True,align="center",max_width=pt(3.00))
    text(pt(l),pt(5.80+0.08),title,size=8,color=BEIGE,bold=True,align="center",max_width=pt(3.00))
    text(pt(l),pt(6.10+0.08),desc,size=7,color=BEIGE3,align="center",max_width=pt(3.00))
c.showPage()

# SLIDE 4 КАКТУСА ДОХОД
rect(0,0,W,H,fill=WHITE)
add_header("Ф И Н А Н С Ы   К А К Т У С Ы  •  1","Доход от билетов","04 / 09")
draw_image(IMG["bar"],pt(0.55),pt(1.80),pt(5.20),pt(3.20))
text(pt(0.55),pt(5.15+0.08),"Foster Club — бар и танцпол, Яндекс Карты",size=6,color=MUTED,align="center",max_width=pt(5.20))
y=pt(1.95)
for label,body in [("Источник:","продажа билетов 300–350 гостей."),("Стратегия:","3 000 ₽ (старт) → 3 500 ₽ → 4 000 ₽ в день события."),("Базовый прогноз:","300 чел. × 3 400 ₽ средний чек = 1 020 000 ₽ выручки.")]:
    text(pt(6.20),y,"◆",size=8,color=GOLD,bold=True)
    text_wrap(pt(6.40),y,pt(6.30),f"{label} {body}",size=8.5,color=HexColor("#2A2A30"),leading=10)
    y+=pt(0.42)
rect(pt(6.20),pt(5.00),pt(6.50),pt(0.95),fill=BURGUNDY,radius=6)
text(pt(6.20),pt(5.15+0.10),"Базовый сценарий — 1 020 000 ₽ выручки",size=8,color=WHITE,bold=True,align="center",max_width=pt(6.50))
text(pt(6.20),pt(5.45+0.10),"При SOLD OUT 350 чел. × 3 400 ₽ = 1 190 000 ₽ → прибыль ~70 000 ₽",size=6,color=GOLD_LIGHT,align="center",max_width=pt(6.50))
c.showPage()

# SLIDE 5 КАКТУСА ЗАТРАТЫ
rect(0,0,W,H,fill=WHITE)
add_header("Ф И Н А Н С Ы   К А К Т У С Ы  •  2","Затраты организатора","05 / 09")
draw_image(IMG["hall1"],pt(0.55),pt(1.80),pt(6.20),pt(2.80))
draw_image(IMG["hall2"],pt(6.95),pt(1.80),pt(5.83),pt(2.80))
text(pt(0.55),pt(4.75+0.08),"Сцена и бар Foster — фото Яндекс Карты",size=6,color=MUTED,align="center",max_width=pt(6.20))
text(pt(6.95),pt(4.75+0.08),"Зеркальный бар — фото Яндекс Карты",size=6,color=MUTED,align="center",max_width=pt(5.83))
costs=[("Гонорар","550 000 ₽",BURGUNDY),("Коктейли","180 000 ₽",GOLD),("Шаттл","20 040 ₽",GOLD),("Логистика","~80 000 ₽",DARK),("Промо","~100 000 ₽",DARK),("Артисты","75 000 ₽",DARK),("Тех/Видео/Фото","85 000 ₽",DARK),("Охрана","30 000 ₽",DARK)]
for i,(name,price,fill) in enumerate(costs):
    row=i//4; col=i%4; l=0.55+col*3.25; t=5.15+row*1.20
    rect(pt(l),pt(t),pt(3.00),pt(1.00),fill=fill,radius=6)
    text(pt(l),pt(t+0.12),name,size=7,color=WHITE if fill==BURGUNDY else (DARK if fill==GOLD else GOLD_LIGHT),bold=True,align="center",max_width=pt(3.00))
    text(pt(l),pt(t+0.40),price,size=8,color=WHITE if fill!=GOLD else DARK,bold=True,align="center",max_width=pt(3.00))
    if name in ["Коктейли","Шаттл"]: text(pt(l),pt(t+0.70),"(фикс → вам)",size=5,color=WHITE if fill==BURGUNDY else MUTED,align="center",max_width=pt(3.00))
rect(pt(0.55),pt(6.65),pt(12.23),pt(0.50),fill=DARK,radius=6)
text(pt(0.55),pt(6.75+0.10),"ИТОГО ЗАТРАТЫ КАКТУСЫ:  ~1 120 040 ₽",size=8,color=WHITE,bold=True,align="center",max_width=pt(12.23))
c.showPage()

# SLIDE 6 КАКТУСА РЕЗУЛЬТАТ
rect(0,0,W,H,fill=WHITE)
add_header("Ф И Н А Н С Ы   К А К Т У С Ы  •  3","Результат: риск на них","06 / 09")
draw_image(IMG["timeline"],pt(0.55),pt(1.80),pt(12.23),pt(2.40))
text(pt(0.55),pt(4.35+0.08),"Тайминг продаж билетов — как в шаблоне Foster Club",size=6,color=MUTED,align="center",max_width=pt(12.23))
scenarios=[("250 чел.","Пессимистичный","825 000 ₽","–295 040 ₽",DARK),("300 чел.","Базовый","1 020 000 ₽","–100 040 ₽",DARK),("350 чел.","SOLD OUT","1 190 000 ₽","+69 960 ₽",BURGUNDY)]
for i,(cnt,title,rev,res,fill) in enumerate(scenarios):
    l=0.55+i*4.20
    rect(pt(l),pt(4.80),pt(3.95),pt(2.00),fill=fill,radius=6)
    text(pt(l),pt(4.95+0.10),cnt,size=12,color=WHITE if fill==BURGUNDY else GOLD_LIGHT,bold=True,align="center",max_width=pt(3.95))
    text(pt(l),pt(5.35+0.08),title,size=8,color=BEIGE,bold=True,align="center",max_width=pt(3.95))
    text(pt(l),pt(5.70+0.08),"Выручка",size=6,color=MUTED,align="center",max_width=pt(1.97))
    text(pt(l+2.08),pt(5.70+0.08),"Итог",size=6,color=MUTED,align="center",max_width=pt(1.97))
    text(pt(l),pt(5.95+0.10),rev,size=9,color=WHITE if fill==BURGUNDY else BEIGE,bold=True,align="center",max_width=pt(1.97))
    text(pt(l+2.08),pt(5.95+0.10),res,size=9,color=HexColor("#FF8A90") if "–" in res else HexColor("#7ED3A0"),bold=True,align="center",max_width=pt(1.97))
rect(pt(0.55),pt(6.05),pt(12.23),pt(0.65),fill=BEIGE,radius=6)
text(pt(0.75),pt(6.15+0.10),"💡 Вывод:",size=8,color=BURGUNDY,bold=True)
text_wrap(pt(0.75),pt(6.45),pt(11.8),"Кактуса выходит в плюс только при SOLD OUT — значит, будет продавать до последнего билета. Ваши 200 040 ₽ зафиксированы в договоре.",size=7,color=DARK,leading=9)
c.showPage()

# SLIDE 7 FOSTER ЗАТРАТЫ
rect(0,0,W,H,fill=WHITE)
add_header("Ф И Н А Н С Ы   F O S T E R  •  1","Наши затраты — 284 476 ₽","07 / 09")
draw_image(IMG["multi1"],pt(0.55),pt(1.80),pt(2.80),pt(3.40))
draw_image(IMG["multi2"],pt(3.50),pt(1.80),pt(3.90),pt(3.40))
text(pt(0.55),pt(5.35+0.08),"Танцпол Foster",size=6,color=MUTED,align="center",max_width=pt(2.80))
text(pt(3.50),pt(5.35+0.08),"VIP-зона — фото Яндекс Карты",size=6,color=MUTED,align="center",max_width=pt(3.90))
y=pt(1.95)
for name,price,note in [("Персонал клуба","165 470 ₽","официанты, бармены, повара"),("Себестоимость бара","92 006 ₽","25% от 368 025 ₽"),("Себестоимость кухни","27 000 ₽","27% от 100 000 ₽")]:
    text(pt(7.60),y,name,size=8,color=DARK,bold=True)
    text(pt(7.60),y+pt(0.18),price,size=10,color=DARK,bold=True,align="right",max_width=pt(5.2))
    text(pt(7.60),y+pt(0.32),note,size=7,color=MUTED)
    y+=pt(0.48)
text(pt(7.60),y+pt(0.15),"Итого затраты:  284 476 ₽",size=9,color=BURGUNDY,bold=True)
card(0.55,5.80,3.95,1.20,fill=DARK,header="ПЕРСОНАЛ 58%",body="165 470 ₽ — официанты,\nбармены, повара, менеджмент.\nКрупнейшая статья.",header_size=9,body_size=8)
card(4.70,5.80,3.95,1.20,fill=DARK,header="ФУДКОСТ 42%",body="Бар 25% + Кухня 27% —\nнормальные показатели.\nПокрываются выручкой.",header_size=9,body_size=8)
card(8.85,5.80,3.93,1.20,fill=DARK,header="0 ₽ ПРОЖИВАНИЕ",body="Полностью исключено.\nНи скидок, ни затрат.\nАренда — 0 ₽.",header_size=9,body_size=8)
c.showPage()

# SLIDE 8 FOSTER ДОХОДЫ
rect(0,0,W,H,fill=WHITE)
add_header("Ф И Н А Н С Ы   F O S T E R  •  2","Наши доходы — 768 025 ₽","08 / 09")
draw_image(IMG["bar"],pt(0.55),pt(1.80),pt(3.90),pt(3.30))
text(pt(0.55),pt(5.25+0.08),"Бар Foster — Яндекс Карты",size=6,color=MUTED,align="center",max_width=pt(3.90))
draw_image(IMG["loft1"],pt(4.65),pt(1.80),pt(3.90),pt(3.30))
text(pt(4.65),pt(5.25+0.08),"Лаунж-зона — фото клуба",size=6,color=MUTED,align="center",max_width=pt(3.90))
y=pt(1.95)
for name,price,note in [("Бар","368 025 ₽","валовая выручка"),("Кухня","100 000 ₽","валовая выручка"),("◆ Коктейли","180 000 ₽","300 × 600 ₽ → вам"),("◆ Столы","50 000 ₽","4×5 000 + 3×10 000"),("◆ VIP №1","40 000 ₽","депозит"),("◆ VIP №2","30 000 ₽","депозит")]:
    is_hl="◆" in name; col=GOLD if is_hl else DARK
    text(pt(8.85),y,name,size=8,color=col,bold=True)
    text(pt(10.20),y,price,size=8.5,color=DARK,bold=True)
    text(pt(11.30),y,"— "+note,size=6,color=MUTED)
    y+=pt(0.30)
text(pt(8.85),y+pt(0.12),"Итого: 768 025 ₽ + кальяны",size=8,color=BURGUNDY,bold=True)
rect(pt(0.55),pt(5.70),pt(12.23),pt(1.10),fill=BURGUNDY,radius=6)
text(pt(0.55),pt(5.85+0.10),"ГАРАНТИРОВАНО ДОГОВОРОМ — 330 000 ₽",size=9,color=WHITE,bold=True,align="center",max_width=pt(12.23))
text_wrap(pt(0.75),pt(6.20),pt(11.8),"180 000 ₽ коктейли (фикс) + 50 000 ₽ столы + 70 000 ₽ VIP-залы — независимо от явки. Кальяны от Макао — по факту.",size=7.5,color=GOLD_LIGHT,leading=9,align="center")
c.showPage()

# SLIDE 9 ФИНАЛ
draw_image(IMG["cover"],0,0,W,H,darken=0.52)
text(pt(0.55),pt(0.28+0.12),"Г Л А В Н Ы Й   И Т О Г",size=8.5,color=GOLD,bold=True)
text_wrap(pt(0.55),pt(0.65+0.12),pt(9),"Сделка безопасна и",size=22,color=BEIGE,bold=True,leading=23)
text_wrap(pt(0.55),pt(1.10+0.12),pt(9),"высокорентабельна",size=22,color=GOLD,bold=True,leading=23)
rect(pt(0.57),pt(1.65),pt(1.0),pt(0.06),fill=GOLD)
rect(pt(0.55),pt(1.95),pt(6.5),pt(1.80),fill=DARK,radius=6)
text(pt(0.75),pt(2.10+0.08),"ВАША ЧИСТАЯ ПРИБЫЛЬ",size=7,color=GOLD_LIGHT,bold=True)
text(pt(0.75),pt(2.40+0.10),"483 549 ₽",size=20,color=WHITE,bold=True)
text(pt(4.50),pt(2.45+0.10),"~63%\nрентабельность",size=8,color=GOLD,bold=True,align="center",max_width=pt(1.8))
text(pt(0.75),pt(3.05+0.08),"768 025 ₽ доходы – 284 476 ₽ затраты  + прибыль с кальянов",size=7,color=BEIGE3)
rect(pt(7.30),pt(1.95),pt(5.30),pt(1.80),fill=BEIGE,radius=6)
text(pt(7.30),pt(2.10+0.08),"ЧТО ЗАФИКСИРОВАТЬ В ДОГОВОРЕ",size=8,color=BURGUNDY,bold=True,align="center",max_width=pt(5.30))
text_wrap(pt(7.50),pt(2.45),pt(4.90),"✓  180 000 ₽ — коктейли (фикс)\n✓  20 040 ₽ — шаттл (фикс)\n✓  1–3 рабочих дня после события",size=7.5,color=DARK,leading=10)
rect(pt(0.55),pt(5.90),pt(12.23),pt(0.95),fill=BEIGE,radius=6)
text(pt(0.55),pt(6.05+0.10),"Вы получаете почти полмиллиона рублей чистыми, не вкладываясь в артистов, аренду и оборудование.",size=8,color=DARK,bold=True,align="center",max_width=pt(12.23))
text(pt(0.55),pt(6.40+0.10),"Задача — жёстко зафиксировать компенсации. Организатор максимально мотивирован продать SOLD OUT.",size=7,color=MUTED,align="center",max_width=pt(12.23))
text(pt(0.55),pt(7.05+0.10),"Презентация для переговоров • Шрифт увеличен для чтения • Концепция «Чёрным по красному»",size=6,color=MUTED)
text(pt(11.0),pt(7.05+0.10),"09 / 09",size=8,color=WHITE,align="right",max_width=pt(1.8))
c.showPage()
c.save()
print(f"Saved PDF {OUT}")

