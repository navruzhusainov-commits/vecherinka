# -*- coding: utf-8 -*-
"""PDF «Вложения и планируемый доход» — вечер Gleb & Noemy в Foster Club (5 декабря 2026).
Данные из презентаций репозитория + уточнения заказчика + райдер артистов (ред. 28.10.2026, файл в репозитории):
  - аренда клуба 200 000 ₽ и аренда оборудования 250 000 ₽ предоставляются Кактусе БЕСПЛАТНО;
  - персонал — вызов в переработку ~160 000 ₽;
  - бар 368 025 ₽ + кухня 100 000 ₽ = 468 025 ₽ — фактическая выручка прошлого года (бенчмарк);
  - Кактуса привозит других артистов сам и берёт ВЕСЬ РАЙДЕР (тех + бытовой) на себя.
Оформление — красно-чёрный (#E60613 / #141419).
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, PageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------- шрифты (кириллица) ----------
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Italic', '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'))

# ---------- палитра ----------
RED = colors.HexColor('#E60613')
DARK = colors.HexColor('#141419')
GREY = colors.HexColor('#EDEAE2')
MUTED = colors.HexColor('#6B6558')
WHITE = colors.white

PAGE_W, PAGE_H = A4
MARGIN = 14 * mm


def st(name, **kw):
    base = dict(fontName='DejaVu', fontSize=9, leading=12, textColor=DARK, alignment=TA_LEFT)
    base.update(kw)
    return ParagraphStyle(name, **base)


S_TITLE = st('title', fontName='DejaVu-Bold', fontSize=20, leading=24, textColor=WHITE)
S_SUB = st('sub', fontSize=9.5, leading=13, textColor=colors.HexColor('#F5F1E8'))
S_H2 = st('h2', fontName='DejaVu-Bold', fontSize=13, leading=16, textColor=WHITE)
S_CELL = st('cell', fontSize=8.5, leading=11)
S_CELL_B = st('cellb', fontName='DejaVu-Bold', fontSize=8.5, leading=11)
S_CELL_W = st('cellw', fontName='DejaVu-Bold', fontSize=8.5, leading=11, textColor=WHITE)
S_CELL_BR = st('cellbr', fontName='DejaVu-Bold', fontSize=8.5, leading=11, alignment=TA_RIGHT)
S_NOTE = st('note', fontName='DejaVu-Italic', fontSize=8, leading=10.5, textColor=MUTED)
S_BODY = st('body', fontSize=9.5, leading=13)


def P(txt, style=S_CELL):
    return Paragraph(txt, style)


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(DARK)
    canvas.rect(0, PAGE_H - 10 * mm, PAGE_W, 10 * mm, stroke=0, fill=1)
    canvas.setFillColor(RED)
    canvas.rect(0, PAGE_H - 10.8 * mm, PAGE_W, 0.8 * mm, stroke=0, fill=1)
    canvas.setFillColor(WHITE)
    canvas.setFont('DejaVu-Bold', 8)
    canvas.drawString(MARGIN, PAGE_H - 6.8 * mm, 'FOSTER × КАКТУСА — ВЛОЖЕНИЯ И ПЛАНИРУЕМЫЙ ДОХОД')
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 6.8 * mm, '5 ДЕКАБРЯ 2026 • FOSTER NIGHTCLUB, КРЫМ')
    canvas.setFillColor(MUTED)
    canvas.setFont('DejaVu', 7.5)
    canvas.drawString(MARGIN, 8 * mm, 'Источники: презентации репозитория + райдер артистов ред. 28.10.2026 + уточнения заказчика (аренда 200К/оборудование 250К бесплатно, персонал ~160К, факт-2025 468 025 ₽, весь райдер на Кактусе)')
    canvas.drawRightString(PAGE_W - MARGIN, 8 * mm, f'Стр. {doc.page}')
    canvas.restoreState()


doc = BaseDocTemplate('/home/user/vecherinka/Вложения_и_планируемый_доход.pdf',
                      pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                      topMargin=16 * mm, bottomMargin=14 * mm,
                      title='Вложения и планируемый доход — Gleb & Noemy, Foster Club, 5 декабря 2026',
                      author='vecherinka project')
frame = Frame(MARGIN, 14 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 30 * mm, id='main')
doc.addPageTemplates([PageTemplate(id='page', frames=[frame], onPage=header_footer)])

CW = PAGE_W - 2 * MARGIN
story = []


def section_bar(txt):
    t = Table([[P(txt, S_H2)]], colWidths=[CW])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), RED),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    return t


def make_table(header, rows, widths, total_rows=(), right_cols=()):
    data = [[P(h, S_CELL_W) for h in header]]
    for row in rows:
        data.append([P(str(cell), S_CELL_BR if i in right_cols else S_CELL) for i, cell in enumerate(row)])
    n_body = len(rows)
    for values, red in total_rows:
        data.append([P(str(c), S_CELL_BR if i in right_cols else S_CELL_W) for i, c in enumerate(values)])
    t = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ('BACKGROUND', (0, 0), (-1, 0), DARK),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#B9B2A4')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 3.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
    ]
    for i in range(1, n_body + 1):
        if i % 2 == 0:
            style.append(('BACKGROUND', (0, i), (-1, i), GREY))
    for j, (values, red) in enumerate(total_rows):
        row_idx = n_body + 1 + j
        style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), RED if red else DARK))
    t.setStyle(TableStyle(style))
    return t


# ================= ТИТУЛ =================
title_tbl = Table([
    [P('ВЛОЖЕНИЯ И ПЛАНИРУЕМЫЙ ДОХОД', S_TITLE)],
    [P('Вечер <b>Gleb Filipchenkow &amp; Noemy Abrantes</b> — «Чёрным по красному» • первый привоз в Крым (хиты «Mallorca», «Garota»/«HOJE)', S_SUB)],
    [P('<b>Суббота, 5 декабря 2026 • 22:00–06:00 • Foster Nightclub, Крым</b> (Mriya Resort &amp; SPA, Ялта, с. Оползневое) • 300–350 гостей, цель SOLD OUT • билеты 3 000 → 3 500 → 4 000 ₽ • старт продаж 01.11.2026', S_SUB)],
    [P('<b>Условия площадки:</b> аренда клуба (рыночная 200 000 ₽) и оборудования (250 000 ₽) — Кактусе <b>бесплатно</b> • персонал — в переработку (~160 000 ₽) • бар + кухня в прошлом году дали 468 025 ₽ • <b>Кактуса привозит других артистов сам и берёт весь райдер на себя</b> (райдер ред. 28.10.2026 — файл в репозитории)', S_SUB)],
], colWidths=[CW])
title_tbl.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), DARK),
    ('LEFTPADDING', (0, 0), (-1, -1), 12),
    ('RIGHTPADDING', (0, 0), (-1, -1), 12),
    ('TOPPADDING', (0, 0), (0, 0), 12),
    ('BOTTOMPADDING', (0, -1), (-1, -1), 12),
    ('LINEBELOW', (0, -1), (-1, -1), 3, RED),
]))
story.append(title_tbl)
story.append(Spacer(1, 6 * mm))

# KPI-плашки
kpi_style = lambda: st('k', fontName='DejaVu-Bold', fontSize=12.5, leading=15, textColor=WHITE, alignment=TA_CENTER)
kpi = Table([[
    P('729 006 ₽<br/><font size="7" color="#B9B2A4">ВСЕГО ВЛОЖЕНИЙ FOSTER</font>', kpi_style()),
    P('~1 240 040 ₽<br/><font size="7" color="#B9B2A4">ВЛОЖЕНИЯ КАКТУСЫ (весь райдер)</font>', kpi_style()),
    P('768 025 ₽<br/><font size="7" color="#B9B2A4">ДОХОД FOSTER + КАЛЬЯНЫ</font>', kpi_style()),
    P('489 019 ₽<br/><font size="7" color="#B9B2A4">ОПЕРАЦ. ПРИБЫЛЬ (~64%)</font>', kpi_style()),
    P('330 000 ₽<br/><font size="7" color="#B9B2A4">ГАРАНТ ПО ДОГОВОРУ</font>', kpi_style()),
]], colWidths=[CW / 5.0] * 5)
kpi.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (2, 0), DARK),
    ('BACKGROUND', (3, 0), (4, 0), RED),
    ('GRID', (0, 0), (-1, -1), 1, WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('TOPPADDING', (0, 0), (-1, -1), 7),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
]))
story.append(kpi)
story.append(Spacer(1, 6 * mm))

# ================= СВОДКА =================
story.append(section_bar('ИТОГОВАЯ СВОДКА ПО СТОРОНАМ'))
story.append(Spacer(1, 2 * mm))
summary_rows = [
    ['Вложения (затраты)',
     '729 006 ₽ всего = 279 006 ₽ операционка (персонал ~160К + бар 92К + кухня 27К) + 450 000 ₽ бесплатная аренда клуба (200К) и оборудования (250К); артисты и райдер клубу — 0 ₽',
     '~1 240 040 ₽ — гонорар 550К + логистика 80К + проживание 60К + суточные 40К + бытовой райдер 20К + техрайдер 30К + артисты 75К + промо 100К + видео/фото 55К + охрана 30К + компенсации клубу 200 040 ₽; аренда и оборудование — 0 ₽ (экономия 450 000 ₽)'],
    ['Планируемый доход',
     '768 025 ₽ + кальяны Макао (бар 368 025 + кухня 100 000 — факт прошлого года; + коктейли 180 000 + столы 50 000 + VIP 70 000)',
     'Билеты: 825 000 ₽ (250 × 3 300) / 1 020 000 ₽ (300 × 3 400) / 1 190 000 ₽ (350 × 3 400) / 1 400 000 ₽ (350 × 4 000)'],
    ['Планируемый результат',
     'Операц. прибыль +489 019 ₽ (~64%); с учётом бесплатной аренды/площадки +39 019 ₽',
     '–415 040 ₽ (250) / –220 040 ₽ (300) / –50 040 ₽ (350 × 3 400) / +159 960 ₽ (350 × 4 000)'],
    ['Гарант / риск',
     '330 000 ₽ гарантированы договором (коктейли 180К + столы 50К + VIP 70К) + 20 040 ₽ шаттл + бесплатная аренда/площадка 450К; артисты и райдер — 0 ₽; оплата 1–3 дня после события',
     'Весь райдер на нём → в плюс только при средней цене билета выше ~3 543 ₽ (продажи по 4 000 ₽ к дате события)'],
]
story.append(make_table(['Показатель', 'Клуб FOSTER (площадка)', 'Организатор КАКТУСА'],
                        summary_rows, [CW * 0.16, CW * 0.42, CW * 0.42]))
story.append(Spacer(1, 3 * mm))
story.append(P('Участники: <b>FOSTER</b> — площадка (клуб, бар, кухня, персонал, VIP; аренда клуба и оборудования Кактусе — бесплатно; артисты и райдер — 0 ₽) • <b>КАКТУСА</b> — организатор (все риски, весь райдер, других артистов привозит сам, продажи билетов) • <b>MRIYA</b> — партнёр (проживание артистов без скидки — за счёт Кактусы по райдеру).', S_BODY))

story.append(PageBreak())

# ================= ВЛОЖЕНИЯ FOSTER =================
story.append(section_bar('1. ВЛОЖЕНИЯ КЛУБА FOSTER — ОПЕРАЦИОНКА 279 006 ₽ + АРЕНДА/ПЛОЩАДКА 450 000 ₽ = 729 006 ₽'))
story.append(Spacer(1, 2 * mm))
story.append(P('<b>1.1 Операционные (денежные) затраты</b>', S_BODY))
story.append(Spacer(1, 1.5 * mm))
foster_oper = [
    ['Персонал — вызов в переработку', '~160 000 ₽', 'Подготовка и проведение мероприятия (ориентир; вместо 165 470 ₽ из презентации)'],
    ['Себестоимость бара (фудкост)', '92 006 ₽', '25% от валовой выручки бара 368 025 ₽'],
    ['Себестоимость кухни (фудкост)', '27 000 ₽', '27% от валовой выручки кухни 100 000 ₽'],
]
story.append(make_table(['Статья затрат', 'Сумма', 'Примечание'], foster_oper,
                        [CW * 0.36, CW * 0.16, CW * 0.48],
                        total_rows=[(['ИТОГО операционные затраты', '279 006 ₽', 'Полностью покрываются доходами клуба'], False)],
                        right_cols=(1,)))
story.append(Spacer(1, 4 * mm))
story.append(P('<b>1.2 Предоставлено Кактусе бесплатно (вложение клуба, 0 ₽ для организатора)</b> — артисты и райдер также 0 ₽ для клуба', S_BODY))
story.append(Spacer(1, 1.5 * mm))
foster_free = [
    ['Аренда клуба Foster', '200 000 ₽', 'Рыночная стоимость; для Кактусы — 0 ₽ (бесплатно)'],
    ['Аренда оборудования (свет/звук/сцена)', '250 000 ₽', 'Рыночная стоимость; для Кактусы — 0 ₽ (бесплатно)'],
    ['Артисты (все, включая «других», которых привозит Кактуса) и райдер', '0 ₽', 'Полностью за счёт Кактусы по райдеру — клубу бесплатно'],
]
story.append(make_table(['Статья', 'Рыночная стоимость', 'Условие'], foster_free,
                        [CW * 0.40, CW * 0.18, CW * 0.42],
                        total_rows=[(['ИТОГО бесплатно для Кактусы (аренда + оборудование)', '450 000 ₽', 'Экономия организатора = 450 000 ₽'], False),
                                    (['ВСЕГО ВЛОЖЕНИЙ FOSTER', '729 006 ₽', '279 006 операционка + 450 000 аренда/площадка и оборудование'], True)],
                        right_cols=(1,)))
story.append(Spacer(1, 6 * mm))

# ================= ВЛОЖЕНИЯ КАКТУСЫ =================
story.append(section_bar('2. ВЛОЖЕНИЯ КАКТУСЫ — ВСЕГО ~1 240 040 ₽ (ВЕСЬ РАЙДЕР НА НЁМ, ред. 28.10.2026)'))
story.append(Spacer(1, 2 * mm))
kaktus_rows = [
    ['Гонорар хедлайнеров', '550 000 ₽', 'Фикс с УСН — райдер п.1.1 (05.12.2026, Крым)'],
    ['Компенсация клубу за welcome-коктейли (300 × 600 ₽)', '180 000 ₽', '→ Foster, фикс в договоре, независимо от явки'],
    ['Оплата шаттла (охрана, клининг, хостес)', '20 040 ₽', '→ Foster, фикс в договоре'],
    ['Логистика: 2 авиабилета эконом+/бизнес + трансфер (Vito/Zeekr, бизнес-класс)', '~80 000 ₽', 'Райдер п.2 — проезд за счёт заказчика'],
    ['Проживание артистов: 2 номера Люкс/Полулюкс 4*, минибар + завтраки, ~2 ночи', '~60 000 ₽', 'Райдер п.3 — Mriya без скидки'],
    ['Суточные: 10 000 ₽/чел/сут × 2 чел × 2 сут', '40 000 ₽', 'Райдер п.2.8 (или питание п.4)'],
    ['Бытовой райдер: гримёрная (вода, сок, сэндвичи, фрукты, вино, пиво, чай/кофе, официант)', '~20 000 ₽', 'Райдер п.4–5 — оценка'],
    ['Техрайдер: CDJ-3000 ×2, DJM-A9/V10, ULX-D ×2, in-ear, подиум 3×2 м, мониторы', '~30 000 ₽', 'Райдер п.6 — ПОЛНОСТЬЮ на Кактусе, сверх бесплатного оборудования клуба'],
    ['Другие артисты от Кактусы (перформанс, go-go, локальный лайн-ап)', '75 000 ₽', 'Кактуса привозит сам; клубу 0 ₽'],
    ['Продвижение, блогеры, афиши', '~100 000 ₽', 'Макет афиши — по п.7.3'],
    ['Видеосъёмка и монтаж + фотоотчёт', '55 000 ₽', 'Продакшн (п.7.2)'],
    ['Внешняя охрана (в т.ч. личная охрана артистов, п.5.7)', '~30 000 ₽', 'Сверх шаттла'],
    ['Аренда клуба Foster', '0 ₽', 'Рыночная 200 000 ₽ — БЕСПЛАТНО от Foster'],
    ['Аренда оборудования клуба', '0 ₽', 'Рыночная 250 000 ₽ — БЕСПЛАТНО от Foster'],
]
story.append(make_table(['Статья затрат', 'Сумма', 'Источник / примечание'], kaktus_rows,
                        [CW * 0.44, CW * 0.14, CW * 0.42],
                        total_rows=[(['ИТОГО вложения Кактусы', '~1 240 040 ₽', 'В плюс только при ср. цене билета выше ~3 543 ₽; экономия на аренде/оборудовании 450 000 ₽'], True)],
                        right_cols=(1,)))
story.append(Spacer(1, 2 * mm))
story.append(P('Фиксированные компенсации клубу Foster: <b>200 040 ₽</b> = 180 000 ₽ (коктейли) + 20 040 ₽ (шаттл) — не зависят от явки. По п.7.12 райдера артистам положены 10 билетов на танцпол + 2 стола (~34 000 ₽ упущенной выручки от билетов).', S_NOTE))

story.append(PageBreak())

# ================= ДОХОДЫ И ПРИБЫЛЬ =================
story.append(section_bar('3. ПЛАНИРУЕМЫЙ ДОХОД КЛУБА FOSTER — 768 025 ₽ + КАЛЬЯНЫ МАКАО'))
story.append(Spacer(1, 2 * mm))
income_rows = [
    ['Бар (валовая выручка)', '368 025 ₽', 'ФАКТ прошлого года — такое же мероприятие с Кактусой'],
    ['Кухня (валовая выручка)', '100 000 ₽', 'ФАКТ прошлого года (бар + кухня = 468 025 ₽ — бенчмарк)'],
    ['Welcome-коктейли (компенсация от Кактусы)', '180 000 ₽', '300 × 600 ₽ — фикс, независимо от явки'],
    ['Депозиты столов', '50 000 ₽', '4 × 5 000 + 3 × 10 000'],
    ['VIP №1 (депозит)', '40 000 ₽', ''],
    ['VIP №2 (депозит)', '30 000 ₽', ''],
]
story.append(make_table(['Источник дохода', 'Сумма', 'Примечание'], income_rows,
                        [CW * 0.40, CW * 0.16, CW * 0.44],
                        total_rows=[(['ИТОГО доходы Foster', '768 025 ₽', '+ кальяны от Макао — по факту продаж, полностью клубу'], False)],
                        right_cols=(1,)))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('4. ПРИБЫЛЬ КЛУБА FOSTER ЗА ВЕЧЕР'))
story.append(Spacer(1, 2 * mm))
profit_rows = [
    ['Доходы', '768 025 ₽', 'Бар 368 025 + кухня 100 000 (факт прошлого года) + коктейли 180 000 + столы 50 000 + VIP 70 000'],
    ['– Операционные затраты', '–279 006 ₽', 'Персонал ~160 000 (переработка) + бар 92 006 + кухня 27 000'],
    ['= Операционная прибыль', '489 019 ₽', '768 025 – 279 006 = 489 019 ₽ + кальяны Макао сверху'],
    ['Рентабельность (операционная)', '~64%', '489 019 / 768 025'],
    ['– Бесплатная аренда клуба + оборудования', '–450 000 ₽', 'Рыночная 200 000 + 250 000 ₽ — Кактусе бесплатно (вложение клуба)'],
    ['= Чистая прибыль с учётом аренды/площадки', '39 019 ₽', '768 025 – 729 006 = 39 019 ₽ — клуб всё равно в плюсе'],
    ['Гарантировано договором (фикс)', '330 000 ₽', '180 000 коктейли + 50 000 столы + 70 000 VIP №1/№2 — независимо от явки; даже при 200 гостях в плюсе (330 000 – 279 006 = +50 994 ₽). + 20 040 ₽ шаттл. Оплата 1–3 рабочих дня после события'],
]
story.append(make_table(['Показатель', 'Сумма', 'Формула / примечание'], profit_rows,
                        [CW * 0.30, CW * 0.16, CW * 0.54], right_cols=(1,)))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('5. СЦЕНАРИИ ДОХОДА КАКТУСЫ — ПРОДАЖА БИЛЕТОВ (ЗАТРАТЫ 1 240 040 ₽)'))
story.append(Spacer(1, 2 * mm))
scen_rows = [
    ['Пессимистичный', '250', '3 300 ₽', '825 000 ₽', '–415 040 ₽ (убыток)'],
    ['Базовый', '300', '3 400 ₽', '1 020 000 ₽', '–220 040 ₽ (убыток)'],
    ['SOLD OUT × средняя цена', '350', '3 400 ₽', '1 190 000 ₽', '–50 040 ₽ (убыток)'],
    ['SOLD OUT × финальная цена', '350', '4 000 ₽', '1 400 000 ₽', '+159 960 ₽ (прибыль)'],
]
story.append(make_table(['Сценарий', 'Гости, чел.', 'Средний билет', 'Выручка', 'Результат (– 1 240 040 ₽ затрат)'],
                        scen_rows, [CW * 0.24, CW * 0.10, CW * 0.14, CW * 0.15, CW * 0.37], right_cols=(1, 2, 3, 4)))
story.append(Spacer(1, 2 * mm))
story.append(P('Ступенчатая цена билетов: <b>3 000 ₽</b> (01–07.11) → <b>3 500 ₽</b> (08–14.11) → <b>4 000 ₽</b> (с 15.11 / в день события). Поскольку Кактуса берёт весь райдер на себя, при средней цене 3 400 ₽ он в минусе даже при SOLD OUT 350; в плюс выходит только при средней цене выше ~3 543 ₽/билет (1 240 040 / 350), т.е. при продаже билетов по финальной цене 4 000 ₽. Клуб Foster защищён: 330 000 ₽ гарант + 200 040 ₽ компенсаций зафиксированы в договоре независимо от результата; аренда и оборудование — бесплатно (экономия 450 000 ₽).', S_NOTE))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('ЧТО ФИКСИРОВАТЬ В ДОГОВОРЕ'))
story.append(Spacer(1, 2 * mm))
checklist = [
    '☑ 180 000 ₽ — компенсация коктейлей (300 × 600 ₽, фикс)',
    '☑ 20 040 ₽ — оплата шаттла (охрана / клининг / хостес)',
    '☑ 50 000 ₽ — депозиты столов + 70 000 ₽ — VIP №1/№2 (итого гарант 330 000 ₽)',
    '☑ Аренда клуба (рыночная 200 000 ₽) и оборудования (250 000 ₽) — БЕСПЛАТНО для Кактусы',
    '☑ ВЕСЬ РАЙДЕР артистов (тех + бытовой) и других артистов — за счёт Кактусы, клубу 0 ₽',
    '☑ Оплата 1–3 рабочих дня после события',
    '☑ Кальяны Макао — по факту продаж, полностью в пользу клуба',
]
for item in checklist:
    story.append(P(item, S_BODY))
    story.append(Spacer(1, 1.2 * mm))
story.append(Spacer(1, 2 * mm))
story.append(P('<b>Следующие шаги:</b> согласовать договор (фикс 330 000 ₽ + компенсации + бесплатная аренда/площадка 450К + райдер на Кактусе) → старт продаж 01.11.2026 → 5 декабря 2026, SOLD OUT 350 гостей.', S_BODY))

doc.build(story)
print('Сохранено: /home/user/vecherinka/Вложения_и_планируемый_доход.pdf')
