# -*- coding: utf-8 -*-
"""PDF «Вложения и планируемый доход» — вечер Gleb & Noemy в Foster Club (5 декабря 2026).
Данные собраны из презентаций репозитория. Оформление — красно-чёрный (#E60613 / #141419).
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, PageBreak, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------- шрифты (кириллица) ----------
pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Italic', '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'))

# ---------- палитра ----------
RED = colors.HexColor('#E60613')
DARK = colors.HexColor('#141419')
LIGHT = colors.HexColor('#F5F1E8')
GREY = colors.HexColor('#EDEAE2')
MUTED = colors.HexColor('#6B6558')
WHITE = colors.white

PAGE_W, PAGE_H = A4
MARGIN = 14 * mm

# ---------- стили ----------
def st(name, **kw):
    base = dict(fontName='DejaVu', fontSize=9, leading=12, textColor=DARK, alignment=TA_LEFT)
    base.update(kw)
    return ParagraphStyle(name, **base)

S_TITLE = st('title', fontName='DejaVu-Bold', fontSize=20, leading=24, textColor=WHITE)
S_SUB = st('sub', fontSize=9.5, leading=13, textColor=colors.HexColor('#F5F1E8'))
S_H2 = st('h2', fontName='DejaVu-Bold', fontSize=13, leading=16, textColor=WHITE)
S_H3 = st('h3', fontName='DejaVu-Bold', fontSize=10.5, leading=13, textColor=RED)
S_CELL = st('cell', fontSize=8.5, leading=11)
S_CELL_B = st('cellb', fontName='DejaVu-Bold', fontSize=8.5, leading=11)
S_CELL_W = st('cellw', fontName='DejaVu-Bold', fontSize=8.5, leading=11, textColor=WHITE)
S_CELL_R = st('cellr', fontSize=8.5, leading=11, alignment=TA_RIGHT)
S_CELL_BR = st('cellbr', fontName='DejaVu-Bold', fontSize=8.5, leading=11, alignment=TA_RIGHT)
S_NOTE = st('note', fontName='DejaVu-Italic', fontSize=8, leading=10.5, textColor=MUTED)
S_BODY = st('body', fontSize=9.5, leading=13)


def P(txt, style=S_CELL):
    return Paragraph(txt, style)


def header_footer(canvas, doc):
    canvas.saveState()
    # верхняя плашка
    canvas.setFillColor(DARK)
    canvas.rect(0, PAGE_H - 10 * mm, PAGE_W, 10 * mm, stroke=0, fill=1)
    canvas.setFillColor(RED)
    canvas.rect(0, PAGE_H - 10.8 * mm, PAGE_W, 0.8 * mm, stroke=0, fill=1)
    canvas.setFillColor(WHITE)
    canvas.setFont('DejaVu-Bold', 8)
    canvas.drawString(MARGIN, PAGE_H - 6.8 * mm, 'FOSTER × КАКТУСА — ВЛОЖЕНИЯ И ПЛАНИРУЕМЫЙ ДОХОД')
    canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 6.8 * mm, '5 ДЕКАБРЯ 2026 • FOSTER NIGHTCLUB, КРЫМ')
    # низ
    canvas.setFillColor(MUTED)
    canvas.setFont('DejaVu', 7.5)
    canvas.drawString(MARGIN, 8 * mm, 'Источник: презентации репозитория — Foster_Gleb_Noemy_Presentation.pptx/.pdf, FOSTER_Club_презентация.pptx')
    canvas.drawRightString(PAGE_W - MARGIN, 8 * mm, f'Стр. {doc.page}')
    canvas.restoreState()


doc = BaseDocTemplate('/home/user/vecherinka/Вложения_и_планируемый_доход.pdf',
                      pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                      topMargin=16 * mm, bottomMargin=14 * mm,
                      title='Вложения и планируемый доход — Gleb & Noemy, Foster Club, 5 декабря 2026',
                      author='vecherinka project')
frame = Frame(MARGIN, 14 * mm, PAGE_W - 2 * MARGIN, PAGE_H - 30 * mm, id='main')
doc.addPageTemplates([PageTemplate(id='page', frames=[frame], onPage=header_footer)])

CW = PAGE_W - 2 * MARGIN  # ширина контента
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


def make_table(header, rows, widths, total_row=None, right_cols=(), red_last=False):
    data = [[P(h, S_CELL_W) for h in header]]
    for row in rows:
        line = []
        for i, cell in enumerate(row):
            style = S_CELL_B if (isinstance(cell, str) and cell.startswith('ИТОГО')) else (
                S_CELL_BR if (i in right_cols and isinstance(cell, str) and ('ИТОГО' in cell or cell.replace(' ', '').replace('₽', '').replace('–', '-').replace(',', '').lstrip('-').isdigit())) else
                S_CELL_R if i in right_cols else S_CELL)
            line.append(P(str(cell), style))
        data.append(line)
    if total_row:
        data.append([P(str(c), S_CELL_W if i == 0 else (S_CELL_BR if i in right_cols else S_CELL_W)) for i, c in enumerate(total_row)])
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
    for i in range(1, len(data) - (1 if total_row else 0)):
        if i % 2 == 0:
            style.append(('BACKGROUND', (0, i), (-1, i), GREY))
    if total_row:
        style.append(('BACKGROUND', (0, -1), (-1, -1), RED if red_last else DARK))
    t.setStyle(TableStyle(style))
    return t


# ================= ТИТУЛ =================
title_tbl = Table([
    [P('ВЛОЖЕНИЯ И ПЛАНИРУЕМЫЙ ДОХОД', S_TITLE)],
    [P('Вечер <b>Gleb Filipchenkow &amp; Noemy Abrantes</b> — «Чёрным по красному» • первый привоз в Крым (хиты «Mallorca», «Garota»/«HOJE)', S_SUB)],
    [P('<b>Суббота, 5 декабря 2026 • 22:00–06:00 • Foster Nightclub, Крым</b> (Mriya Resort &amp; SPA, Ялта, с. Оползневое) • 300–350 гостей, цель SOLD OUT • билеты 3 000 → 3 500 → 4 000 ₽ • старт продаж 01.11.2026', S_SUB)],
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
kpi = Table([[
    P('284 476 ₽<br/><font size="7.5" color="#B9B2A4">ВЛОЖЕНИЯ FOSTER</font>', st('k1', fontName='DejaVu-Bold', fontSize=13, leading=15, textColor=WHITE, alignment=TA_CENTER)),
    P('~1 120 040 ₽<br/><font size="7.5" color="#B9B2A4">ВЛОЖЕНИЯ КАКТУСЫ</font>', st('k2', fontName='DejaVu-Bold', fontSize=13, leading=15, textColor=WHITE, alignment=TA_CENTER)),
    P('768 025 ₽<br/><font size="7.5" color="#B9B2A4">ДОХОД FOSTER + КАЛЬЯНЫ</font>', st('k3', fontName='DejaVu-Bold', fontSize=13, leading=15, textColor=WHITE, alignment=TA_CENTER)),
    P('483 549 ₽<br/><font size="7.5" color="#B9B2A4">ЧИСТАЯ ПРИБЫЛЬ (63%)</font>', st('k4', fontName='DejaVu-Bold', fontSize=13, leading=15, textColor=WHITE, alignment=TA_CENTER)),
    P('330 000 ₽<br/><font size="7.5" color="#B9B2A4">ГАРАНТ ПО ДОГОВОРУ</font>', st('k5', fontName='DejaVu-Bold', fontSize=13, leading=15, textColor=WHITE, alignment=TA_CENTER)),
]], colWidths=[CW / 5.0] * 5)
kpi.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (0, 0), DARK),
    ('BACKGROUND', (1, 0), (1, 0), DARK),
    ('BACKGROUND', (2, 0), (2, 0), DARK),
    ('BACKGROUND', (3, 0), (3, 0), RED),
    ('BACKGROUND', (4, 0), (4, 0), RED),
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
    ['Вложения (затраты)', '284 476 ₽ — только операционка (аренда, артисты, оборудование, проживание = 0 ₽)',
     '~1 120 040 ₽ — гонорар 550К + логистика 80К + промо 100К + артисты 75К + тех/видео/фото 85К + охрана 30К + райдер 30К + компенсации клубу 200 040 ₽'],
    ['Планируемый доход', '768 025 ₽ + кальяны Макао (бар 368 025 + кухня 100 000 + коктейли 180 000 + столы 50 000 + VIP 70 000)',
     'Билеты: 825 000 ₽ (250 чел.) / 1 020 000 ₽ (300 чел.) / 1 190 000 ₽ (350 чел., средний чек 3 400 ₽)'],
    ['Планируемый результат', '+483 549 ₽ чистыми (~63% рентабельности)',
     '–295 040 ₽ (250 чел.) / –100 040 ₽ (300 чел.) / +69 960 ₽ только при SOLD OUT 350'],
    ['Гарант / риск', '330 000 ₽ гарантированы договором (коктейли 180К + столы 50К + VIP 70К) + 20 040 ₽ шаттл; оплата 1–3 дня после события',
     'Высокий риск — в плюс только при 350 гостях; мотивация продавать до последнего билета'],
]
story.append(make_table(['Показатель', 'Клуб FOSTER (площадка)', 'Организатор КАКТУСА'],
                        summary_rows, [CW * 0.18, CW * 0.40, CW * 0.42]))
story.append(Spacer(1, 3 * mm))
story.append(P('Участники: <b>FOSTER</b> — площадка (клуб, бар, кухня, персонал, VIP-залы, аренда 0 ₽) • <b>КАКТУСА</b> — организатор (все риски: гонорары, логистика, продакшн, промо, продажи билетов) • <b>MRIYA</b> — партнёр (проживание артистов без скидки).', S_BODY))

story.append(PageBreak())

# ================= ВЛОЖЕНИЯ FOSTER =================
story.append(section_bar('1. ВЛОЖЕНИЯ (ЗАТРАТЫ) КЛУБА FOSTER — ВСЕГО 284 476 ₽'))
story.append(Spacer(1, 2 * mm))
foster_rows = [
    ['Персонал клуба', '165 470 ₽', 'Официанты, бармены, повара, менеджмент (~58% всех затрат)'],
    ['Себестоимость бара (фудкост)', '92 006 ₽', '25% от валовой выручки бара 368 025 ₽'],
    ['Себестоимость кухни (фудкост)', '27 000 ₽', '27% от валовой выручки кухни 100 000 ₽'],
    ['Аренда площадки', '0 ₽', 'Площадка своя'],
    ['Гонорар артистов', '0 ₽', 'Платит Кактуса'],
    ['Оборудование / свет / звук', '0 ₽', 'Уже на месте'],
    ['Проживание артистов (скидка Mriya)', '0 ₽', 'Отменено по договорённости'],
]
story.append(make_table(['Статья затрат', 'Сумма', 'Примечание'], foster_rows,
                        [CW * 0.38, CW * 0.16, CW * 0.46],
                        total_row=['ИТОГО вложения Foster', '284 476 ₽', 'Фактически нулевые инвестиции — только операционка, заложенная в выручку'],
                        right_cols=(1,), red_last=True))
story.append(Spacer(1, 6 * mm))

# ================= ВЛОЖЕНИЯ КАКТУСЫ =================
story.append(section_bar('2. ВЛОЖЕНИЯ ОРГАНИЗАТОРА КАКТУСА — ВСЕГО ~1 120 040 ₽'))
story.append(Spacer(1, 2 * mm))
kaktus_rows = [
    ['Гонорар хедлайнеров (Gleb Filipchenkow & Noemy Abrantes)', '550 000 ₽', 'Фикс'],
    ['Компенсация клубу за welcome-коктейли (300 × 600 ₽)', '180 000 ₽', '→ Foster, фикс в договоре, независимо от явки'],
    ['Оплата шаттла (охрана, клининг, хостес Велком-центра)', '20 040 ₽', '→ Foster, фикс в договоре'],
    ['Перелёт / поезд / трансфер (логистика)', '~80 000 ₽', 'Допущение'],
    ['Продвижение, блогеры, афиши (промо)', '~100 000 ₽', 'Допущение'],
    ['Местные артисты (перформанс, go-go)', '75 000 ₽', 'Среднее'],
    ['Технический райдер (сверх клуба)', '~30 000 ₽', 'Допущение'],
    ['Видеосъёмка и монтаж', '30 000 ₽', 'Продакшн'],
    ['Фотоотчёт', '25 000 ₽', 'Продакшн'],
    ['Внешняя охрана', '~30 000 ₽', 'Сверх шаттла'],
]
story.append(make_table(['Статья затрат', 'Сумма', 'Тип / примечание'], kaktus_rows,
                        [CW * 0.44, CW * 0.15, CW * 0.41],
                        total_row=['ИТОГО вложения Кактусы', '~1 120 040 ₽', 'Все риски на организаторе; в плюс только при SOLD OUT 350 гостей'],
                        right_cols=(1,), red_last=True))
story.append(Spacer(1, 2 * mm))
story.append(P('Из них фиксированные компенсации клубу Foster: <b>200 040 ₽</b> = 180 000 ₽ (коктейли) + 20 040 ₽ (шаттл) — не зависят от явки.', S_NOTE))

story.append(PageBreak())

# ================= ДОХОДЫ И ПРИБЫЛЬ =================
story.append(section_bar('3. ПЛАНИРУЕМЫЙ ДОХОД КЛУБА FOSTER — 768 025 ₽ + КАЛЬЯНЫ МАКАО'))
story.append(Spacer(1, 2 * mm))
income_rows = [
    ['Бар (валовая выручка)', '368 025 ₽', ''],
    ['Кухня (валовая выручка)', '100 000 ₽', ''],
    ['Welcome-коктейли (компенсация от Кактусы)', '180 000 ₽', '300 × 600 ₽ — фикс, независимо от явки'],
    ['Депозиты столов', '50 000 ₽', '4 × 5 000 + 3 × 10 000'],
    ['VIP №1 (депозит)', '40 000 ₽', ''],
    ['VIP №2 (депозит)', '30 000 ₽', ''],
]
story.append(make_table(['Источник дохода', 'Сумма', 'Примечание'], income_rows,
                        [CW * 0.40, CW * 0.16, CW * 0.44],
                        total_row=['ИТОГО доходы Foster', '768 025 ₽', '+ кальяны от Макао — по факту продаж, полностью клубу'],
                        right_cols=(1,)))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('4. ЧИСТАЯ ПРИБЫЛЬ КЛУБА FOSTER ЗА ВЕЧЕР'))
story.append(Spacer(1, 2 * mm))
profit_rows = [
    ['Доходы', '768 025 ₽', 'Бар 368 025 + кухня 100 000 + коктейли 180 000 + столы 50 000 + VIP 70 000'],
    ['– Затраты (вложения клуба)', '–284 476 ₽', 'Персонал 165 470 + бар 92 006 + кухня 27 000'],
    ['= Чистая прибыль за вечер', '483 549 ₽', '768 025 – 284 476 = 483 549 ₽ + кальяны Макао сверху'],
    ['Рентабельность', '~63%', '483 549 / 768 025'],
    ['Гарантировано договором (фикс)', '330 000 ₽', '180 000 коктейли + 50 000 столы + 70 000 VIP №1/№2 — независимо от явки; даже при 200 гостях клуб в плюсе. + 20 040 ₽ шаттл. Оплата 1–3 рабочих дня после события'],
]
story.append(make_table(['Показатель', 'Сумма', 'Формула / примечание'], profit_rows,
                        [CW * 0.30, CW * 0.16, CW * 0.54], right_cols=(1,)))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('5. СЦЕНАРИИ ДОХОДА КАКТУСЫ — ПРОДАЖА БИЛЕТОВ (СРЕДНИЙ ЧЕК 3 400 ₽)'))
story.append(Spacer(1, 2 * mm))
scen_rows = [
    ['Пессимистичный', '250', '3 300 ₽', '825 000 ₽', '–295 040 ₽ (убыток)'],
    ['Базовый', '300', '3 400 ₽', '1 020 000 ₽', '–100 040 ₽ (убыток)'],
    ['Оптимистичный / SOLD OUT', '350', '3 400 ₽', '1 190 000 ₽', '+69 960 ₽ (прибыль)'],
]
story.append(make_table(['Сценарий', 'Гости, чел.', 'Средний билет', 'Выручка', 'Результат (– 1 120 040 ₽ затрат)'],
                        scen_rows, [CW * 0.24, CW * 0.11, CW * 0.14, CW * 0.15, CW * 0.36], right_cols=(1, 2, 3, 4)))
story.append(Spacer(1, 2 * mm))
story.append(P('Ступенчатая цена билетов: <b>3 000 ₽</b> (01–07.11) → <b>3 500 ₽</b> (08–14.11) → <b>4 000 ₽</b> (с 15.11 / в день события). Кактуса в плюс только при SOLD OUT 350 — мотивация продавать до последнего билета. Клуб Foster защищён: его 200 040 ₽ компенсаций зафиксированы в договоре независимо от явки.', S_NOTE))
story.append(Spacer(1, 5 * mm))

story.append(section_bar('ЧТО ФИКСИРОВАТЬ В ДОГОВОРЕ'))
story.append(Spacer(1, 2 * mm))
checklist = [
    '☑ 180 000 ₽ — компенсация коктейлей (300 × 600 ₽, фикс)',
    '☑ 20 040 ₽ — оплата шаттла (охрана / клининг / хостес)',
    '☑ 50 000 ₽ — депозиты столов + 70 000 ₽ — VIP №1/№2 (итого гарант 330 000 ₽)',
    '☑ Оплата 1–3 рабочих дня после события',
    '☑ Кальяны Макао — по факту продаж, полностью в пользу клуба',
]
for item in checklist:
    story.append(P(item, S_BODY))
    story.append(Spacer(1, 1.2 * mm))
story.append(Spacer(1, 2 * mm))
story.append(P('<b>Следующие шаги:</b> согласовать договор (фикс 330 000 ₽ + компенсации) → старт продаж 01.11.2026 → 5 декабря 2026, SOLD OUT 350 гостей.', S_BODY))

doc.build(story)
print('Сохранено: /home/user/vecherinka/Вложения_и_планируемый_доход.pdf')
